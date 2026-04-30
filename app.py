import gradio as gr
import google.generativeai as genai
import os
import requests
import json
import io
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

# ==========================================
# 1. 基础配置
# ==========================================
genai.configure(api_key=os.environ.get("GEMINI_API_KEY"))
model = genai.GenerativeModel('gemini-2.5-flash')
FS_APP_ID = os.environ.get("FEISHU_APP_ID")
FS_APP_SECRET = os.environ.get("FEISHU_APP_SECRET")

# ==========================================
# 2. AI 语言处理 (保留原有全部功能)
# ==========================================
prompt_typo = """你的唯一任务是：检查文档或图片中的【中文错别字和语病】。
1. 请保持极高的敏感度，哪怕只有 10% 的错漏把握也请指出，宁可错杀不可放过。
2. 请忽略文档中的纯英文内容，专注中文。
3. 【核心输出格式】：请务必【逐页】输出检查结果。
   - 如果该页没有任何错别字或语病，请严格输出：“第X页：OKOK”
   - 如果该页有需要修改的地方，请清晰列出：“第X页：[原文] -> [修改建议]及原因”"""

prompt_proofread = "做专业的【中英双语校对】。对比中英文翻译是否对齐，检查英文语法和拼写。"
prompt_translate = "做地道的【中译英翻译】。提取内容并输出符合商务规范的纯英文结果。"

def process_ai_task(file_obj, prompt_text):
    if file_obj is None: return "⚠️ 请先上传文件哦！"
    file_name = file_obj.name.lower()
    if not (file_name.endswith('.pdf') or file_name.endswith('.png') or file_name.endswith('.jpg') or file_name.endswith('.jpeg')):
        return "❌ 仅支持 PDF 或图片。请将 PPT/Word 导出为 PDF 后上传。"
    try:
        gemini_file = genai.upload_file(file_obj.name)
        response = model.generate_content([prompt_text, gemini_file])
        return response.text
    except Exception as e:
        return f"❌ 错误: {str(e)}"

# ==========================================
# 3. AI 视觉理解核心 (用于飞书图片摘要)
# ==========================================
def get_image_summary(image_bytes):
    """视觉识别 B 列图片并生成 5 字摘要"""
    if not image_bytes: return ""
    try:
        img_part = {"mime_type": "image/png", "data": image_bytes}
        prompt = "这是演讲幻灯片的下一页内容，请用5个字以内总结其核心要点，作为给演讲者的‘下一页预告’提示（例如：业务增长图表、年度目标展望）"
        response = model.generate_content([prompt, img_part])
        return response.text.strip()
    except:
        return "图片解析失败"

# ==========================================
# 4. 飞书数据引擎 (读取与回写)
# ==========================================
def get_feishu_token():
    url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal"
    r = requests.post(url, json={"app_id": FS_APP_ID, "app_secret": FS_APP_SECRET})
    return r.json().get("tenant_access_token")

def download_fs_media(file_token, token):
    url = f"https://open.feishu.cn/open-apis/drive/v1/medias/{file_token}/download"
    headers = {"Authorization": f"Bearer {token}"}
    r = requests.get(url, headers=headers)
    return r.content if r.status_code == 200 else None

def update_feishu_cell(ss_token, sheet_id, row_index, text, token):
    """【回写引擎】将摘要写回 D 列"""
    col_name = "D"
    range_str = f"{sheet_id}!{col_name}{row_index+1}:{col_name}{row_index+1}"
    url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{ss_token}/values"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    body = {"valueRange": {"range": range_str, "values": [[text]]}}
    requests.put(url, headers=headers, json=body)

def parse_link(url_link, token):
    headers = {"Authorization": f"Bearer {token}"}
    if "/sheets/" in url_link:
        ss_token = url_link.split("/sheets/")[1].split("?")[0].split("#")[0]
    elif "/wiki/" in url_link:
        wiki_token = url_link.split("/wiki/")[1].split("?")[0].split("#")[0]
        node_res = requests.get(f"https://open.feishu.cn/open-apis/wiki/v2/spaces/get_node?token={wiki_token}", headers=headers).json()
        ss_token = node_res.get("data", {}).get("node", {}).get("obj_token")
    else: return None, None, "格式不支持"
    
    meta_res = requests.get(f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{ss_token}/metainfo", headers=headers).json()
    sheet_id = meta_res["data"]["sheets"][0]["sheetId"]
    return ss_token, sheet_id, "OK"

# ==========================================
# 5. 引擎 A：自动生成表格摘要并回写 D 列
# ==========================================
def generate_summaries_handler(link):
    if not link: return "⚠️ 请先粘贴链接"
    token = get_feishu_token()
    ss_token, sheet_id, msg = parse_link(link, token)
    if not ss_token: return msg
    
    data_url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{ss_token}/values/{sheet_id}!A1:Z500?valueRenderOption=Formula"
    headers = {"Authorization": f"Bearer {token}"}
    raw_data = requests.get(data_url, headers=headers).json().get("data", {}).get("valueRange", {}).get("values", [])
    
    processed_count = 0
    img_col = ord('B') - ord('A')
    
    for i, row in enumerate(raw_data):
        if len(row) > img_col:
            cell_data = row[img_col]
            if isinstance(cell_data, list) and len(cell_data) > 0 and 'fileToken' in cell_data[0]:
                file_token = cell_data[0]['fileToken']
                img_bytes = download_fs_media(file_token, token)
                summary = get_image_summary(img_bytes)
                if summary:
                    update_feishu_cell(ss_token, sheet_id, i, summary, token)
                    processed_count += 1
                    
    return f"✅ 摘要生成完毕！已成功在 D 列回写 {processed_count} 条摘要。请刷新飞书表格查看并确认。"

# ==========================================
# 6. 引擎 B：导出智能 PPT (左下角红框，C列正文，D列摘要)
# ==========================================
def export_ppt_handler(link, col_letter):
    if not link: return None, "⚠️ 请先粘贴链接"
    token = get_feishu_token()
    ss_token, sheet_