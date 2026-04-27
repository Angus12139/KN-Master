import gradio as gr
import google.generativeai as genai
import os
import requests
import json
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

# ==========================================
# 1. 基础配置
# ==========================================
genai.configure(api_key=os.environ.get("GEMINI_API_KEY"))
model = genai.GenerativeModel('gemini-2.5-flash')
FS_APP_ID = os.environ.get("FEISHU_APP_ID")
FS_APP_SECRET = os.environ.get("FEISHU_APP_SECRET")

# ==========================================
# 2. AI 语言处理 (保持高敏纠错指令)
# ==========================================
prompt_typo = """你的唯一任务是：检查文档或图片中的【中文错别字和语病】。
1. 请保持极高的敏感度，哪怕只有 10% 的错漏把握也请指出。
2. 【核心输出格式】：请务必【逐页】输出检查结果。
   - 如果该页没有问题，输出：“第X页：OKOK”
   - 如果该页有问题，输出：“第X页：[原文] -> [修改建议]及原因”"""

def process_ai_task(file_obj, prompt_text):
    if file_obj is None: return "⚠️ 请先上传文件"
    if not (file_obj.name.lower().endswith(('.pdf', '.png', '.jpg', '.jpeg'))):
        return "❌ 仅支持 PDF 或图片格式。"
    try:
        gemini_file = genai.upload_file(file_obj.name)
        response = model.generate_content([prompt_text, gemini_file])
        return response.text
    except Exception as e: return f"❌ 错误: {str(e)}"

# ==========================================
# 3. 飞书 API (升级：请求原始富文本数据)
# ==========================================
def get_feishu_token():
    url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal"
    r = requests.post(url, json={"app_id": FS_APP_ID, "app_secret": FS_APP_SECRET})
    return r.json().get("tenant_access_token")

def fetch_feishu_data(url_link):
    token = get_feishu_token()
    headers = {"Authorization": f"Bearer {token}"}
    
    if "/sheets/" in url_link:
        ss_token = url_link.split("/sheets/")[1].split("?")[0].split("#")[0]
    elif "/wiki/" in url_link:
        wiki_token = url_link.split("/wiki/")[1].split("?")[0].split("#")[0]
        node_res = requests.get(f"https://open.feishu.cn/open-apis/wiki/v2/spaces/get_node?token={wiki_token}", headers=headers).json()
        ss_token = node_res.get("data", {}).get("node", {}).get("obj_token")
    else: return None, "❌ 链接格式不支持"

    meta_res = requests.get(f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{ss_token}/metainfo", headers=headers).json()
    first_sheet_id = meta_res["data"]["sheets"][0]["sheetId"]

    # 关键修改：增加 valueRenderOption=Formula 参数，确保飞书返回富文本 JSON 结构
    data_url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{ss_token}/values/{first_sheet_id}!A1:Z500?valueRenderOption=Formula"
    r = requests.get(data_url, headers=headers)
    return r.json().get("data", {}).get("valueRange", {}).get("values", []), "OK"

# ==========================================
# 4. PPT 生成逻辑 (样式同步 + 左对齐)
# ==========================================
def hex_to_rgb(hex_color):
    """将 #RRGGBB 转换为 RGB 元组"""
    hex_color = hex_color.lstrip('#')
    return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))

def link_to_pptx(link, col_letter):
    if not link: return None, "⚠️ 请输入链接"
    raw_data, msg = fetch_feishu_data(link)
    if raw_data is None: return None, msg
    
    col_index = ord(col_letter.upper()) - ord('A')
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    
    count = 0
    for row in raw_data:
        if len(row) <= col_index: continue
        content_obj = row[col_index]
        
        # 统一处理：不管是字符串还是富文本列表，都转为处理列表
        segments = []
        full_text = ""
        
        if isinstance(content_obj, list): # 如果是富文本格式
            segments = content_obj
            for s in segments: full_text += s.get('text', '')
        else: # 如果是普通字符串
            full_text = str(content_obj).strip()
            if full_text in ["None", "", "nan"]: continue
            segments = [{'text': full_text, 'segmentStyle': {'foreColor': '#FFFFFF'}}]

        if not full_text.strip(): continue

        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = RGBColor(0, 0, 0)
        
        margin = Inches(0.6)
        txBox = slide.shapes.add_textbox(margin, margin, prs.slide_width - margin*2, prs.slide_height - margin*2)
        tf = txBox.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = 1 # 垂直居中，但文字内部左对齐
        
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT # 【需求修改】：左对齐
        
        # 智能字号计算
        text_len = len(full_text)
        if text_len <= 15: font_size = 90
        elif text_len <= 40: font_size = 70
        elif text_len <= 80: font_size = 50
        elif text_len <= 150: font_size = 38
        else: font_size = 30

        # 遍历片段，添加 Run（样式块）
        for seg in segments:
            text_part = seg.get('text', '')
            if not text_part: continue
            
            run = p.add_run()
            run.text = text_part
            run.font.name = '微软雅黑'
            run.font.size = Pt(font_size)
            run.font.bold = seg.get('segmentStyle', {}).get('bold', True)
            
            # 颜色处理逻辑
            color_hex = seg.get('segmentStyle', {}).get('foreColor', '#FFFFFF')
            # 如果是黑色或未设置，在黑底PPT上显示为白色
            if color_hex.upper() == "#000000" or color_hex.upper() == "#121212":
                run.font.color.rgb = RGBColor(255, 255, 255)
            else:
                try:
                    rgb = hex_to_rgb(color_hex)
                    run.font.color.rgb = RGBColor(*rgb)
                except:
                    run.font.color.rgb = RGBColor(255, 255, 255)
        
        count += 1

    output_name = "发布会样式同步提词器.pptx"
    prs.save(output_name)
    return output_name, f"✅ 已同步样式并导出 {count} 页提词！"

# ==========================================
# 5. 界面
# ==========================================
with gr.Blocks(title="AI 智能工作站") as demo:
    gr.Markdown("# 🚀 AI 智能工作站 V5.4 (样式同步版)")
    with gr.Tabs():
        with gr.TabItem("📝 AI 语言处理"):
            with gr.Row():
                with gr.Column():
                    ai_file = gr.File(label="上传 PDF/图片")
                    btn_typo = gr.Button("🇨🇳 纠错 (按页显示)", variant="primary")
                ai_output = gr.Textbox(label="结果", lines=15)
            btn_typo.click(fn=lambda f: process_ai_task(f, prompt_typo), inputs=ai_file, outputs=ai_output)

        with gr.TabItem("🎬 飞书一键转 PPT"):
            gr.Markdown("### 提取飞书 C 列题词：支持标红文字同步、左对齐、微软雅黑")
            with gr.Row():
                with gr.Column():
                    link_input = gr.Textbox(label="飞书链接")
                    col_input = gr.Textbox(label="题词列", value="C")
                    ppt_btn = gr.Button("🚀 立即生成 PPT", variant="primary")
                with gr.Column():
                    ppt_download = gr.File(label="下载 PPT")
                    ppt_status = gr.Markdown("状态：准备就绪")
            ppt_btn.click(fn=link_to_pptx, inputs=[link_input, col_input], outputs=[ppt_download, ppt_status])

demo.launch()