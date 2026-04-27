import gradio as gr
import google.generativeai as genai
import os
import requests
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

# 1. 密钥配置
genai.configure(api_key=os.environ.get("GEMINI_API_KEY"))
model = genai.GenerativeModel('gemini-2.5-flash')
FS_APP_ID = os.environ.get("FEISHU_APP_ID")
FS_APP_SECRET = os.environ.get("FEISHU_APP_SECRET")

# ==========================================
# 2. AI 语言处理
# ==========================================
prompt_typo = """你的唯一任务是：检查文档或图片中的【中文错别字和语病】。
1. 哪怕只有 10% 的把握觉得某个词、某句话不对劲，也请务必指出来！
2. 请清晰列出：“原词/原句” -> “修改建议” -> “怀疑理由”。
3. 请忽略所有的英文内容，不要翻译，只专注于中文纠错。如果这一页没有错误,直接输出okokok,每一页都需要输出修改建议或者ok的结论"""
prompt_proofread = "做专业的【中英双语校对】。"
prompt_translate = "做地道的【中译英翻译】。"

def process_ai_task(file_obj, prompt_text):
    if file_obj is None: return "⚠️ 请先上传文件哦！"
    
    # 增加：智能拦截不支持的文件格式
    file_name = file_obj.name.lower()
    if not (file_name.endswith('.pdf') or file_name.endswith('.png') or file_name.endswith('.jpg') or file_name.endswith('.jpeg')):
        return "❌ 格式不支持：AI 目前只长了看 PDF 和图片的“眼睛”。请把你的 PPT 或 Word 导出为 PDF 格式后再上传哦！"

    try:
        gemini_file = genai.upload_file(file_obj.name)
        response = model.generate_content([prompt_text, gemini_file])
        return response.text
    except Exception as e:
        error_msg = str(e)
        if "429" in error_msg: return "⏳ 速度太快啦，请等待 1 分钟再试~"
        return f"❌ 错误: {error_msg}"
# ==========================================
# 3. 飞书 API 智能解析引擎 (核心升级区)
# ==========================================
def get_feishu_token():
    url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal"
    r = requests.post(url, json={"app_id": FS_APP_ID, "app_secret": FS_APP_SECRET})
    return r.json().get("tenant_access_token")

def fetch_feishu_data(url_link):
    token = get_feishu_token()
    if not token: return None, "❌ 无法获取飞书授权，请检查 Secret 配置"
    headers = {"Authorization": f"Bearer {token}"}
    
    # 1. 智能判断：是普通表格还是知识库？
    if "/sheets/" in url_link:
        ss_token = url_link.split("/sheets/")[1].split("?")[0].split("#")[0]
        obj_type = "sheet"
    elif "/wiki/" in url_link:
        wiki_token = url_link.split("/wiki/")[1].split("?")[0].split("#")[0]
        node_url = f"https://open.feishu.cn/open-apis/wiki/v2/spaces/get_node?token={wiki_token}"
        node_res = requests.get(node_url, headers=headers).json()
        if node_res.get("code") != 0:
            return None, f"❌ 知识库解析失败: {node_res.get('msg')}"
        node_info = node_res.get("data", {}).get("node", {})
        ss_token = node_info.get("obj_token")
        obj_type = node_info.get("obj_type")
    else:
        return None, "❌ 链接格式不支持，请粘贴包含 /sheets/ 或 /wiki/ 的链接。"

    if obj_type != "sheet":
        return None, f"❌ 抱歉，链接里装的是【{obj_type}】，本工具仅支持【电子表格】。"

    # 2. 获取表格元数据，精准抓取第一页的 sheetId
    meta_url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{ss_token}/metainfo"
    meta_res = requests.get(meta_url, headers=headers).json()
    if meta_res.get("code") != 0:
        return None, f"❌ 获取基本信息失败: {meta_res.get('msg')}"
    
    try:
        first_sheet_id = meta_res["data"]["sheets"][0]["sheetId"]
    except Exception as e:
        return None, f"❌ 解析工作表 ID 失败。飞书实际返回: {meta_res}"

    # 3. 带上正确的 sheetId 去拿数据
    data_url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{ss_token}/values/{first_sheet_id}!A1:Z500"
    r = requests.get(data_url, headers=headers)
    res_data = r.json()
    
    if res_data.get("code") != 0:
        return None, f"❌ 读取表格报错: {res_data.get('code')} - {res_data.get('msg')}"
    
    return res_data.get("data", {}).get("valueRange", {}).get("values", []), "OK"

# ==========================================
# 4. PPT 生成逻辑
# ==========================================
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
        content = str(row[col_index]).strip()
        if content == "None" or not content or content == "nan": continue
        
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = RGBColor(0, 0, 0)
        
        txBox = slide.shapes.add_textbox(0, 0, prs.slide_width, prs.slide_height)
        tf = txBox.text_frame
        tf.vertical_anchor = 1
        p = tf.paragraphs[0]
        p.text = content
        p.alignment = PP_ALIGN.CENTER
        p.font.size, p.font.bold, p.font.color.rgb = Pt(60), True, RGBColor(255, 255, 255)
        count += 1

    output_name = "飞书智能提词器.pptx"
    prs.save(output_name)
    return output_name, f"✅ 成功穿透知识库，提取 {count} 页提词！"

# ==========================================
# 5. 构建界面
# ==========================================
with gr.Blocks(title="AI 智能文档工作站") as demo:
    gr.Markdown("# 🚀 AI 智能文档工作站 V5.1 (智能解析版)")
    
    with gr.Tabs():
        with gr.TabItem("📝 AI 语言处理"):
            with gr.Row():
                with gr.Column():
                    ai_file = gr.File(label="上传 PDF 或图片")
                    with gr.Row():
                        btn_typo = gr.Button("🇨🇳 纠错", variant="primary")
                        btn_proof = gr.Button("⚖️ 校对", variant="secondary")
                        btn_trans = gr.Button("🌍 翻译", variant="secondary")
                ai_output = gr.Textbox(label="处理结果", lines=15)
            
            btn_typo.click(fn=lambda f: process_ai_task(f, prompt_typo), inputs=ai_file, outputs=ai_output)
            btn_proof.click(fn=lambda f: process_ai_task(f, prompt_proofread), inputs=ai_file, outputs=ai_output)
            btn_trans.click(fn=lambda f: process_ai_task(f, prompt_translate), inputs=ai_file, outputs=ai_output)

        with gr.TabItem("🎬 飞书一键转 PPT"):
            gr.Markdown("### 支持直接粘贴普通表格 /sheets/ 或 知识库 /wiki/ 链接")
            with gr.Row():
                with gr.Column():
                    link_input = gr.Textbox(label="1. 粘贴飞书链接", placeholder="https://xxx.feishu.cn/wiki/...")
                    col_input = gr.Textbox(label="2. 题词所在列 (如 C)", value="C")
                    ppt_btn = gr.Button("🚀 立即生成 PPT", variant="primary")
                with gr.Column():
                    ppt_status = gr.Markdown("状态：准备就绪")
                    ppt_download = gr.File(label="3. 下载生成的 PPT")
            
            ppt_btn.click(fn=link_to_pptx, inputs=[link_input, col_input], outputs=[ppt_download, ppt_status])

demo.launch()