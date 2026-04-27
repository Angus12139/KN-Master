import gradio as gr
import google.generativeai as genai
import os
import requests
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
# 2. AI 语言处理 (高敏加强版指令)
# ==========================================
prompt_typo = """你的唯一任务是：检查文档或图片中的【中文错别字和语病】。
1. 请保持极高的敏感度，哪怕只有 10% 的错漏把握也请指出，宁可错杀不可放过。
2. 请忽略文档中的纯英文内容，专注中文。
3. 【核心输出格式】：请务必【逐页】输出检查结果。
   - 如果该页没有任何错别字或语病，请严格输出：“第X页：OKOK”
   - 如果该页有需要修改的地方，请清晰列出：“第X页：[原文] -> [修改建议]及原因”"""

prompt_proofread = "你的唯一任务是：做专业的【中英双语校对】。对比中英文翻译是否对齐，检查英文语法和拼写。"
prompt_translate = "你的唯一任务是：做地道的【中译英翻译】。提取内容并输出符合商务规范的纯英文结果。"

def process_ai_task(file_obj, prompt_text):
    if file_obj is None: return "⚠️ 请先上传文件哦！"
    
    # 文件格式拦截保护
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
# 3. 飞书 API 智能解析引擎
# ==========================================
def get_feishu_token():
    url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal"
    r = requests.post(url, json={"app_id": FS_APP_ID, "app_secret": FS_APP_SECRET})
    return r.json().get("tenant_access_token")

def fetch_feishu_data(url_link):
    token = get_feishu_token()
    if not token: return None, "❌ 无法获取飞书授权，请检查 Secret 配置"
    headers = {"Authorization": f"Bearer {token}"}
    
    # 智能判断链接类型
    if "/sheets/" in url_link:
        ss_token = url_link.split("/sheets/")[1].split("?")[0].split("#")[0]
        obj_type = "sheet"
    elif "/wiki/" in url_link:
        wiki_token = url_link.split("/wiki/")[1].split("?")[0].split("#")[0]
        node_url = f"https://open.feishu.cn/open-apis/wiki/v2/spaces/get_node?token={wiki_token}"
        node_res = requests.get(node_url, headers=headers).json()
        if node_res.get("code") != 0: return None, f"❌ 知识库解析失败: {node_res.get('msg')}"
        ss_token = node_res.get("data", {}).get("node", {}).get("obj_token")
        obj_type = node_res.get("data", {}).get("node", {}).get("obj_type")
    else:
        return None, "❌ 链接格式不支持，请粘贴包含 /sheets/ 或 /wiki/ 的链接。"

    if obj_type != "sheet":
        return None, f"❌ 抱歉，链接里装的是【{obj_type}】，本工具仅支持【电子表格】。"

    # 获取表格第一页的 sheetId
    meta_url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{ss_token}/metainfo"
    meta_res = requests.get(meta_url, headers=headers).json()
    if meta_res.get("code") != 0: return None, f"❌ 获取基本信息失败: {meta_res.get('msg')}"
    
    try:
        first_sheet_id = meta_res["data"]["sheets"][0]["sheetId"]
    except Exception as e:
        return None, f"❌ 解析工作表 ID 失败。飞书实际返回: {meta_res}"

    # 读取表格数据
    data_url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{ss_token}/values/{first_sheet_id}!A1:Z500"
    r = requests.get(data_url, headers=headers)
    res_data = r.json()
    
    if res_data.get("code") != 0:
        return None, f"❌ 读取表格报错: {res_data.get('code')} - {res_data.get('msg')}"
    
    return res_data.get("data", {}).get("valueRange", {}).get("values", []), "OK"

# ==========================================
# 4. PPT 生成逻辑 (智能排版版)
# ==========================================
def link_to_pptx(link, col_letter):
    if not link: return None, "⚠️ 请输入链接"
    raw_data, msg = fetch_feishu_data(link)
    if raw_data is None: return None, msg
    
    col_index = ord(col_letter.upper()) - ord('A')
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5) # 16:9
    
    count = 0
    for row in raw_data:
        if len(row) <= col_index: continue
        content = str(row[col_index]).strip()
        if content == "None" or not content or content == "nan": continue
        
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = RGBColor(0, 0, 0)
        
        # 设置安全边距，防止字出界
        margin = Inches(0.5)
        txBox = slide.shapes.add_textbox(margin, margin, prs.slide_width - margin*2, prs.slide_height - margin*2)
        tf = txBox.text_frame
        tf.word_wrap = True  # 开启自动换行
        tf.vertical_anchor = 1 # 垂直居中
        
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        p.text = content
        
        # 智能字号算法
        text_len = len(content)
        if text_len <= 15: font_size = 100
        elif text_len <= 40: font_size = 75
        elif text_len <= 80: font_size = 55
        elif text_len <= 150: font_size = 40
        else: font_size = 32
            
        p.font.size = Pt(font_size)
        p.font.bold = True
        p.font.color.rgb = RGBColor(255, 255, 255)
        p.font.name = '微软雅黑'
        
        count += 1

    output_name = "发布会智能提词器.pptx"
    prs.save(output_name)
    return output_name, f"✅ 已智能排版 {count} 页提词！"

# ==========================================
# 5. 整合界面
# ==========================================
with gr.Blocks(title="AI 智能文档工作站") as demo:
    gr.Markdown("# 🚀 AI 智能文档工作站 V5.3 (高敏排版版)")
    
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
            gr.Markdown("### 粘贴飞书表格/知识库链接，一键生成防出界提词 PPT")
            with gr.Row():
                with gr.Column():
                    link_input = gr.Textbox(label="1. 粘贴飞书链接", placeholder="https://xxx.feishu.cn/...")
                    col_input = gr.Textbox(label="2. 题词所在列 (如 C)", value="C")
                    ppt_btn = gr.Button("🚀 立即生成 PPT", variant="primary")
                with gr.Column():
                    ppt_status = gr.Markdown("状态：准备就绪")
                    ppt_download = gr.File(label="3. 下载生成的 PPT")
            
            ppt_btn.click(fn=link_to_pptx, inputs=[link_input, col_input], outputs=[ppt_download, ppt_status])

demo.launch()