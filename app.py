import gradio as gr
import google.generativeai as genai
import os
import requests
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

# 1. 配置所有密钥 (从环境变量读取)
genai.configure(api_key=os.environ.get("GEMINI_API_KEY"))
model = genai.GenerativeModel('gemini-2.5-flash')
FS_APP_ID = os.environ.get("FEISHU_APP_ID")
FS_APP_SECRET = os.environ.get("FEISHU_APP_SECRET")

# ==========================================
# 2. AI 语言处理逻辑 (纠错、校对、翻译)
# ==========================================
prompt_typo = "你的唯一任务是：检查文档或图片中的【中文错别字和语病】。哪怕只有 10% 的把握也请指出。"
prompt_proofread = "你的唯一任务是：做专业的【中英双语校对】。对比中英文翻译是否对齐，检查英文语法拼写。"
prompt_translate = "你的唯一任务是：做地道的【中译英翻译】。提取内容并输出符合商务规范的纯英文结果。"

def process_ai_task(file_obj, prompt_text):
    if file_obj is None: return "⚠️ 请先上传文件哦！"
    try:
        gemini_file = genai.upload_file(file_obj.name)
        response = model.generate_content([prompt_text, gemini_file])
        return response.text
    except Exception as e:
        error_msg = str(e)
        if "429" in error_msg: return "⏳ 速度太快啦，请等待 1 分钟再试~"
        return f"❌ 错误: {error_msg}"

# ==========================================
# 3. 飞书链接解析逻辑
# ==========================================
def get_feishu_token():
    url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal"
    req_body = {"app_id": FS_APP_ID, "app_secret": FS_APP_SECRET}
    r = requests.post(url, json=req_body)
    return r.json().get("tenant_access_token")

def fetch_feishu_data(url_link):
    try:
        # 提取 spreadsheet_token
        ss_token = url_link.split("/sheets/")[1].split("?")[0]
    except:
        return None, "❌ 链接格式似乎不对，请确保是飞书表格链接"
    
    token = get_feishu_token()
    if not token: return None, "❌ 无法获取飞书授权，请检查 Secret 配置"
    
    # 读取 A1 到 Z500 的范围
    data_url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{ss_token}/values/A1:Z500"
    headers = {"Authorization": f"Bearer {token}"}
    r = requests.get(data_url, headers=headers)
    res_data = r.json()
    
    if res_data.get("code") != 0:
        return None, f"❌ 飞书接口报错: {res_data.get('msg')}"
    
    return res_data.get("data").get("valueRange").get("values"), "OK"

# ==========================================
# 4. PPT 生成逻辑 (黑底白字提词器)
# ==========================================
def link_to_pptx(link, col_letter):
    if not link: return None, "⚠️ 请输入链接"
    raw_data, msg = fetch_feishu_data(link)
    if raw_data is None: return None, msg
    
    # 字母转索引
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
        
        txBox = slide.shapes.add_textbox(0, 0, prs.slide_width, prs.slide_height)
        tf = txBox.text_frame
        tf.vertical_anchor = 1
        p = tf.paragraphs[0]
        p.text = content
        p.alignment = PP_ALIGN.CENTER
        p.font.size, p.font.bold, p.font.color.rgb = Pt(60), True, RGBColor(255, 255, 255)
        count += 1

    output_name = "飞书提词器导出.pptx"
    prs.save(output_name)
    return output_name, f"✅ 成功提取 {count} 页提词！"

# ==========================================
# 5. 构建整合界面
# ==========================================
with gr.Blocks(title="AI 智能文档工作站") as demo:
    gr.Markdown("# 🚀 AI 智能文档工作站 V5.0")
    
    with gr.Tabs():
        # 标签页一：AI 任务
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

        # 标签页二：飞书链接转 PPT
        with gr.TabItem("🎬 飞书一键转 PPT"):
            gr.Markdown("### 粘贴飞书表格链接，指定列号，一键生成")
            with gr.Row():
                with gr.Column():
                    link_input = gr.Textbox(label="1. 粘贴飞书表格链接", placeholder="https://xxx.feishu.cn/sheets/...")
                    col_input = gr.Textbox(label="2. 题词所在列 (如 C)", value="C")
                    ppt_btn = gr.Button("🚀 立即生成 PPT", variant="primary")
                with gr.Column():
                    ppt_status = gr.Markdown("状态：准备就绪")
                    ppt_download = gr.File(label="3. 下载生成的 PPT")
            
            ppt_btn.click(fn=link_to_pptx, inputs=[link_input, col_input], outputs=[ppt_download, ppt_status])

demo.launch()