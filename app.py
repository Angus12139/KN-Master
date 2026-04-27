import gradio as gr
import google.generativeai as genai
import os
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

# 1. 配置 AI 通行证
my_api_key = os.environ.get("GEMINI_API_KEY") 
genai.configure(api_key=my_api_key)
model = genai.GenerativeModel('gemini-2.5-flash')

# ==========================================
# 2. AI 逻辑部分 (三个按钮的专属指令)
# ==========================================
prompt_typo = "你的唯一任务是：检查文档或图片中的【中文错别字和语病】。哪怕只有 10% 的把握也请指出。请忽略英文。"
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
# 3. 提词器 PPT 导出逻辑部分
# ==========================================
def table_to_pptx(file_obj):
    if file_obj is None: return None, "⚠️ 请先上传表格"
    try:
        df = pd.read_csv(file_obj.name) if file_obj.name.endswith('.csv') else pd.read_excel(file_obj.name)
        prs = Presentation()
        prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
        for _, row in df.iterrows():
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            slide.background.fill.solid()
            slide.background.fill.fore_color.rgb = RGBColor(0, 0, 0)
            txBox = slide.shapes.add_textbox(0, 0, prs.slide_width, prs.slide_height)
            tf = txBox.text_frame
            tf.vertical_anchor = 1
            p = tf.paragraphs[0]
            p.text = str(row[0])
            p.alignment = PP_ALIGN.CENTER
            p.font.size, p.font.bold, p.font.color.rgb = Pt(60), True, RGBColor(255, 255, 255)
        output_name = "发布会提词器.pptx"
        prs.save(output_name)
        return output_name, f"✅ 成功生成 {len(df)} 页提词！"
    except Exception as e:
        return None, f"❌ 失败: {str(e)}"

# ==========================================
# 4. 构建整合界面 (使用 Tabs 标签页)
# ==========================================
with gr.Blocks(title="AI 智能文档工作站") as demo:
    gr.Markdown("# 🚀 AI 智能文档工作站 V4.0 (完整版)")
    
    with gr.Tabs():
        # 标签页一：AI 语言处理
        with gr.TabItem("📝 AI 语言校对与翻译"):
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

        # 标签页二：提词器自动化
        with gr.TabItem("🎬 提词器 PPT 导出"):
            with gr.Row():
                table_file = gr.File(label="上传飞书导出的 Excel/CSV")
                with gr.Column():
                    ppt_btn = gr.Button("🔥 生成提词 PPT", variant="primary")
                    ppt_status = gr.Markdown("状态：等待上传")
            ppt_download = gr.File(label="下载生成的 PPT")
            
            ppt_btn.click(fn=table_to_pptx, inputs=table_file, outputs=[ppt_download, ppt_status])

demo.launch()