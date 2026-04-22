import gradio as gr
import google.generativeai as genai
import os

# 1. 配置 AI
my_api_key = os.environ.get("GEMINI_API_KEY") 
genai.configure(api_key=my_api_key)
model = genai.GenerativeModel('gemini-2.5-flash')

# 2. 【升级版】中英双语深度校对提示词
bilingual_prompt = """
你是一个顶级的“中英双语翻译校对”和“语言规范检查”专家。
请对用户提供的文档或图片进行全方位的深度扫描。

你的检查任务包括：
1. 【中英对照检查】：如果页面内存在中英文对照，请检查翻译是否准确、语义是否对齐，指出任何意思偏差、术语不统一或漏译的地方。
2. 【英文拼写与语法】：检查所有英文内容的拼写（Spelling）、语法（Grammar）以及用词地道性。
3. 【格式规范检查】：检查标点符号（如中英文标点混用）、大小写规范、缩进以及段落排版错误。
4. 【中文错别字】：延续之前的逻辑，检查中文部分的错别字和语病。

【输出要求】：
- 哪怕只有 10% 的把握，也请指出潜在风险。
- 请使用表格或清晰的列表形式列出：【原内容】 -> 【建议修改】 -> 【风险等级及理由】。
- 风险等级分为：高风险（确定错误）、中风险（建议优化）、低风险（风格建议）。
"""

# 3. 核心工作函数
def check_typos(file_obj):
    if file_obj is None:
        return "请上传文件（PDF 或图片）"
    
    try:
        # 使用 Google 高速通道处理
        gemini_file = genai.upload_file(file_obj.name)
        # 发送提示词和文件
        response = model.generate_content([bilingual_prompt, gemini_file])
        return response.text
    except Exception as e:
        return f"处理过程中出现错误: {str(e)}"

# 4. 界面搭建
with gr.Blocks(title="AI 双语深度校对雷达") as demo:
    gr.Markdown("# 🛰️ AI 双语深度校对雷达")
    gr.Markdown("支持 PDF 和图片。自动识别中英翻译对照、英文语法拼写及格式规范。")
    
    with gr.Row():
        with gr.Column():
            file_input = gr.File(label="上传 PDF 或图片")
            submit_btn = gr.Button("🔍 开始深度校对", variant="primary")
        
    with gr.Row():
        output_text = gr.Textbox(label="校对报告（包含翻译对比与语法建议）", lines=20)
        
    submit_btn.click(fn=check_typos, inputs=file_input, outputs=output_text)
    gr.ClearButton([file_input, output_text], value="🗑️ 清空")

# 5. 启动
demo.launch()