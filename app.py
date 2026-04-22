import gradio as gr
import google.generativeai as genai
import os

# 1. 配置 AI 通行证
my_api_key = os.environ.get("GEMINI_API_KEY") 
genai.configure(api_key=my_api_key)
model = genai.GenerativeModel('gemini-2.5-flash')

# ==========================================
# 2. 为三个不同的按钮定制专属“大脑指令”
# ==========================================

# 指令 A：纯粹的错别字检查 (保留你的高敏感度雷达要求)
prompt_typo = """
你的唯一任务是：检查文档或图片中的【中文错别字和语病】。
1. 哪怕只有 10% 的把握觉得某个词、某句话不对劲，也请务必指出来！
2. 请清晰列出：“原词/原句” -> “修改建议” -> “怀疑理由”。
3. 请忽略所有的英文内容，不要翻译，只专注于中文纠错。
"""

# 指令 B：纯粹的中英文对照检查
prompt_proofread = """
你的唯一任务是：做专业的【中英双语校对】。
1. 请对比页面上的中文和英文，检查翻译是否准确、语义是否对齐。指出漏译、误译或专业术语不一致的地方。
2. 检查英文部分的拼写 (Spelling) 和语法 (Grammar) 错误。
3. 如果页面上只有中文或只有英文，请提示用户：“未检测到双语对照内容”。
4. 请用列表输出：“原文” -> “当前翻译/表达” -> “校对修改建议”。
"""

# 指令 C：全新的中译英翻译功能
prompt_translate = """
你的唯一任务是：做极其专业的【中译英翻译】。
1. 请提取文档或图片中的所有中文内容。
2. 将其翻译为极其地道、符合商务/专业规范的英文。
3. 保持原有的段落排版格式。
4. 不要输出任何解释性的废话，直接输出翻译好的纯英文结果。
"""

# ==========================================
# 3. 核心处理引擎
# ==========================================
def process_task(file_obj, prompt_text):
    if file_obj is None:
        return "⚠️ 请先上传要处理的 PDF 或图片哦！"
    try:
        # 上传文件到高速通道
        gemini_file = genai.upload_file(file_obj.name)
        # 将特定的指令和文件一起发给 AI
        response = model.generate_content([prompt_text, gemini_file])
        return response.text
    except Exception as e:
        return f"哎呀，处理时出错了: {str(e)}"

# 为了配合三个按钮，我们做三个小助手函数来转交任务
def run_typo(file): return process_task(file, prompt_typo)
def run_proofread(file): return process_task(file, prompt_proofread)
def run_translate(file): return process_task(file, prompt_translate)

# ==========================================
# 4. 搭积木：构建高级分屏界面
# ==========================================
with gr.Blocks(title="AI 智能文档工作站") as demo:
    gr.Markdown("# 🚀 AI 智能文档工作站 V3.0")
    gr.Markdown("请先上传文件，然后点击下方对应的按钮执行您需要的专属任务。")
    
    with gr.Row():
        # 左侧：文件上传区
        with gr.Column(scale=1):
            file_input = gr.File(label="📂 1. 上传 PDF 或图片")
            
            # 把三个按钮并排放在文件上传下方
            gr.Markdown("### 2. 选择要执行的任务")
            with gr.Row():
                btn_typo = gr.Button("🇨🇳 错别字检查", variant="primary")
                btn_proofread = gr.Button("⚖️ 中英文校对", variant="secondary")
                btn_translate = gr.Button("🌍 中翻英翻译", variant="secondary")
                
            clear_btn = gr.ClearButton(value="🗑️ 清空重来")

        # 右侧：结果显示区
        with gr.Column(scale=1):
            output_text = gr.Textbox(label="📝 3. AI 处理结果", lines=20)
            
    # 把按钮和对应的功能绑定起来！
    btn_typo.click(fn=run_typo, inputs=file_input, outputs=output_text)
    btn_proofread.click(fn=run_proofread, inputs=file_input, outputs=output_text)
    btn_translate.click(fn=run_translate, inputs=file_input, outputs=output_text)
    
    # 让清空按钮能同时清空上传框和结果框
    clear_btn.add([file_input, output_text])

# 启动！
demo.launch()