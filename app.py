import gradio as gr
import google.generativeai as genai
import os  # 新增：让程序能读取系统的“保险箱”

# 1. 从云端保险箱中读取你的 API 通行证
# 注意：这里不要再写你真实的 AIza... 密钥了！
my_api_key = os.environ.get("GEMINI_API_KEY") 
genai.configure(api_key=my_api_key)
model = genai.GenerativeModel('gemini-2.5-flash')

# 2. 最高敏感度的提示词
sensitive_prompt = """
你是一个极其敏锐的中文错别字和语病检查专家。
请检查用户提供的文档或图片中的错别字。

【重要原则：全面雷达扫描，绝不漏报！】
1. 哪怕你只有 10% 的把握觉得某个词、某句话不对劲，也请务必指出来！
2. 无论是明显的错别字、标点错误、语句不顺，还是专业术语可能有风险的用法，统统列出来。
3. 请在反馈时标注你的“怀疑程度”。
4. 请清晰地列出“原词/原句”、“修改建议”以及“你的怀疑理由”。
"""

# 3. 核心工作函数
def check_typos(file_obj):
    if file_obj is None:
        return "请先上传一个文件哦！"
    
    file_path = file_obj.name 

    try:
        gemini_file = genai.upload_file(file_path)
        response = model.generate_content([sensitive_prompt, gemini_file])
        return response.text
    except Exception as e:
        return f"哎呀，处理时出错了: {e}"

# 4. 界面搭建
interface = gr.Interface(
    fn=check_typos,
    inputs=gr.File(label="📂 上传文档 (支持 PDF / 图片)"), 
    outputs=gr.Textbox(label="AI 深度检查结果", lines=15), 
    title="⚡ 极速全覆盖 AI 错别字雷达",                    
    description="搭载原生高速上传通道！自动免疫 PDF 乱码问题，最高敏感度全面扫描！",
    submit_btn="🔍 闪电扫描",       
    clear_btn="🗑️ 清空",        
    flagging_mode="never"         
)

# 5. 启动软件！
interface.launch()