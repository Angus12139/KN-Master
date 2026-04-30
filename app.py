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
# 2. AI 语言处理：高敏纠错/校对/翻译提示词
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

def get_next_slide_hint(image_file_bytes):
    """【V5.5 新增】使用 Gemini 总结下一页图片的演讲核心"""
    if not image_file_bytes: return "（无视觉参考）"
    try:
        img_part = {"mime_type": "image/png", "data": image_file_bytes}
        prompt = "用5个字以内，告诉我这张图片的核心内容是什么？用于演讲提醒（如：销售额对比图、架构图等）"
        response = model.generate_content([prompt, img_part])
        return response.text.strip()
    except:
        return "（图片解析中...）"

# ==========================================
# 3. 飞书 API 深度穿透 (支持图片下载)
# ==========================================
def get_feishu_token():
    url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal"
    r = requests.post(url, json={"app_id": FS_APP_ID, "app_secret": FS_APP_SECRET})
    return r.json().get("tenant_access_token")

def download_feishu_image(file_token, token):
    """【V5.5 新增】从飞书下载单元格中的图片"""
    url = f"https://open.feishu.cn/open-apis/drive/v1/medias/{file_token}/download"
    headers = {"Authorization": f"Bearer {token}"}
    r = requests.get(url, headers=headers)
    return r.content if r.status_code == 200 else None

def fetch_feishu_data(url_link):
    token = get_feishu_token()
    if not token: return None, None, "❌ 授权失败"
    headers = {"Authorization": f"Bearer {token}"}
    
    if "/sheets/" in url_link:
        ss_token = url_link.split("/sheets/")[1].split("?")[0].split("#")[0]
    elif "/wiki/" in url_link:
        wiki_token = url_link.split("/wiki/")[1].split("?")[0].split("#")[0]
        node_res = requests.get(f"https://open.feishu.cn/open-apis/wiki/v2/spaces/get_node?token={wiki_token}", headers=headers).json()
        ss_token = node_res.get("data", {}).get("node", {}).get("obj_token")
    else: return None, None, "❌ 链接不支持"

    meta_res = requests.get(f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{ss_token}/metainfo", headers=headers).json()
    first_sheet_id = meta_res["data"]["sheets"][0]["sheetId"]
    data_url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{ss_token}/values/{first_sheet_id}!A1:Z500?valueRenderOption=Formula"
    r = requests.get(data_url, headers=headers)
    return r.json().get("data", {}).get("valueRange", {}).get("values", []), token, "OK"

# ==========================================
# 4. PPT 高级排版引擎 (样式同步 + 逻辑预判)
# ==========================================
def hex_to_rgb(hex_color):
    hex_color = hex_color.lstrip('#')
    return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))

def link_to_pptx_full(link, col_letter):
    if not link: return None, "⚠️ 请输入链接"
    raw_data, fs_token, msg = fetch_feishu_data(link)
    if raw_data is None: return None, msg
    
    col_index = ord(col_letter.upper()) - ord('A')
    img_col_index = ord('B') - ord('A') # 图片固定在 B 列
    
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5) # 16:9
    
    valid_rows = [r for r in raw_data if len(r) > col_index]
    count = 0

    for i in range(len(valid_rows)):
        row = valid_rows[i]
        content_obj = row[col_index]
        
        # --- A. 处理正文文本与样式 ---
        full_text = ""
        segments = []
        if isinstance(content_obj, list):
            segments = content_obj
            for s in segments: full_text += s.get('text', '')
        else:
            full_text = str(content_obj).strip()
            if full_text in ["None", "", "nan"]: continue
            segments = [{'text': full_text, 'segmentStyle': {'foreColor': '#FFFFFF'}}]
        
        if not full_text.strip(): continue

        # --- B. 智能获取下一页提示 (分析下一行 B 列图片) ---
        next_hint = "（演讲结束）"
        if i + 1 < len(valid_rows):
            next_row = valid_rows[i+1]
            if len(next_row) > img_col_index:
                img_cell = next_row[img_col_index]
                if isinstance(img_cell, list) and len(img_cell) > 0 and 'fileToken' in img_cell[0]:
                    img_bytes = download_feishu_image(img_cell[0]['fileToken'], fs_token)
                    next_hint = get_next_slide_hint(img_bytes)

        # --- C. 绘图渲染 ---
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = RGBColor(0, 0, 0)
        
        # 1. 主题词文本框 (左对齐)
        margin = Inches(0.8)
        txBox = slide.shapes.add_textbox(margin, margin, prs.slide_width - margin*2, prs.slide_height - Inches(2))
        tf = txBox.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = 1 
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        
        # 动态字号
        text_len = len(full_text)
        font_size = 85 if text_len <= 15 else 65 if text_len <= 40 else 45 if text_len <= 100 else 32

        for seg in segments:
            text_part = seg.get('text', '')
            run = p.add_run()
            run.text = text_part
            run.font.name = '微软雅黑'
            run.font.size = Pt(font_size)
            run.font.bold = seg.get('segmentStyle', {}).get('bold', True)
            
            color_hex = seg.get('segmentStyle', {}).get('foreColor', '#FFFFFF')
            if color_hex.upper() in ["#000000", "#121212"]:
                run.font.color.rgb = RGBColor(255, 255, 255)
            else:
                try: run.font.color.rgb = RGBColor(*hex_to_rgb(color_hex))
                except: run.font.color.rgb = RGBColor(255, 255, 255)

        # 2. 【V5.5 新增】红色提示框 (左下角)
        box_w, box_h = Inches(5.0), Inches(0.6)
        left, top = margin, prs.slide_height - Inches(1.2)
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, box_w, box_h)
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(255, 0, 0) # 纯红背景
        shape.line.fill.background() # 无边框
        
        tf_hint = shape.text_frame
        tf_hint.vertical_anchor = 1
        p_hint = tf_hint.paragraphs[0]
        p_hint.alignment = PP_ALIGN.LEFT
        run_hint = p_hint.add_run()
        run_hint.text = f" 下一页预告：{next_hint}"
        run_hint.font.name = '微软雅黑'
        run_hint.font.size = Pt(20)
        run_hint.font.bold = True
        run_hint.font.color.rgb = RGBColor(255, 255, 255) # 白色文字
        
        count += 1

    output_name = "V5.5_智能样式同步提词器.pptx"
    prs.save(output_name)
    return output_name, f"✅ 已成功同步 {count} 页带样式及预测的提词！"

# ==========================================
# 5. 整合 UI 界面
# ==========================================
with gr.Blocks(title="AI 智能工作站 V5.5") as demo:
    gr.Markdown("# 🚀 AI 智能文档工作站 V5.5")
    
    with gr.Tabs():
        # 标签页一：AI 语言处理 (保留纠错、翻译、校对功能)
        with gr.TabItem("📝 AI 语言处理"):
            with gr.Row():
                with gr.Column():
                    ai_file = gr.File(label="上传 PDF 或图片")
                    with gr.Row():
                        btn_typo = gr.Button("🇨🇳 逐页纠错", variant="primary")
                        btn_proof = gr.Button("⚖️ 双语校对", variant="secondary")
                        btn_trans = gr.Button("🌍 翻译结果", variant="secondary")
                ai_output = gr.Textbox(label="AI 处理结果", lines=20)
            
            btn_typo.click(fn=lambda f: process_ai_task(f, prompt_typo), inputs=ai_file, outputs=ai_output)
            btn_proof.click(fn=lambda f: process_ai_task(f, prompt_proofread), inputs=ai_file, outputs=ai_output)
            btn_trans.click(fn=lambda f: process_ai_task(f, prompt_translate), inputs=ai_file, outputs=ai_output)

        # 标签页二：飞书转 PPT (V5.5 样式同步 + 视觉预测)
        with gr.TabItem("🎬 飞书一键转 PPT"):
            gr.Markdown("### 功能：C列题词 + B列预测 + 样式同步 + 左对齐")
            with gr.Row():
                with gr.Column():
                    link_input = gr.Textbox(label="飞书表格/知识库链接")
                    col_input = gr.Textbox(label="正文列 (如 C)", value="C")
                    ppt_btn = gr.Button("🚀 立即生成 PPT", variant="primary")
                with gr.Column():
                    ppt_download = gr.File(label="下载排版好的 PPT")
                    ppt_status = gr.Markdown("状态：就绪")
            
            ppt_btn.click(fn=link_to_pptx_full, inputs=[link_input, col_input], outputs=[ppt_download, ppt_status])

demo.launch()