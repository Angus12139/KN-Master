import gradio as gr
from google import genai
from google.genai import types
import os
import requests
import json
import io
import re
import time

# ==========================================
# 1. 基础配置
# ==========================================
client = genai.Client(api_key=os.environ.get("GEMINI_API_KEY"))
FS_APP_ID = os.environ.get("FEISHU_APP_ID")
FS_APP_SECRET = os.environ.get("FEISHU_APP_SECRET")

# ==========================================
# 2. AI 语言处理 (Tab 1 核心功能全面保留)
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
        if file_name.endswith('.pdf'): mime_type = 'application/pdf'
        elif file_name.endswith(('.jpg', '.jpeg')): mime_type = 'image/jpeg'
        else: mime_type = 'image/png'
            
        with open(file_obj.name, 'rb') as f:
            file_bytes = f.read()
            
        file_part = types.Part.from_bytes(data=file_bytes, mime_type=mime_type)
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=[file_part, prompt_text]
        )
        return response.text
    except Exception as e:
        return f"❌ 错误: {str(e)}"

# ==========================================
# 3. AI 视觉理解核心 (【深度优化】带详细报错跟踪)
# ==========================================
def get_image_summary(image_bytes):
    """视觉识别 B 列图片并生成极致精简摘要（带防封锁控速机制）"""
    if not image_bytes: return "ERROR: 空图片流"
    
    mime_type = "image/jpeg"
    if image_bytes.startswith(b'\x89PNG\r\n\x1a\n'): mime_type = "image/png"
    elif image_bytes.startswith(b'RIFF') and b'WEBP' in image_bytes[:16]: mime_type = "image/webp"
    elif image_bytes.startswith(b'GIF8'): mime_type = "image/gif"

    last_error = ""
    try:
        img_part = types.Part.from_bytes(data=image_bytes, mime_type=mime_type)
        prompt = """
        这是演讲幻灯片的下一页图片，请提取它的核心演讲主题，用于提词器。
        必须严格遵守以下 3 条铁律：
        1. 字数绝对控制在 10 个字以内！
        2. 绝对不要使用任何标点符号（包括句号、逗号、冒号等）！
        3. 直接输出核心词，绝不能出现“这张图片展示了”、“核心内容是”等废话。
        正确示例：年度销售数据盘点
        """
        
        # 3次重试抗波动机制，每次失败休息更长时间
        for attempt in range(3):
            try:
                response = client.models.generate_content(
                    model='gemini-2.5-flash',
                    contents=[prompt, img_part]
                )
                result = response.text.strip()
                
                # 物理洗刷废话
                for prefix in ["图片展示了", "核心是", "这张图", "核心内容是", "下一页内容是", "主要展示", "总结：", "提示："]:
                    if result.startswith(prefix): result = result[len(prefix):]
                        
                # 剃光标点
                result = re.sub(r'[^\w\u4e00-\u9fa5]', '', result)
                if len(result) > 10: result = result[:10]
                
                # 防止正则洗完后变成空字符串
                if not result: return "无有效文本"
                
                return result
            except Exception as api_err:
                last_error = str(api_err)
                time.sleep(3) # 失败后冷却 3 秒
                
        return f"ERROR: API超载或报错 ({last_error})"
    except Exception as e:
        return f"ERROR: 数据封装失败 ({e})"

# ==========================================
# 4. 飞书数据底座
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
# 5. 🛠 引擎 A：【滚动日志 + 控速版】自动生成表格摘要并回写
# ==========================================
def generate_summaries_handler(link):
    # 建立一个日记本，记录每一步的操作，让用户看清楚
    logs = []
    def update_ui(msg):
        logs.append(msg)
        # 只显示最后 12 条日志，防止界面过长，用换行符连接
        return "\n\n".join(logs[-12:])
        
    if not link:
        yield update_ui("⚠️ 错误：请先粘贴链接！")
        return
    
    yield update_ui("🔄 [步骤 1] 正在获取飞书云端安全授权...")
    token = get_feishu_token()
    
    yield update_ui("🔄 [步骤 2] 正在穿透多维表格节点...")
    ss_token, sheet_id, msg = parse_link(link, token)
    if not ss_token:
        yield update_ui(f"❌ 连接失败: {msg}")
        return
    
    yield update_ui("🔄 [步骤 3] 正在拉取全量数据...")
    data_url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{ss_token}/values/{sheet_id}!A1:Z500?valueRenderOption=Formula"
    headers = {"Authorization": f"Bearer {token}"}
    raw_data = requests.get(data_url, headers=headers).json().get("data", {}).get("valueRange", {}).get("values", [])
    
    if not raw_data:
        yield update_ui("❌ 未读取到有效数据。")
        return

    processed_count = 0
    img_col = 1 # B 列
    total_rows = len(raw_data)
    
    yield update_ui(f"✅ 成功加载！共发现 {total_rows} 行数据。准备启动 AI 解析...")
    time.sleep(1)

    for i, row in enumerate(raw_data):
        if len(row) <= img_col: continue
        cell_data = row[img_col]
        file_token = None
        
        # 深度检索 Token
        def find_token(data):
            if isinstance(data, dict):
                for key in ['fileToken', 'imageToken', 'token', 'file_token']:
                    if data.get(key): return data.get(key)
                for val in data.values():
                    res = find_token(val)
                    if res: return res
            elif isinstance(data, list):
                for item in data:
                    res = find_token(item)
                    if res: return res
            return None

        file_token = find_token(cell_data)
        
        if file_token:
            yield update_ui(f"⏳ 正在处理第 {i+1} 行图片...")
            img_bytes = download_fs_media(file_token, token)
            
            if img_bytes:
                summary = get_image_summary(img_bytes)
                
                if summary.startswith("ERROR"):
                    yield update_ui(f"⚠️ 第 {i+1} 行跳过：{summary}")
                else:
                    update_feishu_cell(ss_token, sheet_id, i, summary, token)
                    processed_count += 1
                    yield update_ui(f"✨ 第 {i+1} 行成功写入 ➡️ 【{summary}】")
                    
                    # 【核心修复】：成功后强制休息 4 秒，防止触发 Google 的 15 RPM 频控报错
                    yield update_ui(f"⏸️ 为防止接口封锁，安全冷却 4 秒钟...")
                    time.sleep(4)
            else:
                yield update_ui(f"❌ 第 {i+1} 行：图片下载失败，请检查飞书附件权限。")
                    
    yield update_ui(f"🎉 大功告成！全篇共在 D 列回写了 {processed_count} 条精简摘要。")

# ==========================================
# 6. 🚀 引擎 B：导出智能 PPT
# ==========================================
def export_ppt_handler(link, col_letter):
    if not link: return None, "⚠️ 请先粘贴链接"
    token = get_feishu_token()
    ss_token, sheet_id, msg = parse_link(link, token)
    if not ss_token: return None, msg
    
    headers = {"Authorization": f"Bearer {token}"}
    data_url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{ss_token}/values/{sheet_id}!A1:Z500?valueRenderOption=Formula"
    raw_data = requests.get(data_url, headers=headers).json().get("data", {}).get("valueRange", {}).get("values", [])
    
    col_idx = ord(col_letter.upper()) - ord('A')
    hint_col_idx = ord('D') - ord('A')
    
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    
    valid_rows = [r for r in raw_data if len(r) > col_idx]
    
    for i in range(len(valid_rows)):
        row = valid_rows[i]
        
        # A. 正文处理 (C列)
        content_obj = row[col_idx]
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

        # B. 查找下一页提示 (下一行的 D 列)
        next_hint = "演讲结束"
        if i + 1 < len(valid_rows):
            next_row = valid_rows[i+1]
            if len(next_row) > hint_col_idx:
                hint_val = next_row[hint_col_idx]
                if isinstance(hint_val, list):
                    next_hint = "".join([s.get('text','') for s in hint_val]).strip()
                else:
                    next_hint = str(hint_val).strip()
                
                if not next_hint or next_hint == "None":
                    next_hint = "演讲结束"

        # C. 渲染 Slide
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = RGBColor(0, 0, 0)
        
        margin = Inches(0.8)
        txBox = slide.shapes.add_textbox(margin, margin, prs.slide_width - margin*2, prs.slide_height - Inches(2))
        tf = txBox.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = 1 
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        
        text_len = len(full_text)
        font_size = 80 if text_len <= 20 else 60 if text_len <= 60 else 40
        
        for seg in segments:
            run = p.add_run()
            run.text = seg.get('text', '')
            run.font.name, run.font.size = '微软雅黑', Pt(font_size)
            run.font.bold = seg.get('segmentStyle', {}).get('bold', True)
            c = seg.get('segmentStyle', {}).get('foreColor', '#FFFFFF')
            if c.upper() in ["#000000", "#121212"]: run.font.color.rgb = RGBColor(255, 255, 255)
            else:
                try: 
                    hex_c = c.lstrip('#')
                    run.font.color.rgb = RGBColor(*(int(hex_c[k:k+2], 16) for k in (0, 2, 4)))
                except: run.font.color.rgb = RGBColor(255, 255, 255)

        # D. 左下角红色预告框
        box_w, box_h = Inches(5.0), Inches(0.6)
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, margin, prs.slide_height - Inches(1.2), box_w, box_h)
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(225, 29, 72)
        shape.line.fill.background()
        
        tf_h = shape.text_frame
        tf_h.vertical_anchor = 1
        ph = tf_h.paragraphs[0]
        ph.alignment = PP_ALIGN.LEFT
        rh = ph.add_run()
        rh.text = f" 下一页预告：{next_hint}"
        rh.font.name, rh.font.size, rh.font.bold = '微软雅黑', Pt(20), True
        rh.font.color.rgb = RGBColor(255, 255, 255)

    out = "V5.6_智能预测提词器.pptx"
    prs.save(out)
    return out, f"✅ 已成功导出 {len(valid_rows)} 页带下一页预告的 PPT！"

# ==========================================
# 7. UI 界面整合
# ==========================================
with gr.Blocks(title="AI 智能文档工作站 V5.6") as demo:
    gr.Markdown("# 🚀 AI 智能文档工作站 V5.6 (安全滚动日志版)")
    
    with gr.Tabs():
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

        with gr.TabItem("🎬 飞书一键转 PPT"):
            gr.Markdown("### 双步工作流：1. 读取 B 列图片生成摘要写回 D 列 ➡️ 2. 读取 C 列正文与下页 D 列生成 PPT")
            with gr.Row():
                link_input = gr.Textbox(label="第一步：粘贴飞书链接 (Sheet/Wiki)", placeholder="https://...")
                col_input = gr.Textbox(label="正文所在列 (默认 C 列)", value="C")

            with gr.Row():
                with gr.Column():
                    gr.Markdown("### 🛠 引擎 A：表格内容生产")
                    summary_btn = gr.Button("🤖 识别图片生成提示词 (自动写入 D 列)", variant="secondary")
                    # 将状态提示框改为多行文本显示，更直观
                    summary_status = gr.Markdown("状态：等待指令（运行时将显示滚动日志）")
                
                with gr.Column():
                    gr.Markdown("### 🚀 引擎 B：PPT 智能导出")
                    export_btn = gr.Button("🔥 立即导出智能预测 PPT", variant="primary")
                    ppt_file = gr.File(label="下载导出的 PPT")
                    export_status = gr.Markdown("状态：等待指令")

            summary_btn.click(fn=generate_summaries_handler, inputs=link_input, outputs=summary_status)
            export_btn.click(fn=export_ppt_handler, inputs=[link_input, col_input], outputs=[ppt_file, export_status])

demo.launch()