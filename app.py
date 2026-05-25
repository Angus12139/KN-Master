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
   - If there are areas to change, clearly list them as: “第X页：[原文] -> [修改建议]及原因”"""

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
# 3. AI 视觉理解核心 (【深度优化】消灭解析失败)
# ==========================================
def get_image_summary(image_bytes):
    """视觉识别 B 列图片并生成极致精简摘要（智能格式探测 + 3次抗波动重试）"""
    if not image_bytes: return ""
    
    # 【核心修复1】：智能嗅探图片真实格式，拒绝硬编码带来的 400 校验错误
    mime_type = "image/jpeg" # 默认兜底
    if image_bytes.startswith(b'\x89PNG\r\n\x1a\n'):
        mime_type = "image/png"
    elif image_bytes.startswith(b'RIFF') and b'WEBP' in image_bytes[:16]:
        mime_type = "image/webp"
    elif image_bytes.startswith(b'GIF8'):
        mime_type = "image/gif"

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
        
        # 【核心修复2】：内置 3 次自动重试机制，强力抵抗网络抖动与接口频控
        for attempt in range(3):
            try:
                response = client.models.generate_content(
                    model='gemini-2.5-flash',
                    contents=[prompt, img_part]
                )
                result = response.text.strip()
                
                # 物理洗刷废话
                prefixes_to_remove = ["图片展示了", "核心是", "这张图", "核心内容是", "下一页内容是", "主要展示", "总结：", "提示："]
                for prefix in prefixes_to_remove:
                    if result.startswith(prefix):
                        result = result[len(prefix):]
                        
                # 剃光标点
                result = re.sub(r'[^\w\u4e00-\u9fa5]', '', result)
                if len(result) > 10: result = result[:10]
                return result
            except Exception as api_err:
                print(f"Gemini 接口第 {attempt+1} 次尝试失败: {api_err}")
                time.sleep(1.5) # 稍微静置后重试
                
        return "解析失败"
    except Exception as e:
        print(f"图片参数封装失败: {e}")
        return "解析失败"

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
# 5. 🛠 引擎 A：【流式动态反馈版】自动生成表格摘要并回写 D 列
# ==========================================
def generate_summaries_handler(link):
    if not link:
        yield "⚠️ 请先粘贴链接"
        return
    
    # 使用 yield 实现即时视觉反馈
    yield "🔄 [1/4] 正在获取飞书云端安全授权..."
    token = get_feishu_token()
    
    yield "🔄 [2/4] 正在穿透多维表格/知识库节点..."
    ss_token, sheet_id, msg = parse_link(link, token)
    if not ss_token:
        yield f"❌ 连接失败: {msg}"
        return
    
    yield "🔄 [3/4] 正在拉取 A1:Z500 全量数据矩阵..."
    data_url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{ss_token}/values/{sheet_id}!A1:Z500?valueRenderOption=Formula"
    headers = {"Authorization": f"Bearer {token}"}
    raw_data = requests.get(data_url, headers=headers).json().get("data", {}).get("valueRange", {}).get("values", [])
    
    if not raw_data:
        yield "❌ 未能在该工作表中读取到有效数据，请检查文档是否为空或机器人权限。"
        return

    processed_count = 0
    img_col = 1 # B 列
    total_rows = len(raw_data)
    
    yield f"📊 [4/4] 成功加载！共发现 {total_rows} 行数据。开始启动 AI 视觉引擎逐行解析..."
    time.sleep(0.5)

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
            # 动态反馈正在处理哪一行，让用户心中有数
            yield f"⏳ 正在深度解析第 {i+1}/{total_rows} 行的单元格图片..."
            img_bytes = download_fs_media(file_token, token)
            if img_bytes:
                summary = get_image_summary(img_bytes)
                # 只有真正生成了有效文字才写回，防止用“解析失败”四个字污染用户的干净表格
                if summary and summary != "解析失败":
                    update_feishu_cell(ss_token, sheet_id, i, summary, token)
                    processed_count += 1
                    yield f"✨ 进展：第 {i+1} 行图片成功识别为 ➡️【{summary}】"
                else:
                    yield f"⚠️ 警告：第 {i+1} 行图片多轮重试后依然无法识别，已安全跳过。"
            else:
                yield f"❌ 阻碍：第 {i+1} 行图片文件流下载失败，请检查飞书后台‘导出附件’权限。"
        
        # 引入 0.2 秒科学微休眠，优雅规避并发频控
        time.sleep(0.2)
                    
    yield f"🎉 **全篇大功告成！** 成功在 D 列回写了 {processed_count} 条极致精简的下一页提示词。请刷新飞书表格查阅！"

# ==========================================
# 6. 🚀 引擎 B：导出智能 PPT (C列正文，下行D列摘要)
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
                
                if not next_hint or next_hint == "None" or next_hint == "解析失败":
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
    gr.Markdown("# 🚀 AI 智能文档工作站 V5.6 (高阶反馈版)")
    
    with gr.Tabs():
        # --- Tab 1：纠错校对 ---
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

        # --- Tab 2：飞书双引擎 ---
        with gr.TabItem("🎬 飞书一键转 PPT"):
            gr.Markdown("### 双步工作流：1. 读取 B 列图片生成摘要写回 D 列 ➡️ 2. 读取 C 列正文与下页 D 列生成 PPT")
            with gr.Row():
                link_input = gr.Textbox(label="第一步：粘贴飞书链接 (Sheet/Wiki)", placeholder="https://...")
                col_input = gr.Textbox(label="正文所在列 (默认 C 列)", value="C")

            with gr.Row():
                with gr.Column():
                    gr.Markdown("### 🛠 引擎 A：表格内容生产")
                    summary_btn = gr.Button("🤖 识别图片生成提示词 (将结果写入 D 列)", variant="secondary")
                    summary_status = gr.Markdown("状态：等待指令")
                
                with gr.Column():
                    gr.Markdown("### 🚀 引擎 B：PPT 智能导出")
                    export_btn = gr.Button("🔥 立即导出智能预测 PPT", variant="primary")
                    ppt_file = gr.File(label="下载导出的 PPT")
                    export_status = gr.Markdown("状态：等待指令")

            # 注意：此处 generate_summaries_handler 升级为流式生成器，会自动刷新状态框
            summary_btn.click(fn=generate_summaries_handler, inputs=link_input, outputs=summary_status)
            export_btn.click(fn=export_ppt_handler, inputs=[link_input, col_input], outputs=[ppt_file, export_status])

demo.launch()