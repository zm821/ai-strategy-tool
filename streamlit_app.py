import os
import json
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Side, Font
from openpyxl.utils import get_column_letter
from openai import OpenAI
import base64
import tempfile

# ===================== 登录 =====================
def check_login():
    if "login_pass" not in st.session_state:
        st.session_state["login_pass"] = False

    if not st.session_state["login_pass"]:
        st.set_page_config(page_title="AI战略分析工具", layout="centered")
        st.title("🔒 AI战略分析工具 管理员登录")
        login_pwd = st.text_input("请输入登录密码", type="password")
        YOUR_LOGIN_PASSWORD = "Ai@2026666"

        if st.button("登录", type="primary"):
            if login_pwd == YOUR_LOGIN_PASSWORD:
                st.session_state["login_pass"] = True
                st.rerun()
            else:
                st.error("密码错误")
        return False
    return True

if not check_login():
    st.stop()

# ===================== API配置 =====================
st.set_page_config(page_title="AI战略分析工具", layout="wide")
st.title("📊 AI战略分析表生成工具（全文档AI提取）")

api_key = st.secrets["API_KEY"]
BASE_URL = "https://dashscope.aliyuncs.com/compatible-mode/v1"
TEXT_MODEL = "qwen3-max"
VISION_MODEL = "qwen-vl-max"

client = OpenAI(api_key=api_key, base_url=BASE_URL)

# ===================== 核心：所有文档统一AI提取 =====================
def extract_content_by_ai(file_bytes, filename):
    """
    所有文档统一走AI提取：
    PDF → 转图片 → 视觉大模型识别
    Word/PPT/Excel → 上传内容 → 大模型标准化提取
    彻底告别本地库解析
    """
    ext = os.path.splitext(filename)[1].lower()

    # 写入临时文件
    with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp:
        tmp.write(file_bytes)
        tmp_path = tmp.name

    content = f"=== 文档：{filename} ===\n"

    try:
        # ===== PDF：视觉大模型识图提取（豆包同款）=====
        if ext == ".pdf":
            import pdfplumber
            with pdfplumber.open(tmp_path) as pdf:
                for idx, page in enumerate(pdf.pages, 1):
                    img = page.to_image()
                    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as img_tmp:
                        img.save(img_tmp.name)
                        with open(img_tmp.name, "rb") as f:
                            b64 = base64.b64encode(f.read()).decode()

                    messages = [
                        {
                            "role": "user",
                            "content": [
                                {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{b64}"}},
                                {"type": "text", "text": "逐字提取图片所有文字，保持阅读顺序，不要总结"}
                            ]
                        }
                    ]

                    res = client.chat.completions.create(model=VISION_MODEL, messages=messages)
                    page_text = res.choices[0].message.content
                    content += f"\n【第{idx}页】\n{page_text}\n"
                    os.unlink(img_tmp.name)

        # ===== Word/PPT/Excel：大模型标准化提取 =====
        else:
            # 先简单读取原始内容，再丢给大模型清洗整理
            raw = ""
            if ext == ".docx":
                from docx import Document
                doc = Document(tmp_path)
                raw = "\n".join(p.text for p in doc.paragraphs)
            elif ext == ".pptx":
                from pptx import Presentation
                prs = Presentation(tmp_path)
                for s in prs.slides:
                    for sh in s.shapes:
                        if hasattr(sh, "text"):
                            raw += sh.text + "\n"
            elif ext in [".xlsx", ".xls"]:
                from openpyxl import load_workbook
                wb = load_workbook(tmp_path, read_only=True)
                for ws in wb:
                    for row in ws.iter_rows(values_only=True):
                        raw += " ".join(str(c) for c in row if c) + "\n"

            # 交给大模型整理成规范文本
            prompt = f"请把下面的文档内容完整提取出来，保持段落清晰、信息不乱：\n{raw}"
            res = client.chat.completions.create(model=TEXT_MODEL, messages=[{"role":"user","content":prompt}])
            content += res.choices[0].message.content

    except Exception as e:
        content += f"\n提取异常：{str(e)}"
    finally:
        os.unlink(tmp_path)

    return content

# ===================== AI分析 =====================
def analyze_with_ai(content, feedback=None):
    system_prompt = """你是专业战略分析师，提取4类内容，每类≥20条，每类条数可以不一致，用英语简洁短句，不允许有中文，严格返回JSON：
{
    "战略目标2030": [],
    "年度目标2026": [],
    "主要改进事项": [],
    "改进指标": []
}"""
    messages = [{"role": "system", "content": system_prompt}]
    for h in st.session_state["current_session"].get("history", []):
        messages.append({"role": h["role"], "content": h["text"]})

    if feedback:
        messages.append({"role":"user","content":f"文档内容：{content}\n修改要求：{feedback}"})
    else:
        messages.append({"role":"user","content":content})

    resp = client.chat.completions.create(model=TEXT_MODEL, messages=messages, response_format={"type":"json_object"})
    return json.loads(resp.choices[0].message.content)

# ===================== Excel导出（不变） =====================
def save_excel(data, base_name="分析结果"):
    out = f"{base_name}_战略表.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "战略规划"
    thin = Side(style='thin', color='000000')
    border = Border(top=thin, bottom=thin, left=thin, right=thin)
    font = Font(name="宋体", size=11)

    CR, CC = 30, 30
    ws.cell(CR, CC, value="年度目标2026")

    # 上：改进事项
    items = data.get("主要改进事项", [])
    r, c = CR - 1, CC
    for x in items[:25]:
        ws.cell(r, c, x).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        r -= 1
    # 左：2026
    items = data.get("年度目标2026", [])
    r, c = CR, CC - 1
    for x in items[:25]:
        ws.cell(r, c, x).alignment = Alignment(textRotation=90, horizontal='center', vertical='center')
        c -= 1
    # 下：2030
    items = data.get("战略目标2030", [])
    r, c = CR + 1, CC
    for x in items[:25]:
        ws.cell(r, c, x).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        r += 1
    # 右：指标
    items = data.get("改进指标", [])
    r, c = CR, CC + 1
    for x in items[:25]:
        ws.cell(r, c, x).alignment = Alignment(textRotation=90, horizontal='center', vertical='center')
        c += 1

    # 清理空行空列
    for r in range(ws.max_row, 0, -1):
        if all(ws.cell(r, c).value is None for c in range(1, ws.max_column+1)):
            ws.delete_rows(r)
    for c in range(ws.max_column, 0, -1):
        if all(ws.cell(r, c).value is None for r in range(1, ws.max_row+1)):
            ws.delete_cols(c)

    # 中心图
    ar, ac = CR, CC
    img_path = "four.png"
    if os.path.exists(img_path):
        from openpyxl.drawing.image import Image
        img = Image(img_path)
        img.width, img.height = 400, 400
        ws.add_image(img, f"{get_column_letter(ac)}{ar}")

    # 样式
    for r in range(1, ws.max_row+1):
        ws.row_dimensions[r].height = 15
    ws.row_dimensions[ar].height = 300
    for c in range(1, ws.max_column+1):
        ws.column_dimensions[get_column_letter(c)].width = 3
    ws.column_dimensions[get_column_letter(ac)].width = 50

    for r in range(1, ws.max_row+1):
        for c in range(1, ws.max_column+1):
            cell = ws.cell(r, c)
            cell.border = border
            cell.font = font

    wb.save(out)
    return out

# ===================== 界面 =====================
if "current_session" not in st.session_state:
    st.session_state["current_session"] = {"history": [], "last_data": None, "original_content": ""}

st.subheader("📁 上传文件（全部由AI识别提取）")
# 🔥 修复这里：补全 label + type 参数
uploaded_files = st.file_uploader("选择文件", type=["docx","pptx","pdf","xlsx","xls"], accept_multiple_files=True)

if st.button("🚀 生成Excel", type="primary"):
    if not uploaded_files:
        st.warning("请上传文件")
    else:
        with st.spinner("AI正在全量提取文档内容..."):
            all_text = ""
            for f in uploaded_files:
                all_text += extract_content_by_ai(f.getbuffer(), f.name) + "\n\n"

            st.session_state["current_session"]["original_content"] = all_text
            result = analyze_with_ai(all_text)
            st.session_state["current_session"]["last_data"] = result

            xlsx = save_excel(result)
            with open(xlsx, "rb") as f:
                st.download_button("📥 下载Excel", f, file_name=xlsx)

# 修改意见
st.subheader("✍️ 输入修改意见重新生成")
feedback = st.text_area("修改要求")
if st.button("发送并重新生成Excel"):
    if not feedback or not st.session_state["current_session"]["last_data"]:
        st.warning("先生成一次或输入意见")
    else:
        with st.spinner("AI重新生成中..."):
            st.session_state["current_session"]["history"].append({"role":"user","text":feedback})
            orig = st.session_state["current_session"]["original_content"]
            result = analyze_with_ai(orig, feedback)
            st.session_state["current_session"]["last_data"] = result

            xlsx = save_excel(result)
            with open(xlsx, "rb") as f:
                st.download_button("📥 下载修改后Excel", f, file_name=xlsx)

with st.expander("📜 历史对话"):
    for item in st.session_state["current_session"]["history"]:
        st.write(f"**你**: {item['text']}")
