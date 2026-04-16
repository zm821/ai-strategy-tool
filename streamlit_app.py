import os
import json
import streamlit as st
from docx import Document
from pptx import Presentation
import PyPDF2
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Side, Font
from openpyxl.utils import get_column_letter
from openai import OpenAI
import datetime


# ===================== 1. 登录密码保护（第一道锁） =====================
def check_login():
    if "login_pass" not in st.session_state:
        st.session_state["login_pass"] = False

    if not st.session_state["login_pass"]:
        st.set_page_config(page_title="AI战略分析工具", layout="centered")
        st.title("🔒 AI战略分析工具 用户登录")
        login_pwd = st.text_input("请输入登录密码", type="password")

        # 你自己设置的登录密码（可修改）
        YOUR_LOGIN_PASSWORD = "Ai@2026666"

        if st.button("登录", type="primary"):
            if login_pwd == YOUR_LOGIN_PASSWORD:
                st.session_state["login_pass"] = True
                st.rerun()
            else:
                st.error("密码错误，无权访问")
        return False
    return True


# 未登录直接拦截
if not check_login():
    st.stop()

# ===================== 2. API密钥配置（第二道锁） =====================
st.set_page_config(page_title="AI战略分析工具", layout="wide")
st.title("📊 AI战略分析表生成工具")

# 优先从Streamlit Secrets读取密钥（部署用），也支持手动输入（本地用）
# with st.sidebar:
#     st.header("⚙️ API配置")
#     # 部署时把密钥填在Streamlit Secrets里，这里自动读取，不用手动输
#     default_key = st.secrets.get("API_KEY", "")
#     api_key = st.text_input("通义千问API密钥", value=default_key, type="password")
#     if not api_key:
#         st.warning("请输入API密钥后使用")
#         st.stop()
api_key = st.secrets["API_KEY"]

BASE_URL = "https://dashscope.aliyuncs.com/compatible-mode/v1"
MODEL_NAME = "qwen3-max"


# ===================== 3. 文件读取功能 =====================
def read_file(file_bytes, filename):
    ext = os.path.splitext(filename)[1].lower()
    content = ""
    try:
        if ext == ".docx":
            doc = Document(file_bytes)
            content = "\n".join([p.text for p in doc.paragraphs])
        elif ext == ".pptx":
            prs = Presentation(file_bytes)
            for slide in prs.slides:
                for shape in slide.shapes:
                    if hasattr(shape, "text"):
                        content += shape.text + "\n"
        elif ext == ".pdf":
            reader = PyPDF2.PdfReader(file_bytes)
            for page in reader.pages:
                page_text = page.extract_text()
                if page_text:
                    content += page_text + "\n"
        elif ext in [".xlsx", ".xls"]:
            from openpyxl import load_workbook
            wb = load_workbook(file_bytes, read_only=True)
            for sheet in wb:
                for row in sheet.iter_rows(values_only=True):
                    row_str = " ".join([str(cell) for cell in row if cell is not None])
                    if row_str:
                        content += row_str + "\n"
    except Exception as e:
        st.error(f"文件{filename}读取失败：{str(e)}")
    return content


# ===================== 4. AI分析功能 =====================
def analyze_with_ai(content, feedback=None):
    system_prompt = """你是专业战略分析师，提取4类内容，每类≥20条，每类条数可以不一致，用英语简洁短句，不允许有中文，严格返回JSON：
{
    "战略目标2030": [],
    "年度目标2026": [],
    "主要改进事项": [],
    "改进指标": []
}"""
    messages = [{"role": "system", "content": system_prompt}]
    # 加载历史对话
    hist = st.session_state.get("current_session", {}).get("history", [])
    for h in hist:
        messages.append({"role": h["role"], "content": h["text"]})
    # 拼接当前输入
    if feedback:
        messages.append({"role": "user", "content": feedback})
    else:
        messages.append({"role": "user", "content": content})

    client = OpenAI(api_key=api_key, base_url=BASE_URL)
    resp = client.chat.completions.create(
        model=MODEL_NAME,
        messages=messages,
        response_format={"type": "json_object"}
    )
    return json.loads(resp.choices[0].message.content)


# ===================== 5. Excel生成功能 =====================
def save_excel(data, base_name="分析结果"):
    out = f"{base_name}_战略表.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "战略规划"
    thin = Side(style='thin', color='000000')
    border = Border(top=thin, bottom=thin, left=thin, right=thin)
    font = Font(name="宋体", size=11)

    # 中心单元格
    CR, CC = 30, 30
    ws.cell(CR, CC, value="年度目标2026")

    # 填充四个方向的内容
    # 上：主要改进事项
    items = data.get("主要改进事项", [])
    r, c = CR - 1, CC
    for x in items[:25]:
        ws.cell(r, c, x).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        r -= 1
    # 左：年度目标2026
    items = data.get("年度目标2026", [])
    r, c = CR, CC - 1
    for x in items[:25]:
        ws.cell(r, c, x).alignment = Alignment(textRotation=90, horizontal='center', vertical='center')
        c -= 1
    # 下：战略目标2030
    items = data.get("战略目标2030", [])
    r, c = CR + 1, CC
    for x in items[:25]:
        ws.cell(r, c, x).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        r += 1
    # 右：改进指标
    items = data.get("改进指标", [])
    r, c = CR, CC + 1
    for x in items[:25]:
        ws.cell(r, c, x).alignment = Alignment(textRotation=90, horizontal='center', vertical='center')
        c += 1

    # 清理空行空列
    for r in range(ws.max_row, 0, -1):
        if all(ws.cell(r, c).value is None for c in range(1, ws.max_column + 1)):
            ws.delete_rows(r)
    for c in range(ws.max_column, 0, -1):
        if all(ws.cell(r, c).value is None for r in range(1, ws.max_row + 1)):
            ws.delete_cols(c)

    # 定位中心单元格
    ar, ac = CR, CC
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            if ws.cell(r, c).value == "年度目标2026":
                ar, ac = r, c
                break

    # 插入中心图片four.png
    img_path = "four.png"
    if os.path.exists(img_path):
        from openpyxl.drawing.image import Image
        img = Image(img_path)
        img.width = 400
        img.height = 400
        col_letter = get_column_letter(ac)
        ws.add_image(img, f"{col_letter}{ar}")

    # 设置行高列宽
    for r in range(1, ws.max_row + 1):
        ws.row_dimensions[r].height = 15
    ws.row_dimensions[ar].height = 300
    for c in range(1, ws.max_column + 1):
        ws.column_dimensions[get_column_letter(c)].width = 3
    ws.column_dimensions[get_column_letter(ac)].width = 50

    # 统一边框字体
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(r, c)
            cell.border = border
            cell.font = font

    wb.save(out)
    return out


# ===================== 6. 主界面 =====================
# 初始化会话状态
if "current_session" not in st.session_state:
    st.session_state["current_session"] = {"history": [], "last_data": None}

# 文件上传区
st.subheader("📁 上传文件（支持多选：Word/PPT/PDF/Excel）")
uploaded_files = st.file_uploader(
    "选择文件",
    type=["docx", "pptx", "pdf", "xlsx", "xls"],
    accept_multiple_files=True
)

# 生成按钮
if st.button("🚀 生成Excel", type="primary"):
    if not uploaded_files:
        st.warning("请先上传文件！")
    else:
        with st.spinner("AI正在分析文档..."):
            # 合并所有文件内容
            all_content = ""
            for f in uploaded_files:
                all_content += f"\n===== 文件：{f.name} =====\n"
                all_content += read_file(f, f.name)
            # AI分析
            data = analyze_with_ai(all_content)
            st.session_state["current_session"]["last_data"] = data
            # 生成Excel
            out_path = save_excel(data)
            # 下载按钮
            with open(out_path, "rb") as f:
                st.download_button(
                    label="📥 下载生成的Excel",
                    data=f,
                    file_name=out_path,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

# 修改意见区
st.subheader("✍️ 修改意见（重新生成）")
feedback = st.text_area("输入你的修改要求，AI会基于上一版调整")
if st.button("发送并重新生成Excel"):
    if not feedback:
        st.warning("请输入修改意见！")
    elif not st.session_state["current_session"]["last_data"]:
        st.warning("请先生成一次Excel！")
    else:
        with st.spinner("AI正在修改并重新生成..."):
            # 记录历史对话
            st.session_state["current_session"]["history"].append({"role": "user", "text": feedback})
            # AI重新生成
            data = analyze_with_ai("", feedback=feedback)
            st.session_state["current_session"]["last_data"] = data
            # 生成新Excel
            out_path = save_excel(data)
            # 下载按钮
            with open(out_path, "rb") as f:
                st.download_button(
                    label="📥 下载修改后的Excel",
                    data=f,
                    file_name=out_path,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

# 历史对话区
with st.expander("📜 查看历史对话记录"):
    for item in st.session_state["current_session"]["history"]:
        st.write(f"**你**：{item['text']}")
