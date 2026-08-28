"""
发票/收据整理工具 —— 网页版 (Streamlit + Gemini API)

用户只需要：打开网页 -> 上传照片 -> 点处理 -> 下载zip（里面是重命名后的图片+Excel台账）

部署方式见文末的部署说明，或参考对话中的分步指南。
"""

import io
import re
import json
import base64
import zipfile
import mimetypes
import time
from datetime import datetime

import streamlit as st
import requests
import openpyxl
from openpyxl.styles import Font

# ========== 配置 ==========
MODEL_NAME = "gemini-3.6-flash"   # 如果之后又被下线，改成报错信息里提示的新模型名即可
SUPPORTED_TYPES = ["jpg", "jpeg", "png", "pdf"]
# ===========================

EXTRACTION_PROMPT = """你是一个财务票据识别助手。请仔细查看这张发票/收据，提取以下信息，只返回JSON，不要有任何其他文字、不要markdown代码块标记：

{
  "date": "YYYY-MM-DD 格式的交易日期，看不清就填 unknown",
  "amount": "总金额的数字，只保留数字和小数点，不要货币符号，看不清就填 0",
  "currency": "货币代码，如 USD/CNY/EUR，看不清填 unknown",
  "merchant": "商家/公司名称，看不清填 unknown"
}
"""


def call_gemini(file_bytes: bytes, mime_type: str, api_key: str, max_retries: int = 3) -> dict:
    data = base64.standard_b64encode(file_bytes).decode("utf-8")
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{MODEL_NAME}:generateContent?key={api_key}"
    body = {
        "contents": [
            {
                "parts": [
                    {"inline_data": {"mime_type": mime_type, "data": data}},
                    {"text": EXTRACTION_PROMPT},
                ]
            }
        ],
        "generationConfig": {"response_mime_type": "application/json"},
    }

    last_error = None
    for attempt in range(max_retries):
        try:
            resp = requests.post(url, json=body, timeout=120)
            resp.raise_for_status()
            raw_text = resp.json()["candidates"][0]["content"]["parts"][0]["text"].strip()
            raw_text = re.sub(r"^```json|```$", "", raw_text, flags=re.MULTILINE).strip()
            try:
                return json.loads(raw_text)
            except json.JSONDecodeError:
                return {"date": "unknown", "amount": "0", "currency": "unknown", "merchant": "unknown"}
        except requests.exceptions.HTTPError as e:
            last_error = e
            # 503(服务器暂时过载) 和 429(请求太快) 值得重试；其他错误(如401/404)重试也没用，直接抛出
            if resp.status_code in (503, 429) and attempt < max_retries - 1:
                time.sleep(3 * (attempt + 1))  # 等待时间逐次增加：3秒、6秒...
                continue
            raise

    raise last_error


def sanitize(s) -> str:
    s = str(s).strip().replace(" ", "")
    s = re.sub(r'[\\/:*?"<>|]', "", s)
    return s or "unknown"


def build_filename(date_str, amount_str, merchant, ext, used_names: set) -> str:
    base = f"{sanitize(date_str)}_{sanitize(amount_str)}_{sanitize(merchant)}"
    name = f"{base}{ext}"
    counter = 1
    while name in used_names:
        name = f"{base}({counter}){ext}"
        counter += 1
    used_names.add(name)
    return name


def build_excel(rows) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "报销台账"
    ws.append(["日期", "金额", "货币", "商家", "原文件名", "新文件名", "处理时间"])
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for row in rows:
        ws.append(row)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ---------------- 页面 ----------------

st.set_page_config(page_title="发票整理工具", page_icon="🧾")
st.title("🧾 发票整理工具")
st.caption("上传发票照片，自动识别日期/金额/商家，生成重命名文件 + Excel台账")

# 简单密码保护，防止链接被陌生人乱用消耗你的API额度
correct_password = st.secrets.get("APP_PASSWORD")
if correct_password:
    pw = st.text_input("访问密码", type="password")
    if pw != correct_password:
        st.info("请输入访问密码后继续")
        st.stop()

api_key = st.secrets.get("GEMINI_API_KEY")
if not api_key:
    st.error("服务器没有配置API Key，请联系管理员")
    st.stop()

uploaded_files = st.file_uploader(
    "上传发票照片（支持 jpg / png / pdf，可一次选多张）",
    type=SUPPORTED_TYPES,
    accept_multiple_files=True,
)

if uploaded_files and st.button(f"开始处理 {len(uploaded_files)} 个文件", type="primary"):
    rows = []
    zip_buf = io.BytesIO()
    used_names = set()
    progress = st.progress(0, text="准备开始...")

    with zipfile.ZipFile(zip_buf, "w") as zf:
        for i, uf in enumerate(uploaded_files):
            progress.progress((i) / len(uploaded_files), text=f"正在识别: {uf.name}")

            file_bytes = uf.read()
            mime_type = uf.type or mimetypes.guess_type(uf.name)[0] or "image/jpeg"
            ext = "." + uf.name.split(".")[-1].lower()

            try:
                info = call_gemini(file_bytes, mime_type, api_key)
            except Exception as e:
                st.warning(f"⚠️ {uf.name} 识别失败: {e}")
                info = {"date": "unknown", "amount": "0", "currency": "unknown", "merchant": "unknown"}

            date_str = info.get("date", "unknown")
            amount_str = info.get("amount", "0")
            currency = info.get("currency", "unknown")
            merchant = info.get("merchant", "unknown")

            new_name = build_filename(date_str, amount_str, merchant, ext, used_names)
            zf.writestr(new_name, file_bytes)

            rows.append([date_str, amount_str, currency, merchant, uf.name, new_name,
                         datetime.now().strftime("%Y-%m-%d %H:%M:%S")])

        progress.progress(1.0, text="生成Excel台账...")
        zf.writestr("报销台账.xlsx", build_excel(rows))

    progress.empty()
    st.success(f"✅ 处理完成，共 {len(uploaded_files)} 个文件")

    st.dataframe(
        {"日期": [r[0] for r in rows], "金额": [r[1] for r in rows],
         "货币": [r[2] for r in rows], "商家": [r[3] for r in rows]},
        use_container_width=True,
    )

    st.download_button(
        "⬇️ 下载结果 (zip，含重命名图片+Excel)",
        data=zip_buf.getvalue(),
        file_name=f"发票整理_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
        mime="application/zip",
        type="primary",
    )
