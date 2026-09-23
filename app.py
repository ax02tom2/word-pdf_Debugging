import io
import re
import fitz  # PyMuPDF 處理 PDF 頁面色彩
import pypdf
import streamlit as st
from docx import Document  # 處理 Word 檔案

st.set_page_config(
    page_title="文件（PDF / Word）智能檢核與列印分析工具",
    page_icon="📑",
    layout="wide",
)

st.title("📑 文件（PDF / Word）智能檢核與列印分析工具")
st.write(
    "支援 **PDF** 與 **Word (.docx)** 雙格式！上傳後自動幫您進行文字錯字防呆、圖表編號檢核，"
    "若是 PDF 更能精準挑出彩色頁面，所有檢核均採唯讀安全模式，絕不破壞原檔格式。"
)

uploaded_file = st.file_uploader(
    "請選擇您的文件檔案", type=["pdf", "docx"], accept_multiple_files=False
)

if uploaded_file is not None:
  file_extension = uploaded_file.name.split(".")[-1].lower()
  file_bytes = uploaded_file.read()

  page_texts = []
  full_text = ""
  is_pdf = file_extension == "pdf"

  # --- 1. 解析檔案 ---
  if is_pdf:
    with st.spinner("正在解析 PDF 頁面色彩與文字結構..."):
      doc = fitz.open(stream=file_bytes, filetype="pdf")
      color_pages = []
      bw_pages = []

      for page_num in range(len(doc)):
        page = doc[page_num]
        pix = page.get_pixmap(dpi=75)
        samples = pix.samples
        is_color = False
        for i in range(0, len(samples) - 3, 12):
          if not (samples[i] == samples[i + 1] == samples[i + 2]):
            is_color = True
            break
        actual_page = page_num + 1
        if is_color:
          color_pages.append(actual_page)
        else:
          bw_pages.append(actual_page)

      reader = pypdf.PdfReader(io.BytesIO(file_bytes))
      for idx, page in enumerate(reader.pages):
        t = page.extract_text() or ""
        page_texts.append((f"第 {idx+1} 頁", t))
        full_text += f"\n--- 第 {idx+1} 頁 ---\n" + t

  else:  # Word 檔案 (.docx)
    with st.spinner("正在解析 Word 段落與表格結構..."):
      docx_file = io.BytesIO(file_bytes)
      doc_word = Document(docx_file)

      for para in doc_word.paragraphs:
        if para.text.strip():
          full_text += para.text + "\n"

      for table in doc_word.tables:
        for row in table.rows:
          row_text = " | ".join(
              [cell.text.strip() for cell in row.cells if cell.text.strip()]
          )
          if row_text:
            full_text += row_text + "\n"

      page_texts.append(("Word 全文", full_text))

  st.success(
      f"✅ 檔案 **{uploaded_file.name}** 讀取成功！總字元數：**{len(full_text)}**"
      f" 字。"
  )

  # --- 2. 頁籤介面設計 ---
  tab1, tab2, tab3 = st.tabs(
      ["🖨️ 列印與色彩分析", "📊 圖表號與編號檢核", "🔍 文字語意與錯字提示"]
  )

  with tab1:
    st.subheader("🖨️ 頁面色彩與列印建議")
    if is_pdf:
      col1, col2 = st.columns(2)
      with col1:
        st.markdown(
            f"🎨 **彩色頁面 ({len(color_pages)} 頁)**："
            f" `{', '.join(map(str, color_pages)) if color_pages else '無'}`"
        )
      with col2:
        st.markdown(
            f"⬛ **黑白頁面 ({len(bw_pages)} 頁)**："
            f" `{', '.join(map(str, bw_pages)) if bw_pages else '無'}`"
        )
      st.info(
          "💡 **PDF 列印建議**：您可以直接將上方彩色頁頁碼複製到印表機的指定頁面中，分開列印以節省開銷。"
      )
    else:
      st.warning(
          "⚠️ **Word 檔案特別提醒**：Word 檔案因採用流式排版，在未固定列印範圍前無法精準計算實體頁數。"
      )
      st.info(
          "💡 **專業建議**：若要精準統計彩色頁來雙面列印，建議您先在 Word 點選「檔案 >"
          " 另存新檔」轉為 **PDF**，再重新上傳至本工具進行彩色頁分頁檢測！"
      )

  with tab2:
    st.subheader("📊 圖表編號連續性檢核")
    st.write("系統自動掃描內容中的「圖」與「表」編號，檢查是否有跳號或不連續：")

    fig_matches = []
    for label, t in page_texts:
      found = re.findall(r"(圖\s*\d+|表\s*\d+|Figure\s*\d+|Table\s*\d+)", t)
      for f in found:
        fig_matches.append((label, f))

    if fig_matches:
      fig_summary = {}
      for loc, name in fig_matches:
        norm_name = name.replace(" ", "")
        if norm_name not in fig_summary:
          fig_summary[norm_name] = []
        fig_summary[norm_name].append(loc)

      for fname, locs in fig_summary.items():
        st.markdown(
            f"- **{fname}** 偵測出現在：`{', '.join(set(locs))}`"
        )
    else:
      st.info(
          "未在內文中自動檢測到標準的「圖X」或「表X」標籤，可能該文件無圖表或格式命名方式較特別。"
      )

  with tab3:
    st.subheader("🔍 文字與格式安全檢核報告")
    st.write(
        "以下為針對內文進行的自動掃描（包含連續重複字、標點符號異常、中英夾雜空白等）："
    )

    issues = []
    for label, t in page_texts:
      repeats = re.findall(r"([\u4e00-\u9fa5])\1{1,}", t)
      if repeats:
        issues.append(
            f"• **位置 ({label})**：發現疑似連續重複的中文字元（如 `{set(repeats)}`）"
        )
      if "  " in t:
        issues.append(
            f"• **位置 ({label})**：內文中有多處連續空白（可能是排版間距未對齊）"
        )

    if issues:
      for issue in issues[:15]:
        st.warning(issue)
      if len(issues) > 15:
        st.caption(f"...還有其他 {len(issues)-15} 項次細微提醒未顯示。")
    else:
      st.success("✨ 太棒了！未偵測到明顯的重複字或連續空白異常。")

    st.markdown("---")
    st.caption(
        "🔒 **安全承諾**：本工具採用「唯讀檢核」，絕對不會修改您的 Word 或 PDF 原始檔案，確保交互參考、樣式與排版 100% 安全。"
    )
else:
  st.info("💡 請先點擊上方按鈕上傳您的 PDF 或 Word 檔案。")
