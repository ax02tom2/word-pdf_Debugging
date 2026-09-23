import io
import re
import fitz  # PyMuPDF 處理 PDF 頁面色彩
import pypdf
import streamlit as st
from docx import Document

st.set_page_config(
    page_title="文件（PDF / Word）智能檢核與列印分析工具",
    page_icon="📑",
    layout="wide",
)

st.title("📑 文件（PDF / Word）智能檢核與列印分析工具")
st.write(
    "支援 **PDF** 與 **Word (.docx)** 雙格式！上傳後自動幫您進行文字錯字防呆、圖表編號檢核，"
    "若是 PDF 更能精準挑出彩色頁面。"
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

  if is_pdf:
    # 建立進度條，解決畫面卡住的問題
    progress_bar = st.progress(0)
    status_text = st.empty()

    doc = fitz.open(stream=file_bytes, filetype="pdf")
    total_pages = len(doc)
    color_pages = []
    bw_pages = []

    reader = pypdf.PdfReader(io.BytesIO(file_bytes))

    for page_num in range(total_pages):
      # 更新進度條與狀態文字
      current_progress = (page_num + 1) / total_pages
      progress_bar.progress(current_progress)
      status_text.text(
          f"正在分析 PDF 頁面色彩與文字... (進度: {page_num+1} /"
          f" {total_pages} 頁)"
      )

      # 1. 色彩快速檢測（降低 DPI 提升速度，避免卡死）
      page = doc[page_num]
      pix = page.get_pixmap(dpi=50)  # 降低至 50 提升效能
      samples = pix.samples
      is_color = False
      # 加大抽樣間距，大幅減少迴圈運算量
      for i in range(0, len(samples) - 3, 36):
        if not (samples[i] == samples[i + 1] == samples[i + 2]):
          is_color = True
          break

      actual_page = page_num + 1
      if is_color:
        color_pages.append(actual_page)
      else:
        bw_pages.append(actual_page)

      # 2. 文字萃取
      t = reader.pages[page_num].extract_text() or ""
      page_texts.append((f"第 {actual_page} 頁", t))
      full_text += f"\n--- 第 {actual_page} 頁 ---\n" + t

    # 清除進度條
    progress_bar.empty()
    status_text.empty()

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
      f"✅ 檔案 **{uploaded_file.name}** 解析完成！總字元數：**{len(full_text)}**"
      f" 字。"
  )

  # --- 頁籤介面設計 ---
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
    else:
      st.warning(
          "⚠️ **Word 檔案提示**：建議另存為 PDF 後再上傳，以精準計算彩色頁。"
      )

  with tab2:
    st.subheader("📊 圖表編號連續性檢核")
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
      st.info("未在內文中自動檢測到標準的「圖X」或「表X」標籤。")

  with tab3:
    st.subheader("🔍 文字與格式安全檢核報告")
    issues = []
    for label, t in page_texts:
      repeats = re.findall(r"([\u4e00-\u9fa5])\1{1,}", t)
      if repeats:
        issues.append(
            f"• **位置 ({label})**：發現連續重複的中文字元（如 `{set(repeats)}`）"
        )
      if "  " in t:
        issues.append(f"• **位置 ({label})**：內文有多處連續空白")

    if issues:
      for issue in issues[:15]:
        st.warning(issue)
    else:
      st.success("✨ 未偵測到明顯的重複字或連續空白異常。")
else:
  st.info("💡 請先點擊上方按鈕上傳您的 PDF 或 Word 檔案。")
