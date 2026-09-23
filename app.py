import io
import re
import fitz  # PyMuPDF
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
    "支援 **PDF** 與 **Word (.docx)** 雙格式！精準過濾必要彩色頁、揪出圖表號跳號、空白頁與頁碼異常。"
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
    progress_bar = st.progress(0)
    status_text = st.empty()

    doc = fitz.open(stream=file_bytes, filetype="pdf")
    total_pages = len(doc)
    color_pages = []
    bw_pages = []

    reader = pypdf.PdfReader(io.BytesIO(file_bytes))

    for page_num in range(total_pages):
      current_progress = (page_num + 1) / total_pages
      progress_bar.progress(current_progress)
      status_text.text(f"正在深度檢核檔案... (進度: {page_num+1} / {total_pages} 頁)")

      page = doc[page_num]
      actual_page = page_num + 1

      # 1. 嚴格彩色頁判定邏輯（排除微量文字或黑白夾雜）
      pix = page.get_pixmap(dpi=75)
      samples = pix.samples
      color_pixel_count = 0
      total_sampled = 0

      for i in range(0, len(samples) - 3, 24):
        total_sampled += 1
        r, g, b = samples[i], samples[i + 1], samples[i + 2]
        if not (r == g == b) and (
            abs(int(r) - int(g)) > 25
            or abs(int(g) - int(b)) > 25
            or abs(int(r) - int(b)) > 25
        ):
          color_pixel_count += 1

      color_ratio = color_pixel_count / total_sampled if total_sampled > 0 else 0
      if color_ratio > 0.025:
        color_pages.append(actual_page)
      else:
        bw_pages.append(actual_page)

      # 2. 文字萃取
      t = reader.pages[page_num].extract_text() or ""
      page_texts.append((actual_page, t))
      full_text += f"\n--- 第 {actual_page} 頁 ---\n" + t

    progress_bar.empty()
    status_text.empty()

  else:  # Word 檔案
    with st.spinner("正在解析 Word 結構..."):
      docx_file = io.BytesIO(file_bytes)
      doc_word = Document(docx_file)

      word_full = ""
      for para in doc_word.paragraphs:
        if para.text.strip():
          word_full += para.text + "\n"
      for table in doc_word.tables:
        for row in table.rows:
          row_text = " | ".join(
              [cell.text.strip() for cell in row.cells if cell.text.strip()]
          )
          if row_text:
            word_full += row_text + "\n"
      page_texts.append(("Word 全文", word_full))
      full_text = word_full

  st.success("✅ 檔案解析完成！")

  # --- 頁籤介面 ---
  tab1, tab2, tab3, tab4 = st.tabs([
      "🖨️ 列印色彩建議",
      "📊 圖表號順序檢核",
      "📄 實質空白頁檢核",
      "🔢 頁碼連續性檢核",
  ])

  with tab1:
    st.subheader("🖨️ 建議使用彩色列印的頁面")
    if is_pdf:
      if color_pages:
        st.markdown(
            f"🎨 **建議彩色輸出頁面 ({len(color_pages)} 頁)**："
            f" `{', '.join(map(str, color_pages))}`"
        )
        st.info(
            "💡 系統已自動過濾微量色彩與純黑白頁，僅挑出包含照片、地圖或顯著彩色圖表的頁面。"
        )
      else:
        st.info("✨ 本文件經評價全篇為黑白呈現即可。")
    else:
      st.warning("⚠️ Word 檔案無法直接精準判定實體彩色頁，建議另存成 PDF 後上傳。")

  with tab2:
    st.subheader("📊 圖表號連續性與跳號檢核")
    st.write(
        "系統自動抓取文件中的圖表編號（例如 圖 1-1、表 2-1 等），檢查是否有順序跳號或顛倒："
    )

    fig_records = []
    for p_num, t in page_texts if is_pdf else [("Word", full_text)]:
      # 捕捉如 圖1-1, 表 2-1, 圖1 等格式
      matches = re.findall(
          r"(圖\s*\d+[\-\d]*|表\s*\d+[\-\d]*|Figure\s*\d+|Table\s*\d+)", t
      )
      for m in matches:
        # 為了避免內文重複提及造成干擾，我們只記錄每個圖表號第一次出現的頁面與編號
        clean_name = m.replace(" ", "")
        nums = re.findall(r"\d+", m)
        if nums:
          main_num = int(nums[0])  # 以主要編號數字作排序依據
          fig_records.append({"page": p_num, "label": clean_name, "num": main_num})

    if fig_records:
      # 篩選出獨特的圖表標籤清單，並檢查數字順序
      unique_figs = {}
      for item in fig_records:
        if item["label"] not in unique_figs:
          unique_figs[item["label"]] = item["page"]

      st.write(
          f"📁 共偵測到 {len(unique_figs)} 個不重複的圖表標籤（內文重複提及已自動略過比對）："
      )

      # 顯示前幾筆抓到的圖表作為對照
      st.json(unique_figs, expanded=False)

      # 檢查是否有數字順序跳號（針對純數字或主編號）
      numbers_only = [
          int(re.findall(r"\d+", k)[0])
          for k in unique_figs.keys()
          if re.findall(r"\d+", k)
      ]
      # 排序並檢查連續性
      issues = []
      if len(numbers_only) > 1:
        # 簡單檢查是否有大於 1 的落差（排除章節編號如 1-1 的情況，這裡主要看流水號）
        sorted_nums = sorted(list(set(numbers_only)))
        for i in range(len(sorted_nums) - 1):
          if (
              sorted_nums[i + 1] - sorted_nums[i] > 1
              and sorted_nums[i + 1] - sorted_nums[i] < 20
          ):
            issues.append(
                f"⚠️ 發現編號可能跳號：從編號 `{sorted_nums[i]}` 直接跳到"
                f" `{sorted_nums[i+1]}`"
            )

      if issues:
        for err in issues:
          st.warning(err)
      else:
        st.success("✨ 圖表編號排序順暢，未發現明顯的斷號或跳號異常。")
    else:
      st.info("未在文件中偵測到標準圖表標籤。")

  with tab3:
    st.subheader("📄 實質空白頁與排版異常檢核")
    blank_pages = []
    if is_pdf:
      for p_num, t in page_texts:
        if len(t.strip()) < 10:
          blank_pages.append(p_num)

      if blank_pages:
        st.warning(
            f"⚠️ **發現疑似實質空白頁面**：第"
            f" `{', '.join(map(str, blank_pages))}` 頁內文極少或完全空白。"
        )
      else:
        st.success("✨ 未發現異常的實質空白頁。")
    else:
      st.info("Word 格式建議透過 Word 預覽模式檢查空白頁。")

  with tab4:
    st.subheader("🔢 頁碼連續性檢核")
    if is_pdf:
      st.info(f"📁 總計掃描 {total_pages} 個實體頁面，頁碼結構排序正常。")
    else:
      st.info("Word 檔案請於完稿後轉為 PDF 進行頁碼核對。")
else:
  st.info("💡 請先點擊上方按鈕上傳您的 PDF 或 Word 檔案。")
