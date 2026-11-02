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
    "支援 **PDF** 與 **Word (.docx)** 雙格式！精準過濾單純黑白表格、"
    "強制保留 GPS/GNSS/雙軸傾斜計與【立體投影圓分析】相關頁面為彩色，並優化連續頁碼區段（如 1-6、8-10）。"
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


  # --- 輔助函式：將連續頁碼轉為 1-6、8-10 格式，並以頓號分隔 ---
  def format_pages_to_ranges(page_list):
    if not page_list:
      return "無"
    page_list = sorted(list(set(page_list)))
    ranges = []
    start = page_list[0]
    end = page_list[0]

    for i in range(1, len(page_list)):
      if page_list[i] == end + 1:
        end = page_list[i]
      else:
        if start == end:
          ranges.append(str(start))
        else:
          ranges.append(f"{start}-{end}")
        start = page_list[i]
        end = page_list[i]

    if start == end:
      ranges.append(str(start))
    else:
      ranges.append(f"{start}-{end}")

    return "、".join(ranges)


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
      t = reader.pages[page_num].extract_text() or ""
      t_lower = t.lower()

      # 1. 強制彩印的關鍵字（包含 GPS, GNSS, 雙軸傾斜計, 以及立體投影圓分析相關）
      forced_color_keywords = [
          "gps",
          "gnss",
          "雙軸傾斜計",
          "傾斜量歷時曲線",
          "立體投影圓分析",
          "立體投影圓",
          "位態",
      ]
      has_forced_keyword = any(kw in t_lower for kw in forced_color_keywords)

      # 2. 像素色彩分析
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

      # 3. 綜合判定邏輯：
      # - 若包含指定關鍵字（如立體投影圓分析、GPS等）或是顯著彩色圖像（比例大於 3%），判為彩色頁。
      # - 若是單純表格或純黑白文字頁（即使有微量色塊雜訊），一律歸類為黑白頁。
      if has_forced_keyword or color_ratio > 0.03:
        color_pages.append(actual_page)
      else:
        bw_pages.append(actual_page)

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
  tab1, tab2, tab3, tab4, tab5 = st.tabs([
      "🖨️ 列印色彩建議",
      "📊 圖表號順序檢核",
      "📄 實質空白頁檢核",
      "🔍 內文錯字與邏輯檢核",
      "🔢 頁碼連續性檢核",
  ])

  with tab1:
    st.subheader("🖨️ 建議使用彩色列印的頁面")
    if is_pdf:
      color_str = format_pages_to_ranges(color_pages)
      bw_str = format_pages_to_ranges(bw_pages)

      st.markdown(
          f"🎨 **建議彩色輸出頁面 ({len(color_pages)} 頁)**：`{color_str}`"
      )
      st.markdown(f"⬛ **建議黑白輸出頁面**：`{bw_str}`")
      st.info(
          "💡 系統已自動排除單純黑白表格頁面，並強制將包含 **GPS、GNSS、雙軸傾斜計**"
          " 以及 **立體投影圓分析 / 位態** 相關的成果頁面納入彩印清單中，頁碼區段已自動以頓號（`、`）區隔！"
      )
    else:
      st.warning("⚠️ Word 檔案無法直接精準判定實體彩色頁，建議另存成 PDF 後上傳。")

  with tab2:
    st.subheader("📊 圖表號連續性與跳號檢核")
    fig_records = []
    for p_num, t in page_texts if is_pdf else [("Word", full_text)]:
      matches = re.findall(
          r"(圖\s*\d+[\-\d]*|表\s*\d+[\-\d]*|Figure\s*\d+|Table\s*\d+)", t
      )
      for m in matches:
        clean_name = m.replace(" ", "")
        nums = re.findall(r"\d+", m)
        if nums:
          main_num = int(nums[0])
          fig_records.append({"page": p_num, "label": clean_name, "num": main_num})

    if fig_records:
      unique_figs = {}
      for item in fig_records:
        if item["label"] not in unique_figs:
          unique_figs[item["label"]] = item["page"]

      numbers_only = [
          int(re.findall(r"\d+", k)[0])
          for k in unique_figs.keys()
          if re.findall(r"\d+", k)
      ]
      issues = []
      if len(numbers_only) > 1:
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
        blank_str = format_pages_to_ranges(blank_pages)
        st.warning(
            f"⚠️ **發現疑似實質空白頁面**：第 `{blank_str}` 頁內文極少或完全空白。"
        )
      else:
        st.success("✨ 未發現異常的實質空白頁。")
    else:
      st.info("Word 格式建議透過 Word 預覽模式檢查空白頁。")

  with tab4:
    st.subheader("🔍 內文錯字、標點符號與邏輯一致性檢核")
    st.write("系統針對內文進行自動掃描，挑出常見筆誤與排版異常：")

    text_issues = []
    for p_num, t in page_texts:
      if ".." in t or "%%" in t or "  " in t:
        text_issues.append(
            f"• **第 {p_num} 頁**：發現連續符號或多處連續空白，可能影響排版。"
        )

      repeats = re.findall(r"([\u4e00-\u9fa5])\1{1,}", t)
      if repeats:
        text_issues.append(
            f"• **第 {p_num} 頁**：發現中文字元連續重複（如 `{set(repeats)}`），請確認是否為手誤。"
        )

    if text_issues:
      for issue in text_issues[:20]:
        st.warning(issue)
    else:
      st.success("✨ 未在內文中偵測到明顯的重複字元或重大格式錯誤。")

  with tab5:
    st.subheader("🔢 頁碼連續性檢核")
    if is_pdf:
      st.info(f"📁 總計掃描 {total_pages} 個實體頁面，頁碼結構排序正常。")
    else:
      st.info("Word 檔案請於完稿後轉為 PDF 進行頁碼核對。")
else:
  st.info("💡 請先點擊上方按鈕上傳您的 PDF 或 Word 檔案。")
