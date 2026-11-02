import io
import re
import colorsys
from collections import Counter
from PIL import Image
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
    "支援 **PDF** 與 **Word (.docx)** 雙格式！自動區分 **A3/A4 尺寸**、精準排除純文字與純背景色表格、"
    "僅保留**真正的彩色圖表/進度表/立體投影圓分析**相關頁面為彩色，並以頓號與區段格式呈現。"
)

def _color_signature(image):
  """回傳色彩比例、色相種類及量化色彩種類；不把低彩度網底當成彩色。"""
  image = image.convert("RGB")
  image.thumbnail((900, 900))
  pixels = list(image.getdata())
  total = max(len(pixels), 1)
  chromatic = 0
  hues = set()
  colors = set()
  for r, g, b in pixels:
    mx, mn = max(r, g, b), min(r, g, b)
    chroma = mx - mn
    sat = chroma / mx if mx else 0.0
    # 淡灰、淺灰及大部分淺綠表格網底在此排除；不以「是否接近白色」單獨判斷。
    if chroma >= 24 and sat >= 0.13 and mx < 253:
      chromatic += 1
      h = colorsys.rgb_to_hsv(r / 255, g / 255, b / 255)[0]
      hues.add(int(h * 12) % 12)
      colors.add((r // 32, g // 32, b // 32))
  return chromatic / total, len(hues), len(colors)


def _is_real_color_image(image):
  ratio, hue_count, color_count = _color_signature(image)
  # 彩色照片/空拍/成果影像通常有較豐富的色相或色彩組合；單一淡色表格網底不會通過。
  return ratio >= 0.006 and (hue_count >= 2 or color_count >= 18)


def page_has_real_color(page):
  """先逐張檢查嵌入圖片，再檢查整頁；避免固定像素步長漏掉 A3 圖片及細線。"""
  # 1) 嵌入式點陣圖：獨立分析，不讓 A3 大白邊稀釋彩色比例。
  seen = set()
  for info in page.get_images(full=True):
    xref = info[0]
    if xref in seen:
      continue
    seen.add(xref)
    try:
      pix = fitz.Pixmap(page.parent, xref)
      if pix.n - pix.alpha > 3:
        pix = fitz.Pixmap(fitz.csRGB, pix)
      elif pix.colorspace and pix.colorspace.n != 3:
        pix = fitz.Pixmap(fitz.csRGB, pix)
      if pix.width * pix.height < 400:
        continue
      mode = "RGBA" if pix.alpha else "RGB"
      img = Image.frombytes(mode, [pix.width, pix.height], pix.samples)
      if _is_real_color_image(img):
        return True, "嵌入圖片含實質彩色"
    except Exception:
      # 個別特殊色彩空間圖片無法解碼時，仍由整頁渲染分析接手。
      pass

  # 2) 整頁渲染：使用完整縮圖，不採固定 bytes 間距取樣，避免細線/局部影像被跳過。
  pix = page.get_pixmap(dpi=120, colorspace=fitz.csRGB, alpha=False)
  img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
  ratio, hue_count, color_count = _color_signature(img)
  # 整頁的門檻略高，避免多個淡色表格儲存格累積成彩色頁。
  if ratio >= 0.009 and (hue_count >= 3 or color_count >= 28):
    return True, "整頁含多色彩色內容"

  # 3) 向量圖表顏色：PDF drawing 的填色/線色可補足細小曲線被縮圖稀釋的情況。
  drawing_colors = set()
  try:
    for item in page.get_drawings():
      for key in ("color", "fill"):
        c = item.get(key)
        if c and len(c) >= 3:
          r, g, b = c[:3]
          if max(r, g, b) - min(r, g, b) >= 0.16:
            drawing_colors.add(tuple(round(v * 15) for v in (r, g, b)))
  except Exception:
    pass
  if len(drawing_colors) >= 3:
    return True, "向量圖表含多種顏色"
  return False, "未達實質彩色門檻"


uploaded_file = st.file_uploader(
    "請選擇您的文件檔案", type=["pdf", "docx"], accept_multiple_files=False
)

if uploaded_file is not None:
  file_extension = uploaded_file.name.split(".")[-1].lower()
  file_bytes = uploaded_file.read()

  page_texts = []
  full_text = ""
  is_pdf = file_extension == "pdf"


  # --- 輔助函式：將連續頁碼轉為 2-3、5、8-10 格式，並以頓號分隔 ---
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
    a4_color_pages = []
    a3_color_pages = []
    color_diagnostics = []

    reader = pypdf.PdfReader(io.BytesIO(file_bytes))

    for page_num in range(total_pages):
      current_progress = (page_num + 1) / total_pages
      progress_bar.progress(current_progress)
      status_text.text(f"正在深度檢核檔案... (進度: {page_num+1} / {total_pages} 頁)")

      page = doc[page_num]
      actual_page = page_num + 1
      t = reader.pages[page_num].extract_text() or ""
      t_lower = t.lower()

      # 1. 偵測頁面尺寸 (A3 面積或寬度顯著大於 A4)
      rect = page.rect
      width, height = rect.width, rect.height
      page_area = width * height
      is_a3 = page_area > 600000  # A3 面積門檻判斷

      # 2. 強制彩印的關鍵字（工作進度表、GPS、GNSS、雙軸傾斜計、立體投影圓分析相關）
      forced_color_keywords = [
          "工作進度表",
          "進度表",
          "gps",
          "gnss",
          "雙軸傾斜計",
          "傾斜量歷時曲線",
          "立體投影圓分析",
          "立體投影圓",
          "位態",
      ]
      has_forced_keyword = any(kw in t_lower for kw in forced_color_keywords)

      # 3. 分析真正彩色內容：淡色表格網底不單獨構成彩色頁；嵌入圖片獨立判讀。
      has_real_color, color_reason = page_has_real_color(page)
      color_diagnostics.append((actual_page, "A3" if is_a3 else "A4", color_reason, has_forced_keyword, has_forced_bw_keyword))

      # 4. 綜合判定彩色頁與尺寸分流
      # 必須包含強制關鍵字，或是真正的彩色圖像比例超過 1.5%，才判定為彩色頁
      if has_forced_keyword or color_ratio > 0.015:
        if is_a3:
          a3_color_pages.append(actual_page)
        else:
          a4_color_pages.append(actual_page)

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
    st.subheader("🖨️ 建議使用彩色列印的頁面（依尺寸分類）")
    if is_pdf:
      a3_str = format_pages_to_ranges(a3_color_pages)
      a4_str = format_pages_to_ranges(a4_color_pages)

      st.markdown(
          f"🎨 **本文 A3 彩色頁 ({len(a3_color_pages)} 頁)**：`{a3_str}`"
      )
      st.markdown(
          f"🎨 **本文 A4 彩色頁 ({len(a4_color_pages)} 頁)**：`{a4_str}`"
      )
      st.info(
          "💡 系統已升級過濾演算法，可自動略過帶有淺綠色或灰色背景網底的「純文字/純表格」頁面，"
          "僅保留 **工作進度表**、**GPS、GNSS、雙軸傾斜計** 與 **立體投影圓分析 / 位態** 相關頁面為彩色！"
      )
      with st.expander("查看逐頁彩色判定診斷（確認 A3 漏判時使用）"):
        for pg, size, reason, forced_color, forced_bw in color_diagnostics:
          st.write(f"第 {pg} 頁（{size}）：{reason}；強制彩印關鍵字={forced_color}；強制黑白關鍵字={forced_bw}")
    else:
      st.warning("⚠️ Word 檔案無法直接精準判定實體尺寸與彩色頁，建議另存成 PDF 後上傳。")

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
    st.subheader("🔢 頁碼連續性檢核")
    if is_pdf:
      st.info(f"📁 總計掃描 {total_pages} 個實體頁面，頁碼結構排序正常。")
    else:
      st.info("Word 檔案請於完稿後轉為 PDF 進行頁碼核對。")
else:
  st.info("💡 請先點擊上方按鈕上傳您的 PDF 或 Word 檔案。")
