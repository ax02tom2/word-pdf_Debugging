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

def _color_signature(image, ignore_mask=None):
  """計算非文字區域的實質彩色比例，避免彩色字體/超連結單獨觸發。"""
  image = image.convert("RGB")
  original_size = image.size
  image.thumbnail((900, 900))
  if ignore_mask is not None:
    ignore_mask = ignore_mask.resize(image.size, Image.Resampling.NEAREST)
  pixels = list(image.getdata())
  ignored = list(ignore_mask.getdata()) if ignore_mask is not None else None
  total = max(len(pixels) - (sum(1 for v in ignored if v) if ignored else 0), 1)
  chromatic = 0
  hues = set()
  colors = set()
  for i, (r, g, b) in enumerate(pixels):
    if ignored is not None and ignored[i]:
      continue
    mx, mn = max(r, g, b), min(r, g, b)
    chroma = mx - mn
    sat = chroma / mx if mx else 0.0
    # 灰色、低彩度網底及近白色抗鋸齒像素不列入彩色。
    if chroma >= 30 and sat >= 0.18 and mx < 253:
      chromatic += 1
      h = colorsys.rgb_to_hsv(r / 255, g / 255, b / 255)[0]
      hues.add(int(h * 12) % 12)
      colors.add((r // 32, g // 32, b // 32))
  return chromatic / total, len(hues), len(colors)


def _is_real_color_image(image, strict=False, color_context=False):
  ratio, hue_count, color_count = _color_signature(image)
  if strict:
    min_ratio = 0.010 if color_context else 0.015
    min_colors = 22 if color_context else 30
    # 指定黑白頁僅讓色彩豐富、較像實質底圖/照片的影像取得彩色優先；
    # 雨量曲線的藍紅雙色，以及風險矩陣的少量純色區塊不會通過。
    return ratio >= min_ratio and hue_count >= 3 and color_count >= min_colors
  return ratio >= 0.006 and (hue_count >= 2 or color_count >= 18)


def _text_ignore_mask(page, pix):
  """建立 PDF 文字區域遮罩；文字顏色不單獨作為整頁彩色證據。"""
  scale_x = pix.width / max(page.rect.width, 1)
  scale_y = pix.height / max(page.rect.height, 1)
  mask = Image.new("L", (pix.width, pix.height), 0)
  from PIL import ImageDraw
  draw = ImageDraw.Draw(mask)
  try:
    data = page.get_text("dict")
    for block in data.get("blocks", []):
      for line in block.get("lines", []):
        for span in line.get("spans", []):
          x0, y0, x1, y1 = span.get("bbox", (0, 0, 0, 0))
          # 僅遮文字字框，保留相鄰的圖表線條及底圖。
          pad = 1.2
          box = (
              int((x0 - page.rect.x0 - pad) * scale_x),
              int((y0 - page.rect.y0 - pad) * scale_y),
              int((x1 - page.rect.x0 + pad) * scale_x),
              int((y1 - page.rect.y0 + pad) * scale_y),
          )
          draw.rectangle(box, fill=255)
  except Exception:
    # 無法取得文字框時不套用遮罩，避免影響原有圖片辨識。
    return None
  return mask


def _has_substantial_colored_vectors(page):
  """向量彩色判斷須具備足夠圖形規模，排除零星彩色符號/裝飾線。"""
  colors = set()
  colored_count = 0
  covered_area = 0.0
  line_extent = 0.0
  page_area = max(page.rect.width * page.rect.height, 1.0)
  page_diagonal = max((page.rect.width ** 2 + page.rect.height ** 2) ** 0.5, 1.0)
  try:
    for item in page.get_drawings():
      color_found = False
      for key in ("color", "fill"):
        c = item.get(key)
        if c and len(c) >= 3 and max(c[:3]) - min(c[:3]) >= 0.16:
          colors.add(tuple(round(v * 15) for v in c[:3]))
          color_found = True
      if not color_found:
        continue
      rect = item.get("rect")
      if rect:
        w, h = max(rect.width, 0), max(rect.height, 0)
        covered_area += w * h
        # 細長線段以長度累積，避免面積接近零而漏判彩色線圖。
        line_extent += max(w, h)
      colored_count += 1
  except Exception:
    return False
  return (
      len(colors) >= 3
      and colored_count >= 4
      and (covered_area / page_area >= 0.0025 or line_extent / page_diagonal >= 0.12)
  )


def _has_rich_colored_vectors(page):
  """指定黑白頁中的彩色優先例外：需是多色且具一定規模的向量底圖。"""
  colors = set()
  colored_count = 0
  covered_area = 0.0
  line_extent = 0.0
  page_area = max(page.rect.width * page.rect.height, 1.0)
  page_diagonal = max((page.rect.width ** 2 + page.rect.height ** 2) ** 0.5, 1.0)
  try:
    for item in page.get_drawings():
      found = False
      for key in ("color", "fill"):
        c = item.get(key)
        if c and len(c) >= 3 and max(c[:3]) - min(c[:3]) >= 0.16:
          colors.add(tuple(round(v * 15) for v in c[:3]))
          found = True
      if found:
        rect = item.get("rect")
        if rect:
          covered_area += max(rect.width, 0) * max(rect.height, 0)
          line_extent += max(rect.width, rect.height)
        colored_count += 1
  except Exception:
    return False
  return (
      len(colors) >= 8
      and colored_count >= 12
      and (covered_area / page_area >= 0.008 or line_extent / page_diagonal >= 0.20)
  )


def page_has_real_color(page, bw_priority=False, color_context=False):
  """辨識實質彩色內容；指定黑白圖表採較嚴格的彩色底圖門檻。"""
  seen = set()
  page_area = max(page.rect.width * page.rect.height, 1.0)
  for info in page.get_images(full=True):
    xref = info[0]
    if xref in seen:
      continue
    seen.add(xref)
    try:
      placements = page.get_image_rects(xref)
      largest_placement = max(
          (r.width * r.height / page_area for r in placements), default=0
      )
      if largest_placement < 0.004:
        continue
      pix = fitz.Pixmap(page.parent, xref)
      if pix.n - pix.alpha > 3 or (pix.colorspace and pix.colorspace.n != 3):
        pix = fitz.Pixmap(fitz.csRGB, pix)
      if pix.width * pix.height < 400:
        continue
      mode = "RGBA" if pix.alpha else "RGB"
      img = Image.frombytes(mode, [pix.width, pix.height], pix.samples)
      if _is_real_color_image(img, strict=bw_priority, color_context=color_context):
        reason = "指定黑白頁中仍偵測到色彩豐富的實質底圖/影像" if bw_priority else "嵌入圖片含實質彩色且版面佔比足夠"
        return True, reason
    except Exception:
      pass

  # 指定黑白頁（雨量圖、風險矩陣）不以圖表線條/區塊顏色觸發彩印。
  # 但同頁若有色彩豐富的嵌入式平面圖/空拍影像，上方影像檢查仍可保留彩色。
  if bw_priority:
    if _has_rich_colored_vectors(page):
      return True, "指定黑白頁中另有多色且具規模的向量底圖，彩色優先"
    return False, "命中指定黑白圖表規則；未發現色彩豐富的獨立底圖/影像"

  pix = page.get_pixmap(dpi=120, colorspace=fitz.csRGB, alpha=False)
  img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
  text_mask = _text_ignore_mask(page, pix)
  ratio, hue_count, color_count = _color_signature(img, text_mask)
  if ratio >= 0.009 and (hue_count >= 3 or color_count >= 28):
    return True, "非文字區域含實質多色內容"

  if _has_substantial_colored_vectors(page):
    return True, "向量彩色圖形達實質面積/線段門檻"
  return False, "未偵測到實質彩色圖片或圖形（文字與零星裝飾已排除）"


DEFAULT_BW_KEYWORDS = [
    "雨量", "降雨", "雨量歷線", "降雨歷線", "累積雨量", "時雨量",
    "日雨量", "雨量站", "雨量圖", "風險等級", "風險評估",
    "分數評估", "評估矩陣", "風險矩陣", "風險等級評估矩陣",
]
DEFAULT_COLOR_CONTEXT_KEYWORDS = [
    "平面圖", "空拍圖", "正射影像", "航照圖", "彩色影像",
    "地形影像", "崩塌範圍圖",
]

def _parse_keyword_lines(value):
  return [line.strip().lower() for line in value.replace("，", "\n").replace(",", "\n").splitlines() if line.strip()]

with st.expander("⚙️ 彩色／黑白判定規則（可自行調整）", expanded=False):
  st.caption("黑白關鍵字只會限制雨量圖、風險矩陣等指定圖表；同頁若另有色彩豐富的平面圖／空拍圖，仍保留彩色優先。關鍵字不會單獨決定頁面顏色。")
  bw_keyword_text = st.text_area(
      "指定黑白內容關鍵字（每行一項）",
      value="\n".join(DEFAULT_BW_KEYWORDS), height=150, key="bw_keyword_rules"
  )
  color_context_text = st.text_area(
      "彩色成果內容提示關鍵字（每行一項）",
      value="\n".join(DEFAULT_COLOR_CONTEXT_KEYWORDS), height=100, key="color_context_rules"
  )
  st.caption("設定只影響本次分析；若要永久保存，可將常用詞保留在預設清單中。")

BW_KEYWORDS = _parse_keyword_lines(bw_keyword_text)
COLOR_CONTEXT_KEYWORDS = _parse_keyword_lines(color_context_text)

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

      # 2. 強制彩印只保留明確指向彩色圖表/成果的詞組。
      # GPS、GNSS、雙軸傾斜計等設備或一般文字本身不代表需要彩印。
      forced_color_keywords = [
          "工作進度表",
          "進度表",
          "傾斜量歷時曲線",
          "立體投影圓分析",
          "立體投影圓",
          "位態分析",
      ]
      has_forced_keyword = any(kw in t_lower for kw in forced_color_keywords)

      # 指定黑白關鍵字涵蓋雨量圖與風險評估矩陣；彩色成果提示只作輔助辨識。
      has_forced_bw_keyword = any(kw in t_lower for kw in BW_KEYWORDS)
      has_color_context_keyword = any(kw in t_lower for kw in COLOR_CONTEXT_KEYWORDS)

      # 3. 一般頁維持原彩色辨識；指定黑白頁採嚴格影像豐富度門檻，
      #    避免藍紅雨量曲線或四色風險矩陣把整頁誤判為彩色。
      has_real_color, color_reason = page_has_real_color(
          page, bw_priority=has_forced_bw_keyword,
          color_context=has_color_context_keyword
      )
      color_diagnostics.append((actual_page, "A3" if is_a3 else "A4", has_real_color, color_reason, has_forced_keyword, has_forced_bw_keyword, has_color_context_keyword))

      # 4. 優先順序：實質彩色內容 > 黑白限制。只有未偵測到實質彩色時，雨量頁才維持黑白。
      if has_real_color:
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
          "💡 系統會略過灰色及低彩度淺綠色表格網底，並排除文字顏色與零星向量裝飾。"
          "雨量圖及風險評估矩陣列為指定黑白內容；藍紅曲線或少量色塊不會單獨觸發彩印。"
          "同頁若另有色彩豐富的平面圖、空拍圖或正射影像，仍保留彩色優先。"
      )
      with st.expander("查看逐頁彩色判定診斷（確認誤判/漏判時使用）"):
        for pg, size, has_color, reason, forced_color, forced_bw, color_context in color_diagnostics:
          decision = "彩色優先" if has_color else ("指定黑白（未發現獨立彩色底圖）" if forced_bw else "黑白")
          st.write(f"第 {pg} 頁（{size}）：{decision}；{reason}；彩色內容提示={color_context}；指定黑白關鍵字={forced_bw}")
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
