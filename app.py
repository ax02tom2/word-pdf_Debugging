import io
import re
import unicodedata
import colorsys
from collections import Counter
from PIL import Image, ImageChops
import fitz  # PyMuPDF
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
  """快速計算非文字區域的色彩特徵；先縮圖，再利用 Pillow 直方圖統計。"""
  image = image.convert("RGB")
  image.thumbnail((500, 500))
  width, height = image.size
  total_pixels = max(width * height, 1)

  # 以 HSV 飽和度／亮度遮罩找彩色像素，避免逐一對數十萬像素呼叫 colorsys。
  h, sat, val = image.convert("HSV").split()
  chroma_mask = sat.point(lambda p: 255 if p >= 46 else 0)
  visible_mask = val.point(lambda p: 255 if p < 253 else 0)
  min_chroma_mask = ImageChops.multiply(sat, val).point(lambda p: 255 if p >= 30 else 0)
  chroma_mask = ImageChops.multiply(chroma_mask, visible_mask)
  chroma_mask = ImageChops.multiply(chroma_mask, min_chroma_mask)

  ignored_pixels = 0
  keep_mask = None
  if ignore_mask is not None:
    ignore_mask = ignore_mask.convert("L").resize(image.size, Image.Resampling.NEAREST)
    ignored_pixels = sum(ignore_mask.histogram()[128:])
    keep_mask = ignore_mask.point(lambda p: 0 if p >= 128 else 255)
    chroma_mask = ImageChops.multiply(chroma_mask, keep_mask)

  hue_hist = h.histogram(mask=chroma_mask)
  chromatic = sum(hue_hist)
  eligible = max(total_pixels - ignored_pixels, 1)

  # 將 256 個 Hue 細格合併為 12 個色相區，排除雨量圖常見的藍色＋紅色雙系列。
  hue_bins = [0] * 12
  for i, count in enumerate(hue_hist):
    if count:
      hue_bins[min(i * 12 // 256, 11)] += count
  hue_count = sum(1 for count in hue_bins if count >= max(2, chromatic * 0.002))
  dominant_hue_share = (max(hue_bins) / chromatic) if chromatic else 1.0

  # 量化後只檢查最多 256 個調色盤顏色，不再逐像素建立 Python set。
  quantized = image.quantize(colors=256, method=Image.Quantize.FASTOCTREE)
  palette = quantized.getpalette() or []
  palette_hist = quantized.histogram(mask=keep_mask) if keep_mask is not None else quantized.histogram()
  color_count = 0
  for idx, count in enumerate(palette_hist[:256]):
    if count <= 0:
      continue
    base = idx * 3
    if base + 2 >= len(palette):
      continue
    r, g, b = palette[base:base + 3]
    mx, mn = max(r, g, b), min(r, g, b)
    chroma = mx - mn
    if mx and chroma >= 30 and chroma / mx >= 0.18 and mx < 253:
      color_count += 1

  return chromatic / eligible, hue_count, color_count, dominant_hue_share


def _is_real_color_image(image, strict=False):
  ratio, hue_count, color_count, dominant_hue_share = _color_signature(image)
  if strict:
    # 嚴格模式排除雨量圖常見的藍色柱線＋紅色標記；需要多個實質色相共同構成內容。
    return (
        ratio >= 0.018
        and hue_count >= 4
        and color_count >= 24
        and dominant_hue_share <= 0.86
    )
  return ratio >= 0.006 and (hue_count >= 2 or color_count >= 18)

def _text_ignore_mask(page, pix, text_dict=None):
  """建立 PDF 文字區域遮罩；文字顏色不單獨作為整頁彩色證據。"""
  scale_x = pix.width / max(page.rect.width, 1)
  scale_y = pix.height / max(page.rect.height, 1)
  mask = Image.new("L", (pix.width, pix.height), 0)
  from PIL import ImageDraw
  draw = ImageDraw.Draw(mask)
  try:
    data = text_dict if text_dict is not None else page.get_text("dict")
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


def _vector_color_profile(page):
  """一次掃描 PDF 向量圖形，回傳顏色、色相分布及圖形規模。"""
  colors = set()
  hue_counts = Counter()
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
          rgb = tuple(round(v * 255) for v in c[:3])
          colors.add(tuple(v // 24 for v in rgb))
          hue = colorsys.rgb_to_hsv(*(v / 255 for v in rgb))[0]
          hue_counts[int(hue * 12) % 12] += 1
          found = True
      if not found:
        continue
      rect = item.get("rect")
      if rect:
        w, h = max(rect.width, 0), max(rect.height, 0)
        covered_area += w * h
        line_extent += max(w, h)
      colored_count += 1
  except Exception:
    return None
  return {
      "colors": colors,
      "hue_count": sum(1 for n in hue_counts.values() if n >= 2),
      "colored_count": colored_count,
      "area_ratio": covered_area / page_area,
      "line_ratio": line_extent / page_diagonal,
  }


def _has_substantial_colored_vectors(page, profile=None):
  """一般頁面的向量彩色判斷。"""
  profile = profile if profile is not None else _vector_color_profile(page)
  if not profile:
    return False
  return (
      len(profile["colors"]) >= 3
      and profile["hue_count"] >= 3
      and profile["colored_count"] >= 4
      and (profile["area_ratio"] >= 0.0025 or profile["line_ratio"] >= 0.12)
  )


def _has_rich_colored_vectors(page, profile=None):
  """雨量等指定黑白頁的彩色例外：排除藍紅雙色曲線及零星線條。"""
  profile = profile if profile is not None else _vector_color_profile(page)
  if not profile:
    return False
  return (
      len(profile["colors"]) >= 4
      and profile["hue_count"] >= 4
      and profile["colored_count"] >= 12
      and (profile["area_ratio"] >= 0.008 or profile["line_ratio"] >= 0.20)
  )

def page_has_real_color(page, bw_priority=False, image_color_cache=None):
  """辨識實質彩色內容；指定黑白頁只接受多色且具規模的獨立內容。"""
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
      min_placement = 0.015 if bw_priority else 0.004
      if largest_placement < min_placement:
        continue
      cache_key = (xref, bool(bw_priority))
      cached_color = image_color_cache.get(cache_key) if image_color_cache is not None else None
      if cached_color is None:
        pix = fitz.Pixmap(page.parent, xref)
        if pix.n - pix.alpha > 3 or (pix.colorspace and pix.colorspace.n != 3):
          pix = fitz.Pixmap(fitz.csRGB, pix)
        if pix.width * pix.height < 400:
          continue
        mode = "RGBA" if pix.alpha else "RGB"
        img = Image.frombytes(mode, [pix.width, pix.height], pix.samples)
        cached_color = _is_real_color_image(img, strict=bw_priority)
        if image_color_cache is not None:
          image_color_cache[cache_key] = cached_color
      if cached_color:
        reason = "指定黑白頁中另有多色且具規模的獨立影像" if bw_priority else "嵌入圖片含實質彩色且版面佔比足夠"
        return True, reason
    except Exception:
      pass

  # 向量內容只掃描一次；雨量圖常見的藍色柱狀＋紅色標記不會通過嚴格色相數。
  vector_profile = _vector_color_profile(page)
  if bw_priority:
    if _has_rich_colored_vectors(page, vector_profile):
      return True, "指定黑白頁中另有多色且具規模的向量成果圖，彩色優先"
    return False, "命中指定黑白圖表規則；未發現其他多色且具規模的彩色成果"

  if _has_substantial_colored_vectors(page, vector_profile):
    return True, "向量彩色圖形達實質面積／線段門檻"

  # 只有前述快速檢查無法判斷時才渲染頁面；由 120 dpi 降至 90 dpi，減少 A3 頁運算量。
  pix = page.get_pixmap(dpi=90, colorspace=fitz.csRGB, alpha=False)
  img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
  text_dict = page.get_text("dict")
  text_mask = _text_ignore_mask(page, pix, text_dict=text_dict)
  ratio, hue_count, color_count, _ = _color_signature(img, text_mask)
  if ratio >= 0.009 and (hue_count >= 3 or color_count >= 28):
    return True, "非文字區域含實質多色內容"

  return False, "未偵測到實質彩色圖片或圖形（文字與零星裝飾已排除）"


# 雨量規則採「主要圖表為雨量」判定；避免單純提到雨量就把整頁成果圖改成黑白。
DEFAULT_BW_KEYWORDS = [
    "雨量圖", "雨量歷線", "降雨歷線", "降雨量歷線", "雨量柱狀圖",
    "雨量分布圖", "月雨量圖", "日雨量圖", "時雨量圖", "累積雨量圖",
    "雨量統計圖", "歷年雨量圖", "歷年月份雨量圖", "雨量站歷年",
]
DEFAULT_RISK_BW_KEYWORDS = [
    "風險等級評估矩陣", "風險評估矩陣", "分數評估矩陣", "風險矩陣", "評估矩陣",
]
DEFAULT_FORCE_COLOR_KEYWORDS = [
    "工作進度表", "進度表", "傾斜量歷時曲線", "立體投影圓分析", "立體投影圓", "位態分析",
]
DEFAULT_COLOR_CONTEXT_KEYWORDS = [
    "平面圖", "空拍圖", "正射影像", "航照圖", "彩色影像", "地形影像",
    "崩塌範圍圖", "影響範圍圖", "成果圖", "成果分析圖", "分析成果圖",
    "分析結果圖", "結果圖", "監測成果圖", "監測結果圖", "成果分析",
    "變形分析圖", "位移向量圖", "地表變位圖", "流動深度圖", "流速圖",
    "堆積厚度圖",
]


def _parse_keyword_lines(value):
  return [unicodedata.normalize("NFKC", line.strip()).lower()
          for line in value.replace("，", "\n").replace(",", "\n").splitlines() if line.strip()]

with st.expander("⚙️ 彩色／黑白判定規則（可自行調整）", expanded=False):
  st.caption("強制彩色關鍵字優先級最高；雨量只有在頁面主要內容為雨量圖時才套用黑白。若同頁另有成果圖、平面圖或空拍圖，則彩色優先。")
  force_color_keyword_text = st.text_area(
      "強制彩色頁面關鍵字（每行一項）",
      value="\n".join(DEFAULT_FORCE_COLOR_KEYWORDS), height=110, key="force_color_keyword_rules"
  )
  bw_keyword_text = st.text_area(
      "雨量等指定黑白圖表關鍵字（每行一項）",
      value="\n".join(DEFAULT_BW_KEYWORDS), height=130, key="bw_keyword_rules"
  )
  risk_bw_keyword_text = st.text_area(
      "指定黑白風險矩陣關鍵字（每行一項）",
      value="\n".join(DEFAULT_RISK_BW_KEYWORDS), height=90, key="risk_bw_keyword_rules"
  )
  color_context_text = st.text_area(
      "其他彩色成果提示關鍵字（與雨量同頁時可優先彩色）",
      value="\n".join(DEFAULT_COLOR_CONTEXT_KEYWORDS), height=130, key="color_context_rules"
  )
  st.caption("設定只影響本次分析；可依報告圖表名稱自行增減。雨量與風險矩陣是不同規則：雨量允許同頁其他成果內容優先彩色，風險矩陣則預設黑白。")

FORCE_COLOR_KEYWORDS = _parse_keyword_lines(force_color_keyword_text)
BW_KEYWORDS = _parse_keyword_lines(bw_keyword_text)
RISK_BW_KEYWORDS = _parse_keyword_lines(risk_bw_keyword_text)
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
    image_color_cache = {}  # 同一 PDF 重複使用的嵌入圖片只分析一次。
    progress_step = max(1, total_pages // 80)
    status_step = max(1, total_pages // 30)

    for page_num in range(total_pages):
      current_progress = (page_num + 1) / total_pages
      if page_num == 0 or (page_num + 1) % progress_step == 0 or page_num + 1 == total_pages:
        progress_bar.progress(current_progress)
      if page_num == 0 or (page_num + 1) % status_step == 0 or page_num + 1 == total_pages:
        status_text.text(f"正在深度檢核檔案... (進度: {page_num+1} / {total_pages} 頁)")

      page = doc[page_num]
      actual_page = page_num + 1
      t = page.get_text("text") or ""
      # PDF 字型常將「量、歷、年、進、度」等中文字抽取為相容字形（如 量、歷、年、進度）；
      # 先做 NFKC 正規化，避免雨量黑白及工作進度彩色關鍵字漏判。
      t_lower = unicodedata.normalize("NFKC", t).lower()
      t_rule = re.sub(r"\s+", "", t_lower)

      # 1. 偵測頁面尺寸 (A3 面積或寬度顯著大於 A4)
      rect = page.rect
      width, height = rect.width, rect.height
      page_area = width * height
      is_a3 = page_area > 600000  # A3 面積門檻判斷

      # 2. 分別辨識強制彩色、雨量黑白、風險矩陣黑白及其他成果脈絡。
      force_color_hits = [kw for kw in FORCE_COLOR_KEYWORDS if kw in t_rule]
      # 「雨量／降雨」只代表頁面含有雨量內容；搭配成果圖脈絡時不能套用整頁黑白。
      has_rain_mention = any(kw in t_rule for kw in ("雨量", "降雨", "降雨量"))
      rain_rule_text = re.sub(r"(?:無|沒有|未含|未繪製|未提供)雨量圖(?:表)?", "", t_rule)
      rain_bw_hits = [kw for kw in BW_KEYWORDS if kw in rain_rule_text]
      risk_bw_hits = [kw for kw in RISK_BW_KEYWORDS if kw in t_rule]
      color_context_hits = [kw for kw in COLOR_CONTEXT_KEYWORDS if kw in t_rule]
      has_forced_keyword = bool(force_color_hits)
      has_rain_bw_keyword = bool(rain_bw_hits)
      has_risk_bw_keyword = bool(risk_bw_hits)
      has_color_context_keyword = bool(color_context_hits)
      has_forced_bw_keyword = has_rain_bw_keyword or has_risk_bw_keyword

      # 3. 優先順序：明確強制彩色 > 同頁其他成果脈絡 > 指定黑白圖表 > 一般色彩辨識。
      if has_forced_keyword:
        has_real_color, color_reason = True, "命中強制彩色關鍵字，彩色優先"
      elif has_rain_mention and has_color_context_keyword:
        has_real_color, color_reason = True, "雨量內容與其他成果內容同頁，依規則彩色優先"
      elif has_risk_bw_keyword and not has_color_context_keyword:
        has_real_color, color_reason = False, "命中風險等級／分數評估矩陣黑白規則"
      elif has_forced_bw_keyword:
        has_real_color, color_reason = page_has_real_color(
            page, bw_priority=True, image_color_cache=image_color_cache
        )
      else:
        has_real_color, color_reason = page_has_real_color(
            page, image_color_cache=image_color_cache
        )

      color_diagnostics.append((
          actual_page, "A3" if is_a3 else "A4", has_real_color, color_reason,
          force_color_hits, rain_bw_hits, risk_bw_hits, color_context_hits
      ))

      if has_real_color:
        if is_a3:
          a3_color_pages.append(actual_page)
        else:
          a4_color_pages.append(actual_page)

      page_texts.append((actual_page, t))

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
          "單獨雨量圖與風險評估矩陣依規則列為黑白；雨量若與工作進度、成果圖、平面圖或空拍圖同頁，則彩色優先。"
      )
      with st.expander("查看逐頁彩色判定診斷（確認誤判/漏判時使用）"):
        for pg, size, has_color, reason, force_hits, rain_hits, risk_hits, context_hits in color_diagnostics:
          decision = "彩色優先" if has_color else ("指定黑白" if (rain_hits or risk_hits) else "黑白")
          st.write(
              f"第 {pg} 頁（{size}）：{decision}；{reason}；"
              f"強制彩色={', '.join(force_hits) or '無'}；"
              f"雨量黑白={', '.join(rain_hits) or '無'}；"
              f"風險矩陣黑白={', '.join(risk_hits) or '無'}；"
              f"彩色成果提示={', '.join(context_hits) or '無'}"
          )
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
