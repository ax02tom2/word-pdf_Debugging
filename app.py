import io
import json
from pathlib import Path
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


def _color_rule_ok(sig, level="normal"):
  """依判定等級檢查色彩特徵。
  strict ：原本的雨量藍紅雙色排除門檻（過嚴，會漏掉色調單純的 A3 空拍／地形圖，已不用於一般頁面）。
  normal ：一般頁面；色彩占比夠大且有足夠色階，或多色相即成立。
  lenient：已有關鍵字佐證（GPS／GNSS 等）時，只需少量實質彩色。
  雨量與風險矩陣改由關鍵字規則處理，不再靠提高像素門檻來硬排除。
  """
  ratio, hue_count, color_count, dominant_hue_share = sig
  if level == "strict":
    return ratio >= 0.018 and hue_count >= 4 and color_count >= 24 and dominant_hue_share <= 0.86
  if level == "rich":
    # 黑白標題頁的例外：必須是色相與色階都很豐富的照片／成果圖，雨量藍紅線與紅黃綠矩陣色塊達不到。
    return ratio >= 0.02 and hue_count >= 4 and color_count >= 80
  if level == "lenient":
    return ratio >= 0.006 and (hue_count >= 2 or color_count >= 18)
  big = ratio >= 0.015 and color_count >= 10 and (hue_count >= 3 or ratio >= 0.05)
  # 面積占比不大，但色相與色階都很豐富（例如 A3 上的小幅地圖／成果圖）也算彩色。
  rich = ratio >= 0.004 and hue_count >= 3 and color_count >= 40
  return big or rich


def _is_real_color_image(image, level="normal"):
  return _color_rule_ok(_color_signature(image), level)

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
          pad = 2.5
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

def page_has_real_color(page, image_color_cache=None, level="normal"):
  """以實際圖片、向量、整頁渲染三層判定實質彩色；回傳 (是否彩色, 原因, 色彩特徵)。"""
  seen = set()
  page_area = max(page.rect.width * page.rect.height, 1.0)
  image_errors = 0
  best_sig = None
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
      sig = image_color_cache.get(xref) if image_color_cache is not None else None
      if sig is None:
        pix = fitz.Pixmap(page.parent, xref)
        if pix.n - pix.alpha > 3 or (pix.colorspace and pix.colorspace.n != 3):
          pix = fitz.Pixmap(fitz.csRGB, pix)
        if pix.width * pix.height < 400:
          continue
        mode = "RGBA" if pix.alpha else "RGB"
        img = Image.frombytes(mode, [pix.width, pix.height], pix.samples)
        if mode == "RGBA":
          img = img.convert("RGB")
        sig = _color_signature(img)
        if image_color_cache is not None:
          image_color_cache[xref] = sig
      best_sig = sig
      if _color_rule_ok(sig, level):
        return True, "嵌入圖片含實質彩色內容", sig
    except Exception:
      # 圖片解碼失敗時不放棄，改由後面的向量／整頁渲染判定（原本是直接略過，會造成 A3 圖漏判）。
      image_errors += 1

  vector_profile = _vector_color_profile(page)
  if _has_substantial_colored_vectors(page, vector_profile):
    return True, "向量彩色圖形達實質面積／線段門檻", best_sig

  # 整頁渲染備援（含無法單獨解碼的圖片）；A3 頁面也使用同一組門檻。
  pix = page.get_pixmap(dpi=90, colorspace=fitz.csRGB, alpha=False)
  img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
  text_mask = _text_ignore_mask(page, pix, text_dict=page.get_text("dict"))
  sig = _color_signature(img, text_mask)
  if _color_rule_ok(sig, level):
    return True, "整頁渲染：非文字區域含實質彩色內容", sig

  note = f"（{image_errors} 張圖片解碼失敗，已改用整頁渲染）" if image_errors else ""
  return False, "未偵測到實質彩色圖片或圖形" + note, sig


# ── 判定規則（只有兩份關鍵字清單，且只看「圖表標題／短行」，內文敘述提到不算）──
DEFAULT_COLOR_KEYWORDS = [
    "工作進度表", "進度表", "傾斜量歷時曲線", "立體投影圓分析", "立體投影圓", "位態分析",
    "gps", "gnss", "雙軸傾斜計", "變位歷時曲線", "變位歷時", "地質剖面圖",
]
DEFAULT_BW_KEYWORDS = [
    "雨量圖", "雨量歷線", "降雨歷線", "降雨量歷線", "雨量柱狀圖",
    "雨量分布圖", "月雨量圖", "日雨量圖", "時雨量圖", "累積雨量圖",
    "雨量統計圖", "歷年雨量圖", "歷年月份雨量圖", "雨量站歷年",
    "風險等級評估矩陣", "風險評估矩陣", "分數評估矩陣", "風險矩陣", "評估矩陣",
]


def _looks_like_toc(text):
  """目錄頁（工作進度表 ........ 12）不套用標題規則。"""
  leaders = re.findall(r"[.．…·‧⋯]{3,}\s*\d+\s*$", text, flags=re.M)
  return len(leaders) >= 3 or any(k in text[:200] for k in ("目錄", "目次"))


def _title_hits(raw_text, keywords, caption_only=False):
  """只計入出現在短行（<=40 字）的關鍵字；caption_only 時須為「圖／表 x」開頭的標題行。"""
  hits = set()
  for line in raw_text.splitlines():
    n = re.sub(r"\s+", "", unicodedata.normalize("NFKC", line).lower())
    if not 0 < len(n) <= 60:
      continue
    if caption_only and not re.match(r"^(圖|表|figure|table)", n):
      continue
    if not caption_only and len(n) > 40:
      continue
    hits.update(kw for kw in keywords if kw in n)
  return sorted(hits)


def decide_page_color(page, t_raw, image_cache):
  """回傳 (是否彩色, 規則, 原因, 彩色關鍵字, 黑白關鍵字)。

  判定順序（由上而下，命中即停）：
    1. 圖表標題含「彩色關鍵字」→ 彩色
    2. 短行含「彩色關鍵字」且頁面有少量實質彩色 → 彩色
    3. 圖表標題含「黑白關鍵字」→ 黑白（同頁有豐富彩色圖片則例外）
    4. 其餘 → 依圖片／向量／整頁渲染的實際色彩判定
  目錄頁略過 1~3。
  """
  toc = _looks_like_toc(t_raw)
  c_cap = [] if toc else _title_hits(t_raw, COLOR_KEYWORDS, caption_only=True)
  c_short = [] if toc else _title_hits(t_raw, COLOR_KEYWORDS)
  # 黑白關鍵字只認「圖／表 x」標題行；圖內座標軸文字（如「日雨量(mm)」）不算。
  bw_short = [] if toc else _title_hits(t_raw, BW_KEYWORDS, caption_only=True)

  if c_cap:
    return True, "①彩色標題", "圖表標題命中彩色關鍵字", c_short, bw_short
  if c_short:
    ok, why, sig = page_has_real_color(page, image_cache, level="lenient")
    if ok:
      return True, "②標題＋實際彩色", why, c_short, bw_short
  if bw_short:
    ok, why, sig = page_has_real_color(page, image_cache, level="rich")
    if ok:
      return True, "③黑白標題但有豐富彩色", "同頁含多色相彩色圖片／圖形，不套用黑白：" + why, c_short, bw_short
    return False, "③黑白標題", "圖表標題命中黑白關鍵字", c_short, bw_short

  has_color, reason, sig = page_has_real_color(page, image_cache, level="normal")
  if sig:
    reason += f"［占比{sig[0]:.3f}／色相{sig[1]}／色階{sig[2]}］"
  return has_color, "④實際色彩", reason, c_short, bw_short


def _parse_keyword_lines(value):
  return [unicodedata.normalize("NFKC", line.strip()).lower()
          for line in value.replace("，", "\n").replace(",", "\n").splitlines() if line.strip()]

RULES_PATH = Path(__file__).with_name("keyword_rules.json")


def _load_saved_rules():
  try:
    data = json.loads(RULES_PATH.read_text(encoding="utf-8"))
    return data["color"], data["bw"]
  except Exception:
    return DEFAULT_COLOR_KEYWORDS, DEFAULT_BW_KEYWORDS


def _reset_rules():
  st.session_state["color_keyword_rules"] = "\n".join(DEFAULT_COLOR_KEYWORDS)
  st.session_state["bw_keyword_rules"] = "\n".join(DEFAULT_BW_KEYWORDS)


def _save_rules():
  data = {
      "color": [x for x in st.session_state["color_keyword_rules"].splitlines() if x.strip()],
      "bw": [x for x in st.session_state["bw_keyword_rules"].splitlines() if x.strip()],
  }
  try:
    RULES_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    st.session_state["rules_msg"] = "✅ 已儲存到伺服器（重開瀏覽器仍在；若主機重啟或重新部署可能消失，建議另外下載備份）。"
  except Exception as exc:
    st.session_state["rules_msg"] = f"⚠️ 儲存失敗：{exc}，請改用下載備份。"


if "color_keyword_rules" not in st.session_state:
  _c, _b = _load_saved_rules()
  st.session_state["color_keyword_rules"] = "\n".join(_c)
  st.session_state["bw_keyword_rules"] = "\n".join(_b)

with st.expander("⚙️ 彩色／黑白判定規則（可自行調整）", expanded=False):
  st.markdown(
      "**判定順序**：①圖表標題含彩色關鍵字 → 彩色　②短行含彩色關鍵字且頁面確有彩色 → 彩色　"
      "③圖表標題（圖／表 x…）含黑白關鍵字 → 黑白（同頁若有豐富多色圖片則仍判彩色）　"
      "④其餘依實際圖片／向量色彩判定。  \n"
      "彩色關鍵字看圖表標題或短行；黑白關鍵字只認「圖／表」開頭的標題行；目錄頁不套用。"
  )
  # 匯入備份：必須在文字框建立之前處理
  imported = st.file_uploader("匯入先前備份的規則（.json）", type=["json"], key="rules_import")
  if imported is not None and st.session_state.get("rules_imported_id") != (imported.name, imported.size):
    try:
      data = json.loads(imported.getvalue().decode("utf-8"))
      st.session_state["color_keyword_rules"] = "\n".join(data["color"])
      st.session_state["bw_keyword_rules"] = "\n".join(data["bw"])
      st.session_state["rules_imported_id"] = (imported.name, imported.size)
      st.session_state["rules_msg"] = "✅ 已匯入規則。"
    except Exception:
      st.session_state["rules_imported_id"] = (imported.name, imported.size)
      st.session_state["rules_msg"] = "⚠️ 匯入失敗：檔案格式不正確。"

  col_a, col_b = st.columns(2)
  with col_a:
    color_keyword_text = st.text_area(
        "🎨 彩色關鍵字（每行一項）", height=260, key="color_keyword_rules"
    )
  with col_b:
    bw_keyword_text = st.text_area(
        "⬛ 黑白關鍵字：雨量、風險矩陣等（每行一項）", height=260, key="bw_keyword_rules"
    )
  b1, b2, b3 = st.columns(3)
  b1.button("💾 儲存規則", on_click=_save_rules, use_container_width=True)
  b2.button("↩️ 還原預設", on_click=_reset_rules, use_container_width=True)
  b3.download_button(
      "⬇️ 下載備份 (.json)",
      data=json.dumps(
          {"color": [x for x in color_keyword_text.splitlines() if x.strip()],
           "bw": [x for x in bw_keyword_text.splitlines() if x.strip()]},
          ensure_ascii=False, indent=2),
      file_name="keyword_rules.json", mime="application/json", use_container_width=True,
  )
  if st.session_state.get("rules_msg"):
    st.caption(st.session_state["rules_msg"])

COLOR_KEYWORDS = _parse_keyword_lines(color_keyword_text)
BW_KEYWORDS = _parse_keyword_lines(bw_keyword_text)


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

      # 1. 偵測頁面尺寸 (A3 面積或寬度顯著大於 A4)
      rect = page.rect
      width, height = rect.width, rect.height
      page_area = width * height
      is_a3 = page_area > 600000  # A3 面積門檻判斷

      # 2. 標題關鍵字規則＋實際色彩判定。
      has_real_color, rule_label, color_reason, c_hits, bw_hits = decide_page_color(
          page, t, image_color_cache
      )
      color_diagnostics.append({
          "頁碼": actual_page,
          "尺寸": "A3" if is_a3 else "A4",
          "結果": "🎨 彩色" if has_real_color else "⬛ 黑白",
          "依據": rule_label,
          "說明": color_reason,
          "彩色關鍵字": "、".join(c_hits),
          "黑白關鍵字": "、".join(bw_hits),
      })

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
          "圖表標題含彩色關鍵字（工作進度表、GNSS 等）固定彩色；含黑白關鍵字（雨量、風險矩陣）固定黑白；其餘依實際色彩判定。"
      )
      with st.expander("查看逐頁彩色判定診斷（確認誤判/漏判時使用）"):
        only_hit = st.checkbox("只顯示彩色頁或有關鍵字命中的頁面", value=False, key="diag_filter")
        rows = [
            r for r in color_diagnostics
            if not only_hit or r["結果"].startswith("🎨") or r["彩色關鍵字"] or r["黑白關鍵字"]
        ]
        st.dataframe(rows, use_container_width=True, hide_index=True)
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
