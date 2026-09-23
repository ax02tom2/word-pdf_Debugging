import streamlit as st
from PIL import Image, ImageDraw, ImageFont
import os
import urllib.request

st.title("📁 影像時序 GIF 動態圖製作工具")
st.write("請直接上傳依照日期命名的圖片，先即時預覽浮水印效果，再一鍵產出 GIF。")

# --- 側邊欄設定 ---
st.sidebar.header("⚙️ 播放設定")
duration_ms = st.sidebar.slider("每張圖片停留時間 (毫秒)", 200, 3000, 1000, 100)

st.sidebar.markdown("---")
st.sidebar.header("📝 日期浮水印設定")
add_watermark = st.sidebar.checkbox("自動加入日期浮水印", value=True)

if add_watermark:
    text_color = st.sidebar.color_picker("文字顏色", "#FF0000")
    font_size = st.sidebar.slider("文字大小", 10, 150, 40, step=2)
    text_pos = st.sidebar.selectbox("文字位置", ["右下", "左下", "右上", "左上"])
    margin_pct = st.sidebar.slider("邊距往內縮 (百分比 %)", 1, 40, 5, step=1)
else:
    text_color, font_size, text_pos, margin_pct = "#FF0000", 40, "右下", 5

# --- 自動下載清晰字型 ---
font_path = "Roboto-Bold.ttf"
if not os.path.exists(font_path):
    try:
        url = "https://github.com/googlefonts/roboto/raw/main/src/hinted/Roboto-Bold.ttf"
        urllib.request.urlretrieve(url, font_path)
    except:
        pass

# --- 浮水印繪製工具 ---
def apply_watermark(img, file_name_str, f_size, t_color, t_pos, m_pct):
    draw = ImageDraw.Draw(img)
    file_name = os.path.splitext(file_name_str)[0]
    
    if len(file_name) == 7 and file_name.isdigit():
        display_text = f"{file_name[:3]} / {file_name[3:5]} / {file_name[5:]}"
    else:
        display_text = file_name
    
    try:
        font = ImageFont.truetype(font_path, f_size)
    except:
        font = ImageFont.load_default()

    w, h = img.size
    try:
        left, top, right, bottom = font.getbbox(display_text)
        tw, th = right - left, bottom - top
    except:
        tw, th = f_size * len(display_text) * 0.6, f_size

    margin_x = int(w * (m_pct / 100.0))
    margin_y = int(h * (m_pct / 100.0))
    
    if t_pos == "右下": xy = (w - tw - margin_x, h - th - margin_y)
    elif t_pos == "左下": xy = (margin_x, h - th - margin_y)
    elif t_pos == "右上": xy = (w - tw - margin_x, margin_y)
    else: xy = (margin_x, margin_y)

    shadow_offset = max(1, int(f_size * 0.08))
    draw.text((xy[0]+shadow_offset, xy[1]+shadow_offset), display_text, fill="black", font=font)
    draw.text(xy, display_text, fill=t_color, font=font)
    
    return img

# --- 主程式區塊 ---
uploaded_files = st.file_uploader("請選擇或拖曳圖片檔案 (JPG/PNG)", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

if uploaded_files:
    # --- 🛡️ 檔案大小防呆計算機制 ---
    total_size_bytes = sum(getattr(f, "size", 0) for f in uploaded_files)
    total_size_mb = total_size_bytes / (1024 * 1024)
    MAX_LIMIT_MB = 1000.0  # 1GB 門檻

    # 1. 顯示檔案總大小與數量
    st.info(f"📊 已上傳 **{len(uploaded_files)}** 張圖片，總容量：**{total_size_mb:.2f} MB** / {int(MAX_LIMIT_MB)} MB")

    # 2. 超過 1GB 嚴格攔截，中斷程式執行
    if total_size_mb > MAX_LIMIT_MB:
        st.error(
            f"❌ 【容量超標警告】上傳檔案總大小為 **{total_size_mb:.2f} MB**，已超過 **{int(MAX_LIMIT_MB)} MB (1 GB)** 限制！\n\n"
            "為避免伺服器記憶體崩潰，系統已強制暫停處理。請刪除部分圖片或壓縮後重新上傳。"
        )
        st.stop()  # 立即停止下方所有運算

    # 3. 接近記憶體上限時提供黃色溫馨提醒
    elif total_size_mb > 700.0:
        st.warning(f"⚠️ 提醒：目前檔案總大小較大 ({total_size_mb:.2f} MB)，生成 GIF 可能需耗時數十秒，請耐心等候。")

    # --- 圖片排序與即時預覽 ---
    uploaded_files = sorted(uploaded_files, key=lambda x: x.name)
    
    st.markdown("### 👁️ 浮水印即時預覽 (第一張圖片)")
    preview_file = uploaded_files[0]
    preview_file.seek(0)
    preview_img = Image.open(preview_file).convert("RGBA")
    
    # 限制預覽圖解析度避免卡頓
    preview_img.thumbnail((1200, 1200), Image.Resampling.LANCZOS)
    
    if add_watermark:
        preview_img = apply_watermark(preview_img, preview_file.name, font_size, text_color, text_pos, margin_pct)
    
    st.image(preview_img.convert("RGB"), caption="💡 調整左側拉桿，此預覽圖會即時輕量化更新！", use_container_width=True)
    
    st.markdown("---")
    
    # --- 產出 GIF 區塊 ---
    if st.button("🚀 確認預覽無誤，開始製作 GIF"):
        with st.spinner("正在合成全部圖片中，請稍候..."):
            images = []
            base_size = None
            
            for idx, f in enumerate(uploaded_files):
                f.seek(0)
                img = Image.open(f).convert("RGBA")
                
                if idx == 0:
                    img.thumbnail((1200, 1200), Image.Resampling.LANCZOS)
                    base_size = img.size
                else:
                    img = img.resize(base_size, Image.Resampling.LANCZOS)
                
                if add_watermark:
                    img = apply_watermark(img, f.name, font_size, text_color, text_pos, margin_pct)
                images.append(img.convert("RGB"))
            
            output_path = "slideshow.gif"
            images[0].save(
                output_path,
                save_all=True,
                append_images=images[1:],
                duration=duration_ms,
                loop=0
            )
            
            with open(output_path, "rb") as file:
                gif_bytes = file.read()
                
        st.success("🎉 GIF 製作完成！")
        st.download_button(label="📥 下載您的 GIF 檔案", data=gif_bytes, file_name="slideshow.gif", mime="image/gif")
else:
    st.info("💡 請先由上方按鈕上傳圖片。")
