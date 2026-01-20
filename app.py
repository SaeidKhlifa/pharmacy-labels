import streamlit as st
import pandas as pd
import io
import os
import fitz  # PyMuPDF: Essential for fixing the Chrome Block issue
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.graphics.barcode import code128
import arabic_reshaper
from bidi.algorithm import get_display

# ==========================================
# 1. SETUP & FONTS
# ==========================================
st.set_page_config(page_title="Offers Generator", layout="wide", page_icon="🏷️")

FONT_PATH = "arial.ttf"
FONT_NAME = "CustomArial"

def setup_fonts():
    if os.path.exists(FONT_PATH):
        try:
            pdfmetrics.registerFont(TTFont(FONT_NAME, FONT_PATH))
            return True
        except:
            return False
    return False

has_font = setup_fonts()

def process_text(text, is_arabic=False):
    if pd.isna(text) or text == "": return ""
    text = str(text)
    if is_arabic and has_font:
        return get_display(arabic_reshaper.reshape(text))
    return text

# ==========================================
# 2. PDF GENERATION ENGINE
# ==========================================
def draw_label(c, x, y, w, h, row, settings):
    # Extract Data
    item_code = str(row.get('Item Number', '')).replace('.0', '')
    desc_en = row.get('Item Description EN', '')[:35]
    desc_ar = row.get('Item Description AR', '')
    brand_txt = row.get('Brand', '')
    offer_txt = row.get('Offer Description EN', '')

    center_x = x + (w / 2)
    
    # Debug Borders
    if settings['show_borders']:
        c.setLineWidth(0.5)
        c.setStrokeColorRGB(0.8, 0.8, 0.8)
        c.rect(x, y, w, h)

    # --- POSITIONING LOGIC ---
    # Top of the 8cm yellow area
    area_top_y = (y + h) - settings['top_offset_skip']
    
    # 1. Brand
    c.setFillColorRGB(0, 0, 0)
    c.setFont(FONT_NAME if has_font else "Helvetica-Bold", settings['brand_font_size'])
    brand_y = area_top_y - 15 
    c.drawCentredString(center_x, brand_y, str(brand_txt))

    # 2. English Name
    c.setFont(FONT_NAME if has_font else "Helvetica", settings['name_font_size'])
    name_en_y = brand_y - settings['spacing_brand_to_name']
    c.drawCentredString(center_x, name_en_y, str(desc_en))

    # 3. Arabic Name
    name_ar_y = name_en_y - settings['spacing_en_to_ar']
    ar_text = process_text(desc_ar, is_arabic=True)
    c.drawCentredString(center_x, name_ar_y, ar_text)

    # 4. Price (Optional)
    c.setFont(FONT_NAME if has_font else "Helvetica-Bold", settings['price_font_size'])
    c.setFillColorRGB(0.85, 0.21, 0.27)
    price_y = name_ar_y - settings['spacing_ar_to_price']
    c.drawCentredString(center_x, price_y, str(offer_txt))

    # 5. Barcode
    barcode_y = y + settings['barcode_bottom_margin']
    if item_code:
        try:
            bc_height = settings['barcode_height']
            barcode = code128.Code128(item_code, barHeight=bc_height, barWidth=1.2)
            bc_x = center_x - (barcode.width / 2)
            barcode.drawOn(c, bc_x, barcode_y + 12) 
            
            c.setFillColorRGB(0, 0, 0)
            c.setFont("Helvetica", settings['barcode_font_size'])
            c.drawCentredString(center_x, barcode_y, item_code)
        except:
            pass

def generate_pdf(df, settings):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    page_w, page_h = A4
    cols, rows = 3, 2
    block_w, block_h = page_w / cols, page_h / rows
    
    for i, (_, row) in enumerate(df.iterrows()):
        if i > 0 and i % (cols * rows) == 0:
            c.showPage()
        
        pos = i % (cols * rows)
        x = (pos % cols) * block_w
        y = page_h - ((pos // cols + 1) * block_h)
        draw_label(c, x, y, block_w, block_h, row, settings)
        
    c.save()
    buffer.seek(0)
    return buffer

# ==========================================
# 3. UI & FLOW
# ==========================================
st.title("🏷️ Offers Generator Pro")

if not has_font:
    st.warning("⚠️ Font `arial.ttf` missing. Arabic will look broken.")

# --- SIDEBAR (SETTINGS) ---
st.sidebar.header("1. Data")
offers_file = st.sidebar.file_uploader("Offers Excel", type=['xlsx'])
stock_file = st.sidebar.file_uploader("Stock Excel", type=['xlsx'])
min_qty = st.sidebar.number_input("Min Qty", 2, 100, 2)

st.sidebar.markdown("---")
st.sidebar.header("2. Design")
show_borders = st.sidebar.checkbox("Show Grid Lines", False)

with st.sidebar.expander("↕️ Positions", expanded=True):
    s_top_offset = st.slider("Top Offset", 0, 80, 50)
    s_bc_bottom = st.slider("Barcode Bottom", 0, 80, 20)
    s_brand_gap = st.slider("Gap: Brand->Name", 5, 50, 20)
    s_en_ar_gap = st.slider("Gap: En->Ar", 5, 50, 15)
    s_ar_price_gap = st.slider("Gap: Ar->Price", 5, 50, 20)

with st.sidebar.expander("🅰️ Font Sizes", expanded=False):
    s_brand_font = st.slider("Brand Size", 10, 24, 14)
    s_name_font = st.slider("Name Size", 8, 20, 11)
    s_price_font = st.slider("Price Size", 10, 50, 20)
    s_bc_h = st.slider("BC Height", 10, 50, 25)
    s_bc_font = st.slider("BC Num Size", 6, 14, 10)

user_settings = {
    'show_borders': show_borders, 'top_offset_skip': s_top_offset,
    'barcode_bottom_margin': s_bc_bottom, 'spacing_brand_to_name': s_brand_gap,
    'spacing_en_to_ar': s_en_ar_gap, 'spacing_ar_to_price': s_ar_price_gap,
    'brand_font_size': s_brand_font, 'name_font_size': s_name_font,
    'price_font_size': s_price_font, 'barcode_height': s_bc_h, 'barcode_font_size': s_bc_font
}

# --- MAIN AREA ---
if offers_file and stock_file:
    try:
        # Load Data
        df1 = pd.read_excel(offers_file)
        df2 = pd.read_excel(stock_file)
        
        # Clean & Merge
        df1['Item Number'] = df1['Item Number'].astype(str).str.replace('.0', '')
        df2['Item Number'] = df2['Item Number'].astype(str).str.replace('.0', '')
        merged = pd.merge(df1, df2[['Item Number', 'Quantity']], on='Item Number', how='left')
        final_df = merged[merged['Quantity'] >= min_qty].copy()

        if final_df.empty:
            st.error("No items found.")
        else:
            # Filters
            c1, c2, c3 = st.columns(3)
            cats = ['All'] + sorted(list(final_df['Category'].dropna().unique()))
            brands = ['All'] + sorted(list(final_df['Brand'].dropna().unique()))
            sel_cat = c1.selectbox("Category", cats)
            sel_brand = c2.selectbox("Brand", brands)
            
            if sel_cat != 'All': final_df = final_df[final_df['Category'] == sel_cat]
            if sel_brand != 'All': final_df = final_df[final_df['Brand'] == sel_brand]
            
            st.info(f"Ready to process **{len(final_df)}** labels.")
            st.markdown("---")

            # === THE PREVIEW STEP ===
            if st.button("Generate Preview (Page 1)", type="primary"):
                # 1. Generate ONLY first page for speed
                preview_pdf = generate_pdf(final_df.head(6), user_settings)
                
                # 2. Save PDF to session state (so it doesn't vanish)
                st.session_state['preview_pdf'] = preview_pdf
                st.session_state['data_hash'] = str(user_settings) # Track changes
            
            # If preview exists, show it
            if 'preview_pdf' in st.session_state:
                st.subheader("👁️ Preview Result")
                
                # Render PDF as IMAGE (Fixes Chrome Block)
                doc = fitz.open(stream=st.session_state['preview_pdf'].getvalue(), filetype="pdf")
                page = doc.load_page(0)
                pix = page.get_pixmap(dpi=150)
                st.image(pix.tobytes("png"), caption="First Page Preview", width=700)
                
                # === THE DOWNLOAD STEP ===
                st.markdown("---")
                st.success("Looks good?")
                
                # Button to generate FULL file
                if st.download_button(
                    label="📥 Download Full PDF Now",
                    data=generate_pdf(final_df, user_settings),
                    file_name="Final_Offers_Stock.pdf",
                    mime="application/pdf"
                ):
                    st.balloons()

    except Exception as e:
        st.error(f"Error: {e}")
