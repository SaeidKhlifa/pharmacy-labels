import streamlit as st
import pandas as pd
import io
import os
import base64
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.graphics.barcode import code128
import arabic_reshaper
from bidi.algorithm import get_display

# ==========================================
# 1. CONFIGURATION & FONTS
# ==========================================
st.set_page_config(page_title="Offers Generator - Pre-Printed", layout="wide", page_icon="🏷️")

FONT_PATH = "arial.ttf" 
FONT_NAME = "CustomArial"

def setup_fonts():
    try:
        if os.path.exists(FONT_PATH):
            pdfmetrics.registerFont(TTFont(FONT_NAME, FONT_PATH))
            return True
        else:
            return False
    except Exception:
        return False

has_font = setup_fonts()

def process_text(text, is_arabic=False):
    if pd.isna(text) or text == "": return ""
    text = str(text)
    if is_arabic and has_font:
        reshaped = arabic_reshaper.reshape(text)
        return get_display(reshaped)
    return text

# ==========================================
# 2. PDF ENGINE (ADAPTED FOR PRE-PRINTED PAPER)
# ==========================================
def draw_label(c, x, y, w, h, row, settings):
    # Extract Data
    item_code = str(row.get('Item Number', '')).replace('.0', '')
    desc_en = row.get('Item Description EN', '')[:35]
    desc_ar = row.get('Item Description AR', '')
    brand_txt = row.get('Brand', '')
    offer_txt = row.get('Offer Description EN', '') # Keeping offer price just in case

    center_x = x + (w / 2)
    
    # --- DEBUG BORDER (Optional) ---
    if settings['show_borders']:
        c.setLineWidth(0.5)
        c.setStrokeColorRGB(0.8, 0.8, 0.8) # Light gray for debugging
        c.rect(x, y, w, h)

    # --- CALCULATION OF PRINT AREA ---
    # The 'y' parameter is the BOTTOM of the cell. 'h' is the height (half page).
    # The cell Top is y + h.
    # We need to skip the Top Margin (White strip + Red Header).
    
    # 1. Start of the 8cm Writing Area (Top limit)
    area_top_y = (y + h) - settings['top_offset_skip']
    
    # 2. BRAND NAME (At the top of the 8cm area)
    c.setFillColorRGB(0, 0, 0)
    c.setFont(FONT_NAME if has_font else "Helvetica-Bold", settings['brand_font_size'])
    # Draw slightly below the top of the 8cm area
    brand_y = area_top_y - 15 
    c.drawCentredString(center_x, brand_y, str(brand_txt))

    # 3. ENGLISH NAME (Below Brand)
    c.setFont(FONT_NAME if has_font else "Helvetica", settings['name_font_size'])
    name_en_y = brand_y - settings['spacing_brand_to_name']
    c.drawCentredString(center_x, name_en_y, str(desc_en))

    # 4. ARABIC NAME (Below English Name)
    name_ar_y = name_en_y - settings['spacing_en_to_ar']
    ar_text = process_text(desc_ar, is_arabic=True)
    c.drawCentredString(center_x, name_ar_y, ar_text)

    # 5. PRICE / OFFER (Optional - Middle of space)
    # If you want the price to appear, uncomment the lines below. 
    # Based on the prompt, user asked for names/barcode, but usually price is needed.
    c.setFont(FONT_NAME if has_font else "Helvetica-Bold", settings['price_font_size'])
    c.setFillColorRGB(0.85, 0.21, 0.27) # Red
    price_y = name_ar_y - settings['spacing_ar_to_price']
    c.drawCentredString(center_x, price_y, str(offer_txt))

    # 6. BARCODE & NUMBER (At the bottom of the 8cm area)
    # We calculate position from the BOTTOM of the 8cm area upwards.
    # The bottom of the 8cm area is roughly (area_top_y - 8cm).
    # Or we can just use the global 'y' (bottom of paper section) + a margin.
    
    barcode_y = y + settings['barcode_bottom_margin']
    
    if item_code:
        try:
            # Barcode
            bc_height = settings['barcode_height']
            barcode = code128.Code128(item_code, barHeight=bc_height, barWidth=1.2)
            bc_x = center_x - (barcode.width / 2)
            barcode.drawOn(c, bc_x, barcode_y + 12) 
            
            # Number (Below Barcode)
            c.setFillColorRGB(0, 0, 0)
            c.setFont("Helvetica", settings['barcode_font_size'])
            c.drawCentredString(center_x, barcode_y, item_code)
        except:
            pass

def generate_pdf(df, settings):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    
    page_w, page_h = A4
    margin = 0 # No side margins, the paper is full width
    
    # 3 Columns, 2 Rows (Top half / Bottom half)
    cols = 3
    rows = 2
    
    block_w = page_w / cols
    block_h = page_h / rows
    
    for i, (_, row) in enumerate(df.iterrows()):
        if i > 0 and i % (cols * rows) == 0:
            c.showPage()
            
        pos_on_page = i % (cols * rows)
        col_idx = pos_on_page % cols
        row_idx = pos_on_page // cols
        
        x = col_idx * block_w
        # y is calculated from bottom up. 
        # Row 0 is Top (y = 148.5mm), Row 1 is Bottom (y = 0mm)
        y = page_h - ((row_idx + 1) * block_h)
        
        draw_label(c, x, y, block_w, block_h, row, settings)
        
    c.save()
    buffer.seek(0)
    return buffer

def display_pdf_preview(pdf_bytes):
    base64_pdf = base64.b64encode(pdf_bytes.getvalue()).decode('utf-8')
    pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="100%" height="600px" type="application/pdf"></iframe>'
    st.markdown(pdf_display, unsafe_allow_html=True)

# ==========================================
# 3. UI LAYOUT
# ==========================================
st.title("🏷️ Offers Generator (Stock Paper Edition)")

if not has_font:
    st.warning("⚠️ Font `arial.ttf` missing. Arabic text will fail.")

col_controls, col_preview = st.columns([1, 1.2])

with col_controls:
    st.subheader("1. Data Input")
    offers_file = st.file_uploader("Upload Offers (Excel)", type=['xlsx'])
    stock_file = st.file_uploader("Upload Stock (Excel)", type=['xlsx'])
    min_qty = st.number_input("Min Stock Qty", value=2, min_value=1)

    st.markdown("---")
    st.subheader("2. Layout Calibration")
    st.info("Adjust these sliders to match your pre-printed paper.")

    show_borders = st.checkbox("Show Grid Lines (For debugging)", value=False)
    
    with st.expander("↕️ Vertical Positions (Critical)", expanded=True):
        # Default ~ 148mm total height. Header is usually 30-40mm.
        s_top_offset = st.slider("Top Offset (Skip Red Header)", 0, 80, 50, help="Distance from the very top of the paper section to where text starts.")
        s_bc_bottom = st.slider("Barcode Position (From Bottom)", 0, 80, 20)
        
    with st.expander("📏 Spacing Between Items", expanded=True):
        s_brand_gap = st.slider("Gap: Brand -> En Name", 5, 50, 20)
        s_en_ar_gap = st.slider("Gap: En Name -> Ar Name", 5, 50, 15)
        s_ar_price_gap = st.slider("Gap: Ar Name -> Price", 5, 50, 20)
    
    with st.expander("🅰️ Font Sizes", expanded=False):
        s_brand_font = st.slider("Brand Size", 10, 24, 14)
        s_name_font = st.slider("Product Name Size", 8, 20, 11)
        s_price_font = st.slider("Price Size", 10, 50, 20)
        s_bc_h = st.slider("Barcode Height", 10, 50, 25)
        s_bc_font = st.slider("Barcode Num Size", 6, 14, 10)

user_settings = {
    'show_borders': show_borders,
    'top_offset_skip': s_top_offset,      # Pushes text down below red header
    'barcode_bottom_margin': s_bc_bottom, # Pushes barcode up from bottom edge
    'spacing_brand_to_name': s_brand_gap,
    'spacing_en_to_ar': s_en_ar_gap,
    'spacing_ar_to_price': s_ar_price_gap,
    'brand_font_size': s_brand_font,
    'name_font_size': s_name_font,
    'price_font_size': s_price_font,
    'barcode_height': s_bc_h,
    'barcode_font_size': s_bc_font
}

# --- MAIN LOGIC ---
if offers_file and stock_file:
    try:
        df_offers = pd.read_excel(offers_file)
        df_stock = pd.read_excel(stock_file)
        
        # Merge
        df_offers['Item Number'] = df_offers['Item Number'].astype(str).str.replace('.0', '')
        df_stock['Item Number'] = df_stock['Item Number'].astype(str).str.replace('.0', '')
        merged_df = pd.merge(df_offers, df_stock[['Item Number', 'Quantity']], on='Item Number', how='left')
        final_df = merged_df[merged_df['Quantity'] >= min_qty].copy()

        if final_df.empty:
            st.error("❌ No items match.")
        else:
            with col_controls:
                # Filters
                cat_list = ['All'] + sorted(list(final_df['Category'].dropna().unique()))
                brand_list = ['All'] + sorted(list(final_df['Brand'].dropna().unique()))
                
                sel_cat = st.selectbox("Category", cat_list)
                sel_brand = st.selectbox("Brand", brand_list)
                
                if sel_cat != 'All': final_df = final_df[final_df['Category'] == sel_cat]
                if sel_brand != 'All': final_df = final_df[final_df['Brand'] == sel_brand]

                st.success(f"Printing {len(final_df)} labels")
                
                if st.button("📥 Download Final PDF", type="primary"):
                    full_pdf = generate_pdf(final_df, user_settings)
                    st.download_button("Click to Save", full_pdf, "Labels_Stock.pdf", "application/pdf")

            with col_preview:
                st.subheader("Live Preview")
                # Preview first 6 items
                preview_data = final_df.head(6)
                if not preview_data.empty:
                    preview_pdf = generate_pdf(preview_data, user_settings)
                    display_pdf_preview(preview_pdf)

    except Exception as e:
        st.error(f"Error: {e}")
