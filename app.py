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
st.set_page_config(page_title="Offers Generator Pro", layout="wide", page_icon="🏷️")

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
# 2. PDF ENGINE
# ==========================================
def draw_label(c, x, y, w, h, row, settings):
    # Extract Data
    item_code = str(row.get('Item Number', '')).replace('.0', '')
    desc_en = row.get('Item Description EN', '')[:35]
    desc_ar = row.get('Item Description AR', '')
    offer_txt = row.get('Offer Description EN', '')

    center_x = x + (w / 2)
    
    # Border
    c.setLineWidth(0.5)
    c.rect(x, y, w, h)

    # 1. Header
    header_y_pos = y + h - settings['header_margin_top']
    c.setFont(FONT_NAME if has_font else "Helvetica", settings['header_font_size'])
    c.setFillColorRGB(0.4, 0.4, 0.4)
    header_text = process_text("Al-Dawaa Pharmacy | صيدلية الدواء", is_arabic=True)
    c.drawCentredString(center_x, header_y_pos, header_text)

    # 2. English Name
    en_name_y_pos = header_y_pos - settings['spacing_header_to_name']
    c.setFillColorRGB(0, 0, 0)
    c.setFont(FONT_NAME if has_font else "Helvetica", settings['name_font_size'])
    c.drawCentredString(center_x, en_name_y_pos, str(desc_en))

    # 3. Arabic Name
    ar_name_y_pos = en_name_y_pos - settings['spacing_en_to_ar']
    ar_text = process_text(desc_ar, is_arabic=True)
    c.drawCentredString(center_x, ar_name_y_pos, ar_text)

    # 4. Offer
    offer_y_pos = y + (h / 2) + settings['offer_vertical_shift']
    c.setFont(FONT_NAME if has_font else "Helvetica-Bold", settings['offer_font_size'])
    c.setFillColorRGB(0.85, 0.21, 0.27) 
    c.drawCentredString(center_x, offer_y_pos, str(offer_txt))

    # 5. Barcode
    barcode_base_y = y + settings['barcode_margin_bottom']
    if item_code:
        try:
            bc_height = settings['barcode_height']
            barcode = code128.Code128(item_code, barHeight=bc_height, barWidth=1.2)
            bc_x = center_x - (barcode.width / 2)
            barcode.drawOn(c, bc_x, barcode_base_y + 10) 
            
            c.setFillColorRGB(0, 0, 0)
            c.setFont("Helvetica", settings['barcode_font_size'])
            c.drawCentredString(center_x, barcode_base_y, item_code)
        except:
            pass

def generate_pdf(df, settings):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    
    page_w, page_h = A4
    margin = 20
    cols = 3
    rows = 2
    
    block_w = (page_w - (2 * margin)) / cols
    block_h = (page_h - (2 * margin)) / rows
    
    for i, (_, row) in enumerate(df.iterrows()):
        if i > 0 and i % (cols * rows) == 0:
            c.showPage()
            
        pos_on_page = i % (cols * rows)
        col_idx = pos_on_page % cols
        row_idx = pos_on_page // cols
        
        x = margin + (col_idx * block_w)
        y = page_h - margin - ((row_idx + 1) * block_h)
        
        draw_label(c, x, y, block_w, block_h, row, settings)
        
    c.save()
    buffer.seek(0)
    return buffer

def display_pdf_preview(pdf_bytes):
    """
    Embeds the PDF binary in an HTML iframe for preview.
    """
    base64_pdf = base64.b64encode(pdf_bytes.getvalue()).decode('utf-8')
    pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="100%" height="600px" type="application/pdf"></iframe>'
    st.markdown(pdf_display, unsafe_allow_html=True)

# ==========================================
# 3. UI
# ==========================================
st.title("🏷️ Offers Generator Pro")

if not has_font:
    st.warning("⚠️ Font `arial.ttf` missing. Arabic will look wrong.")

# Layout: Split Sidebar and Main
col_controls, col_preview = st.columns([1, 1.5])

with col_controls:
    st.header("1. Data & Filters")
    offers_file = st.file_uploader("Upload Offers", type=['xlsx'])
    stock_file = st.file_uploader("Upload Stock", type=['xlsx'])
    min_qty = st.number_input("Min Qty", value=2, min_value=1)

    st.markdown("---")
    st.header("2. Design Controls")
    
    with st.expander("🅰️ Font Sizes", expanded=False):
        s_header_font = st.slider("Header Size", 6, 14, 8)
        s_name_font = st.slider("Name Size", 8, 18, 11)
        s_offer_font = st.slider("Offer Size", 15, 60, 24)
        s_bc_font = st.slider("BC Num Size", 6, 14, 10)

    with st.expander("📏 Margins & Spacing", expanded=True):
        s_head_top = st.slider("Header Top", 5, 40, 15)
        s_head_name_gap = st.slider("Gap: Header->Name", 5, 40, 20)
        s_en_ar_gap = st.slider("Gap: En->Ar", 5, 30, 15)
        s_offer_shift = st.slider("Offer Shift Y", -30, 30, -5)
        s_bc_bottom = st.slider("Barcode Bottom", 5, 50, 15)
        s_bc_height = st.slider("Barcode Height", 10, 50, 25)

user_settings = {
    'header_font_size': s_header_font, 'name_font_size': s_name_font,
    'offer_font_size': s_offer_font, 'barcode_font_size': s_bc_font,
    'header_margin_top': s_head_top, 'spacing_header_to_name': s_head_name_gap,
    'spacing_en_to_ar': s_en_ar_gap, 'offer_vertical_shift': s_offer_shift,
    'barcode_margin_bottom': s_bc_bottom, 'barcode_height': s_bc_height
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
            st.error("❌ No items match filters.")
        else:
            # Filters inside the left column
            with col_controls:
                cat_filter = st.selectbox("Category", ['All'] + sorted(list(final_df['Category'].dropna().unique())))
                brand_filter = st.selectbox("Brand", ['All'] + sorted(list(final_df['Brand'].dropna().unique())))
                
                if cat_filter != 'All': final_df = final_df[final_df['Category'] == cat_filter]
                if brand_filter != 'All': final_df = final_df[final_df['Brand'] == brand_filter]

                st.success(f"Selected: {len(final_df)} items")
                
                # Full Download Button
                if st.button("📥 Download Full PDF", type="primary"):
                    full_pdf = generate_pdf(final_df, user_settings)
                    st.download_button("Click to Save PDF", full_pdf, "Labels.pdf", "application/pdf")

            # --- PREVIEW SECTION (Right Column) ---
            with col_preview:
                st.subheader("👁️ Live Preview (First Page)")
                # Generate PDF for ONLY the first 6 items for speed
                preview_data = final_df.head(6) 
                
                if not preview_data.empty:
                    # Generate temporary PDF bytes
                    preview_pdf = generate_pdf(preview_data, user_settings)
                    # Display using iframe
                    display_pdf_preview(preview_pdf)
                else:
                    st.info("No data to preview.")

    except Exception as e:
        st.error(f"Error: {e}")
