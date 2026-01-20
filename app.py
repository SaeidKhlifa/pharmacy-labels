import streamlit as st
import pandas as pd
import io
import os
import fitz  # PyMuPDF
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
st.set_page_config(page_title="Offers Generator Pro", layout="wide", page_icon="🏷️")

FONT_PATH = "arial.ttf"
FONT_NAME = "CustomArial"

def setup_fonts():
    """Registers the font for Arabic support."""
    if os.path.exists(FONT_PATH):
        try:
            pdfmetrics.registerFont(TTFont(FONT_NAME, FONT_PATH))
            return True
        except:
            return False
    return False

has_font = setup_fonts()

def process_text(text, is_arabic=False):
    """Handles Arabic reshaping and bidi algorithm."""
    if pd.isna(text) or text == "": return ""
    text = str(text)
    if is_arabic and has_font:
        reshaped = arabic_reshaper.reshape(text)
        return get_display(reshaped)
    return text

# ==========================================
# 2. PDF GENERATION ENGINE
# ==========================================
def draw_label(c, x, y, w, h, row, settings):
    """Draws a single label on the PDF canvas."""
    # Extract Data
    item_code = str(row.get('Item Number', '')).replace('.0', '')
    desc_en = row.get('Item Description EN', '')[:35]
    desc_ar = row.get('Item Description AR', '')
    brand_txt = row.get('Brand', '')
    offer_txt = row.get('Offer Description EN', '')

    center_x = x + (w / 2)
    
    # --- DEBUG BORDER (Optional) ---
    if settings['show_borders']:
        c.setLineWidth(0.5)
        c.setStrokeColorRGB(0.8, 0.8, 0.8) # Light gray
        c.rect(x, y, w, h)

    # --- POSITIONING LOGIC ---
    # We calculate position relative to the BOTTOM of the label area (y)
    # The label height is h. So Top is (y + h).
    
    # 1. Start of the 8cm Writing Area (Top limit)
    # 'top_offset_skip' skips the red header on your paper
    area_top_y = (y + h) - settings['top_offset_skip']
    
    # 2. BRAND NAME (At the top of the 8cm area)
    c.setFillColorRGB(0, 0, 0)
    c.setFont(FONT_NAME if has_font else "Helvetica-Bold", settings['brand_font_size'])
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
    c.setFont(FONT_NAME if has_font else "Helvetica-Bold", settings['price_font_size'])
    c.setFillColorRGB(0.85, 0.21, 0.27) # Red
    price_y = name_ar_y - settings['spacing_ar_to_price']
    c.drawCentredString(center_x, price_y, str(offer_txt))

    # 6. BARCODE & NUMBER (At the bottom of the 8cm area)
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
    """Generates the full PDF file in memory."""
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    
    page_w, page_h = A4
    margin = 0 # No side margins as per stock paper design
    
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
        # Row 0 is Top, Row 1 is Bottom
        y = page_h - ((row_idx + 1) * block_h)
        
        draw_label(c, x, y, block_w, block_h, row, settings)
        
    c.save()
    buffer.seek(0)
    return buffer

# ==========================================
# 3. UI LAYOUT & LOGIC
# ==========================================
st.title("🏷️ Offers Generator Pro (Stock Paper)")

if not has_font:
    st.warning("⚠️ Font `arial.ttf` missing. Arabic text will not display correctly.")

# --- SIDEBAR: SETTINGS ---
st.sidebar.header("1. Data Input")
offers_file = st.sidebar.file_uploader("Upload Offers (Excel)", type=['xlsx'])
stock_file = st.sidebar.file_uploader("Upload Stock (Excel)", type=['xlsx'])
min_qty = st.sidebar.number_input("Minimum Stock Qty", value=2, min_value=1)

st.sidebar.markdown("---")
st.sidebar.header("2. Design Calibration")
show_borders = st.sidebar.checkbox("Show Grid Lines (For debugging)", False)

with st.sidebar.expander("↕️ Vertical Positions", expanded=True):
    st.caption("Adjust to match your pre-printed paper lines")
    s_top_offset = st.slider("Top Offset (Skip Header)", 0, 100, 50, help="Push text down below the red strip")
    s_bc_bottom = st.slider("Barcode Position (From Bottom)", 0, 80, 20)

with st.sidebar.expander("📏 Spacing Between Items", expanded=True):
    s_brand_gap = st.slider("Gap: Brand -> En Name", 5, 50, 20)
    s_en_ar_gap = st.slider("Gap: En Name -> Ar Name", 5, 50, 15)
    s_ar_price_gap = st.slider("Gap: Ar Name -> Price", 5, 50, 20)

with st.sidebar.expander("🅰️ Font Sizes", expanded=False):
    s_brand_font = st.slider("Brand Size", 10, 24, 14)
    s_name_font = st.slider("Product Name Size", 8, 20, 11)
    s_price_font = st.slider("Price Size", 10, 50, 20)
    s_bc_h = st.slider("Barcode Height", 10, 50, 25)
    s_bc_font = st.slider("Barcode Num Size", 6, 14, 10)

# Pack settings
user_settings = {
    'show_borders': show_borders,
    'top_offset_skip': s_top_offset,
    'barcode_bottom_margin': s_bc_bottom,
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
        # Load Files
        df_offers = pd.read_excel(offers_file)
        df_stock = pd.read_excel(stock_file)
        
        # Clean Keys for Merging
        df_offers['Item Number'] = df_offers['Item Number'].astype(str).str.replace('.0', '')
        df_stock['Item Number'] = df_stock['Item Number'].astype(str).str.replace('.0', '')
        
        # Merge & Filter
        merged_df = pd.merge(df_offers, df_stock[['Item Number', 'Quantity']], on='Item Number', how='left')
        final_df = merged_df[merged_df['Quantity'] >= min_qty].copy()

        if final_df.empty:
            st.error("❌ No items match the stock requirements.")
        else:
            # --- FILTERS SECTION ---
            st.markdown("### 🔍 Filter Results")
            c1, c2, c3 = st.columns(3)
            
            # Populate dropdowns
            cats = ['All'] + sorted(list(final_df['Category'].dropna().unique()))
            brands = ['All'] + sorted(list(final_df['Brand'].dropna().unique()))
            
            sel_cat = c1.selectbox("Filter Category", cats)
            sel_brand = c2.selectbox("Filter Brand", brands)
            
            # Apply Filtering
            if sel_cat != 'All': final_df = final_df[final_df['Category'] == sel_cat]
            if sel_brand != 'All': final_df = final_df[final_df['Brand'] == sel_brand]
            
            st.info(f"✅ Ready to generate **{len(final_df)}** labels.")
            st.markdown("---")

            # --- PREVIEW BUTTON ---
            if st.button("👁️ Generate Preview (First Page)", type="primary"):
                # Generate PDF for first 6 items only (1 page)
                preview_pdf = generate_pdf(final_df.head(6), user_settings)
                st.session_state['preview_pdf'] = preview_pdf
            
            # --- PREVIEW DISPLAY ---
            if 'preview_pdf' in st.session_state:
                st.subheader("📄 Preview")
                
                # Convert PDF Page to Image for Display (Bypasses Chrome Blocking)
                doc = fitz.open(stream=st.session_state['preview_pdf'].getvalue(), filetype="pdf")
                page = doc.load_page(0)  # First page
                pix = page.get_pixmap(dpi=150) # Render image
                st.image(pix.tobytes("png"), caption="Live Preview of First Page", width=700)
                
                # --- DOWNLOAD BUTTON ---
                st.markdown("---")
                st.success("Does the preview look correct?")
                
                # Generate Full PDF only when requested
                full_pdf_data = generate_pdf(final_df, user_settings)
                st.download_button(
                    label="📥 Download Full PDF",
                    data=full_pdf_data,
                    file_name="Offers_Labels_Stock.pdf",
                    mime="application/pdf"
                )

    except Exception as e:
        st.error(f"Error processing files: {e}")
else:
    st.info("👋 Please upload your Excel files in the sidebar to begin.")
