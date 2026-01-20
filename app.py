import streamlit as st
import pandas as pd
import io
import os
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

FONT_PATH = "arial.ttf"  # Ensure this file exists in the directory
FONT_NAME = "CustomArial"

def setup_fonts():
    """Registers the font for Arabic support."""
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
    """Handles Arabic reshaping and bidi algorithm."""
    if pd.isna(text) or text == "":
        return ""
    text = str(text)
    if is_arabic and has_font:
        reshaped = arabic_reshaper.reshape(text)
        return get_display(reshaped)
    return text

# ==========================================
# 2. PDF GENERATION ENGINE
# ==========================================
def draw_label(c, x, y, w, h, row):
    """Draws a single label at coordinates x, y."""
    # Data Extraction (Using Column Names from your Excel)
    item_code = str(row.get('Item Number', '')).replace('.0', '')
    desc_en = row.get('Item Description EN', '')[:35] # Truncate if too long
    desc_ar = row.get('Item Description AR', '')
    offer_txt = row.get('Offer Description EN', '')
    
    # Coordinates Calculation
    center_x = x + (w / 2)
    
    # 1. Header
    c.setLineWidth(0.5)
    c.rect(x, y, w, h) # Border
    
    c.setFont(FONT_NAME if has_font else "Helvetica", 8)
    c.setFillColorRGB(0.4, 0.4, 0.4)
    header_text = process_text("Al-Dawaa Pharmacy | صيدلية الدواء", is_arabic=True)
    c.drawCentredString(center_x, y + h - 15, header_text)
    
    # 2. English Name
    c.setFillColorRGB(0, 0, 0)
    c.setFont(FONT_NAME if has_font else "Helvetica", 11)
    c.drawCentredString(center_x, y + h - 35, str(desc_en))
    
    # 3. Arabic Name
    c.setFont(FONT_NAME if has_font else "Helvetica", 11)
    ar_text = process_text(desc_ar, is_arabic=True)
    c.drawCentredString(center_x, y + h - 50, ar_text)
    
    # 4. The Offer (Big Red Text)
    c.setFont(FONT_NAME if has_font else "Helvetica-Bold", 24)
    c.setFillColorRGB(0.85, 0.21, 0.27) # Red Color
    c.drawCentredString(center_x, y + (h/2) - 5, str(offer_txt))
    
    # 5. Barcode & Number
    if item_code:
        try:
            # Draw Barcode
            barcode = code128.Code128(item_code, barHeight=25, barWidth=1.2)
            # Center the barcode
            bc_x = center_x - (barcode.width / 2)
            barcode.drawOn(c, bc_x, y + 25)
            
            # Draw Number below barcode
            c.setFillColorRGB(0, 0, 0)
            c.setFont("Helvetica", 10)
            c.drawCentredString(center_x, y + 15, item_code)
        except:
            pass

def generate_pdf(df):
    """Main loop to create PDF pages."""
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    
    page_w, page_h = A4
    margin = 20
    cols = 3
    rows = 2
    
    block_w = (page_w - (2 * margin)) / cols
    block_h = (page_h - (2 * margin)) / rows
    
    for i, (_, row) in enumerate(df.iterrows()):
        # Check if we need a new page
        if i > 0 and i % (cols * rows) == 0:
            c.showPage()
            
        # Calculate Position
        pos_on_page = i % (cols * rows)
        col_idx = pos_on_page % cols
        row_idx = pos_on_page // cols
        
        # X grows left to right, Y grows bottom to top in PDF
        x = margin + (col_idx * block_w)
        # Invert row index for top-to-bottom filling
        y = page_h - margin - ((row_idx + 1) * block_h)
        
        draw_label(c, x, y, block_w, block_h, row)
        
    c.save()
    buffer.seek(0)
    return buffer

# ==========================================
# 3. STREAMLIT UI
# ==========================================
st.title("🏷️ Offers Generator Pro (Merged)")

if not has_font:
    st.warning("⚠️ Font file `arial.ttf` not found. Arabic text will not render correctly.")

# --- Sidebar: Inputs ---
st.sidebar.header("1. Upload Files")
offers_file = st.sidebar.file_uploader("Upload Offers (Excel)", type=['xlsx'])
stock_file = st.sidebar.file_uploader("Upload Stock (Excel)", type=['xlsx'])

st.sidebar.header("2. Settings")
min_qty = st.sidebar.number_input("Minimum Stock Qty", value=2, min_value=1)

# --- Main Logic ---
if offers_file and stock_file:
    try:
        # Load Data
        df_offers = pd.read_excel(offers_file)
        df_stock = pd.read_excel(stock_file)
        
        # Ensure Item Numbers are Strings for merging
        df_offers['Item Number'] = df_offers['Item Number'].astype(str).str.replace('.0', '')
        df_stock['Item Number'] = df_stock['Item Number'].astype(str).str.replace('.0', '')
        
        # Merge Logic (Left Join on Offers)
        merged_df = pd.merge(df_offers, df_stock[['Item Number', 'Quantity']], on='Item Number', how='left')
        
        # Filter by Quantity
        filtered_df = merged_df[merged_df['Quantity'] >= min_qty].copy()
        
        if filtered_df.empty:
            st.error("❌ No items match the minimum quantity requirement.")
        else:
            st.success(f"✅ Loaded {len(filtered_df)} items matching stock criteria.")
            
            # --- Dynamic Filters (Like the HTML Version) ---
            col1, col2, col3 = st.columns(3)
            
            with col1:
                # Category Filter
                cats = ['All'] + sorted(list(filtered_df['Category'].dropna().unique()))
                selected_cat = st.selectbox("Filter by Category", cats)

            with col2:
                # Brand Filter
                brands = ['All'] + sorted(list(filtered_df['Brand'].dropna().unique()))
                selected_brand = st.selectbox("Filter by Brand", brands)
                
            with col3:
                 # Offer Type Filter
                offer_types = ['All'] + sorted(list(filtered_df['Offer Description EN'].dropna().unique()))
                selected_offer = st.selectbox("Filter by Offer", offer_types)

            # Apply Filters
            final_df = filtered_df.copy()
            if selected_cat != 'All':
                final_df = final_df[final_df['Category'] == selected_cat]
            if selected_brand != 'All':
                final_df = final_df[final_df['Brand'] == selected_brand]
            if selected_offer != 'All':
                final_df = final_df[final_df['Offer Description EN'] == selected_offer]
                
            st.divider()
            st.subheader(f"🖨️ Ready to Print: {len(final_df)} Labels")
            st.dataframe(final_df.head())
            
            if st.button("Generate PDF", type="primary"):
                if len(final_df) > 0:
                    pdf_data = generate_pdf(final_df)
                    st.download_button(
                        label="📥 Download PDF Labels",
                        data=pdf_data,
                        file_name="Offers_Labels_Merged.pdf",
                        mime="application/pdf"
                    )
                else:
                    st.warning("No data to print based on current filters.")

    except Exception as e:
        st.error(f"An error occurred: {e}")
else:
    st.info("👋 Please upload both Offers and Stock Excel files in the sidebar to begin.")
