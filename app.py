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
from reportlab.lib.utils import simpleSplit
import arabic_reshaper
from bidi.algorithm import get_display

# ==========================================
# 1. إعدادات القياسات (بالسنتيمتر)
# ==========================================
DIM_ROW1_TOP_CM = 7.7    
DIM_ROW2_TOP_CM = 22.5   
DIM_YELLOW_H_CM = 7.5    

CENTERS_FROM_RIGHT_CM = [3.5, 10.5, 17.9]

POS_BRAND_Y_CM = 0.6
POS_EN_Y_CM = 1.6
POS_AR_Y_CM = 2.6
POS_BARCODE_BOTTOM_CM = 0.8

FONT_PATH = "arial.ttf"
FONT_NAME = "CustomArial"
TEMPLATE_PATH = "template.png"

st.set_page_config(page_title="Offers Generator (Fixed)", layout="wide", page_icon="✅")

def cm2p(cm):
    return cm * 28.3465

def setup_resources():
    font_ok = False
    if os.path.exists(FONT_PATH):
        try:
            pdfmetrics.registerFont(TTFont(FONT_NAME, FONT_PATH))
            font_ok = True
        except:
            pass
    template_ok = os.path.exists(TEMPLATE_PATH)
    return font_ok, template_ok

has_font, has_template = setup_resources()

def process_text(text, is_arabic=False):
    if pd.isna(text) or text == "": return ""
    text = str(text)
    if is_arabic and has_font:
        reshaped = arabic_reshaper.reshape(text)
        return get_display(reshaped)
    return text

# ==========================================
# 2. دوال الرسم (تم إصلاح الخطأ هنا)
# ==========================================

def draw_text_multiline(c, text, y_center, max_width, font_name, max_font_size, min_font_size=6, color=(0,0,0), is_bold=False):
    """
    رسم نص ذكي (سطر أو سطرين) مع إصلاح خطأ الـ Canvas Attribute
    """
    c.setFillColorRGB(*color)
    c.setStrokeColorRGB(*color)
    
    current_size = max_font_size
    lines = []
    
    # 1. حساب حجم الخط المناسب
    while current_size >= min_font_size:
        lines = simpleSplit(text, font_name, current_size, max_width)
        if len(lines) <= 2:
            break
        current_size -= 0.5
        
    # 2. حساب موقع البداية الرأسي لتوسيط الكتلة النصية
    leading = current_size * 1.2
    total_height = leading * len(lines)
    start_y = y_center + (total_height / 2) - (current_size * 0.8)

    # 3. رسم كل سطر
    for line in lines:
        text_width = pdfmetrics.stringWidth(line, font_name, current_size)
        start_x = -(text_width / 2) # التوسيط الأفقي
        
        if is_bold:
            # طريقة الرسم "السميك" (Fake Bold)
            c.setLineWidth(0.5 if current_size < 12 else 0.8)
            text_obj = c.beginText()
            text_obj.setFont(font_name, current_size)
            text_obj.setTextRenderMode(2) # تفعيل التعبئة + الحدود (Fill + Stroke)
            text_obj.setTextOrigin(start_x, start_y)
            text_obj.textOut(line)
            c.drawText(text_obj)
            c.setLineWidth(0) # إعادة الحدود للوضع الطبيعي
        else:
            # طريقة الرسم العادي
            c.setFont(font_name, current_size)
            c.drawString(start_x, start_y, line)
            
        start_y -= leading # الانتقال للسطر التالي

    # تنظيف الإعدادات
    c.setFillColorRGB(0, 0, 0)
    c.setStrokeColorRGB(0, 0, 0)

def draw_card_content(c, row, settings):
    item_code = str(row.get('Item Number', '')).replace('.0', '')
    desc_en = row.get('Item Description EN', '') 
    desc_ar = row.get('Item Description AR', '')
    brand_txt = row.get('Brand', '')
    offer_txt = row.get('Offer Description EN', '')

    height = cm2p(DIM_YELLOW_H_CM)
    max_text_width = cm2p(7.0) * 0.90 

    # 1. البراند
    brand_y = -cm2p(POS_BRAND_Y_CM)
    if has_font:
        draw_text_multiline(c, str(brand_txt), brand_y, max_text_width, 
                            FONT_NAME, settings['font_brand'], min_font_size=8, is_bold=True)
    else:
        c.setFont("Helvetica-Bold", settings['font_brand'])
        c.drawCentredString(0, brand_y, str(brand_txt))

    # 2. الاسم الإنجليزي
    en_y = -cm2p(POS_EN_Y_CM)
    draw_text_multiline(c, str(desc_en), en_y, max_text_width, 
                        FONT_NAME if has_font else "Helvetica", settings['font_name'], min_font_size=7)

    # 3. الاسم العربي
    ar_y = -cm2p(POS_AR_Y_CM)
    ar_txt_proc = process_text(desc_ar, is_arabic=True)
    draw_text_multiline(c, ar_txt_proc, ar_y, max_text_width, 
                        FONT_NAME if has_font else "Helvetica", settings['font_name'], min_font_size=7)

    # 4. العرض
    offer_y = -(height / 2) - 5 
    if has_font:
        draw_text_multiline(c, str(offer_txt), offer_y, max_text_width, 
                            FONT_NAME, settings['font_offer'], min_font_size=12, 
                            color=(0.85, 0.21, 0.27), is_bold=True)
    else:
        c.setFont("Helvetica-Bold", settings['font_offer'])
        c.setFillColorRGB(0.85, 0.21, 0.27)
        c.drawCentredString(0, offer_y, str(offer_txt))

    # 5. الباركود
    barcode_y = -height + cm2p(POS_BARCODE_BOTTOM_CM)
    if item_code:
        try:
            barcode = code128.Code128(item_code, barHeight=20, barWidth=1.2)
            bc_x = -(barcode.width / 2)
            barcode.drawOn(c, bc_x, barcode_y + 10)
            c.setFillColorRGB(0, 0, 0)
            c.setFont("Helvetica", 8)
            c.drawCentredString(0, barcode_y, item_code)
        except:
            pass

def generate_pdf(df, settings, preview_mode=False):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    page_w_pt, page_h_pt = A4 
    
    row1_top = cm2p(DIM_ROW1_TOP_CM)
    row2_top = cm2p(DIM_ROW2_TOP_CM)
    
    abs_centers_from_left = [
        cm2p(21.0 - CENTERS_FROM_RIGHT_CM[2]), 
        cm2p(21.0 - CENTERS_FROM_RIGHT_CM[1]), 
        cm2p(21.0 - CENTERS_FROM_RIGHT_CM[0])  
    ]
    
    cols = 3
    cards_per_page = 6
    
    if preview_mode:
        df_to_process = df.head(cards_per_page)
    else:
        df_to_process = df
        
    for i, (_, row) in enumerate(df_to_process.iterrows()):
        if i > 0 and i % cards_per_page == 0:
            c.showPage()
            
        if preview_mode and (i % cards_per_page == 0) and has_template:
            c.drawImage(TEMPLATE_PATH, 0, 0, width=page_w_pt, height=page_h_pt)
        
        pos_in_page = i % cards_per_page
        col_idx = pos_in_page % cols
        row_idx = pos_in_page // cols
        
        x_center = abs_centers_from_left[col_idx]
        
        if row_idx == 0:
            y_start = page_h_pt - row1_top
        else:
            y_start = page_h_pt - row2_top

        offset_x = 0
        offset_y = 0
        if row_idx == 0:
            offset_x = cm2p(settings['top_x_cm'])
            offset_y = cm2p(settings['top_y_cm'])
        else:
            offset_x = cm2p(settings['bot_x_cm'])
            offset_y = cm2p(settings['bot_y_cm'])

        final_x = x_center + offset_x
        final_y = y_start + offset_y
            
        c.saveState()
        c.translate(final_x, final_y)
        draw_card_content(c, row, settings)
        c.restoreState()
        
    c.save()
    buffer.seek(0)
    return buffer

def create_preview_image(df, settings):
    pdf_buffer = generate_pdf(df, settings, preview_mode=True)
    doc = fitz.open(stream=pdf_buffer.getvalue(), filetype="pdf")
    page = doc.load_page(0)
    pix = page.get_pixmap(dpi=150)
    return pix.tobytes("png")

# ==========================================
# 3. الواجهة
# ==========================================
st.title("🖨️ Offers Generator")

if not has_font:
    st.error("⚠️ ملف الخط `arial.ttf` مفقود!")

st.sidebar.header("1. البيانات")
offers_file = st.sidebar.file_uploader("ملف العروض", type=['xlsx'])
stock_file = st.sidebar.file_uploader("ملف المخزون", type=['xlsx'])
min_qty = st.sidebar.number_input("أقل كمية", 1, 100, 2)

st.sidebar.markdown("---")
st.sidebar.header("2. 🎛️ ضبط الإزاحات (Calibration)")

tab_top, tab_bot = st.sidebar.tabs(["⬆️ الصف العلوي", "⬇️ الصف السفلي"])
with tab_top:
    s_top_x = st.number_input("يمين/يسار (Top X)", -5.0, 5.0, 0.0, step=0.1, key='tx')
    s_top_y = st.number_input("فوق/تحت (Top Y)", -5.0, 5.0, 0.0, step=0.1, key='ty')
with tab_bot:
    s_bot_x = st.number_input("يمين/يسار (Bot X)", -5.0, 5.0, 0.0, step=0.1, key='bx')
    s_bot_y = st.number_input("فوق/تحت (Bot Y)", -5.0, 5.0, 0.0, step=0.1, key='by')

st.sidebar.markdown("---")
with st.sidebar.expander("🅰️ أحجام الخطوط", expanded=True):
    s_f_brand = st.slider("أقصى خط للبراند", 8, 20, 12)
    s_f_name = st.slider("أقصى خط للأسماء", 6, 18, 10)
    s_f_offer = st.slider("أقصى خط للعرض", 10, 40, 24)

user_settings = {
    'top_x_cm': s_top_x, 'top_y_cm': s_top_y,
    'bot_x_cm': s_bot_x, 'bot_y_cm': s_bot_y,
    'font_brand': s_f_brand, 'font_name': s_f_name, 'font_offer': s_f_offer
}

if offers_file and stock_file:
    try:
        df1 = pd.read_excel(offers_file)
        df2 = pd.read_excel(stock_file)
        df1['Item Number'] = df1['Item Number'].astype(str).str.replace('.0', '')
        df2['Item Number'] = df2['Item Number'].astype(str).str.replace('.0', '')
        merged = pd.merge(df1, df2[['Item Number', 'Quantity']], on='Item Number', how='left')
        base_df = merged[merged['Quantity'] >= min_qty].copy()

        if base_df.empty:
            st.warning("لا توجد أصناف.")
        else:
            st.markdown("---")
            c1, c2, c3 = st.columns(3)
            cats = ['All'] + sorted(list(base_df['Category'].astype(str).unique()))
            sel_cat = c1.selectbox("القسم", cats)
            df_cat = base_df if sel_cat == 'All' else base_df[base_df['Category'].astype(str) == sel_cat]
            
            brands = ['All'] + sorted(list(df_cat['Brand'].astype(str).unique()))
            sel_brand = c2.selectbox("البراند", brands)
            df_brand = df_cat if sel_brand == 'All' else df_cat[df_cat['Brand'].astype(str) == sel_brand]
            
            offers = ['All'] + sorted(list(df_brand['Offer Description EN'].astype(str).unique()))
            sel_offer = c3.selectbox("العرض", offers)
            final_df = df_brand if sel_offer == 'All' else df_brand[df_brand['Offer Description EN'].astype(str) == sel_offer]

            st.info(f"العدد: {len(final_df)}")
            
            if not final_df.empty:
                if has_template:
                    if st.button("👁️ معاينة حية"):
                        img = create_preview_image(final_df, user_settings)
                        st.image(img, caption="معاينة دقيقة")
                
                pdf_data = generate_pdf(final_df, user_settings)
                st.download_button("📥 تحميل PDF", pdf_data, "Offers.pdf", "application/pdf", type="primary")

    except Exception as e:
        st.error(f"Error: {e}")
