import streamlit as st
import pandas as pd
import io
import os
import fitz  # PyMuPDF
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.utils import simpleSplit
from reportlab.graphics.barcode import code128
import arabic_reshaper
from bidi.algorithm import get_display

# ==========================================
# 1. إعدادات الخطوط
# ==========================================
st.set_page_config(page_title="Offers Generator Pro", layout="wide", page_icon="🏷️")

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
        reshaped = arabic_reshaper.reshape(text)
        return get_display(reshaped)
    return text

# ==========================================
# 2. دوال الرسم (محسنة)
# ==========================================

def draw_bold_centered(c, text, center_x, y, font_name, font_size, color_rgb=(0,0,0), stroke_width=0.7):
    """رسم نص عريض (Bold) في سطر واحد"""
    text_width = pdfmetrics.stringWidth(text, font_name, font_size)
    start_x = center_x - (text_width / 2)
    c.setStrokeColorRGB(*color_rgb)
    c.setFillColorRGB(*color_rgb)
    c.setLineWidth(stroke_width)
    text_obj = c.beginText()
    text_obj.setTextRenderMode(2)
    text_obj.setFont(font_name, font_size)
    text_obj.setTextOrigin(start_x, y)
    text_obj.textOut(text)
    c.drawText(text_obj)
    # إعادة الضبط
    c.setFillColorRGB(0, 0, 0)
    c.setStrokeColorRGB(0, 0, 0)
    c.setLineWidth(0)

def draw_wrapped_text(c, text, x, y, max_width, font_name, font_size, line_spacing=4):
    """رسم نص عادي متعدد الأسطر"""
    c.setFont(font_name, font_size)
    lines = simpleSplit(text, font_name, font_size, max_width)
    current_y = y
    for line in lines:
        c.drawCentredString(x, current_y, line)
        current_y -= (font_size + line_spacing)
    return current_y

def draw_bold_wrapped_text(c, text, center_x, y, max_width, font_name, font_size, color_rgb, stroke_width=1.0, line_spacing=4):
    """
    دالة جديدة: رسم نص عريض (Bold) ومتعدد الأسطر (Wrapped)
    تستخدم للعروض الطويلة
    """
    # تقسيم النص إلى أسطر
    lines = simpleSplit(text, font_name, font_size, max_width)
    
    current_y = y
    # رسم كل سطر باستخدام دالة الـ Bold
    for line in lines:
        draw_bold_centered(c, line, center_x, current_y, font_name, font_size, color_rgb, stroke_width)
        current_y -= (font_size + line_spacing)
        
    return current_y

def draw_label(c, x, y, w, h, row, settings):
    # البيانات
    item_code = str(row.get('Item Number', '')).replace('.0', '')
    desc_en = row.get('Item Description EN', '') 
    desc_ar = row.get('Item Description AR', '')
    brand_txt = row.get('Brand', '')
    offer_txt = row.get('Offer Description EN', '')

    center_x = x + (w / 2)
    max_text_width = w * 0.90 

    if settings['show_borders']:
        c.setLineWidth(0.5)
        c.setStrokeColorRGB(0.8, 0.8, 0.8)
        c.rect(x, y, w, h)

    # 1. المنطقة العلوية (البيضاء)
    header_y_position = (y + h) - 30 
    c.setFillColorRGB(0.4, 0.4, 0.4) 
    c.setFont(FONT_NAME if has_font else "Helvetica", settings['header_font_size'])
    pharmacy_name = process_text("Al-Dawaa Pharmacy | صيدلية الدواء", is_arabic=True)
    c.drawCentredString(center_x, header_y_position, pharmacy_name)

    # 2. منطقة المحتوى (الصفراء)
    content_start_y = (y + h) - settings['top_offset_skip']

    # أ. البراند
    current_y = content_start_y
    c.setFillColorRGB(0, 0, 0)
    if has_font:
        draw_bold_centered(c, str(brand_txt), center_x, current_y, FONT_NAME, settings['brand_font_size'], (0,0,0), stroke_width=0.7)
    else:
        c.setFont("Helvetica-Bold", settings['brand_font_size'])
        c.drawCentredString(center_x, current_y, str(brand_txt))

    current_y -= settings['spacing_brand_to_name']

    # ب. الاسم الإنجليزي
    font_used = FONT_NAME if has_font else "Helvetica"
    c.setFillColorRGB(0, 0, 0)
    current_y = draw_wrapped_text(c, str(desc_en), center_x, current_y, max_text_width, font_used, settings['name_font_size'])

    # ج. مسافة
    current_y -= settings['spacing_en_to_ar']

    # د. الاسم العربي
    ar_text = process_text(desc_ar, is_arabic=True)
    current_y = draw_wrapped_text(c, ar_text, center_x, current_y, max_text_width, font_used, settings['name_font_size'])

    # هـ. مسافة قبل العرض
    current_y -= settings['spacing_ar_to_offer']

    # و. العرض / السعر (تم التعديل ليقبل سطرين)
    if has_font:
        # استخدام الدالة الجديدة التي تدعم التفاف النص مع الـ Bold
        draw_bold_wrapped_text(c, str(offer_txt), center_x, current_y, max_text_width, FONT_NAME, settings['price_font_size'], (0.85, 0.21, 0.27), stroke_width=1.0)
    else:
        # Fallback للخطوط العادية
        c.setFont("Helvetica-Bold", settings['price_font_size'])
        c.setFillColorRGB(0.85, 0.21, 0.27)
        # تقسيم النص يدوياً للخطوط العادية
        lines = simpleSplit(str(offer_txt), "Helvetica-Bold", settings['price_font_size'], max_text_width)
        temp_y = current_y
        for line in lines:
            c.drawCentredString(center_x, temp_y, line)
            temp_y -= (settings['price_font_size'] + 4)

    # ز. الباركود
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
# 3. واجهة المستخدم
# ==========================================
st.title("🏷️ Offers Generator Pro (Multi-line Offers)")

if not has_font:
    st.warning("⚠️ Font `arial.ttf` missing.")

st.sidebar.header("1. البيانات")
offers_file = st.sidebar.file_uploader("ملف العروض", type=['xlsx'])
stock_file = st.sidebar.file_uploader("ملف المخزون", type=['xlsx'])
min_qty = st.sidebar.number_input("أقل كمية", 2, 100, 2)

st.sidebar.markdown("---")
st.sidebar.header("2. ضبط المسافات")
show_borders = st.sidebar.checkbox("حدود للتجربة", False)

with st.sidebar.expander("📏 المسافات الرأسية", expanded=True):
    s_top_offset = st.slider("بداية الكلام (تخطي الأبيض والأحمر)", 0, 250, 190)
    s_brand_name_gap = st.slider("مسافة: براند -> إنجليزي", 5, 50, 20)
    s_en_ar_gap = st.slider("مسافة: إنجليزي -> عربي", 5, 60, 15)
    s_ar_offer_gap = st.slider("مسافة: عربي -> العرض", 10, 120, 60)
    s_bc_bottom = st.slider("ارتفاع الباركود (من الأسفل)", 0, 80, 25)

with st.sidebar.expander("🅰️ أحجام الخطوط", expanded=False):
    s_header_font = st.slider("حجم اسم الصيدلية", 6, 14, 8)
    s_brand_font = st.slider("حجم البراند", 10, 24, 14)
    s_name_font = st.slider("حجم اسم الصنف", 8, 20, 10)
    s_price_font = st.slider("حجم العرض", 10, 60, 24)
    s_bc_h = st.slider("طول خطوط الباركود", 10, 50, 25)
    s_bc_font = st.slider("رقم الباركود", 6, 14, 10)

user_settings = {
    'show_borders': show_borders, 
    'top_offset_skip': s_top_offset,
    'barcode_bottom_margin': s_bc_bottom, 
    'spacing_brand_to_name': s_brand_name_gap,
    'spacing_en_to_ar': s_en_ar_gap, 
    'spacing_ar_to_offer': s_ar_offer_gap,
    'header_font_size': s_header_font,
    'brand_font_size': s_brand_font, 
    'name_font_size': s_name_font,
    'price_font_size': s_price_font, 
    'barcode_height': s_bc_h, 
    'barcode_font_size': s_bc_font
}

# المنطق الرئيسي
if offers_file and stock_file:
    try:
        df1 = pd.read_excel(offers_file)
        df2 = pd.read_excel(stock_file)
        df1['Item Number'] = df1['Item Number'].astype(str).str.replace('.0', '')
        df2['Item Number'] = df2['Item Number'].astype(str).str.replace('.0', '')
        merged = pd.merge(df1, df2[['Item Number', 'Quantity']], on='Item Number', how='left')
        final_df = merged[merged['Quantity'] >= min_qty].copy()

        if final_df.
