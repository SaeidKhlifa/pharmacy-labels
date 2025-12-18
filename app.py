import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.graphics.barcode import code128
import arabic_reshaper
from bidi.algorithm import get_display
import io
import os

# --- إعدادات الصفحة ---
st.set_page_config(page_title="مولد ملصقات العروض", page_icon="🏷️")
st.title("🏷️ برنامج طباعة ملصقات العروض")

# --- القائمة الجانبية (تم التحديث: 4 مفاتيح تحكم) ---
st.sidebar.header("⚙️ إعدادات الصف العلوي (Top Row)")
top_logo_shift = st.sidebar.number_input("1. إزاحة الشعار (العلوي)", value=0, step=1, help="يحرك كلمة الدواء فقط")
top_content_shift = st.sidebar.number_input("2. إزاحة المحتوى (العلوي)", value=-10, step=1, help="يحرك الاسم والسعر والباركود")

st.sidebar.markdown("---") # فاصل خطي

st.sidebar.header("⚙️ إعدادات الصف السفلي (Bottom Row)")
bottom_logo_shift = st.sidebar.number_input("3. إزاحة الشعار (السفلي)", value=0, step=1)
bottom_content_shift = st.sidebar.number_input("4. إزاحة المحتوى (السفلي)", value=-10, step=1)

# --- تعريف الخطوط ---
FONT_NAME = "CustomFont"
FONT_BOLD = "CustomFontBold"

def setup_fonts():
    try:
        # محاولة استخدام الخطوط المرفقة
        if os.path.exists("arial.ttf"):
            pdfmetrics.registerFont(TTFont(FONT_NAME, "arial.ttf"))
        else:
            st.warning("⚠️ ملف arial.ttf غير موجود.")
            
        if os.path.exists("arialbd.ttf"):
            pdfmetrics.registerFont(TTFont(FONT_BOLD, "arialbd.ttf")) 
        else:
             if os.path.exists("arial.ttf"):
                pdfmetrics.registerFont(TTFont(FONT_BOLD, "arial.ttf"))
    except Exception as e:
        st.error(f"خطأ في الخطوط: {e}")

def process_arabic(text):
    if not text or pd.isna(text): return ""
    text = str(text)
    reshaped = arabic_reshaper.reshape(text)
    return get_display(reshaped)

def clean_offer_value(raw_value):
    str_val = str(raw_value).strip()
    try:
        float_val = float(str_val)
        if 0 < float_val < 1:
            percentage = float_val * 100
            if percentage.is_integer(): return str(int(percentage)), True
            return str(round(percentage, 1)), True
        if float_val.is_integer(): return str(int(float_val)), True
        return str(float_val), True
    except ValueError:
        return str_val, False

def draw_block(c, x, y, width, height, data, row_index):
    center_x = x + (width / 2)
    
    # تحديد قيم الإزاحة بناءً على رقم الصف
    if row_index == 0:
        current_logo_shift = top_logo_shift
        current_content_shift = top_content_shift
    else:
        current_logo_shift = bottom_logo_shift
        current_content_shift = bottom_content_shift

    # 1. رسم الشعار (يتأثر بإزاحة الشعار فقط)
    brand_ar = process_arabic("الدواء")
    c.setFont(FONT_BOLD, 18)
    # مكان الشعار الأساسي + إزاحة الشعار
    logo_y_pos = y + (height * 0.83) + current_logo_shift
    c.drawCentredString(center_x, logo_y_pos, f"al-dawaa | {brand_ar}")

    # --- حساب نقطة ارتكاز المحتوى (تتأثر بإزاحة المحتوى فقط) ---
    yellow_center_y = y + (height * 0.38) + current_content_shift

    # 2. الاسم الإنجليزي
    item_en = str(data.get('English Name', ''))[:28]
    c.setFont(FONT_NAME, 11)
    c.drawCentredString(center_x, yellow_center_y + 45, item_en)

    # 3. الاسم العربي
    item_ar = process_arabic(data.get('Arabic Name', ''))
    c.setFont(FONT_NAME, 11)
    c.drawCentredString(center_x, yellow_center_y + 25, item_ar)

    # 4. العرض
    clean_val, is_number = clean_offer_value(data.get('Current Offer', ''))
    if is_number:
        offer_en = f"{clean_val}% off"
        offer_ar = process_arabic(f"خصم {clean_val}%")
    else:
        offer_en = clean_val
        offer_ar = process_arabic(clean_val)

    c.setFont(FONT_BOLD, 30)
    c.drawCentredString(center_x, yellow_center_y - 20, offer_en)
    
    if is_number:
        c.setFont(FONT_BOLD, 18)
        c.drawCentredString(center_x, yellow_center_y - 45, offer_ar)

    # 5. الباركود
    raw_code = str(data.get('Item Code', '')).replace('.0', '')
    barcode_y = yellow_center_y - 85
    if raw_code:
        try:
            barcode = code128.Code128(raw_code, barHeight=25, barWidth=1.2)
            barcode.drawOn(c, center_x - (barcode.width/2), barcode_y)
            c.setFont(FONT_NAME, 10)
            c.drawCentredString(center_x, barcode_y - 12, raw_code)
        except:
            c.drawCentredString(center_x, barcode_y, raw_code)

def create_pdf(df):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    setup_fonts()
    
    PAGE_WIDTH, PAGE_HEIGHT = A4
    MARGIN_X, MARGIN_Y = 20, 20
    COLS, ROWS = 3, 2
    BLOCK_WIDTH = (PAGE_WIDTH - (2 * MARGIN_X)) / COLS
    BLOCK_HEIGHT = (PAGE_HEIGHT - (2 * MARGIN_Y)) / ROWS

    col_counter = 0
    row_counter = 0
    
    for _, row in df.iterrows():
        x_pos = MARGIN_X + (col_counter * BLOCK_WIDTH)
        y_pos = PAGE_HEIGHT - MARGIN_Y - ((row_counter + 1) * BLOCK_HEIGHT)
        
        draw_block(c, x_pos, y_pos, BLOCK_WIDTH, BLOCK_HEIGHT, row, row_counter)
        
        col_counter += 1
        if col_counter >= COLS:
            col_counter = 0
            row_counter += 1
        if row_counter >= ROWS:
            c.showPage()
            col_counter, row_counter = 0, 0
            
    c.save()
    buffer.seek(0)
    return buffer

# --- الواجهة ---
st.write("قم برفع ملف الإكسيل وسيقوم البرنامج بتحويله إلى PDF.")
uploaded_file = st.file_uploader("اختر ملف الإكسيل (Excel)", type=['xlsx'])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        st.success(f"تم تحميل الملف: {len(df)} صنف")
        if st.button("تحويل إلى PDF"):
            pdf_bytes = create_pdf(df)
            st.download_button("📥 تحميل الملف", pdf_bytes, "offers_v2.pdf", "application/pdf")
    except Exception as e:
        st.error(f"خطأ: {e}")
