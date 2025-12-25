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

# --- القائمة الجانبية: التحكم في الخطوط ---
st.sidebar.header("🔠 إعدادات حجم الخط")
name_font_size = st.sidebar.number_input("حجم خط اسم الصنف", value=11, min_value=5, max_value=25, step=1)
offer_font_size = st.sidebar.number_input("حجم خط العرض/السعر", value=30, min_value=10, max_value=60, step=1)

# --- إعدادات الإزاحة الثابتة ---
TOP_LOGO_SHIFT = 15       
TOP_CONTENT_SHIFT = -10   
BOTTOM_LOGO_SHIFT = 0     
BOTTOM_CONTENT_SHIFT = -20 

# --- تعريف الخطوط ---
FONT_NAME = "CustomFont"
FONT_BOLD = "CustomFontBold"

def setup_fonts():
    try:
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
    
    if row_index == 0:
        current_logo_shift = TOP_LOGO_SHIFT
        current_content_shift = TOP_CONTENT_SHIFT
    else:
        current_logo_shift = BOTTOM_LOGO_SHIFT
        current_content_shift = BOTTOM_CONTENT_SHIFT

    # 1. رسم الشعار
    brand_ar = process_arabic("الدواء")
    c.setFont(FONT_BOLD, 18)
    logo_y_pos = y + (height * 0.83) + current_logo_shift
    c.drawCentredString(center_x, logo_y_pos, f"al-dawaa | {brand_ar}")

    # نقطة ارتكاز المحتوى
    yellow_center_y = y + (height * 0.38) + current_content_shift

    # --- استخراج البيانات بناءً على ترتيب الأعمدة (Index) وليس الاسم ---
    # iloc[0] = العمود الأول (الكود)
    # iloc[1] = العمود الثاني (الاسم العربي)
    # iloc[2] = العمود الثالث (الاسم الإنجليزي)
    # iloc[3] = العمود الرابع (العرض)

    # 2. الاسم الإنجليزي (العمود الثالث - رقم 2)
    item_en = str(data.iloc[2])[:28] if len(data) > 2 else ""
    if item_en == 'nan': item_en = ""
    c.setFont(FONT_NAME, name_font_size)
    c.drawCentredString(center_x, yellow_center_y + 45, item_en)

    # 3. الاسم العربي (العمود الثاني - رقم 1)
    item_ar_raw = data.iloc[1] if len(data) > 1 else ""
    item_ar = process_arabic(item_ar_raw)
    c.setFont(FONT_NAME, name_font_size)
    c.drawCentredString(center_x, yellow_center_y + 25, item_ar)

    # 4. العرض (العمود الرابع - رقم 3)
    offer_raw = data.iloc[3] if len(data) > 3 else ""
    clean_val, is_number = clean_offer_value(offer_raw)
    
    if is_number:
        offer_en = f"{clean_val}% off"
        offer_ar = process_arabic(f"خصم {clean_val}%")
    else:
        offer_en = clean_val
        offer_ar = process_arabic(clean_val)

    # رسم العرض بالإنجليزية
    c.setFont(FONT_BOLD, offer_font_size)
    c.drawCentredString(center_x, yellow_center_y - 20, offer_en)
    
    # رسم العرض بالعربية
    if is_number:
        arabic_offer_size = int(offer_font_size * 0.6)
        c.setFont(FONT_BOLD, arabic_offer_size)
        c.drawCentredString(center_x, yellow_center_y - 45, offer_ar)

    # 5. الباركود (العمود الأول - رقم 0)
    raw_code = str(data.iloc[0]).replace('.0', '') if len(data) > 0 else ""
    if raw_code == 'nan': raw_code = ""

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
st.write("### 📂 تعليمات الملف")
st.info("""
**ملاحظة هامة:** لا يهم اسم الأعمدة في الملف، ولكن **يجب** أن يكون ترتيب البيانات كالتالي:
1. العمود الأول: **كود الصنف** (Code)
2. العمود الثاني: **الاسم العربي**
3. العمود الثالث: **الاسم الإنجليزي**
4. العمود الرابع: **قيمة العرض** (السعر/الخصم)
""")

uploaded_file = st.file_uploader("اختر ملف الإكسيل (Excel)", type=['xlsx'])

if uploaded_file is not None:
    try:
        # قراءة الملف (header=0 يعني يعتبر الصف الأول عناوين ولكنا سنتجاهل أسماءها)
        df = pd.read_excel(uploaded_file)
        
        # التأكد من أن الملف يحتوي على 4 أعمدة على الأقل
        if len(df.columns) < 4:
            st.error("❌ خطأ: الملف المرفوع يحتوي على أقل من 4 أعمدة. يرجى التأكد من الملف.")
        else:
            st.success(f"✅ تم تحميل الملف: {len(df)} صنف")
            
            # عرض معاينة للمستخدم ليتأكد من الترتيب
            st.write("👀 **معاينة البيانات (تأكد أن الترتيب صحيح):**")
            st.dataframe(df.head())

            if st.button("تحويل إلى PDF"):
                pdf_bytes = create_pdf(df)
                st.download_button("📥 تحميل الملف", pdf_bytes, "offers_print.pdf", "application/pdf")
                
    except Exception as e:
        st.error(f"خطأ: {e}")
