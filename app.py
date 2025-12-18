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

# --- إعدادات الصفحة في المتصفح ---
st.set_page_config(page_title="مولد ملصقات العروض", page_icon="🏷️")

st.title("🏷️ برنامج طباعة ملصقات العروض")
st.write("قم برفع ملف الإكسيل وسيقوم البرنامج بتحويله إلى PDF جاهز للطباعة.")

# --- القائمة الجانبية للإعدادات ---
st.sidebar.header("⚙️ إعدادات الطباعة")
shift_top = st.sidebar.number_input("إزاحة الصف العلوي (لأعلى/لأسفل)", value=-10, step=1)
shift_bottom = st.sidebar.number_input("إزاحة الصف السفلي (لأعلى/لأسفل)", value=-10, step=1)

# --- الدوال المساعدة ---
# ملاحظة: نحتاج لملف خط بجانب الكود لأن السيرفر لا يحتوي على خطوط ويندوز
FONT_NAME = "CustomFont"
FONT_BOLD = "CustomFontBold"

def setup_fonts():
    # يجب وضع ملف arial.ttf في نفس مجلد البرنامج
    try:
        pdfmetrics.registerFont(TTFont(FONT_NAME, "arial.ttf"))
        pdfmetrics.registerFont(TTFont(FONT_BOLD, "arialbd.ttf")) 
    except:
        # خط احتياطي في حال عدم وجود الملفات
        st.warning("لم يتم العثور على ملفات الخطوط (arial.ttf)، سيتم استخدام الخط الافتراضي (قد لا تظهر العربية بشكل صحيح).")

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
    current_shift = shift_top if row_index == 0 else shift_bottom
    yellow_center_y = y + (height * 0.38) + current_shift

    
    # الاسم الإنجليزي
    item_en = str(data.get('English Name', ''))[:28]
    c.setFont(FONT_NAME, 11)
    c.drawCentredString(center_x, yellow_center_y + 45, item_en)

    # الاسم العربي
    item_ar = process_arabic(data.get('Arabic Name', ''))
    c.setFont(FONT_NAME, 11)
    c.drawCentredString(center_x, yellow_center_y + 25, item_ar)

    # العرض
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

    # الباركود
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
    buffer = io.BytesIO() # إنشاء ملف في الذاكرة بدلاً من القرص الصلب
    c = canvas.Canvas(buffer, pagesize=A4)
    setup_fonts()
    
    # إعدادات الصفحة
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

# --- واجهة التطبيق الرئيسية ---
uploaded_file = st.file_uploader("اختر ملف الإكسيل (Excel)", type=['xlsx'])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        st.success("تم قراءة الملف بنجاح! عدد الأصناف: " + str(len(df)))
        
        # عرض عينة من البيانات
        st.dataframe(df.head())

        if st.button("إنشاء ملف PDF"):
            pdf_bytes = create_pdf(df)
            st.success("تم إنشاء الملف بنجاح! اضغط بالأسفل للتحميل.")
            
            st.download_button(
                label="📥 تحميل ملف PDF",
                data=pdf_bytes,
                file_name="offers_labels.pdf",
                mime="application/pdf"
            )
            
    except Exception as e:
        st.error(f"حدث خطأ: {e}")
