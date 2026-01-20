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
# 1. إعدادات الخطوط والنظام
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
    """معالجة النص العربي"""
    if pd.isna(text) or text == "": return ""
    text = str(text)
    if is_arabic and has_font:
        reshaped = arabic_reshaper.reshape(text)
        return get_display(reshaped)
    return text

# ==========================================
# 2. محرك رسم PDF (المحدث)
# ==========================================

def draw_wrapped_text(c, text, x, y, max_width, font_name, font_size, line_spacing=4):
    """
    دالة مساعدة لكتابة نص طويل وتنزيله للسطر التالي إذا لم يكفِ العرض.
    ترجع دالة الـ y الجديدة بعد الانتهاء من الكتابة (لمعرفة مكان العنصر التالي).
    """
    c.setFont(font_name, font_size)
    # تقسيم النص إلى أسطر بناءً على العرض المتاح
    lines = simpleSplit(text, font_name, font_size, max_width)
    
    current_y = y
    for line in lines:
        c.drawCentredString(x, current_y, line)
        current_y -= (font_size + line_spacing) # النزول للأسفل
    
    # إرجاع الارتفاع الذي توقفنا عنده لنكمل الرسم تحته
    # نعيد حساب الفرق لنعرف كم نزلنا بالضبط
    height_consumed = y - current_y
    return current_y, height_consumed

def draw_label(c, x, y, w, h, row, settings):
    # استخراج البيانات
    item_code = str(row.get('Item Number', '')).replace('.0', '')
    desc_en = row.get('Item Description EN', '') # الاسم كاملاً بدون قص
    desc_ar = row.get('Item Description AR', '')
    brand_txt = row.get('Brand', '')
    offer_txt = row.get('Offer Description EN', '')

    center_x = x + (w / 2)
    max_text_width = w * 0.90 # ترك هامش 10%

    # رسم إطار للتجربة (اختياري)
    if settings['show_borders']:
        c.setLineWidth(0.5)
        c.setStrokeColorRGB(0.8, 0.8, 0.8)
        c.rect(x, y, w, h)

    # --- بداية الحسابات ---
    # النقطة المرجعية العليا (بداية منطقة الطباعة الصفراء 8 سم)
    current_y = (y + h) - settings['top_offset_skip']

    # 1. اسم الصيدلية (Pharmacy Name) - تمت إضافته
    c.setFillColorRGB(0.4, 0.4, 0.4) # لون رمادي غامق
    c.setFont(FONT_NAME if has_font else "Helvetica", settings['header_font_size'])
    pharmacy_name = process_text("Al-Dawaa Pharmacy | صيدلية الدواء", is_arabic=True)
    c.drawCentredString(center_x, current_y, pharmacy_name)
    
    # تحديث المكان للمكون التالي
    current_y -= settings['spacing_header_to_brand']

    # 2. البراند (Brand)
    c.setFillColorRGB(0, 0, 0) # أسود
    c.setFont(FONT_NAME if has_font else "Helvetica-Bold", settings['brand_font_size'])
    c.drawCentredString(center_x, current_y, str(brand_txt))
    
    # تحديث المكان
    current_y -= settings['spacing_brand_to_name']

    # 3. الاسم الإنجليزي (English Name) - مع التفاف النص
    font_used = FONT_NAME if has_font else "Helvetica"
    new_y, _ = draw_wrapped_text(c, str(desc_en), center_x, current_y, max_text_width, font_used, settings['name_font_size'])
    current_y = new_y # اعتماد المكان الجديد بعد كتابة الأسطر

    # 4. الفراغ الأول (1 سم تقريباً) بين الإنجليزي والعربي
    current_y -= settings['spacing_en_to_ar']

    # 5. الاسم العربي (Arabic Name) - مع التفاف النص
    ar_text = process_text(desc_ar, is_arabic=True)
    new_y, _ = draw_wrapped_text(c, ar_text, center_x, current_y, max_text_width, font_used, settings['name_font_size'])
    current_y = new_y

    # 6. الفراغ الثاني (3 سم تقريباً) بين الأسماء والعرض
    current_y -= settings['spacing_ar_to_offer']

    # 7. العرض / السعر (Offer/Price)
    c.setFont(FONT_NAME if has_font else "Helvetica-Bold", settings['price_font_size'])
    c.setFillColorRGB(0.85, 0.21, 0.27) # أحمر
    c.drawCentredString(center_x, current_y, str(offer_txt))

    # 8. الباركود والرقم (في الأسفل تماماً)
    # يتم حسابه من الأسفل للأعلى لضمان عدم خروجه من الورقة
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
st.title("🏷️ Offers Generator Pro (Full Control)")

if not has_font:
    st.warning("⚠️ Font `arial.ttf` missing. Arabic will look broken.")

# --- القائمة الجانبية ---
st.sidebar.header("1. البيانات")
offers_file = st.sidebar.file_uploader("ملف العروض", type=['xlsx'])
stock_file = st.sidebar.file_uploader("ملف المخزون", type=['xlsx'])
min_qty = st.sidebar.number_input("أقل كمية", 2, 100, 2)

st.sidebar.markdown("---")
st.sidebar.header("2. تصميم المسافات والأحجام")
show_borders = st.sidebar.checkbox("إظهار حدود للتجربة", False)

with st.sidebar.expander("📏 المسافات (المحور الرأسي)", expanded=True):
    st.info("ملاحظة: 28 نقطة ≈ 1 سم")
    
    s_top_offset = st.slider("إزاحة علوية (لتخطي الهيدر الأحمر)", 0, 100, 40)
    s_head_brand_gap = st.slider("مسافة: صيدلية -> براند", 5, 50, 15)
    s_brand_name_gap = st.slider("مسافة: براند -> اسم إنجليزي", 5, 50, 15)
    
    # الطلب: 1 سم بين الاسمين
    s_en_ar_gap = st.slider("مسافة: إنجليزي -> عربي (تقريباً 1 سم = 28)", 5, 60, 28)
    
    # الطلب: 3 سم بين الاسماء والعرض
    s_ar_offer_gap = st.slider("مسافة: عربي -> العرض (تقريباً 3 سم = 85)", 10, 120, 85)
    
    s_bc_bottom = st.slider("مكان الباركود (من الأسفل)", 0, 80, 20)

with st.sidebar.expander("🅰️ أحجام الخطوط", expanded=False):
    s_header_font = st.slider("حجم اسم الصيدلية", 6, 14, 8)
    s_brand_font = st.slider("حجم البراند", 10, 24, 14)
    s_name_font = st.slider("حجم اسم الصنف", 8, 20, 11)
    s_price_font = st.slider("حجم السعر/العرض", 10, 60, 24)
    s_bc_h = st.slider("ارتفاع الباركود", 10, 50, 25)
    s_bc_font = st.slider("حجم رقم الباركود", 6, 14, 10)

user_settings = {
    'show_borders': show_borders, 
    'top_offset_skip': s_top_offset,
    'barcode_bottom_margin': s_bc_bottom, 
    
    'spacing_header_to_brand': s_head_brand_gap,
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

# --- المنطق الرئيسي ---
if offers_file and stock_file:
    try:
        df1 = pd.read_excel(offers_file)
        df2 = pd.read_excel(stock_file)
        
        df1['Item Number'] = df1['Item Number'].astype(str).str.replace('.0', '')
        df2['Item Number'] = df2['Item Number'].astype(str).str.replace('.0', '')
        merged = pd.merge(df1, df2[['Item Number', 'Quantity']], on='Item Number', how='left')
        final_df = merged[merged['Quantity'] >= min_qty].copy()

        if final_df.empty:
            st.error("لا توجد أصناف مطابقة.")
        else:
            c1, c2, c3 = st.columns(3)
            cats = ['All'] + sorted(list(final_df['Category'].dropna().unique()))
            brands = ['All'] + sorted(list(final_df['Brand'].dropna().unique()))
            sel_cat = c1.selectbox("القسم", cats)
            sel_brand = c2.selectbox("البراند", brands)
            
            if sel_cat != 'All': final_df = final_df[final_df['Category'] == sel_cat]
            if sel_brand != 'All': final_df = final_df[final_df['Brand'] == sel_brand]
            
            st.info(f"جاهز لطباعة **{len(final_df)}** صنف.")
            
            # زر المعاينة
            if st.button("👁️ معاينة الصفحة الأولى", type="primary"):
                preview_pdf = generate_pdf(final_df.head(6), user_settings)
                st.session_state['preview_pdf'] = preview_pdf
            
            # عرض المعاينة
            if 'preview_pdf' in st.session_state:
                st.markdown("---")
                col_prev, col_down = st.columns([2, 1])
                
                with col_prev:
                    st.subheader("صورة المعاينة")
                    doc = fitz.open(stream=st.session_state['preview_pdf'].getvalue(), filetype="pdf")
                    pix = doc.load_page(0).get_pixmap(dpi=150)
                    st.image(pix.tobytes("png"), caption="معاينة حية", width=600)
                
                with col_down:
                    st.success("هل التصميم مناسب؟")
                    full_pdf = generate_pdf(final_df, user_settings)
                    st.download_button("📥 تحميل الملف كامل", full_pdf, "Final_Offers.pdf", "application/pdf")

    except Exception as e:
        st.error(f"حدث خطأ: {e}")
