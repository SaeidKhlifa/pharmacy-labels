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
# 1. إعدادات النظام والخطوط
# ==========================================
st.set_page_config(page_title="Offers Generator Pro (Fixed Layout)", layout="wide", page_icon="🏷️")

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
# 2. دوال الرسم الذكية (Smart Drawing)
# ==========================================

def draw_text_auto_shrink(c, text, center_x, y, max_width, font_name, max_font_size, min_font_size=6, color=(0,0,0), is_bold=False):
    """
    دالة ذكية تقوم بتصغير الخط تلقائياً إذا كان النص أعرض من المساحة المتاحة (max_width).
    """
    current_size = max_font_size
    text_width = pdfmetrics.stringWidth(text, font_name, current_size)
    
    # حلقة تكرارية لتصغير الخط حتى يناسب العرض
    while text_width > max_width and current_size > min_font_size:
        current_size -= 0.5
        text_width = pdfmetrics.stringWidth(text, font_name, current_size)
    
    # إعداد الألوان
    c.setFillColorRGB(*color)
    c.setStrokeColorRGB(*color)
    
    # إذا كان مطلوب Bold نقوم بتطبيق Stroke
    if is_bold:
        c.setLineWidth(0.5 if current_size < 12 else 0.8)
        text_obj = c.beginText()
        text_obj.setTextRenderMode(2) # Fill + Stroke
        text_obj.setFont(font_name, current_size)
        text_obj.setTextOrigin(center_x - (text_width / 2), y)
        text_obj.textOut(text)
        c.drawText(text_obj)
        # إعادة تعيين
        c.setLineWidth(0)
        c.setTextRenderMode(0)
    else:
        # رسم عادي
        c.setFont(font_name, current_size)
        c.drawCentredString(center_x, y, text)
    
    # إعادة اللون للأسود افتراضياً
    c.setFillColorRGB(0, 0, 0)
    c.setStrokeColorRGB(0, 0, 0)

def draw_label(c, x, y, w, h, row, settings):
    # البيانات
    item_code = str(row.get('Item Number', '')).replace('.0', '')
    desc_en = row.get('Item Description EN', '') 
    desc_ar = row.get('Item Description AR', '')
    brand_txt = row.get('Brand', '')
    offer_txt = row.get('Offer Description EN', '')

    center_x = x + (w / 2)
    max_text_width = w * 0.92  # هامش أمان 8%

    # رسم حدود للتجربة
    if settings['show_borders']:
        c.setLineWidth(0.5)
        c.setStrokeColorRGB(0.8, 0.8, 0.8)
        c.rect(x, y, w, h) # حدود الكارت كامل
        # رسم خط يوضح بداية المنطقة الصفراء
        yellow_start_y = (y + h) - settings['top_offset_skip']
        c.setStrokeColorRGB(1, 0, 0) # خط أحمر
        c.line(x, yellow_start_y, x+w, yellow_start_y)

    # ==========================================
    # 1. المنطقة العلوية (البيضاء) - موقع ثابت مطلق
    # ==========================================
    # دائماً تبعد 30 نقطة عن الحافة العلوية للكارت
    header_y = (y + h) - 30 
    c.setFillColorRGB(0.4, 0.4, 0.4) 
    c.setFont(FONT_NAME if has_font else "Helvetica", settings['header_font_size'])
    pharmacy_name = process_text("Al-Dawaa Pharmacy | صيدلية الدواء", is_arabic=True)
    c.drawCentredString(center_x, header_y, pharmacy_name)

    # ==========================================
    # 2. المنطقة الصفراء (نظام المواقع الثابتة)
    # ==========================================
    # نقطة الصفر للمنطقة الصفراء (الخط الفاصل بين الأحمر والأصفر)
    yellow_zero_y = (y + h) - settings['top_offset_skip']

    # أ. البراند (Brand)
    # الموقع: ننزل من خط الصفر بمقدار brand_pos_y
    brand_y = yellow_zero_y - settings['brand_pos_y']
    if has_font:
        draw_text_auto_shrink(c, str(brand_txt), center_x, brand_y, max_text_width, 
                              FONT_NAME, settings['brand_font_size'], min_font_size=8, is_bold=True)
    else:
        c.setFont("Helvetica-Bold", settings['brand_font_size'])
        c.drawCentredString(center_x, brand_y, str(brand_txt))

    # ب. الاسم الإنجليزي (English Name)
    # الموقع: ننزل من خط الصفر بمقدار en_pos_y
    en_y = yellow_zero_y - settings['en_pos_y']
    draw_text_auto_shrink(c, str(desc_en), center_x, en_y, max_text_width, 
                          FONT_NAME if has_font else "Helvetica", settings['name_font_size'], min_font_size=8)

    # ج. الاسم العربي (Arabic Name)
    # الموقع: ننزل من خط الصفر بمقدار ar_pos_y
    ar_y = yellow_zero_y - settings['ar_pos_y']
    ar_txt_proc = process_text(desc_ar, is_arabic=True)
    draw_text_auto_shrink(c, ar_txt_proc, center_x, ar_y, max_text_width, 
                          FONT_NAME if has_font else "Helvetica", settings['name_font_size'], min_font_size=8)

    # د. العرض (Offer) - مساحة كبيرة
    # الموقع: ننزل من خط الصفر بمقدار offer_pos_y
    offer_y = yellow_zero_y - settings['offer_pos_y']
    if has_font:
        # هنا نسمح بتصغير أقل قليلاً لأننا نريد العرض كبيراً، ولكن إذا كان طويلاً جداً سيصغر
        draw_text_auto_shrink(c, str(offer_txt), center_x, offer_y, max_text_width, 
                              FONT_NAME, settings['price_font_size'], min_font_size=12, 
                              color=(0.85, 0.21, 0.27), is_bold=True)
    else:
        c.setFont("Helvetica-Bold", settings['price_font_size'])
        c.setFillColorRGB(0.85, 0.21, 0.27)
        c.drawCentredString(center_x, offer_y, str(offer_txt))

    # هـ. الباركود (Barcode) - ثابت من الأسفل
    # موقعه يعتمد على قاع الورقة (y) وليس المنطقة الصفراء العلوية
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
# 3. واجهة المستخدم (التحكم بالمواقع)
# ==========================================
st.title("🏷️ Offers Generator Pro (Fixed Layout & Auto-Size)")

if not has_font:
    st.warning("⚠️ Font `arial.ttf` missing.")

st.sidebar.header("1. البيانات")
offers_file = st.sidebar.file_uploader("ملف العروض", type=['xlsx'])
stock_file = st.sidebar.file_uploader("ملف المخزون", type=['xlsx'])
min_qty = st.sidebar.number_input("أقل كمية", 2, 100, 2)

st.sidebar.markdown("---")
st.sidebar.header("2. ضبط المواقع الثابتة")
st.sidebar.info("حرّك المؤشرات لتثبيت مكان كل عنصر داخل المنطقة الصفراء.")
show_borders = st.sidebar.checkbox("إظهار حدود للتجربة", False)

with st.sidebar.expander("📍 التحكم في المواقع (Y Position)", expanded=True):
    # 1. بداية المنطقة الصفراء
    s_top_offset = st.slider("بداية المنطقة الصفراء (تخطي الأحمر)", 100, 250, 190)
    
    # 2. مواقع العناصر داخل الأصفر (المسافة من بداية الأصفر)
    st.markdown("---")
    st.caption("المسافة من بداية المنطقة الصفراء (لأسفل):")
    
    s_brand_pos = st.slider("موقع البراند (Brand)", 10, 100, 20)
    s_en_pos = st.slider("موقع الاسم الإنجليزي", 20, 150, 50)
    s_ar_pos = st.slider("موقع الاسم العربي", 30, 200, 80)
    
    st.markdown("---")
    s_offer_pos = st.slider("موقع العرض (Offer) - الوسط", 50, 250, 140)
    
    st.markdown("---")
    st.caption("المسافة من أسفل الورقة (لأعلى):")
    s_bc_bottom = st.slider("موقع الباركود (ثابت في القاع)", 0, 80, 25)

with st.sidebar.expander("🅰️ أحجام الخطوط (الحد الأقصى)", expanded=False):
    st.caption("سيتم تصغير الخط تلقائياً إذا كان الكلام كثيراً")
    s_header_font = st.slider("حجم اسم الصيدلية", 6, 14, 8)
    s_brand_font = st.slider("حجم البراند (Max)", 10, 30, 14)
    s_name_font = st.slider("حجم الأسماء (Max)", 8, 25, 12)
    s_price_font = st.slider("حجم العرض (Max)", 10, 60, 30) # كبير جداً للعرض
    s_bc_h = st.slider("ارتفاع الباركود", 10, 50, 25)
    s_bc_font = st.slider("رقم الباركود", 6, 14, 10)

user_settings = {
    'show_borders': show_borders, 
    'top_offset_skip': s_top_offset,
    
    # مواقع ثابتة
    'brand_pos_y': s_brand_pos,
    'en_pos_y': s_en_pos,
    'ar_pos_y': s_ar_pos,
    'offer_pos_y': s_offer_pos,
    'barcode_bottom_margin': s_bc_bottom,
    
    # أحجام الخطوط
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

        if final_df.empty:
            st.error("لا توجد أصناف.")
        else:
            st.subheader("🔍 الفلتر")
            c1, c2, c3 = st.columns(3)
            cats = ['All'] + sorted(list(final_df['Category'].dropna().unique()))
            brands = ['All'] + sorted(list(final_df['Brand'].dropna().unique()))
            offers_list = ['All'] + sorted(list(final_df['Offer Description EN'].astype(str).dropna().unique()))

            sel_cat = c1.selectbox("القسم", cats)
            sel_brand = c2.selectbox("البراند", brands)
            sel_offer = c3.selectbox("العرض", offers_list)

            if sel_cat != 'All': final_df = final_df[final_df['Category'] == sel_cat]
            if sel_brand != 'All': final_df = final_df[final_df['Brand'] == sel_brand]
            if sel_offer != 'All': final_df = final_df[final_df['Offer Description EN'].astype(str) == sel_offer]
            
            st.success(f"العدد: {len(final_df)}")
            
            if st.button("👁️ معاينة", type="primary"):
                preview_pdf = generate_pdf(final_df.head(6), user_settings)
                st.session_state['preview_pdf'] = preview_pdf
            
            if 'preview_pdf' in st.session_state:
                st.markdown("---")
                col_prev, col_down = st.columns([2, 1])
                with col_prev:
                    doc = fitz.open(stream=st.session_state['preview_pdf'].getvalue(), filetype="pdf")
                    pix = doc.load_page(0).get_pixmap(dpi=150)
                    st.image(pix.tobytes("png"), width=600)
                with col_down:
                    full_pdf = generate_pdf(final_df, user_settings)
                    st.download_button("📥 تحميل PDF", full_pdf, "Fixed_Layout_Offers.pdf", "application/pdf")

    except Exception as e:
        st.error(f"خطأ: {e}")
