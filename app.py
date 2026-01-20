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
# 1. إعدادات النظام والتحويلات
# ==========================================
st.set_page_config(page_title="Offers Generator Pro (Fixed Alignment)", layout="wide", page_icon="🖨️")

FONT_PATH = "arial.ttf"
FONT_NAME = "CustomArial"

def mm2p(mm):
    """تحويل من مليمتر إلى نقاط بدقة عالية"""
    return mm * 2.83465

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
# 2. دوال الرسم (معزولة وثابتة)
# ==========================================

def draw_text_auto_shrink(c, text, center_x, y, max_width, font_name, max_font_size, min_font_size=6, color=(0,0,0), is_bold=False):
    """رسم نص ذكي يقوم بتصغير نفسه ليدخل في الصندوق"""
    current_size = max_font_size
    text_width = pdfmetrics.stringWidth(text, font_name, current_size)
    
    while text_width > max_width and current_size > min_font_size:
        current_size -= 0.5
        text_width = pdfmetrics.stringWidth(text, font_name, current_size)
    
    c.setFillColorRGB(*color)
    c.setStrokeColorRGB(*color)
    
    if is_bold:
        c.setLineWidth(0.5 if current_size < 12 else 0.8)
        text_obj = c.beginText()
        text_obj.setTextRenderMode(2) # Fill + Stroke
        text_obj.setFont(font_name, current_size)
        start_x = center_x - (text_width / 2)
        text_obj.setTextOrigin(start_x, y)
        text_obj.textOut(text)
        c.drawText(text_obj)
        c.setLineWidth(0)
    else:
        c.setFont(font_name, current_size)
        c.drawCentredString(center_x, y, text)
    
    # إعادة تعيين الألوان للأسود لتجنب التأثير على العناصر التالية
    c.setFillColorRGB(0, 0, 0)
    c.setStrokeColorRGB(0, 0, 0)

def draw_label(c, x, y, w, h, row, settings):
    """رسم كارت واحد في مكان محدد"""
    item_code = str(row.get('Item Number', '')).replace('.0', '')
    desc_en = row.get('Item Description EN', '') 
    desc_ar = row.get('Item Description AR', '')
    brand_txt = row.get('Brand', '')
    offer_txt = row.get('Offer Description EN', '')

    center_x = x + (w / 2)
    max_text_width = w * 0.92

    # --- منطقة الحدود (للتجربة) ---
    if settings['show_borders']:
        c.setLineWidth(0.5)
        c.setStrokeColorRGB(0.8, 0.8, 0.8)
        c.rect(x, y, w, h)
        
        yellow_start_y = (y + h) - mm2p(settings['yellow_start_mm'])
        c.setStrokeColorRGB(1, 0, 0)
        c.setLineWidth(1)
        c.line(x, yellow_start_y, x+w, yellow_start_y)

    # 1. المنطقة العلوية (اسم الصيدلية)
    header_y = (y + h) - mm2p(10)
    c.setFillColorRGB(0.4, 0.4, 0.4) 
    c.setFont(FONT_NAME if has_font else "Helvetica", settings['header_font_size'])
    pharmacy_name = process_text("Al-Dawaa Pharmacy | صيدلية الدواء", is_arabic=True)
    c.drawCentredString(center_x, header_y, pharmacy_name)

    # 2. المنطقة الصفراء
    yellow_zero_y = (y + h) - mm2p(settings['yellow_start_mm'])

    # أ. البراند
    brand_y = yellow_zero_y - mm2p(settings['brand_pos_mm'])
    if has_font:
        draw_text_auto_shrink(c, str(brand_txt), center_x, brand_y, max_text_width, 
                              FONT_NAME, settings['brand_font_size'], min_font_size=8, is_bold=True)
    else:
        c.setFont("Helvetica-Bold", settings['brand_font_size'])
        c.drawCentredString(center_x, brand_y, str(brand_txt))

    # ب. الاسم الإنجليزي
    en_y = yellow_zero_y - mm2p(settings['en_pos_mm'])
    draw_text_auto_shrink(c, str(desc_en), center_x, en_y, max_text_width, 
                          FONT_NAME if has_font else "Helvetica", settings['name_font_size'], min_font_size=8)

    # ج. الاسم العربي
    ar_y = yellow_zero_y - mm2p(settings['ar_pos_mm'])
    ar_txt_proc = process_text(desc_ar, is_arabic=True)
    draw_text_auto_shrink(c, ar_txt_proc, center_x, ar_y, max_text_width, 
                          FONT_NAME if has_font else "Helvetica", settings['name_font_size'], min_font_size=8)

    # د. العرض
    offer_y = yellow_zero_y - mm2p(settings['offer_pos_mm'])
    if has_font:
        draw_text_auto_shrink(c, str(offer_txt), center_x, offer_y, max_text_width, 
                              FONT_NAME, settings['price_font_size'], min_font_size=12, 
                              color=(0.85, 0.21, 0.27), is_bold=True)
    else:
        c.setFont("Helvetica-Bold", settings['price_font_size'])
        c.setFillColorRGB(0.85, 0.21, 0.27)
        c.drawCentredString(center_x, offer_y, str(offer_txt))

    # هـ. الباركود
    barcode_y = y + mm2p(settings['barcode_bottom_mm'])
    
    if item_code:
        try:
            bc_height = settings['barcode_height']
            barcode = code128.Code128(item_code, barHeight=bc_height, barWidth=1.2)
            bc_x = center_x - (barcode.width / 2)
            barcode.drawOn(c, bc_x, barcode_y + 10)
            
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
        # إدارة الصفحات الجديدة
        if i > 0 and i % (cols * rows) == 0:
            c.showPage() # إعادة ضبط الإحداثيات للصفحة الجديدة
        
        # حساب الموقع
        pos = i % (cols * rows)
        col_idx = pos % cols
        row_idx = pos // cols 
        
        base_x = col_idx * block_w
        base_y = page_h - ((row_idx + 1) * block_h)
        
        # حساب الإزاحة الخاصة بكل نصف
        if row_idx == 0:
            final_x = base_x + mm2p(settings['top_x_mm'])
            final_y = base_y + mm2p(settings['top_y_mm'])
        else:
            final_x = base_x + mm2p(settings['bottom_x_mm'])
            final_y = base_y + mm2p(settings['bottom_y_mm'])
        
        # === هام جداً: عزل الحالة (State Isolation) ===
        # هذا الأمر يحفظ حالة الصفحة قبل الرسم ويستعيدها بعده
        # مما يمنع أي تداخل أو انحراف في القياسات بين الكروت
        c.saveState()
        draw_label(c, final_x, final_y, block_w, block_h, row, settings)
        c.restoreState()
        # ============================================
        
    c.save()
    buffer.seek(0)
    return buffer

# ==========================================
# 3. واجهة المستخدم
# ==========================================
st.title("🖨️ Offers Generator Pro (Stable Alignment)")

if not has_font:
    st.warning("⚠️ Font `arial.ttf` missing.")

# القائمة الجانبية
st.sidebar.header("1. البيانات")
offers_file = st.sidebar.file_uploader("ملف العروض", type=['xlsx'])
stock_file = st.sidebar.file_uploader("ملف المخزون", type=['xlsx'])
min_qty = st.sidebar.number_input("أقل كمية", 2, 100, 2)

st.sidebar.markdown("---")
st.sidebar.header("2. 🎚️ معايرة الطابعة (Printer Calibration)")
st.sidebar.info("القيم هنا بالميليمتر. استخدمها لضبط الطباعة على الورق.")

tab_top, tab_bottom = st.sidebar.tabs(["⬆️ النصف العلوي", "⬇️ النصف السفلي"])

with tab_top:
    st.caption("ضبط الصف الأول (الثلاثة كروت العليا)")
    s_top_x = st.number_input("تحريك أفقي (Top X)", -50.0, 50.0, 0.0, step=0.5, key="tx")
    s_top_y = st.number_input("تحريك رأسي (Top Y)", -50.0, 50.0, 0.0, step=0.5, key="ty")

with tab_bottom:
    st.caption("ضبط الصف الثاني (الثلاثة كروت السفلى)")
    s_bot_x = st.number_input("تحريك أفقي (Bottom X)", -50.0, 50.0, 0.0, step=0.5, key="bx")
    s_bot_y = st.number_input("تحريك رأسي (Bottom Y)", -50.0, 50.0, 0.0, step=0.5, key="by")

st.sidebar.markdown("---")
st.sidebar.header("3. ضبط التصميم الداخلي")
show_borders = st.sidebar.checkbox("إظهار حدود (للضبط)", False)

with st.sidebar.expander("📍 المسافات (بالمليمتر)", expanded=True):
    # القيم الثابتة الافتراضية كما طلبت
    s_yellow_start = st.slider("بداية الأصفر (من أعلى الكارت)", 40, 80, 50)
    
    st.caption("مواقع العناصر (من بداية الأصفر لأسفل):")
    s_brand_pos = st.slider("موقع البراند", 2, 30, 10)
    s_en_pos = st.slider("موقع الإنجليزي", 5, 50, 31)
    s_ar_pos = st.slider("موقع العربي", 10, 60, 54)
    s_offer_pos = st.slider("موقع العرض (الوسط)", 20, 80, 84)
    s_bc_bottom = st.slider("الباركود من الأسفل", 2, 40, 15)

with st.sidebar.expander("🅰️ أحجام الخطوط", expanded=False):
    s_header_font = st.slider("اسم الصيدلية", 6, 14, 8)
    s_brand_font = st.slider("البراند", 8, 24, 12)
    s_name_font = st.slider("الأسماء", 6, 20, 10)
    s_price_font = st.slider("العرض", 10, 50, 24)
    s_bc_h = st.slider("ارتفاع الباركود", 10, 40, 20)
    s_bc_font = st.slider("خط الباركود", 6, 12, 8)

user_settings = {
    'top_x_mm': s_top_x, 'top_y_mm': s_top_y,
    'bottom_x_mm': s_bot_x, 'bottom_y_mm': s_bot_y,
    
    'show_borders': show_borders,
    'yellow_start_mm': s_yellow_start,
    'brand_pos_mm': s_brand_pos,
    'en_pos_mm': s_en_pos,
    'ar_pos_mm': s_ar_pos,
    'offer_pos_mm': s_offer_pos,
    'barcode_bottom_mm': s_bc_bottom,
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
            st.error("لا توجد بيانات.")
        else:
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
                    st.download_button("📥 تحميل PDF", full_pdf, "Stable_Offers.pdf", "application/pdf")

    except Exception as e:
        st.error(f"خطأ: {e}")
