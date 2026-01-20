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
st.set_page_config(page_title="Offers Generator (Final Dimensions)", layout="wide", page_icon="📏")

FONT_PATH = "arial.ttf"
FONT_NAME = "CustomArial"

def cm2p(cm):
    """تحويل من سنتيمتر إلى نقاط (Points)"""
    return cm * 28.3465

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
# 2. محرك الرسم (الإحداثيات المطلقة)
# ==========================================

def draw_text_auto_shrink(c, text, center_x, y, max_width, font_name, max_font_size, min_font_size=6, color=(0,0,0), is_bold=False):
    """رسم نص مع تصغير تلقائي"""
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
        text_obj.setTextRenderMode(2) 
        text_obj.setFont(font_name, current_size)
        start_x = center_x - (text_width / 2)
        text_obj.setTextOrigin(start_x, y)
        text_obj.textOut(text)
        c.drawText(text_obj)
        c.setLineWidth(0)
    else:
        c.setFont(font_name, current_size)
        c.drawCentredString(center_x, y, text)
    
    c.setFillColorRGB(0, 0, 0)
    c.setStrokeColorRGB(0, 0, 0)

def draw_card_content(c, width, height, row, settings):
    """
    رسم محتوى الكارت داخل المنطقة الصفراء فقط.
    نقطة (0,0) هي الزاوية اليسرى العليا للمنطقة الصفراء.
    """
    item_code = str(row.get('Item Number', '')).replace('.0', '')
    desc_en = row.get('Item Description EN', '') 
    desc_ar = row.get('Item Description AR', '')
    brand_txt = row.get('Brand', '')
    offer_txt = row.get('Offer Description EN', '')

    center_x = width / 2 
    max_text_width = width * 0.95 # هامش جانبي بسيط

    # --- حدود التجربة (داخل الأصفر فقط) ---
    if settings['show_borders']:
        c.setLineWidth(1)
        c.setStrokeColorRGB(1, 0, 0) # أحمر
        c.rect(0, -height, width, height) # رسم مستطيل يمثل المنطقة الصفراء
        c.setLineWidth(0)

    # جميع الإحداثيات Y ستكون بالسالب لأننا ننزل لأسفل من قمة الأصفر

    # 1. البراند (Brand)
    brand_y = -cm2p(settings['pos_brand_cm'])
    if has_font:
        draw_text_auto_shrink(c, str(brand_txt), center_x, brand_y, max_text_width, 
                              FONT_NAME, settings['font_brand'], min_font_size=8, is_bold=True)
    else:
        c.setFont("Helvetica-Bold", settings['font_brand'])
        c.drawCentredString(center_x, brand_y, str(brand_txt))

    # 2. الاسم الإنجليزي
    en_y = -cm2p(settings['pos_en_cm'])
    draw_text_auto_shrink(c, str(desc_en), center_x, en_y, max_text_width, 
                          FONT_NAME if has_font else "Helvetica", settings['font_name'], min_font_size=8)

    # 3. الاسم العربي
    ar_y = -cm2p(settings['pos_ar_cm'])
    ar_txt_proc = process_text(desc_ar, is_arabic=True)
    draw_text_auto_shrink(c, ar_txt_proc, center_x, ar_y, max_text_width, 
                          FONT_NAME if has_font else "Helvetica", settings['font_name'], min_font_size=8)

    # 4. العرض (Offer) - في المنتصف
    # يتم حسابه تلقائياً بناءً على ارتفاع المنطقة الصفراء
    offer_y = -(height / 2) - 5 # إزاحة بسيطة للأعلى
    if has_font:
        draw_text_auto_shrink(c, str(offer_txt), center_x, offer_y, max_text_width, 
                              FONT_NAME, settings['font_offer'], min_font_size=12, 
                              color=(0.85, 0.21, 0.27), is_bold=True)
    else:
        c.setFont("Helvetica-Bold", settings['font_offer'])
        c.setFillColorRGB(0.85, 0.21, 0.27)
        c.drawCentredString(center_x, offer_y, str(offer_txt))

    # 5. الباركود (أسفل الأصفر)
    # يتم وضعه قبل نهاية الأصفر بمسافة محددة
    barcode_y = -height + cm2p(settings['pos_barcode_bottom_cm'])
    
    if item_code:
        try:
            bc_height = settings['barcode_height']
            barcode = code128.Code128(item_code, barHeight=bc_height, barWidth=1.2)
            bc_x = center_x - (barcode.width / 2)
            barcode.drawOn(c, bc_x, barcode_y + 10)
            
            c.setFillColorRGB(0, 0, 0)
            c.setFont("Helvetica", settings['font_barcode'])
            c.drawCentredString(center_x, barcode_y, item_code)
        except:
            pass

def generate_pdf(df, settings):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    page_w_pt, page_h_pt = A4 # 21.0cm x 29.7cm
    
    # تحويل القيم المدخلة من سم إلى نقاط
    row1_top_margin = cm2p(settings['row1_top_cm'])   # 7.7 cm
    row2_top_margin = cm2p(settings['row2_top_cm'])   # 22.5 cm
    yellow_h = cm2p(settings['yellow_height_cm'])     # 7.5 cm
    card_w = cm2p(settings['card_width_cm'])          # 7.0 cm
    gap_w = cm2p(settings['col_gap_cm'])              # 0.7 cm
    
    # تصحيح بداية الصفحة (X Offset Global)
    global_x = cm2p(settings['global_x_cm'])
    
    cols = 3
    cards_per_page = 6
    
    for i, (_, row) in enumerate(df.iterrows()):
        # صفحة جديدة
        if i > 0 and i % cards_per_page == 0:
            c.showPage()
        
        pos_in_page = i % cards_per_page
        col_idx = pos_in_page % cols  # 0, 1, 2
        row_idx = pos_in_page // cols # 0 (Top), 1 (Bottom)
        
        # === حساب الإحداثيات المطلقة ===
        
        # 1. حساب X (الأفقي)
        # المعادلة: إزاحة عامة + (رقم العمود * (عرض الكارت + الفاصل))
        x_start = global_x + (col_idx * (card_w + gap_w))
        
        # 2. حساب Y (الرأسي - بداية الأصفر)
        # نستخدم نظام الإحداثيات من الأسفل (PDF Standard)
        # الصفحة 29.7 سم.
        # الصف الأول يبدأ عند 7.7 سم من القمة -> 29.7 - 7.7
        # الصف الثاني يبدأ عند 22.5 سم من القمة -> 29.7 - 22.5
        
        if row_idx == 0:
            # الصف العلوي
            y_start = page_h_pt - row1_top_margin
        else:
            # الصف السفلي
            y_start = page_h_pt - row2_top_margin
            
        # === العزل والرسم ===
        c.saveState()
        # نقل نقطة الصفر إلى (الركن الأيسر العلوي للمنطقة الصفراء)
        c.translate(x_start, y_start)
        
        # استدعاء دالة الرسم
        draw_card_content(c, card_w, yellow_h, row, settings)
        
        c.restoreState()
        
    c.save()
    buffer.seek(0)
    return buffer

# ==========================================
# 3. واجهة المستخدم
# ==========================================
st.title("🖨️ Offers Generator (Exact Dimensions)")

if not has_font:
    st.warning("⚠️ Font `arial.ttf` missing.")

# --- القائمة الجانبية ---
st.sidebar.header("1. البيانات")
offers_file = st.sidebar.file_uploader("ملف العروض", type=['xlsx'])
stock_file = st.sidebar.file_uploader("ملف المخزون", type=['xlsx'])
min_qty = st.sidebar.number_input("أقل كمية", 2, 100, 2)

st.sidebar.markdown("---")
st.sidebar.header("2. 📏 أبعاد الورقة (سم)")
st.sidebar.info("تم ضبط الأرقام بناءً على قياساتك.")

with st.sidebar.expander("تعديل القياسات الأساسية", expanded=True):
    # القيم الافتراضية كما طلبتها
    s_row1_top = st.number_input("بداية الأصفر العلوي (من القمة)", 0.0, 29.7, 7.7, step=0.1)
    s_row2_top = st.number_input("بداية الأصفر السفلي (من القمة)", 0.0, 29.7, 22.5, step=0.1)
    s_yellow_h = st.number_input("ارتفاع المستطيل الأصفر", 1.0, 15.0, 7.5, step=0.1)
    s_card_w = st.number_input("عرض المنطقة الصفراء", 1.0, 10.0, 7.0, step=0.1)
    s_gap = st.number_input("المسافة الفاصلة بين العمودين", 0.0, 5.0, 0.7, step=0.1)

st.sidebar.markdown("---")
st.sidebar.header("3. 🔧 معايرة الطابعة")
s_global_x = st.number_input("تحريك الصفحة يمين/يسار (سم)", -5.0, 5.0, 0.0, step=0.1, help="موجب لليمين، سالب لليسار")

st.sidebar.markdown("---")
st.sidebar.header("4. التنسيق الداخلي")
show_borders = st.sidebar.checkbox("إظهار حدود الأصفر (للتجربة)", True)

with st.sidebar.expander("مواقع النصوص (داخل الأصفر)", expanded=True):
    st.caption("المسافات بالسنتيمتر من قمة المستطيل الأصفر:")
    s_pos_brand = st.slider("موقع البراند", 0.1, 5.0, 0.5)
    s_pos_en = st.slider("موقع الإنجليزي", 0.5, 5.0, 1.5)
    s_pos_ar = st.slider("موقع العربي", 1.0, 6.0, 2.5)
    s_pos_bc_bot = st.slider("الباركود (من نهاية الأصفر لأعلى)", 0.1, 3.0, 0.8)

with st.sidebar.expander("أحجام الخطوط", expanded=False):
    s_f_brand = st.slider("خط البراند", 8, 20, 12)
    s_f_name = st.slider("خط الأسماء", 6, 18, 10)
    s_f_offer = st.slider("خط العرض", 10, 40, 24)
    s_f_bc = st.slider("خط رقم الباركود", 6, 12, 8)
    s_bc_h = st.slider("ارتفاع أعمدة الباركود", 10, 40, 20)

user_settings = {
    'row1_top_cm': s_row1_top,
    'row2_top_cm': s_row2_top,
    'yellow_height_cm': s_yellow_h,
    'card_width_cm': s_card_w,
    'col_gap_cm': s_gap,
    'global_x_cm': s_global_x,
    
    'show_borders': show_borders,
    'pos_brand_cm': s_pos_brand,
    'pos_en_cm': s_pos_en,
    'pos_ar_cm': s_pos_ar,
    'pos_barcode_bottom_cm': s_pos_bc_bot,
    
    'font_brand': s_f_brand,
    'font_name': s_f_name, 
    'font_offer': s_f_offer,
    'font_barcode': s_f_bc,
    'barcode_height': s_bc_h
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
                    st.download_button("📥 تحميل PDF", full_pdf, "Final_Exact_Cm.pdf", "application/pdf")

    except Exception as e:
        st.error(f"خطأ: {e}")
