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
# 1. الثوابت والأبعاد (سم) - Hardcoded
# ==========================================
DIM_ROW1_TOP_CM = 7.7    # بداية الأصفر العلوي
DIM_ROW2_TOP_CM = 22.5   # بداية الأصفر السفلي
DIM_YELLOW_H_CM = 7.5    # ارتفاع المنطقة الصفراء
DIM_CARD_W_CM = 7.0      # عرض المنطقة الصفراء
DIM_GAP_CM = 0.7         # الفاصل بين الأعمدة

# الأبعاد الداخلية (تنسيق النصوص داخل الأصفر)
POS_BRAND_Y_CM = 0.6
POS_EN_Y_CM = 1.6
POS_AR_Y_CM = 2.6
POS_BARCODE_BOTTOM_CM = 0.8

# إعدادات الخطوط
FONT_PATH = "arial.ttf"
FONT_NAME = "CustomArial"

st.set_page_config(page_title="Offers Generator (Fixed + Filters)", layout="wide", page_icon="🏷️")

def cm2p(cm):
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
# 2. دوال الرسم
# ==========================================

def draw_text_auto_shrink(c, text, center_x, y, max_width, font_name, max_font_size, min_font_size=6, color=(0,0,0), is_bold=False):
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

def draw_card_content(c, row):
    item_code = str(row.get('Item Number', '')).replace('.0', '')
    desc_en = row.get('Item Description EN', '') 
    desc_ar = row.get('Item Description AR', '')
    brand_txt = row.get('Brand', '')
    offer_txt = row.get('Offer Description EN', '')

    width = cm2p(DIM_CARD_W_CM)
    height = cm2p(DIM_YELLOW_H_CM)
    center_x = width / 2 
    max_text_width = width * 0.95

    # 1. البراند
    brand_y = -cm2p(POS_BRAND_Y_CM)
    if has_font:
        draw_text_auto_shrink(c, str(brand_txt), center_x, brand_y, max_text_width, 
                              FONT_NAME, 12, min_font_size=8, is_bold=True)
    else:
        c.setFont("Helvetica-Bold", 12)
        c.drawCentredString(center_x, brand_y, str(brand_txt))

    # 2. الاسم الإنجليزي
    en_y = -cm2p(POS_EN_Y_CM)
    draw_text_auto_shrink(c, str(desc_en), center_x, en_y, max_text_width, 
                          FONT_NAME if has_font else "Helvetica", 10, min_font_size=8)

    # 3. الاسم العربي
    ar_y = -cm2p(POS_AR_Y_CM)
    ar_txt_proc = process_text(desc_ar, is_arabic=True)
    draw_text_auto_shrink(c, ar_txt_proc, center_x, ar_y, max_text_width, 
                          FONT_NAME if has_font else "Helvetica", 10, min_font_size=8)

    # 4. العرض (في المنتصف)
    offer_y = -(height / 2) - 5 
    if has_font:
        draw_text_auto_shrink(c, str(offer_txt), center_x, offer_y, max_text_width, 
                              FONT_NAME, 24, min_font_size=12, 
                              color=(0.85, 0.21, 0.27), is_bold=True)
    else:
        c.setFont("Helvetica-Bold", 24)
        c.setFillColorRGB(0.85, 0.21, 0.27)
        c.drawCentredString(center_x, offer_y, str(offer_txt))

    # 5. الباركود
    barcode_y = -height + cm2p(POS_BARCODE_BOTTOM_CM)
    
    if item_code:
        try:
            barcode = code128.Code128(item_code, barHeight=20, barWidth=1.2)
            bc_x = center_x - (barcode.width / 2)
            barcode.drawOn(c, bc_x, barcode_y + 10)
            
            c.setFillColorRGB(0, 0, 0)
            c.setFont("Helvetica", 8)
            c.drawCentredString(center_x, barcode_y, item_code)
        except:
            pass

def generate_pdf(df):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    page_w_pt, page_h_pt = A4 
    
    row1_top = cm2p(DIM_ROW1_TOP_CM)
    row2_top = cm2p(DIM_ROW2_TOP_CM)
    card_w_pt = cm2p(DIM_CARD_W_CM)
    gap_w_pt = cm2p(DIM_GAP_CM)
    
    cols = 3
    cards_per_page = 6
    
    for i, (_, row) in enumerate(df.iterrows()):
        if i > 0 and i % cards_per_page == 0:
            c.showPage()
        
        pos_in_page = i % cards_per_page
        col_idx = pos_in_page % cols
        row_idx = pos_in_page // cols
        
        x_start = col_idx * (card_w_pt + gap_w_pt)
        
        if row_idx == 0:
            y_start = page_h_pt - row1_top
        else:
            y_start = page_h_pt - row2_top
            
        c.saveState()
        c.translate(x_start, y_start)
        draw_card_content(c, row)
        c.restoreState()
        
    c.save()
    buffer.seek(0)
    return buffer

# ==========================================
# 3. الواجهة (مع الفلاتر)
# ==========================================
st.title("🖨️ Offers Generator (Fixed Layout + Filters)")

if not has_font:
    st.error("⚠️ ملف الخط `arial.ttf` مفقود! اللغة العربية لن تظهر.")

st.sidebar.header("1. البيانات")
offers_file = st.sidebar.file_uploader("ملف العروض (Excel)", type=['xlsx'])
stock_file = st.sidebar.file_uploader("ملف المخزون (Excel)", type=['xlsx'])
min_qty = st.sidebar.number_input("أقل كمية للطباعة", 1, 100, 2)

if offers_file and stock_file:
    try:
        # 1. قراءة البيانات
        df1 = pd.read_excel(offers_file)
        df2 = pd.read_excel(stock_file)
        
        df1['Item Number'] = df1['Item Number'].astype(str).str.replace('.0', '')
        df2['Item Number'] = df2['Item Number'].astype(str).str.replace('.0', '')
        
        merged = pd.merge(df1, df2[['Item Number', 'Quantity']], on='Item Number', how='left')
        
        # 2. فلتر الكمية الأول
        base_df = merged[merged['Quantity'] >= min_qty].copy()

        if base_df.empty:
            st.warning("لا توجد أصناف تحقق شرط الكمية.")
        else:
            st.markdown("---")
            st.subheader("🔍 تصفية البيانات (Filters)")
            
            # 3. إعداد القوائم المنسدلة
            # نستخدم astype(str) لضمان عدم حدوث خطأ إذا كانت البيانات مختلطة
            cats = ['All'] + sorted(list(base_df['Category'].astype(str).unique()))
            brands = ['All'] + sorted(list(base_df['Brand'].astype(str).unique()))
            offers_list = ['All'] + sorted(list(base_df['Offer Description EN'].astype(str).unique()))

            c1, c2, c3 = st.columns(3)
            
            sel_cat = c1.selectbox("القسم (Category)", cats)
            sel_brand = c2.selectbox("البراند (Brand)", brands)
            sel_offer = c3.selectbox("نوع العرض (Offer)", offers_list)

            # 4. تطبيق الفلاتر
            final_df = base_df.copy()
            
            if sel_cat != 'All':
                final_df = final_df[final_df['Category'].astype(str) == sel_cat]
            
            if sel_brand != 'All':
                final_df = final_df[final_df['Brand'].astype(str) == sel_brand]
                
            if sel_offer != 'All':
                final_df = final_df[final_df['Offer Description EN'].astype(str) == sel_offer]

            # 5. النتائج والتحميل
            st.info(f"عدد البطاقات التي سيتم طباعتها: {len(final_df)}")
            
            if not final_df.empty:
                pdf_data = generate_pdf(final_df)
                st.download_button(
                    label="📥 تحميل ملف الطباعة (PDF)",
                    data=pdf_data,
                    file_name="Filtered_Offers.pdf",
                    mime="application/pdf",
                    type="primary"
                )
            else:
                st.warning("لا توجد نتائج تطابق الفلاتر المختارة.")

    except Exception as e:
        st.error(f"حدث خطأ في قراءة الملفات: {e}")
