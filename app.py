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
from PIL import Image

# ==========================================
# 1. الثوابت والأبعاد (سم) - Hardcoded
# ==========================================
DIM_ROW1_TOP_CM = 7.7    
DIM_ROW2_TOP_CM = 22.5   
DIM_YELLOW_H_CM = 7.5    
DIM_CARD_W_CM = 7.0      
DIM_GAP_CM = 0.7         

# الأبعاد الداخلية
POS_BRAND_Y_CM = 0.6
POS_EN_Y_CM = 1.6
POS_AR_Y_CM = 2.6
POS_BARCODE_BOTTOM_CM = 0.8

# إعدادات الملفات
FONT_PATH = "arial.ttf"
FONT_NAME = "CustomArial"
TEMPLATE_PATH = "template.png" # اسم ملف صورة الخلفية

st.set_page_config(page_title="Offers Generator (Live Preview)", layout="wide", page_icon="👁️")

def cm2p(cm):
    return cm * 28.3465

def setup_resources():
    # التحقق من الخط
    font_ok = False
    if os.path.exists(FONT_PATH):
        try:
            pdfmetrics.registerFont(TTFont(FONT_NAME, FONT_PATH))
            font_ok = True
        except:
            pass
            
    # التحقق من صورة القالب
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

    # 4. العرض
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

def generate_pdf(df, preview_mode=False):
    """
    إنشاء ملف PDF.
    preview_mode=True: ينشئ صفحة واحدة فقط ويضع صورة القالب كخلفية.
    preview_mode=False: ينشئ ملف الطباعة النهائي بدون خلفية.
    """
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    page_w_pt, page_h_pt = A4 
    
    row1_top = cm2p(DIM_ROW1_TOP_CM)
    row2_top = cm2p(DIM_ROW2_TOP_CM)
    card_w_pt = cm2p(DIM_CARD_W_CM)
    gap_w_pt = cm2p(DIM_GAP_CM)
    
    cols = 3
    cards_per_page = 6
    
    # إذا كنا في وضع المعاينة، نأخذ أول 6 كروت فقط
    if preview_mode:
        df_to_process = df.head(cards_per_page)
    else:
        df_to_process = df
        
    for i, (_, row) in enumerate(df_to_process.iterrows()):
        if i > 0 and i % cards_per_page == 0:
            c.showPage()
            
        # في بداية كل صفحة، إذا كنا في وضع المعاينة، نرسم الخلفية
        if preview_mode and (i % cards_per_page == 0):
            if has_template:
                # رسم صورة القالب لتملأ الصفحة بالكامل
                c.drawImage(TEMPLATE_PATH, 0, 0, width=page_w_pt, height=page_h_pt)
        
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

def create_preview_image(df):
    """إنشاء صورة معاينة من أول صفحة PDF مع الخلفية"""
    # 1. إنشاء PDF صفحة واحدة مع الخلفية
    pdf_buffer = generate_pdf(df, preview_mode=True)
    
    # 2. تحويل PDF إلى صورة عالية الدقة
    doc = fitz.open(stream=pdf_buffer.getvalue(), filetype="pdf")
    page = doc.load_page(0)
    pix = page.get_pixmap(dpi=150) # دقة 150 مناسبة للعرض
    
    # 3. تحويل Pixmap إلى تنسيق يمكن عرضه في Streamlit
    img_data = pix.tobytes("png")
    return img_data

# ==========================================
# 3. الواجهة
# ==========================================
st.title("🖨️ Offers Generator (Live Preview on Template)")

if not has_font:
    st.warning("⚠️ ملف الخط `arial.ttf` مفقود! اللغة العربية لن تظهر.")
if not has_template:
    st.error(f"⚠️ ملف صورة القالب `{TEMPLATE_PATH}` مفقود! المعاينة الحية لن تعمل.")

st.sidebar.header("1. البيانات")
offers_file = st.sidebar.file_uploader("ملف العروض (Excel)", type=['xlsx'])
stock_file = st.sidebar.file_uploader("ملف المخزون (Excel)", type=['xlsx'])
min_qty = st.sidebar.number_input("أقل كمية للطباعة", 1, 100, 2)

if offers_file and stock_file:
    try:
        df1 = pd.read_excel(offers_file)
        df2 = pd.read_excel(stock_file)
        
        df1['Item Number'] = df1['Item Number'].astype(str).str.replace('.0', '')
        df2['Item Number'] = df2['Item Number'].astype(str).str.replace('.0', '')
        
        merged = pd.merge(df1, df2[['Item Number', 'Quantity']], on='Item Number', how='left')
        
        base_df = merged[merged['Quantity'] >= min_qty].copy()

        if base_df.empty:
            st.warning("لا توجد أصناف تحقق شرط الكمية.")
        else:
            st.markdown("---")
            st.subheader("🔍 الفلاتر الديناميكية")
            
            c1, c2, c3 = st.columns(3)

            all_cats = ['All'] + sorted(list(base_df['Category'].astype(str).unique()))
            sel_cat = c1.selectbox("1. القسم (Category)", all_cats)
            
            if sel_cat == 'All':
                df_after_cat = base_df
            else:
                df_after_cat = base_df[base_df['Category'].astype(str) == sel_cat]

            available_brands = ['All'] + sorted(list(df_after_cat['Brand'].astype(str).unique()))
            sel_brand = c2.selectbox("2. البراند (Brand)", available_brands)
            
            if sel_brand == 'All':
                df_after_brand = df_after_cat
            else:
                df_after_brand = df_after_cat[df_after_cat['Brand'].astype(str) == sel_brand]

            available_offers = ['All'] + sorted(list(df_after_brand['Offer Description EN'].astype(str).unique()))
            sel_offer = c3.selectbox("3. العرض (Offer)", available_offers)
            
            if sel_offer == 'All':
                final_df = df_after_brand
            else:
                final_df = df_after_brand[df_after_brand['Offer Description EN'].astype(str) == sel_offer]

            st.info(f"العدد النهائي للطباعة: {len(final_df)}")
            
            if not final_df.empty:
                # زر المعاينة
                if has_template:
                    if st.button("👁️ معاينة حية على القالب", type="primary"):
                        with st.spinner("جاري إنشاء المعاينة..."):
                            # إنشاء صورة المعاينة
                            preview_img = create_preview_image(final_df)
                            st.session_state['preview_img'] = preview_img
                
                # عرض المعاينة إذا كانت موجودة في الذاكرة
                if 'preview_img' in st.session_state:
                    st.markdown("---")
                    st.subheader("معاينة الطباعة (أول صفحة)")
                    st.image(st.session_state['preview_img'], caption="هكذا ستظهر الطباعة على الورق", use_column_width=True, output_format="PNG")
                    st.markdown("---")

                # زر التحميل النهائي
                # نولد PDF الطباعة (بدون خلفية)
                pdf_data = generate_pdf(final_df, preview_mode=False)
                st.download_button(
                    label="📥 تحميل ملف الطباعة النهائي (PDF)",
                    data=pdf_data,
                    file_name="Final_Print_Offers.pdf",
                    mime="application/pdf",
                )
            else:
                st.warning("لا توجد نتائج.")

    except Exception as e:
        st.error(f"حدث خطأ: {e}")
