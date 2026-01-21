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
from reportlab.lib.utils import simpleSplit
import arabic_reshaper
from bidi.algorithm import get_display

# ==========================================
# 1. إعدادات القياسات (بالسنتيمتر)
# ==========================================
DIM_ROW1_TOP_CM = 7.7
DIM_ROW2_TOP_CM = 22.5
DIM_YELLOW_H_CM = 7.5

CENTERS_FROM_RIGHT_CM = [3.5, 10.5, 17.9]

POS_BRAND_Y_CM = 0.6
POS_EN_Y_CM = 1.6
POS_AR_Y_CM = 2.6
POS_BARCODE_BOTTOM_CM = 0.8

FONT_PATH = "arial.ttf"
FONT_NAME = "CustomArial"
TEMPLATE_PATH = "template.png"

st.set_page_config(page_title="Offers Generator (Fixed)", layout="wide", page_icon="✅")

def cm2p(cm):
    return cm * 28.3465

def setup_resources():
    font_ok = False
    if os.path.exists(FONT_PATH):
        try:
            pdfmetrics.registerFont(TTFont(FONT_NAME, FONT_PATH))
            font_ok = True
        except:
            pass
    template_ok = os.path.exists(TEMPLATE_PATH)
    return font_ok, template_ok

has_font, has_template = setup_resources()

# caches for speed
_barcode_cache = {}
_arabic_cache = {}

def process_text(text, is_arabic=False):
    # caching processed Arabic text to avoid repeated expensive reshaping
    if pd.isna(text) or text == "":
        return ""
    text = str(text)
    if is_arabic and has_font:
        cached = _arabic_cache.get(text)
        if cached is not None:
            return cached
        try:
            reshaped = arabic_reshaper.reshape(text)
            displayed = get_display(reshaped)
            _arabic_cache[text] = displayed
            return displayed
        except Exception:
            return text
    return text

# ==========================================
# 2. دوال الرسم (تم إصلاح الأداء هنا)
# ==========================================

def _fit_font_size_for_lines(text, font_name, max_font_size, min_font_size, max_width, max_lines=2):
    """
    Find the largest font size that produces <= max_lines using binary search,
    which reduces calls to simpleSplit vs linear decrement.
    Returns chosen_size and corresponding lines list.
    """
    lo = min_font_size
    hi = max_font_size
    best_size = min_font_size
    best_lines = simpleSplit(text, font_name, best_size, max_width)
    # if even min_font_size yields too many lines, return it
    if len(best_lines) <= max_lines:
        # try to grow to the max via binary search
        while lo <= hi:
            mid = (lo + hi) / 2.0
            lines_mid = simpleSplit(text, font_name, mid, max_width)
            if len(lines_mid) <= max_lines:
                best_size = mid
                best_lines = lines_mid
                lo = mid + 0.1
            else:
                hi = mid - 0.1
        return best_size, best_lines
    else:
        # min font already produces > max_lines, return min
        return best_size, best_lines

def draw_text_multiline(c, text, y_center, max_width, font_name, max_font_size, min_font_size=6, color=(0,0,0), is_bold=False):
    """
    Faster drawing: uses binary-search font fitting and caches few things.
    """
    c.setFillColorRGB(*color)
    c.setStrokeColorRGB(*color)

    # 1. compute font size and lines using a faster search
    current_size, lines = _fit_font_size_for_lines(text, font_name, max_font_size, min_font_size, max_width, max_lines=2)

    # 2. compute vertical start
    leading = current_size * 1.2
    total_height = leading * len(lines)
    start_y = y_center + (total_height / 2) - (current_size * 0.8)

    # 3. draw each line (local bindings for speed)
    pdfmetrics_stringWidth = pdfmetrics.stringWidth
    for line in lines:
        text_width = pdfmetrics_stringWidth(line, font_name, current_size)
        start_x = -(text_width / 2)  # center

        if is_bold:
            # fake bold using stroke + fill (kept as before)
            c.setLineWidth(0.5 if current_size < 12 else 0.8)
            text_obj = c.beginText()
            text_obj.setFont(font_name, current_size)
            # render mode 2 -> Fill + Stroke
            try:
                text_obj.setTextRenderMode(2)
            except Exception:
                pass
            text_obj.setTextOrigin(start_x, start_y)
            text_obj.textOut(line)
            c.drawText(text_obj)
            c.setLineWidth(0)
        else:
            c.setFont(font_name, current_size)
            c.drawString(start_x, start_y, line)

        start_y -= leading

    # reset
    c.setFillColorRGB(0,0,0)
    c.setStrokeColorRGB(0,0,0)

def _get_barcode(item_code):
    """
    Return a cached Code128 barcode object (faster when many rows share codes).
    """
    if not item_code:
        return None
    bc = _barcode_cache.get(item_code)
    if bc is None:
        try:
            bc = code128.Code128(item_code, barHeight=20, barWidth=1.2)
            _barcode_cache[item_code] = bc
        except Exception:
            return None
    return bc

def draw_card_content(c, row, settings, font_base_name):
    # row is a plain dict
    item_code = str(row.get('Item Number', '')).replace('.0', '')
    desc_en = row.get('Item Description EN', '')
    desc_ar = row.get('Item Description AR', '')
    brand_txt = row.get('Brand', '')
    offer_txt = row.get('Offer Description EN', '')

    height = cm2p(DIM_YELLOW_H_CM)
    max_text_width = cm2p(7.0) * 0.90

    # 1. brand
    brand_y = -cm2p(POS_BRAND_Y_CM)
    if has_font:
        draw_text_multiline(c, str(brand_txt or ""), brand_y, max_text_width,
                            font_base_name, settings['font_brand'], min_font_size=8, is_bold=True)
    else:
        c.setFont("Helvetica-Bold", settings['font_brand'])
        c.drawCentredString(0, brand_y, str(brand_txt or ""))

    # 2. English name
    en_y = -cm2p(POS_EN_Y_CM)
    draw_text_multiline(c, str(desc_en or ""), en_y, max_text_width,
                        font_base_name, settings['font_name'], min_font_size=7)

    # 3. Arabic name (processed + cached)
    ar_y = -cm2p(POS_AR_Y_CM)
    ar_txt_proc = process_text(desc_ar, is_arabic=True)
    draw_text_multiline(c, ar_txt_proc, ar_y, max_text_width,
                        font_base_name, settings['font_name'], min_font_size=7)

    # 4. offer
    offer_y = -(height / 2) - 5
    if has_font:
        draw_text_multiline(c*

