import os
import re
import json
import base64
import io
import zipfile
from datetime import datetime

from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from PIL import Image, ImageDraw, ImageFont
from barcode import Code128
from barcode.writer import ImageWriter
import arabic_reshaper
from bidi.algorithm import get_display
import pandas as pd
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
import tempfile
import traceback

# ══════════════════════════════════════════
#  SETUP
# ══════════════════════════════════════════
app = Flask(__name__, static_folder='.', static_url_path='')
CORS(app)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# مسارات الخطوط
FONT_BOLD = os.path.join(BASE_DIR, "sst-arabic-bold.ttf")
FONT_LIGHT = os.path.join(BASE_DIR, "alfont_com_AlFont_com_SST-Arabic-Light-2 (1).ttf")


def load_font(path, size):
    try:
        return ImageFont.truetype(path, size)
    except Exception:
        return ImageFont.load_default()


def arabic(text):
    """تشكيل النص العربي بشكل صحيح"""
    try:
        return get_display(arabic_reshaper.reshape(str(text)))
    except Exception:
        return str(text)


def safe_filename(text):
    return re.sub(r'[\\/:*?"<>|]', '', str(text)).strip() or "card"


def wrap_text_lines(text, font, max_width, draw):
    """تقسيم النص لأسطر مع مراعاة الكلمات العربية"""
    if not text:
        return []

    bbox = draw.textbbox((0, 0), text, font=font)
    if (bbox[2] - bbox[0]) <= max_width:
        return [text]

    words = text.split()
    lines = []
    current_line = ""

    for word in words:
        test_line = f"{current_line} {word}".strip() if current_line else word
        bbox = draw.textbbox((0, 0), test_line, font=font)
        if (bbox[2] - bbox[0]) <= max_width:
            current_line = test_line
        else:
            if current_line:
                lines.append(current_line)
            current_line = word

    if current_line:
        lines.append(current_line)

    return lines if lines else [text]


def make_barcode_img(value, w, h):
    """إنشاء باركود Code128 حقيقي"""
    try:
        value_str = str(value).strip()
        if not value_str or value_str == "nan":
            return None

        writer = ImageWriter()
        writer.set_options({
            "write_text": False,
            "quiet_zone": 1,
            "module_width": 0.45,
            "module_height": 28,
            "background": "white",
            "foreground": "black"
        })
        buf = io.BytesIO()
        barcode_obj = Code128(value_str, writer=writer)
        barcode_obj.write(buf)
        buf.seek(0)
        img = Image.open(buf).convert("RGBA")

        # تحويل الأبيض لشفاف
        data = img.getdata()
        new_data = []
        for p in data:
            if p[0] > 240 and p[1] > 240 and p[2] > 240:
                new_data.append((255, 255, 255, 0))
            else:
                new_data.append(p)

        img.putdata(new_data)
        img = img.resize((int(w), int(h)), Image.Resampling.LANCZOS)
        return img
    except Exception as e:
        print(f"Barcode error: {e}")
        return None


def load_image_from_base64(base64_str):
    """تحميل صورة من base64"""
    try:
        if "," in base64_str:
            base64_str = base64_str.split(",", 1)[1]
        img_data = base64.b64decode(base64_str)
        img = Image.open(io.BytesIO(img_data))
        return img.convert("RGBA")
    except Exception as e:
        print(f"Error loading image: {e}")
        return None


def draw_field_on_image(draw, img, field, row_data):
    """رسم حقل واحد على الصورة"""
    try:
        ftype = field.get("type", "text")
        col = field.get("col", "")
        x = int(field.get("x", 0))
        y = int(field.get("y", 0))
        w = int(field.get("w", 200))
        h = int(field.get("h", 50))
        size = int(field.get("size", 28))
        color = field.get("color", "#000000")
        bold = field.get("bold", False)
        static = field.get("staticVal", "")
        image_data = field.get("imageData", "")  # للصور المرفوعة

        cx = x + w // 2
        cy = y + h // 2

        # ── الصورة ──
        if ftype == "image":
            if image_data:
                img_obj = load_image_from_base64(image_data)
                if img_obj:
                    # تغيير حجم الصورة
                    img_resized = img_obj.resize((w, h), Image.Resampling.LANCZOS)
                    # لصق الصورة
                    if img_resized.mode == 'RGBA':
                        img.paste(img_resized, (x, y), img_resized)
                    else:
                        img.paste(img_resized, (x, y))
                    return
            # إذا لم توجد صورة، ارسم إطار
            draw.rectangle([x, y, x + w, y + h], outline="gray", width=1)
            draw.text((cx, cy), "📷 صورة", fill="gray", anchor="mm", font=load_font(FONT_LIGHT, 14))
            return

        # ── الباركود ──
        if ftype == "barcode":
            val = row_data.get(col, "") if col else ""
            if not val or str(val).strip() == "" or str(val).lower() == "nan":
                val = static
            if val and str(val).strip():
                bc_img = make_barcode_img(str(val), w, h)
                if bc_img:
                    img.paste(bc_img, (x, y), bc_img)
                    return
                else:
                    draw.rectangle([x, y, x + w, y + h], outline="red", width=2)
                    draw.text((cx, cy), "خطأ في الباركود", fill="red", anchor="mm", font=load_font(FONT_LIGHT, 12))
            else:
                draw.rectangle([x, y, x + w, y + h], outline="gray", width=1)
                draw.text((cx, cy), "لا توجد قيمة", fill="gray", anchor="mm", font=load_font(FONT_LIGHT, 12))
            return

        # ── النص المساعد الثابت (static_label) ──
        if ftype == "static_label":
            raw_text = static
            if not raw_text or raw_text == "nan":
                return
            font_path = FONT_BOLD if bold else FONT_LIGHT
            font = load_font(font_path, size)
            text = arabic(raw_text)
            bbox = draw.textbbox((0, 0), text, font=font)
            tw = bbox[2] - bbox[0]
            th = bbox[3] - bbox[1]
            lx = cx - tw // 2
            ly = cy - th // 2
            draw.text((lx, ly), text, fill=color, font=font)
            return

        # ── النص الديناميكي ──
        if ftype == "static":
            raw_text = static
        else:
            raw = row_data.get(col, "") if col else ""
            raw_text = str(raw) if raw is not None else ""

        if not raw_text or raw_text == "nan":
            return

        font_path = FONT_BOLD if bold else FONT_LIGHT
        font = load_font(font_path, size)
        text = arabic(raw_text)

        lines = wrap_text_lines(text, font, w - 20, draw)
        if not lines:
            return

        line_h = size * 1.3
        total_h = len(lines) * line_h
        start_y = cy - total_h / 2

        for i, ln in enumerate(lines):
            ly = start_y + i * line_h + line_h / 2
            bbox = draw.textbbox((0, 0), ln, font=font)
            tw = bbox[2] - bbox[0]
            lx = cx - tw // 2
            draw.text((lx, ly), ln, fill=color, font=font, anchor="lm")

            if ftype == "price_strike":
                draw.line(
                    (lx, ly - size // 4, lx + tw, ly - size // 4),
                    fill=color,
                    width=max(2, size // 14)
                )
    except Exception as e:
        print(f"Error drawing field {field.get('name')}: {e}")
        traceback.print_exc()


# ══════════════════════════════════════════
#  PDF GENERATION
# ══════════════════════════════════════════
def generate_pdf(cards_images, cards_per_page=4, page_size="A4"):
    if page_size == "A4":
        page_size = A4
    elif page_size == "A4-L":
        page_size = landscape(A4)
    else:
        page_size = A4

    pdf_buffer = io.BytesIO()
    c = canvas.Canvas(pdf_buffer, pagesize=page_size)
    page_width, page_height = page_size

    if cards_per_page == 1:
        card_width = page_width * 0.8
        card_height = page_height * 0.8
        margin_x = (page_width - card_width) / 2
        margin_y = (page_height - card_height) / 2
        positions = [(margin_x, margin_y, card_width, card_height)]
    elif cards_per_page == 2:
        card_width = page_width * 0.7
        card_height = page_height * 0.45
        margin_y = (page_height - card_height * 2) / 3
        positions = [
            ((page_width - card_width) / 2, margin_y, card_width, card_height),
            ((page_width - card_width) / 2, margin_y * 2 + card_height, card_width, card_height)
        ]
    else:
        card_width = page_width * 0.45
        card_height = page_height * 0.45
        margin_x = (page_width - card_width * 2) / 3
        margin_y = (page_height - card_height * 2) / 3
        positions = [
            (margin_x, margin_y, card_width, card_height),
            (margin_x * 2 + card_width, margin_y, card_width, card_height),
            (margin_x, margin_y * 2 + card_height, card_width, card_height),
            (margin_x * 2 + card_width, margin_y * 2 + card_height, card_width, card_height)
        ]

    for i, img in enumerate(cards_images):
        pos_index = i % cards_per_page

        if pos_index == 0 and i > 0:
            c.showPage()

        if pos_index < len(positions):
            x, y, w, h = positions[pos_index]
            img_buffer = io.BytesIO()
            img.save(img_buffer, format="PNG", dpi=(150, 150))
            img_buffer.seek(0)
            img_reader = ImageReader(img_buffer)
            c.drawImage(img_reader, x, y, width=w, height=h, preserveAspectRatio=True)

    c.save()
    pdf_buffer.seek(0)
    return pdf_buffer


# ══════════════════════════════════════════
#  ROUTES
# ══════════════════════════════════════════

@app.route('/')
def index():
    return send_file('index.html')


@app.route('/generate-pdf', methods=['POST'])
def generate_pdf_endpoint():
    try:
        data = request.get_json(force=True)
        print("Received PDF generation request")

        template_b64 = data.get("template", "")
        if not template_b64:
            return jsonify({"error": "لا يوجد قالب"}), 400

        if "," in template_b64:
            template_b64 = template_b64.split(",", 1)[1]

        template_bytes = base64.b64decode(template_b64)
        template_img = Image.open(io.BytesIO(template_bytes)).convert("RGB")

        fields = data.get("fields", [])
        if not fields:
            return jsonify({"error": "لا توجد حقول"}), 400

        rows = data.get("rows", [])
        if not rows:
            rows = [{}]

        cards_per_page = data.get("cards_per_page", 4)
        page_size = data.get("page_size", "A4")

        cards_images = []
        for i, row in enumerate(rows):
            card = template_img.copy()
            draw = ImageDraw.Draw(card)

            for field in fields:
                draw_field_on_image(draw, card, field, row)

            cards_images.append(card)

        pdf_buffer = generate_pdf(cards_images, cards_per_page, page_size)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"price_cards_{timestamp}.pdf"

        return send_file(
            pdf_buffer,
            mimetype='application/pdf',
            as_attachment=True,
            download_name=filename
        )

    except Exception as e:
        print(f"Error in generate-pdf: {e}")
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route('/generate', methods=['POST'])
def generate():
    try:
        data = request.get_json(force=True)

        template_b64 = data.get("template", "")
        if not template_b64:
            return jsonify({"error": "لا يوجد قالب"}), 400

        if "," in template_b64:
            template_b64 = template_b64.split(",", 1)[1]

        template_bytes = base64.b64decode(template_b64)
        template_img = Image.open(io.BytesIO(template_bytes)).convert("RGB")

        fields = data.get("fields", [])
        if not fields:
            return jsonify({"error": "لا توجد حقول"}), 400

        rows = data.get("rows", [])
        if not rows:
            rows = [{}]

        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
            for i, row in enumerate(rows):
                card = template_img.copy()
                draw = ImageDraw.Draw(card)

                for field in fields:
                    draw_field_on_image(draw, card, field, row)

                name_col = next(
                    (f["col"] for f in fields if f.get("type") in ("text", "price") and f.get("col")),
                    None
                )
                fname = safe_filename(
                    row.get("الموديل") or
                    row.get("اسم الصنف المعتمد") or
                    (row.get(name_col) if name_col else None) or
                    f"card_{i + 1}"
                )

                img_buf = io.BytesIO()
                card.save(img_buf, format="PNG", dpi=(300, 300))
                img_buf.seek(0)
                zf.writestr(f"{fname}.png", img_buf.read())

        zip_buf.seek(0)
        return send_file(
            zip_buf,
            mimetype='application/zip',
            as_attachment=True,
            download_name='price-cards.zip'
        )

    except Exception as e:
        print(f"Error in generate: {e}")
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route('/preview', methods=['POST'])
def preview():
    try:
        data = request.get_json(force=True)
        template_b64 = data.get("template", "")
        if "," in template_b64:
            template_b64 = template_b64.split(",", 1)[1]

        template_bytes = base64.b64decode(template_b64)
        template_img = Image.open(io.BytesIO(template_bytes)).convert("RGB")

        fields = data.get("fields", [])
        row = data.get("row", {})

        card = template_img.copy()
        draw = ImageDraw.Draw(card)
        for field in fields:
            draw_field_on_image(draw, card, field, row)

        buf = io.BytesIO()
        card.save(buf, format="PNG")
        buf.seek(0)
        img_b64 = base64.b64encode(buf.read()).decode()
        return jsonify({"image": f"data:image/png;base64,{img_b64}"})

    except Exception as e:
        print(f"Error in preview: {e}")
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


if __name__ == '__main__':
    print("=" * 50)
    print("  مصمم كروت الأسعار — السيرفر يعمل")
    print("  افتح المتصفح على: http://localhost:5000")
    print("=" * 50)
    app.run(debug=True, host='0.0.0.0', port=8001)
