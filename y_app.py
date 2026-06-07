import os
import re
import pandas as pd
from PIL import Image, ImageDraw, ImageFont  # ⭐ تأكد من ImageFont
from barcode import Code128
from barcode.writer import ImageWriter
import arabic_reshaper
from bidi.algorithm import get_display

# ================== المسارات ==================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

TEMPLATE = os.path.join(BASE_DIR, "كرت-سعر-اصفر.jpg")
EXCEL = os.path.join(BASE_DIR, "data2.xlsx")
OUT_DIR = os.path.join(BASE_DIR, "output")

FONT_NUMBERS = os.path.join(BASE_DIR, "sst-arabic-bold.ttf")
FONT_TEXT = os.path.join(BASE_DIR, "alfont_com_AlFont_com_SST-Arabic-Light-2 (1).ttf")

os.makedirs(OUT_DIR, exist_ok=True)

# ================== الخطوط بعد التصغير ==================
font_price_big = ImageFont.truetype(FONT_NUMBERS, 75)     # السعر الجديد
font_price_small = ImageFont.truetype(FONT_NUMBERS, 45)  # السعر القديم
font_label = ImageFont.truetype(FONT_TEXT, 28)
font_product_name = ImageFont.truetype(FONT_TEXT, 25)
font_model = ImageFont.truetype(FONT_TEXT, 18)

# ================== أدوات مساعدة ==================
def clean_price(val):
    try:
        val = str(val).replace(",", "").strip()
        if val == "" or val.lower() == "nan":
            return None
        return int(float(val))
    except:
        return None

def safe_filename(text):
    return re.sub(r'[\\/:*?"<>|]', '', str(text)).strip()

def draw_arabic_text(draw, text, x, y, font, fill="black", anchor="mm"):
    if not text:
        return
    reshaped = arabic_reshaper.reshape(str(text))
    bidi_text = get_display(reshaped)
    draw.text((x, y), bidi_text, fill=fill, font=font, anchor=anchor)

def create_clean_barcode(barcode_value, w=200, h=60):
    """
    إنشاء باركود نظيف وشفاف وحجمه أصغر
    بدون الرقم أسفل الباركود
    """
    writer = ImageWriter()
    writer.set_options({
        "write_text": False,  # منع الرقم أسفل الباركود
        "quiet_zone": 1,
        "module_width": 0.4,
        "module_height": 30,
        "background": "white"
    })

    temp = os.path.join(OUT_DIR, "temp_barcode")
    Code128(barcode_value, writer=writer).save(temp)

    # فتح الصورة وتحويلها لشفافة
    img = Image.open(temp + ".png").convert("RGBA")
    os.remove(temp + ".png")

    # إزالة أي نص أسفل الباركود (أي بكسلات بيضاء في الأسفل)
    # هذا يضمن عدم ظهور الرقم تحت الباركود
    datas = img.getdata()
    newData = []
    for item in datas:
        # كل الأبيض يتحول شفاف
        if item[0] > 240 and item[1] > 240 and item[2] > 240:
            newData.append((255, 255, 255, 0))
        else:
            newData.append(item)
    img.putdata(newData)

    # إعادة الحجم النهائي
    img = img.resize((w, h), Image.Resampling.LANCZOS)
    return img



# ================== تقسيم النصوص الطويلة ==================
def wrap_arabic_text(text, font, max_width):
    """تقسيم النص العربي إلى عدة أسطر حسب عرض معين"""
    words = text.split()
    lines = []
    line = ""
    dummy_img = Image.new("RGB", (1, 1))
    draw = ImageDraw.Draw(dummy_img)

    for word in words:
        test_line = f"{line} {word}".strip()
        bbox = draw.textbbox((0, 0), get_display(arabic_reshaper.reshape(test_line)), font=font)
        width = bbox[2] - bbox[0]
        if width <= max_width:
            line = test_line
        else:
            if line:
                lines.append(line)
            line = word
    if line:
        lines.append(line)
    return lines


# ================== قراءة الإكسل ==================
df = pd.read_excel(EXCEL)
df.columns = df.columns.str.strip()

# ================== التنفيذ ==================
for i, row in df.iterrows():
    price_after = clean_price(row.get("سعر البيع للفرع"))
    if price_after is None:
        continue

    price_before = clean_price(row.get("سعر قبل الخصم"))
    arabic_name = str(row.get("اسم الصنف المعتمد", "")).strip()
    model_name = str(row.get("الموديل", "")).strip()
    barcode_value = str(row.get("الباركود", "")).strip()

    img = Image.open(TEMPLATE).convert("RGB")
    draw = ImageDraw.Draw(img)

    WIDTH, HEIGHT = img.size
    CENTER_X = WIDTH // 2

    # ================== توزيع عمودي ==================
    y = 390  # بداية أقرب إلى الأعلى

    # ===== السعر الجديد =====
    draw.text(
        (CENTER_X, y),
        f"{price_after}",
        font=font_price_big,
        fill="black",
        anchor="mm"
    )
    draw_arabic_text(draw, "بعد", CENTER_X + 120, y + 5, font_label, fill="red")

    y += 65

    # ===== السعر القديم =====
    # ===== السعر القديم =====
    if price_before:
        draw.text(
            (CENTER_X, y),
            f"{price_before}",
            font=font_price_small,
            fill="red",
            anchor="mm"
        )
        # كلمة "قبل" واضحة ومستقيمة
        draw_arabic_text(draw, "قبل", CENTER_X + 90, y, font_label, fill="red", anchor="mm")

        # شطب السعر القديم بشكل مستقيم وواضح
        bbox = draw.textbbox((0, 0), f"{price_before}", font=font_price_small)
        left = CENTER_X - (bbox[2] - bbox[0]) // 2
        right = CENTER_X + (bbox[2] - bbox[0]) // 2
        line_y = y  # نفس مستوى السعر
        draw.line(
            (left, line_y, right, line_y),
            fill="black",
            width=3  # زيادة عرض الخط ليكون واضح
        )
        y += 70

    # ===== اسم المنتج (مع تقسيم الأسطر) =====
    if arabic_name:
        max_width = WIDTH - 40
        lines = wrap_arabic_text(arabic_name, font_product_name, max_width)
        for l in lines:
            draw_arabic_text(draw, l, CENTER_X, y - 20, font_product_name)
            y += 40  # المسافة بين الأسطر

    # ===== الموديل (مع تقسيم إذا كان طويل) =====
    if model_name:
        max_width = WIDTH - 40
        lines = wrap_arabic_text(f"الموديل: {model_name}", font_model, max_width)
        for l in lines:
            draw_arabic_text(draw, l, CENTER_X, y, font_model)
            y += 80
    # ===== الباركود =====
    if barcode_value and barcode_value != "nan":
        barcode_img = create_clean_barcode(barcode_value)
        img.paste(barcode_img, (CENTER_X - barcode_img.width // 2, y), barcode_img)

    # ===== حفظ =====
    filename = safe_filename(model_name or arabic_name or f"product_{i}")
    img.save(os.path.join(OUT_DIR, f"{filename}.png"), dpi=(300, 300))
    print(f"✅ تم إنشاء {filename}.png")

print("🎉 تم إنشاء جميع كروت الأسعار بنجاح")
