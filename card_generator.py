import os
import re
import io
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
from barcode import Code128
from barcode.writer import ImageWriter
import arabic_reshaper
from bidi.algorithm import get_display

# ═══════════════════════════════════════════
#  الإعدادات - نفس موقع index.html
# ═══════════════════════════════════════════
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ⭐ نفس إعدادات index.html بالضبط ⭐
# القالب الافتراضي 600x720
TEMPLATE_WIDTH = 600
TEMPLATE_HEIGHT = 720

# مجلدات
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")
OUT_DIR = os.path.join(BASE_DIR, "output")
os.makedirs(TEMPLATES_DIR, exist_ok=True)
os.makedirs(OUT_DIR, exist_ok=True)

# ملف الإكسل
EXCEL_FILE = os.path.join(BASE_DIR, "book6.xlsx")

# ملفات الخطوط
FONT_BOLD = os.path.join(BASE_DIR, "sst-arabic-bold.ttf")
FONT_LIGHT = os.path.join(BASE_DIR, "alfont_com_AlFont_com_SST-Arabic-Light-2 (1).ttf")

def load_font(path, size):
    """تحميل الخط"""
    try:
        return ImageFont.truetype(path, size)
    except:
        return ImageFont.load_default()

def arabic(text):
    """تشكيل النص العربي"""
    try:
        return get_display(arabic_reshaper.reshape(str(text)))
    except:
        return str(text)

# ═══════════════════════════════════════════
#  نفس حقول index.html بالضبط
# ═══════════════════════════════════════════
FIELDS = [
    # السعر الجديد (type: price)
    {"name": "السعر الجديد", "col": "السعر بعد الخصم", "type": "price",
     "x": 190, "y": 295, "w": 220, "h": 100, "size": 78, "color": "#000000", "bold": True},

    # السعر القديم (type: price_strike)
    {"name": "السعر القديم", "col": "السعر قبل الخصم", "type": "price_strike",
     "x": 195, "y": 415, "w": 210, "h": 65, "size": 46, "color": "#cc0000", "bold": True},

    # اسم المنتج
    {"name": "اسم المنتج", "col": "اسم الصنف", "type": "text",
     "x": 40, "y": 500, "w": 520, "h": 55, "size": 28, "color": "#111111", "bold": False},

    # الموديل
    {"name": "الموديل", "col": "الموديل", "type": "text",
     "x": 80, "y": 565, "w": 440, "h": 38, "size": 20, "color": "#333333", "bold": False},

    # الباركود
    {"name": "الباركود", "col": "Barcode", "type": "barcode",
     "x": 170, "y": 615, "w": 260, "h": 55, "size": 0, "color": "#000000", "bold": False},
    
    # نص "قبل الخصم" (static_label)
    {"name": "نص قبل الخصم", "col": "", "type": "static_label",
     "x": 150, "y": 385, "w": 100, "h": 35, "size": 24, "color": "#ff0000", "bold": True, "staticVal": "قبل الخصم"},
    
    # نص "السعر الجديد" (static_label)
    {"name": "نص السعر الجديد", "col": "", "type": "static_label",
     "x": 150, "y": 265, "w": 100, "h": 35, "size": 24, "color": "#00aa00", "bold": True, "staticVal": "السعر الجديد"},
]

# ═══════════════════════════════════════════
#  أدوات مساعدة
# ═══════════════════════════════════════════
def clean_price(val):
    """تنظيف السعر"""
    try:
        val = str(val).replace(",", "").strip()
        if val == "" or val.lower() == "nan":
            return None
        result = round(float(val))
        return result
    except Exception as e:
        print(f"      ❌ خطأ في تنظيف السعر '{val}': {e}")
        return None

def calculate_discount(price_before, price_after):
    """حساب نسبة الخصم"""
    try:
        if price_before and price_after and price_before > price_after:
            return int(((price_before - price_after) / price_before) * 100)
        return None
    except:
        return None

def safe_filename(text):
    """اسم ملف آمن"""
    return re.sub(r'[\\/:*?"<>|]', '', str(text)).strip() or "product"

# ═══════════════════════════════════════════
#  باركود
# ═══════════════════════════════════════════
def create_barcode(value, w, h):
    """إنشاء باركود Code128"""
    try:
        value_str = str(value).strip()
        if not value_str or value_str == "nan":
            return None
        
        writer = ImageWriter()
        writer.set_options({
            "write_text": False,
            "quiet_zone": 1,
            "module_width": 0.45,
            "module_height": 6,
            "foreground": "black",
            "background": "white"
        })
        
        buf = io.BytesIO()
        Code128(value_str, writer=writer).write(buf)
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
        
        return img.resize((w, h), Image.Resampling.LANCZOS)
    except Exception as e:
        print(f"Barcode error: {e}")
        return None

def wrap_text_lines(text, font, max_width, draw):
    """تقسيم النص لأسطر"""
    if not text:
        return []
    
    bbox = draw.textbbox((0, 0), text, font=font)
    if (bbox[2] - bbox[0]) <= max_width:
        return [text]
    
    words = text.split()
    lines = []
    current = ""
    
    for word in words:
        test = f"{current} {word}".strip() if current else word
        bbox = draw.textbbox((0, 0), test, font=font)
        if (bbox[2] - bbox[0]) <= max_width:
            current = test
        else:
            if current:
                lines.append(current)
            current = word
        if len(lines) == 2:
            break
    
    if current and len(lines) < 2:
        lines.append(current)
    
    return lines if lines else [text]

# ═══════════════════════════════════════════
#  الرسم على الصورة - نفس index.html
# ═══════════════════════════════════════════
def draw_field(draw, img, field, row_data):
    """رسم حقل واحد - يتجاهل الحقول الناقصة"""
    try:
        ftype = field.get("type", "text")
        col = field.get("col", "")
        x = field.get("x", 0)
        y = field.get("y", 0)
        w = field.get("w", 200)
        h = field.get("h", 50)
        size = field.get("size", 28)
        color = field.get("color", "#000000")
        bold = field.get("bold", False)
        static = field.get("staticVal", "")

        cx = x + w // 2
        cy = y + h // 2

        # ── الباركود ──
        if ftype == "barcode":
            if not col or col not in row_data:
                return  # تجاهل إذا العمود غير موجود
            val = row_data.get(col, "")
            if val and str(val).strip() and str(val).lower() != "nan":
                bc_img = create_barcode(val, w, h)
                if bc_img:
                    img.paste(bc_img, (x, y), bc_img)
            return

        # ── النص الثابت (static_label) ──
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
        if ftype in ["price", "price_strike", "text"]:
            # تجاهل إذا العمود غير موجود
            if not col:
                return
            if col not in row_data:
                print(f"    ⚠️ العمود '{col}' غير موجود في الإكسل")
                return

            raw = row_data.get(col, "")
            if raw is None or str(raw).strip() == "" or str(raw).lower() == "nan":
                print(f"    ⚠️ العمود '{col}' فارغ أو nan: '{raw}'")
                return

            # تنسيق السعر
            if ftype in ["price", "price_strike"]:
                price_val = clean_price(raw)
                if price_val is None:
                    return
                raw_text = f"{price_val:,}"
            else:
                raw_text = str(raw)

            font_path = FONT_BOLD if bold else FONT_LIGHT
            font = load_font(font_path, size)
            text = arabic(raw_text)

            # تعديل المحاذاة حسب النوع
            if ftype == "price":
                # محاذاة يمين للسعر
                bbox = draw.textbbox((0, 0), text, font=font)
                tw = bbox[2] - bbox[0]
                lx = x + w - tw - 10  # 10px margin
                ly = cy - (bbox[3] - bbox[1]) // 2
                draw.text((lx, ly), text, fill=color, font=font)
                print(f"    ✓ تم رسم السعر: {raw_text}")

            elif ftype == "price_strike":
                # محاذاة يمين + خط شطب
                bbox = draw.textbbox((0, 0), text, font=font)
                tw = bbox[2] - bbox[0]
                lx = x + w - tw - 10
                ly = cy - (bbox[3] - bbox[1]) // 2
                draw.text((lx, ly), text, fill=color, font=font)
                # خط الشطب
                line_y = ly + (bbox[3] - bbox[1]) // 2
                draw.line([(lx, line_y), (lx + tw, line_y)], fill=color, width=max(2, size // 14))
                print(f"    ✓ تم رسم السعر القديم: {raw_text}")

            else:
                # نص عادي - توسيط
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

    except Exception as e:
        # تجاهل الأخطاء الصغيرة
        pass

# ═══════════════════════════════════════════
#  معالجة منتج واحد
# ═══════════════════════════════════════════
def process_product(template_path, row, output_dir, index=0):
    """معالجة منتج واحد على القالب"""
    try:
        # Debug: عرض الأعمدة المتاحة
        if index == 0:
            print(f"\n📋 الأعمدة في الإكسل: {list(row.keys())}")
        
        # قراءة بيانات المنتج
        product_name = row.get("اسم الصنف", "")
        price_before_raw = row.get("السعر قبل الخصم", "")
        price_after_raw = row.get("السعر بعد الخصم", "")
        discount_val_raw = row.get("نسبة الخصم", "")
        
        price_before = clean_price(price_before_raw)
        price_after = clean_price(price_after_raw)
        discount_val = clean_price(discount_val_raw)
        
        model_name = str(row.get("الموديل", "")).strip()
        barcode_value = str(row.get("Barcode", "")).strip()
        
        # Debug للمنتج الأول
        if index == 0:
            print(f"📝 منتج: {product_name}")
            print(f"💰 سعر قبل (raw): '{price_before_raw}' → clean: {price_before}")
            print(f"💰 سعر بعد (raw): '{price_after_raw}' → clean: {price_after}")
        
        # معالجة السعر
        if price_after is None and price_before is not None:
            price_after = price_before
        if price_after is None:
            price_after = 0
        
        # تحميل القالب أو إنشاء قالب افتراضي
        if os.path.exists(template_path):
            img = Image.open(template_path).convert("RGB")
        else:
            # إنشاء قالب افتراضي 600x720 نفس index.html
            img = Image.new('RGB', (TEMPLATE_WIDTH, TEMPLATE_HEIGHT), '#f5f5f5')
            draw_temp = ImageDraw.Draw(img)
            # خلفية بيضاء في الأسفل
            draw_temp.rectangle([0, 500, TEMPLATE_WIDTH, TEMPLATE_HEIGHT], fill='white')
            # خط أحمر في الأسفل
            draw_temp.rectangle([0, TEMPLATE_HEIGHT-70, TEMPLATE_WIDTH, TEMPLATE_HEIGHT], fill='#c50000')
            draw_temp.text((TEMPLATE_WIDTH//2, TEMPLATE_HEIGHT-35), arabic("ريال سعودي"), 
                          fill='#ffd700', font=load_font(FONT_BOLD, 22), anchor="mm")
            # إطار ذهبي
            draw_temp.rectangle([12, 12, TEMPLATE_WIDTH-12, TEMPLATE_HEIGHT-12], outline='#ffd700', width=3)
        
        draw = ImageDraw.Draw(img)
        
        # رسم جميع الحقول
        for field in FIELDS:
            draw_field(draw, img, field, row)
        
        # اسم الملف
        if model_name and model_name != "nan":
            filename = safe_filename(f"{model_name}")
        elif product_name and product_name != "nan":
            filename = safe_filename(f"{product_name}")
        else:
            filename = f"product"
        
        output_path = os.path.join(output_dir, f"{filename}.png")
        img.save(output_path, dpi=(300, 300), quality=95)
        
        return filename
        
    except Exception as e:
        print(f"Error processing product: {e}")
        import traceback
        traceback.print_exc()
        return None

# ═══════════════════════════════════════════
#  البرنامج الرئيسي
# ═══════════════════════════════════════════
def main():
    print("\n" + "=" * 60)
    print("🎨 نظام توليد كروت الأسعار - نفس إعدادات الموقع")
    print("=" * 60)
    
    # البحث عن القوالب
    templates = [f for f in os.listdir(TEMPLATES_DIR) 
                 if f.lower().endswith(('.jpg', '.jpeg', '.png'))] if os.path.exists(TEMPLATES_DIR) else []
    
    # قراءة الإكسل
    try:
        df = pd.read_excel(EXCEL_FILE)
        df.columns = df.columns.str.strip()
        print(f"📊 تم قراءة {len(df)} منتج من الإكسل")
    except Exception as e:
        print(f"❌ خطأ في قراءة الإكسل: {e}")
        return
    
    total_images = 0
    
    # إذا وجدنا قوالب
    if templates:
        print(f"\n✅ تم العثور على {len(templates)} قالب:")
        for i, t in enumerate(templates, 1):
            print(f"   {i}. {t}")
        
        for template_file in templates:
            template_path = os.path.join(TEMPLATES_DIR, template_file)
            template_name = os.path.splitext(template_file)[0]
            
            # مجلد فرعي لكل قالب
            output_subdir = os.path.join(OUT_DIR, template_name)
            os.makedirs(output_subdir, exist_ok=True)
            
            print(f"\n{'─' * 60}")
            print(f"🎨 معالجة القالب: {template_file}")
            print(f"{'─' * 60}")
            
            for i, row in df.iterrows():
                try:
                    filename = process_product(template_path, row, output_subdir, i)
                    if filename:
                        print(f"  ✅ [{i + 1}/{len(df)}] تم حفظ: {filename}.png")
                        total_images += 1
                    else:
                        print(f"  ⚠️ [{i + 1}/{len(df)}] تم تخطي المنتج")
                except Exception as e:
                    print(f"  ❌ [{i + 1}/{len(df)}] خطأ: {e}")
    else:
        # استخدام القالب الافتراضي
        print("\n⚠️ لم يوجد قوالب، سيتم استخدام القالب الافتراضي")
        print(f"{'─' * 60}")
        
        output_subdir = os.path.join(OUT_DIR, "default")
        os.makedirs(output_subdir, exist_ok=True)
        
        for i, row in df.iterrows():
            try:
                filename = process_product(None, row, output_subdir, i)
                if filename:
                    print(f"  ✅ [{i + 1}/{len(df)}] تم حفظ: {filename}.png")
                    total_images += 1
            except Exception as e:
                print(f"  ❌ [{i + 1}/{len(df)}] خطأ: {e}")
    
    # الملخص
    print("\n" + "=" * 60)
    print("✅ اكتمل التوليد!")
    print("=" * 60)
    print(f"📊 إجمالي الصور: {total_images}")
    print(f"📁 مجلد الإخراج: {OUT_DIR}")
    print("=" * 60 + "\n")

if __name__ == "__main__":
    main()
