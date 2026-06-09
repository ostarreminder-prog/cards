import os
import re
import json
import base64
import io
import zipfile
import hashlib
import secrets
import pickle
import numpy as np
from datetime import datetime
from functools import wraps
from urllib.parse import urlencode

from flask import Flask, request, jsonify, send_file, session, redirect, render_template_string, url_for
from flask_cors import CORS

# Load environment variables from .env file
from dotenv import load_dotenv
load_dotenv()

# Google APIs
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import Flow
from google.oauth2 import service_account
from googleapiclient.discovery import build
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
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE = os.path.join(BASE_DIR, 'products.xlsx')

EXCEL_COLUMN_MAP = {
    'اسم الصنف المعتمد': 'اسم الصنف',
    'الباركود': 'Barcode',
    'الماركة': 'Brand',
    'سعر البيع للفرع': 'السعر بعد الخصم',
    'السعر بعد الخصم': 'السعر بعد الخصم',
    'سعر قبل الخصم': 'السعر قبل الخصم',
}

def clean_col(name):
    """تنظيف اسم العمود من المسافات الغريبة"""
    return str(name).replace('\xa0', ' ').replace('\u200b', '').strip()

def delete_excel_file():
    """حذف ملف الإكسل المرفوع وبياناته الوصفية بعد انتهاء العملية"""
    try:
        if os.path.exists(EXCEL_FILE):
            os.remove(EXCEL_FILE)
        meta_file = os.path.join(BASE_DIR, 'excel_meta.json')
        if os.path.exists(meta_file):
            os.remove(meta_file)
        print("[EXCEL] 🗑️ تم حذف ملف الإكسل بعد انتهاء العملية")
    except Exception as e:
        print(f"[EXCEL] ✗ خطأ في حذف الملف: {e}")

def read_excel_products():
    """قراءة المنتجات من ملف Excel المرفوع مع تحويل أسماء الأعمدة"""
    if not os.path.exists(EXCEL_FILE):
        return []
    try:
        df = pd.read_excel(EXCEL_FILE, dtype=str)
        df = df.fillna('')
        # تنظيف أسماء الأعمدة من المسافات الغريبة (xa0 وغيرها)
        df.columns = [clean_col(c) for c in df.columns]
        print(f"[EXCEL] أعمدة بعد التنظيف: {list(df.columns)}")
        # تطبيق mapping
        df.rename(columns={k: v for k, v in EXCEL_COLUMN_MAP.items() if k in df.columns}, inplace=True)
        rows = []
        for _, row in df.iterrows():
            row_dict = {str(col).strip(): str(val).strip() for col, val in row.items()}
            if not row_dict.get('اسم الصنف', '').strip():
                continue
            rows.append(row_dict)
        print(f"[EXCEL] ✓ قُرئ الملف: {len(rows)} منتج")
        return rows
    except Exception as e:
        print(f"[EXCEL] ✗ خطأ في قراءة الملف: {e}")
        return []

# ================== PRODUCT TRACKER SYSTEM ==================
PRODUCTS_DB_FILE = os.path.join(BASE_DIR, 'products_tracker.json')

def load_products_tracker():
    """تحميل قاعدة بيانات المنتجات"""
    if os.path.exists(PRODUCTS_DB_FILE):
        try:
            with open(PRODUCTS_DB_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return {}
    return {}

def save_products_tracker(data):
    """حفظ قاعدة بيانات المنتجات"""
    with open(PRODUCTS_DB_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def get_product_hash(product):
    """إنشاء معرف فريد للمنتج بناءً على Barcode أو الموديل + اسم الصنف"""
    barcode = str(product.get('Barcode', '')).strip()
    if barcode and barcode != 'nan':
        return barcode
    
    model = str(product.get('الموديل', '')).strip()
    name = str(product.get('اسم الصنف', '')).strip()
    return f"{model}_{name}" if model or name else None

def check_product_changes(product, tracked_data):
    """فحص إذا تغيرت بيانات المنتج في أي عمود"""
    if not tracked_data:
        return None, False, None
    
    changes = []
    has_changes = False
    product_name = str(product.get('اسم الصنف', ''))[0:20]
    
    # قائمة جميع الأعمدة للمقارنة
    columns_to_compare = [
        ('السعر بعد الخصم', ['السعر بعد الخصم ', 'السعر بعد الخصم']),
        ('السعر قبل الخصم', ['السعر قبل الخصم']),
        ('اسم الصنف', ['اسم الصنف']),
        ('الموديل', ['الموديل']),
        ('Brand', ['Brand']),
        ('Barcode', ['Barcode']),
        ('نسبة الخصم', ['نسبة الخصم']),
    ]
    
    for tracked_key, product_keys in columns_to_compare:
        old_val = str(tracked_data.get(tracked_key, '')).strip()
        
        # نجرب جميع مفاتيح الممكنة
        new_val = ''
        for key in product_keys:
            new_val = str(product.get(key, '')).strip()
            if new_val:
                break
        
        # شرط التغيير: نقارن حتى لو old_val فارغة، فقط نتخطى لو الاثنين فارغين
        if old_val != new_val and not (old_val == '' and new_val == ''):
            changes.append(f"{tracked_key}: {old_val} → {new_val}")
            has_changes = True
            if product_name:
                print(f"[CHANGED] ⭐ {product_name}: {tracked_key} تغير! '{old_val}' → '{new_val}'")
    
    last_updated = tracked_data.get('last_updated', '')
    
    return changes, has_changes, last_updated

def get_sheet_last_modified(sheets_service, spreadsheet_id):
    """جلب تاريخ آخر تعديل على الشيت باستخدام Google Drive API"""
    try:
        # ننشئ خدمة Drive باستخدام نفس الـ credentials
        from googleapiclient.discovery import build
        
        # نجلب الـ credentials من الـ sheets service
        creds = sheets_service._http.credentials
        
        # ننشئ خدمة Drive
        drive_service = build('drive', 'v3', credentials=creds, cache_discovery=False)
        
        # نجلب معلومات الملف
        file_info = drive_service.files().get(
            fileId=spreadsheet_id,
            fields='modifiedTime'
        ).execute()
        
        # نحول الـ ISO format لـ readable format
        modified_time = file_info.get('modifiedTime', '')
        if modified_time:
            from datetime import datetime as dt
            # 2025-06-08T10:30:00.000Z -> 2025-06-08 13:30 (بالتوقيت المحلي)
            parsed = dt.fromisoformat(modified_time.replace('Z', '+00:00'))
            # نحول للتوقيت المحلي (UTC+3 للسعودية)
            local_time = parsed.astimezone()
            return local_time.strftime('%Y-%m-%d %H:%M')
        
        return datetime.now().strftime('%Y-%m-%d %H:%M')
    except Exception as e:
        print(f"⚠️ خطأ في جلب تاريخ التعديل: {e}")
        return datetime.now().strftime('%Y-%m-%d %H:%M')

def process_fetched_products(products, sheet_modified_time=None):
    """
    يقارن البيانات الحالية من الشيت مع آخر نسخة محفوظة (من آخر جلسة).
    - إذا تغير شيء → يعرضه كـ 'updated' تلقائياً
    - يحفظ النسخة الحالية بعد المقارنة (للجلسة القادمة)
    - لا يحتاج أي تدخل يدوي
    """
    tracker = load_products_tracker()
    processed = []
    current_time = sheet_modified_time or datetime.now().strftime('%Y-%m-%d %H:%M')
    updated_count = 0
    new_tracker = {}  # نسخة جديدة تُحفظ بعد المقارنة

    for product in products:
        product_id = get_product_hash(product)
        if not product_id:
            processed.append({**product, '_status': 'new', '_hidden': False, '_last_updated': current_time})
            continue

        tracked = tracker.get(product_id)

        # القيم الحالية من الشيت
        current_vals = {
            'اسم الصنف': str(product.get('اسم الصنف', '')).strip(),
            'السعر بعد الخصم': str(product.get('السعر بعد الخصم ', product.get('السعر بعد الخصم', ''))).strip(),
            'السعر قبل الخصم': str(product.get('السعر قبل الخصم', '')).strip(),
            'الموديل': str(product.get('الموديل', '')).strip(),
            'Brand': str(product.get('Brand', '')).strip(),
            'Barcode': str(product.get('Barcode', '')).strip(),
            'نسبة الخصم': str(product.get('نسبة الخصم', '')).strip(),
        }

        if tracked:
            # قارن مع آخر نسخة محفوظة
            old_price = tracked.get('السعر بعد الخصم', '')
            new_price = current_vals['السعر بعد الخصم']
            print(f"[CMP] {current_vals['اسم الصنف'][0:20]}: tracker={old_price!r} | sheet={new_price!r} | match={old_price==new_price}")
            changes, has_changes, _ = check_product_changes(product, tracked)
            hidden = tracked.get('hidden', False)

            if has_changes:
                updated_count += 1
                print(f"[DEBUG] ✅ محدث: {current_vals['اسم الصنف'][0:25]} | {changes}")
                processed.append({
                    **product,
                    '_status': 'updated',
                    '_changes': changes,
                    '_hidden': False,
                    '_last_updated': current_time
                })
            else:
                processed.append({
                    **product,
                    '_status': 'same',
                    '_hidden': hidden,
                    '_last_updated': tracked.get('last_updated', current_time)
                })
        else:
            # منتج جديد (أول مرة نشوفه)
            processed.append({**product, '_status': 'new', '_hidden': False, '_last_updated': current_time})

        if not tracked:
            # منتج جديد فقط: احفظه كنقطة مرجعية
            new_tracker[product_id] = {
                **current_vals,
                'hidden': False,
                'last_updated': current_time,
            }
        else:
            # منتج موجود: احتفظ بالقيم القديمة كما هي (لا تحدّثها)
            # ستتحدث فقط عند الطباعة
            new_tracker[product_id] = {
                **tracker[product_id],  # ← القيم القديمة تبقى
            }

    print(f"[DEBUG] {updated_count} محدث | {len([p for p in processed if p['_status']=='new'])} جديد | {len([p for p in processed if p['_status']=='same'])} بدون تغيير")
    save_products_tracker(new_tracker)
    return processed

# ================== GOOGLE SERVICE ACCOUNT CONFIG ==================
GOOGLE_CLIENT_ID = os.environ.get('GOOGLE_CLIENT_ID', '')
GOOGLE_CLIENT_SECRET = os.environ.get('GOOGLE_CLIENT_SECRET', '')
SCOPES = [
    'https://www.googleapis.com/auth/userinfo.email',
    'https://www.googleapis.com/auth/userinfo.profile',
    'https://www.googleapis.com/auth/spreadsheets.readonly',
    'openid'
]
ALLOWED_EMAILS = []  # سيملأ من البيئة أو يُترك فارغ للسماح للجميع

# Service Account Configuration
SERVICE_ACCOUNT_FILE = os.path.join(BASE_DIR, 'service_account.json')
SHEETS_SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets.readonly',
    'https://www.googleapis.com/auth/drive.metadata.readonly'  # لجلب تاريخ التعديل
]

# ================== TEMPLATES CONFIG ==================
TEMPLATES = {
    'offers': {
        'name': 'عروض',
        'file': 'rrsa.jpg',
        'type': 'offers'
    },
    'products': {
        'name': 'منتجات',
        'file': 'template.jpg',
        'type': 'products'
    },
    'products_other': {
        'name': 'منتجات أخرى',
        'file': 'LLL.jpg',
        'type': 'products_other'
    }
}

# Google Sheet ID (من الرابط) - افتراضي للنظام التلقائي
SHEET_ID = os.environ.get('GOOGLE_SHEET_ID', '')

# ID الشيت الافتراضي للموظفين (ثابت)
DEFAULT_SHEET_ID = os.environ.get('DEFAULT_SHEET_ID', '')

# شيتات منفصلة لكل نوع
SHEET_IDS = {
    'offers': os.environ.get('SHEET_OFFERS_ID', ''),
    'products': os.environ.get('SHEET_PRODUCTS_ID', ''),
    'products_other': os.environ.get('SHEET_PRODUCTS_OTHER_ID', '')
}

app = Flask(__name__, static_folder='.', static_url_path='')
app.secret_key = os.environ.get('SECRET_KEY', 'dev-secret-key-change-in-production')
CORS(app)

# ══════════════════════════════════════════
#  AUTHENTICATION SYSTEM
# ══════════════════════════════════════════
USERS_FILE = os.path.join(BASE_DIR, 'users.json')

DEFAULT_ADMIN = {
    "admin": {
        "password_hash": hashlib.sha256("123456".encode()).hexdigest(),
        "role": "admin",
        "created_at": datetime.now().isoformat()
    }
}

# ================== GOOGLE AUTH FUNCTIONS ==================
def get_google_auth_url():
    if not GOOGLE_CLIENT_ID:
        return None
    
    flow = Flow.from_client_config(
        {
            "web": {
                "client_id": GOOGLE_CLIENT_ID,
                "client_secret": GOOGLE_CLIENT_SECRET,
                "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                "token_uri": "https://oauth2.googleapis.com/token",
                "redirect_uris": ["http://localhost:8001/oauth/callback"]
            }
        },
        scopes=SCOPES,
        redirect_uri="http://localhost:8001/oauth/callback"
    )
    
    auth_url, state = flow.authorization_url(
        access_type='offline',
        include_granted_scopes='true',
        prompt='consent'
    )
    
    return auth_url, state

def fetch_sheet_data(service, spreadsheet_id, range_name='A1:Z200'):
    try:
        print(f"[FETCH] جاري الجلب من Google Sheets API... ID={spreadsheet_id[:20]}")
        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=range_name
        ).execute()
        
        values = result.get('values', [])
        print(f"[FETCH] ✓ وصلت {len(values)} صفوف من الـ API")
        if not values:
            return []
        
        # تحويل لـ DataFrame
        headers = values[0]
        data = values[1:]
        
        rows = []
        for row in data:
            # تخطي الصفوف الفارغة
            if not row or all(str(cell).strip() == '' for cell in row):
                continue
            
            row_dict = {}
            for i, header in enumerate(headers):
                row_dict[header] = row[i] if i < len(row) else ''
            
            # تخطي الصفوف التي لا تحتوي على اسم صنف
            if not row_dict.get('اسم الصنف', '').strip():
                continue
            
            rows.append(row_dict)
        
        return rows
    except Exception as e:
        print(f"Error fetching sheet: {e}")
        return []

def create_sheets_service(credentials_dict=None):
    """إنشاء خدمة Sheets باستخدام Service Account أو OAuth credentials"""
    if credentials_dict:
        # OAuth mode (legacy)
        creds = Credentials.from_authorized_user_info(credentials_dict, SCOPES)
        if creds.expired and creds.refresh_token:
            creds.refresh(Request())
        return build('sheets', 'v4', credentials=creds, cache_discovery=False)
    
    # Service Account mode (default)
    if os.path.exists(SERVICE_ACCOUNT_FILE):
        creds = service_account.Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE,
            scopes=SHEETS_SCOPES
        )
        return build('sheets', 'v4', credentials=creds, cache_discovery=False)
    
    return None

def load_users():
    if os.path.exists(USERS_FILE):
        with open(USERS_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return DEFAULT_ADMIN.copy()

def save_users(users):
    with open(USERS_FILE, 'w', encoding='utf-8') as f:
        json.dump(users, f, indent=2, ensure_ascii=False)

def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

def require_auth(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user' not in session:
            return redirect('/login')
        return f(*args, **kwargs)
    return decorated

def require_admin(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user' not in session:
            return redirect('/login')
        users = load_users()
        if users.get(session['user'], {}).get('role') != 'admin':
            return jsonify({"error": "Unauthorized"}), 403
        return f(*args, **kwargs)
    return decorated

# مسارات الخطوط
FONT_BOLD = os.path.join(BASE_DIR, "sst-arabic-bold.ttf")
FONT_LIGHT = os.path.join(BASE_DIR, "alfont_com_AlFont_com_SST-Arabic-Light-2 (1).ttf")


def load_font(path, size):
    try:
        return ImageFont.truetype(path, size)
    except Exception:
        return ImageFont.load_default()


# هل Pillow مبني مع raqm؟ إذا نعم فهو يعالج العربية تلقائياً
try:
    from PIL import features as _pil_features
    RAQM_AVAILABLE = _pil_features.check('raqm')
except Exception:
    RAQM_AVAILABLE = False


def arabic(text):
    """تشكيل النص العربي بشكل صحيح.
    إذا كان Pillow يدعم raqm فهو يتولى الـ shaping + bidi تلقائياً،
    لذا نمرّر النص الخام لتفادي المعالجة المزدوجة (النص المقلوب).
    """
    if RAQM_AVAILABLE:
        return str(text)
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

        # ── معالجة الأسعار ──
        if ftype in ["price", "price_strike"]:
            price_val = clean_price(raw_text)
            if price_val is None:
                return
            raw_text = f"{price_val:,}"
            
            font_path = FONT_BOLD if bold else FONT_LIGHT
            font = load_font(font_path, size)
            text = arabic(raw_text)
            
            # محاذاة يمين للأسعار
            bbox = draw.textbbox((0, 0), text, font=font)
            tw = bbox[2] - bbox[0]
            th = bbox[3] - bbox[1]
            lx = x + w - tw - 10  # 10px margin من اليمين
            ly = cy - th // 2
            draw.text((lx, ly), text, fill=color, font=font)
            
            # خط الشطب للسعر القديم
            if ftype == "price_strike":
                line_y = ly + th // 2
                draw.line(
                    [(lx, line_y), (lx + tw, line_y)],
                    fill=color,
                    width=max(2, size // 14)
                )
            return

        # ── النص العادي ──
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
    except Exception as e:
        print(f"Error drawing field {field.get('name')}: {e}")
        traceback.print_exc()


# ══════════════════════════════════════════
#  CARD PROCESSING (from user's code)
# ══════════════════════════════════════════

# الخطوط
font_price_big = ImageFont.truetype(FONT_BOLD, 165)
font_price_small = ImageFont.truetype(FONT_BOLD, 50)
font_label = ImageFont.truetype(FONT_LIGHT, 30)
font_model = ImageFont.truetype(FONT_LIGHT, 18)
font_discount = ImageFont.truetype(FONT_BOLD, 35)

def clean_price(val):
    try:
        val = str(val).replace(",", "").strip()
        if val == "" or val.lower() == "nan":
            return None
        return round(float(val))
    except:
        return None

def calculate_discount(price_before, price_after):
    try:
        if price_before and price_after and price_before > price_after:
            return int(((price_before - price_after) / price_before) * 100)
        return None
    except:
        return None

def create_clean_barcode(barcode_value, target_width=150, target_height=35):
    writer = ImageWriter()
    writer.set_options({
        "write_text": False,
        "quiet_zone": 1,
        "module_width": 0.4,
        "module_height": 6,
        "foreground": "black",
        "background": "white",
    })
    
    buf = io.BytesIO()
    Code128(str(barcode_value), writer=writer).write(buf)
    buf.seek(0)
    img = Image.open(buf).convert("RGBA")
    
    bbox = img.getbbox()
    if bbox:
        img = img.crop(bbox)
    
    img = img.resize((target_width, target_height), Image.Resampling.LANCZOS)
    return img

def draw_product_name_safe(draw, center_x, start_y, text, fill, max_width):
    if not text or str(text).strip() == "":
        return
    
    raw_text = str(text)
    for size in [38, 36, 34, 32, 30]:
        try:
            font = ImageFont.truetype(FONT_LIGHT, size)
        except:
            font = ImageFont.load_default()
        
        words = raw_text.split(" ")
        lines = []
        current = ""
        
        for w in words:
            test = current + " " + w if current else w
            bbox = draw.textbbox((0, 0), test, font=font)
            w_width = bbox[2] - bbox[0]
            if w_width <= max_width:
                current = test
            else:
                lines.append(current)
                current = w
            if len(lines) == 2:
                break
        
        if current and len(lines) < 2:
            lines.append(current)
        
        if lines:
            y = start_y
            for line in lines:
                bidi_line = arabic(line)
                draw.text((center_x, y), bidi_line, fill=fill, font=font, anchor="mm")
                y += size + 6
            return

def process_product_card(template_path, row, base_riyal_img=None):
    """معالجة منتج واحد على قالب واحد — أي عمود مفقود يُتجاهل"""
    try:
        return _process_product_card_inner(template_path, row, base_riyal_img)
    except Exception as e:
        print(f"⚠️ خطأ في معالجة المنتج: {e} | المنتج: {row.get('اسم الصنف','?')}")
        try:
            img = Image.open(template_path).convert("RGB")
            buf = io.BytesIO()
            img.save(buf, format="PNG")
            buf.seek(0)
            return buf, "error_card"
        except:
            return None, "error_card"

def _process_product_card_inner(template_path, row, base_riyal_img=None):
    """المعالجة الفعلية للكرت"""
    
    product_name = row.get("اسم الصنف", "")
    price_before = clean_price(row.get("السعر قبل الخصم", ""))
    price_after = clean_price(row.get("السعر بعد الخصم", "") or row.get("السعر بعد الخصم ", ""))
    discount_raw = str(row.get("نسبة الخصم", "")).replace("%", "").replace("٪", "").replace(",", "").strip()
    # نقرأ القيمة كعدد عشري قبل أي تقريب
    try:
        _dval = float(discount_raw) if discount_raw and discount_raw.lower() != "nan" else None
    except Exception:
        _dval = None
    if _dval is None or _dval <= 0:
        discount_val = None
    elif _dval < 1:
        # كسر مثل 0.25 يعني 25%
        discount_val = round(_dval * 100)
    else:
        # رقم صحيح مثل 25 أو 50 يعني نسبة مباشرة
        discount_val = round(_dval)
    model_name = str(row.get("الموديل", "")).strip()
    barcode_value = str(row.get("Barcode", "")).strip()
    
    if price_after is None and price_before is not None:
        price_after = price_before
    if price_after is None:
        price_after = 0
    
    # تحميل القالب
    img = Image.open(template_path).convert("RGB")
    draw = ImageDraw.Draw(img)
    CENTER_X = img.width // 2
    
    # الإحداثيات
    PRICE_BEFORE_POS = (CENTER_X + 170, 250)
    LABEL_BEFORE_POS = (CENTER_X + 170, 285)
    PRICE_AFTER_POS = (CENTER_X, 405)
    LABEL_AFTER_POS = (CENTER_X - 25, 455)
    DISCOUNT_POS = (PRICE_AFTER_POS[0] - 350, PRICE_AFTER_POS[1] - 200)
    
    RED_LINE_Y = img.height - 100
    BARCODE_POS_X = 40
    BARCODE_POS_Y = RED_LINE_Y - 5
    
    WHITE_TOP_Y = 535
    
    # السعر قبل
    if price_before:
        before_text = f"{price_before:,}"
        draw.text(PRICE_BEFORE_POS, before_text, fill="red", font=font_price_small, anchor="mm")
        # خط الشطب على السعر القديم
        bb = draw.textbbox(PRICE_BEFORE_POS, before_text, font=font_price_small, anchor="mm")
        draw.line([(bb[0], PRICE_BEFORE_POS[1]), (bb[2], PRICE_BEFORE_POS[1])], fill="red", width=3)
        
        # نسبة الخصم
        if discount_val:
            display_discount = discount_val
        else:
            display_discount = calculate_discount(price_before, price_after)
        
        if display_discount:
            discount_text = f"%{display_discount} خصم"
            bidi = arabic(discount_text)
            draw.text(DISCOUNT_POS, bidi, fill="red", font=font_discount, anchor="lm")
    elif discount_val:
        # نسبة الخصم بدون سعر قبل
        discount_text = f"%{discount_val} خصم"
        bidi = arabic(discount_text)
        draw.text(DISCOUNT_POS, bidi, fill="red", font=font_discount, anchor="lm")
    
    # السعر بعد
    price_text = f"{price_after:,}"
    draw.text(PRICE_AFTER_POS, price_text, fill="black", font=font_price_big, anchor="mm")
    
    bbox = draw.textbbox((0, 0), price_text, font=font_price_big)
    pw = bbox[2] - bbox[0]
    
    # اسم المنتج
    draw_product_name_safe(draw, CENTER_X, WHITE_TOP_Y, product_name, fill="black", max_width=860)
    
    # التسميات
    if price_before:
        label = arabic("قبل")
        draw.text(LABEL_BEFORE_POS, label, fill="red", font=font_label, anchor="mm")
    
    label_after = arabic("بعد")
    draw.text(LABEL_AFTER_POS, label_after, fill="red", font=font_label)
    
    # الباركود والموديل
    if barcode_value and barcode_value != "nan" and barcode_value != "":
        try:
            barcode_img = create_clean_barcode(barcode_value, target_width=140, target_height=32)
            img.paste(barcode_img, (BARCODE_POS_X, BARCODE_POS_Y), barcode_img)
            
            if model_name and model_name != "nan" and model_name != "":
                model_text = f"الموديل: {model_name}"
                bidi = arabic(model_text)
                model_x = BARCODE_POS_X + 330
                model_y = BARCODE_POS_Y - 1
                draw.text((model_x, model_y), bidi, fill="red", font=font_model, anchor="lm")
        except Exception as e:
            print(f"⚠️ خطأ في الباركود: {e}")
    
    # حفظ في buffer
    buf = io.BytesIO()
    img.save(buf, format="PNG", dpi=(300, 300))
    buf.seek(0)
    
    # اسم الملف
    if model_name and model_name != "nan":
        filename = safe_filename(model_name)
    elif product_name and product_name != "nan":
        filename = safe_filename(product_name)
    else:
        filename = "product"
    
    return buf, filename

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

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'GET':
        return send_file('login.html')
    
    data = request.get_json()
    username = data.get('username', '').strip().lower()
    password = data.get('password', '').strip()
    
    if not username or not password:
        return jsonify({"error": "اسم المستخدم وكلمة المرور مطلوبان"}), 400
    
    if len(password) != 6 or not password.isdigit():
        return jsonify({"error": "كلمة المرور يجب أن تكون 6 أرقام"}), 400
    
    users = load_users()
    if username not in users:
        return jsonify({"error": "اسم المستخدم غير موجود"}), 401
    
    if users[username]['password_hash'] != hash_password(password):
        return jsonify({"error": "كلمة المرور غير صحيحة"}), 401
    
    session['user'] = username
    session['role'] = users[username]['role']
    
    return jsonify({
        "success": True,
        "role": users[username]['role'],
        "redirect": "/"
    })

@app.route('/logout')
def logout():
    session.clear()
    return redirect('/login')

@app.route('/')
@require_auth
def index():
    return redirect('/auto')

@app.route('/sheets')
@require_auth
def sheets_page():
    return redirect('/auto')

@app.route('/auto')
@require_auth
def auto_page():
    return send_file('auto.html')

# ══════════════════════════════════════════
#  USER MANAGEMENT API (Admin only)
# ══════════════════════════════════════════

@app.route('/api/users', methods=['GET'])
@require_admin
def get_users():
    users = load_users()
    return jsonify({
        k: {"role": v["role"], "created_at": v.get("created_at", "")}
        for k, v in users.items()
    })

@app.route('/api/users', methods=['POST'])
@require_admin
def create_user():
    data = request.get_json()
    username = data.get('username', '').strip().lower()
    password = data.get('password', '').strip()
    role = data.get('role', 'user')
    
    if not username or not password:
        return jsonify({"error": "اسم المستخدم وكلمة المرور مطلوبان"}), 400
    
    if len(password) != 6 or not password.isdigit():
        return jsonify({"error": "كلمة المرور يجب أن تكون 6 أرقام"}), 400
    
    users = load_users()
    if username in users:
        return jsonify({"error": "اسم المستخدم موجود مسبقاً"}), 400
    
    users[username] = {
        "password_hash": hash_password(password),
        "role": role,
        "created_at": datetime.now().isoformat()
    }
    save_users(users)
    
    return jsonify({"success": True, "message": "تم إنشاء المستخدم بنجاح"})

@app.route('/api/users/<username>', methods=['DELETE'])
@require_admin
def delete_user(username):
    username = username.lower()
    if username == 'admin':
        return jsonify({"error": "لا يمكن حذف حساب الأدمن الرئيسي"}), 400
    
    users = load_users()
    if username not in users:
        return jsonify({"error": "المستخدم غير موجود"}), 404
    
    del users[username]
    save_users(users)
    
    return jsonify({"success": True, "message": "تم حذف المستخدم"})

@app.route('/api/me', methods=['GET'])
def get_current_user():
    if 'user' not in session:
        return jsonify({"authenticated": False})
    return jsonify({
        "authenticated": True,
        "username": session['user'],
        "role": session.get('role', 'user')
    })


@app.route('/generate-pdf', methods=['POST'])
@require_auth
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
@require_auth
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
                    row.get("اسم الصنف") or
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
@require_auth
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


# ══════════════════════════════════════════
#  GOOGLE OAUTH & SHEETS ENDPOINTS
# ══════════════════════════════════════════

@app.route('/api/google/login')
def google_login():
    """بدء تسجيل الدخول بـ Google"""
    if not GOOGLE_CLIENT_ID:
        # Fallback: use local demo mode
        session['user'] = 'demo@local.com'
        session['role'] = 'admin'
        return redirect('/')
    
    flow = Flow.from_client_config(
        {
            "web": {
                "client_id": GOOGLE_CLIENT_ID,
                "client_secret": GOOGLE_CLIENT_SECRET,
                "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                "token_uri": "https://oauth2.googleapis.com/token",
                "redirect_uris": ["http://localhost:8001/oauth/callback"]
            }
        },
        scopes=SCOPES,
        redirect_uri="http://localhost:8001/oauth/callback"
    )
    
    auth_url, state = flow.authorization_url(
        access_type='offline',
        include_granted_scopes='true',
        prompt='consent'
    )
    
    session['oauth_state'] = state
    return redirect(auth_url)

@app.route('/oauth/callback')
def oauth_callback():
    """معالجة رد Google OAuth"""
    try:
        flow = Flow.from_client_config(
            {
                "web": {
                    "client_id": GOOGLE_CLIENT_ID,
                    "client_secret": GOOGLE_CLIENT_SECRET,
                    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                    "token_uri": "https://oauth2.googleapis.com/token",
                    "redirect_uris": ["http://localhost:8001/oauth/callback"]
                }
            },
            scopes=SCOPES,
            redirect_uri="http://localhost:8001/oauth/callback"
        )
        
        flow.fetch_token(authorization_response=request.url)
        credentials = flow.credentials
        
        # جلب معلومات المستخدم
        userinfo_service = build('oauth2', 'v2', credentials=credentials)
        user_info = userinfo_service.userinfo().get().execute()
        
        email = user_info.get('email')
        
        # التحقق من الإيميل إذا كان هناك قائمة مسموح بها
        if ALLOWED_EMAILS and email not in ALLOWED_EMAILS:
            return "Access denied. Email not authorized.", 403
        
        # حفظ بيانات المستخدم في الجلسة
        session['user'] = email
        session['role'] = 'user'
        session['google_creds'] = {
            'token': credentials.token,
            'refresh_token': credentials.refresh_token,
            'token_uri': credentials.token_uri,
            'client_id': credentials.client_id,
            'client_secret': credentials.client_secret,
            'scopes': credentials.scopes
        }
        
        return redirect('/')
        
    except Exception as e:
        print(f"OAuth error: {e}")
        return f"Authentication failed: {str(e)}", 400

@app.route('/api/sheets/data', methods=['POST'])
@require_auth
def get_sheets_data():
    """جلب بيانات من Google Sheet أو من ملف Excel محلي"""
    try:
        data = request.get_json() or {}
        sheet_id = data.get('sheet_id', SHEET_ID)
        
        # Try Google Sheets with Service Account first
        rows = []
        service = create_sheets_service()
        sheet_modified_time = None
        if service:
            try:
                rows = fetch_sheet_data(service, sheet_id, 'A1:Z1000')
                # جلب تاريخ التعديل الحقيقي
                sheet_modified_time = get_sheet_last_modified(service, sheet_id)
            except Exception as e:
                print(f"Service Account error: {e}")
                rows = []
        
        # Fallback: OAuth فقط (بدون Excel)
        if not rows:
            if 'google_creds' in session:
                try:
                    service = create_sheets_service(session['google_creds'])
                    rows = fetch_sheet_data(service, sheet_id, 'A1:Z1000')
                    sheet_modified_time = get_sheet_last_modified(service, sheet_id)
                except Exception as e:
                    print(f"OAuth error: {e}")
                    return jsonify({"error": f"Failed to fetch from Google Sheets: {str(e)}"}), 500
            else:
                return jsonify({"error": "No Google Sheets connection available"}), 500
        
        # معالجة المنتجات ومقارنتها مع المخزن (مع التاريخ الحقيقي)
        processed_rows = process_fetched_products(rows, sheet_modified_time)
        
        return jsonify({
            "success": True,
            "count": len(processed_rows),
            "sheet_modified": sheet_modified_time,
            "data": processed_rows
        })
        
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/upload-excel', methods=['POST'])
@require_auth
def upload_excel():
    """رفع ملف Excel وحفظه كمصدر البيانات"""
    try:
        if 'file' not in request.files:
            return jsonify({"error": "لم يتم إرسال ملف"}), 400
        file = request.files['file']
        if not file.filename.endswith(('.xlsx', '.xls')):
            return jsonify({"error": "يجب أن يكون الملف بصيغة .xlsx أو .xls"}), 400
        file.save(EXCEL_FILE)
        file_size = os.path.getsize(EXCEL_FILE)
        # اقرأ الملف مباشرة بعد الرفع
        rows = read_excel_products()
        if not rows:
            return jsonify({"error": "الملف فارغ أو لا يحتوي على بيانات صحيحة"}), 400
        # احفظ اسم الملف ووقت الرفع
        meta = {"filename": file.filename, "uploaded_at": datetime.now().strftime('%Y-%m-%d %H:%M'), "count": len(rows)}
        with open(os.path.join(BASE_DIR, 'excel_meta.json'), 'w', encoding='utf-8') as f:
            json.dump(meta, f, ensure_ascii=False)
        processed_rows = process_fetched_products(rows, meta['uploaded_at'])
        response = jsonify({
            "success": True,
            "source": "excel",
            "message": f"تم رفع الملف ✅ ({len(rows)} منتج)",
            "filename": file.filename,
            "uploaded_at": meta['uploaded_at'],
            "count": len(processed_rows),
            "data": processed_rows
        })
        response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate'
        response.headers['Pragma'] = 'no-cache'
        return response
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route('/api/sheets/auto-fetch', methods=['GET'])
@require_auth
def auto_fetch_data():
    """جلب البيانات من ملف Excel المرفوع"""
    try:
        # قراءة من ملف Excel المرفوع
        rows = read_excel_products()
        if rows:
            # اقرأ metadata الملف
            meta_file = os.path.join(BASE_DIR, 'excel_meta.json')
            meta = {}
            if os.path.exists(meta_file):
                with open(meta_file, 'r', encoding='utf-8') as f:
                    meta = json.load(f)
            file_time = meta.get('uploaded_at', datetime.now().strftime('%Y-%m-%d %H:%M'))
            filename = meta.get('filename', 'products.xlsx')
            # استخدم آخر وقت تعديل للملف (أدق من وقت الرفع)
            file_mtime = datetime.fromtimestamp(os.path.getmtime(EXCEL_FILE)).strftime('%Y-%m-%d %H:%M')
            processed_rows = process_fetched_products(rows, file_mtime)
            response = jsonify({
                "success": True,
                "source": "excel",
                "message": f"ملف Excel: {filename} ✅",
                "filename": filename,
                "sheet_modified": file_mtime,
                "count": len(processed_rows),
                "data": processed_rows
            })
            response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate'
            response.headers['Pragma'] = 'no-cache'
            return response
        else:
            return jsonify({"error": "no_file", "message": "لم يتم رفع ملف Excel بعد"}), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/sheets/fetch-by-type/<sheet_type>', methods=['GET'])
@require_auth
def fetch_by_type(sheet_type):
    """جلب بيانات من شيت محدد حسب النوع (offers, products, products_other)"""
    try:
        if sheet_type not in SHEET_IDS:
            return jsonify({"error": "Invalid sheet type"}), 400
        
        sheet_id = SHEET_IDS[sheet_type]
        
        if not sheet_id:
            return jsonify({"error": f"Sheet ID not configured for {sheet_type}"}), 400
        
        # Use Service Account or OAuth
        service = create_sheets_service()
        if not service and 'google_creds' in session:
            service = create_sheets_service(session['google_creds'])
        
        if service:
            try:
                rows = fetch_sheet_data(service, sheet_id, 'A1:Z1000')
                return jsonify({
                    "success": True,
                    "source": "google_sheets",
                    "type": sheet_type,
                    "count": len(rows),
                    "data": rows
                })
            except Exception as e:
                print(f"Google Sheets error: {e}")
                return jsonify({"error": "Failed to fetch from Google Sheets"}), 500
        else:
            return jsonify({"error": "Google authentication required"}), 401
            
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/sheets/fetch-all', methods=['GET'])
@require_auth
def fetch_all_sheets():
    """جلب بيانات من جميع الشيتات ودمجها مع تحديد النوع"""
    try:
        all_products = []
        
        # Use Service Account or OAuth
        service = create_sheets_service()
        if not service and 'google_creds' in session:
            service = create_sheets_service(session['google_creds'])
        
        if not service:
            return jsonify({"error": "Google authentication required"}), 401
        
        for sheet_type, sheet_id in SHEET_IDS.items():
            if not sheet_id:
                continue
            
            try:
                rows = fetch_sheet_data(service, sheet_id, 'A1:Z1000')
                # Add template type to each product
                for row in rows:
                    row['_template_type'] = sheet_type
                all_products.extend(rows)
            except Exception as e:
                print(f"Error fetching {sheet_type}: {e}")
                continue
        
        return jsonify({
            "success": True,
            "source": "google_sheets",
            "count": len(all_products),
            "data": all_products
        })
            
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ══════════════════════════════════════════
#  PRODUCT TRACKER API
# ══════════════════════════════════════════

@app.route('/api/products/hide', methods=['POST'])
@require_auth
def hide_product():
    """إخفاء منتج من القائمة"""
    try:
        data = request.get_json()
        product = data.get('product', {})
        product_id = get_product_hash(product)
        
        if not product_id:
            return jsonify({"error": "Invalid product"}), 400
        
        tracker = load_products_tracker()
        if product_id not in tracker:
            tracker[product_id] = {}
        
        tracker[product_id].update({
            'hidden': True,
            'اسم الصنف': product.get('اسم الصنف', ''),
            'السعر بعد الخصم': str(product.get('السعر بعد الخصم ', '')).strip(),
            'السعر قبل الخصم': str(product.get('السعر قبل الخصم', '')).strip(),
            'الموديل': str(product.get('الموديل', '')).strip(),
            'Barcode': str(product.get('Barcode', '')).strip(),
            'hidden_at': datetime.now().isoformat(),
            'last_updated': product.get('_last_updated', datetime.now().strftime('%Y-%m-%d %H:%M'))
        })
        
        save_products_tracker(tracker)
        return jsonify({"success": True, "message": "Product hidden"})
        
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route('/api/products/unhide', methods=['POST'])
@require_auth
def unhide_product():
    """إظهار منتج مخفي"""
    try:
        data = request.get_json()
        product = data.get('product', {})
        product_id = get_product_hash(product)
        
        if not product_id:
            return jsonify({"error": "Invalid product"}), 400
        
        tracker = load_products_tracker()
        if product_id in tracker:
            tracker[product_id]['hidden'] = False
            tracker[product_id]['unhidden_at'] = datetime.now().isoformat()
            save_products_tracker(tracker)
        
        return jsonify({"success": True, "message": "Product unhidden"})
        
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route('/api/products/mark-processed', methods=['POST'])
@require_auth
def mark_product_processed():
    """تحديد أن المنتج تم معالجته (ولكن ما يُخفى تلقائياً)"""
    try:
        data = request.get_json()
        product = data.get('product', {})
        product_id = get_product_hash(product)
        
        if not product_id:
            return jsonify({"error": "Invalid product"}), 400
        
        tracker = load_products_tracker()
        if product_id not in tracker:
            tracker[product_id] = {}
        
        tracker[product_id].update({
            'processed': True,
            'اسم الصنف': str(product.get('اسم الصنف', '')).strip(),
            'السعر بعد الخصم': str(product.get('السعر بعد الخصم ', product.get('السعر بعد الخصم', ''))).strip(),
            'السعر قبل الخصم': str(product.get('السعر قبل الخصم', '')).strip(),
            'الموديل': str(product.get('الموديل', '')).strip(),
            'Brand': str(product.get('Brand', '')).strip(),
            'Barcode': str(product.get('Barcode', '')).strip(),
            'نسبة الخصم': str(product.get('نسبة الخصم', '')).strip(),
            'processed_at': datetime.now().isoformat(),
            'last_updated': datetime.now().strftime('%Y-%m-%d %H:%M'),
        })
        tracker[product_id].pop('change_detected_at', None)  # نمسح وقت الاكتشاف
        
        save_products_tracker(tracker)
        return jsonify({"success": True, "message": "Product marked as processed"})
        
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route('/api/products/reset', methods=['POST'])
@require_auth
def reset_products_tracker():
    """إعادة تعيين قاعدة البيانات (حذف كل المنتجات المخفية)"""
    try:
        if os.path.exists(PRODUCTS_DB_FILE):
            os.remove(PRODUCTS_DB_FILE)
        return jsonify({"success": True, "message": "Tracker reset successfully"})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route('/api/products/hidden', methods=['GET'])
@require_auth
def get_hidden_products():
    """جلب قائمة المنتجات المخفية"""
    try:
        tracker = load_products_tracker()
        hidden = {k: v for k, v in tracker.items() if v.get('hidden', False)}
        return jsonify({"success": True, "count": len(hidden), "data": hidden})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route('/api/sheets/info', methods=['GET'])
@require_auth
def get_sheet_info():
    """جلب معلومات الشيت (تاريخ آخر تحديث)"""
    try:
        service = create_sheets_service()
        if not service:
            return jsonify({"error": "No service account configured"}), 400
        
        target_sheet = request.args.get('sheet_id') or DEFAULT_SHEET_ID or SHEET_ID
        
        # محاولة جلب معلومات من Google Drive API (للmodifiedTime)
        last_modified = get_sheet_last_modified(service, target_sheet)
        
        return jsonify({
            "success": True,
            "sheet_id": target_sheet,
            "last_modified": last_modified,
            "current_time": datetime.now().strftime('%Y-%m-%d %H:%M')
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


def get_local_excel_data():
    """قراءة بيانات من ملف Excel محلي"""
    excel_files = [f for f in os.listdir(BASE_DIR) if f.endswith(('.xlsx', '.xls'))]
    
    if not excel_files:
        return jsonify({"error": "No local Excel files found"}), 404
    
    # Use first Excel file found
    excel_path = os.path.join(BASE_DIR, excel_files[0])
    
    try:
        df = pd.read_excel(excel_path)
        df.columns = df.columns.str.strip()
        
        print(f"✓ Excel file: {excel_files[0]}")
        print(f"✓ Columns found: {list(df.columns)}")
        print(f"✓ Total rows before filter: {len(df)}")
        
        # Filter: keep only rows with product name
        if 'اسم الصنف' in df.columns:
            df = df[df['اسم الصنف'].notna()]
            df = df[df['اسم الصنف'].astype(str).str.strip() != '']
        elif 'اسم الصنف المعتمد' in df.columns:
            df = df[df['اسم الصنف المعتمد'].notna()]
            df = df[df['اسم الصنف المعتمد'].astype(str).str.strip() != '']
        
        # Convert to list of dicts
        rows = df.replace({np.nan: None}).to_dict('records')
        
        print(f"✓ Total rows after filter: {len(rows)}")
        
        return jsonify({
            "success": True,
            "source": excel_files[0],
            "count": len(rows),
            "data": rows
        })
    except Exception as e:
        return jsonify({"error": f"Error reading Excel: {str(e)}"}), 500

@app.route('/api/templates/preview/<template_type>')
@require_auth
def template_preview_image(template_type):
    """إرجاع صورة القالب كـ thumbnail للبريف"""
    if template_type not in TEMPLATES:
        return "Not found", 404
    file_path = os.path.join(BASE_DIR, TEMPLATES[template_type]['file'])
    if not os.path.exists(file_path):
        return "File not found", 404
    ext = os.path.splitext(file_path)[1].lower()
    mime = 'image/jpeg' if ext in ('.jpg', '.jpeg') else 'image/png'
    return send_file(file_path, mimetype=mime)


@app.route('/api/templates', methods=['GET'])
@require_auth
def get_templates():
    """عرض قائمة القوالب المتاحة"""
    return jsonify({
        "success": True,
        "templates": TEMPLATES
    })

@app.route('/api/cards/generate', methods=['POST'])
@require_auth
def generate_cards_zip():
    """توليد كروت ZIP من بيانات مختارة"""
    try:
        data = request.get_json()
        products = data.get('products', [])
        template_type = data.get('template_type', 'products')
        
        if not products:
            return jsonify({"error": "No products selected"}), 400
        
        if template_type not in TEMPLATES:
            return jsonify({"error": "Invalid template type"}), 400
        
        template_file = TEMPLATES[template_type]['file']
        template_path = os.path.join(BASE_DIR, template_file)
        
        if not os.path.exists(template_path):
            return jsonify({"error": f"Template not found: {template_file}"}), 404
        
        # توليد الكروت
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
            for i, product in enumerate(products):
                try:
                    img_buf, filename = process_product_card(template_path, product)
                    zf.writestr(f"{filename}_{i+1}.png", img_buf.getvalue())
                except Exception as e:
                    print(f"Error processing product {i}: {e}")
                    continue
        
        zip_buf.seek(0)
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        return send_file(
            zip_buf,
            mimetype='application/zip',
            as_attachment=True,
            download_name=f'cards_{template_type}_{timestamp}.zip'
        )
        
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/cards/generate-pdf', methods=['POST'])
@require_auth
def generate_cards_pdf():
    """توليد PDF بـ 6 كروت في صفحة واحدة"""
    try:
        data = request.get_json()
        products = data.get('products', [])
        
        if not products:
            return jsonify({"error": "No products selected"}), 400
        
        # توليد كل الكروت كصور
        card_images = []
        for i, product in enumerate(products):
            template_type = product.get('template_type', 'products')
            if template_type not in TEMPLATES:
                template_type = 'products'
            
            template_file = TEMPLATES[template_type]['file']
            template_path = os.path.join(BASE_DIR, template_file)
            
            if not os.path.exists(template_path):
                continue
            
            try:
                img_buf, _ = process_product_card(template_path, product)
                img = Image.open(img_buf).convert("RGB")
                # عدد النسخ المطلوبة لنفس الكرت (افتراضي 1)
                try:
                    qty = int(float(product.get('quantity', 1)))
                except (ValueError, TypeError):
                    qty = 1
                qty = max(1, min(qty, 100))
                for _ in range(qty):
                    card_images.append(img)
            except Exception as e:
                print(f"Error processing product {i}: {e}")
                continue
        
        if not card_images:
            return jsonify({"error": "No cards generated"}), 400
        
        # ✅ بعد الطباعة: حدّث المخزن بالقيم الجديدة (نقطة مرجعية جديدة)
        tracker = load_products_tracker()
        current_time = datetime.now().strftime('%Y-%m-%d %H:%M')
        for product in products:
            product_id = get_product_hash(product)
            if not product_id:
                continue
            if product_id not in tracker:
                tracker[product_id] = {}
            tracker[product_id].update({
                'اسم الصنف': str(product.get('اسم الصنف', '')).strip(),
                'السعر بعد الخصم': str(product.get('السعر بعد الخصم ', product.get('السعر بعد الخصم', ''))).strip(),
                'السعر قبل الخصم': str(product.get('السعر قبل الخصم', '')).strip(),
                'الموديل': str(product.get('الموديل', '')).strip(),
                'Brand': str(product.get('Brand', '')).strip(),
                'Barcode': str(product.get('Barcode', '')).strip(),
                'نسبة الخصم': str(product.get('نسبة الخصم', '')).strip(),
                'last_updated': current_time,
                'printed_at': current_time,
            })
        save_products_tracker(tracker)
        print(f"[TRACKER] ✅ حُدّث المخزن بعد طباعة {len(products)} منتج")

        # إنشاء PDF مع 6 كروت في الصفحة (2×3)
        pdf_buffer = generate_pdf_6_per_page(card_images)

        # ✅ انتهت العملية: احذف ملف الإكسل المرفوع لينتظر النظام رفع ملف جديد
        delete_excel_file()
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        return send_file(
            pdf_buffer,
            mimetype='application/pdf',
            as_attachment=True,
            download_name=f'cards_6perpage_{timestamp}.pdf'
        )
        
    except Exception as e:
        return jsonify({"error": str(e)}), 500

def generate_pdf_6_per_page(card_images):
    """إنشاء PDF مع 6 كروت في الصفحة (2×3)"""
    from reportlab.lib.pagesizes import A4
    
    page_width, page_height = A4
    cards_per_row = 2
    cards_per_col = 3
    cards_per_page = cards_per_row * cards_per_col
    
    # حساب حجم الكرت
    margin_x = 30
    margin_y = 30
    gap_x = 20
    gap_y = 20
    
    available_width = page_width - (2 * margin_x) - ((cards_per_row - 1) * gap_x)
    available_height = page_height - (2 * margin_y) - ((cards_per_col - 1) * gap_y)
    
    card_width = available_width / cards_per_row
    card_height = available_height / cards_per_col

    # معامل تصغير البطاقة (0.9 = أصغر بـ 10%)
    CARD_SCALE = 0.9
    
    pdf_buffer = io.BytesIO()
    c = canvas.Canvas(pdf_buffer, pagesize=A4)
    
    for i, img in enumerate(card_images):
        pos_index = i % cards_per_page
        
        if pos_index == 0 and i > 0:
            c.showPage()
        
        row = pos_index // cards_per_row
        col = pos_index % cards_per_row
        
        x = margin_x + (col * (card_width + gap_x))
        y = page_height - margin_y - ((row + 1) * (card_height + gap_y)) + gap_y

        # تصغير البطاقة مع توسيطها داخل خانتها
        draw_w = card_width * CARD_SCALE
        draw_h = card_height * CARD_SCALE
        draw_x = x + (card_width - draw_w) / 2
        draw_y = y + (card_height - draw_h) / 2
        
        # حفظ الصورة مؤقتاً
        img_buffer = io.BytesIO()
        img.save(img_buffer, format="PNG", dpi=(300, 300))
        img_buffer.seek(0)
        
        # رسم الصورة
        img_reader = ImageReader(img_buffer)
        c.drawImage(img_reader, draw_x, draw_y, width=draw_w, height=draw_h, preserveAspectRatio=True)
    
    c.save()
    pdf_buffer.seek(0)
    return pdf_buffer


@app.route('/api/cards/generate-png', methods=['POST'])
@require_auth
def generate_cards_png():
    """توليد صور PNG (ZIP) بنفس إعدادات card_generator.py"""
    try:
        data = request.get_json()
        products = data.get('products', [])
        
        if not products:
            return jsonify({"error": "No products selected"}), 400
        
        # نفس FIELDS من card_generator.py
        FIELDS_CONFIG = [
            {"name": "السعر الجديد", "col": "السعر بعد الخصم ", "type": "price",
             "x": 190, "y": 295, "w": 220, "h": 100, "size": 78, "color": "#000000", "bold": True},
            {"name": "السعر القديم", "col": "السعر قبل الخصم", "type": "price_strike",
             "x": 195, "y": 415, "w": 210, "h": 65, "size": 46, "color": "#cc0000", "bold": True},
            {"name": "اسم المنتج", "col": "اسم الصنف", "type": "text",
             "x": 40, "y": 500, "w": 520, "h": 55, "size": 28, "color": "#111111", "bold": False},
            {"name": "الموديل", "col": "الموديل", "type": "text",
             "x": 80, "y": 565, "w": 440, "h": 38, "size": 20, "color": "#333333", "bold": False},
            {"name": "الباركود", "col": "Barcode", "type": "barcode",
             "x": 170, "y": 615, "w": 260, "h": 55, "size": 0, "color": "#000000", "bold": False},
            {"name": "نص قبل الخصم", "col": "", "type": "static_label",
             "x": 150, "y": 385, "w": 100, "h": 35, "size": 24, "color": "#ff0000", "bold": True, "staticVal": "قبل الخصم"},
            {"name": "نص السعر الجديد", "col": "", "type": "static_label",
             "x": 150, "y": 265, "w": 100, "h": 35, "size": 24, "color": "#00aa00", "bold": True, "staticVal": "السعر الجديد"},
        ]
        
        # إنشاء ZIP
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
            for i, product in enumerate(products):
                template_type = product.get('template_type', 'products')
                if template_type not in TEMPLATES:
                    template_type = 'products'
                
                template_file = TEMPLATES[template_type]['file']
                template_path = os.path.join(BASE_DIR, template_file)
                
                # تحميل القالب
                if os.path.exists(template_path):
                    img = Image.open(template_path).convert("RGB")
                else:
                    # قالب افتراضي 600x720
                    img = Image.new('RGB', (600, 720), '#f5f5f5')
                    draw_temp = ImageDraw.Draw(img)
                    draw_temp.rectangle([0, 500, 600, 720], fill='white')
                    draw_temp.rectangle([0, 650, 600, 720], fill='#c50000')
                    draw_temp.rectangle([12, 12, 588, 708], outline='#ffd700', width=3)
                
                draw = ImageDraw.Draw(img)
                
                # رسم كل الحقول
                for field in FIELDS_CONFIG:
                    draw_field_on_image(draw, img, field, product)
                
                # اسم الملف
                model_name = str(product.get('الموديل', '')).strip()
                product_name = str(product.get('اسم الصنف', '')).strip()
                
                if model_name and model_name != "nan":
                    fname = re.sub(r'[\\/:*?"<>|]', '', model_name).strip() or f"product_{i+1}"
                elif product_name and product_name != "nan":
                    fname = re.sub(r'[\\/:*?"<>|]', '', product_name).strip() or f"product_{i+1}"
                else:
                    fname = f"product_{i+1}"
                
                # حفظ في ZIP
                img_buf = io.BytesIO()
                img.save(img_buf, format="PNG", dpi=(300, 300))
                img_buf.seek(0)
                zf.writestr(f"{fname}.png", img_buf.read())
        
        zip_buf.seek(0)

        # ✅ انتهت العملية: احذف ملف الإكسل المرفوع لينتظر النظام رفع ملف جديد
        delete_excel_file()

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        return send_file(
            zip_buf,
            mimetype='application/zip',
            as_attachment=True,
            download_name=f'cards_png_{timestamp}.zip'
        )
        
    except Exception as e:
        print(f"Error in generate-png: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


if __name__ == '__main__':
    print("=" * 50)
    print("  مصمم كروت الأسعار — السيرفر يعمل")
    print("  افتح المتصفح على: http://localhost:5000")
    print("=" * 50)
    app.run(debug=True, host='0.0.0.0', port=8001)
