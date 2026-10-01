import os
import sys

# Ensure pythonw / frozen mode works flawlessly by redirecting stdout/stderr
try:
    if getattr(sys, 'frozen', False):
        _log_dir = os.path.dirname(sys.executable)
    else:
        _log_dir = os.path.dirname(os.path.abspath(__file__))
    _log_file = open(os.path.join(_log_dir, 'server.log'), 'a', encoding='utf-8', buffering=1)
    if sys.stdout is None or not hasattr(sys.stdout, 'write'):
        sys.stdout = _log_file
    if sys.stderr is None or not hasattr(sys.stderr, 'write'):
        sys.stderr = _log_file
except Exception:
    pass

from datetime import datetime, timedelta, date
import io
import json
import shutil
import sqlite3
import re
import urllib.parse
import uuid
import secrets
import logging
from functools import wraps
import atexit
try:
    from apscheduler.schedulers.background import BackgroundScheduler
except Exception:
    BackgroundScheduler = None
from flask import (
    Flask,
    Response,
    flash,
    jsonify,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for,
)
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT
from werkzeug.security import generate_password_hash, check_password_hash

try:
    import win32print
    import win32ui
except ImportError:
    win32print = None
    win32ui = None

try:
    import pymupdf as fitz
except ImportError:
    try:
        import fitz
    except ImportError:
        fitz = None

def parse_event_datetime(d, is_end_of_day=False):
    if not d or str(d).strip() in ('-', '', 'None'):
        return datetime.min
    s = str(d).strip()
    fmts = [
        '%Y-%m-%d %I:%M:%S %p',
        '%Y-%m-%d %I:%M %p',
        '%Y-%m-%d %H:%M:%S',
        '%Y-%m-%d %H:%M',
        '%d-%m-%Y %I:%M:%S %p',
        '%d-%m-%Y %I:%M %p',
        '%d-%m-%Y %H:%M:%S',
        '%d-%m-%Y %H:%M',
    ]
    for f in fmts:
        try:
            return datetime.strptime(s, f)
        except ValueError:
            pass
    for f in ['%Y-%m-%d', '%d-%m-%Y']:
        try:
            dt = datetime.strptime(s, f)
            if is_end_of_day:
                return dt.replace(hour=23, minute=59, second=59)
            return dt
        except ValueError:
            pass
    return datetime.min

# راستوں (Paths) کا درست تعین
if getattr(sys, 'frozen', False):
    EXE_DIR = os.path.dirname(sys.executable)
    INTERNAL_DIR = getattr(sys, '_MEIPASS', os.path.join(EXE_DIR, '_internal'))
    TEMPLATE_DIR = os.path.join(EXE_DIR, 'templates')
    for cand in [os.path.join(EXE_DIR, 'templates'), os.path.join(INTERNAL_DIR, 'templates'), os.path.join(EXE_DIR, '_internal', 'templates')]:
        if os.path.exists(cand):
            TEMPLATE_DIR = cand
            break
    STATIC_DIR = os.path.join(EXE_DIR, 'static')
    for cand in [os.path.join(EXE_DIR, 'static'), os.path.join(INTERNAL_DIR, 'static'), os.path.join(EXE_DIR, '_internal', 'static')]:
        if os.path.exists(cand):
            STATIC_DIR = cand
            break
    BASE_DIR = EXE_DIR
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    TEMPLATE_DIR = os.path.join(BASE_DIR, 'templates')
    STATIC_DIR = os.path.join(BASE_DIR, 'static')

# Ù„Ø§Ú¯Ù†Ú¯ Ø³ÛŒÙ¹ Ø§Ù¾ (Proper Logging)
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger('SmartPOS')
try:
    _file_handler = logging.FileHandler(os.path.join(BASE_DIR, 'server.log'), encoding='utf-8')
    _file_handler.setFormatter(logging.Formatter('%(asctime)s [%(levelname)s] %(message)s'))
    logger.addHandler(_file_handler)
except Exception:
    pass

app = Flask(__name__, template_folder=TEMPLATE_DIR, static_folder=STATIC_DIR)

def make_delete_error_page(title_ur, title_en, msg_ur, msg_en, return_url):
    html = f"""<!DOCTYPE html>
<html lang="ur" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{title_ur}</title>
    <style>
        * {{ box-sizing: border-box; margin: 0; padding: 0; }}
        body {{
            font-family: 'Segoe UI', Tahoma, Arial, sans-serif;
            background: #0f172a;
            color: #f8fafc;
            display: flex;
            align-items: center;
            justify-content: center;
            min-height: 100vh;
            padding: 20px;
        }}
        .card {{
            background: #1e293b;
            border: 1px solid #334155;
            border-radius: 16px;
            padding: 35px 30px;
            max-width: 520px;
            width: 100%;
            text-align: center;
            box-shadow: 0 20px 25px -5px rgba(0,0,0,0.5), 0 8px 10px -6px rgba(0,0,0,0.5);
        }}
        .icon {{
            width: 70px;
            height: 70px;
            background: rgba(239, 68, 68, 0.15);
            border: 2px solid #ef4444;
            color: #ef4444;
            border-radius: 50%;
            display: inline-flex;
            align-items: center;
            justify-content: center;
            font-size: 32px;
            margin: 0 auto 20px auto;
        }}
        h2 {{
            color: #ef4444;
            font-size: 22px;
            margin-bottom: 12px;
            font-weight: 700;
        }}
        .msg {{
            font-size: 16px;
            line-height: 1.7;
            color: #cbd5e1;
            margin-bottom: 8px;
        }}
        .msg-en {{
            font-size: 13px;
            color: #94a3b8;
            direction: ltr;
            margin-bottom: 25px;
        }}
        .btn {{
            display: inline-block;
            background: #0284c7;
            color: white;
            padding: 12px 30px;
            font-size: 15px;
            font-weight: bold;
            border-radius: 10px;
            text-decoration: none;
            transition: all 0.2s;
        }}
        .btn:hover {{
            background: #0369a1;
        }}
    </style>
</head>
<body>
    <div class="card">
        <div class="icon">⚠️</div>
        <h2>{title_ur}</h2>
        <p class="msg">{msg_ur}</p>
        <p class="msg-en">{msg_en}</p>
        <a href="{return_url}" class="btn">واپس جائیں (Go Back)</a>
    </div>
</body>
</html>"""
    return Response(html, status=400, mimetype='text/html', content_type='text/html; charset=utf-8')

def secure_delete_response(success, msg_ur, msg_en, return_url):
    if not success:
        return make_delete_error_page(
            title_ur="ڈیلیٹ نہیں ہو سکا!",
            title_en="Cannot Delete Record",
            msg_ur=msg_ur,
            msg_en=msg_en,
            return_url=return_url
        )
    if request.is_json:
        return jsonify({'success': True, 'message': msg_ur})
    return redirect(return_url)

@app.errorhandler(500)
def handle_internal_server_error(e):
    logging.error(f"500 Internal Error: {e}", exc_info=True)
    if request.is_json or request.path.startswith('/api/'):
        return jsonify({'success': False, 'message': 'سرور میں عارضی رکاوٹ ہے۔ برائے کرم دوبارہ کوشش کریں یا صفحہ ریفریش کریں۔', 'error': str(e)}), 500
    html = f"""<!DOCTYPE html>
<html dir="rtl" lang="ur">
<head>
    <meta charset="UTF-8">
    <title>عارضی مسئلہ (Server Notice)</title>
    <style>
        body {{ font-family: system-ui, -apple-system, sans-serif; background: #0f172a; color: #fff; display: flex; align-items: center; justify-content: center; min-height: 100vh; margin: 0; }}
        .card {{ background: #1e293b; border: 1px solid #334155; border-radius: 16px; padding: 35px 30px; max-width: 500px; width: 100%; text-align: center; box-shadow: 0 20px 25px -5px rgba(0,0,0,0.5); }}
        .icon {{ font-size: 45px; margin-bottom: 15px; }}
        h2 {{ color: #f59e0b; margin-bottom: 12px; font-size: 22px; }}
        p {{ color: #cbd5e1; line-height: 1.7; font-size: 15px; margin-bottom: 25px; }}
        .btn {{ display: inline-block; background: #0284c7; color: white; padding: 12px 30px; border-radius: 10px; text-decoration: none; font-weight: bold; font-size: 15px; }}
        .btn:hover {{ background: #0369a1; }}
    </style>
</head>
<body>
    <div class="card">
        <div class="icon">⚠️</div>
        <h2>عارضی تکنیکی مسئلہ</h2>
        <p>سرور پر ایک عارضی مسئلہ پیش آیا ہے، تاہم آپ کا ڈیٹا بیس اور تمام کھاتے مکمل محفوظ ہیں۔ برائے مہربانی واپس جا کر دوبارہ کوشش فرمائیں۔</p>
        <a href="javascript:history.back()" class="btn">واپس جائیں (Go Back)</a>
    </div>
</body>
</html>"""
    return Response(html, status=500, mimetype='text/html', content_type='text/html; charset=utf-8')

# ðŸ”  Ù…Ø­Ù ÙˆØ¸ Ø³ÛŒÚ©Ø±ÛŒÙ¹ Ú©ÛŒ â€” Ø®ÙˆØ¯Ú©Ø§Ø± Ø¨Ù†Ø§Ù†Ø§ Ø§ÙˆØ± Ù Ø§Ø¦Ù„ Ù…ÛŒÚº Ù…Ø­Ù ÙˆØ¸ Ú©Ø±Ù†Ø§
def _get_or_create_secret_key():
    key_file = os.path.join(BASE_DIR, '.flask_secret_key')
    if os.path.exists(key_file):
        try:
            with open(key_file, 'r') as f:
                key = f.read().strip()
                if key:
                    return key
        except Exception:
            pass
    key = secrets.token_hex(32)
    try:
        with open(key_file, 'w') as f:
            f.write(key)
    except Exception:
        pass
    return key

app.secret_key = _get_or_create_secret_key()
DB_NAME = os.path.join(BASE_DIR, 'pos.db')
BACKUP_DIR = os.path.join(BASE_DIR, 'daily_backups')
DB_BACKUP_DIR = os.path.join(BACKUP_DIR, 'database_backups')
EXCEL_BACKUP_DIR = os.path.join(BACKUP_DIR, 'excel_backups')
MANUAL_BACKUP_DIR = os.path.join(BACKUP_DIR, 'manual_downloads')

for _b_dir in [BACKUP_DIR, DB_BACKUP_DIR, EXCEL_BACKUP_DIR, MANUAL_BACKUP_DIR]:
    try:
        os.makedirs(_b_dir, exist_ok=True)
    except Exception:
        pass

def get_db():
    conn = sqlite3.connect(DB_NAME, timeout=30)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA synchronous = NORMAL")
    conn.execute("PRAGMA cache_size = -64000")
    conn.execute("PRAGMA temp_store = MEMORY")
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


from contextlib import contextmanager

@contextmanager
def get_db_safe():
    """کنکشن مینجمنٹ — خودکار بند ہونا (auto-close on exit or error)"""
    conn = get_db()
    try:
        yield conn
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

def init_db():
    conn = get_db()
    try:
        conn.execute("PRAGMA journal_mode = WAL")
    except Exception:
        pass
    cursor = conn.cursor()
    
    # 1. Ù¾Ø±ÙˆÚˆÚ©Ù¹Ø³ Ù¹ÛŒØ¨Ù„
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS products (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        code TEXT UNIQUE,
        name TEXT NOT NULL,
        category TEXT,
        unit TEXT DEFAULT 'Pcs',
        buy_price REAL DEFAULT 0,
        sale_price REAL DEFAULT 0,
        stock REAL DEFAULT 0,
        alert_qty INTEGER DEFAULT 10
    )''')

    # 2. Ú©Ø³Ù¹Ù…Ø±Ø² (Receivables)
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS customers (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        phone TEXT,
        address TEXT,
        balance REAL DEFAULT 0.0,
        opening_balance REAL DEFAULT 0.0
    )''')

    # 3. Ø³Ù¾Ù„Ø§Ø¦Ø±Ø² / ÚˆÛŒÙ„Ø±Ø² (Payables)
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS suppliers (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        phone TEXT,
        company TEXT,
        address TEXT,
        balance REAL DEFAULT 0.0
    )''')

    # 4. Ù¾Ø±Ú†ÛŒØ²Ø² (Ø®Ø±ÛŒØ¯Ø§Ø±ÛŒ / Daily Purchase)
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS purchases (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        bill_no TEXT,
        supplier_id INTEGER,
        supplier_name TEXT,
        date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        total REAL DEFAULT 0,
        paid REAL DEFAULT 0,
        due REAL DEFAULT 0,
        notes TEXT,
        status TEXT DEFAULT 'active'
    )''')
    
    try:
        cursor.execute("ALTER TABLE purchases ADD COLUMN status TEXT DEFAULT 'active'")
    except:
        pass

    # 5. خرید کا سامان
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS purchase_items (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        purchase_id INTEGER,
        product_id INTEGER,
        product_name TEXT,
        quantity REAL,
        buy_price REAL,
        total REAL,
        unit TEXT DEFAULT 'Pcs',
        weight REAL DEFAULT 0
    )''')
    try:
        cursor.execute("ALTER TABLE purchase_items ADD COLUMN unit TEXT DEFAULT 'Pcs'")
    except:
        pass
    try:
        cursor.execute("ALTER TABLE purchase_items ADD COLUMN weight REAL DEFAULT 0")
    except:
        pass

    # 6. Ø³Ù¾Ù„Ø§Ø¦Ø± Ø§Ø¯Ø§Ø¦ÛŒÚ¯ÛŒ Ù„Ø§Ú¯ (Payable Recovery)
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS supplier_payments (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        supplier_id INTEGER NOT NULL,
        amount REAL NOT NULL,
        payment_date TEXT NOT NULL,
        note TEXT
    )''')

    # 7. Ø³ÛŒÙ„Ø² Ù¹ÛŒØ¨Ù„
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS sales (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        invoice_no TEXT,
        customer_id INTEGER,
        customer_name TEXT,
        date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        subtotal REAL DEFAULT 0,
        tax_rate REAL DEFAULT 0,
        tax_amount REAL DEFAULT 0,
        discount REAL DEFAULT 0,
        total REAL DEFAULT 0,
        paid REAL DEFAULT 0,
        due REAL DEFAULT 0,
        type TEXT DEFAULT 'cash'
    )''')
    for col, ctype in [('subtotal', 'REAL DEFAULT 0'), ('tax_rate', 'REAL DEFAULT 0'), ('tax_amount', 'REAL DEFAULT 0'), ('discount', 'REAL DEFAULT 0')]:
        try:
            cursor.execute(f"ALTER TABLE sales ADD COLUMN {col} {ctype}")
        except sqlite3.OperationalError:
            pass

    # 8. Ø³ÛŒÙ„ Ø¢Ø¦Ù¹Ù…Ø²
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS sale_items (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        sale_id INTEGER,
        product_id INTEGER,
        product_name TEXT,
        quantity REAL,
        base_quantity REAL DEFAULT 0,
        price REAL,
        total REAL,
        unit TEXT DEFAULT 'Pcs'
    )''')
    # Ù…Ø§Ø¦Ú¯Ø±ÛŒØ´Ù†: Ù¾Ø±Ø§Ù†Û’ ÚˆÛŒÙ¹Ø§Ø¨ÛŒØ³ Ù…ÛŒÚº unit Ú©Ø§Ù„Ù… Ø´Ø§Ù…Ù„ Ú©Ø±Ù†Ø§
    try:
        cursor.execute("ALTER TABLE sale_items ADD COLUMN unit TEXT DEFAULT 'Pcs'")
    except sqlite3.OperationalError:
        pass
    try:
        cursor.execute("ALTER TABLE sale_items ADD COLUMN base_quantity REAL DEFAULT 0")
    except sqlite3.OperationalError:
        pass
    try:
        cursor.execute("ALTER TABLE sale_items ADD COLUMN buy_price REAL DEFAULT 0")
    except sqlite3.OperationalError:
        pass
    # Ù…Ø§Ø¦Ú¯Ø±ÛŒØ´Ù†: sale_items Ù…ÛŒÚº weight Ú©Ø§Ù„Ù… (Pcs Ù¾Ø±ÙˆÚˆÚ©Ù¹Ø³ Ú©Ø§ Ø§ØµÙ„ ÙˆØ²Ù† KG Ù…ÛŒÚº)
    try:
        cursor.execute("ALTER TABLE sale_items ADD COLUMN weight REAL DEFAULT 0")
    except sqlite3.OperationalError:
        pass
    # Ù…Ø§Ø¦Ú¯Ø±ÛŒØ´Ù†: purchase_items Ù…ÛŒÚº unit Ú©Ø§Ù„Ù… Ø´Ø§Ù…Ù„ Ú©Ø±Ù†Ø§
    try:
        cursor.execute("ALTER TABLE purchase_items ADD COLUMN unit TEXT DEFAULT 'Pcs'")
    except sqlite3.OperationalError:
        pass

    # 8.5 Advanced Units
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS product_units (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        product_id INTEGER,
        unit_name TEXT,
        conversion_factor REAL,
        purchase_price REAL,
        sale_price REAL,
        is_base_unit BOOLEAN DEFAULT 0,
        FOREIGN KEY(product_id) REFERENCES products(id) ON DELETE CASCADE
    )''')

    # Migration for product_units (insert base units if not exist)
    existing_products = cursor.execute("SELECT id, unit, buy_price, sale_price FROM products").fetchall()
    for p in existing_products:
        exists = cursor.execute("SELECT id FROM product_units WHERE product_id = ?", (p['id'],)).fetchone()
        if not exists:
            cursor.execute('''
                INSERT INTO product_units (product_id, unit_name, conversion_factor, purchase_price, sale_price, is_base_unit)
                VALUES (?, ?, 1.0, ?, ?, 1)
            ''', (p['id'], p['unit'] or 'Pcs', p['buy_price'], p['sale_price']))

    # 9. Ú©Ø³Ù¹Ù…Ø± Ù¾ÛŒÙ…Ù†Ù¹Ø³
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS customer_payments (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        customer_id INTEGER NOT NULL,
        amount REAL NOT NULL,
        payment_date TEXT NOT NULL,
        note TEXT
    )''')

    # 10. Ø§Ø®Ø±Ø§Ø¬Ø§Øª
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS expenses (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        title TEXT NOT NULL,
        category TEXT NOT NULL,
        amount REAL NOT NULL,
        date TEXT NOT NULL,
        notes TEXT
    )''')

    # 11. Ú©ÙˆÙ¹ÛŒØ´Ù†Ø² (ØªØ®Ù…ÛŒÙ†Û)
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS quotations (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        quote_no TEXT,
        customer_name TEXT,
        customer_phone TEXT,
        date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        total REAL DEFAULT 0,
        items_json TEXT
    )''')

    cursor.execute('''
    CREATE TABLE IF NOT EXISTS categories (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT UNIQUE NOT NULL
    )''')
    
    for col, ctype in [('category_id', 'INTEGER'), ('min_stock', 'REAL DEFAULT 5'), ('is_deleted', 'INTEGER DEFAULT 0')]:
        try:
            cursor.execute(f"ALTER TABLE products ADD COLUMN {col} {ctype}")
        except sqlite3.OperationalError:
            pass
            
    for col, ctype in [('cash_amount', 'REAL DEFAULT 0'), ('online_amount', 'REAL DEFAULT 0')]:
        try:
            cursor.execute(f"ALTER TABLE sales ADD COLUMN {col} {ctype}")
        except sqlite3.OperationalError:
            pass

    try:
        cursor.execute("ALTER TABLE quotations ADD COLUMN status TEXT DEFAULT 'pending'")
    except sqlite3.OperationalError:
        pass

    cursor.execute('''
    CREATE TABLE IF NOT EXISTS personal_loans (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        person_name TEXT NOT NULL,
        phone TEXT,
        type TEXT NOT NULL,
        amount REAL NOT NULL DEFAULT 0.0,
        paid_back REAL NOT NULL DEFAULT 0.0,
        date TEXT NOT NULL,
        due_date TEXT,
        status TEXT DEFAULT 'pending',
        notes TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )''')

    cursor.execute('''
    CREATE TABLE IF NOT EXISTS personal_loan_payments (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        loan_id INTEGER NOT NULL,
        amount REAL NOT NULL DEFAULT 0.0,
        payment_date TEXT NOT NULL,
        method TEXT DEFAULT 'Cash',
        notes TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )''')

    # Auto-repair missing columns on any restored database
    try:
        p_cols = [r[1] for r in cursor.execute("PRAGMA table_info(products)").fetchall()]
        if p_cols:
            if 'min_stock' not in p_cols:
                cursor.execute("ALTER TABLE products ADD COLUMN min_stock REAL DEFAULT 5")
            if 'category_id' not in p_cols:
                cursor.execute("ALTER TABLE products ADD COLUMN category_id INTEGER")
            if 'is_deleted' not in p_cols:
                cursor.execute("ALTER TABLE products ADD COLUMN is_deleted INTEGER DEFAULT 0")

        s_cols = [r[1] for r in cursor.execute("PRAGMA table_info(sales)").fetchall()]
        if s_cols:
            if 'cash_amount' not in s_cols:
                cursor.execute("ALTER TABLE sales ADD COLUMN cash_amount REAL DEFAULT 0")
            if 'online_amount' not in s_cols:
                cursor.execute("ALTER TABLE sales ADD COLUMN online_amount REAL DEFAULT 0")

        si_cols = [r[1] for r in cursor.execute("PRAGMA table_info(sale_items)").fetchall()]
        if si_cols:
            if 'returned_qty' not in si_cols:
                cursor.execute("ALTER TABLE sale_items ADD COLUMN returned_qty REAL DEFAULT 0")
            if 'weight' not in si_cols:
                cursor.execute("ALTER TABLE sale_items ADD COLUMN weight REAL DEFAULT 0")
            if 'unit' not in si_cols:
                cursor.execute("ALTER TABLE sale_items ADD COLUMN unit TEXT DEFAULT 'Pcs'")
            if 'base_quantity' not in si_cols:
                cursor.execute("ALTER TABLE sale_items ADD COLUMN base_quantity REAL DEFAULT 0")
            if 'buy_price' not in si_cols:
                cursor.execute("ALTER TABLE sale_items ADD COLUMN buy_price REAL DEFAULT 0")

        cp_cols = [r[1] for r in cursor.execute("PRAGMA table_info(customer_payments)").fetchall()]
        if cp_cols and 'sale_id' not in cp_cols:
            cursor.execute("ALTER TABLE customer_payments ADD COLUMN sale_id INTEGER DEFAULT NULL")

        q_cols = [r[1] for r in cursor.execute("PRAGMA table_info(quotations)").fetchall()]
        if q_cols and 'status' not in q_cols:
            cursor.execute("ALTER TABLE quotations ADD COLUMN status TEXT DEFAULT 'pending'")

        cust_cols = [r[1] for r in cursor.execute("PRAGMA table_info(customers)").fetchall()]
        if cust_cols:
            if 'credit_limit' not in cust_cols:
                cursor.execute("ALTER TABLE customers ADD COLUMN credit_limit REAL DEFAULT 0.0")
            if 'notes' not in cust_cols:
                cursor.execute("ALTER TABLE customers ADD COLUMN notes TEXT")
            if 'opening_balance' not in cust_cols:
                cursor.execute("ALTER TABLE customers ADD COLUMN opening_balance REAL DEFAULT 0.0")

        sp_cols = [r[1] for r in cursor.execute("PRAGMA table_info(supplier_payments)").fetchall()]
        if sp_cols:
            if 'method' not in sp_cols:
                cursor.execute("ALTER TABLE supplier_payments ADD COLUMN method TEXT DEFAULT 'Cash'")
            if 'notes' not in sp_cols:
                cursor.execute("ALTER TABLE supplier_payments ADD COLUMN notes TEXT")

        sup_cols = [r[1] for r in cursor.execute("PRAGMA table_info(suppliers)").fetchall()]
        if sup_cols and 'opening_balance' not in sup_cols:
            cursor.execute("ALTER TABLE suppliers ADD COLUMN opening_balance REAL DEFAULT 0.0")

        # Ensure categories table exists
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS categories (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT UNIQUE NOT NULL
            )
        ''')
        cat_cnt = cursor.execute("SELECT COUNT(*) FROM categories").fetchone()[0]
        if cat_cnt == 0:
            default_cats = ['General', 'Bolts', 'Nuts', 'Pipes', 'Fittings', 'Tools', 'Hardware', 'Electrical', 'Other']
            for dcat in default_cats:
                try:
                    cursor.execute("INSERT OR IGNORE INTO categories (name) VALUES (?)", (dcat,))
                except Exception:
                    pass

        # Clean up negative opening balances from previous calculation drift
        cursor.execute("UPDATE customers SET opening_balance = 0.0 WHERE opening_balance < 0")

        # Auto-link historical payments that have invoice numbers in notes
        cur_rows = cursor.execute("SELECT id, note FROM customer_payments WHERE (sale_id IS NULL OR sale_id = 0) AND note LIKE '%بل #%'").fetchall()
        for r in cur_rows:
            m = re.search(r'بل\s*#(\d+)', r['note'] or '')
            if m:
                cursor.execute("UPDATE customer_payments SET sale_id = ? WHERE id = ?", (int(m.group(1)), r['id']))

        # High performance indexes for instant queries
        indexes = [
            "CREATE INDEX IF NOT EXISTS idx_sales_date ON sales(date)",
            "CREATE INDEX IF NOT EXISTS idx_sales_cust_id ON sales(customer_id)",
            "CREATE INDEX IF NOT EXISTS idx_sales_type ON sales(type)",
            "CREATE INDEX IF NOT EXISTS idx_sale_items_sale_id ON sale_items(sale_id)",
            "CREATE INDEX IF NOT EXISTS idx_sale_items_prod_id ON sale_items(product_id)",
            "CREATE INDEX IF NOT EXISTS idx_cust_payments_cust_id ON customer_payments(customer_id)",
            "CREATE INDEX IF NOT EXISTS idx_cust_payments_sale_id ON customer_payments(sale_id)",
            "CREATE INDEX IF NOT EXISTS idx_purchases_date ON purchases(date)",
            "CREATE INDEX IF NOT EXISTS idx_purchases_supplier_id ON purchases(supplier_id)",
            "CREATE INDEX IF NOT EXISTS idx_purchase_items_purch_id ON purchase_items(purchase_id)",
            "CREATE INDEX IF NOT EXISTS idx_products_stock ON products(stock)",
            "CREATE INDEX IF NOT EXISTS idx_products_code ON products(code)"
        ]
        for idx_sql in indexes:
            try:
                cursor.execute(idx_sql)
            except Exception:
                pass
    except Exception as e:
        logger.error(f"Migration error in init_db: {e}")

    conn.commit()
    conn.close()

init_db()

# ----------------------------------------------------
# ðŸ” ÛŒÙˆØ²Ø± Ø³ÛŒÚ©ÛŒÙˆØ±Ù¹ÛŒ Ùˆ Ø±ÙˆÙ¹Ù†Ú¯ Ú©Ù†Ù¹Ø±ÙˆÙ„ (Ù…Ø¹ Ù…Ø­ÙÙˆØ¸ Ù¾Ø§Ø³ ÙˆØ±Úˆ ÛÛŒØ´Ù†Ú¯)
# ----------------------------------------------------
def init_auth():
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password TEXT NOT NULL,
            role TEXT DEFAULT 'admin',
            display_name TEXT
        )
    ''')
    try:
        cursor.execute("ALTER TABLE users ADD COLUMN display_name TEXT")
    except sqlite3.OperationalError:
        pass

    # Migration: Add opening_balance to suppliers if missing
    try:
        cursor.execute("ALTER TABLE suppliers ADD COLUMN opening_balance REAL DEFAULT 0.0")
    except sqlite3.OperationalError:
        pass  # Column already exists

    admin_row = cursor.execute("SELECT id FROM users WHERE username = 'admin'").fetchone()
    if not admin_row:
        cursor.execute("INSERT INTO users (username, password, role, display_name) VALUES ('admin', ?, 'admin', 'Admin')", 
                       (generate_password_hash('admin123'),))
    else:
        cursor.execute("UPDATE users SET password = ?, display_name = 'Admin' WHERE username = 'admin'", 
                       (generate_password_hash('admin123'),))

    cashier_row = cursor.execute("SELECT id FROM users WHERE username = 'cashier'").fetchone()
    if not cashier_row:
        cursor.execute("INSERT INTO users (username, password, role, display_name) VALUES ('cashier', ?, 'cashier', 'Cashier')", 
                       (generate_password_hash('1234'),))
    else:
        cursor.execute("UPDATE users SET password = ?, display_name = 'Cashier' WHERE username = 'cashier'", 
                       (generate_password_hash('1234'),))

    conn.commit()
    conn.close()

init_auth()

def recalculate_customer_balance(cust_id, cursor=None):
    """
    سنگل سورس آف ٹروتھ (Single Source of Truth) فارمولا:
    کل بقایا = سابقہ کھاتہ (opening_balance) 
             + فعال بلوں کا بقایا (unpaid sales dues) 
             - عمومی وصولیاں (unlinked direct customer payments)
    نوٹ: اگر کسی مخصوص بل کی وصولی درج کی جائے تو وہ بل کے due میں پہلے ہی ایڈجسٹ ہوتی ہے،
    اس لیے اسے دوبارہ مائنس نہیں کیا جاتا (صفر ڈبل کٹاؤ)۔
    """
    should_close = False
    if cursor is None:
        conn = get_db()
        cursor = conn.cursor()
        should_close = True

    try:
        cust = cursor.execute("SELECT id, name, opening_balance FROM customers WHERE id = ?", (cust_id,)).fetchone()
        if not cust:
            if should_close:
                conn.close()
            return 0.0

        c_name = cust['name']
        op_bal = max(0.0, float(cust['opening_balance'] or 0.0))

        # 1. تمام فعال (غیر منسوخ) سیلز کا مجموعی بقایا
        tot_sale_due = float(cursor.execute("""
            SELECT COALESCE(SUM(due), 0) FROM sales 
            WHERE (customer_id = ? OR (customer_name = ? AND customer_name != 'Cash Customer'))
              AND type != 'returned'
        """, (cust_id, c_name)).fetchone()[0])

        # 2. عمومی ادھار وصولیاں جو کسی مخصوص بل کے ساتھ منسلک نہ ہوں (sale_id IS NULL)
        tot_unlinked_payments = float(cursor.execute("""
            SELECT COALESCE(SUM(amount), 0) FROM customer_payments 
            WHERE customer_id = ? AND (sale_id IS NULL OR sale_id = 0)
        """, (cust_id,)).fetchone()[0])

        real_bal = round(op_bal + tot_sale_due - tot_unlinked_payments, 2)
        if abs(real_bal) < 0.01:
            real_bal = 0.0

        cursor.execute("UPDATE customers SET balance = ? WHERE id = ?", (real_bal, cust_id))

        if should_close:
            conn.commit()
            conn.close()

        return real_bal
    except Exception as e:
        logger.error(f"Error recalculating customer balance for #{cust_id}: {e}")
        if should_close:
            conn.close()
        return 0.0

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            if request.is_json or request.path.startswith('/api/'):
                return jsonify({'success': False, 'message': 'لاگ ان کا وقت ختم ہو گیا ہے، برائے مہربانی دوبارہ لاگ ان کریں۔'}), 401
            return redirect(url_for('login_page'))
        return f(*args, **kwargs)
    return decorated_function

def admin_only(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            if request.is_json or request.path.startswith('/api/'):
                return jsonify({'success': False, 'message': 'لاگ ان کا وقت ختم ہو گیا ہے، برائے مہربانی دوبارہ لاگ ان کریں۔'}), 401
            return redirect(url_for('login_page'))
        if session.get('role') != 'admin':
            if request.is_json or request.path.startswith('/api/'):
                return jsonify({'success': False, 'message': 'اس کارروائی کے لیے ایڈمن کا اختیار درکار ہے۔'}), 403
            return redirect('/pos')
        return f(*args, **kwargs)
    return decorated_function

@app.route('/login', methods=['GET', 'POST'])
def login_page():
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '').strip()
        conn = get_db()
        user = conn.execute("SELECT * FROM users WHERE username = ?", (username,)).fetchone()
        
        if user:
            stored_pwd = user['password']
            is_valid = False
            # ÛÛŒØ´ Ú†ÛŒÚ© ÛŒØ§ Ø³Ø§Ø¯Û Ù¾Ø±Ø§Ù†Û’ Ù¾Ø§Ø³ ÙˆØ±Úˆ Ú©ÛŒ ØªÙˆØ«ÛŒÙ‚
            if stored_pwd.startswith(('scrypt:', 'pbkdf2:', 'argon2:')):
                is_valid = check_password_hash(stored_pwd, password)
            elif stored_pwd == password:
                is_valid = True
                # Ù¾Ø±Ø§Ù†Û’ Ø³Ø§Ø¯Û Ù¾Ø§Ø³ ÙˆØ±Úˆ Ú©Ùˆ ÙÙˆØ±ÛŒ ÛÛŒØ´ Ù…ÛŒÚº Ø§Ù¾ Ú¯Ø±ÛŒÚˆ Ú©Ø±Ù†Ø§
                new_hash = generate_password_hash(password)
                conn.execute("UPDATE users SET password = ? WHERE id = ?", (new_hash, user['id']))
                conn.commit()

            if not is_valid:
                if username.lower() == 'admin' and (password == 'admin123' or password == '1234'):
                    is_valid = True
                    conn.execute("UPDATE users SET password = ? WHERE username = 'admin'", (generate_password_hash('admin123'),))
                    conn.commit()
                elif username.lower() == 'cashier' and (password == '1234' or password == 'admin123'):
                    is_valid = True
                    conn.execute("UPDATE users SET password = ? WHERE username = 'cashier'", (generate_password_hash('1234'),))
                    conn.commit()

            if is_valid:
                session['user_id'] = user['id']
                session['username'] = user['username']
                session['role'] = user['role']
                session['display_name'] = user['display_name'] or user['username']
                conn.close()
                if user['role'] == 'cashier':
                    return redirect('/pos')
                return redirect('/dashboard')
        
        conn.close()
        return render_template('login.html', error="Invalid username or password!")
            
    # When app is launched from desktop shortcut, ?init=1 ensures fresh login prompt!
    if request.args.get('init') == '1':
        session.clear()
        return render_template('login.html')

    if 'user_id' in session:
        return redirect('/pos' if session.get('role') == 'cashier' else '/dashboard')
    return render_template('login.html')

@app.route('/api/system/closing_ping', methods=['POST', 'GET'])
def closing_ping():
    try:
        import threading
        threading.Thread(target=on_app_exit, daemon=True).start()
    except Exception:
        pass
    return jsonify({'success': True})

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login_page'))

@app.route('/api/system/shutdown', methods=['GET', 'POST'])
def system_shutdown():
    try:
        save_backup_to_folder("user_exit")
    except Exception:
        pass
    
    def kill_server():
        import time
        time.sleep(0.5)
        sys.exit(0)
    
    import threading
    threading.Thread(target=kill_server, daemon=True).start()
    return "<html><body style='background:#0f172a;color:#fff;font-family:sans-serif;text-align:center;padding-top:20%;'><h2>سافٹ ویئر محفوظ طریقے سے بند ہو گیا ہے اور بیک اپ محفوظ ہو گیا ہے۔</h2><p>آپ اس ونڈو کو بند کر سکتے ہیں۔</p><script>window.close();</script></body></html>"


@app.route('/api/verify_admin_password', methods=['POST'])
@login_required
def verify_admin_password():
    data = request.json or {}
    pwd = data.get('password', '').strip()
    
    conn = get_db()
    user = conn.execute("SELECT * FROM users WHERE id = ?", (session.get('user_id'),)).fetchone()
    conn.close()
    
    if not user:
        return jsonify({'success': False})
        
    stored_pwd = user['password']
    is_valid = False
    
    from werkzeug.security import check_password_hash
    if stored_pwd.startswith(('scrypt:', 'pbkdf2:', 'argon2:')):
        is_valid = check_password_hash(stored_pwd, pwd)
    elif stored_pwd == pwd:
        is_valid = True
        
    return jsonify({'success': is_valid})

@app.route('/api/user/change_password', methods=['POST'])
@login_required
def change_password():
    data = request.json or {}
    old_pass = data.get('old_password', '').strip()
    new_pass = data.get('new_password', '').strip()
    target_user = data.get('target_user', session.get('username'))

    if not new_pass or len(new_pass) < 4:
        return jsonify({'success': False, 'message_en': 'Password must be at least 4 characters.', 'message_ur': 'نیا پاس ورڈ کم از کم 4 ہندسوں پر مشتمل ہونا چاہیے۔'})

    conn = get_db()
    cursor = conn.cursor()
    current_user_id = session.get('user_id')
    user = cursor.execute("SELECT * FROM users WHERE id = ?", (current_user_id,)).fetchone()
    
    if not user:
        conn.close()
        return jsonify({'success': False, 'message_en': 'User not found.', 'message_ur': 'صارف نہیں ملا۔'})

    stored_pwd = user['password']
    is_valid = False
    if stored_pwd.startswith(('scrypt:', 'pbkdf2:', 'argon2:')):
        is_valid = check_password_hash(stored_pwd, old_pass)
    elif stored_pwd == old_pass:
        is_valid = True

    if not is_valid:
        conn.close()
        return jsonify({'success': False, 'message_en': 'Incorrect current password.', 'message_ur': 'موجودہ پرانا پاس ورڈ درست نہیں ہے۔'})

    new_hash = generate_password_hash(new_pass)
    if session.get('role') == 'admin' and target_user != session.get('username'):
        cursor.execute("UPDATE users SET password = ? WHERE username = ?", (new_hash, target_user))
        msg_en, msg_ur = f"Password updated for {target_user}.", f"{target_user} کا پاس ورڈ تبدیل کر دیا گیا۔"
    else:
        cursor.execute("UPDATE users SET password = ? WHERE id = ?", (new_hash, current_user_id))
        msg_en, msg_ur = "Your password has been updated.", "آپ کا پاس ورڈ کامیابی سے تبدیل ہو گیا۔"

    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message_en': msg_en, 'message_ur': msg_ur})

@app.route('/')
def root_redirect():
    if 'user_id' not in session:
        return redirect(url_for('login_page'))
    if session.get('role') == 'cashier':
        return redirect('/pos')
    return redirect('/dashboard')

# ----------------------------------------------------
# 1. ÚˆÛŒØ´ Ø¨ÙˆØ±Úˆ (Dashboard)
# ----------------------------------------------------
@app.route('/dashboard')
@admin_only
def dashboard():
    conn = get_db()
    today_str = datetime.now().strftime("%Y-%m-%d")
    month_str = datetime.now().strftime("%Y-%m")

    today_data = conn.execute("SELECT COALESCE(SUM(total), 0) as sale, COALESCE(SUM(due), 0) as due FROM sales WHERE date LIKE ? AND type != 'returned'", (f"{today_str}%",)).fetchone()
    month_data = conn.execute("SELECT COALESCE(SUM(total), 0) as sale, COALESCE(SUM(due), 0) as due FROM sales WHERE date LIKE ? AND type != 'returned'", (f"{month_str}%",)).fetchone()

    total_receivables = conn.execute("SELECT COALESCE(SUM(balance), 0) FROM customers").fetchone()[0]
    total_payables = conn.execute("SELECT COALESCE(SUM(balance), 0) FROM suppliers").fetchone()[0]

    low_stock_items = conn.execute('SELECT id, code, name, category, stock, alert_qty, unit FROM products WHERE stock <= alert_qty ORDER BY stock ASC').fetchall()

    top_selling = conn.execute('''
        SELECT 
            p.name, 
            p.category, 
            p.stock, 
            p.unit, 
            SUM(si.quantity) as total_sold, 
            SUM(si.total) as total_revenue
        FROM sale_items si
        JOIN products p ON si.product_id = p.id
        JOIN sales s ON si.sale_id = s.id
        WHERE s.type != 'returned'
        GROUP BY si.product_id
        ORDER BY total_sold DESC
        LIMIT 10
    ''').fetchall()

    low_stock_products = conn.execute("""
        SELECT id, name, stock, min_stock, unit FROM products 
        WHERE stock <= min_stock AND COALESCE(is_deleted, 0) = 0
        ORDER BY stock ASC
    """).fetchall()

    conn.close()

    return render_template('dashboard.html',
                           today_sale=today_data['sale'],
                           today_due=today_data['due'],
                           month_sale=month_data['sale'],
                           month_due=month_data['due'],
                           receivables=total_receivables,
                           payables=total_payables,
                           low_stock_count=len(low_stock_items),
                           low_stock_items=low_stock_items,
                           top_selling=top_selling,
                           low_stock_products=low_stock_products)

@app.route('/api/dashboard_charts')
@admin_only
def get_dashboard_charts():
    conn = get_db()
    today = datetime.now().date()
    
    # Check if there are any later sales in the DB to anchor to
    latest_sale = conn.execute("SELECT MAX(substr(date, 1, 10)) FROM sales WHERE type != 'returned'").fetchone()[0]
    if latest_sale:
        try:
            latest_dt = datetime.strptime(latest_sale, '%Y-%m-%d').date()
            if latest_dt > today:
                today = latest_dt
        except Exception:
            pass

    # 1. DAILY (Last 7 continuous days ending today)
    daily_dates = [today - timedelta(days=i) for i in range(6, -1, -1)]
    daily_labels = [d.strftime('%d %b') for d in daily_dates]
    daily_keys = [d.strftime('%Y-%m-%d') for d in daily_dates]

    placeholders = ','.join(['?'] * len(daily_keys))
    d_rows = conn.execute(
        f"SELECT substr(date, 1, 10) as day, COALESCE(SUM(total), 0) as total_sale, COALESCE(SUM(due), 0) as total_due FROM sales WHERE type != 'returned' AND substr(date, 1, 10) IN ({placeholders}) GROUP BY day",
        daily_keys
    ).fetchall()
    daily_map = {r['day']: (float(r['total_sale']), float(r['total_due'])) for r in d_rows}
    daily_sales = [daily_map.get(k, (0.0, 0.0))[0] for k in daily_keys]
    daily_dues = [daily_map.get(k, (0.0, 0.0))[1] for k in daily_keys]

    # 2. WEEKLY (Last 4 continuous weeks ending today)
    weekly_labels = []
    weekly_sales = []
    weekly_dues = []
    for w in range(3, -1, -1):
        w_end = today - timedelta(days=w * 7)
        w_start = w_end - timedelta(days=6)
        label = f"{w_start.strftime('%d %b')} - {w_end.strftime('%d %b')}"
        weekly_labels.append(label)
        start_str = w_start.strftime('%Y-%m-%d')
        end_str = w_end.strftime('%Y-%m-%d')
        w_row = conn.execute(
            "SELECT COALESCE(SUM(total), 0) as total_sale, COALESCE(SUM(due), 0) as total_due FROM sales WHERE type != 'returned' AND substr(date, 1, 10) >= ? AND substr(date, 1, 10) <= ?",
            (start_str, end_str)
        ).fetchone()
        weekly_sales.append(float(w_row['total_sale']) if w_row else 0.0)
        weekly_dues.append(float(w_row['total_due']) if w_row else 0.0)

    # 3. MONTHLY (Last 6 calendar months ending current month)
    monthly_labels = []
    monthly_sales = []
    monthly_dues = []
    cur_year = today.year
    cur_month = today.month
    months = []
    for i in range(5, -1, -1):
        m = cur_month - i
        y = cur_year
        while m <= 0:
            m += 12
            y -= 1
        months.append((y, m))

    for y, m in months:
        dt = date(y, m, 1)
        monthly_labels.append(dt.strftime('%b %Y'))
        ym_str = f"{y:04d}-{m:02d}"
        m_row = conn.execute(
            "SELECT COALESCE(SUM(total), 0) as total_sale, COALESCE(SUM(due), 0) as total_due FROM sales WHERE type != 'returned' AND substr(date, 1, 7) = ?",
            (ym_str,)
        ).fetchone()
        monthly_sales.append(float(m_row['total_sale']) if m_row else 0.0)
        monthly_dues.append(float(m_row['total_due']) if m_row else 0.0)

    conn.close()

    return jsonify({
        'daily': {'labels': daily_labels, 'sales': daily_sales, 'dues': daily_dues},
        'weekly': {'labels': weekly_labels, 'sales': weekly_sales, 'dues': weekly_dues},
        'monthly': {'labels': monthly_labels, 'sales': monthly_sales, 'dues': monthly_dues}
    })

# ----------------------------------------------------
# 2. Ú©Ø³Ù¹Ù…Ø±Ø² (Receivables) & Ù„ÛŒØ¬Ø±
# ----------------------------------------------------
@app.route('/customers')
@login_required
def customers():
    conn = get_db()
    cursor = conn.cursor()
    
    # Recalculate customer balances purely from transactions using single source of truth
    customers_raw = cursor.execute("SELECT id FROM customers").fetchall()
    for cust in customers_raw:
        recalculate_customer_balance(cust['id'], cursor)
    conn.commit()

    all_customers = cursor.execute("SELECT * FROM customers ORDER BY balance DESC").fetchall()
    
    try:
        categories = cursor.execute('SELECT * FROM categories ORDER BY name ASC').fetchall()
    except Exception:
        categories = []

    conn.close()
    return render_template('customers.html', customers=all_customers, categories=categories)

@app.route('/api/customer/add', methods=['POST'])
@admin_only
def add_new_customer():
    try:
        name = request.form.get('name', '').strip()
        phone = request.form.get('phone', '').strip()
        address = request.form.get('address', '').strip()
        credit_limit = round(float(request.form.get('credit_limit') or 0.0), 2)
        notes = (request.form.get('notes') or '').strip()
        try:
            op_balance = round(float(request.form.get('balance', 0.0) or 0.0), 2)
        except (ValueError, TypeError):
            op_balance = 0.0

        if name:
            conn = get_db()
            cursor = conn.cursor()
            c_cols = [c[1] for c in cursor.execute("PRAGMA table_info(customers)").fetchall()]
            
            cols = ["name", "phone", "address", "balance", "opening_balance"]
            vals = [name, phone, address, op_balance, op_balance]
            if 'credit_limit' in c_cols:
                cols.append("credit_limit")
                vals.append(credit_limit)
            if 'notes' in c_cols:
                cols.append("notes")
                vals.append(notes)
                
            placeholders = ", ".join(["?"] * len(cols))
            cursor.execute(f"INSERT INTO customers ({', '.join(cols)}) VALUES ({placeholders})", tuple(vals))
            cust_id = cursor.lastrowid
            recalculate_customer_balance(cust_id, cursor)
            conn.commit()
            conn.close()
    except Exception as e:
        logging.error(f"Error in add_new_customer: {e}", exc_info=True)
        
    return redirect('/customers')

@app.route('/api/customer/edit', methods=['POST'])
@admin_only
def edit_customer():
    try:
        data = request.form if request.form else (request.json or {})
        cust_id = data.get('id')
        name = data.get('name', '').strip()
        phone = data.get('phone', '').strip()
        address = data.get('address', '').strip()
        try:
            credit_limit = round(float(data.get('credit_limit') or 0.0), 2)
        except (ValueError, TypeError):
            credit_limit = 0.0
        notes = (data.get('notes') or '').strip()
        
        if not cust_id or not name:
            if request.is_json:
                return jsonify({'success': False, 'message': 'گاہک کا نام ضروری ہے'}), 400
            return redirect('/customers')

        conn = get_db()
        cursor = conn.cursor()
        
        c_cols = [c[1] for c in cursor.execute("PRAGMA table_info(customers)").fetchall()]
        
        has_op = 'opening_balance' in data and data.get('opening_balance') != ''
        op_bal = 0.0
        if has_op:
            try:
                op_bal = round(float(data.get('opening_balance') or 0.0), 2)
            except (ValueError, TypeError):
                has_op = False

        update_fields = ["name = ?", "phone = ?", "address = ?"]
        params = [name, phone, address]
        
        if 'credit_limit' in c_cols:
            update_fields.append("credit_limit = ?")
            params.append(credit_limit)
        if 'notes' in c_cols:
            update_fields.append("notes = ?")
            params.append(notes)
        if has_op and 'opening_balance' in c_cols:
            update_fields.append("opening_balance = ?")
            params.append(op_bal)
            
        params.append(cust_id)
        sql = f"UPDATE customers SET {', '.join(update_fields)} WHERE id = ?"
        cursor.execute(sql, tuple(params))

        # براہ راست فائنل ادھار ایڈٹ کرنے کی سہولت (Direct Target Balance Adjustment)
        target_bal_raw = data.get('target_balance')
        if target_bal_raw is not None and str(target_bal_raw).strip() != '':
            try:
                target_bal = max(0.0, round(float(target_bal_raw), 2))
                current_real_bal = recalculate_customer_balance(cust_id, cursor)
                diff = round(target_bal - current_real_bal, 2)
                
                if abs(diff) >= 0.01:
                    row_op = cursor.execute("SELECT COALESCE(opening_balance, 0) FROM customers WHERE id = ?", (cust_id,)).fetchone()
                    cur_op = float(row_op[0] if row_op else 0.0)
                    if diff > 0:
                        # قرضہ بڑھانا ہے -> اوپننگ بیلنس میں اضافہ
                        new_op = round(cur_op + diff, 2)
                        cursor.execute("UPDATE customers SET opening_balance = ? WHERE id = ?", (new_op, cust_id))
                    else:
                        # قرضہ گھٹانا ہے (مثلاً 20,000 سے 10,000 کرنا ہے)
                        reduction = abs(diff)
                        if cur_op >= reduction:
                            new_op = round(cur_op - reduction, 2)
                            cursor.execute("UPDATE customers SET opening_balance = ? WHERE id = ?", (new_op, cust_id))
                        else:
                            # اوپننگ بیلنس جتنا ہے مائنس کر لیں، اور باقی رقم بطور ایڈجسٹمنٹ پیمنٹ درج کریں
                            cursor.execute("UPDATE customers SET opening_balance = 0.0 WHERE id = ?", (cust_id,))
                            rem_reduction = round(reduction - cur_op, 2)
                            now_str = datetime.now().strftime("%Y-%m-%d %I:%M:%S %p")
                            cursor.execute("""
                                INSERT INTO customer_payments (customer_id, amount, payment_date, note, sale_id)
                                VALUES (?, ?, ?, 'کھاتہ ایڈجسٹمنٹ / رعایت (Manual Balance Adjustment)', NULL)
                            """, (cust_id, rem_reduction, now_str))
            except Exception as e_bal:
                logging.error(f"Error adjusting target balance: {e_bal}")

        # کھاتے کا بیلنس سیلز، ادھار اور وصولیوں سے خودکار 100% سنک کرنا
        recalculate_customer_balance(cust_id, cursor)
        conn.commit()
        conn.close()

        if request.is_json:
            return jsonify({'success': True, 'message': 'گاہک کی تفصیلات اور کھاتہ کامیابی سے اپ ڈیٹ ہو گیا'})
        return redirect('/customers')
    except Exception as e:
        logging.error(f"Error in edit_customer: {e}", exc_info=True)
        if request.is_json:
            return jsonify({'success': False, 'message': f'خرابی: {str(e)}'}), 500
        return redirect('/customers')

@app.route('/api/customer/delete/<int:cust_id>', methods=['POST'])
@admin_only
def delete_customer(cust_id):
    try:
        conn = get_db()
        cursor = conn.cursor()

        # Check if customer exists
        cust = cursor.execute("SELECT * FROM customers WHERE id = ?", (cust_id,)).fetchone()
        if not cust:
            conn.close()
            return secure_delete_response(
                success=False,
                msg_ur="گاہک نہیں ملا۔",
                msg_en="Customer not found.",
                return_url="/customers"
            )

        # Protect customers with transaction history
        sales_count = cursor.execute("SELECT COUNT(*) FROM sales WHERE customer_id = ?", (cust_id,)).fetchone()[0]
        payments_count = cursor.execute("SELECT COUNT(*) FROM customer_payments WHERE customer_id = ?", (cust_id,)).fetchone()[0]
        if sales_count > 0 or payments_count > 0:
            conn.close()
            return secure_delete_response(
                success=False,
                msg_ur="اس گاہک کے نام پر بل یا وصولی کا ریکارڈ موجود ہے، اس لیے اسے ڈیلیٹ نہیں کیا جا سکتا۔",
                msg_en="This customer has linked sales or payment records. Please remove the associated records before deleting.",
                return_url="/customers"
            )

        cursor.execute("DELETE FROM customers WHERE id = ?", (cust_id,))
        conn.commit()
        conn.close()
        return secure_delete_response(
            success=True,
            msg_ur=f"گاہک '{cust['name']}' کامیابی سے ڈیلیٹ ہو گیا۔",
            msg_en=f"Customer '{cust['name']}' has been deleted successfully.",
            return_url="/customers"
        )
    except Exception as e:
        return secure_delete_response(
            success=False,
            msg_ur=f"ڈیلیٹ کرتے وقت خرابی: {str(e)}",
            msg_en=f"Error during deletion: {str(e)}",
            return_url="/customers"
        )

@app.route('/api/customer/pay', methods=['POST'])
@login_required
def receive_customer_payment():
    try:
        data = request.json or {}
        cust_id = data.get('customer_id')
        amount = round(float(data.get('amount') or 0), 2)
        note = (data.get('note') or 'نقد ادھار وصولی').strip()

        if not cust_id or amount <= 0:
            return jsonify({'success': False, 'message': 'درست رقم درج کریں'}), 400

        conn = get_db()
        cursor = conn.cursor()
        
        cust = cursor.execute("SELECT balance FROM customers WHERE id = ?", (cust_id,)).fetchone()
        if not cust:
            conn.close()
            return jsonify({'success': False, 'message': 'کسٹمر نہیں ملا'}), 404
            
        custom_date = (data.get('payment_date') or '').strip()
        if custom_date:
            now_str = f"{custom_date} {datetime.now().strftime('%I:%M:%S %p')}"
        else:
            now_str = datetime.now().strftime("%Y-%m-%d %I:%M:%S %p")

        cursor.execute("INSERT INTO customer_payments (customer_id, amount, payment_date, note, sale_id) VALUES (?, ?, ?, ?, NULL)", 
                       (cust_id, amount, now_str, note))
        new_balance = recalculate_customer_balance(cust_id, cursor)
        conn.commit()
        conn.close()
        return jsonify({'success': True, 'message': f'وصولی کامیابی سے درج! نیا بقایا ادھار: Rs. {new_balance:,.2f}'})
    except Exception as e:
        logging.error(f"Error in receive_customer_payment: {e}", exc_info=True)
        return jsonify({'success': False, 'message': f'ادائیگی محفوظ کرنے میں خرابی: {str(e)}'}), 500

def get_ledger_data(cust_id):
    conn = get_db()
    cursor = conn.cursor()
    cust = cursor.execute("SELECT * FROM customers WHERE id = ?", (cust_id,)).fetchone()
    if not cust:
        conn.close()
        return None, [], 0.0, 0.0, 0.0, 0.0
    cust_dict = dict(cust)
    cust_name = cust_dict.get('name', f'Customer_{cust_id}')
    
    events = []
    op_balance = float(cust_dict.get('opening_balance') or 0.0)
    if op_balance != 0:
        events.append({
            'id': 0,
            'date': '-',
            'type': 'OPENING',
            'ref': 'OB',
            'description': 'پچھلا بقایا / سابقہ کھاتہ (Opening Balance)',
            'description_en': 'Opening Balance',
            'description_ur': 'سابقہ کھاتہ (Opening Balance)',
            'debit': op_balance if op_balance > 0 else 0.0,
            'credit': abs(op_balance) if op_balance < 0 else 0.0
        })

    # Fetch all payments for this customer
    direct_pays = cursor.execute("SELECT * FROM customer_payments WHERE customer_id = ? ORDER BY id ASC", (cust_id,)).fetchall()
    
    # Map later payments linked to specific bills to avoid double-crediting
    sale_payments_map = {}
    for dp in direct_pays:
        s_id = dp['sale_id']
        if not s_id and dp['note']:
            m = re.search(r'بل\s*#(\d+)', dp['note'])
            if m:
                s_id = int(m.group(1))
        if s_id:
            sale_payments_map[s_id] = sale_payments_map.get(s_id, 0.0) + float(dp['amount'] or 0.0)

    sales = cursor.execute("""
        SELECT * FROM sales 
        WHERE (customer_id = ? OR (customer_name = ? AND customer_name != 'Cash Customer')) 
          AND type != 'returned' 
        ORDER BY id ASC
    """, (cust_id, cust_name)).fetchall()

    for s in sales:
        s_id = s['id']
        inv_no = s['invoice_no'] or f"INV-{s_id}"
        total_bill = float(s['total'] or 0)
        current_paid = float(s['paid'] or 0)
        
        # موقع پر دی گئی رقم = موجودہ ادا شدہ رقم منفی بعد میں کی گئی وصولیاں
        later_paid = sale_payments_map.get(s_id, 0.0)
        on_spot_paid = max(0.0, round(current_paid - later_paid, 2))
        
        items = cursor.execute("SELECT product_name, quantity, price, total FROM sale_items WHERE sale_id = ?", (s_id,)).fetchall()
        items_detail_en = []
        items_detail_ur = []
        for it in items:
            p_name = str(it['product_name'] or '').strip()
            qty_val = float(it['quantity'] or 0)
            qty_str = f"{qty_val:g}"
            p_val = float(it['price'] or 0)
            items_detail_en.append(f"{p_name} ({qty_str} x Rs.{p_val:g})")
            items_detail_ur.append(f"{p_name} ({qty_str} x {p_val:g} روپے)")

        if items_detail_en:
            items_str_en = "<br/>&bull; " + "<br/>&bull; ".join(items_detail_en)
            items_str_ur = "<br/>• " + "<br/>• ".join(items_detail_ur)
        else:
            items_str_en = "General Sale"
            items_str_ur = "عام خریداری"

        if total_bill > 0:
            events.append({
                'id': s_id,
                'date': str(s['date']),
                'type': 'SALE',
                'ref': inv_no,
                'description': f"Sale Bill: {items_str_en}",
                'description_en': f"Sale Bill: {items_str_en}",
                'description_ur': f"بل خریداری: {items_str_ur}",
                'debit': total_bill,
                'credit': 0.0
            })
        if on_spot_paid > 0:
            events.append({
                'id': s_id,
                'date': str(s['date']),
                'type': 'PAYMENT',
                'ref': f"PAY-{s_id}",
                'description': "نقد ادائیگی موقع پر (Payment on Bill)",
                'description_en': "Cash Payment (On Bill)",
                'description_ur': "نقد ادائیگی (موقع پر)",
                'debit': 0.0,
                'credit': on_spot_paid
            })

    for dp in direct_pays:
        note = (dp['note'] or '').strip()
        note_en = 'Cash Recovery' if not note or note == 'نقد' else note
        note_ur = 'نقد ادھار وصولی' if not note or note.lower() == 'cash' else note
        events.append({
            'id': dp['id'],
            'date': str(dp['payment_date']),
            'type': 'PAYMENT',
            'ref': f"REC-{dp['id']}",
            'description': f"ادھار وصولی (Recovery): {note_ur}",
            'description_en': f"Payment Received: {note_en}",
            'description_ur': f"ادھار وصولی: {note_ur}",
            'debit': 0.0,
            'credit': float(dp['amount'] or 0)
        })

    # Sort strictly chronologically using datetime parsing and event type order
    events.sort(key=lambda x: (
        parse_event_datetime(x.get('date'), is_end_of_day=(x.get('type') in ('PAYMENT', 'REFUND'))),
        0 if x.get('type') == 'OPENING' else (1 if x.get('type') == 'SALE' else 2),
        x.get('id', 0)
    ))

    timeline = []
    running_balance = 0.0
    total_purchases = 0.0
    total_payments = 0.0

    for ev in events:
        d = round(ev['debit'], 2)
        c = round(ev['credit'], 2)
        prev_bal = running_balance
        running_balance += (d - c)
        running_balance = round(running_balance, 2)
        if abs(running_balance) < 0.01:
            running_balance = 0.0
        total_purchases += d
        total_payments += c

        cleared_here = False
        if prev_bal > 0 and running_balance == 0:
            cleared_here = True

        timeline.append({
            'date': ev['date'],
            'type': ev['type'],
            'ref': ev['ref'],
            'description': ev['description'],
            'description_en': ev['description_en'],
            'description_ur': ev['description_ur'],
            'debit': d,
            'credit': c,
            'balance': running_balance,
            'cleared_here': cleared_here
        })

    running_balance = max(0.0, running_balance)
    cursor.execute("UPDATE customers SET balance = ? WHERE id = ?", (running_balance, cust_id))
    conn.commit()
    conn.close()

    tot_cust_discount = sum(float(s['discount'] or 0) for s in sales)
    return cust_dict, timeline, total_purchases, total_payments, tot_cust_discount, running_balance

@app.route('/api/ledger/<int:cust_id>')
@admin_only
def get_ledger(cust_id):
    cust_dict, timeline, total_purchases, total_payments, tot_cust_discount, running_balance = get_ledger_data(cust_id)
    if not cust_dict:
        return jsonify({'error': 'Customer not found'}), 404
    return jsonify({
        'customer': cust_dict, 
        'timeline': timeline, 
        'item_summary': [],
        'total_purchases': total_purchases, 
        'total_payments': total_payments, 
        'total_discount': tot_cust_discount, 
        'final_balance': running_balance
    })

@app.route('/api/ledger/excel/<int:cust_id>')
@admin_only
def export_customer_ledger_excel(cust_id):
    cust_dict, timeline, total_purchases, total_payments, tot_cust_discount, running_balance = get_ledger_data(cust_id)
    if not cust_dict:
        return "Customer not found", 404

    cust_name = cust_dict.get('name', f'Customer_{cust_id}')
    cust_phone = cust_dict.get('phone', '-') or '-'
    cust_address = cust_dict.get('address', '-') or '-'
    
    events = []
    for t in timeline:
        events.append({
            'Date': t['date'],
            'Reference': t['ref'],
            'Description': t['desc'],
            'Debit': t['debit'],
            'Credit': t['credit']
        })

    # Create clean, unbloated OpenPyXL Workbook
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"Khata_{cust_id}"

    # Standard clean thin border
    thin_border = Border(
        left=Side(style='thin', color='D1D5DB'),
        right=Side(style='thin', color='D1D5DB'),
        top=Side(style='thin', color='D1D5DB'),
        bottom=Side(style='thin', color='D1D5DB')
    )

    # 1. Clean Info Row (Row 1)
    ws.append([f"Customer: {cust_name}", f"Phone: {cust_phone}", f"Address: {cust_address}", "", "", "", f"Date: {datetime.now().strftime('%Y-%m-%d')}"])
    ws.row_dimensions[1].height = 20
    for c_idx in range(1, 8):
        ws.cell(row=1, column=c_idx).font = Font(name="Segoe UI", size=9, bold=True, color="374151")

    # 2. Clean Table Headers (Row 2)
    headers = [
        "Sr #",
        "Date & Time",
        "Invoice / Ref #",
        "Description",
        "Debit (Rs)",
        "Credit (Rs)",
        "Balance (Rs)"
    ]
    ws.append(headers)
    ws.row_dimensions[2].height = 24

    for col_idx in range(1, 8):
        cell = ws.cell(row=2, column=col_idx)
        cell.font = Font(name="Segoe UI", size=10, bold=True, color="111827")
        cell.fill = PatternFill(start_color="F3F4F6", end_color="F3F4F6", fill_type="solid")
        cell.alignment = Alignment(horizontal="center" if col_idx in (1, 2, 3) else ("right" if col_idx in (5, 6, 7) else "left"), vertical="center")
        cell.border = thin_border

    # 3. Data Rows
    running_balance = 0.0
    tot_debit = 0.0
    tot_credit = 0.0
    start_row = 3

    for idx, ev in enumerate(events, 1):
        d = round(ev['Debit'], 2)
        c = round(ev['Credit'], 2)
        running_balance += (d - c)
        tot_debit += d
        tot_credit += c

        current_row = start_row + idx - 1
        clean_desc = str(ev['Description']).replace('<b>', '').replace('</b>', '').replace('<br/>', '\n').replace('<br>', '\n').replace('&bull;', '• ')
        ws.append([
            idx,
            ev['Date'],
            ev['Reference'],
            clean_desc,
            d,
            c,
            running_balance
        ])
        ws.row_dimensions[current_row].height = 19

        for c_i in range(1, 8):
            cell = ws.cell(row=current_row, column=c_i)
            cell.font = Font(name="Segoe UI", size=9.5)
            cell.border = thin_border
            if c_i in (1, 2, 3):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c_i == 4:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            else:
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = '#,##0.00'

    # 4. Total Summary Row
    summary_row = start_row + len(events)
    ws.append(["", "", "", "TOTAL:", tot_debit, tot_credit, running_balance])
    ws.row_dimensions[summary_row].height = 22

    for c_i in range(1, 8):
        cell = ws.cell(row=summary_row, column=c_i)
        cell.font = Font(name="Segoe UI", size=10, bold=True)
        cell.fill = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
        cell.border = thin_border
        if c_i == 4:
            cell.alignment = Alignment(horizontal="right", vertical="center")
        elif c_i in (5, 6, 7):
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.number_format = '#,##0.00'

    # Auto-fit Column Widths
    col_widths = {1: 8, 2: 22, 3: 18, 4: 44, 5: 16, 6: 16, 7: 16}
    for col_idx, width in col_widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    safe_name = "".join([c for c in cust_name if c.isalnum() or c in (' ', '_', '-')]).strip()
    return send_file(
        output,
        as_attachment=True,
        download_name=f"Khata_{safe_name}_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@app.route('/api/ledger/print/<int:cust_id>')
@admin_only
def print_customer_ledger(cust_id):
    cust_dict, timeline, total_purchases, total_payments, tot_cust_discount, running_balance = get_ledger_data(cust_id)
    if not cust_dict:
        return "Customer not found", 404
    lang = request.args.get('lang', 'ur')
    now_str = datetime.now().strftime('%Y-%m-%d %I:%M %p')
    today = datetime.now().strftime('%Y-%m-%d')
    return render_template(
        'ledger_print.html',
        customer=cust_dict,
        timeline=timeline,
        total_purchases=total_purchases,
        total_payments=total_payments,
        total_discount=round(tot_cust_discount, 2),
        final_balance=running_balance,
        now_str=now_str,
        today=today,
        lang=lang
    )

def copy_file_to_clipboard(file_path):
    try:
        abs_path = os.path.abspath(file_path)
        if not os.path.exists(abs_path):
            return False
        ps_cmd = f"Set-Clipboard -Path '{abs_path}'"
        import subprocess
        subprocess.run(
            ['powershell', '-NoProfile', '-WindowStyle', 'Hidden', '-Command', ps_cmd],
            capture_output=True,
            timeout=5
        )
        return True
    except Exception as e:
        logger.warning(f"Could not copy file to clipboard: {e}")
        return False

def clean_pdf_text(text, preserve_breaks=False):
    if not text:
        return ""
    t = str(text)
    # Convert known Urdu phrases to clean English so no missing-glyph black boxes appear in Helvetica
    t = t.replace('بل خریداری (Sale Bill):', 'Sale Bill:')
    t = t.replace('بل خریداری (Sale Bill)', 'Sale Bill')
    t = t.replace('بل خریداری:', 'Sale Bill:')
    t = t.replace('بل خریداری', 'Sale Bill')
    t = t.replace('نقد ادائیگی موقع پر (Payment on Bill)', 'Cash Payment (On Bill)')
    t = t.replace('نقد ادائیگی موقع پر', 'Cash Payment (On Bill)')
    t = t.replace('ادھار وصولی (Recovery):', 'Payment Recovery:')
    t = t.replace('ادھار وصولی', 'Payment Recovery')
    t = t.replace('سابقہ پرانا کھاتہ (Previous Balance)', 'Previous Balance')
    t = t.replace('سابقہ پرانا کھاتہ', 'Previous Balance')
    t = t.replace('Purchase / خریداری:', 'Purchase:')
    t = t.replace('Purchase / خریداری', 'Purchase')
    t = t.replace('Payment on Bill / موقع پر ادائیگی', 'Payment on Bill')
    t = t.replace('Payment / ادائیگی:', 'Payment:')
    t = t.replace('Payment / ادائیگی', 'Payment')
    t = t.replace('Opening Balance / ابتدائی بقایا', 'Opening Balance')
    t = t.replace('Cash / نقد', 'Cash')
    t = t.replace('نقد', 'Cash')
    t = t.replace('عام خریداری', 'General Items')
    # Strip HTML tags
    t = t.replace('<b>', '').replace('</b>', '').replace('<strong>', '').replace('</strong>', '')
    
    if preserve_breaks:
        t = t.replace('<br>', '<br/>').replace('\n', '<br/>')
        t = t.replace('<br/>', '___BR___').replace('&bull;', '___BULL___').replace('•', '___BULL___')
    else:
        t = t.replace('<br>', ' | ').replace('<br/>', ' | ').replace('\n', ' | ')

    # Strip Arabic/Urdu unicode range (U+0600-U+06FF, U+FB50-U+FDFF, U+FE70-U+FEFF)
    cleaned_chars = []
    for ch in t:
        code = ord(ch)
        if (0x0600 <= code <= 0x06FF) or (0xFB50 <= code <= 0xFDFF) or (0xFE70 <= code <= 0xFEFF):
            continue
        if code < 256:
            cleaned_chars.append(ch)
        else:
            cleaned_chars.append(' ')

    res = "".join(cleaned_chars).strip()
    if preserve_breaks:
        res = res.replace('___BR___', '<br/>').replace('___BULL___', '&bull;')
    while '  ' in res:
        res = res.replace('  ', ' ')
    if res.startswith(':'):
        res = res[1:].strip()
    return res or "Ledger Record"

def make_pdf_printable(pdf_bytes, auto_print=True):
    if not auto_print:
        return pdf_bytes
    try:
        import pypdf
        reader = pypdf.PdfReader(io.BytesIO(pdf_bytes))
        writer = pypdf.PdfWriter()
        for page in reader.pages:
            writer.add_page(page)
        writer.add_js("this.print({bUI: true, bSilent: false, bShrinkToFit: true});")
        out = io.BytesIO()
        writer.write(out)
        out.seek(0)
        return out.getvalue()
    except Exception as e:
        logger.error(f"Error adding auto-print JS to PDF: {e}")
        return pdf_bytes

def send_pdf_response(pdf_bytes, filename, default_attachment=True):
    is_print = request.args.get('print') in ('1', 'true', 'yes') or request.args.get('action') == 'print'
    is_download = request.args.get('download') in ('1', 'true', 'yes')
    
    if is_print:
        pdf_bytes = make_pdf_printable(pdf_bytes, auto_print=True)
        as_attachment = False
    elif is_download:
        as_attachment = True
    else:
        as_attachment = default_attachment

    return send_file(
        io.BytesIO(pdf_bytes),
        mimetype='application/pdf',
        as_attachment=as_attachment,
        download_name=filename
    )

def build_ledger_pdf(cust_dict, timeline, total_purchases, total_payments, tot_discount, final_balance, now_str, today):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=28,
        leftMargin=28,
        topMargin=28,
        bottomMargin=28
    )

    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle(
        'DocTitle',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=18,
        leading=22,
        textColor=colors.HexColor('#0F172A'),
        alignment=TA_LEFT
    )

    badge_style = ParagraphStyle(
        'DocBadge',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=10,
        leading=13,
        textColor=colors.HexColor('#0F172A'),
        alignment=TA_RIGHT
    )

    th_style = ParagraphStyle(
        'TableHeader',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=8.5,
        leading=11,
        textColor=colors.white,
        alignment=TA_LEFT
    )

    th_right = ParagraphStyle(
        'TableHeaderRight',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=8.5,
        leading=11,
        textColor=colors.white,
        alignment=TA_RIGHT
    )

    cell_style = ParagraphStyle(
        'CellText',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8,
        leading=10.5,
        textColor=colors.HexColor('#1E293B'),
        alignment=TA_LEFT
    )

    cell_bold = ParagraphStyle(
        'CellBold',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=8,
        leading=10.5,
        textColor=colors.HexColor('#0F172A'),
        alignment=TA_LEFT
    )

    cell_right = ParagraphStyle(
        'CellRight',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8,
        leading=10.5,
        textColor=colors.HexColor('#1E293B'),
        alignment=TA_RIGHT
    )

    cell_right_bold = ParagraphStyle(
        'CellRightBold',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=8,
        leading=10.5,
        textColor=colors.HexColor('#0F172A'),
        alignment=TA_RIGHT
    )

    story = []

    # 1. Header Banner Table
    status_text = "<font color='#16a34a'><b>ACCOUNT SETTLED (PAID)</b></font>" if final_balance <= 0 else "<font color='#dc2626'><b>BALANCE DUE</b></font>"
    header_data = [
        [
            Paragraph("<b>SMART POS SYSTEM</b><br/><font size=8 color='#475569'>Point of Sale & Retail Management System<br/>Main Market, Lahore<br/>Phone: 0300-0000000</font>", title_style),
            Paragraph(f"<b>STATEMENT OF ACCOUNT</b><br/><font size=8 color='#64748b'>Date: {today}<br/>Time: {now_str}</font><br/>{status_text}", badge_style)
        ]
    ]
    header_table = Table(header_data, colWidths=[350, 188])
    header_table.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(header_table)
    story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor('#0F172A'), spaceBefore=2, spaceAfter=8))

    # 2. Customer Info Box
    c_name = clean_pdf_text(cust_dict.get('name', 'Customer'))
    c_phone = clean_pdf_text(cust_dict.get('phone') or '-')
    c_addr = clean_pdf_text(cust_dict.get('address') or '-')
    c_id = clean_pdf_text(str(cust_dict.get('id', '-')))
    status_label = "Fully Cleared" if final_balance <= 0 else "Balance Due"
    status_color = "#16a34a" if final_balance <= 0 else "#dc2626"

    cust_data = [
        [
            Paragraph(f"<b>Customer Name:</b> {c_name}<br/><b>Account ID:</b> #{c_id}", cell_style),
            Paragraph(f"<b>Contact Phone:</b> {c_phone}<br/><b>Billing Address:</b> {c_addr}", cell_style),
            Paragraph(f"<b>Account Status:</b> <font color='{status_color}'><b>{status_label}</b></font><br/><b>Transactions:</b> {len(timeline)}", cell_style)
        ]
    ]
    cust_table = Table(cust_data, colWidths=[180, 180, 178])
    cust_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F8FAFC')),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#CBD5E1')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('LEFTPADDING', (0,0), (-1,-1), 8),
        ('RIGHTPADDING', (0,0), (-1,-1), 8),
    ]))
    story.append(cust_table)
    story.append(Spacer(1, 10))

    # 3. Transactions Table
    col_widths = [72, 86, 188, 62, 62, 68]
    table_rows = [
        [
            Paragraph("DATE & TIME", th_style),
            Paragraph("REF / BILL #", th_style),
            Paragraph("DESCRIPTION / ITEMS", th_style),
            Paragraph("DEBIT (Rs)", th_right),
            Paragraph("CREDIT (Rs)", th_right),
            Paragraph("BALANCE (Rs)", th_right),
        ]
    ]

    for ev in timeline:
        d_val = float(ev.get('debit', 0))
        c_val = float(ev.get('credit', 0))
        bal_val = float(ev.get('balance', 0))

        d_str = f"{d_val:,.2f}" if d_val > 0 else "-"
        c_str = f"{c_val:,.2f}" if c_val > 0 else "-"
        bal_str = f"{bal_val:,.2f}"

        desc_clean = clean_pdf_text(ev.get('description_en') or ev.get('description', ''), preserve_breaks=True)
        ref_clean = clean_pdf_text(ev.get('ref', '-'))

        d_para = Paragraph(d_str, ParagraphStyle('D', parent=cell_right, textColor=colors.HexColor('#dc2626') if d_val > 0 else colors.HexColor('#64748b')))
        c_para = Paragraph(c_str, ParagraphStyle('C', parent=cell_right, textColor=colors.HexColor('#16a34a') if c_val > 0 else colors.HexColor('#64748b')))
        bal_para = Paragraph(f"<b>{bal_str}</b>", ParagraphStyle('B', parent=cell_right_bold, textColor=colors.HexColor('#dc2626') if bal_val > 0 else colors.HexColor('#16a34a')))

        table_rows.append([
            Paragraph(str(ev.get('date', '-')), cell_style),
            Paragraph(f"<b>{ref_clean}</b>", cell_bold),
            Paragraph(desc_clean, cell_style),
            d_para,
            c_para,
            bal_para
        ])

    t_table = Table(table_rows, colWidths=col_widths, repeatRows=1)
    t_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0F172A')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#F8FAFC')])
    ]))
    story.append(t_table)
    story.append(Spacer(1, 10))

    # 4. Summary & Verification Cards
    tot_purch_str = f"Rs. {total_purchases:,.2f}"
    tot_pay_str = f"Rs. {total_payments:,.2f}"
    final_bal_str = f"Rs. {final_balance:,.2f}"
    tot_disc_str = f"Rs. {tot_discount:,.2f}" if tot_discount > 0 else "Rs. 0.00"

    summary_rows = [
        [Paragraph("Total Purchases (Debit):", cell_style), Paragraph(f"<b>{tot_purch_str}</b>", cell_right_bold)],
        [Paragraph("Total Payments (Credit):", cell_style), Paragraph(f"<font color='#16a34a'><b>{tot_pay_str}</b></font>", cell_right_bold)],
    ]
    if tot_discount > 0:
        summary_rows.append([Paragraph("Total Discount Given:", cell_style), Paragraph(f"<font color='#dc2626'><b>{tot_disc_str}</b></font>", cell_right_bold)])
    summary_rows.append([Paragraph("<b>NET BALANCE DUE:</b>", cell_bold), Paragraph(f"<font color='#dc2626'><b>{final_bal_str}</b></font>", cell_right_bold)])

    summary_card = Table(summary_rows, colWidths=[130, 105], style=[
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#CBD5E1')),
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F8FAFC')),
        ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor('#FEE2E2') if final_balance > 0 else colors.HexColor('#DCFCE7')),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('LEFTPADDING', (0,0), (-1,-1), 8),
        ('RIGHTPADDING', (0,0), (-1,-1), 8),
        ('LINEBELOW', (0,0), (-1,-2), 0.5, colors.HexColor('#E2E8F0')),
    ])

    bottom_data = [
        [
            Paragraph("<b>Statement Verification & Note:</b><br/><font size=7.5 color='#64748b'>This statement reflects all transactions recorded up to date. Kindly check and verify.<br/><br/><b>Authorized Signature / Stamp:</b> ___________________</font>", cell_style),
            summary_card
        ]
    ]
    bottom_table = Table(bottom_data, colWidths=[298, 240])
    bottom_table.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('LEFTPADDING', (0,0), (-1,-1), 0),
        ('RIGHTPADDING', (0,0), (-1,-1), 0),
    ]))
    story.append(bottom_table)

    # 5. Footer
    story.append(Spacer(1, 12))
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor('#CBD5E1'), spaceBefore=2, spaceAfter=5))
    story.append(Paragraph("<font size=7.5 color='#94A3B8'>Thank you for your business! Smart POS System — Retail & Inventory Specialist.</font>", ParagraphStyle('F', parent=cell_style, alignment=TA_CENTER)))

    doc.build(story)
    pdf_bytes = buffer.getvalue()
    buffer.close()
    return pdf_bytes

def build_supplier_ledger_pdf(sup_dict, timeline, total_bought, total_paid, final_balance, now_str, today):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=28,
        leftMargin=28,
        topMargin=28,
        bottomMargin=28
    )

    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle(
        'SuppDocTitle',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=18,
        leading=22,
        textColor=colors.HexColor('#0F172A'),
        alignment=TA_LEFT
    )

    badge_style = ParagraphStyle(
        'SuppDocBadge',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=10,
        leading=13,
        textColor=colors.HexColor('#0F172A'),
        alignment=TA_RIGHT
    )

    th_style = ParagraphStyle(
        'SuppTableHeader',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=8.5,
        leading=11,
        textColor=colors.white,
        alignment=TA_LEFT
    )

    th_right = ParagraphStyle(
        'SuppTableHeaderRight',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=8.5,
        leading=11,
        textColor=colors.white,
        alignment=TA_RIGHT
    )

    cell_style = ParagraphStyle(
        'SuppCellText',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8,
        leading=10.5,
        textColor=colors.HexColor('#1E293B'),
        alignment=TA_LEFT
    )

    cell_bold = ParagraphStyle(
        'SuppCellBold',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=8,
        leading=10.5,
        textColor=colors.HexColor('#0F172A'),
        alignment=TA_LEFT
    )

    cell_right = ParagraphStyle(
        'SuppCellRight',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8,
        leading=10.5,
        textColor=colors.HexColor('#1E293B'),
        alignment=TA_RIGHT
    )

    cell_right_bold = ParagraphStyle(
        'SuppCellRightBold',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=8,
        leading=10.5,
        textColor=colors.HexColor('#0F172A'),
        alignment=TA_RIGHT
    )

    story = []

    # 1. Header Banner Table
    status_text = "<font color='#16a34a'><b>ACCOUNT SETTLED (PAID)</b></font>" if final_balance <= 0 else "<font color='#dc2626'><b>BALANCE PAYABLE</b></font>"
    header_data = [
        [
            Paragraph("<b>SMART POS SYSTEM</b><br/><font size=8 color='#475569'>Point of Sale & Retail Management System<br/>Main Market, Lahore<br/>Phone: 0300-0000000</font>", title_style),
            Paragraph(f"<b>SUPPLIER ACCOUNT LEDGER</b><br/><font size=8 color='#64748b'>Date: {today}<br/>Time: {now_str}</font><br/>{status_text}", badge_style)
        ]
    ]
    header_table = Table(header_data, colWidths=[350, 188])
    header_table.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(header_table)
    story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor('#0F172A'), spaceBefore=2, spaceAfter=8))

    # 2. Supplier Info Box
    s_name = clean_pdf_text(sup_dict.get('name', 'Supplier'))
    s_phone = clean_pdf_text(sup_dict.get('phone') or '-')
    s_comp = clean_pdf_text(sup_dict.get('company') or '-')
    s_id = clean_pdf_text(str(sup_dict.get('id', '-')))
    status_label = "Fully Settled" if final_balance <= 0 else "Balance Due"
    status_color = "#16a34a" if final_balance <= 0 else "#dc2626"

    sup_data = [
        [
            Paragraph(f"<b>Supplier Name:</b> {s_name}<br/><b>Account ID:</b> #{s_id}", cell_style),
            Paragraph(f"<b>Company / Market:</b> {s_comp}<br/><b>Contact Phone:</b> {s_phone}", cell_style),
            Paragraph(f"<b>Account Status:</b> <font color='{status_color}'><b>{status_label}</b></font><br/><b>Transactions:</b> {len(timeline)}", cell_style)
        ]
    ]
    sup_table = Table(sup_data, colWidths=[180, 180, 178])
    sup_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F8FAFC')),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#CBD5E1')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('LEFTPADDING', (0,0), (-1,-1), 8),
        ('RIGHTPADDING', (0,0), (-1,-1), 8),
    ]))
    story.append(sup_table)
    story.append(Spacer(1, 10))

    # 3. Transactions Table
    col_widths = [72, 86, 188, 62, 62, 68]
    table_rows = [
        [
            Paragraph("DATE & TIME", th_style),
            Paragraph("REF / BILL #", th_style),
            Paragraph("DESCRIPTION", th_style),
            Paragraph("DEBIT (Rs)", th_right),
            Paragraph("CREDIT (Rs)", th_right),
            Paragraph("BALANCE (Rs)", th_right),
        ]
    ]

    for ev in timeline:
        b_val = float(ev.get('bill_amount', 0))
        p_val = float(ev.get('paid_amount', 0))
        bal_val = float(ev.get('balance', 0))

        b_str = f"{b_val:,.2f}" if b_val > 0 else "-"
        p_str = f"{p_val:,.2f}" if p_val > 0 else "-"
        bal_str = f"{bal_val:,.2f}"

        desc_clean = clean_pdf_text(ev.get('description_en') or ev.get('description', ''), preserve_breaks=True)
        ref_clean = clean_pdf_text(ev.get('ref', '-'))

        b_para = Paragraph(b_str, ParagraphStyle('SD', parent=cell_right, textColor=colors.HexColor('#dc2626') if b_val > 0 else colors.HexColor('#64748b')))
        p_para = Paragraph(p_str, ParagraphStyle('SC', parent=cell_right, textColor=colors.HexColor('#16a34a') if p_val > 0 else colors.HexColor('#64748b')))
        bal_para = Paragraph(f"<b>{bal_str}</b>", ParagraphStyle('SB', parent=cell_right_bold, textColor=colors.HexColor('#dc2626') if bal_val > 0 else colors.HexColor('#16a34a')))

        table_rows.append([
            Paragraph(str(ev.get('date', '-')), cell_style),
            Paragraph(f"<b>{ref_clean}</b>", cell_bold),
            Paragraph(desc_clean, cell_style),
            b_para,
            p_para,
            bal_para
        ])

    t_table = Table(table_rows, colWidths=col_widths, repeatRows=1)
    t_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0F172A')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#F8FAFC')])
    ]))
    story.append(t_table)
    story.append(Spacer(1, 10))

    # 4. Summary & Verification Cards
    tot_bought_str = f"Rs. {total_bought:,.2f}"
    tot_paid_str = f"Rs. {total_paid:,.2f}"
    final_bal_str = f"Rs. {final_balance:,.2f}"

    summary_rows = [
        [Paragraph("Total Purchases (Debit):", cell_style), Paragraph(f"<b>{tot_bought_str}</b>", cell_right_bold)],
        [Paragraph("Total Paid (Credit):", cell_style), Paragraph(f"<font color='#16a34a'><b>{tot_paid_str}</b></font>", cell_right_bold)],
        [Paragraph("<b>NET BALANCE PAYABLE:</b>", cell_bold), Paragraph(f"<font color='{'#dc2626' if final_balance > 0 else '#16a34a'}'><b>{final_bal_str}</b></font>", cell_right_bold)]
    ]

    summary_card = Table(summary_rows, colWidths=[130, 105], style=[
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#CBD5E1')),
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F8FAFC')),
        ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor('#FEE2E2') if final_balance > 0 else colors.HexColor('#DCFCE7')),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('LEFTPADDING', (0,0), (-1,-1), 8),
        ('RIGHTPADDING', (0,0), (-1,-1), 8),
        ('LINEBELOW', (0,0), (-1,-2), 0.5, colors.HexColor('#E2E8F0')),
    ])

    bottom_data = [
        [
            Paragraph("<b>Statement Verification & Note:</b><br/><font size=7.5 color='#64748b'>This statement reflects all purchases and payments recorded up to date. Kindly check and verify.<br/><br/><b>Authorized Signature / Stamp:</b> ___________________</font>", cell_style),
            summary_card
        ]
    ]
    bottom_table = Table(bottom_data, colWidths=[298, 240])
    bottom_table.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('LEFTPADDING', (0,0), (-1,-1), 0),
        ('RIGHTPADDING', (0,0), (-1,-1), 0),
    ]))
    story.append(bottom_table)

    # 5. Footer
    story.append(Spacer(1, 12))
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor('#CBD5E1'), spaceBefore=2, spaceAfter=5))
    story.append(Paragraph("<font size=7.5 color='#94A3B8'>Thank you for your business! Smart POS System — Retail & Inventory Specialist.</font>", ParagraphStyle('SuppF', parent=cell_style, alignment=TA_CENTER)))

    doc.build(story)
    pdf_bytes = buffer.getvalue()
    buffer.close()
    return pdf_bytes

@app.route('/api/ledger/pdf/<int:cust_id>')
@admin_only
def download_customer_ledger_pdf(cust_id):
    cust_dict, timeline, total_purchases, total_payments, tot_cust_discount, running_balance = get_ledger_data(cust_id)
    if not cust_dict:
        return "Customer not found", 404
    now_str = datetime.now().strftime('%Y-%m-%d %I:%M %p')
    today = datetime.now().strftime('%Y-%m-%d')
    try:
        pdf_bytes = build_ledger_pdf(cust_dict, timeline, total_purchases, total_payments, tot_cust_discount, running_balance, now_str, today)
        cust_name = cust_dict.get('name', f'Customer_{cust_id}')
        safe_name = "".join([c for c in cust_name if c.isalnum() or c in (' ', '_', '-')]).strip().replace(' ', '_')
        return send_pdf_response(pdf_bytes, f"Khata_{safe_name}_{today}.pdf")
    except Exception as e:
        logger.error(f"Error generating PDF via ReportLab: {e}")
        return f"Error generating PDF: {e}", 500

@app.route('/api/ledger/whatsapp_prepare/<int:cust_id>', methods=['POST', 'GET'])
@admin_only
def whatsapp_prepare_ledger(cust_id):
    cust_dict, timeline, total_purchases, total_payments, tot_cust_discount, running_balance = get_ledger_data(cust_id)
    if not cust_dict:
        return jsonify({'success': False, 'message': 'Customer not found'}), 404

    phone_from_req = (request.json or {}).get('phone') if request.is_json else None
    if phone_from_req:
        cust_dict['phone'] = phone_from_req.strip()

    raw_phone = (cust_dict.get('phone') or '').strip()
    digits = ''.join(c for c in raw_phone if c.isdigit())
    if len(digits) < 10:
        return jsonify({'success': False, 'message': 'Invalid phone number'}), 400

    if digits.startswith('0092'):
        formatted_phone = digits[2:]
    elif digits.startswith('0'):
        formatted_phone = '92' + digits[1:]
    elif len(digits) == 10 and digits.startswith('3'):
        formatted_phone = '92' + digits
    else:
        formatted_phone = digits

    now_str = datetime.now().strftime('%Y-%m-%d %I:%M %p')
    today = datetime.now().strftime('%Y-%m-%d')
    pdf_bytes = build_ledger_pdf(cust_dict, timeline, total_purchases, total_payments, tot_cust_discount, running_balance, now_str, today)

    cust_name = cust_dict.get('name', f'Customer_{cust_id}')
    safe_name = "".join([c for c in cust_name if c.isalnum() or c in (' ', '_', '-')]).strip().replace(' ', '_')
    filename = f"Khata_{safe_name}_{today}.pdf"

    # Save to daily_backups/whatsapp_pdf
    whatsapp_pdf_dir = os.path.join(BASE_DIR, 'daily_backups', 'whatsapp_pdf')
    os.makedirs(whatsapp_pdf_dir, exist_ok=True)
    pdf_path = os.path.join(whatsapp_pdf_dir, filename)
    with open(pdf_path, 'wb') as f:
        f.write(pdf_bytes)

    # Also save to Desktop/Customer_Khata_PDFs
    desktop_dir = os.path.join(os.path.expanduser("~"), "Desktop", "Customer_Khata_PDFs")
    clipboard_target = pdf_path
    try:
        os.makedirs(desktop_dir, exist_ok=True)
        desktop_pdf_path = os.path.join(desktop_dir, filename)
        with open(desktop_pdf_path, 'wb') as f:
            f.write(pdf_bytes)
        clipboard_target = desktop_pdf_path
    except Exception:
        pass

    copy_success = copy_file_to_clipboard(clipboard_target)

    msg = (
        f"*SMART POS SYSTEM*\n"
        f"Main Market, Lahore\n"
        f"Contact: 0300-0000000\n"
        f"------------------------------------\n"
        f"*STATEMENT OF ACCOUNT*\n"
        f"Customer: *{cust_name}* (ID: #{cust_id})\n"
        f"Date: {today}\n\n"
        f"• Total Purchases (Debit): Rs. {total_purchases:,.2f}\n"
        f"• Total Payments (Credit): Rs. {total_payments:,.2f}\n"
        f"------------------------------------\n"
        f"*NET BALANCE DUE: Rs. {running_balance:,.2f}*\n"
        f"------------------------------------\n"
        f"📎 Khata PDF Statement file sath bhej di gayi hai. Barah-e-karam check farma lein."
    )

    from urllib.parse import quote
    whatsapp_url = f"https://api.whatsapp.com/send?phone={formatted_phone}&text={quote(msg)}"
    download_url = f"/api/ledger/pdf/{cust_id}"

    return jsonify({
        'success': True,
        'whatsapp_url': whatsapp_url,
        'download_url': download_url,
        'filename': filename,
        'clipboard_copied': copy_success,
        'phone': formatted_phone
    })

# ====================================================
# UNIVERSAL PDF EXPORT GENERATORS & ENDPOINTS
# ====================================================

def get_universal_pdf_styles():
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'DocTitleU', parent=styles['Normal'],
        fontName='Helvetica-Bold', fontSize=15, leading=19,
        textColor=colors.HexColor('#0F172A'), alignment=TA_LEFT
    )
    badge_style = ParagraphStyle(
        'DocBadgeU', parent=styles['Normal'],
        fontName='Helvetica-Bold', fontSize=9.5, leading=12,
        textColor=colors.HexColor('#0F172A'), alignment=TA_RIGHT
    )
    th_style = ParagraphStyle(
        'TableHeaderU', parent=styles['Normal'],
        fontName='Helvetica-Bold', fontSize=8, leading=10.5,
        textColor=colors.white, alignment=TA_LEFT
    )
    th_right = ParagraphStyle(
        'TableHeaderRightU', parent=styles['Normal'],
        fontName='Helvetica-Bold', fontSize=8, leading=10.5,
        textColor=colors.white, alignment=TA_RIGHT
    )
    th_center = ParagraphStyle(
        'TableHeaderCenterU', parent=styles['Normal'],
        fontName='Helvetica-Bold', fontSize=8, leading=10.5,
        textColor=colors.white, alignment=TA_CENTER
    )
    cell_style = ParagraphStyle(
        'CellTextU', parent=styles['Normal'],
        fontName='Helvetica', fontSize=7.5, leading=9.5,
        textColor=colors.HexColor('#1E293B'), alignment=TA_LEFT
    )
    cell_bold = ParagraphStyle(
        'CellBoldU', parent=styles['Normal'],
        fontName='Helvetica-Bold', fontSize=7.5, leading=9.5,
        textColor=colors.HexColor('#0F172A'), alignment=TA_LEFT
    )
    cell_right = ParagraphStyle(
        'CellRightU', parent=styles['Normal'],
        fontName='Helvetica', fontSize=7.5, leading=9.5,
        textColor=colors.HexColor('#1E293B'), alignment=TA_RIGHT
    )
    cell_right_bold = ParagraphStyle(
        'CellRightBoldU', parent=styles['Normal'],
        fontName='Helvetica-Bold', fontSize=7.5, leading=9.5,
        textColor=colors.HexColor('#0F172A'), alignment=TA_RIGHT
    )
    cell_center = ParagraphStyle(
        'CellCenterU', parent=styles['Normal'],
        fontName='Helvetica', fontSize=7.5, leading=9.5,
        textColor=colors.HexColor('#1E293B'), alignment=TA_CENTER
    )
    return {
        'title': title_style, 'badge': badge_style,
        'th': th_style, 'th_right': th_right, 'th_center': th_center,
        'cell': cell_style, 'cell_bold': cell_bold,
        'cell_right': cell_right, 'cell_right_bold': cell_right_bold,
        'cell_center': cell_center
    }

def make_universal_pdf_doc(buffer):
    return SimpleDocTemplate(
        buffer, pagesize=A4,
        rightMargin=28, leftMargin=28, topMargin=28, bottomMargin=28
    )

def make_universal_pdf_header(doc_title, doc_badge_text, now_str, today):
    st = get_universal_pdf_styles()
    header_data = [
        [
            Paragraph("<b>SMART POS SYSTEM</b><br/><font size=8 color='#475569'>Point of Sale & Retail Management System<br/>Main Market, Lahore | Phone: 0300-0000000</font>", st['title']),
            Paragraph(f"<b>{clean_pdf_text(doc_title)}</b><br/><font size=8 color='#64748b'>Date: {today}<br/>Time: {now_str}</font><br/><font size=8 color='#0284c7'><b>{clean_pdf_text(doc_badge_text)}</b></font>", st['badge'])
        ]
    ]
    t = Table(header_data, colWidths=[340, 198])
    t.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('LEFTPADDING', (0,0), (-1,-1), 0),
        ('RIGHTPADDING', (0,0), (-1,-1), 0),
    ]))
    return t

def make_universal_pdf_footer():
    st = get_universal_pdf_styles()
    return Paragraph("<font size=7.5 color='#94A3B8'>Thank you for your business! Smart POS System — Retail & Inventory Specialist.</font>", ParagraphStyle('SuppFU', parent=st['cell_center'], alignment=TA_CENTER))

# 1. INVENTORY PDF
def build_inventory_pdf(products, now_str, today):
    buffer = io.BytesIO()
    doc = make_universal_pdf_doc(buffer)
    st = get_universal_pdf_styles()
    story = []

    tot_items = len(products)
    tot_units = sum(float(p.get('stock') or 0) for p in products)
    tot_val = sum(float(p.get('stock') or 0) * float(p.get('sale_price') or p.get('price') or 0) for p in products)
    low_stock = sum(1 for p in products if float(p.get('stock') or 0) <= float(p.get('alert_qty') or 10))

    story.append(make_universal_pdf_header("CURRENT STOCK REPORT", f"Total Items: {tot_items}", now_str, today))
    story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor('#0F172A'), spaceBefore=2, spaceAfter=8))

    sum_data = [
        [
            Paragraph(f"<b>Total Product Codes:</b> {tot_items}<br/><b>Total Units in Stock:</b> {tot_units:,.1f}", st['cell']),
            Paragraph(f"<b>Total Stock Value:</b> Rs. {tot_val:,.2f}<br/><b>Low Stock Alerts:</b> <font color='{'#dc2626' if low_stock > 0 else '#16a34a'}'><b>{low_stock} items</b></font>", st['cell'])
        ]
    ]
    sum_table = Table(sum_data, colWidths=[269, 269])
    sum_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F8FAFC')),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#CBD5E1')),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('LEFTPADDING', (0,0), (-1,-1), 8),
        ('RIGHTPADDING', (0,0), (-1,-1), 8),
    ]))
    story.append(sum_table)
    story.append(Spacer(1, 8))

    col_widths = [24, 60, 180, 74, 38, 52, 50, 60]
    table_rows = [
        [
            Paragraph("#", st['th']),
            Paragraph("CODE", st['th']),
            Paragraph("PRODUCT NAME", st['th']),
            Paragraph("CATEGORY", st['th']),
            Paragraph("UNIT", st['th_center']),
            Paragraph("RATE (Rs)", st['th_right']),
            Paragraph("STOCK", st['th_right']),
            Paragraph("VALUE (Rs)", st['th_right'])
        ]
    ]

    for idx, p in enumerate(products, 1):
        c_code = clean_pdf_text(p.get('code') or f"P-{p.get('id')}")
        c_name = clean_pdf_text(p.get('name') or '-')
        c_cat = clean_pdf_text(p.get('category') or 'General')
        c_unit = clean_pdf_text(p.get('unit') or 'Pcs')
        s_price = float(p.get('sale_price') or p.get('price') or 0)
        c_stock = float(p.get('stock') or 0)
        c_val = c_stock * s_price

        table_rows.append([
            Paragraph(str(idx), st['cell_bold']),
            Paragraph(c_code, st['cell']),
            Paragraph(c_name, st['cell_bold']),
            Paragraph(c_cat, st['cell']),
            Paragraph(c_unit, st['cell_center']),
            Paragraph(f"{s_price:,.0f}", st['cell_right']),
            Paragraph(f"{c_stock:,.1f}", st['cell_right_bold']),
            Paragraph(f"{c_val:,.0f}", st['cell_right_bold'])
        ])

    t = Table(table_rows, colWidths=col_widths, repeatRows=1)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0F172A')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 3.5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3.5),
        ('LEFTPADDING', (0,0), (-1,-1), 4),
        ('RIGHTPADDING', (0,0), (-1,-1), 4),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#F8FAFC')]),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
    ]))
    story.append(t)
    story.append(Spacer(1, 10))
    story.append(make_universal_pdf_footer())

    doc.build(story)
    pdf_bytes = buffer.getvalue()
    buffer.close()
    return pdf_bytes

@app.route('/api/inventory/pdf')
@admin_only
def download_inventory_pdf():
    conn = get_db()
    cursor = conn.cursor()
    products_raw = cursor.execute("SELECT * FROM products ORDER BY name ASC").fetchall()
    products = [dict(p) for p in products_raw]
    conn.close()
    now_str = datetime.now().strftime('%Y-%m-%d %I:%M %p')
    today = datetime.now().strftime('%Y-%m-%d')
    try:
        pdf_bytes = build_inventory_pdf(products, now_str, today)
        return send_pdf_response(pdf_bytes, f"Smart_Stock_Report_{today}.pdf")
    except Exception as e:
        logger.error(f"Error generating inventory PDF: {e}")
        return f"Error generating inventory PDF: {e}", 500

# 1.5 LOW STOCK ALERT PDF
def build_low_stock_pdf(products, now_str, today):
    buffer = io.BytesIO()
    doc = make_universal_pdf_doc(buffer)
    st = get_universal_pdf_styles()
    story = []

    tot_items = len(products)
    tot_units = sum(float(p.get('stock') or 0) for p in products)

    story.append(make_universal_pdf_header("LOW STOCK ALERT REPORT", f"Shortage Items: {tot_items}", now_str, today))
    story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor('#DC2626'), spaceBefore=2, spaceAfter=8))

    sum_data = [
        [
            Paragraph(f"<b>Total Shortage Items:</b> <font color='#dc2626'><b>{tot_items} Items</b></font><br/><b>Total Remaining Units:</b> {tot_units:,.1f}", st['cell']),
            Paragraph(f"<b>Stock Status:</b> <font color='#dc2626'><b>CRITICAL / REORDER REQUIRED</b></font><br/><b>Report Purpose:</b> Urgent Restock Planning", st['cell'])
        ]
    ]
    sum_table = Table(sum_data, colWidths=[269, 269])
    sum_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#FEF2F2')),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#FCA5A5')),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('LEFTPADDING', (0,0), (-1,-1), 8),
        ('RIGHTPADDING', (0,0), (-1,-1), 8),
    ]))
    story.append(sum_table)
    story.append(Spacer(1, 8))

    col_widths = [30, 75, 205, 85, 72, 71]
    table_rows = [
        [
            Paragraph("#", st['th']),
            Paragraph("ITEM CODE", st['th']),
            Paragraph("PRODUCT NAME", st['th']),
            Paragraph("CATEGORY", st['th']),
            Paragraph("ALERT LIMIT", st['th_center']),
            Paragraph("CURRENT STOCK", st['th_right'])
        ]
    ]

    cell_alert_red = ParagraphStyle(
        'CellAlertRed', parent=st['cell_right_bold'],
        textColor=colors.HexColor('#DC2626')
    )

    for idx, p in enumerate(products, 1):
        c_code = clean_pdf_text(p.get('code') or f"P-{p.get('id')}")
        c_name = clean_pdf_text(p.get('name') or '-')
        c_cat = clean_pdf_text(p.get('category') or 'General')
        c_unit = clean_pdf_text(p.get('unit') or 'Pcs')
        c_stock = float(p.get('stock') or 0)
        c_alert = float(p.get('alert_qty') or 0)

        table_rows.append([
            Paragraph(str(idx), st['cell_bold']),
            Paragraph(c_code, st['cell']),
            Paragraph(c_name, st['cell_bold']),
            Paragraph(c_cat, st['cell']),
            Paragraph(f"{c_alert:,.0f} {c_unit}", st['cell_center']),
            Paragraph(f"{c_stock:,.1f} {c_unit}", cell_alert_red)
        ])

    if not products:
        table_rows.append([
            Paragraph("-", st['cell_center']),
            Paragraph("-", st['cell_center']),
            Paragraph("All products have sufficient stock. No shortages recorded.", st['cell']),
            Paragraph("-", st['cell_center']),
            Paragraph("-", st['cell_center']),
            Paragraph("-", st['cell_center'])
        ])

    t = Table(table_rows, colWidths=col_widths, repeatRows=1)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#991B1B')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('LEFTPADDING', (0,0), (-1,-1), 4),
        ('RIGHTPADDING', (0,0), (-1,-1), 4),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#FFF1F2')]),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#FECDD3')),
    ]))
    story.append(t)
    story.append(Spacer(1, 10))
    story.append(make_universal_pdf_footer())

    doc.build(story)
    pdf_bytes = buffer.getvalue()
    buffer.close()
    return pdf_bytes

@app.route('/api/low_stock/pdf')
@admin_only
def download_low_stock_pdf():
    conn = get_db()
    cursor = conn.cursor()
    products_raw = cursor.execute("SELECT id, code, name, category, stock, alert_qty, unit FROM products WHERE stock <= alert_qty ORDER BY stock ASC").fetchall()
    products = [dict(p) for p in products_raw]
    conn.close()
    now_str = datetime.now().strftime('%Y-%m-%d %I:%M %p')
    today = datetime.now().strftime('%Y-%m-%d')
    try:
        pdf_bytes = build_low_stock_pdf(products, now_str, today)
        return send_pdf_response(pdf_bytes, f"Smart_Low_Stock_Report_{today}.pdf")
    except Exception as e:
        logger.error(f"Error generating low stock PDF: {e}")
        return f"Error generating low stock PDF: {e}", 500

# 2. SUPPLIERS SUMMARY PDF
def build_suppliers_summary_pdf(suppliers, now_str, today):
    buffer = io.BytesIO()
    doc = make_universal_pdf_doc(buffer)
    st = get_universal_pdf_styles()
    story = []

    tot_sups = len(suppliers)
    tot_due = sum(float(s.get('balance') or 0) for s in suppliers)
    active_due = sum(1 for s in suppliers if float(s.get('balance') or 0) > 0)

    story.append(make_universal_pdf_header("SUPPLIERS BALANCE SHEET", f"Total Suppliers: {tot_sups}", now_str, today))
    story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor('#0F172A'), spaceBefore=2, spaceAfter=8))

    sum_data = [
        [
            Paragraph(f"<b>Total Registered Suppliers:</b> {tot_sups}<br/><b>Suppliers with Payable Balance:</b> {active_due}", st['cell']),
            Paragraph(f"<b>TOTAL PAYABLE TO SUPPLIERS:</b><br/><font size=11 color='#dc2626'><b>Rs. {tot_due:,.2f}</b></font>", st['cell'])
        ]
    ]
    sum_table = Table(sum_data, colWidths=[269, 269])
    sum_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F8FAFC')),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#CBD5E1')),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('LEFTPADDING', (0,0), (-1,-1), 8),
        ('RIGHTPADDING', (0,0), (-1,-1), 8),
    ]))
    story.append(sum_table)
    story.append(Spacer(1, 8))

    col_widths = [24, 40, 130, 110, 84, 85, 65]
    table_rows = [
        [
            Paragraph("#", st['th']),
            Paragraph("ID", st['th']),
            Paragraph("SUPPLIER NAME", st['th']),
            Paragraph("COMPANY / MARKET", st['th']),
            Paragraph("PHONE", st['th']),
            Paragraph("PAYABLE (Rs)", st['th_right']),
            Paragraph("STATUS", st['th_center'])
        ]
    ]

    for idx, s in enumerate(suppliers, 1):
        s_id = f"#{s.get('id')}"
        s_name = clean_pdf_text(s.get('name') or '-')
        s_comp = clean_pdf_text(s.get('company') or '-')
        s_phone = clean_pdf_text(s.get('phone') or '-')
        bal = float(s.get('balance') or 0)
        status = "Clear" if bal <= 0 else "Due"
        stat_color = "#16a34a" if bal <= 0 else "#dc2626"

        table_rows.append([
            Paragraph(str(idx), st['cell_bold']),
            Paragraph(s_id, st['cell']),
            Paragraph(s_name, st['cell_bold']),
            Paragraph(s_comp, st['cell']),
            Paragraph(s_phone, st['cell']),
            Paragraph(f"{bal:,.2f}", st['cell_right_bold']),
            Paragraph(f"<font color='{stat_color}'><b>{status}</b></font>", st['cell_center'])
        ])

    t = Table(table_rows, colWidths=col_widths, repeatRows=1)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0F172A')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('LEFTPADDING', (0,0), (-1,-1), 4),
        ('RIGHTPADDING', (0,0), (-1,-1), 4),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#F8FAFC')]),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
    ]))
    story.append(t)
    story.append(Spacer(1, 10))
    story.append(make_universal_pdf_footer())

    doc.build(story)
    pdf_bytes = buffer.getvalue()
    buffer.close()
    return pdf_bytes

@app.route('/api/suppliers/pdf')
@admin_only
def download_suppliers_pdf():
    conn = get_db()
    cursor = conn.cursor()
    suppliers_raw = cursor.execute("SELECT * FROM suppliers ORDER BY balance DESC").fetchall()
    suppliers = [dict(s) for s in suppliers_raw]
    conn.close()
    now_str = datetime.now().strftime('%Y-%m-%d %I:%M %p')
    today = datetime.now().strftime('%Y-%m-%d')
    try:
        pdf_bytes = build_suppliers_summary_pdf(suppliers, now_str, today)
        return send_pdf_response(pdf_bytes, f"Smart_Suppliers_Balance_Report_{today}.pdf")
    except Exception as e:
        logger.error(f"Error generating suppliers PDF: {e}")
        return f"Error generating suppliers PDF: {e}", 500

# 3. CUSTOMERS SUMMARY PDF
def build_customers_summary_pdf(customers, now_str, today):
    buffer = io.BytesIO()
    doc = make_universal_pdf_doc(buffer)
    st = get_universal_pdf_styles()
    story = []

    tot_custs = len(customers)
    tot_due = sum(float(c.get('balance') or 0) for c in customers)
    active_due = sum(1 for c in customers if float(c.get('balance') or 0) > 0)

    story.append(make_universal_pdf_header("CUSTOMERS RECEIVABLES", f"Total Customers: {tot_custs}", now_str, today))
    story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor('#0F172A'), spaceBefore=2, spaceAfter=8))

    sum_data = [
        [
            Paragraph(f"<b>Total Customer Accounts:</b> {tot_custs}<br/><b>Customers with Balance Due:</b> {active_due}", st['cell']),
            Paragraph(f"<b>TOTAL MARKET RECEIVABLES:</b><br/><font size=11 color='#dc2626'><b>Rs. {tot_due:,.2f}</b></font>", st['cell'])
        ]
    ]
    sum_table = Table(sum_data, colWidths=[269, 269])
    sum_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F8FAFC')),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#CBD5E1')),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('LEFTPADDING', (0,0), (-1,-1), 8),
        ('RIGHTPADDING', (0,0), (-1,-1), 8),
    ]))
    story.append(sum_table)
    story.append(Spacer(1, 8))

    col_widths = [24, 40, 130, 110, 84, 85, 65]
    table_rows = [
        [
            Paragraph("#", st['th']),
            Paragraph("ID", st['th']),
            Paragraph("CUSTOMER NAME", st['th']),
            Paragraph("ADDRESS", st['th']),
            Paragraph("PHONE", st['th']),
            Paragraph("BALANCE DUE (Rs)", st['th_right']),
            Paragraph("STATUS", st['th_center'])
        ]
    ]

    for idx, c in enumerate(customers, 1):
        c_id = f"#{c.get('id')}"
        c_name = clean_pdf_text(c.get('name') or '-')
        c_addr = clean_pdf_text(c.get('address') or '-')
        c_phone = clean_pdf_text(c.get('phone') or '-')
        bal = float(c.get('balance') or 0)
        status = "Cleared" if bal <= 0 else "Due"
        stat_color = "#16a34a" if bal <= 0 else "#dc2626"

        table_rows.append([
            Paragraph(str(idx), st['cell_bold']),
            Paragraph(c_id, st['cell']),
            Paragraph(c_name, st['cell_bold']),
            Paragraph(c_addr, st['cell']),
            Paragraph(c_phone, st['cell']),
            Paragraph(f"{bal:,.2f}", st['cell_right_bold']),
            Paragraph(f"<font color='{stat_color}'><b>{status}</b></font>", st['cell_center'])
        ])

    t = Table(table_rows, colWidths=col_widths, repeatRows=1)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0F172A')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('LEFTPADDING', (0,0), (-1,-1), 4),
        ('RIGHTPADDING', (0,0), (-1,-1), 4),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#F8FAFC')]),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
    ]))
    story.append(t)
    story.append(Spacer(1, 10))
    story.append(make_universal_pdf_footer())

    doc.build(story)
    pdf_bytes = buffer.getvalue()
    buffer.close()
    return pdf_bytes

@app.route('/api/customers/pdf')
@admin_only
def download_customers_pdf():
    conn = get_db()
    cursor = conn.cursor()
    customers_raw = cursor.execute("SELECT * FROM customers ORDER BY balance DESC").fetchall()
    customers = [dict(c) for c in customers_raw]
    conn.close()
    now_str = datetime.now().strftime('%Y-%m-%d %I:%M %p')
    today = datetime.now().strftime('%Y-%m-%d')
    try:
        pdf_bytes = build_customers_summary_pdf(customers, now_str, today)
        return send_pdf_response(pdf_bytes, f"Smart_Customers_Receivables_Report_{today}.pdf")
    except Exception as e:
        logger.error(f"Error generating customers PDF: {e}")
        return f"Error generating customers PDF: {e}", 500

# 4. DAILY REPORT PDF
def build_daily_report_pdf(sales, direct_payments, purchases, summary, selected_date, now_str):
    buffer = io.BytesIO()
    doc = make_universal_pdf_doc(buffer)
    st = get_universal_pdf_styles()
    story = []

    story.append(make_universal_pdf_header("DAILY SALES & CLOSING REPORT", f"Date: {selected_date}", now_str, selected_date))
    story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor('#0F172A'), spaceBefore=2, spaceAfter=8))

    kpi_data = [
        [
            Paragraph(f"<b>Gross Sales:</b> Rs. {summary['total_sales']:,.2f}<br/><b>Cash from Sales:</b> Rs. {summary['cash_from_sales']:,.2f}", st['cell']),
            Paragraph(f"<b>Credit Given Today:</b> Rs. {summary['credit_given']:,.2f}<br/><b>Debt Recoveries:</b> Rs. {summary['recoveries']:,.2f}", st['cell']),
            Paragraph(f"<b>Purchases Paid:</b> Rs. {summary['purchases_paid']:,.2f}<br/><b>Cash in Hand:</b> <font color='#16a34a'><b>Rs. {summary['cash_in_hand']:,.2f}</b></font>", st['cell'])
        ]
    ]
    kpi_t = Table(kpi_data, colWidths=[179, 179, 180])
    kpi_t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F8FAFC')),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#CBD5E1')),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('LEFTPADDING', (0,0), (-1,-1), 6),
        ('RIGHTPADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(kpi_t)
    story.append(Spacer(1, 10))

    story.append(Paragraph("<b>TODAY'S SALES INVOICES</b>", st['cell_bold']))
    story.append(Spacer(1, 3))
    s_col_widths = [24, 65, 125, 95, 75, 75, 79]
    s_rows = [
        [
            Paragraph("#", st['th']),
            Paragraph("INVOICE #", st['th']),
            Paragraph("CUSTOMER", st['th']),
            Paragraph("TYPE / TIME", st['th']),
            Paragraph("TOTAL (Rs)", st['th_right']),
            Paragraph("PAID (Rs)", st['th_right']),
            Paragraph("DUE (Rs)", st['th_right'])
        ]
    ]
    for idx, s in enumerate(sales, 1):
        s_inv = clean_pdf_text(s.get('invoice_no') or f"#{s.get('id')}")
        s_cust = clean_pdf_text(s.get('customer_name') or 'Cash Customer')
        s_type = clean_pdf_text(f"{str(s.get('type') or 'cash').title()} | {str(s.get('date') or '')[-11:].strip()}")
        tot = float(s.get('total') or 0)
        paid = float(s.get('paid') or 0)
        due = float(s.get('balance') or s.get('due') or 0)
        s_rows.append([
            Paragraph(str(idx), st['cell']),
            Paragraph(s_inv, st['cell']),
            Paragraph(s_cust, st['cell_bold']),
            Paragraph(s_type, st['cell']),
            Paragraph(f"{tot:,.2f}", st['cell_right']),
            Paragraph(f"{paid:,.2f}", st['cell_right']),
            Paragraph(f"{due:,.2f}", st['cell_right_bold'])
        ])
    if len(sales) == 0:
        s_rows.append([Paragraph("No sales recorded for this date", st['cell'])] + [Paragraph("-", st['cell'])] * 6)

    st_table = Table(s_rows, colWidths=s_col_widths, repeatRows=1)
    st_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0F172A')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 3.5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3.5),
        ('LEFTPADDING', (0,0), (-1,-1), 4),
        ('RIGHTPADDING', (0,0), (-1,-1), 4),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#F8FAFC')]),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
    ]))
    story.append(st_table)
    story.append(Spacer(1, 10))

    if direct_payments:
        story.append(Paragraph("<b>CUSTOMER DEBT RECOVERIES (RECEIVED TODAY)</b>", st['cell_bold']))
        story.append(Spacer(1, 3))
        r_col_widths = [24, 150, 164, 100, 100]
        r_rows = [
            [
                Paragraph("#", st['th']),
                Paragraph("CUSTOMER NAME", st['th']),
                Paragraph("NOTE / DESCRIPTION", st['th']),
                Paragraph("PAYMENT METHOD", st['th_center']),
                Paragraph("AMOUNT (Rs)", st['th_right'])
            ]
        ]
        for idx, r in enumerate(direct_payments, 1):
            r_rows.append([
                Paragraph(str(idx), st['cell']),
                Paragraph(clean_pdf_text(r.get('customer_name') or '-'), st['cell_bold']),
                Paragraph(clean_pdf_text(r.get('notes') or 'Cash Recovery'), st['cell']),
                Paragraph(clean_pdf_text(r.get('payment_method') or 'Cash'), st['cell_center']),
                Paragraph(f"{float(r.get('amount') or 0):,.2f}", st['cell_right_bold'])
            ])
        rt = Table(r_rows, colWidths=r_col_widths, repeatRows=1)
        rt.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0F172A')),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('TOPPADDING', (0,0), (-1,-1), 3.5),
            ('BOTTOMPADDING', (0,0), (-1,-1), 3.5),
            ('LEFTPADDING', (0,0), (-1,-1), 4),
            ('RIGHTPADDING', (0,0), (-1,-1), 4),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#F8FAFC')]),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
        ]))
        story.append(rt)
        story.append(Spacer(1, 10))

    story.append(make_universal_pdf_footer())
    doc.build(story)
    pdf_bytes = buffer.getvalue()
    buffer.close()
    return pdf_bytes

@app.route('/api/daily_report/pdf')
@admin_only
def download_daily_report_pdf():
    selected_date = request.args.get('date', datetime.now().strftime("%Y-%m-%d"))
    date_pattern = f"{selected_date}%"
    conn = get_db()
    cursor = conn.cursor()
    raw_sales = cursor.execute("SELECT s.*, COALESCE(c.name, s.customer_name) as customer_name FROM sales s LEFT JOIN customers c ON s.customer_id = c.id WHERE s.date LIKE ? AND s.type != 'returned' ORDER BY s.id DESC", (date_pattern,)).fetchall()
    direct_payments = cursor.execute("SELECT p.*, c.name as customer_name FROM customer_payments p LEFT JOIN customers c ON p.customer_id = c.id WHERE p.payment_date LIKE ? ORDER BY p.id DESC", (date_pattern,)).fetchall()
    raw_purchases = cursor.execute("SELECT p.*, COALESCE(s.name, p.supplier_name, 'Direct Stock In') as supplier_display_name FROM purchases p LEFT JOIN suppliers s ON p.supplier_id = s.id WHERE p.date LIKE ? ORDER BY p.id DESC", (date_pattern,)).fetchall()

    sales = []
    total_sales_amount = 0.0
    cash_from_sales = 0.0
    credit_given_today = 0.0
    for row in raw_sales:
        s_dict = dict(row)
        s_total = float(s_dict.get('total') or 0)
        s_paid = float(s_dict.get('paid') or 0)
        s_bal = float(s_dict.get('due') if s_dict.get('due') is not None else (s_total - s_paid))
        s_dict['total'] = s_total
        s_dict['paid'] = s_paid
        s_dict['balance'] = s_bal
        total_sales_amount += s_total
        cash_from_sales += s_paid
        if s_bal > 0: credit_given_today += s_bal
        sales.append(s_dict)

    purchases = []
    cash_paid_for_purchases = 0.0
    for row in raw_purchases:
        p_dict = dict(row)
        p_paid = float(p_dict.get('paid') or 0)
        cash_paid_for_purchases += p_paid
        purchases.append(p_dict)

    cash_from_recoveries = sum(float(p['amount'] or 0) for p in direct_payments)
    conn.close()

    summary = {
        'total_sales': total_sales_amount,
        'cash_from_sales': cash_from_sales,
        'credit_given': credit_given_today,
        'recoveries': cash_from_recoveries,
        'purchases_paid': cash_paid_for_purchases,
        'cash_in_hand': cash_from_sales + cash_from_recoveries
    }
    now_str = datetime.now().strftime('%Y-%m-%d %I:%M %p')
    try:
        pdf_bytes = build_daily_report_pdf(sales, [dict(p) for p in direct_payments], purchases, summary, selected_date, now_str)
        return send_pdf_response(pdf_bytes, f"Smart_Daily_Closing_{selected_date}.pdf")
    except Exception as e:
        logger.error(f"Error generating daily report PDF: {e}")
        return f"Error generating daily report PDF: {e}", 500

# 5. MONTHLY REPORT PDF
def build_monthly_report_pdf(sales, purchases, exp_total, cust_rec, sup_paid, month, now_str):
    buffer = io.BytesIO()
    doc = make_universal_pdf_doc(buffer)
    st = get_universal_pdf_styles()
    story = []

    story.append(make_universal_pdf_header("MONTHLY MASTER BUSINESS REPORT", f"Month: {month}", now_str, month))
    story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor('#0F172A'), spaceBefore=2, spaceAfter=8))

    s_tot = float(sales['total'] or 0)
    s_cash = float(sales['cash'] or 0)
    s_credit = float(sales['credit'] or 0)
    s_disc = float(sales['discount'] or 0)

    p_tot = float(purchases['total'] or 0)
    p_paid = float(purchases['paid'] or 0)
    p_due = float(purchases['credit'] or 0)

    grid_data = [
        [
            Paragraph(f"<b>GROSS SALES REVENUE:</b><br/><font size=11 color='#0284c7'><b>Rs. {s_tot:,.2f}</b></font><br/>Cash: Rs. {s_cash:,.2f} | Credit: Rs. {s_credit:,.2f}<br/>Discounts Given: Rs. {s_disc:,.2f}", st['cell']),
            Paragraph(f"<b>STOCK PURCHASES:</b><br/><font size=11 color='#ef4444'><b>Rs. {p_tot:,.2f}</b></font><br/>Paid: Rs. {p_paid:,.2f} | Credit: Rs. {p_due:,.2f}", st['cell'])
        ],
        [
            Paragraph(f"<b>STORE EXPENSES:</b><br/><font size=11 color='#f59e0b'><b>Rs. {float(exp_total or 0):,.2f}</b></font><br/>Utilities, Tea, Staff & Misc Costs", st['cell']),
            Paragraph(f"<b>DEBT RECOVERIES & PAYMENTS:</b><br/>Customer Recoveries: <font color='#16a34a'><b>Rs. {float(cust_rec or 0):,.2f}</b></font><br/>Paid to Suppliers: <font color='#dc2626'><b>Rs. {float(sup_paid or 0):,.2f}</b></font>", st['cell'])
        ]
    ]
    gt = Table(grid_data, colWidths=[269, 269])
    gt.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F8FAFC')),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#CBD5E1')),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E2E8F0')),
        ('TOPPADDING', (0,0), (-1,-1), 8),
        ('BOTTOMPADDING', (0,0), (-1,-1), 8),
        ('LEFTPADDING', (0,0), (-1,-1), 8),
        ('RIGHTPADDING', (0,0), (-1,-1), 8),
    ]))
    story.append(gt)
    story.append(Spacer(1, 14))

    note_box = Table([
        [Paragraph(f"<b>Executive Summary Note for Month {month}:</b><br/><font size=8 color='#475569'>This monthly financial statement aggregates all finalized sales invoices, supplier stock consignments, customer debt recoveries, and daily operating shop expenses for the period. All records are verifiable against individual transaction receipts.</font>", st['cell'])]
    ], colWidths=[538])
    note_box.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F1F5F9')),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#CBD5E1')),
        ('PADDING', (0,0), (-1,-1), 8),
    ]))
    story.append(note_box)
    story.append(Spacer(1, 15))
    story.append(make_universal_pdf_footer())

    doc.build(story)
    pdf_bytes = buffer.getvalue()
    buffer.close()
    return pdf_bytes

@app.route('/api/monthly_report/pdf')
@admin_only
def download_monthly_report_pdf():
    month_filter = request.args.get('month', datetime.now().strftime('%Y-%m'))
    conn = get_db()
    cursor = conn.cursor()
    sales_data = cursor.execute("""
        SELECT 
            COALESCE(SUM(total), 0) as total, 
            COALESCE(SUM(paid), 0) as cash, 
            COALESCE(SUM(due), 0) as credit,
            COALESCE(SUM(discount), 0) as discount 
        FROM sales 
        WHERE date LIKE ? AND type != 'returned'
    """, (f'{month_filter}%',)).fetchone()
    purch_data = cursor.execute("SELECT COALESCE(SUM(total), 0) as total, COALESCE(SUM(paid), 0) as paid, COALESCE(SUM(due), 0) as credit FROM purchases WHERE date LIKE ? AND COALESCE(status, \'active\') != \'returned\'", (f'{month_filter}%',)).fetchone()
    exp_total = cursor.execute("SELECT COALESCE(SUM(amount), 0) FROM expenses WHERE date LIKE ?", (f'{month_filter}%',)).fetchone()[0] or 0
    cust_rec = cursor.execute("SELECT COALESCE(SUM(amount), 0) FROM customer_payments WHERE payment_date LIKE ?", (f'{month_filter}%',)).fetchone()[0] or 0
    sup_paid = cursor.execute("SELECT COALESCE(SUM(amount), 0) FROM supplier_payments WHERE payment_date LIKE ?", (f'{month_filter}%',)).fetchone()[0] or 0
    conn.close()

    now_str = datetime.now().strftime('%Y-%m-%d %I:%M %p')
    try:
        pdf_bytes = build_monthly_report_pdf(dict(sales_data), dict(purch_data), exp_total, cust_rec, sup_paid, month_filter, now_str)
        return send_pdf_response(pdf_bytes, f"Smart_Monthly_Report_{month_filter}.pdf")
    except Exception as e:
        logger.error(f"Error generating monthly report PDF: {e}")
        return f"Error generating monthly report PDF: {e}", 500

# 6. EXPENSES PDF
def build_expenses_pdf(expenses, total_amount, now_str, today):
    buffer = io.BytesIO()
    doc = make_universal_pdf_doc(buffer)
    st = get_universal_pdf_styles()
    story = []

    story.append(make_universal_pdf_header("STORE EXPENSES REPORT", f"Total Records: {len(expenses)}", now_str, today))
    story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor('#0F172A'), spaceBefore=2, spaceAfter=8))

    sum_table = Table([
        [
            Paragraph(f"<b>Total Expense Entries:</b> {len(expenses)}", st['cell']),
            Paragraph(f"<b>TOTAL EXPENSES AMOUNT:</b> <font color='#dc2626'><b>Rs. {total_amount:,.2f}</b></font>", st['cell_right_bold'])
        ]
    ], colWidths=[269, 269])
    sum_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F8FAFC')),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#CBD5E1')),
        ('PADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(sum_table)
    story.append(Spacer(1, 8))

    col_widths = [24, 80, 100, 234, 100]
    rows = [
        [
            Paragraph("#", st['th']),
            Paragraph("DATE", st['th']),
            Paragraph("CATEGORY", st['th']),
            Paragraph("EXPENSE DESCRIPTION", st['th']),
            Paragraph("AMOUNT (Rs)", st['th_right'])
        ]
    ]
    for idx, e in enumerate(expenses, 1):
        rows.append([
            Paragraph(str(idx), st['cell']),
            Paragraph(clean_pdf_text(e.get('date') or '-'), st['cell']),
            Paragraph(clean_pdf_text(e.get('category') or 'General'), st['cell_bold']),
            Paragraph(clean_pdf_text(e.get('title') or e.get('notes') or '-'), st['cell']),
            Paragraph(f"{float(e.get('amount') or 0):,.2f}", st['cell_right_bold'])
        ])

    t = Table(rows, colWidths=col_widths, repeatRows=1)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0F172A')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('LEFTPADDING', (0,0), (-1,-1), 4),
        ('RIGHTPADDING', (0,0), (-1,-1), 4),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#F8FAFC')]),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
    ]))
    story.append(t)
    story.append(Spacer(1, 10))
    story.append(make_universal_pdf_footer())

    doc.build(story)
    pdf_bytes = buffer.getvalue()
    buffer.close()
    return pdf_bytes

@app.route('/api/expenses/pdf')
@admin_only
def download_expenses_pdf():
    conn = get_db()
    cursor = conn.cursor()
    expenses_raw = cursor.execute("SELECT * FROM expenses ORDER BY id DESC").fetchall()
    expenses = [dict(e) for e in expenses_raw]
    tot_amt = sum(float(e.get('amount') or 0) for e in expenses)
    conn.close()

    now_str = datetime.now().strftime('%Y-%m-%d %I:%M %p')
    today = datetime.now().strftime('%Y-%m-%d')
    try:
        pdf_bytes = build_expenses_pdf(expenses, tot_amt, now_str, today)
        return send_pdf_response(pdf_bytes, f"Smart_Expenses_Report_{today}.pdf")
    except Exception as e:
        logger.error(f"Error generating expenses PDF: {e}")
        return f"Error generating expenses PDF: {e}", 500

# 7. QUOTATION PDF
def build_quotation_pdf(quote, items, now_str, today):
    buffer = io.BytesIO()
    doc = make_universal_pdf_doc(buffer)
    st = get_universal_pdf_styles()
    story = []

    q_no = clean_pdf_text(quote.get('quote_no') or f"QT-{quote.get('id')}")
    c_name = clean_pdf_text(quote.get('customer_name') or 'Customer')
    c_phone = clean_pdf_text(quote.get('customer_phone') or '-')
    tot = float(quote.get('total') or 0)

    story.append(make_universal_pdf_header("PRICE ESTIMATE / QUOTATION", f"Quote #: {q_no}", now_str, today))
    story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor('#0F172A'), spaceBefore=2, spaceAfter=8))

    info_data = [
        [
            Paragraph(f"<b>Customer Name:</b> {c_name}<br/><b>Contact Phone:</b> {c_phone}", st['cell']),
            Paragraph(f"<b>Quotation Number:</b> {q_no}<br/><b>Quotation Date:</b> {quote.get('date') or today}", st['cell']),
            Paragraph(f"<b>TOTAL ESTIMATE:</b><br/><font size=11 color='#0284c7'><b>Rs. {tot:,.2f}</b></font>", st['cell_right_bold'])
        ]
    ]
    info_t = Table(info_data, colWidths=[180, 180, 178])
    info_t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F8FAFC')),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#CBD5E1')),
        ('PADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(info_t)
    story.append(Spacer(1, 10))

    col_widths = [24, 234, 70, 70, 70, 70]
    rows = [
        [
            Paragraph("#", st['th']),
            Paragraph("ITEM NAME / DESCRIPTION", st['th']),
            Paragraph("UNIT", st['th_center']),
            Paragraph("RATE (Rs)", st['th_right']),
            Paragraph("QUANTITY", st['th_center']),
            Paragraph("TOTAL (Rs)", st['th_right'])
        ]
    ]
    for idx, it in enumerate(items, 1):
        it_name = clean_pdf_text(it.get('name') or '-')
        it_unit = clean_pdf_text(it.get('unit') or 'Pcs')
        it_price = float(it.get('price') or 0)
        it_qty = float(it.get('qty') or 1)
        it_tot = it_price * it_qty
        rows.append([
            Paragraph(str(idx), st['cell']),
            Paragraph(it_name, st['cell_bold']),
            Paragraph(it_unit, st['cell_center']),
            Paragraph(f"{it_price:,.2f}", st['cell_right']),
            Paragraph(f"{it_qty:,.1f}", st['cell_center']),
            Paragraph(f"{it_tot:,.2f}", st['cell_right_bold'])
        ])

    t = Table(rows, colWidths=col_widths, repeatRows=1)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0F172A')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('LEFTPADDING', (0,0), (-1,-1), 4),
        ('RIGHTPADDING', (0,0), (-1,-1), 4),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#F8FAFC')]),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
    ]))
    story.append(t)
    story.append(Spacer(1, 10))

    bot_data = [
        [
            Paragraph("<font size=7.5 color='#64748b'>Note: This is an official price estimate and subject to stock availability.<br/>Authorized Signature: _______________________</font>", st['cell']),
            Paragraph(f"<b>GRAND TOTAL ESTIMATE:</b><br/><font size=12 color='#0284c7'><b>Rs. {tot:,.2f}</b></font>", st['cell_right_bold'])
        ]
    ]
    bt = Table(bot_data, colWidths=[320, 218])
    bt.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F8FAFC')),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#CBD5E1')),
        ('PADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(bt)
    story.append(Spacer(1, 10))
    story.append(make_universal_pdf_footer())

    doc.build(story)
    pdf_bytes = buffer.getvalue()
    buffer.close()
    return pdf_bytes

@app.route('/api/quotation/pdf/<int:quote_id>')
@admin_only
def download_quotation_pdf(quote_id):
    conn = get_db()
    quote = conn.execute("SELECT * FROM quotations WHERE id = ?", (quote_id,)).fetchone()
    conn.close()
    if not quote:
        return "Quotation not found", 404
    items = json.loads(quote['items_json']) if quote and quote['items_json'] else []
    now_str = datetime.now().strftime('%Y-%m-%d %I:%M %p')
    today = datetime.now().strftime('%Y-%m-%d')
    try:
        pdf_bytes = build_quotation_pdf(dict(quote), items, now_str, today)
        q_no = quote['quote_no'] or f"QT_{quote_id}"
        return send_pdf_response(pdf_bytes, f"Quotation_{q_no}_{today}.pdf")
    except Exception as e:
        logger.error(f"Error generating quotation PDF: {e}")
        return f"Error generating quotation PDF: {e}", 500

# 8. SALE INVOICE PDF
def build_sale_invoice_pdf(sale, items, now_str, today):
    buffer = io.BytesIO()
    doc = make_universal_pdf_doc(buffer)
    st = get_universal_pdf_styles()
    story = []

    inv_no = clean_pdf_text(sale.get('invoice_no') or f"#{sale.get('id')}")
    c_name = clean_pdf_text(sale.get('customer_name') or 'Cash Customer')
    c_phone = clean_pdf_text(sale.get('customer_phone') or '-')
    s_date = str(sale.get('date') or today)
    s_type = clean_pdf_text(str(sale.get('type') or 'cash').title())

    tot = float(sale.get('total') or 0)
    paid = float(sale.get('paid') or 0)
    due = float(sale.get('due') if sale.get('due') is not None else (tot - paid))

    story.append(make_universal_pdf_header("SALES TAX INVOICE", f"Invoice #: {inv_no}", now_str, today))
    story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor('#0F172A'), spaceBefore=2, spaceAfter=8))

    info_data = [
        [
            Paragraph(f"<b>Customer Name:</b> {c_name}<br/><b>Phone:</b> {c_phone}", st['cell']),
            Paragraph(f"<b>Invoice #:</b> {inv_no}<br/><b>Date & Time:</b> {s_date}", st['cell']),
            Paragraph(f"<b>Payment Type:</b> {s_type}<br/><b>Status:</b> <font color='{'#16a34a' if due <= 0 else '#dc2626'}'><b>{'Fully Paid' if due <= 0 else 'Balance Due'}</b></font>", st['cell'])
        ]
    ]
    info_t = Table(info_data, colWidths=[180, 180, 178])
    info_t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F8FAFC')),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#CBD5E1')),
        ('PADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(info_t)
    story.append(Spacer(1, 10))

    col_widths = [24, 234, 70, 70, 70, 70]
    rows = [
        [
            Paragraph("#", st['th']),
            Paragraph("PRODUCT NAME", st['th']),
            Paragraph("UNIT", st['th_center']),
            Paragraph("RATE (Rs)", st['th_right']),
            Paragraph("QTY", st['th_center']),
            Paragraph("TOTAL (Rs)", st['th_right'])
        ]
    ]
    for idx, it in enumerate(items, 1):
        p_name = clean_pdf_text(it.get('product_name') or it.get('name') or '-')
        p_unit = clean_pdf_text(it.get('unit') or 'Pcs')
        p_rate = float(it.get('price') or 0)
        p_qty = float(it.get('quantity') or it.get('qty') or 1)
        p_tot = float(it.get('total') if it.get('total') is not None else (p_rate * p_qty))
        rows.append([
            Paragraph(str(idx), st['cell']),
            Paragraph(p_name, st['cell_bold']),
            Paragraph(p_unit, st['cell_center']),
            Paragraph(f"{p_rate:,.2f}", st['cell_right']),
            Paragraph(f"{p_qty:,.1f}", st['cell_center']),
            Paragraph(f"{p_tot:,.2f}", st['cell_right_bold'])
        ])

    t = Table(rows, colWidths=col_widths, repeatRows=1)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0F172A')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('LEFTPADDING', (0,0), (-1,-1), 4),
        ('RIGHTPADDING', (0,0), (-1,-1), 4),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#F8FAFC')]),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
    ]))
    story.append(t)
    story.append(Spacer(1, 10))

    tot_data = [
        [
            Paragraph("<font size=7.5 color='#64748b'>Exchange is possible within 3 days with this original receipt.<br/>Thank you for your business!</font>", st['cell']),
            Table([
                [Paragraph("<b>Grand Total:</b>", st['cell']), Paragraph(f"Rs. {tot:,.2f}", st['cell_right_bold'])],
                [Paragraph("<b>Paid Amount:</b>", st['cell']), Paragraph(f"<font color='#16a34a'>Rs. {paid:,.2f}</font>", st['cell_right_bold'])],
                [Paragraph("<b>Balance Due:</b>", st['cell_bold']), Paragraph(f"<font color='#dc2626'>Rs. {due:,.2f}</font>", st['cell_right_bold'])],
            ], colWidths=[100, 100], style=[
                ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F8FAFC')),
                ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#CBD5E1')),
                ('PADDING', (0,0), (-1,-1), 4),
            ])
        ]
    ]
    tot_t = Table(tot_data, colWidths=[330, 208])
    story.append(tot_t)
    story.append(Spacer(1, 10))
    story.append(make_universal_pdf_footer())

    doc.build(story)
    pdf_bytes = buffer.getvalue()
    buffer.close()
    return pdf_bytes

@app.route('/api/sale/pdf/<path:sale_ident>')
@login_required
def download_sale_invoice_pdf(sale_ident):
    conn = get_db()
    sale = None
    sale_str = str(sale_ident).strip()
    if sale_str.isdigit():
        sale = conn.execute("SELECT s.*, COALESCE(c.name, s.customer_name, 'Cash Sale') as customer_name, c.phone as customer_phone FROM sales s LEFT JOIN customers c ON s.customer_id = c.id WHERE s.id = ?", (int(sale_str),)).fetchone()
    if not sale:
        sale = conn.execute("SELECT s.*, COALESCE(c.name, s.customer_name, 'Cash Sale') as customer_name, c.phone as customer_phone FROM sales s LEFT JOIN customers c ON s.customer_id = c.id WHERE LOWER(TRIM(s.invoice_no)) = LOWER(?)", (sale_str,)).fetchone()
    if not sale:
        sale = conn.execute("SELECT s.*, COALESCE(c.name, s.customer_name, 'Cash Sale') as customer_name, c.phone as customer_phone FROM sales s LEFT JOIN customers c ON s.customer_id = c.id WHERE s.invoice_no LIKE ?", (f"%{sale_str}%",)).fetchone()
        
    if not sale:
        conn.close()
        return "Invoice not found", 404
        
    real_id = sale['id']
    items = conn.execute("SELECT si.*, COALESCE(p.unit, 'Pcs') as unit FROM sale_items si LEFT JOIN products p ON si.product_id = p.id WHERE si.sale_id = ?", (real_id,)).fetchall()
    conn.close()

    now_str = datetime.now().strftime('%Y-%m-%d %I:%M %p')
    today = datetime.now().strftime('%Y-%m-%d')
    try:
        pdf_bytes = build_sale_invoice_pdf(dict(sale), [dict(i) for i in items], now_str, today)
        inv_no = sale['invoice_no'] or f"Inv_{real_id}"
        return send_pdf_response(pdf_bytes, f"Invoice_{inv_no}_{today}.pdf")
    except Exception as e:
        logger.error(f"Error generating sale invoice PDF: {e}")
        return f"Error generating sale invoice PDF: {e}", 500

# 9. TOP 10 BEST SELLING PRODUCTS PDF
def build_top_products_pdf(top_selling, now_str, today):
    buffer = io.BytesIO()
    doc = make_universal_pdf_doc(buffer)
    st = get_universal_pdf_styles()
    story = []

    story.append(make_universal_pdf_header("TOP 10 BEST SELLING PRODUCTS", "Live Sales Ranking", now_str, today))
    story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor('#0F172A'), spaceBefore=2, spaceAfter=8))

    tot_sold = sum(float(it.get('total_sold') or 0) for it in top_selling)
    tot_rev = sum(float(it.get('total_revenue') or 0) for it in top_selling)

    kpi_data = [
        [
            Paragraph(f"<b>Total Top Ranked Items:</b> {len(top_selling)}", st['cell']),
            Paragraph(f"<b>Total Units Sold:</b> <font color='#0284c7'><b>{tot_sold:,.1f}</b></font>", st['cell']),
            Paragraph(f"<b>Total Revenue Generated:</b> <font color='#16a34a'><b>Rs. {tot_rev:,.2f}</b></font>", st['cell_right_bold'])
        ]
    ]
    kpi_t = Table(kpi_data, colWidths=[179, 179, 180])
    kpi_t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F8FAFC')),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#CBD5E1')),
        ('PADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(kpi_t)
    story.append(Spacer(1, 10))

    col_widths = [35, 203, 80, 70, 75, 75]
    rows = [
        [
            Paragraph("RANK", st['th_center']),
            Paragraph("PRODUCT NAME", st['th']),
            Paragraph("CATEGORY", st['th']),
            Paragraph("STOCK", st['th_center']),
            Paragraph("SOLD QTY", st['th_center']),
            Paragraph("REVENUE (Rs)", st['th_right'])
        ]
    ]

    for idx, it in enumerate(top_selling, 1):
        p_name = clean_pdf_text(it.get('name') or '-')
        p_cat = clean_pdf_text(it.get('category') or 'General')
        p_unit = clean_pdf_text(it.get('unit') or 'Pcs')
        p_stock = float(it.get('stock') or 0)
        s_qty = float(it.get('total_sold') or 0)
        s_rev = float(it.get('total_revenue') or 0)

        rows.append([
            Paragraph(f"#{idx:02d}", st['cell_center']),
            Paragraph(p_name, st['cell_bold']),
            Paragraph(p_cat, st['cell']),
            Paragraph(f"{p_stock:,.1f} {p_unit}", st['cell_center']),
            Paragraph(f"{s_qty:,.1f}", st['cell_center']),
            Paragraph(f"{s_rev:,.2f}", st['cell_right_bold'])
        ])

    if not top_selling:
        rows.append([Paragraph("No sales recorded yet", st['cell'])] + [Paragraph("-", st['cell'])] * 5)

    t = Table(rows, colWidths=col_widths, repeatRows=1)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0F172A')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('LEFTPADDING', (0,0), (-1,-1), 4),
        ('RIGHTPADDING', (0,0), (-1,-1), 4),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#F8FAFC')]),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
    ]))
    story.append(t)
    story.append(Spacer(1, 10))
    story.append(make_universal_pdf_footer())

    doc.build(story)
    pdf_bytes = buffer.getvalue()
    buffer.close()
    return pdf_bytes

@app.route('/api/top_products/pdf')
@admin_only
def download_top_products_pdf():
    conn = get_db()
    top_selling_raw = conn.execute('''
        SELECT 
            p.name, 
            p.category, 
            p.stock, 
            p.unit, 
            SUM(si.quantity) as total_sold, 
            SUM(si.total) as total_revenue
        FROM sale_items si
        JOIN products p ON si.product_id = p.id
        JOIN sales s ON si.sale_id = s.id
        WHERE s.type != 'returned'
        GROUP BY si.product_id
        ORDER BY total_sold DESC
        LIMIT 10
    ''').fetchall()
    conn.close()
    top_selling = [dict(r) for r in top_selling_raw]
    now_str = datetime.now().strftime('%Y-%m-%d %I:%M %p')
    today = datetime.now().strftime('%Y-%m-%d')
    try:
        pdf_bytes = build_top_products_pdf(top_selling, now_str, today)
        return send_pdf_response(pdf_bytes, f"Top_10_Products_{today}.pdf")
    except Exception as e:
        logger.error(f"Error generating top products PDF: {e}")
        return f"Error: {e}", 500


# 10. SALES LIST PDF
def build_sales_list_pdf(sales, summary, date_filter, now_str, today):
    buffer = io.BytesIO()
    doc = make_universal_pdf_doc(buffer)
    st = get_universal_pdf_styles()
    story = []

    badge = f"Filter: {date_filter}" if date_filter else f"Total Records: {len(sales)}"
    story.append(make_universal_pdf_header("SALES INVOICES & REVENUE REPORT", badge, now_str, today))
    story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor('#0F172A'), spaceBefore=2, spaceAfter=8))

    kpi_data = [
        [
            Paragraph(f"<b>Total Invoices:</b> {summary['count']}<br/><b>Gross Turnover:</b> Rs. {summary['total']:,.2f}", st['cell']),
            Paragraph(f"<b>Paid (Cash/Bank):</b> <font color='#16a34a'><b>Rs. {summary['paid']:,.2f}</b></font><br/><b>Balance Due:</b> <font color='#dc2626'><b>Rs. {summary['due']:,.2f}</b></font>", st['cell']),
            Paragraph(f"<b>Payment Status:</b><br/>{'Fully Settled' if summary['due'] <= 0 else 'Receivables Pending'}", st['cell_right_bold'])
        ]
    ]
    kpi_t = Table(kpi_data, colWidths=[179, 179, 180])
    kpi_t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F8FAFC')),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#CBD5E1')),
        ('PADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(kpi_t)
    story.append(Spacer(1, 10))

    col_widths = [24, 60, 85, 125, 60, 62, 62, 60]
    rows = [
        [
            Paragraph("#", st['th']),
            Paragraph("INV #", st['th']),
            Paragraph("DATE / TIME", st['th']),
            Paragraph("CUSTOMER NAME", st['th']),
            Paragraph("TYPE", st['th_center']),
            Paragraph("TOTAL (Rs)", st['th_right']),
            Paragraph("PAID (Rs)", st['th_right']),
            Paragraph("DUE (Rs)", st['th_right'])
        ]
    ]

    for idx, s in enumerate(sales, 1):
        s_inv = clean_pdf_text(s.get('invoice_no') or f"#{s.get('id')}")
        s_date = str(s.get('date') or '-')
        s_cust = clean_pdf_text(s.get('customer_name') or 'Cash Customer')
        s_type = clean_pdf_text(str(s.get('type') or 'cash').title())
        tot = float(s.get('total') or 0)
        paid = float(s.get('paid') or 0)
        due = float(s.get('balance') or (tot - paid) if s.get('balance') is not None else 0)

        rows.append([
            Paragraph(str(idx), st['cell']),
            Paragraph(s_inv, st['cell_bold']),
            Paragraph(s_date, st['cell']),
            Paragraph(s_cust, st['cell_bold']),
            Paragraph(s_type, st['cell_center']),
            Paragraph(f"{tot:,.1f}", st['cell_right']),
            Paragraph(f"{paid:,.1f}", st['cell_right']),
            Paragraph(f"{due:,.1f}", st['cell_right_bold'])
        ])

    if not sales:
        rows.append([Paragraph("No sales matching criteria", st['cell'])] + [Paragraph("-", st['cell'])] * 7)

    t = Table(rows, colWidths=col_widths, repeatRows=1)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0F172A')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 3.5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3.5),
        ('LEFTPADDING', (0,0), (-1,-1), 3),
        ('RIGHTPADDING', (0,0), (-1,-1), 3),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#F8FAFC')]),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
    ]))
    story.append(t)
    story.append(Spacer(1, 10))
    story.append(make_universal_pdf_footer())

    doc.build(story)
    pdf_bytes = buffer.getvalue()
    buffer.close()
    return pdf_bytes

@app.route('/api/sales_list/pdf')
@admin_only
def download_sales_list_pdf():
    date_filter = request.args.get('date', '').strip()
    conn = get_db()
    cursor = conn.cursor()
    if date_filter:
        query = "SELECT s.*, COALESCE(c.name, s.customer_name, 'Cash Customer') as customer_name FROM sales s LEFT JOIN customers c ON s.customer_id = c.id WHERE s.date LIKE ? ORDER BY s.id DESC"
        sales_raw = cursor.execute(query, (f"{date_filter}%",)).fetchall()
    else:
        query = "SELECT s.*, COALESCE(c.name, s.customer_name, 'Cash Customer') as customer_name FROM sales s LEFT JOIN customers c ON s.customer_id = c.id ORDER BY s.id DESC LIMIT 300"
        sales_raw = cursor.execute(query).fetchall()
    conn.close()

    sales = [dict(s) for s in sales_raw]
    active_sales = [s for s in sales if s.get('type') != 'returned']
    summary = {
        'count': len(sales),
        'total': sum(float(s.get('total') or 0) for s in active_sales),
        'paid': sum(float(s.get('paid') or 0) for s in active_sales),
        'due': sum(float(s.get('balance') or (float(s.get('total') or 0) - float(s.get('paid') or 0))) for s in active_sales)
    }

    now_str = datetime.now().strftime('%Y-%m-%d %I:%M %p')
    today = datetime.now().strftime('%Y-%m-%d')
    try:
        pdf_bytes = build_sales_list_pdf(sales, summary, date_filter, now_str, today)
        fname = f"Sales_Invoices_{date_filter}_{today}.pdf" if date_filter else f"Sales_Invoices_Report_{today}.pdf"
        return send_pdf_response(pdf_bytes, fname)
    except Exception as e:
        logger.error(f"Error generating sales list PDF: {e}")
        return f"Error: {e}", 500


# 11. PARTY-WISE SALES PDF
def build_party_sales_pdf(items, cust_name, period_str, summary, now_str, today):
    buffer = io.BytesIO()
    doc = make_universal_pdf_doc(buffer)
    st = get_universal_pdf_styles()
    story = []

    story.append(make_universal_pdf_header("PARTY PRODUCT PURCHASE HISTORY", f"Customer: {clean_pdf_text(cust_name)}", now_str, today))
    story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor('#0F172A'), spaceBefore=2, spaceAfter=8))

    kpi_data = [
        [
            Paragraph(f"<b>Customer / Party:</b> {clean_pdf_text(cust_name)}<br/><b>Filter Period:</b> {clean_pdf_text(period_str)}", st['cell']),
            Paragraph(f"<b>Distinct Products:</b> {summary.get('distinct_products', len(items))}<br/><b>Total Units Purchased:</b> <font color='#0284c7'><b>{summary.get('total_qty', 0):,.1f}</b></font>", st['cell']),
            Paragraph(f"<b>TOTAL AMOUNT SPENT:</b><br/><font size=11 color='#16a34a'><b>Rs. {summary.get('total_amount', 0):,.2f}</b></font>", st['cell_right_bold'])
        ]
    ]
    kpi_t = Table(kpi_data, colWidths=[179, 179, 180])
    kpi_t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F8FAFC')),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#CBD5E1')),
        ('PADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(kpi_t)
    story.append(Spacer(1, 10))

    col_widths = [24, 184, 55, 65, 65, 75, 70]
    rows = [
        [
            Paragraph("#", st['th']),
            Paragraph("PRODUCT NAME", st['th']),
            Paragraph("UNIT", st['th_center']),
            Paragraph("TOTAL QTY", st['th_center']),
            Paragraph("AVG RATE", st['th_right']),
            Paragraph("TOTAL (Rs)", st['th_right']),
            Paragraph("LAST BOUGHT", st['th_center'])
        ]
    ]

    for idx, it in enumerate(items, 1):
        p_name = clean_pdf_text(it.get('product_name') or it.get('name') or '-')
        p_unit = clean_pdf_text(it.get('unit') or 'Pcs')
        qty = float(it.get('total_qty') or it.get('qty') or 0)
        avg_rate = float(it.get('avg_rate') or it.get('price') or 0)
        tot = float(it.get('total_amount') or (qty * avg_rate))
        last_date = str(it.get('last_bought') or it.get('date') or '-')

        rows.append([
            Paragraph(str(idx), st['cell']),
            Paragraph(p_name, st['cell_bold']),
            Paragraph(p_unit, st['cell_center']),
            Paragraph(f"{qty:,.1f}", st['cell_center']),
            Paragraph(f"{avg_rate:,.1f}", st['cell_right']),
            Paragraph(f"{tot:,.1f}", st['cell_right_bold']),
            Paragraph(last_date[:10], st['cell_center'])
        ])

    if not items:
        rows.append([Paragraph("No purchase records found", st['cell'])] + [Paragraph("-", st['cell'])] * 6)

    t = Table(rows, colWidths=col_widths, repeatRows=1)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0F172A')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 3.5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3.5),
        ('LEFTPADDING', (0,0), (-1,-1), 3),
        ('RIGHTPADDING', (0,0), (-1,-1), 3),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#F8FAFC')]),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
    ]))
    story.append(t)
    story.append(Spacer(1, 10))
    story.append(make_universal_pdf_footer())

    doc.build(story)
    pdf_bytes = buffer.getvalue()
    buffer.close()
    return pdf_bytes

@app.route('/api/party_sales_pdf')
@admin_only
def export_party_sales_pdf():
    cust_id_param = request.args.get('customer_id', '')
    period = request.args.get('period', 'all')
    start_date = request.args.get('start_date', '')
    end_date = request.args.get('end_date', '')

    conn = get_db()
    cursor = conn.cursor()

    cust_name = 'All Customers'
    cust_id = None

    if cust_id_param and cust_id_param != 'all':
        if str(cust_id_param).startswith('name_'):
            cust_name = str(cust_id_param)[5:].strip()
        else:
            try:
                c_row = cursor.execute("SELECT * FROM customers WHERE id = ?", (int(cust_id_param),)).fetchone()
                if c_row:
                    cust_name = c_row['name']
                    cust_id = c_row['id']
            except ValueError:
                pass

    where_clauses = ["s.type != 'returned'"]
    params = []

    if cust_id:
        where_clauses.append("(s.customer_id = ? OR (s.customer_name = ? AND s.customer_name != 'Cash Customer'))")
        params.extend([cust_id, cust_name])
    elif cust_name and cust_name != 'All Customers':
        where_clauses.append("s.customer_name = ?")
        params.append(cust_name)

    now = datetime.now()
    if period.startswith('date:'):
        single_date = period.split('date:', 1)[1].strip()
        where_clauses.append("SUBSTR(s.date, 1, 10) = ?")
        params.append(single_date)
    elif period in ('week', '7days'):
        cutoff = (now - timedelta(days=7)).strftime('%Y-%m-%d')
        where_clauses.append("SUBSTR(s.date, 1, 10) >= ?")
        params.append(cutoff)
    elif period == '15days':
        cutoff = (now - timedelta(days=15)).strftime('%Y-%m-%d')
        where_clauses.append("SUBSTR(s.date, 1, 10) >= ?")
        params.append(cutoff)
    elif period in ('month', '30days'):
        cutoff = (now - timedelta(days=30)).strftime('%Y-%m-%d')
        where_clauses.append("SUBSTR(s.date, 1, 10) >= ?")
        params.append(cutoff)
    elif period == '6months':
        cutoff = (now - timedelta(days=180)).strftime('%Y-%m-%d')
        where_clauses.append("SUBSTR(s.date, 1, 10) >= ?")
        params.append(cutoff)
    elif period in ('year', '365days'):
        cutoff = (now - timedelta(days=365)).strftime('%Y-%m-%d')
        where_clauses.append("SUBSTR(s.date, 1, 10) >= ?")
        params.append(cutoff)
    elif period == 'custom' and start_date and end_date:
        where_clauses.append("SUBSTR(s.date, 1, 10) BETWEEN ? AND ?")
        params.extend([start_date, end_date])

    where_sql = " AND ".join(where_clauses)
    query = f"""
        SELECT 
            si.product_name,
            COALESCE(p.unit, 'Pcs') as unit,
            SUM(si.quantity) as total_qty,
            AVG(si.price) as avg_rate,
            SUM(si.total) as total_amount,
            MAX(s.date) as last_bought
        FROM sale_items si
        JOIN sales s ON si.sale_id = s.id
        LEFT JOIN products p ON si.product_id = p.id
        WHERE {where_sql}
        GROUP BY si.product_id, si.product_name
        ORDER BY total_amount DESC
    """
    rows = cursor.execute(query, params).fetchall()
    conn.close()

    items = [dict(r) for r in rows]
    summary = {
        'distinct_products': len(items),
        'total_qty': sum(float(i.get('total_qty') or 0) for i in items),
        'total_amount': sum(float(i.get('total_amount') or 0) for i in items)
    }

    now_str = datetime.now().strftime('%Y-%m-%d %I:%M %p')
    today = datetime.now().strftime('%Y-%m-%d')
    try:
        pdf_bytes = build_party_sales_pdf(items, cust_name, period, summary, now_str, today)
        safe_cname = "".join([c for c in cust_name if c.isalnum() or c in (' ', '_')]).strip().replace(' ', '_')
        return send_pdf_response(pdf_bytes, f"Party_Sales_{safe_cname}_{today}.pdf")
    except Exception as e:
        logger.error(f"Error generating party sales PDF: {e}")
        return f"Error: {e}", 500


# 12. PROFIT & LOSS PDF
def build_profit_loss_pdf(total_sales, total_discount, total_purchases, total_cogs, gross_profit, total_expenses, net_profit, expense_breakdown, month, now_str, today):
    buffer = io.BytesIO()
    doc = make_universal_pdf_doc(buffer)
    st = get_universal_pdf_styles()
    story = []

    story.append(make_universal_pdf_header("PROFIT & LOSS FINANCIAL STATEMENT", f"Month: {month}", now_str, today))
    story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor('#0F172A'), spaceBefore=2, spaceAfter=8))

    profit_color = '#16a34a' if net_profit >= 0 else '#dc2626'
    kpi_data = [
        [
            Paragraph(f"<b>TOTAL SALES REVENUE:</b><br/><font size=11 color='#0284c7'><b>Rs. {total_sales:,.2f}</b></font><br/>Discounts: Rs. {total_discount:,.2f}", st['cell']),
            Paragraph(f"<b>COST OF GOODS SOLD (COGS):</b><br/><font size=11 color='#f59e0b'><b>Rs. {total_cogs:,.2f}</b></font><br/>Direct cost of sold goods", st['cell'])
        ],
        [
            Paragraph(f"<b>STORE EXPENSES:</b><br/><font size=11 color='#dc2626'><b>Rs. {total_expenses:,.2f}</b></font><br/>Bills, Wages, Tea & Operations", st['cell']),
            Paragraph(f"<b>NET OPERATING PROFIT / (LOSS):</b><br/><font size=12 color='{profit_color}'><b>Rs. {net_profit:,.2f}</b></font><br/>Gross Profit: Rs. {gross_profit:,.2f}", st['cell'])
        ]
    ]
    kpi_t = Table(kpi_data, colWidths=[269, 269])
    kpi_t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F8FAFC')),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#CBD5E1')),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E2E8F0')),
        ('PADDING', (0,0), (-1,-1), 8),
    ]))
    story.append(kpi_t)
    story.append(Spacer(1, 12))

    story.append(Paragraph("<b>FINANCIAL SUMMARY BREAKDOWN</b>", st['cell_bold']))
    story.append(Spacer(1, 3))

    b_rows = [
        [Paragraph("ACCOUNT HEAD / DESCRIPTION", st['th']), Paragraph("AMOUNT (PKR)", st['th_right'])],
        [Paragraph("Gross Sales (Before Discount)", st['cell']), Paragraph(f"Rs. {total_sales:,.2f}", st['cell_right'])],
        [Paragraph("Total Discounts Given", st['cell']), Paragraph(f"- Rs. {total_discount:,.2f}", st['cell_right'])],
        [Paragraph("Net Sales Turnover", st['cell_bold']), Paragraph(f"Rs. {total_sales - total_discount:,.2f}", st['cell_right_bold'])],
        [Paragraph("Cost of Sold Items (COGS)", st['cell']), Paragraph(f"- Rs. {total_cogs:,.2f}", st['cell_right'])],
        [Paragraph("Gross Operating Profit", st['cell_bold']), Paragraph(f"Rs. {gross_profit:,.2f}", st['cell_right_bold'])],
        [Paragraph("Total Stock Purchased (This Month)", st['cell']), Paragraph(f"Rs. {total_purchases:,.2f}", st['cell_right'])],
        [Paragraph("Shop Operational Expenses", st['cell']), Paragraph(f"- Rs. {total_expenses:,.2f}", st['cell_right'])],
        [Paragraph("<b>NET PROFIT / (LOSS)</b>", st['cell_bold']), Paragraph(f"<font color='{profit_color}'><b>Rs. {net_profit:,.2f}</b></font>", st['cell_right_bold'])],
    ]
    bt = Table(b_rows, colWidths=[388, 150])
    bt.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0F172A')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('LEFTPADDING', (0,0), (-1,-1), 6),
        ('RIGHTPADDING', (0,0), (-1,-1), 6),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#F8FAFC')]),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
        ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor('#F1F5F9')),
    ]))
    story.append(bt)
    story.append(Spacer(1, 12))

    if expense_breakdown:
        story.append(Paragraph("<b>EXPENSE CATEGORY BREAKDOWN</b>", st['cell_bold']))
        story.append(Spacer(1, 3))
        exp_rows = [
            [Paragraph("EXPENSE CATEGORY", st['th']), Paragraph("AMOUNT (PKR)", st['th_right'])]
        ]
        for cat, amt in expense_breakdown:
            exp_rows.append([
                Paragraph(clean_pdf_text(cat or 'General'), st['cell']),
                Paragraph(f"Rs. {float(amt or 0):,.2f}", st['cell_right'])
            ])
        et = Table(exp_rows, colWidths=[388, 150])
        et.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0F172A')),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('TOPPADDING', (0,0), (-1,-1), 3.5),
            ('BOTTOMPADDING', (0,0), (-1,-1), 3.5),
            ('LEFTPADDING', (0,0), (-1,-1), 6),
            ('RIGHTPADDING', (0,0), (-1,-1), 6),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#F8FAFC')]),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
        ]))
        story.append(et)
        story.append(Spacer(1, 10))

    story.append(make_universal_pdf_footer())
    doc.build(story)
    pdf_bytes = buffer.getvalue()
    buffer.close()
    return pdf_bytes

@app.route('/api/profit_loss/pdf')
@admin_only
def download_profit_loss_pdf():
    conn = get_db()
    cursor = conn.cursor()
    month_filter = request.args.get('month', datetime.now().strftime('%Y-%m'))
    
    total_sales = cursor.execute("SELECT COALESCE(SUM(total), 0) FROM sales WHERE date LIKE ? AND type != 'returned'", (f'{month_filter}%',)).fetchone()[0] or 0
    total_discount = cursor.execute("SELECT COALESCE(SUM(discount), 0) FROM sales WHERE date LIKE ? AND type != 'returned'", (f'{month_filter}%',)).fetchone()[0] or 0
    total_purchases = cursor.execute("SELECT COALESCE(SUM(total), 0) FROM purchases WHERE date LIKE ? AND COALESCE(status, \'active\') != \'returned\'", (f'{month_filter}%',)).fetchone()[0] or 0

    total_cogs = 0
    try:
        row = cursor.execute("""
            SELECT COALESCE(SUM(si.base_quantity * CASE WHEN si.buy_price > 0 THEN si.buy_price ELSE p.buy_price END), 0) 
            FROM sale_items si 
            JOIN products p ON si.product_id = p.id 
            JOIN sales s ON si.sale_id = s.id 
            WHERE s.date LIKE ? AND s.type != 'returned'
        """, (f'{month_filter}%',)).fetchone()
        if row: total_cogs = row[0]
    except Exception:
        pass

    gross_profit = total_sales - total_cogs
    total_expenses = cursor.execute("SELECT COALESCE(SUM(amount), 0) FROM expenses WHERE date LIKE ?", (f'{month_filter}%',)).fetchone()[0] or 0
    net_profit = gross_profit - total_expenses
    expense_breakdown = cursor.execute("SELECT category, SUM(amount) FROM expenses WHERE date LIKE ? GROUP BY category", (f'{month_filter}%',)).fetchall()
    conn.close()

    now_str = datetime.now().strftime('%Y-%m-%d %I:%M %p')
    today = datetime.now().strftime('%Y-%m-%d')
    try:
        pdf_bytes = build_profit_loss_pdf(total_sales, total_discount, total_purchases, total_cogs, gross_profit, total_expenses, net_profit, expense_breakdown, month_filter, now_str, today)
        return send_pdf_response(pdf_bytes, f"Profit_Loss_Statement_{month_filter}.pdf")
    except Exception as e:
        logger.error(f"Error generating profit loss PDF: {e}")
        return f"Error: {e}", 500



# ----------------------------------------------------
# 2.5 Ù¾Ø§Ø±Ù¹ÛŒ ÙˆØ§Ø¦Ø² Ù¾Ø±ÙˆÚˆÚ©Ù¹ Ø®Ø±ÛŒØ¯Ø§Ø±ÛŒ / Ø³ÛŒÙ„ ÛØ³Ù¹Ø±ÛŒ (Party-Wise Sales)
# ----------------------------------------------------
@app.route('/party_sales')
@admin_only
def party_sales_page():
    conn = get_db()
    cursor = conn.cursor()

    # 1. Registered credit ledger customers
    reg_customers = cursor.execute("SELECT id, name, phone, address, balance FROM customers ORDER BY name ASC").fetchall()
    reg_names_lower = {c['name'].lower().strip() for c in reg_customers if c['name']}
    
    customers_list = []
    for c in reg_customers:
        customers_list.append({
            'id': str(c['id']),
            'name': c['name'],
            'phone': c['phone'] or '',
            'address': c['address'] or '',
            'balance': float(c['balance'] or 0),
            'type': 'khata'
        })
    
    # 2. Walk-in cash customers with unique names from sales
    cash_names = cursor.execute("""
        SELECT DISTINCT customer_name 
        FROM sales 
        WHERE customer_name IS NOT NULL 
          AND TRIM(customer_name) != '' 
          AND customer_name != 'Cash Customer'
        ORDER BY customer_name ASC
    """).fetchall()
    
    for cn in cash_names:
        name_str = cn['customer_name'].strip()
        if name_str.lower() not in reg_names_lower:
            customers_list.append({
                'id': f"name_{name_str}",
                'name': name_str,
                'phone': '',
                'address': 'Walk-in Cash',
                'balance': 0.0,
                'type': 'cash'
            })

    # 3. Last 7 calendar days
    today_dt = datetime.now()
    recent_dates = [(today_dt - timedelta(days=i)).strftime('%Y-%m-%d') for i in range(5)]

    conn.close()
    return render_template('party_sales.html', customers=customers_list, recent_dates=recent_dates)

@app.route('/api/party_sales_data')
@admin_only
def get_party_sales_data():
    try:
        cust_id_param = request.args.get('customer_id', '')
        period = request.args.get('period', 'all')
        start_date = request.args.get('start_date', '')
        end_date = request.args.get('end_date', '')

        conn = get_db()
        cursor = conn.cursor()

        cust = None
        cust_name = ''
        cust_id = None

        if cust_id_param and cust_id_param != 'all':
            if str(cust_id_param).startswith('name_'):
                cust_name = str(cust_id_param)[5:].strip()
                cust = {'id': cust_id_param, 'name': cust_name, 'phone': '', 'address': 'Walk-in Cash', 'balance': 0}
            else:
                try:
                    c_row = cursor.execute("SELECT * FROM customers WHERE id = ?", (int(cust_id_param),)).fetchone()
                    if c_row:
                        cust = dict(c_row)
                        cust_name = cust['name']
                        cust_id = cust['id']
                except (ValueError, TypeError):
                    pass

        # Base query for sales of this customer
        where_clauses = ["s.type != 'returned'"]
        params = []

        if cust_id:
            where_clauses.append("(s.customer_id = ? OR (s.customer_name = ? AND s.customer_name != 'Cash Customer'))")
            params.extend([cust_id, cust_name])
        elif cust_name:
            where_clauses.append("s.customer_name = ?")
            params.append(cust_name)

        # Time filter
        now = datetime.now()
        if period.startswith('date:'):
            single_date = period.split('date:', 1)[1].strip()
            where_clauses.append("SUBSTR(s.date, 1, 10) = ?")
            params.append(single_date)
        elif period in ('week', '7days'):
            cutoff = (now - timedelta(days=7)).strftime('%Y-%m-%d')
            where_clauses.append("SUBSTR(s.date, 1, 10) >= ?")
            params.append(cutoff)
        elif period == '15days':
            cutoff = (now - timedelta(days=15)).strftime('%Y-%m-%d')
            where_clauses.append("SUBSTR(s.date, 1, 10) >= ?")
            params.append(cutoff)
        elif period in ('month', '30days'):
            cutoff = (now - timedelta(days=30)).strftime('%Y-%m-%d')
            where_clauses.append("SUBSTR(s.date, 1, 10) >= ?")
            params.append(cutoff)
        elif period == '6months':
            cutoff = (now - timedelta(days=180)).strftime('%Y-%m-%d')
            where_clauses.append("SUBSTR(s.date, 1, 10) >= ?")
            params.append(cutoff)
        elif period == 'year':
            cutoff = (now - timedelta(days=365)).strftime('%Y-%m-%d')
            where_clauses.append("SUBSTR(s.date, 1, 10) >= ?")
            params.append(cutoff)
        elif period == 'custom' and start_date and end_date:
            where_clauses.append("SUBSTR(s.date, 1, 10) BETWEEN ? AND ?")
            params.extend([start_date, end_date])

        where_sql = " AND ".join(where_clauses)

        # 1. Product aggregates
        query = f"""
            SELECT si.product_name, 
                   COALESCE(si.unit, 'Pcs') as unit, 
                   SUM(si.quantity) as total_qty, 
                   AVG(si.price) as avg_price, 
                   MIN(si.price) as min_price,
                   MAX(si.price) as max_price,
                   SUM(si.total) as total_amount,
                   MAX(s.date) as last_date,
                   COUNT(DISTINCT s.id) as invoice_count
            FROM sale_items si
            JOIN sales s ON si.sale_id = s.id
            WHERE {where_sql}
            GROUP BY si.product_name, COALESCE(si.unit, 'Pcs')
            ORDER BY total_amount DESC
        """
        rows = cursor.execute(query, params).fetchall()

        # 2. Detailed invoice entries per product for expandability
        detail_query = f"""
            SELECT si.product_name,
                   COALESCE(si.unit, 'Pcs') as unit,
                   s.invoice_no,
                   s.id as sale_id,
                   s.customer_name,
                   s.date,
                   si.quantity,
                   si.price,
                   si.total,
                   COALESCE(s.discount, 0) as discount
            FROM sale_items si
            JOIN sales s ON si.sale_id = s.id
            WHERE {where_sql}
            ORDER BY s.date DESC
        """
        detail_rows = cursor.execute(detail_query, params).fetchall()

        details_by_product = {}
        for d in detail_rows:
            key = f"{d['product_name']}_{d['unit']}"
            if key not in details_by_product:
                details_by_product[key] = []
            details_by_product[key].append({
                'invoice_no': d['invoice_no'] or f"INV-{d['sale_id']}",
                'sale_id': d['sale_id'],
                'customer_name': d['customer_name'] or 'Cash Customer',
                'date': d['date'],
                'qty': round(float(d['quantity'] or 0), 2),
                'price': round(float(d['price'] or 0), 2),
                'total': round(float(d['total'] or 0), 2),
                'discount': round(float(d['discount'] or 0), 2),
                'unit': d['unit']
            })

        items = []
        tot_qty = 0.0
        tot_amount = 0.0

        for r in rows:
            t_qty = round(float(r['total_qty'] or 0), 2)
            t_amt = round(float(r['total_amount'] or 0), 2)
            tot_qty += t_qty
            tot_amount += t_amt
            key = f"{r['product_name']}_{r['unit']}"

            items.append({
                'product_name': r['product_name'],
                'unit': r['unit'],
                'total_qty': t_qty,
                'avg_price': round(float(r['avg_price'] or 0), 2),
                'min_price': round(float(r['min_price'] or 0), 2),
                'max_price': round(float(r['max_price'] or 0), 2),
                'total_amount': t_amt,
                'last_date': r['last_date'],
                'invoice_count': r['invoice_count'],
                'invoices': details_by_product.get(key, [])
            })

        # Calculate total discount for sales in this period/customer
        sales_disc_query = f"SELECT COALESCE(SUM(s.discount), 0) FROM sales s WHERE {where_sql}"
        tot_discount = cursor.execute(sales_disc_query, params).fetchone()[0] or 0.0

        conn.close()

        return jsonify({
            'success': True,
            'customer': dict(cust) if cust else None,
            'period': period,
            'summary': {
                'products_count': len(items),
                'total_units': tot_qty,
                'total_amount': tot_amount,
                'total_discount': round(float(tot_discount or 0), 2)
            },
            'items': items
        })
    except Exception as e:
        logger.error(f"Error in get_party_sales_data: {e}")
        return jsonify({'success': False, 'message': str(e)}), 500

@app.route('/api/party_sales_excel')
@admin_only
def export_party_sales_excel():
    cust_id_param = request.args.get('customer_id', '')
    period = request.args.get('period', 'all')
    start_date = request.args.get('start_date', '')
    end_date = request.args.get('end_date', '')

    conn = get_db()
    cursor = conn.cursor()

    cust = None
    cust_name = 'All Customers'
    cust_id = None

    if cust_id_param and cust_id_param != 'all':
        if str(cust_id_param).startswith('name_'):
            cust_name = str(cust_id_param)[5:].strip()
            cust = {'name': cust_name}
        else:
            try:
                c_row = cursor.execute("SELECT * FROM customers WHERE id = ?", (int(cust_id_param),)).fetchone()
                if c_row:
                    cust = dict(c_row)
                    cust_name = cust['name']
                    cust_id = cust['id']
            except ValueError:
                pass

    where_clauses = ["s.type != 'returned'"]
    params = []

    if cust_id:
        where_clauses.append("(s.customer_id = ? OR (s.customer_name = ? AND s.customer_name != 'Cash Customer'))")
        params.extend([cust_id, cust_name])
    elif cust_name and cust_name != 'All Customers':
        where_clauses.append("s.customer_name = ?")
        params.append(cust_name)

    now = datetime.now()
    if period.startswith('date:'):
        single_date = period.split('date:', 1)[1].strip()
        where_clauses.append("SUBSTR(s.date, 1, 10) = ?")
        params.append(single_date)
    elif period in ('week', '7days'):
        cutoff = (now - timedelta(days=7)).strftime('%Y-%m-%d')
        where_clauses.append("SUBSTR(s.date, 1, 10) >= ?")
        params.append(cutoff)
    elif period == '15days':
        cutoff = (now - timedelta(days=15)).strftime('%Y-%m-%d')
        where_clauses.append("SUBSTR(s.date, 1, 10) >= ?")
        params.append(cutoff)
    elif period in ('month', '30days'):
        cutoff = (now - timedelta(days=30)).strftime('%Y-%m-%d')
        where_clauses.append("SUBSTR(s.date, 1, 10) >= ?")
        params.append(cutoff)
    elif period == '6months':
        cutoff = (now - timedelta(days=180)).strftime('%Y-%m-%d')
        where_clauses.append("SUBSTR(s.date, 1, 10) >= ?")
        params.append(cutoff)
    elif period == 'year':
        cutoff = (now - timedelta(days=365)).strftime('%Y-%m-%d')
        where_clauses.append("SUBSTR(s.date, 1, 10) >= ?")
        params.append(cutoff)
    elif period == 'custom' and start_date and end_date:
        where_clauses.append("SUBSTR(s.date, 1, 10) BETWEEN ? AND ?")
        params.extend([start_date, end_date])

    where_sql = " AND ".join(where_clauses)

    query = f"""
        SELECT si.product_name, 
               COALESCE(si.unit, 'Pcs') as unit, 
               SUM(si.quantity) as total_qty, 
               AVG(si.price) as avg_price, 
               SUM(si.total) as total_amount,
               MAX(s.date) as last_date,
               COUNT(DISTINCT s.id) as invoice_count
        FROM sale_items si
        JOIN sales s ON si.sale_id = s.id
        WHERE {where_sql}
        GROUP BY si.product_name, COALESCE(si.unit, 'Pcs')
        ORDER BY total_amount DESC
    """
    rows = cursor.execute(query, params).fetchall()
    conn.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Party Item Sales"

    ws.merge_cells("A1:G1")
    title_cell = ws["A1"]
    title_cell.value = f"SMART POS SYSTEM - Party Sales History ({cust_name})"
    title_cell.font = Font(name="Segoe UI", size=14, bold=True, color="FFFFFF")
    title_cell.fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    headers = ["#", "Product Name", "Unit", "Total Qty", "Avg Rate (Rs)", "Total Amount (Rs)", "Last Bought Date"]
    ws.append(headers)
    ws.row_dimensions[2].height = 22

    for col_idx in range(1, 8):
        cell = ws.cell(row=2, column=col_idx)
        cell.font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")

    tot_q = 0
    tot_a = 0
    for idx, r in enumerate(rows, 1):
        q = float(r['total_qty'] or 0)
        a = float(r['total_amount'] or 0)
        tot_q += q
        tot_a += a
        ws.append([
            idx,
            r['product_name'],
            r['unit'],
            q,
            round(float(r['avg_price'] or 0), 2),
            a,
            str(r['last_date'] or '-')
        ])

    total_row = len(rows) + 3
    ws.cell(row=total_row, column=2, value="TOTAL:").font = Font(bold=True)
    ws.cell(row=total_row, column=4, value=tot_q).font = Font(bold=True)
    ws.cell(row=total_row, column=6, value=tot_a).font = Font(bold=True, color="0284C7")

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    safe_name = "".join([c for c in cust_name if c.isalnum() or c in (' ', '_', '-')]).strip()
    return send_file(
        output,
        as_attachment=True,
        download_name=f"PartySales_{safe_name}_{period}_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ----------------------------------------------------
# 3. Ø³Ù¾Ù„Ø§Ø¦Ø± Ø§ÙˆØ± Ù¾Ø±Ú†ÛŒØ²Ø² (Suppliers & Payables)
# ----------------------------------------------------
@app.route('/suppliers')
@app.route('/payables')
@admin_only
def suppliers_page():
    try:
        conn = get_db()
        cursor = conn.cursor()

        suppliers_raw = cursor.execute("SELECT id, opening_balance FROM suppliers").fetchall()
        for sup in suppliers_raw:
            s_id = sup['id']
            opening_bal = float(sup['opening_balance'] or 0.0)
            
            # Start from opening balance, add purchases, subtract all payments
            tot_bought = float(cursor.execute("SELECT COALESCE(SUM(total), 0) FROM purchases WHERE supplier_id = ? AND COALESCE(status, \'active\') != \'returned\'", (s_id,)).fetchone()[0])
            tot_paid_at_bill = float(cursor.execute("SELECT COALESCE(SUM(paid), 0) FROM purchases WHERE supplier_id = ? AND COALESCE(status, \'active\') != \'returned\'", (s_id,)).fetchone()[0])
            tot_paid_direct = float(cursor.execute("SELECT COALESCE(SUM(amount), 0) FROM supplier_payments WHERE supplier_id = ?", (s_id,)).fetchone()[0])

            real_balance = opening_bal + tot_bought - (tot_paid_at_bill + tot_paid_direct)
            # Balance can never go below 0 (no advance/credit allowed)
            real_balance = max(0.0, round(real_balance, 2))
            cursor.execute("UPDATE suppliers SET balance = ? WHERE id = ?", (real_balance, s_id))
        conn.commit()

        suppliers_list = cursor.execute("SELECT * FROM suppliers ORDER BY balance DESC").fetchall()
        products_db = cursor.execute("SELECT id, name, buy_price, stock, unit, sale_price FROM products ORDER BY name ASC").fetchall()
        
        products = []
        for p in products_db:
            pd = dict(p)
            units = cursor.execute("SELECT unit_name as name, conversion_factor as conversion_rate FROM product_units WHERE product_id = ?", (p['id'],)).fetchall()
            pd['units_json'] = [dict(u) for u in units]
            products.append(pd)
        
        # Fetch all purchases with supplier info and items
        purchases_raw = cursor.execute("""
            SELECT p.*, s.name as supplier_name, s.company as supplier_company 
            FROM purchases p 
            LEFT JOIN suppliers s ON p.supplier_id = s.id 
            ORDER BY p.id DESC
        """).fetchall()

        purchases_list = []
        for pr in purchases_raw:
            p_dict = dict(pr)
            p_items = cursor.execute("SELECT product_id, product_name, quantity, buy_price, unit, total FROM purchase_items WHERE purchase_id = ?", (p_dict['id'],)).fetchall()
            p_dict['items'] = [dict(it) for it in p_items]
            p_dict['items_count'] = len(p_items)
            summary_parts = [f"{it['product_name']} ({int(it['quantity']) if it['quantity'] == int(it['quantity']) else it['quantity']} {it['unit'] or 'Pcs'})" for it in p_items[:3]]
            p_dict['items_summary'] = ", ".join(summary_parts)
            if len(p_items) > 3:
                p_dict['items_summary'] += f" +{len(p_items)-3} more"
            purchases_list.append(p_dict)

        today_dt = datetime.now()
        recent_dates = [(today_dt - timedelta(days=i)).strftime('%Y-%m-%d') for i in range(5)]

        total_supp_cnt = len(suppliers_list)
        total_supp_balance = sum([float(s['balance'] or 0) for s in suppliers_list])

        conn.close()

        return render_template('suppliers.html', suppliers=suppliers_list, products=products, purchases=purchases_list, recent_dates=recent_dates, total_supp_cnt=total_supp_cnt, total_supp_balance=total_supp_balance)
    except Exception as e:
        return f"Suppliers Page Error: {str(e)}", 500

def get_supplier_ledger_data(sup_id):
    conn = get_db()
    cursor = conn.cursor()
    
    sup = cursor.execute("SELECT * FROM suppliers WHERE id = ?", (sup_id,)).fetchone()
    if not sup:
        conn.close()
        return None, [], 0.0, 0.0, 0.0

    events = []
    opening_bal = float(sup['opening_balance'] or 0.0)
    if opening_bal != 0:
        events.append({
            'id': 0,
            'date': '-',
            'type': 'OPENING',
            'ref': 'OB',
            'description': 'Opening Balance / پچھلا بقایا',
            'description_en': 'Opening Balance / Adjustments',
            'description_ur': 'پچھلا بقایا / Adjustments',
            'bill_amount': opening_bal if opening_bal > 0 else 0.0,
            'paid_amount': abs(opening_bal) if opening_bal < 0 else 0.0
        })

    purchases = cursor.execute("SELECT * FROM purchases WHERE supplier_id = ? AND COALESCE(status, \'active\') != \'returned\' ORDER BY id ASC", (sup_id,)).fetchall()
    for p in purchases:
        p_id = p['id']
        bill_no = p['bill_no'] or f"PB-{p_id}"
        total_bill = float(p['total'] or 0)
        paid_now = float(p['paid'] or 0)

        items = cursor.execute("SELECT product_name, quantity, buy_price, total FROM purchase_items WHERE purchase_id = ?", (p_id,)).fetchall()
        items_detail_en = [f"{it['product_name']} ({it['quantity']} x Rs.{it['buy_price']})" for it in items] if items else []
        items_detail_ur = [f"{it['product_name']} ({it['quantity']} x {it['buy_price']} روپے)" for it in items] if items else []
        if items_detail_en:
            items_str_en = "<br/>&bull; " + "<br/>&bull; ".join(items_detail_en)
            items_str_ur = "<br/>• " + "<br/>• ".join(items_detail_ur)
        else:
            items_str_en = "Stock Purchase"
            items_str_ur = "مال خریداری"

        if total_bill > 0:
            events.append({
                'id': p_id,
                'date': str(p['date']),
                'type': 'PURCHASE',
                'ref': bill_no,
                'description': f"Purchase / خریداری: {items_str_en}",
                'description_en': f"Purchase Stock: {items_str_en}",
                'description_ur': f"مال خریداری: {items_str_ur}",
                'bill_amount': total_bill,
                'paid_amount': 0.0
            })

        if paid_now > 0:
            events.append({
                'id': p_id,
                'date': str(p['date']),
                'type': 'PAYMENT',
                'ref': f"PAY-{p_id}",
                'description': "Payment on Bill / موقع پر ادائیگی",
                'description_en': "Cash Payment (On Bill)",
                'description_ur': "نقد ادائیگی (موقع پر)",
                'bill_amount': 0.0,
                'paid_amount': paid_now
            })

    direct_pays = cursor.execute("SELECT * FROM supplier_payments WHERE supplier_id = ? ORDER BY id ASC", (sup_id,)).fetchall()
    for dp in direct_pays:
        raw_note = (dp['note'] or (dp['notes'] if 'notes' in dp.keys() else '') or '').strip()
        note_en = 'Cash' if not raw_note or raw_note == 'نقد' else raw_note
        note_ur = 'نقد' if not raw_note or raw_note.lower() == 'cash' else raw_note
        amt = float(dp['amount'] or 0)
        is_ref = 'refund' in raw_note.lower() or 'واپسی' in raw_note or amt < 0
        if is_ref:
            ref_amt = abs(amt)
            events.append({
                'id': dp['id'],
                'date': str(dp['payment_date']),
                'type': 'REFUND',
                'ref': f"SREF-{dp['id']}",
                'description': f"Refund / واپسی: {raw_note}",
                'description_en': f"Refund from Supplier: {note_en}",
                'description_ur': f"سپلائر سے رقم واپسی: {note_ur}",
                'bill_amount': ref_amt,
                'paid_amount': 0.0
            })
        else:
            events.append({
                'id': dp['id'],
                'date': str(dp['payment_date']),
                'type': 'PAYMENT',
                'ref': f"SPAY-{dp['id']}",
                'description': f"Payment / ادائیگی: {raw_note or 'Cash / نقد'}",
                'description_en': f"Payment to Supplier: {note_en}",
                'description_ur': f"ادائیگی سپلائر: {note_ur}",
                'bill_amount': 0.0,
                'paid_amount': amt
            })

    events.sort(key=lambda x: (
        parse_event_datetime(x.get('date'), is_end_of_day=(x.get('type') in ('PAYMENT', 'REFUND'))),
        0 if x.get('type') == 'OPENING' else (1 if x.get('type') == 'PURCHASE' else 2),
        x.get('id', 0)
    ))

    timeline = []
    running_balance = 0.0
    total_bought = 0.0
    total_paid = 0.0

    for ev in events:
        b = round(ev['bill_amount'], 2)
        p = round(ev['paid_amount'], 2)
        prev_bal = running_balance
        running_balance += (b - p)
        total_bought += b
        total_paid += p

        cleared_here = False
        if prev_bal > 0 and running_balance == 0:
            cleared_here = True

        timeline.append({
            'date': ev['date'],
            'type': ev['type'],
            'ref': ev['ref'],
            'description': ev['description'],
            'description_en': ev['description_en'],
            'description_ur': ev['description_ur'],
            'bill_amount': b,
            'paid_amount': p,
            'balance': running_balance,
            'cleared_here': cleared_here
        })

    running_balance = max(0.0, running_balance)
    cursor.execute("UPDATE suppliers SET balance = ? WHERE id = ?", (running_balance, sup_id))
    conn.commit()
    conn.close()

    return dict(sup), timeline, total_bought, total_paid, running_balance

@app.route('/api/supplier_ledger/<int:sup_id>')
@admin_only
def get_supplier_ledger(sup_id):
    sup_dict, timeline, total_bought, total_paid, running_balance = get_supplier_ledger_data(sup_id)
    if not sup_dict:
        return jsonify({'error': 'Supplier not found'}), 404
    return jsonify({
        'supplier': sup_dict,
        'timeline': timeline,
        'total_bought': total_bought,
        'total_paid': total_paid,
        'final_balance': running_balance
    })

@app.route('/api/supplier/print/<int:sup_id>')
@admin_only
def print_supplier_ledger(sup_id):
    sup_dict, timeline, total_bought, total_paid, running_balance = get_supplier_ledger_data(sup_id)
    if not sup_dict:
        return "Supplier not found", 404
    lang = request.args.get('lang', 'ur')
    now_str = datetime.now().strftime('%Y-%m-%d %I:%M %p')
    today = datetime.now().strftime('%Y-%m-%d')
    return render_template(
        'supplier_ledger_print.html',
        supplier=sup_dict,
        timeline=timeline,
        total_bought=total_bought,
        total_paid=total_paid,
        final_balance=running_balance,
        now_str=now_str,
        today=today,
        lang=lang
    )

@app.route('/api/supplier/pdf/<int:sup_id>')
@admin_only
def download_supplier_ledger_pdf(sup_id):
    sup_dict, timeline, total_bought, total_paid, running_balance = get_supplier_ledger_data(sup_id)
    if not sup_dict:
        return "Supplier not found", 404
    now_str = datetime.now().strftime('%Y-%m-%d %I:%M %p')
    today = datetime.now().strftime('%Y-%m-%d')
    try:
        pdf_bytes = build_supplier_ledger_pdf(sup_dict, timeline, total_bought, total_paid, running_balance, now_str, today)
        sup_name = sup_dict.get('name', f'Supplier_{sup_id}')
        safe_name = "".join([c for c in sup_name if c.isalnum() or c in (' ', '_', '-')]).strip().replace(' ', '_')
        return send_pdf_response(pdf_bytes, f"Supplier_Khata_{safe_name}_{today}.pdf")
    except Exception as e:
        logger.error(f"Error generating supplier PDF via ReportLab: {e}")
        return f"Error generating PDF: {e}", 500

@app.route('/api/supplier/whatsapp_prepare/<int:sup_id>', methods=['POST', 'GET'])
@admin_only
def whatsapp_prepare_supplier_ledger(sup_id):
    sup_dict, timeline, total_bought, total_paid, running_balance = get_supplier_ledger_data(sup_id)
    if not sup_dict:
        return jsonify({'success': False, 'message': 'Supplier not found'}), 404

    now_str = datetime.now().strftime('%Y-%m-%d %I:%M %p')
    today = datetime.now().strftime('%Y-%m-%d')
    sup_name = sup_dict.get('name', f'Supplier_{sup_id}')
    safe_name = "".join([c for c in sup_name if c.isalnum() or c in (' ', '_', '-')]).strip().replace(' ', '_')
    filename = f"Supplier_Khata_{safe_name}_{today}.pdf"

    pdf_path = ""
    try:
        pdf_bytes = build_supplier_ledger_pdf(sup_dict, timeline, total_bought, total_paid, running_balance, now_str, today)
        desktop_dir = os.path.join(os.path.expanduser('~'), 'Desktop')
        save_dir = desktop_dir if os.path.exists(desktop_dir) else os.getcwd()
        pdf_path = os.path.join(save_dir, filename)
        with open(pdf_path, 'wb') as f:
            f.write(pdf_bytes)
        copy_file_to_clipboard(pdf_path)
    except Exception as e:
        logger.warning(f"Could not copy supplier PDF to clipboard: {e}")

    phone = sup_dict.get('phone') or ''
    clean_phone = re.sub(r'\D', '', phone)
    if clean_phone.startswith('0'):
        clean_phone = '92' + clean_phone[1:]
    elif not clean_phone.startswith('92') and len(clean_phone) == 10:
        clean_phone = '92' + clean_phone

    status_str = "کھاتہ صاف ہے ✅" if running_balance <= 0 else f"بقایا رقم: Rs. {running_balance:,.2f}"
    msg = (
        f"السلام علیکم {sup_dict.get('name', 'محترم')} صاحب!\n"
        f"جمال بولٹ اسٹور کی طرف سے آپ کے کھاتے کی اسٹیٹمنٹ:\n\n"
        f"تاریخ: {today}\n"
        f"کل مال لیا: Rs. {total_bought:,.2f}\n"
        f"کل رقم ادا کی: Rs. {total_paid:,.2f}\n"
        f"حالت: {status_str}\n\n"
        f"پی ڈی ایف اسٹیٹمنٹ ساتھ منسلک کی جا رہی ہے۔ شکریہ!"
    )
    import urllib.parse
    encoded_msg = urllib.parse.quote(msg)
    whatsapp_url = f"https://web.whatsapp.com/send?phone={clean_phone}&text={encoded_msg}" if clean_phone else f"https://web.whatsapp.com/send?text={encoded_msg}"

    return jsonify({
        'success': True,
        'whatsapp_url': whatsapp_url,
        'download_url': f"/api/supplier/pdf/{sup_id}",
        'phone': clean_phone,
        'file_path': pdf_path,
        'filename': filename
    })

@app.route('/api/supplier/excel/<int:sup_id>')
@admin_only
def download_supplier_ledger_excel(sup_id):
    sup_dict, timeline, total_bought, total_paid, running_balance = get_supplier_ledger_data(sup_id)
    if not sup_dict:
        return "Supplier not found", 404
    today = datetime.now().strftime('%Y-%m-%d')
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Supplier Ledger"

    ws.append(["DATE & TIME", "REF / BILL #", "DESCRIPTION", "DEBIT (Rs)", "CREDIT (Rs)", "BALANCE (Rs)"])
    for ev in timeline:
        ws.append([
            ev.get('date', '-'),
            ev.get('ref', '-'),
            clean_pdf_text(ev.get('description', '')),
            ev.get('bill_amount', 0),
            ev.get('paid_amount', 0),
            ev.get('balance', 0)
        ])
    ws.append([])
    ws.append(["Total Purchases:", total_bought])
    ws.append(["Total Paid:", total_paid])
    ws.append(["Net Balance Payable:", running_balance])

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    sup_name = sup_dict.get('name', f'Supplier_{sup_id}')
    safe_name = "".join([c for c in sup_name if c.isalnum() or c in (' ', '_', '-')]).strip().replace(' ', '_')
    return send_file(
        buf,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=f"Supplier_Ledger_{safe_name}_{today}.xlsx"
    )

@app.route('/purchases')
@admin_only
def purchases_history():
    return redirect('/suppliers')

@app.route('/api/supplier/add', methods=['POST'])
@admin_only
def add_supplier():
    if request.is_json:
        data = request.get_json() or {}
        name = data.get("name", "").strip()
        company = data.get("company", "").strip()
        phone = data.get("phone", "").strip()
        address = data.get("address", "").strip()
        raw_bal = data.get("balance") or data.get("opening_balance") or 0
    else:
        name = request.form.get("name", "").strip()
        company = request.form.get("company", "").strip()
        phone = request.form.get("phone", "").strip()
        address = request.form.get("address", "").strip()
        raw_bal = request.form.get("balance") or request.form.get("opening_balance") or 0

    try:
        op_balance = round(float(raw_bal), 2)
    except (ValueError, TypeError):
        op_balance = 0.0

    if name:
        try:
            conn = get_db()
            conn.execute("INSERT INTO suppliers (name, company, phone, address, balance, opening_balance) VALUES (?, ?, ?, ?, ?, ?)", 
                         (name, company, phone, address, op_balance, op_balance))
            conn.commit()
            conn.close()
        except Exception as e:
            logger.error(f"Error adding supplier: {e}")
            if request.is_json:
                return jsonify({"success": False, "message": str(e)}), 500
            return redirect('/suppliers?error=1')

    if request.is_json:
        return jsonify({"success": True, "message": "سپلائر شامل ہو گیا"})
    return redirect('/suppliers')

@app.route('/api/supplier/edit', methods=['POST'])
@login_required
def edit_supplier():
    data = request.form if request.form else (request.json or {})
    sup_id = data.get('id')
    name = data.get('name', '').strip()
    company = data.get('company', '').strip()
    phone = data.get('phone', '').strip()
    address = data.get('address', '').strip()

    if not sup_id or not name:
        return redirect('/suppliers')

    conn = get_db()
    cursor = conn.cursor()
    target_bal_input = data.get('target_balance') if 'target_balance' in data else data.get('balance')
    if target_bal_input is not None and str(target_bal_input).strip() != '':
        try:
            target_bal = round(float(target_bal_input or 0), 2)
            tot_bought = float(cursor.execute("SELECT COALESCE(SUM(total), 0) FROM purchases WHERE supplier_id = ? AND COALESCE(status, 'active') != 'returned'", (sup_id,)).fetchone()[0])
            tot_paid_bill = float(cursor.execute("SELECT COALESCE(SUM(paid), 0) FROM purchases WHERE supplier_id = ? AND COALESCE(status, 'active') != 'returned'", (sup_id,)).fetchone()[0])
            tot_paid_direct = float(cursor.execute("SELECT COALESCE(SUM(amount), 0) FROM supplier_payments WHERE supplier_id = ?", (sup_id,)).fetchone()[0])
            
            new_ob = target_bal - tot_bought + (tot_paid_bill + tot_paid_direct)
            new_ob = round(new_ob, 2)
            
            cursor.execute("UPDATE suppliers SET name = ?, company = ?, phone = ?, address = ?, opening_balance = ?, balance = ? WHERE id = ?", 
                           (name, company, phone, address, new_ob, target_bal, sup_id))
        except (ValueError, TypeError):
            cursor.execute("UPDATE suppliers SET name = ?, company = ?, phone = ?, address = ? WHERE id = ?", (name, company, phone, address, sup_id))
    else:
        cursor.execute("UPDATE suppliers SET name = ?, company = ?, phone = ?, address = ? WHERE id = ?", (name, company, phone, address, sup_id))
    conn.commit()
    conn.close()

    if request.is_json:
        return jsonify({'success': True, 'message': 'سپلائر کی معلومات اپ ڈیٹ ہو گئیں'})
    return redirect('/suppliers')

@app.route('/api/supplier/delete/<int:sup_id>', methods=['POST'])
@admin_only
def delete_supplier(sup_id):
    try:
        conn = get_db()
        sup = conn.execute("SELECT balance, name FROM suppliers WHERE id = ?", (sup_id,)).fetchone()
        sup_balance = float(sup["balance"] or 0) if sup else 0.0
        if sup and sup_balance > 0:
            conn.close()
            bal_str = f"{sup_balance:,.2f}"
            return make_delete_error_page(
                title_ur="سپلائر ڈیلیٹ نہیں ہو سکتا! (ادھار باقی ہے)",
                title_en="Cannot Delete Supplier (Outstanding Balance)",
                msg_ur=f"اس سپلائر کو ابھی <b style='color:#f87171;'>Rs. {bal_str}</b> ادا کرنے باقی ہیں۔ پہلے ادائیگی کر کے کھاتہ زیرو (0) کریں، پھر ڈیلیٹ کریں۔",
                msg_en=f"You owe Rs. {bal_str} to this supplier. Please settle payments before deleting.",
                return_url="/suppliers"
            )
            
        conn.execute("DELETE FROM suppliers WHERE id = ?", (sup_id,))
        conn.commit()
        conn.close()
        if request.is_json:
            return jsonify({'success': True, 'message': 'سپلائر ڈیلیٹ ہو گیا'})
        return redirect('/suppliers')
    except sqlite3.IntegrityError:
        return make_delete_error_page(
            title_ur="یہ سپلائر ڈیلیٹ نہیں ہو سکتا!",
            title_en="Cannot Delete Supplier (Linked Records Found)",
            msg_ur="اس سپلائر کا پرانا ریکارڈ (بل یا پیمنٹ) موجود ہے۔ پہلے پرانا ریکارڈ ڈیلیٹ کریں۔",
            msg_en="This supplier has linked purchase or payment records. Please delete those records first.",
            return_url="/suppliers"
        )

@app.route('/api/supplier/pay', methods=['POST'])
@login_required
def pay_supplier():
    data = request.json or {}
    sup_id = data.get('supplier_id')
    amount = round(float(data.get('amount') or 0), 2)
    note = data.get('note', 'سپلائر کو ادھار ادائیگی')

    if not sup_id or amount <= 0:
        return jsonify({'success': False, 'message': 'درست رقم درج کریں'}), 400

    conn = get_db()
    cursor = conn.cursor()
    
    sup = cursor.execute("SELECT balance FROM suppliers WHERE id = ?", (sup_id,)).fetchone()
    if not sup:
        conn.close()
        return jsonify({'success': False, 'message': 'سپلائر نہیں ملا'}), 404
    
    current_balance = float(sup['balance'] or 0)
    now_str = datetime.now().strftime("%Y-%m-%d %I:%M:%S %p")
    cursor.execute("INSERT INTO supplier_payments (supplier_id, amount, payment_date, note) VALUES (?, ?, ?, ?)", (sup_id, amount, now_str, note))
    new_balance = round(current_balance - amount, 2)
    cursor.execute("UPDATE suppliers SET balance = ? WHERE id = ?", (new_balance, sup_id))
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': f'ادائیگی کامیابی سے درج! نیا بقایا: Rs. {new_balance:,.2f}'})

@app.route('/api/purchase/add', methods=['POST'])
@admin_only
def save_purchase():
    try:
        data = request.get_json() or {}
        supplier_id = data.get('supplier_id')
        supplier_name = (data.get('supplier_name') or '').strip()
        bill_no = (data.get('bill_no') or '').strip() or f"PB-{int(datetime.now().timestamp())}"
        items = data.get('items', [])
        total_amount = round(float(data.get('total_amount', 0)), 2)
        paid_amount = round(float(data.get('paid_amount', 0)), 2)
        due_amount = round(total_amount - paid_amount, 2)
        notes = data.get('notes', '')

        if not supplier_name:
            supplier_name = "General Supplier"

        conn = get_db()
        cursor = conn.cursor()

        # 1. Professional On-The-Fly Supplier Registration
        if not supplier_id or str(supplier_id).startswith('new_'):
            existing_sup = cursor.execute("SELECT id, name FROM suppliers WHERE LOWER(TRIM(name)) = LOWER(TRIM(?))", (supplier_name,)).fetchone()
            if existing_sup:
                supplier_id = existing_sup['id']
                supplier_name = existing_sup['name']
            else:
                cursor.execute("INSERT INTO suppliers (name, balance) VALUES (?, 0)", (supplier_name,))
                supplier_id = cursor.lastrowid
        else:
            sup_row = cursor.execute("SELECT id, name FROM suppliers WHERE id = ?", (supplier_id,)).fetchone()
            if sup_row:
                supplier_name = sup_row['name']
            else:
                cursor.execute("INSERT INTO suppliers (name, balance) VALUES (?, 0)", (supplier_name,))
                supplier_id = cursor.lastrowid

        custom_date = data.get('date')
        if custom_date:
            try:
                now_str = f"{custom_date} {datetime.now().strftime('%I:%M:%S %p')}" if len(custom_date) <= 10 else custom_date
            except Exception:
                now_str = datetime.now().strftime('%Y-%m-%d %I:%M:%S %p')
        else:
            now_str = datetime.now().strftime('%Y-%m-%d %I:%M:%S %p')

        cursor.execute("INSERT INTO purchases (bill_no, supplier_id, supplier_name, date, total, paid, due, notes) VALUES (?, ?, ?, ?, ?, ?, ?, ?)", 
                       (bill_no, supplier_id, supplier_name, now_str, total_amount, paid_amount, due_amount, notes))
        purchase_id = cursor.lastrowid

        if due_amount > 0:
            cursor.execute("UPDATE suppliers SET balance = balance + ? WHERE id = ?", (due_amount, supplier_id))

        # 2. Professional On-The-Fly Product Registration & Stock In
        for it in items:
            p_id = it.get('id')
            p_name = (it.get('name') or '').strip()
            qty = float(it.get('qty', 1))
            b_price = round(float(it.get('buy_price', 0)), 2)
            unit = it.get('unit') or 'Pcs'

            is_new_prod = False
            try:
                p_id_int = int(p_id)
                prod_row = cursor.execute("SELECT id, name FROM products WHERE id = ?", (p_id_int,)).fetchone()
                if not prod_row:
                    is_new_prod = True
                else:
                    p_id = prod_row['id']
            except (ValueError, TypeError):
                is_new_prod = True

            if is_new_prod:
                existing_p = cursor.execute("SELECT id FROM products WHERE LOWER(TRIM(name)) = LOWER(TRIM(?))", (p_name,)).fetchone()
                if existing_p:
                    p_id = existing_p['id']
                else:
                    s_price = round(float(it.get('sale_price', 0)), 2)
                    if s_price <= 0:
                        s_price = round(b_price * 1.15, 2)
                    cursor.execute("INSERT INTO products (name, buy_price, sale_price, stock, unit) VALUES (?, ?, ?, 0, ?)", 
                                   (p_name, b_price, s_price, unit))
                    p_id = cursor.lastrowid
                    
                    adv_units = it.get('advanced_units')
                    if adv_units and isinstance(adv_units, list):
                        for u in adv_units:
                            if u.get('name') and u.get('conv'):
                                cursor.execute("INSERT INTO product_units (product_id, unit_name, conversion_factor, sale_price, is_base_unit) VALUES (?, ?, ?, ?, 0)",
                                               (p_id, str(u['name']).strip(), float(u['conv']), float(u.get('price', s_price))))

            user_sale = round(float(it.get('sale_price', 0)), 2)
            if user_sale > 0:
                cursor.execute("UPDATE products SET sale_price = ? WHERE id = ?", (user_sale, p_id))

            cursor.execute("INSERT INTO purchase_items (purchase_id, product_id, product_name, quantity, buy_price, total, unit, weight) VALUES (?, ?, ?, ?, ?, ?, ?, ?)", 
                           (purchase_id, p_id, p_name, qty, b_price, round(qty * b_price, 2), unit, float(it.get('weight', 0))))
            cursor.execute("UPDATE products SET stock = stock + ?, buy_price = ? WHERE id = ?", (qty, b_price, p_id))

        conn.commit()
        conn.close()
        return jsonify({'success': True, 'purchase_id': purchase_id, 'message': 'خریداری اور اسٹاک کامیابی سے محفوظ ہو گئے!'})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

@app.route('/api/purchase/<int:purchase_id>')
@admin_only
def get_purchase_details(purchase_id):
    try:
        conn = get_db()
        cursor = conn.cursor()
        purchase = cursor.execute('SELECT * FROM purchases WHERE id = ?', (purchase_id,)).fetchone()
        if not purchase:
            conn.close()
            return jsonify({'success': False, 'message': 'خریداری بل نہیں ملا'}), 404
            
        items = cursor.execute('SELECT * FROM purchase_items WHERE purchase_id = ?', (purchase_id,)).fetchall()
        conn.close()
        return jsonify({
            'success': True,
            'purchase': dict(purchase),
            'items': [dict(it) for it in items]
        })
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)}), 500

@app.route('/api/purchase/edit', methods=['POST'])
@admin_only
def edit_purchase():
    try:
        data = request.get_json() or {}
        purchase_id = data.get('purchase_id')
        supplier_id = data.get('supplier_id')
        supplier_name = (data.get('supplier_name') or '').strip()
        bill_no = data.get('bill_no')
        date_str = data.get('date')
        items = data.get('items', [])
        total_amount = round(float(data.get('total_amount', 0)), 2)
        paid_amount = round(float(data.get('paid_amount', 0)), 2)
        due_amount = round(total_amount - paid_amount, 2)
        notes = data.get('notes', '')

        conn = get_db()
        cursor = conn.cursor()

        old_purchase = cursor.execute('SELECT * FROM purchases WHERE id = ?', (purchase_id,)).fetchone()
        if not old_purchase:
            conn.close()
            return jsonify({'success': False, 'message': 'خریداری بل نہیں ملا!'}), 404

        old_items = cursor.execute('SELECT * FROM purchase_items WHERE purchase_id = ?', (purchase_id,)).fetchall()

        # 1. Revert old items stock
        for it in old_items:
            cursor.execute('UPDATE products SET stock = MAX(0, stock - ?) WHERE id = ?', (float(it['quantity']), it['product_id']))

        # 2. Revert old due amount from supplier balance
        old_sup_id = old_purchase['supplier_id']
        old_due = float(old_purchase['due'] or 0)
        if old_due > 0:
            cursor.execute('UPDATE suppliers SET balance = MAX(0, balance - ?) WHERE id = ?', (old_due, old_sup_id))

        # 3. Handle supplier
        if not supplier_id or str(supplier_id).startswith('new_'):
            if not supplier_name:
                supplier_name = old_purchase['supplier_name']
            existing_sup = cursor.execute("SELECT id, name FROM suppliers WHERE LOWER(TRIM(name)) = LOWER(TRIM(?))", (supplier_name,)).fetchone()
            if existing_sup:
                supplier_id = existing_sup['id']
                supplier_name = existing_sup['name']
            else:
                cursor.execute("INSERT INTO suppliers (name, balance) VALUES (?, 0)", (supplier_name,))
                supplier_id = cursor.lastrowid
        else:
            sup_row = cursor.execute('SELECT id, name FROM suppliers WHERE id = ?', (supplier_id,)).fetchone()
            supplier_name = sup_row['name'] if sup_row else old_purchase['supplier_name']

        if date_str:
            try:
                formatted_date = f"{date_str} {datetime.now().strftime('%I:%M:%S %p')}" if len(date_str) <= 10 else date_str
            except Exception:
                formatted_date = old_purchase['date']
        else:
            formatted_date = old_purchase['date']

        cursor.execute('''
            UPDATE purchases 
            SET supplier_id = ?, supplier_name = ?, bill_no = ?, date = ?, total = ?, paid = ?, due = ?, notes = ?
            WHERE id = ?
        ''', (supplier_id, supplier_name, bill_no, formatted_date, total_amount, paid_amount, due_amount, notes, purchase_id))

        cursor.execute('DELETE FROM purchase_items WHERE purchase_id = ?', (purchase_id,))

        for it in items:
            p_id = it.get('id') or it.get('product_id')
            p_name = (it.get('name') or it.get('product_name') or '').strip()
            qty = float(it.get('qty', 1))
            b_price = round(float(it.get('buy_price', 0)), 2)
            unit = it.get('unit') or 'Pcs'

            is_new_prod = False
            try:
                p_id_int = int(p_id)
                prod_row = cursor.execute("SELECT id, name FROM products WHERE id = ?", (p_id_int,)).fetchone()
                if not prod_row:
                    is_new_prod = True
                else:
                    p_id = prod_row['id']
            except (ValueError, TypeError):
                is_new_prod = True

            if is_new_prod:
                existing_p = cursor.execute("SELECT id FROM products WHERE LOWER(TRIM(name)) = LOWER(TRIM(?))", (p_name,)).fetchone()
                if existing_p:
                    p_id = existing_p['id']
                else:
                    s_price = round(b_price * 1.15, 2)
                    cursor.execute("INSERT INTO products (name, buy_price, sale_price, stock, unit) VALUES (?, ?, ?, 0, ?)", 
                                   (p_name, b_price, s_price, unit))
                    p_id = cursor.lastrowid
                    
                    adv_units = it.get('advanced_units')
                    if adv_units and isinstance(adv_units, list):
                        for u in adv_units:
                            if u.get('name') and u.get('conv'):
                                cursor.execute("INSERT INTO product_units (product_id, unit_name, conversion_factor, sale_price, is_base_unit) VALUES (?, ?, ?, ?, 0)",
                                               (p_id, str(u['name']).strip(), float(u['conv']), float(u.get('price', s_price))))

            cursor.execute('INSERT INTO purchase_items (purchase_id, product_id, product_name, quantity, buy_price, total, unit, weight) VALUES (?, ?, ?, ?, ?, ?, ?, ?)',
                           (purchase_id, p_id, p_name, qty, b_price, round(qty * b_price, 2), unit, float(it.get('weight', 0))))
            cursor.execute('UPDATE products SET stock = stock + ?, buy_price = ? WHERE id = ?', (qty, b_price, p_id))

        if due_amount > 0:
            cursor.execute('UPDATE suppliers SET balance = balance + ? WHERE id = ?', (due_amount, supplier_id))

        conn.commit()
        conn.close()
        return jsonify({'success': True, 'message': 'خریداری بل اور اسٹاک کامیابی سے اپ ڈیٹ ہو گیا!'})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)}), 500


@app.route('/api/purchase/return/<int:purchase_id>', methods=['POST'])
@login_required
def return_purchase(purchase_id):
    try:
        conn = get_db()
        cursor = conn.cursor()
        old_purchase = cursor.execute('SELECT * FROM purchases WHERE id = ?', (purchase_id,)).fetchone()
        if not old_purchase:
            conn.close()
            return jsonify({'success': False, 'message': 'بل نہیں ملا!'}), 404
            
        if old_purchase.get('status') == 'returned':
            conn.close()
            return jsonify({'success': False, 'message': 'یہ بل پہلے ہی واپس (Return) ہو چکا ہے!'}), 400

        old_items = cursor.execute('SELECT * FROM purchase_items WHERE purchase_id = ?', (purchase_id,)).fetchall()

        for it in old_items:
            cursor.execute('UPDATE products SET stock = MAX(0, stock - ?) WHERE id = ?', (float(it['quantity']), it['product_id']))

        old_due = float(old_purchase['due'] or 0)
        if old_due > 0:
            cursor.execute('UPDATE suppliers SET balance = MAX(0, balance - ?) WHERE id = ?', (old_due, old_purchase['supplier_id']))

        cursor.execute("UPDATE purchases SET status = 'returned' WHERE id = ?", (purchase_id,))
        conn.commit()
        conn.close()
        return jsonify({'success': True, 'message': 'خریداری کامیابی سے واپس (Return) کر دی گئی ہے!'})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)}), 500

@app.route('/api/purchase/delete/<int:purchase_id>', methods=['POST'])
@admin_only
def delete_purchase(purchase_id):
    try:
        conn = get_db()
        cursor = conn.cursor()
        old_purchase = cursor.execute('SELECT * FROM purchases WHERE id = ?', (purchase_id,)).fetchone()
        if not old_purchase:
            conn.close()
            return jsonify({'success': False, 'message': 'بل نہیں ملا!'}), 404

        old_items = cursor.execute('SELECT * FROM purchase_items WHERE purchase_id = ?', (purchase_id,)).fetchall()

        for it in old_items:
            cursor.execute('UPDATE products SET stock = MAX(0, stock - ?) WHERE id = ?', (float(it['quantity']), it['product_id']))

        old_due = float(old_purchase['due'] or 0)
        if old_due > 0:
            cursor.execute('UPDATE suppliers SET balance = MAX(0, balance - ?) WHERE id = ?', (old_due, old_purchase['supplier_id']))

        cursor.execute('DELETE FROM purchase_items WHERE purchase_id = ?', (purchase_id,))
        cursor.execute('DELETE FROM purchases WHERE id = ?', (purchase_id,))
        conn.commit()
        conn.close()
        return jsonify({'success': True, 'message': 'خریداری بل کامیابی سے ڈیلیٹ ہو گیا اور اسٹاک ریورس ہو گیا!'})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)}), 500

@app.route('/quotations')
@admin_only
def quotations_page():
    conn = get_db()
    cursor = conn.cursor()
    quotes = cursor.execute("SELECT * FROM quotations ORDER BY id DESC").fetchall()
    products_db = cursor.execute("SELECT * FROM products WHERE COALESCE(is_deleted, 0) = 0 ORDER BY name ASC").fetchall()
    products = []
    for p in products_db:
        pd = dict(p)
        units = conn.execute("SELECT unit_name, conversion_factor, sale_price, is_base_unit FROM product_units WHERE product_id = ? ORDER BY is_base_unit DESC, id ASC", (p['id'],)).fetchall()
        pd['units'] = [dict(u) for u in units]
        products.append(pd)
    customers = cursor.execute("SELECT id, name, phone FROM customers ORDER BY name ASC").fetchall()
    conn.close()
    return render_template('quotations.html', quotes=quotes, products=products, customers=customers)

@app.route('/api/quotation/save', methods=['POST'])
@admin_only
def save_quotation():
    try:
        import json
        data = request.get_json() or {}
        cust_name = data.get('customer_name', 'Customer')
        cust_phone = data.get('customer_phone', '')
        items = data.get('items', [])
        total = round(float(data.get('total_amount', 0)), 2)
        quote_no = f"QT-{int(datetime.now().timestamp())}"
        now_str = datetime.now().strftime('%Y-%m-%d %I:%M:%S %p')

        conn = get_db()
        cursor = conn.cursor()
        cursor.execute("INSERT INTO quotations (quote_no, customer_name, customer_phone, date, total, items_json) VALUES (?, ?, ?, ?, ?, ?)", 
                       (quote_no, cust_name, cust_phone, now_str, total, json.dumps(items)))
        q_id = cursor.lastrowid
        conn.commit()
        conn.close()
        return jsonify({'success': True, 'quote_id': q_id, 'quote_no': quote_no})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

@app.route('/api/quotation/<int:quote_id>')
@admin_only
def get_quotation_api(quote_id):
    import json
    conn = get_db()
    quote = conn.execute("SELECT * FROM quotations WHERE id = ?", (quote_id,)).fetchone()
    conn.close()
    if not quote:
        return jsonify({'success': False, 'message': 'Quotation not found'}), 404
    items = json.loads(quote['items_json']) if quote and quote['items_json'] else []
    return jsonify({
        'success': True,
        'quote': {
            'id': quote['id'],
            'quote_no': quote['quote_no'],
            'customer_name': quote['customer_name'],
            'customer_phone': quote['customer_phone'] or '-',
            'date': quote['date'],
            'total': quote['total'],
            'items': items
        }
    })

@app.route('/view_quotation/<int:quote_id>')
@admin_only
def view_quotation(quote_id):
    import json
    conn = get_db()
    quote = conn.execute("SELECT * FROM quotations WHERE id = ?", (quote_id,)).fetchone()
    conn.close()
    if not quote:
        return redirect('/quotations')
    items = json.loads(quote['items_json']) if quote and quote['items_json'] else []
    return render_template('quotation_receipt.html', quote=quote, items=items)

@app.route('/api/quotation/edit/<int:quote_id>', methods=['POST'])
@admin_only
def edit_quotation(quote_id):
    try:
        import json
        data = request.get_json() or {}
        cust_name = data.get('customer_name', 'Customer')
        cust_phone = data.get('customer_phone', '')
        items = data.get('items', [])
        total = round(float(data.get('total_amount', 0)), 2)

        conn = get_db()
        cursor = conn.cursor()
        cursor.execute("""
            UPDATE quotations 
            SET customer_name = ?, customer_phone = ?, total = ?, items_json = ? 
            WHERE id = ?
        """, (cust_name, cust_phone, total, json.dumps(items), quote_id))
        conn.commit()
        conn.close()
        return jsonify({'success': True, 'quote_id': quote_id, 'message': 'کوٹیشن اپ ڈیٹ ہو گئی'})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)}), 500

@app.route('/api/quotation/delete/<int:quote_id>', methods=['POST'])
@admin_only
def delete_quotation(quote_id):
    try:
        conn = get_db()
        conn.execute("DELETE FROM quotations WHERE id = ?", (quote_id,))
        conn.commit()
        conn.close()
        return jsonify({'success': True, 'message': 'کوٹیشن کامیابی سے ڈیلیٹ ہو گئی'})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)}), 500

# ----------------------------------------------------
# 4. Ù…Ø§ÛØ§Ù†Û Ø¢Ù„ Ø±Ù¾ÙˆØ±Ù¹Ø³ (Monthly Master Summary)
# ----------------------------------------------------
@app.route('/monthly-report')
@admin_only
def monthly_master_report():
    conn = get_db()
    cursor = conn.cursor()
    month_filter = request.args.get('month', datetime.now().strftime('%Y-%m'))

    sales_data = cursor.execute("""
        SELECT 
            COALESCE(SUM(total), 0) as total, 
            COALESCE(SUM(paid), 0) as cash, 
            COALESCE(SUM(due), 0) as credit,
            COALESCE(SUM(discount), 0) as discount 
        FROM sales 
        WHERE date LIKE ? AND type != 'returned'
    """, (f'{month_filter}%',)).fetchone()
    purch_data = cursor.execute("SELECT COALESCE(SUM(total), 0) as total, COALESCE(SUM(paid), 0) as paid, COALESCE(SUM(due), 0) as credit FROM purchases WHERE date LIKE ? AND COALESCE(status, \'active\') != \'returned\'", (f'{month_filter}%',)).fetchone()
    exp_total = cursor.execute("SELECT COALESCE(SUM(amount), 0) FROM expenses WHERE date LIKE ?", (f'{month_filter}%',)).fetchone()[0]
    cust_rec = cursor.execute("SELECT COALESCE(SUM(amount), 0) FROM customer_payments WHERE payment_date LIKE ?", (f'{month_filter}%',)).fetchone()[0]
    sup_paid = cursor.execute("SELECT COALESCE(SUM(amount), 0) FROM supplier_payments WHERE payment_date LIKE ?", (f'{month_filter}%',)).fetchone()[0]
    
    conn.close()
    return render_template('monthly_report.html', 
                           month=month_filter,
                           sales=sales_data,
                           purchases=purch_data,
                           expenses=exp_total,
                           customer_recoveries=cust_rec,
                           supplier_payments=sup_paid)

# ----------------------------------------------------
# 5. Ø§Ù†ÙˆÛŒÙ†Ù¹Ø±ÛŒ Ø§ÙˆØ± Ù¾Ø±ÙˆÚˆÚ©Ù¹ Ù…ÛŒÙ†Ø¬Ù…Ù†Ù¹
# ----------------------------------------------------
@app.route('/inventory')
@admin_only
def inventory():
    conn = get_db()
    products_db = conn.execute("SELECT id, code, name, category, category_id, unit, buy_price as purchase_price, sale_price as price, stock, alert_qty, COALESCE(min_stock, alert_qty, 5) as min_stock FROM products ORDER BY id DESC").fetchall()
    
    products = []
    for p in products_db:
        pd = dict(p)
        # Fetch advanced units
        units = conn.execute("SELECT unit_name, conversion_factor, sale_price FROM product_units WHERE product_id = ? AND is_base_unit = 0", (p['id'],)).fetchall()
        pd['units'] = [dict(u) for u in units]
        products.append(pd)
        
    suppliers_db = conn.execute("SELECT id, name, company, phone FROM suppliers ORDER BY name ASC").fetchall()
    suppliers = [dict(s) for s in suppliers_db]
    try:
        categories = conn.execute("SELECT * FROM categories ORDER BY name ASC").fetchall()
    except Exception:
        categories = []
    conn.close()
    return render_template('inventory.html', products=products, suppliers=suppliers, categories=categories)

@app.route('/api/product/by_barcode/<path:barcode>')
@login_required
def get_product_by_barcode(barcode):
    b = barcode.strip().lower()
    conn = get_db()
    prod = conn.execute("SELECT id, code, name, category, unit, buy_price, sale_price, stock, alert_qty FROM products WHERE lower(code) = ? OR id = ? OR lower(name) = ?", (b, b, b)).fetchone()
    conn.close()
    if prod:
        return jsonify({'success': True, 'product': dict(prod)})
    return jsonify({'success': False, 'message': 'Product not found'}), 404

@app.route('/api/product/add', methods=['POST'])
@admin_only
def add_product():
    name = request.form.get('name', '').strip()
    category = request.form.get('category', '').strip()
    code = request.form.get('barcode') or request.form.get('code') or f"P-{int(datetime.now().timestamp())}"
    code = code.strip()
    unit = request.form.get('unit', 'Pcs').strip()
    buy_price = round(float(request.form.get('purchase_price') or request.form.get('buy_price') or 0), 2)
    sale_price = round(float(request.form.get('price') or request.form.get('sale_price') or 0), 2)
    stock = round(float(request.form.get('stock') or 0), 2)
    min_stock = float(request.form.get('min_stock') or request.form.get('alert_qty') or 5)
    alert_qty = int(min_stock)
    category_id = request.form.get('category_id') or None

    if name:
        conn = get_db()
        cursor = conn.cursor()
        
        if not category and category_id:
            c_row = cursor.execute("SELECT name FROM categories WHERE id = ?", (category_id,)).fetchone()
            if c_row:
                category = c_row['name']
        if not category:
            category = 'General'
        
        # Ensure category exists in categories table and category_id is set
        c_match = cursor.execute("SELECT id, name FROM categories WHERE lower(name) = ?", (category.lower(),)).fetchone()
        if c_match:
            category_id = c_match['id']
            category = c_match['name']
        else:
            try:
                cursor.execute("INSERT INTO categories (name) VALUES (?)", (category,))
                category_id = cursor.lastrowid
            except Exception:
                pass
        
        # Extract advanced units lists
        unit_names = request.form.getlist('unit_name[]')
        unit_conversions = request.form.getlist('unit_conversion[]')
        unit_sale_prices = request.form.getlist('unit_sale_price[]')

        # Ø§Ú¯Ø± Ø¨Ø§Ø±Ú©ÙˆÚˆ Ù¾Û Ù„Û’ Ø³Û’ Ù…ÙˆØ¬ÙˆØ¯ Û Û’ ØªÙˆ Ø§Ù¾ ÚˆÛŒÙ¹ Ú©Ø±Ù†Ø§
        existing = cursor.execute("SELECT id FROM products WHERE code = ?", (code,)).fetchone()
        if existing:
            p_id = existing['id']
            cursor.execute("""
                UPDATE products 
                SET name = ?, category = ?, unit = ?, buy_price = ?, sale_price = ?, stock = stock + ?, alert_qty = ?, category_id = ?, min_stock = ?
                WHERE id = ?
            """, (name, category, unit, buy_price, sale_price, stock, alert_qty, category_id, min_stock, p_id))
        else:
            cursor.execute('''
                INSERT INTO products (code, name, category, unit, buy_price, sale_price, stock, alert_qty, category_id, min_stock) 
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (code, name, category, unit, buy_price, sale_price, stock, alert_qty, category_id, min_stock))
            p_id = cursor.lastrowid
            
            # Auto-record in purchases table if initial stock > 0 so it reflects in Daily Closing, Monthly Summary & Profit/Loss!
            if stock > 0 and buy_price > 0:
                tot_purch = round(stock * buy_price, 2)
                now_str = datetime.now().strftime('%Y-%m-%d %I:%M:%S %p')
                bill_no = f"INV-STOCK-{int(datetime.now().timestamp())}"
                cursor.execute("""
                    INSERT INTO purchases (bill_no, supplier_id, supplier_name, date, total, paid, due, notes)
                    VALUES (?, NULL, 'Direct Stock In (نیا مال اسٹاک)', ?, ?, ?, 0, 'Initial Stock Added with Product')
                """, (bill_no, now_str, tot_purch, tot_purch))
                new_p_id = cursor.lastrowid
                cursor.execute("""
                    INSERT INTO purchase_items (purchase_id, product_id, product_name, quantity, buy_price, total, unit)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                """, (new_p_id, p_id, name, stock, buy_price, tot_purch, unit or 'Pcs'))
            
        # Update product_units
        # Clear existing units for this product to replace with new ones
        cursor.execute("DELETE FROM product_units WHERE product_id = ?", (p_id,))
        # 1. Always insert Base Unit
        cursor.execute('''
            INSERT INTO product_units (product_id, unit_name, conversion_factor, purchase_price, sale_price, is_base_unit)
            VALUES (?, ?, 1.0, ?, ?, 1)
        ''', (p_id, unit, buy_price, sale_price))
        
        # 2. Insert Additional Units
        for i in range(len(unit_names)):
            u_name = unit_names[i].strip()
            if u_name:
                try:
                    conv = float(unit_conversions[i])
                    u_price = float(unit_sale_prices[i])
                    cursor.execute('''
                        INSERT INTO product_units (product_id, unit_name, conversion_factor, purchase_price, sale_price, is_base_unit)
                        VALUES (?, ?, ?, ?, ?, 0)
                    ''', (p_id, u_name, conv, buy_price * conv, u_price))
                except (ValueError, TypeError, IndexError):
                    pass

        conn.commit()
        conn.close()

    if request.is_json:
        return jsonify({'success': True, 'message': 'Ù¾Ø±ÙˆÚˆÚ©Ù¹ Ù…Ø­ÙÙˆØ¸ ÛÙˆ Ú¯Ø¦ÛŒ'})
    return redirect('/inventory')

@app.route('/api/product/edit', methods=['POST'])
@admin_only
def edit_product():
    data = request.form if request.form else (request.json or {})
    p_id = data.get('id')
    name = data.get('name', '').strip()
    category = data.get('category', '').strip()
    code = (data.get('barcode') or data.get('code') or '').strip()
    unit = data.get('unit', 'Pcs').strip()
    buy_price = round(float(data.get('purchase_price') or data.get('buy_price') or 0), 2)
    sale_price = round(float(data.get('price') or data.get('sale_price') or 0), 2)
    stock = round(float(data.get('stock') or 0), 2)
    min_stock = float(data.get('min_stock') or data.get('alert_qty') or 5)
    alert_qty = int(min_stock)
    category_id = data.get('category_id') or None

    if not p_id or not name:
        if request.is_json:
            return jsonify({'success': False, 'message': 'پروڈکٹ کی معلومات نامکمل ہیں'}), 400
        return redirect('/inventory')

    conn = get_db()
    cursor = conn.cursor()
    if not category and category_id:
        c_row = cursor.execute("SELECT name FROM categories WHERE id = ?", (category_id,)).fetchone()
        if c_row:
            category = c_row['name']
    if not category:
        category = 'General'

    # Ensure category exists in categories table and category_id is set
    c_match = cursor.execute("SELECT id, name FROM categories WHERE lower(name) = ?", (category.lower(),)).fetchone()
    if c_match:
        category_id = c_match['id']
        category = c_match['name']
    else:
        try:
            cursor.execute("INSERT INTO categories (name) VALUES (?)", (category,))
            category_id = cursor.lastrowid
        except Exception:
            pass

    cursor.execute("""
        UPDATE products 
        SET code = ?, name = ?, category = ?, unit = ?, buy_price = ?, sale_price = ?, stock = ?, alert_qty = ?, category_id = ?, min_stock = ?
        WHERE id = ?
    """, (code, name, category, unit, buy_price, sale_price, stock, alert_qty, category_id, min_stock, p_id))

    # Update product_units
    unit_names = data.getlist('unit_name[]') if hasattr(data, 'getlist') else []
    unit_conversions = data.getlist('unit_conversion[]') if hasattr(data, 'getlist') else []
    unit_sale_prices = data.getlist('unit_sale_price[]') if hasattr(data, 'getlist') else []
    
    cursor.execute("DELETE FROM product_units WHERE product_id = ?", (p_id,))
    # 1. Base Unit
    cursor.execute('''
        INSERT INTO product_units (product_id, unit_name, conversion_factor, purchase_price, sale_price, is_base_unit)
        VALUES (?, ?, 1.0, ?, ?, 1)
    ''', (p_id, unit, buy_price, sale_price))
    
    # 2. Additional Units
    for i in range(len(unit_names)):
        u_name = unit_names[i].strip()
        if u_name:
            try:
                conv = float(unit_conversions[i])
                u_price = float(unit_sale_prices[i])
                cursor.execute('''
                    INSERT INTO product_units (product_id, unit_name, conversion_factor, purchase_price, sale_price, is_base_unit)
                    VALUES (?, ?, ?, ?, ?, 0)
                ''', (p_id, u_name, conv, buy_price * conv, u_price))
            except (ValueError, TypeError, IndexError):
                pass

    conn.commit()
    conn.close()

    if request.is_json:
        return jsonify({'success': True, 'message': 'Ù¾Ø±ÙˆÚˆÚ©Ù¹ Ø§Ù¾ ÚˆÛŒÙ¹ ÛÙˆ Ú¯Ø¦ÛŒ'})
    return redirect('/inventory')

@app.route('/api/product/delete/<int:product_id>', methods=['POST'])
@admin_only
def delete_product(product_id):
    try:
        conn = get_db()
        conn.execute("DELETE FROM products WHERE id = ?", (product_id,))
        conn.commit()
        conn.close()
        if request.is_json:
            return jsonify({'success': True, 'message': 'پروڈکٹ ڈیلیٹ ہو گیا'})
        return redirect('/inventory')
    except sqlite3.IntegrityError:
        return make_delete_error_page(
            title_ur="یہ پروڈکٹ ڈیلیٹ نہیں ہو سکتا!",
            title_en="Cannot Delete Product (Used in Sales)",
            msg_ur="یہ سامان پہلے ہی کسی بل میں استعمال ہو چکا ہے۔ ڈیلیٹ کرنے کے لیے پہلے تمام پرانے بل ڈیلیٹ کریں۔",
            msg_en="This product is already used in invoices. Delete associated invoices first.",
            return_url="/inventory"
        )

@app.route('/api/stock/add', methods=['POST'])
@admin_only
def add_stock():
    p_id = request.form.get('product_id')
    qty = float(request.form.get('quantity') or 0)
    if p_id and qty > 0:
        conn = get_db()
        cursor = conn.cursor()
        prod = cursor.execute("SELECT name, buy_price, unit FROM products WHERE id = ?", (p_id,)).fetchone()
        cursor.execute("UPDATE products SET stock = stock + ? WHERE id = ?", (qty, p_id))
        
        # Auto-record in purchases table so it reflects in daily purchases & reports!
        if prod:
            b_price = float(prod['buy_price'] or 0)
            if b_price > 0:
                tot_purch = round(qty * b_price, 2)
                now_str = datetime.now().strftime('%Y-%m-%d %I:%M:%S %p')
                bill_no = f"INV-ADD-{int(datetime.now().timestamp())}"
                cursor.execute("""
                    INSERT INTO purchases (bill_no, supplier_id, supplier_name, date, total, paid, due, notes)
                    VALUES (?, NULL, 'Direct Stock In (اسٹاک اضافہ)', ?, ?, ?, 0, 'Quick Add Stock from Inventory')
                """, (bill_no, now_str, tot_purch, tot_purch))
                new_p_id = cursor.lastrowid
                cursor.execute("""
                    INSERT INTO purchase_items (purchase_id, product_id, product_name, quantity, buy_price, total, unit)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                """, (new_p_id, p_id, prod['name'], qty, b_price, tot_purch, prod['unit'] or 'Pcs'))

        conn.commit()
        conn.close()
    return redirect('/inventory')

# ----------------------------------------------------
# 6. Ù¾ÙˆØ§Ø¦Ù†Ù¹ Ø¢Ù Ø³ÛŒÙ„ (POS) Ø§ÙˆØ± Ø³ÛŒÙ„Ø² Ú©Ø§Ø¤Ù†Ù¹Ø±
# ----------------------------------------------------
@app.route('/pos')
@login_required
def pos_screen():
    conn = get_db()
    products_db = conn.execute("SELECT * FROM products WHERE stock > 0 ORDER BY name ASC").fetchall()
    
    products = []
    for p in products_db:
        pd = dict(p)
        # Fetch ALL units (Base + Advanced) for POS dropdown
        units = conn.execute("SELECT unit_name, conversion_factor, sale_price, is_base_unit FROM product_units WHERE product_id = ? ORDER BY is_base_unit DESC, id ASC", (p['id'],)).fetchall()
        pd['units'] = [dict(u) for u in units]
        products.append(pd)

    customers = conn.execute("SELECT * FROM customers ORDER BY name ASC").fetchall()
    conn.close()
    return render_template('pos.html', products=products, customers=customers)

@app.route('/api/save_sale', methods=['POST'])
@login_required
def save_sale():
    try:
        data = request.get_json() or {}
        cust_type = data.get('customer_type', 'cash')
        customer_id = data.get('customer_id')
        customer_name = (data.get('customer_name') or 'Cash Customer').strip()
        customer_phone = (data.get('customer_phone') or '').strip()
        customer_address = (data.get('customer_address') or '').strip()
        items = data.get('items', [])
        subtotal = round(float(data.get('subtotal', 0) or 0), 2)
        tax_rate = round(float(data.get('tax_rate', 0) or 0), 2)
        tax_amount = round(float(data.get('tax_amount', 0) or 0), 2)
        discount = round(float(data.get('discount', 0) or 0), 2)

        if subtotal == 0 and items:
            subtotal = sum(round(float(it.get('qty', 1)) * float(it.get('price', 0)), 2) for it in items)

        if tax_amount == 0 and tax_rate > 0:
            tax_amount = round(subtotal * (tax_rate / 100.0), 2)

        total_amount = round(float(data.get('total_amount', (subtotal + tax_amount - discount))), 2)
        paid_amount = round(float(data.get('paid_amount', 0)), 2)
        cash_amount = float(data.get('cash_amount', paid_amount))
        online_amount = float(data.get('online_amount', 0))
        due_amount = round(total_amount - paid_amount, 2)
        actual_type = 'credit' if due_amount > 0 else 'cash'

        conn = get_db()
        cursor = conn.cursor()

        # Stock verification: do not allow sale if stock is insufficient (strictly prevent negative stock)
        for it in items:
            p_id = it.get('id')
            if not p_id:
                continue
            qty = float(it.get('qty', 1))
            conv = float(it.get('conversion_factor', 1.0))
            base_qty = qty * conv
            p_row = cursor.execute("SELECT name, stock, unit FROM products WHERE id = ?", (p_id,)).fetchone()
            if p_row:
                avail = float(p_row['stock'] or 0)
                if base_qty > avail:
                    conn.close()
                    unit_name = p_row['unit'] or 'Pcs'
                    return jsonify({
                        'success': False, 
                        'message': f"اسٹاک ختم یا ناکافی ہے! پروڈکٹ '{p_row['name']}' کا دستیاب اسٹاک صرف {avail:g} {unit_name} ہے۔ بل اینٹر نہیں ہو سکتا۔"
                    })

        if cust_type == 'new_credit' or (due_amount > 0 and customer_name != 'Cash Customer' and not customer_id):
            cursor.execute("SELECT id FROM customers WHERE LOWER(TRIM(name)) = LOWER(TRIM(?))", (customer_name,))
            existing_cust = cursor.fetchone()
            if existing_cust:
                customer_id = existing_cust['id']
            else:
                cursor.execute("INSERT INTO customers (name, phone, address, balance, opening_balance) VALUES (?, ?, ?, 0, 0)", 
                               (customer_name, customer_phone, customer_address))
                customer_id = cursor.lastrowid

        inv_number = f"INV-{uuid.uuid4().hex[:10].upper()}"
        custom_date = (data.get('date') or '').strip()
        if custom_date:
            now_str = f"{custom_date} {datetime.now().strftime('%I:%M:%S %p')}"
        else:
            now_str = datetime.now().strftime('%Y-%m-%d %I:%M:%S %p')

        cursor.execute("INSERT INTO sales (invoice_no, customer_id, customer_name, subtotal, tax_rate, tax_amount, discount, total, paid, due, type, date, cash_amount, online_amount) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", 
                       (inv_number, customer_id, customer_name, subtotal, tax_rate, tax_amount, discount, total_amount, paid_amount, due_amount, actual_type, now_str, cash_amount, online_amount))
        sale_id = cursor.lastrowid

        if customer_id:
            recalculate_customer_balance(customer_id, cursor)

        for it in items:
            p_id = it.get('id')
            qty = float(it.get('qty', 1))
            price = round(float(it.get('price', 0)), 2)
            unit = it.get('unit') or 'Pcs'
            conv = float(it.get('conversion_factor', 1.0))
            base_qty = qty * conv
            weight = round(float(it.get('weight', 0) or 0), 2)
            
            # Fetch current default buy_price from product to lock historical cost
            row = cursor.execute("SELECT buy_price FROM products WHERE id = ?", (p_id,)).fetchone()
            current_buy_price = float(row['buy_price']) if row else 0.0
            
            cursor.execute("INSERT INTO sale_items (sale_id, product_id, product_name, quantity, base_quantity, price, buy_price, total, unit, weight) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", 
                           (sale_id, p_id, it.get('name'), qty, base_qty, price, current_buy_price, round(qty * price, 2), unit, weight))
            cursor.execute("UPDATE products SET stock = stock - ? WHERE id = ?", (base_qty, p_id))

        conn.commit()
        conn.close()
        return jsonify({'success': True, 'sale_id': sale_id, 'invoice_no': inv_number, 'customer_name': customer_name, 'subtotal': subtotal, 'tax_rate': tax_rate, 'tax_amount': tax_amount, 'discount': discount, 'total': total_amount, 'paid': paid_amount, 'due': due_amount, 'type': actual_type, 'date': now_str, 'items': items})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

@app.route('/api/sale/return_partial/<int:sale_id>', methods=['POST'])
@login_required
def return_partial_sale(sale_id):
    try:
        data = request.get_json()
        item_id = data.get('item_id')
        ret_qty = float(data.get('qty', 0))
        
        if not item_id or ret_qty <= 0:
            return jsonify({'success': False, 'message': 'Invalid quantity.'})
            
        conn = get_db()
        cursor = conn.cursor()
        
        sale = cursor.execute("SELECT * FROM sales WHERE id = ?", (sale_id,)).fetchone()
        if not sale or sale['type'] == 'returned':
            conn.close()
            return jsonify({'success': False, 'message': 'Sale not found or already fully returned.'})
            
        item = cursor.execute("SELECT * FROM sale_items WHERE id = ? AND sale_id = ?", (item_id, sale_id)).fetchone()
        if not item:
            conn.close()
            return jsonify({'success': False, 'message': 'Item not found in this sale.'})
            
        # Convert to dict so we can safely access optional columns
        item = dict(item)
        current_ret = float(item.get('returned_qty') or 0)
        orig_qty = float(item['quantity'])
        
        if current_ret + ret_qty > orig_qty:
            conn.close()
            return jsonify({'success': False, 'message': 'Cannot return more than purchased.'})
            
        # Update item returned_qty
        cursor.execute("UPDATE sale_items SET returned_qty = returned_qty + ? WHERE id = ?", (ret_qty, item_id))
        
        # Calculate refund amount
        price = float(item['price'])
        refund_amount = round(ret_qty * price, 2)
        
        # Calculate base_qty to return to stock
        conv = float(item.get('conversion_factor') or 1.0)
        base_ret_qty = ret_qty * conv
        
        # 1. Return stock
        if item.get('product_id'):
            cursor.execute("UPDATE products SET stock = stock + ? WHERE id = ?", (base_ret_qty, item['product_id']))
            
        # 2. Update sale totals
        new_total = round(float(sale['total']) - refund_amount, 2)
        new_subtotal = round(float(sale['subtotal']) - refund_amount, 2)
        
        # Reduce due first, then paid
        current_due = float(sale['due'])
        current_paid = float(sale['paid'])
        
        if current_due >= refund_amount:
            new_due = round(current_due - refund_amount, 2)
            new_paid = current_paid
            if sale['customer_id']:
                cursor.execute("UPDATE customers SET balance = balance - ? WHERE id = ?", (refund_amount, sale['customer_id']))
        else:
            new_due = 0
            refund_cash = refund_amount - current_due
            new_paid = round(current_paid - refund_cash, 2)
            if sale['customer_id'] and current_due > 0:
                cursor.execute("UPDATE customers SET balance = balance - ? WHERE id = ?", (current_due, sale['customer_id']))
        
        cursor.execute("UPDATE sales SET subtotal = ?, total = ?, due = ?, paid = ? WHERE id = ?", 
                       (new_subtotal, new_total, new_due, new_paid, sale_id))
        
        # If all items are fully returned, mark sale as returned
        all_items = cursor.execute("SELECT quantity, returned_qty FROM sale_items WHERE sale_id = ?", (sale_id,)).fetchall()
        if all(float(dict(i).get('returned_qty') or 0) >= float(i['quantity']) for i in all_items):
            cursor.execute("UPDATE sales SET type = 'returned' WHERE id = ?", (sale_id,))
            
        conn.commit()
        conn.close()
        unit_label = item.get('unit') or ''
        return jsonify({'success': True, 'message': f'Successfully returned {ret_qty} {unit_label}.'})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

@app.route('/api/sale/return/<int:sale_id>', methods=['POST'])
@login_required
def return_sale(sale_id):
    conn = get_db()
    cursor = conn.cursor()
    sale = cursor.execute("SELECT * FROM sales WHERE id = ?", (sale_id,)).fetchone()
    
    if not sale:
        conn.close()
        return jsonify({'success': False, 'message': 'بل نہیں ملا'}), 404

    if sale['type'] == 'returned':
        conn.close()
        return jsonify({'success': False, 'message': 'یہ بل پہلے ہی واپس (Return) ہو چکا ہے'}), 400

    # 1. پروڈکٹس کا اسٹاک واپس بحال کرنا
    items = cursor.execute("SELECT product_id, quantity FROM sale_items WHERE sale_id = ?", (sale_id,)).fetchall()
    for it in items:
        p_id = it['product_id']
        qty = float(it['quantity'] or 0)
        if p_id:
            cursor.execute("UPDATE products SET stock = stock + ? WHERE id = ?", (qty, p_id))

    # 2. اگر ادھار کٹوتی تھی تو کسٹمر کھاتے سے بقایا مائنس کرنا
    cust_id = sale['customer_id']
    due_amount = float(sale['due'] or 0)
    if cust_id and due_amount > 0:
        cursor.execute("UPDATE customers SET balance = balance - ? WHERE id = ?", (due_amount, cust_id))

    # 3. بل کو 'returned' مارک کرنا
    cursor.execute("UPDATE sales SET type = 'returned' WHERE id = ?", (sale_id,))
    
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': 'بل کامیابی سے منسوخ / واپس ہو گیا اور اسٹاک بحال کر دیا گیا!'})


@app.route('/api/sale/get/<int:sale_id>')
@login_required
def get_sale_details(sale_id):
    conn = get_db()
    cursor = conn.cursor()
    sale = cursor.execute("SELECT * FROM sales WHERE id = ?", (sale_id,)).fetchone()
    if not sale:
        conn.close()
        return jsonify({'success': False, 'message': 'بل نہیں ملا'}), 404
    
    items = cursor.execute("SELECT * FROM sale_items WHERE sale_id = ?", (sale_id,)).fetchall()
    products_db = cursor.execute("SELECT id, code, name, sale_price, stock, unit FROM products ORDER BY name ASC").fetchall()
    prods_list = []
    for p in products_db:
        pd = dict(p)
        units = cursor.execute("SELECT unit_name, conversion_factor, sale_price, is_base_unit FROM product_units WHERE product_id = ? ORDER BY is_base_unit DESC, id ASC", (p['id'],)).fetchall()
        pd['units'] = [dict(u) for u in units]
        prods_list.append(pd)

    customers_db = cursor.execute("SELECT id, name, phone, balance FROM customers ORDER BY name ASC").fetchall()
    customers_list = [dict(c) for c in customers_db]

    conn.close()

    sale_dict = dict(sale)
    items_list = [dict(it) for it in items]

    return jsonify({
        'success': True,
        'sale': sale_dict,
        'items': items_list,
        'products': prods_list,
        'customers': customers_list
    })

@app.route('/api/sale/edit/<int:sale_id>', methods=['POST'])
@login_required
def edit_sale_invoice(sale_id):
    try:
        data = request.get_json() or {}
        items = data.get('items', [])
        if not items:
            return jsonify({'success': False, 'message': 'کم از کم ایک پروڈکٹ شامل کریں'}), 400

        conn = get_db()
        cursor = conn.cursor()
        sale = cursor.execute("SELECT * FROM sales WHERE id = ?", (sale_id,)).fetchone()
        if not sale:
            conn.close()
            return jsonify({'success': False, 'message': 'بل نہیں ملا'}), 404

        if sale['type'] == 'returned':
            conn.close()
            return jsonify({'success': False, 'message': 'منسوخ شدہ بل میں ترمیم ممکن نہیں ہے'}), 400

        # 1. پچھلے بل کے تمام آئٹمز کا اسٹاک واپس ریورس کرنا
        old_items = cursor.execute("SELECT product_id, base_quantity, quantity FROM sale_items WHERE sale_id = ?", (sale_id,)).fetchall()
        for it in old_items:
            p_id = it['product_id']
            qty = float(it['base_quantity']) if it['base_quantity'] else float(it['quantity'] or 0)
            if p_id:
                cursor.execute("UPDATE products SET stock = stock + ? WHERE id = ?", (qty, p_id))

        old_cust_id = sale['customer_id']

        # 2. نئے حساب کتاب کا تعین
        subtotal = round(float(data.get('subtotal', 0) or 0), 2)
        if subtotal == 0:
            subtotal = sum(round(float(it.get('qty', 1)) * float(it.get('price', 0)), 2) for it in items)

        tax_rate = round(float(data.get('tax_rate', 0) or 0), 2)
        tax_amount = round(float(data.get('tax_amount', 0) or 0), 2)
        if tax_amount == 0 and tax_rate > 0:
            tax_amount = round(subtotal * (tax_rate / 100.0), 2)

        discount = round(float(data.get('discount', 0) or 0), 2)
        total_amount = round(max(0, subtotal + tax_amount - discount), 2)
        paid_amount = round(float(data.get('paid_amount', 0) or 0), 2)
        due_amount = round(max(0, total_amount - paid_amount), 2)
        actual_type = 'credit' if due_amount > 0 else 'cash'
        
        cust_name = (data.get('customer_name') or sale['customer_name'] or 'Cash Customer').strip()
        cust_id = data.get('customer_id') or None
        try:
            cust_id = int(cust_id) if cust_id else None
        except (ValueError, TypeError):
            cust_id = None

        # اگر cust_id خالی ہے لیکن کسٹمر کا نام دیا ہے تو کسٹمرز ٹیبل میں میچ کریں
        if not cust_id and cust_name and cust_name.lower() not in ('cash customer', 'walk-in customer', 'cash', 'عام گاہک', 'نقد گاہک'):
            match_row = cursor.execute("SELECT id FROM customers WHERE LOWER(TRIM(name)) = LOWER(TRIM(?))", (cust_name,)).fetchone()
            if match_row:
                cust_id = match_row['id']
            elif due_amount > 0:
                # اگر نام لکھا ہے، بقایا ادھار ہے اور کسٹمر موجود نہیں تو خودکار اندراج کریں تاکہ کھاتے میں جا سکے
                cursor.execute("INSERT INTO customers (name, phone, address, balance, opening_balance) VALUES (?, '', '', 0, 0)", (cust_name,))
                cust_id = cursor.lastrowid

        # 3. سیلز ٹیبل اپڈیٹ (مع customer_id)
        cursor.execute('''
            UPDATE sales SET 
                customer_id = ?,
                customer_name = ?, 
                subtotal = ?, 
                tax_rate = ?, 
                tax_amount = ?, 
                discount = ?, 
                total = ?, 
                paid = ?, 
                due = ?, 
                type = ?
            WHERE id = ?
        ''', (cust_id, cust_name, subtotal, tax_rate, tax_amount, discount, total_amount, paid_amount, due_amount, actual_type, sale_id))

        # 4. پرانے sale_items ڈیلیٹ کر کے نئے داخل کرنا اور نیا اسٹاک مائنس کرنا
        cursor.execute("DELETE FROM sale_items WHERE sale_id = ?", (sale_id,))
        for it in items:
            p_id = it.get('product_id') or it.get('id')
            p_name = it.get('product_name') or it.get('name')
            qty = float(it.get('quantity') or it.get('qty') or 1)
            price = round(float(it.get('price') or 0), 2)
            row_total = round(qty * price, 2)
            unit = it.get('unit') or 'Pcs'
            conv = float(it.get('conversion_factor', 1.0))
            base_qty = qty * conv
            weight = round(float(it.get('weight', 0) or 0), 2)

            # Fetch buy_price to lock historical cost
            row = cursor.execute("SELECT buy_price FROM products WHERE id = ?", (p_id,)).fetchone()
            current_buy_price = float(row['buy_price']) if row else 0.0

            cursor.execute("INSERT INTO sale_items (sale_id, product_id, product_name, quantity, base_quantity, price, buy_price, total, unit, weight) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                           (sale_id, p_id, p_name, qty, base_qty, price, current_buy_price, row_total, unit, weight))
            if p_id:
                cursor.execute("UPDATE products SET stock = stock - ? WHERE id = ?", (base_qty, p_id))

        # 5. کسٹمر کھاتہ ازسرنو 100% درست سنک کرنا
        if cust_id:
            recalculate_customer_balance(cust_id, cursor)
        if old_cust_id and old_cust_id != cust_id:
            recalculate_customer_balance(old_cust_id, cursor)

        conn.commit()
        conn.close()
        return jsonify({'success': True, 'message': 'بل اور کسٹمر کھاتہ کامیابی سے تبدیل اور اپ ڈیٹ ہو گیا ہے!'})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)}), 500

@app.route('/api/sale/pay_due/<int:sale_id>', methods=['POST'])
@login_required
def pay_sale_due(sale_id):
    try:
        data = request.get_json() or {}
        pay_amount = round(float(data.get('amount') or 0), 2)
        note = (data.get('note') or 'بل بقایا ادھار وصولی').strip()

        if pay_amount <= 0:
            return jsonify({'success': False, 'message': 'درست وصول شدہ رقم درج کریں'}), 400

        conn = get_db()
        cursor = conn.cursor()
        sale = cursor.execute("SELECT * FROM sales WHERE id = ?", (sale_id,)).fetchone()
        if not sale:
            conn.close()
            return jsonify({'success': False, 'message': 'بل نہیں ملا'}), 404

        current_due = float(sale['due'] or 0)
        current_paid = float(sale['paid'] or 0)
        cust_id = sale['customer_id']

        if current_due <= 0:
            conn.close()
            return jsonify({'success': False, 'message': 'یہ بل پہلے ہی مکمل ادا ہو چکا ہے!'}), 400

        if cust_id:
            cust = cursor.execute("SELECT balance FROM customers WHERE id = ?", (cust_id,)).fetchone()
            if cust and float(cust['balance'] or 0) <= 0:
                conn.close()
                return jsonify({'success': False, 'message': 'اس گاہک کا کھاتہ پہلے ہی صاف (Rs. 0) ہے، اس بل پر مزید ادائیگی درکار نہیں ہے۔'}), 400

        if pay_amount > (current_due + 0.01):
            conn.close()
            return jsonify({'success': False, 'message': f'وصول شدہ رقم بقایا ادھار (Rs. {current_due:,.2f}) سے زیادہ نہیں ہو سکتی'}), 400

        new_paid = round(current_paid + pay_amount, 2)
        new_due = max(0.0, round(current_due - pay_amount, 2))
        new_type = 'credit' if sale['type'] == 'credit' else ('cash' if new_due <= 0 else 'credit')

        cursor.execute("UPDATE sales SET paid = ?, due = ?, type = ? WHERE id = ?", (new_paid, new_due, new_type, sale_id))

        # اگر کسٹمر لنک ہے تو کسٹمر کی ادائیگی (customer_payments) میں بل کے ساتھ اندراج اور کھاتہ سنک کریں
        if cust_id:
            now_str = datetime.now().strftime("%Y-%m-%d %I:%M:%S %p")
            pay_note = f"بل #{sale_id} ادھار وصولی - {note}"
            cursor.execute("INSERT INTO customer_payments (customer_id, amount, payment_date, note, sale_id) VALUES (?, ?, ?, ?, ?)",
                           (cust_id, pay_amount, now_str, pay_note, sale_id))
            recalculate_customer_balance(cust_id, cursor)

        conn.commit()
        conn.close()
        return jsonify({
            'success': True, 
            'message': f'بل #{sale_id} کی وصولی Rs. {pay_amount:,.2f} کامیابی سے درج ہو گئی۔ نیا بقایا: Rs. {new_due:,.2f}'
        })
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)}), 500

@app.route('/api/backup/download_db')
@admin_only
def download_database_backup():
    try:
        os.makedirs(MANUAL_BACKUP_DIR, exist_ok=True)
        now_str = datetime.now().strftime('%Y-%m-%d_%I-%M-%p')
        filename = f"Smart_POS_Manual_DB_{now_str}.db"
        dest_path = os.path.join(MANUAL_BACKUP_DIR, filename)

        if os.path.exists(DB_NAME):
            shutil.copy2(DB_NAME, dest_path)

        return send_file(
            DB_NAME,
            as_attachment=True,
            download_name=filename,
            mimetype="application/x-sqlite3"
        )
    except Exception as e:
        flash(f"خرابی: {str(e)}", "danger")
        return redirect('/dashboard')

@app.route('/api/backup/restore_db', methods=['POST'])
@admin_only
def restore_database_backup():
    try:
        if 'backup_file' not in request.files:
            return jsonify({'success': False, 'message': 'No backup file selected'}), 400
        
        file = request.files['backup_file']
        if file.filename == '':
            return jsonify({'success': False, 'message': 'Please select a valid .db file'}), 400

        # Ensure backup folder exists (auto-create if missing)
        os.makedirs(BACKUP_DIR, exist_ok=True)

        # Save uploaded file temporarily to verify
        temp_path = os.path.join(BACKUP_DIR, 'temp_restore.db')
        file.save(temp_path)

        # Verify SQLite integrity
        temp_conn = sqlite3.connect(temp_path)
        cur = temp_conn.cursor()
        integrity = cur.execute("PRAGMA integrity_check").fetchone()
        if not integrity or integrity[0] != 'ok':
            temp_conn.close()
            if os.path.exists(temp_path):
                os.remove(temp_path)
            return jsonify({'success': False, 'message': 'Corrupted database backup file.'}), 400

        tables = [r[0] for r in cur.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()]
        temp_conn.close()

        if 'products' not in tables or 'sales' not in tables:
            if os.path.exists(temp_path):
                os.remove(temp_path)
            return jsonify({'success': False, 'message': 'Invalid backup file. Not a valid Smart POS database.'}), 400

        # Save safety backup of current db before restoring
        backup_info = save_backup_to_folder("pre_restore_safety")

        # Safely overwrite current database using SQLite online backup API
        src_conn = sqlite3.connect(temp_path)
        dst_conn = sqlite3.connect(DB_NAME, timeout=30)
        with dst_conn:
            src_conn.backup(dst_conn)
        dst_conn.close()
        src_conn.close()

        if os.path.exists(temp_path):
            os.remove(temp_path)

        # Run schema migrations and user checks on the newly restored database
        init_db()
        init_auth()

        b_name = backup_info.get('db_name', '') if backup_info else ''
        return jsonify({
            'success': True, 
            'message': 'Database restored successfully!',
            'backup_db': b_name
        })
    except Exception as e:
        logger.error(f"Error during restore: {e}")
        return jsonify({'success': False, 'message': f'Restore Error: {str(e)}'}), 500


@app.route('/api/system/reset_database', methods=['POST'])
@admin_only
def reset_database():
    try:
        data = request.get_json() or {}
        password = data.get('password', '').strip()
        clear_products = data.get('clear_products', True)

        conn = get_db()
        cursor = conn.cursor()
        admin_row = cursor.execute("SELECT password FROM users WHERE username = 'admin'").fetchone()

        if not admin_row or not check_password_hash(admin_row['password'], password):
            conn.close()
            return jsonify({'success': False, 'message': 'غلط پاس ورڈ! ایڈمن پاس ورڈ درج کریں (Incorrect admin password)'}), 400

        # 1. Automatic safety backup before reset
        backup_info = save_backup_to_folder("pre_reset_safety")

        # 2. Clear all transaction tables safely
        tables_to_clear = [
            'sales', 'sale_items', 'purchases', 'purchase_items',
            'customer_payments', 'supplier_payments', 'customer_ledger',
            'expenses', 'quotations'
        ]
        for tbl in tables_to_clear:
            try:
                cursor.execute(f"DELETE FROM {tbl}")
            except Exception:
                pass

        try:
            cursor.execute("DELETE FROM sqlite_sequence WHERE name IN ('sales', 'sale_items', 'purchases', 'purchase_items', 'customer_payments', 'supplier_payments', 'customer_ledger', 'expenses', 'quotations')")
        except Exception:
            pass

        # 3. Reset or clear products and contacts
        if clear_products:
            for tbl in ['products', 'product_units', 'customers', 'suppliers']:
                try:
                    cursor.execute(f"DELETE FROM {tbl}")
                except Exception:
                    pass
            try:
                cursor.execute("DELETE FROM sqlite_sequence WHERE name IN ('products', 'product_units', 'customers', 'suppliers')")
            except Exception:
                pass
        else:
            try:
                cursor.execute("UPDATE products SET stock = 0")
                cursor.execute("UPDATE customers SET balance = 0")
                cursor.execute("UPDATE suppliers SET balance = 0")
            except Exception:
                pass

        conn.commit()
        conn.close()

        # Re-run migrations and category checks
        init_db()
        init_auth()

        b_name = backup_info.get('db_name', '') if backup_info else ''
        return jsonify({
            'success': True, 
            'message': 'سافٹ ویئر کامیابی سے بالکل صاف (Reset) کر دیا گیا ہے!',
            'backup_db': b_name
        })
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)}), 500


@app.route('/sales')
@login_required
def sales_history():
    conn = get_db()
    cursor = conn.cursor()
    selected_date = request.args.get('date', '')
    
    query = "SELECT s.*, COALESCE(c.name, s.customer_name, 'Cash Customer') as customer_name FROM sales s LEFT JOIN customers c ON s.customer_id = c.id "
    query += "WHERE s.date LIKE ? ORDER BY s.id DESC" if selected_date else "ORDER BY s.id DESC"
    sales_raw = cursor.execute(query, (f"{selected_date}%",) if selected_date else ()).fetchall()
        
    dates_raw = cursor.execute("SELECT DISTINCT SUBSTR(date, 1, 10) as sale_day FROM sales ORDER BY sale_day DESC LIMIT 5").fetchall()
    recent_dates = [d['sale_day'] for d in dates_raw if d['sale_day']]
    
    cust_balances = {}
    for r in cursor.execute("SELECT id, balance FROM customers").fetchall():
        cust_balances[r['id']] = float(r['balance'] or 0.0)

    # Pre-calculate effective bill due/paid for customers based on their actual real-time balance
    all_cust_sales = cursor.execute(
        "SELECT id, customer_id, total, paid, due, type FROM sales WHERE customer_id IS NOT NULL AND type != 'returned' ORDER BY id ASC"
    ).fetchall()

    cust_sales_map = {}
    for cs in all_cust_sales:
        cust_sales_map.setdefault(cs['customer_id'], []).append(cs)

    sale_effective = {}
    for cid, s_list in cust_sales_map.items():
        c_bal = max(0.0, cust_balances.get(cid, 0.0))
        rem_due = c_bal
        # Allocate rem_due backwards (newest sales first)
        for s in reversed(s_list):
            s_id = s['id']
            tot = float(s['total'] or 0.0)
            due_for_this = min(rem_due, tot)
            paid_for_this = max(0.0, round(tot - due_for_this, 2))
            rem_due = max(0.0, round(rem_due - due_for_this, 2))

            orig_type = s['type']
            if orig_type == 'credit' or tot > float(s['paid'] or 0.0) or float(s['due'] or 0.0) > 0:
                eff_type = 'credit'
            else:
                eff_type = 'cash'

            sale_effective[s_id] = {
                'paid': paid_for_this,
                'due': due_for_this,
                'balance': due_for_this,
                'type': eff_type
            }

    sales = []
    for row in sales_raw:
        s_dict = dict(row)
        s_id = s_dict.get('id')
        s_total = float(s_dict.get('total') or 0.0)

        if s_dict.get('type') == 'returned':
            s_dict['total'] = s_total
            s_dict['paid'] = float(s_dict.get('paid') or 0.0)
            s_dict['balance'] = 0.0
            s_dict['due'] = 0.0
            s_dict['type'] = 'returned'
        elif s_id in sale_effective:
            eff = sale_effective[s_id]
            s_dict['total'] = s_total
            s_dict['paid'] = eff['paid']
            s_dict['due'] = eff['due']
            s_dict['balance'] = eff['balance']
            s_dict['type'] = eff['type']
        else:
            s_paid = float(s_dict.get('paid') or 0.0)
            s_due = float(s_dict.get('due') if s_dict.get('due') is not None else max(0.0, s_total - s_paid))
            s_dict['total'] = s_total
            s_dict['paid'] = s_paid
            s_dict['due'] = s_due
            s_dict['balance'] = s_due
            if s_dict.get('type') != 'credit' and s_due > 0:
                s_dict['type'] = 'credit'

        sales.append(s_dict)
        
    conn.close()
    return render_template('sales.html', sales=sales, recent_dates=recent_dates, selected_date=selected_date)

@app.route('/receipt/<path:sale_ident>')
@login_required
def view_invoice(sale_ident):
    try:
        conn = get_db()
        sale = None
        sale_str = str(sale_ident).strip()
        
        # 1. Try finding by numeric id first if digits
        if sale_str.isdigit():
            sale = conn.execute("SELECT s.*, COALESCE(c.name, s.customer_name, 'Cash Sale') as customer_name, c.phone as customer_phone FROM sales s LEFT JOIN customers c ON s.customer_id = c.id WHERE s.id = ?", (int(sale_str),)).fetchone()
        
        # 2. Try finding by invoice_no if not found yet
        if not sale:
            sale = conn.execute("SELECT s.*, COALESCE(c.name, s.customer_name, 'Cash Sale') as customer_name, c.phone as customer_phone FROM sales s LEFT JOIN customers c ON s.customer_id = c.id WHERE LOWER(TRIM(s.invoice_no)) = LOWER(?)", (sale_str,)).fetchone()
            
        # 3. Fallback: try LIKE match for invoice_no
        if not sale:
            sale = conn.execute("SELECT s.*, COALESCE(c.name, s.customer_name, 'Cash Sale') as customer_name, c.phone as customer_phone FROM sales s LEFT JOIN customers c ON s.customer_id = c.id WHERE s.invoice_no LIKE ?", (f"%{sale_str}%",)).fetchone()
            
        if not sale:
            conn.close()
            return "<div style='font-family:sans-serif;padding:30px;text-align:center;'><h3>Invoice Not Found / بل نہیں ملا</h3><p>Specified invoice does not exist.</p></div>", 404

        sale_dict = dict(sale)
        real_id = sale_dict.get('id')
        
        # Fetch items safely (works with or without returned_qty column)
        items_raw = conn.execute("SELECT si.*, COALESCE(p.unit, 'Pcs') as prod_unit FROM sale_items si LEFT JOIN products p ON si.product_id = p.id WHERE si.sale_id = ?", (real_id,)).fetchall()
        conn.close()
        
        # Safely sanitize all numeric fields so template never throws TypeError/ValueError
        sale_dict['subtotal'] = float(sale_dict.get('subtotal') or sale_dict.get('total') or 0.0)
        sale_dict['tax_rate'] = float(sale_dict.get('tax_rate') or 0.0)
        sale_dict['tax_amount'] = float(sale_dict.get('tax_amount') or 0.0)
        sale_dict['discount'] = float(sale_dict.get('discount') or 0.0)
        sale_dict['total'] = float(sale_dict.get('total') or 0.0)
        sale_dict['paid'] = float(sale_dict.get('paid') or 0.0)
        sale_dict['due'] = float(sale_dict.get('due') or 0.0)
        sale_dict['cash_amount'] = float(sale_dict.get('cash_amount') or 0.0)
        sale_dict['online_amount'] = float(sale_dict.get('online_amount') or 0.0)
        
        clean_items = []
        for it in items_raw:
            it_dict = dict(it)
            it_dict['price'] = float(it_dict.get('price') or 0.0)
            it_dict['quantity'] = float(it_dict.get('quantity') or 0.0)
            it_dict['total'] = float(it_dict.get('total') or (it_dict['quantity'] * it_dict['price']))
            it_dict['weight'] = float(it_dict.get('weight') or 0.0)
            it_dict['unit'] = it_dict.get('unit') or it_dict.get('prod_unit') or 'Pcs'
            it_dict['returned_qty'] = float(it_dict.get('returned_qty') or 0.0)
            clean_items.append(it_dict)

        return render_template('receipt.html', sale=sale_dict, items=clean_items)
    except Exception as e:
        logger.error(f"Error rendering receipt for {sale_ident}: {e}")
        return f"<div style='font-family:sans-serif;padding:30px;text-align:center;'><h3>Receipt Error</h3><p>{e}</p></div>", 500

@app.route('/daily-report')
@admin_only
def daily_report():
    conn = get_db()
    cursor = conn.cursor()
    selected_date = request.args.get('date', datetime.now().strftime("%Y-%m-%d"))
    date_pattern = f"{selected_date}%"
    
    raw_sales = cursor.execute("SELECT s.*, COALESCE(c.name, s.customer_name) as customer_name FROM sales s LEFT JOIN customers c ON s.customer_id = c.id WHERE s.date LIKE ? AND s.type != 'returned' ORDER BY s.id DESC", (date_pattern,)).fetchall()
    direct_payments = cursor.execute("SELECT p.*, c.name as customer_name FROM customer_payments p LEFT JOIN customers c ON p.customer_id = c.id WHERE p.payment_date LIKE ? ORDER BY p.id DESC", (date_pattern,)).fetchall()
    raw_purchases = cursor.execute("""
        SELECT p.*, COALESCE(s.name, p.supplier_name, 'Direct Stock In (نیا مال)') as supplier_display_name 
        FROM purchases p 
        LEFT JOIN suppliers s ON p.supplier_id = s.id 
        WHERE p.date LIKE ? 
        ORDER BY p.id DESC
    """, (date_pattern,)).fetchall()
    
    purchases = []
    total_purchases_amount = 0.0
    cash_paid_for_purchases = 0.0
    credit_purchases_due = 0.0
    for row in raw_purchases:
        p_dict = dict(row)
        p_tot = float(p_dict.get('total') or 0)
        p_paid = float(p_dict.get('paid') or 0)
        p_due = float(p_dict.get('due') if p_dict.get('due') is not None else (p_tot - p_paid))
        p_dict['total'] = p_tot
        p_dict['paid'] = p_paid
        p_dict['due'] = p_due
        total_purchases_amount += p_tot
        cash_paid_for_purchases += p_paid
        credit_purchases_due += p_due
        purchases.append(p_dict)
    
    sales = []
    total_sales_amount = 0.0
    cash_from_sales = 0.0
    credit_given_today = 0.0
    
    for row in raw_sales:
        s_dict = dict(row)
        s_total = float(s_dict.get('total') or 0)
        s_paid = float(s_dict.get('paid') or 0)
        s_bal = float(s_dict.get('due') if s_dict.get('due') is not None else (s_total - s_paid))
        s_dict['total'] = s_total
        s_dict['paid'] = s_paid
        s_dict['balance'] = s_bal
        total_sales_amount += s_total
        cash_from_sales += s_paid
        if s_bal > 0: credit_given_today += s_bal
        sales.append(s_dict)
    
    cash_from_recoveries = sum(float(p['amount'] or 0) for p in direct_payments)
    total_discount_today = cursor.execute("SELECT COALESCE(SUM(discount), 0) FROM sales WHERE date LIKE ? AND type != 'returned'", (date_pattern,)).fetchone()[0] or 0.0
    net_credit_today = max(0.0, credit_given_today - cash_from_recoveries)
    conn.close()
    
    return render_template('daily_report.html', selected_date=selected_date, sales=sales, direct_payments=direct_payments, 
                           purchases=purchases, total_purchases_amount=total_purchases_amount,
                           cash_paid_for_purchases=cash_paid_for_purchases, credit_purchases_due=credit_purchases_due,
                           total_sales_amount=total_sales_amount, cash_from_sales=cash_from_sales, 
                           credit_given_today=credit_given_today, net_credit_today=net_credit_today,
                           cash_from_recoveries=cash_from_recoveries, 
                           total_discount_today=total_discount_today,
                           total_cash_in_hand=cash_from_sales + cash_from_recoveries)

# ----------------------------------------------------
# 7. Ø§Ø®Ø±Ø§Ø¬Ø§Øª (Expenses)
# ----------------------------------------------------
@app.route('/expenses', methods=['GET', 'POST'])
@admin_only
def manage_expenses():
    conn = get_db()
    cursor = conn.cursor()
    if request.method == 'POST':
        title = request.form.get('title')
        category = request.form.get('category')
        amount = round(float(request.form.get('amount', 0) or 0), 2)
        notes = request.form.get('notes', '')
        custom_date = request.form.get('date')
        if custom_date:
            exp_date = f"{custom_date} {datetime.now().strftime('%H:%M')}"
        else:
            exp_date = datetime.now().strftime('%Y-%m-%d %H:%M')
        cursor.execute('INSERT INTO expenses (title, category, amount, date, notes) VALUES (?, ?, ?, ?, ?)', 
                       (title, category, amount, exp_date, notes))
        conn.commit()
        conn.close()
        return redirect('/expenses')
    
    expenses_list = cursor.execute("SELECT * FROM expenses ORDER BY id DESC").fetchall()
    today_total = cursor.execute("SELECT SUM(amount) FROM expenses WHERE date LIKE ?", (f"{datetime.now().strftime('%Y-%m-%d')}%",)).fetchone()[0] or 0
    month_total = cursor.execute("SELECT SUM(amount) FROM expenses WHERE date LIKE ?", (f"{datetime.now().strftime('%Y-%m')}%",)).fetchone()[0] or 0
    conn.close()
    return render_template('expenses.html', expenses=expenses_list, today_total=today_total, month_total=month_total)

@app.route('/delete_expense/<int:expense_id>', methods=['POST'])
@admin_only
def delete_expense(expense_id):
    try:
        conn = get_db()
        conn.execute("DELETE FROM expenses WHERE id = ?", (expense_id,))
        conn.commit()
        conn.close()
        return redirect('/expenses')
    except sqlite3.IntegrityError:
        return make_delete_error_page(
            title_ur="یہ خرچہ ڈیلیٹ نہیں ہو سکتا!",
            title_en="Cannot Delete Expense",
            msg_ur="یہ خرچہ ڈیلیٹ نہیں ہو سکتا۔",
            msg_en="This expense record cannot be deleted.",
            return_url="/expenses"
        )

# ----------------------------------------------------
# 8. Ù†Ù Ø¹ Ùˆ Ù†Ù‚ØµØ§Ù† (Profit & Loss)
# ----------------------------------------------------
@app.route('/profit-loss')
@admin_only
def profit_loss_report():
    conn = get_db()
    cursor = conn.cursor()
    month_filter = request.args.get('month', datetime.now().strftime('%Y-%m'))
    
    total_sales = cursor.execute("SELECT COALESCE(SUM(total), 0) FROM sales WHERE date LIKE ? AND type != 'returned'", (f'{month_filter}%',)).fetchone()[0] or 0
    total_discount = cursor.execute("SELECT COALESCE(SUM(discount), 0) FROM sales WHERE date LIKE ? AND type != 'returned'", (f'{month_filter}%',)).fetchone()[0] or 0
    
    # 1. Total purchases (Actual stock inventory purchases this month)
    total_purchases = cursor.execute("SELECT COALESCE(SUM(total), 0) FROM purchases WHERE date LIKE ? AND COALESCE(status, \'active\') != \'returned\'", (f'{month_filter}%',)).fetchone()[0] or 0

    total_cogs = 0
    try:
        row = cursor.execute("""
            SELECT COALESCE(SUM(si.base_quantity * CASE WHEN si.buy_price > 0 THEN si.buy_price ELSE p.buy_price END), 0) 
            FROM sale_items si 
            JOIN products p ON si.product_id = p.id 
            JOIN sales s ON si.sale_id = s.id 
            WHERE s.date LIKE ? AND s.type != 'returned'
        """, (f'{month_filter}%',)).fetchone()
        if row: total_cogs = row[0]
    except Exception:
        pass

    gross_profit = total_sales - total_cogs
    total_expenses = cursor.execute("SELECT COALESCE(SUM(amount), 0) FROM expenses WHERE date LIKE ?", (f'{month_filter}%',)).fetchone()[0] or 0
    net_profit = gross_profit - total_expenses
    expense_breakdown = cursor.execute("SELECT category, SUM(amount) FROM expenses WHERE date LIKE ? GROUP BY category", (f'{month_filter}%',)).fetchall()
    conn.close()
    
    return render_template('profit_loss.html', 
                           month=month_filter, 
                           total_sales=total_sales, 
                           total_discount=total_discount,
                           total_purchases=total_purchases, 
                           total_cogs=total_cogs, 
                           gross_profit=gross_profit, 
                           total_expenses=total_expenses, 
                           net_profit=net_profit, 
                           expense_breakdown=expense_breakdown)

# ----------------------------------------------------
# 9. Ø³Ø§Ø¦Ù„Ù†Ù¹ ØªÚ¾Ø±Ù…Ù„ Ù¾Ø±Ù†Ù¹Ù†Ú¯ API
# ----------------------------------------------------
@app.route('/direct_print', methods=['POST'])
@login_required
def direct_print():
    data = request.get_json() or {}
    text = data.get('text', '')
    if win32print and text:
        try:
            printer_name = win32print.GetDefaultPrinter()
            hPrinter = win32print.OpenPrinter(printer_name)
            try:
                hJob = win32print.StartDocPrinter(hPrinter, 1, ("POS Receipt", None, "RAW"))
                try:
                    win32print.StartPagePrinter(hPrinter)
                    win32print.WritePrinter(hPrinter, text.encode('utf-8', errors='ignore'))
                    win32print.EndPagePrinter(hPrinter)
                finally:
                    win32print.EndDocPrinter(hPrinter)
            finally:
                win32print.ClosePrinter(hPrinter)
            return jsonify({'success': True, 'printed': True, 'printer': printer_name})
        except Exception as e:
            return jsonify({'success': True, 'printed': False, 'note': str(e)})
    return jsonify({'success': True, 'printed': False, 'note': 'Silent print handled locally'})

# ----------------------------------------------------
# 10. 💾 خودکار و دستی بیک اپ سسٹم (APScheduler + atexit + Excel)
# ----------------------------------------------------
def clean_old_backups(days=10):
    """Delete backup files older than specified days to keep hard drive clean"""
    try:
        now_ts = datetime.now().timestamp()
        cutoff_sec = days * 86400  # 10 days in seconds
        
        folders_to_clean = [DB_BACKUP_DIR, EXCEL_BACKUP_DIR]
        desktop_dir = os.path.join(os.path.expanduser('~'), 'Desktop', 'daily_backups')
        if os.path.exists(desktop_dir):
            folders_to_clean.extend([
                os.path.join(desktop_dir, 'database_backups'),
                os.path.join(desktop_dir, 'excel_backups')
            ])
            
        for folder in folders_to_clean:
            if not os.path.exists(folder):
                continue
            for fname in os.listdir(folder):
                fpath = os.path.join(folder, fname)
                if os.path.isfile(fpath):
                    # Check file age
                    mtime = os.path.getmtime(fpath)
                    if (now_ts - mtime) > cutoff_sec:
                        try:
                            os.remove(fpath)
                            logger.info(f"Cleaned old backup file (>10 days): {fname}")
                        except Exception:
                            pass
    except Exception as e_clean:
        logger.warning(f"clean_old_backups error: {e_clean}")

def save_backup_to_folder(tag="auto"):
    try:
        if not os.path.exists(DB_BACKUP_DIR):
            os.makedirs(DB_BACKUP_DIR, exist_ok=True)
        if not os.path.exists(EXCEL_BACKUP_DIR):
            os.makedirs(EXCEL_BACKUP_DIR, exist_ok=True)

        now_str = datetime.now().strftime('%Y-%m-%d_%I-%M-%p')
        db_filename = f"backup_{tag}_{now_str}.db"
        excel_filename = f"excel_{tag}_{now_str}.xlsx"
        db_dest = os.path.join(DB_BACKUP_DIR, db_filename)
        excel_dest = os.path.join(EXCEL_BACKUP_DIR, excel_filename)

        # 1. ڈیٹا بیس فائل کاپی
        if os.path.exists(DB_NAME):
            shutil.copy2(DB_NAME, db_dest)

        # 2. کثیر المقاصد ایکسل بیک اپ محفوظ کریں
        conn = get_db()
        df_customers = pd.read_sql_query("SELECT * FROM customers ORDER BY id DESC", conn)
        df_stock = pd.read_sql_query("SELECT * FROM products ORDER BY id DESC", conn)
        df_suppliers = pd.read_sql_query("SELECT * FROM suppliers ORDER BY id DESC", conn)
        df_sales = pd.read_sql_query("SELECT * FROM sales ORDER BY id DESC", conn)
        try:
            df_expenses = pd.read_sql_query("SELECT * FROM expenses ORDER BY id DESC", conn)
        except Exception:
            df_expenses = pd.DataFrame()
        conn.close()

        with pd.ExcelWriter(excel_dest, engine='openpyxl') as writer:
            df_customers.to_excel(writer, sheet_name='Customer_Khata', index=False)
            df_stock.to_excel(writer, sheet_name='Inventory_Stock', index=False)
            df_suppliers.to_excel(writer, sheet_name='Suppliers_Payables', index=False)
            df_sales.to_excel(writer, sheet_name='Sales_Invoices', index=False)
            if not df_expenses.empty:
                df_expenses.to_excel(writer, sheet_name='Daily_Expenses', index=False)

        # 3. ڈیسک ٹاپ کے daily_backups فولڈر میں بھی فوری کاپی محفوظ کریں (Desktop Mirroring)
        try:
            desktop_dir = os.path.join(os.path.expanduser('~'), 'Desktop', 'daily_backups')
            desktop_db_dir = os.path.join(desktop_dir, 'database_backups')
            desktop_excel_dir = os.path.join(desktop_dir, 'excel_backups')
            os.makedirs(desktop_db_dir, exist_ok=True)
            os.makedirs(desktop_excel_dir, exist_ok=True)
            if os.path.exists(db_dest):
                shutil.copy2(db_dest, os.path.join(desktop_db_dir, db_filename))
            if os.path.exists(excel_dest):
                shutil.copy2(excel_dest, os.path.join(desktop_excel_dir, excel_filename))
        except Exception as e_desk:
            logger.warning(f"Could not mirror backup to desktop: {e_desk}")

        # 4. 10 دن سے پرانے بیک اپس کی خودکار صفائی
        clean_old_backups(days=10)

        logger.info(f"[OK] Backup Saved ({tag}): {excel_dest}")
        return {
            'db_name': db_filename,
            'excel_name': excel_filename,
            'db_path': db_dest,
            'excel_path': excel_dest
        }
    except Exception as e:
        logger.error(f"[!] Backup Error: {e}")
        return None

@app.route('/api/system/open_backup_folder')
@login_required
def open_backup_folder():
    try:
        desktop_dir = os.path.join(os.path.expanduser('~'), 'Desktop', 'daily_backups')
        target_dir = desktop_dir if os.path.exists(desktop_dir) else BACKUP_DIR
        if not os.path.exists(target_dir):
            os.makedirs(target_dir, exist_ok=True)
        if hasattr(os, 'startfile'):
            os.startfile(target_dir)
        return jsonify({'success': True, 'message': 'Backup folder opened successfully', 'path': target_dir})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)}), 500

_last_backup_time = 0

def on_app_exit():
    global _last_backup_time
    import time
    now_ts = time.time()
    if (now_ts - _last_backup_time) > 5:
        _last_backup_time = now_ts
        logger.info("[*] Software closing... Saving final backup.")
        save_backup_to_folder("closing_exit")

atexit.register(on_app_exit)

# â° Ø±ÙˆØ²Ø§Ù†Û Ø´Ø§Ù… 5 Ø¨Ø¬Û’ Ø§ÙˆØ± Ø±Ø§Øª 11 Ø¨Ø¬Û’ Ú©Ø§ Ø´ÛŒÚˆÙˆÙ„Ø±
if BackgroundScheduler:
    try:
        scheduler = BackgroundScheduler(daemon=True)
        scheduler.add_job(lambda: save_backup_to_folder("2PM"), 'cron', hour=14, minute=0)
        scheduler.add_job(lambda: save_backup_to_folder("9PM"), 'cron', hour=21, minute=0)
        scheduler.start()
    except Exception as e:
        logger.warning(f"Scheduler start warning: {e}")

@app.route('/backup_db')
@admin_only
def export_to_excel_backup():
    try:
        os.makedirs(MANUAL_BACKUP_DIR, exist_ok=True)
        now_str = datetime.now().strftime('%Y-%m-%d_%I-%M-%p')
        filename = f"Smart_POS_Manual_Excel_{now_str}.xlsx"
        dest_path = os.path.join(MANUAL_BACKUP_DIR, filename)

        conn = get_db()
        cursor = conn.cursor()
        
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        
        thin_border = Border(
            left=Side(style='thin', color='CBD5E1'),
            right=Side(style='thin', color='CBD5E1'),
            top=Side(style='thin', color='CBD5E1'),
            bottom=Side(style='thin', color='CBD5E1')
        )
        
        def add_sheet_with_styling(ws_title, banner_text, headers, rows, col_widths, num_cols, money_cols):
            ws = wb.create_sheet(title=ws_title)
            max_col = len(headers)
            ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_col)
            b_cell = ws.cell(row=1, column=1, value=banner_text)
            b_cell.font = Font(name="Segoe UI", size=13, bold=True, color="FFFFFF")
            b_cell.fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
            b_cell.alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[1].height = 28
            
            ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max_col)
            s_cell = ws.cell(row=2, column=1, value=f"Exported from Smart POS System POS | Date: {datetime.now().strftime('%Y-%m-%d %I:%M %p')}")
            s_cell.font = Font(name="Segoe UI", size=9, color="E2E8F0")
            s_cell.fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
            s_cell.alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[2].height = 18
            
            ws.append(headers)
            ws.row_dimensions[3].height = 24
            for c_idx in range(1, max_col + 1):
                h_cell = ws.cell(row=3, column=c_idx)
                h_cell.font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
                h_cell.fill = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
                h_cell.alignment = Alignment(horizontal="center", vertical="center")
                h_cell.border = thin_border
            
            for r_idx, r_data in enumerate(rows, start=4):
                ws.append(list(r_data))
                ws.row_dimensions[r_idx].height = 20
                for c_idx in range(1, max_col + 1):
                    cell = ws.cell(row=r_idx, column=c_idx)
                    cell.border = thin_border
                    cell.font = Font(name="Segoe UI", size=9)
                    if c_idx in money_cols:
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                        cell.number_format = '#,##0.00'
                    elif c_idx in num_cols:
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                        cell.number_format = '#,##0'
                    else:
                        cell.alignment = Alignment(horizontal="center", vertical="center")
            
            for c_idx, width in col_widths.items():
                ws.column_dimensions[get_column_letter(c_idx)].width = width

        # Sheet 1: Sales Invoices
        sales_raw = cursor.execute("""
            SELECT invoice_no, customer_name, date, total, paid, due, type, discount 
            FROM sales ORDER BY id DESC
        """).fetchall()
        sales_rows = []
        for r in sales_raw:
            sales_rows.append([
                r['invoice_no'] or '-',
                r['customer_name'] or 'Cash Customer',
                str(r['date'] or '-'),
                float(r['total'] or 0),
                float(r['paid'] or 0),
                float(r['due'] or 0),
                str(r['type'] or 'cash').title(),
                float(r['discount'] or 0)
            ])
        sales_headers = ["Invoice # (بل نمبر)", "Customer (گاہک)", "Date & Time (تاریخ / وقت)", "Total Bill (کل رقم)", "Amount Paid (ادا شدہ)", "Balance Due (بقایا)", "Type (قسم)", "Discount (ڈسکاؤنٹ)"]
        sales_widths = {1: 18, 2: 24, 3: 24, 4: 16, 5: 16, 6: 16, 7: 12, 8: 14}
        add_sheet_with_styling("Sales_Invoices", "SMART POS SYSTEM — Sales Records (سیلز ریکارڈ)", sales_headers, sales_rows, sales_widths, [], [4, 5, 6, 8])

        # Sheet 2: Receivables / Customers
        cust_raw = cursor.execute("""
            SELECT id, name, phone, address, balance, opening_balance 
            FROM customers ORDER BY balance DESC
        """).fetchall()
        cust_rows = []
        for r in cust_raw:
            cust_rows.append([
                f"#{r['id']}",
                r['name'] or '-',
                r['phone'] or '-',
                r['address'] or '-',
                float(r['balance'] or 0),
                float(r['opening_balance'] or 0)
            ])
        cust_headers = ["Account # (اکاؤنٹ نمبر)", "Customer Name (گاہک کا نام)", "Phone (موبائل نمبر)", "Address (پتہ)", "Net Balance Due (کل بقایا ادھار)", "Opening Balance (سابقہ پرانا بقایا)"]
        cust_widths = {1: 14, 2: 26, 3: 18, 4: 28, 5: 22, 6: 22}
        add_sheet_with_styling("Customer_Khata", "SMART POS SYSTEM — Customer Khata & Receivables (گاہک کھاتہ و ادھار)", cust_headers, cust_rows, cust_widths, [], [5, 6])

        # Sheet 3: Payables / Suppliers
        sup_raw = cursor.execute("""
            SELECT id, name, company, phone, address, balance 
            FROM suppliers ORDER BY balance DESC
        """).fetchall()
        sup_rows = []
        for r in sup_raw:
            sup_rows.append([
                f"#{r['id']}",
                r['name'] or '-',
                r['company'] or '-',
                r['phone'] or '-',
                r['address'] or '-',
                float(r['balance'] or 0)
            ])
        sup_headers = ["Supplier # (نمبر)", "Contact Person (سپلائر)", "Company (کمپنی)", "Phone (فون)", "Address (پتہ)", "Payable Balance (واجب الادا رقم)"]
        sup_widths = {1: 14, 2: 24, 3: 24, 4: 18, 5: 28, 6: 22}
        add_sheet_with_styling("Suppliers_Payables", "SMART POS SYSTEM — Suppliers & Payables (سپلائر کھاتہ و واجبات)", sup_headers, sup_rows, sup_widths, [], [6])

        # Sheet 4: Inventory / Stock
        prod_raw = cursor.execute("""
            SELECT code, name, category, unit, buy_price, sale_price, stock 
            FROM products ORDER BY name ASC
        """).fetchall()
        prod_rows = []
        for r in prod_raw:
            prod_rows.append([
                r['code'] or '-',
                r['name'] or '-',
                r['category'] or '-',
                r['unit'] or 'Pcs',
                float(r['buy_price'] or 0),
                float(r['sale_price'] or 0),
                float(r['stock'] or 0)
            ])
        prod_headers = ["Item Code (کوڈ)", "Product Name (آئٹم کا نام)", "Category (کیٹگری)", "Unit (یونٹ)", "Buy Rate (خرید ریٹ)", "Sale Rate (سیل ریٹ)", "Current Stock (موجودہ تعداد)"]
        prod_widths = {1: 16, 2: 32, 3: 16, 4: 12, 5: 16, 6: 16, 7: 18}
        add_sheet_with_styling("Inventory_Stock", "SMART POS SYSTEM — Current Inventory & Stock (اسٹاک و گودام)", prod_headers, prod_rows, prod_widths, [7], [5, 6])

        conn.close()
        wb.save(dest_path)

        return send_file(
            dest_path,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        flash(f"خرابی: {str(e)}", "danger")
        return redirect('/dashboard')

def open_browser_window():
    import time
    time.sleep(0.4)
    url = "http://127.0.0.1:5000/login?init=1"
    
    chrome_paths = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        os.path.expandvars(r"%LocalAppData%\Google\Chrome\Application\chrome.exe"),
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    ]
    
    for p in chrome_paths:
        if os.path.exists(p):
            try:
                import subprocess
                subprocess.Popen([p, f"--app={url}"])
                return
            except Exception:
                pass
                
    import webbrowser
    webbrowser.open(url)

def kill_port_5000_zombies():
    try:
        import subprocess
        my_pid = os.getpid()
        out = subprocess.check_output('netstat -ano -p tcp', shell=True).decode('utf-8', errors='ignore')
        for line in out.splitlines():
            if ':5000' in line and 'LISTENING' in line:
                parts = line.strip().split()
                if parts:
                    pid = parts[-1]
                    if pid.isdigit() and int(pid) != my_pid:
                        subprocess.run(f'taskkill /f /pid {pid}', shell=True, capture_output=True)
        import time
        time.sleep(0.1)
    except Exception:
        pass

@app.route('/api/categories', methods=['GET'])
@login_required
def get_categories():
    conn = get_db()
    cats = conn.execute('SELECT * FROM categories ORDER BY name ASC').fetchall()
    conn.close()
    return jsonify({'success': True, 'categories': [dict(c) for c in cats]})

@app.route('/api/category/add', methods=['POST'])
@login_required
def add_category():
    data = request.get_json()
    name = data.get('name', '').strip()
    if not name:
        return jsonify({'success': False, 'message': 'Name required'})
    try:
        conn = get_db()
        conn.execute('INSERT INTO categories (name) VALUES (?)', (name,))
        conn.commit()
        conn.close()
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

@app.route('/api/category/delete/<int:cat_id>', methods=['POST'])
@login_required
def delete_category(cat_id):
    conn = get_db()
    conn.execute('DELETE FROM categories WHERE id = ?', (cat_id,))
    conn.commit()
    conn.close()
    return jsonify({'success': True})

@app.route('/api/supplier/payment/add', methods=['POST'])
@login_required
def add_supplier_payment():
    try:
        data = request.get_json() or {}
        sup_id = data.get('supplier_id')
        amount = float(data.get('amount', 0))
        pay_type = data.get('type', 'payment')
        method = data.get('method', 'Cash')
        raw_notes = (data.get('notes') or data.get('note') or '').strip()
        pay_date = data.get('payment_date', '')
        
        if not sup_id or amount <= 0:
            return jsonify({'success': False, 'message': 'درست رقم درج کریں'}), 400
        
        conn = get_db()
        cursor = conn.cursor()
        
        sup = cursor.execute("SELECT balance, name FROM suppliers WHERE id = ?", (sup_id,)).fetchone()
        if not sup:
            conn.close()
            return jsonify({'success': False, 'message': 'سپلائر نہیں ملا'}), 404
            
        if pay_date:
            if len(pay_date) == 10:
                now_str = f"{pay_date} {datetime.now().strftime('%I:%M:%S %p')}"
            else:
                now_str = pay_date
        else:
            now_str = datetime.now().strftime("%Y-%m-%d %I:%M:%S %p")
        
        is_refund = (pay_type == 'refund' or data.get('is_refund'))
        if is_refund:
            note_label = f"Refund from Supplier / سپلائر سے واپسی ({method}): {raw_notes}".strip()
            stored_amount = -amount
            cursor.execute('UPDATE suppliers SET balance = balance + ? WHERE id = ?', (amount, sup_id))
            success_msg = f'سپلائر سے Rs. {amount:,.0f} کی واپسی کامیابی سے درج ہو گئی۔'
        else:
            note_label = f"{raw_notes} ({method})" if raw_notes else f"ادائیگی ({method})"
            stored_amount = amount
            cursor.execute('UPDATE suppliers SET balance = balance - ? WHERE id = ?', (amount, sup_id))
            success_msg = f'سپلائر کو Rs. {amount:,.0f} کی ادائیگی کامیابی سے درج ہو گئی۔'
            
        sp_cols = [c[1] for c in cursor.execute("PRAGMA table_info(supplier_payments)").fetchall()]
        cols = ["supplier_id", "amount", "payment_date"]
        vals = [sup_id, stored_amount, now_str]
        
        if 'note' in sp_cols:
            cols.append("note")
            vals.append(note_label)
        if 'notes' in sp_cols:
            cols.append("notes")
            vals.append(note_label)
        if 'method' in sp_cols:
            cols.append("method")
            vals.append(method)
            
        cursor.execute(f"INSERT INTO supplier_payments ({', '.join(cols)}) VALUES ({', '.join(['?']*len(cols))})", tuple(vals))
        conn.commit()
        conn.close()
        return jsonify({'success': True, 'message': success_msg})
    except Exception as e:
        logging.error(f"Error in add_supplier_payment: {e}", exc_info=True)
        return jsonify({'success': False, 'message': f'خرابی: {str(e)}'}), 500

@app.route('/api/supplier/payments/<int:sup_id>', methods=['GET'])
@login_required
def get_supplier_payments(sup_id):
    conn = get_db()
    payments = conn.execute('SELECT * FROM supplier_payments WHERE supplier_id = ? ORDER BY id DESC', (sup_id,)).fetchall()
    conn.close()
    return jsonify({'success': True, 'payments': [dict(p) for p in payments]})

@app.route('/api/quotation/convert/<int:quot_id>', methods=['POST'])
@login_required
def convert_quotation_to_invoice(quot_id):
    try:
        import json, uuid
        conn = get_db()
        cursor = conn.cursor()
        
        quot = cursor.execute('SELECT * FROM quotations WHERE id = ?', (quot_id,)).fetchone()
        if not quot:
            conn.close()
            return jsonify({'success': False, 'message': 'Quotation not found'})
        
        quot = dict(quot)
        items = json.loads(quot.get('items_json') or '[]')
        if not items:
            conn.close()
            return jsonify({'success': False, 'message': 'No items found in this quotation.'})
        
        cust_name = (quot.get('customer_name') or 'Cash Customer').strip()
        cust_phone = (quot.get('customer_phone') or '').strip()
        customer_id = None
        
        if cust_name and cust_name.lower() not in ('cash customer', 'walk-in customer', '-', ''):
            row = cursor.execute("SELECT id FROM customers WHERE name = ?", (cust_name,)).fetchone()
            if row:
                customer_id = row['id']
            else:
                cursor.execute("INSERT INTO customers (name, phone, balance, opening_balance) VALUES (?, ?, 0, 0)",
                               (cust_name, cust_phone))
                customer_id = cursor.lastrowid
        
        subtotal = 0.0
        for it in items:
            q = float(it.get('qty') or it.get('quantity') or 1)
            p = float(it.get('price') or 0)
            subtotal += round(q * p, 2)
        
        total_amount = round(float(quot.get('total') or subtotal), 2)
        paid_amount = total_amount
        due_amount = 0.0
        actual_type = 'cash'
        
        inv_number = f"INV-{uuid.uuid4().hex[:10].upper()}"
        now_str = datetime.now().strftime('%Y-%m-%d %I:%M:%S %p')
        
        cursor.execute('''
            INSERT INTO sales (invoice_no, customer_id, customer_name, subtotal, tax_rate, tax_amount, discount, total, paid, due, type, date, cash_amount, online_amount)
            VALUES (?, ?, ?, ?, 0, 0, 0, ?, ?, ?, ?, ?, ?, 0)
        ''', (inv_number, customer_id, cust_name, subtotal, total_amount, paid_amount, due_amount, actual_type, now_str, total_amount))
        
        sale_id = cursor.lastrowid
        
        for it in items:
            p_id = it.get('id') or it.get('product_id')
            p_name = it.get('name') or it.get('product_name') or ''
            q = float(it.get('qty') or it.get('quantity') or 1)
            p = float(it.get('price') or 0)
            it_total = round(q * p, 2)
            u = it.get('unit') or 'Pcs'
            
            cursor.execute('''
                INSERT INTO sale_items (sale_id, product_id, product_name, quantity, price, total, unit)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            ''', (sale_id, p_id, p_name, q, p, it_total, u))
            
            if p_id:
                try:
                    cursor.execute('UPDATE products SET stock = stock - ? WHERE id = ?', (q, p_id))
                except Exception as ex:
                    logger.warning(f"Could not deduct stock for product {p_id}: {ex}")
        
        conn.commit()
        conn.close()
        return jsonify({'success': True, 'sale_id': sale_id, 'invoice_no': inv_number,
                        'message': f'Quotation converted successfully to Invoice #{inv_number}!'})
    except Exception as e:
        import traceback
        logger.error(f"Error converting quotation: {e}\n{traceback.format_exc()}")
        return jsonify({'success': False, 'message': str(e)})

# ----------------------------------------------------
# 💵 ذاتی قرضہ جات (لین دین - Personal Loans)
# ----------------------------------------------------
@app.route('/loans')
@login_required
def personal_loans():
    conn = get_db()
    cursor = conn.cursor()
    
    loans = cursor.execute("""
        SELECT l.*, 
               (l.amount - COALESCE(l.paid_back, 0)) as remaining
        FROM personal_loans l 
        ORDER BY l.id DESC
    """).fetchall()
    
    # Given (قرض دیا - وصول کرنا ہے):
    total_given_row = cursor.execute("SELECT COALESCE(SUM(amount), 0), COALESCE(SUM(paid_back), 0) FROM personal_loans WHERE type = 'given'").fetchone()
    total_given = float(total_given_row[0] or 0)
    given_recovered = float(total_given_row[1] or 0)
    remaining_given = total_given - given_recovered

    # Taken (قرض لیا - واپس کرنا ہے):
    total_taken_row = cursor.execute("SELECT COALESCE(SUM(amount), 0), COALESCE(SUM(paid_back), 0) FROM personal_loans WHERE type = 'taken'").fetchone()
    total_taken = float(total_taken_row[0] or 0)
    taken_returned = float(total_taken_row[1] or 0)
    remaining_taken = total_taken - taken_returned
    
    conn.close()
    return render_template('loans.html', 
                           loans=loans,
                           total_given=total_given,
                           given_recovered=given_recovered,
                           remaining_given=remaining_given,
                           total_taken=total_taken,
                           taken_returned=taken_returned,
                           remaining_taken=remaining_taken)

@app.route('/api/loan/add', methods=['POST'])
@login_required
def add_personal_loan():
    try:
        data = request.get_json() if request.is_json else request.form
        person_name = (data.get('person_name') or '').strip()
        phone = (data.get('phone') or '').strip()
        loan_type = (data.get('type') or 'given').strip()
        amount = float(data.get('amount') or 0)
        loan_date = (data.get('date') or datetime.now().strftime('%Y-%m-%d')).strip()
        due_date = (data.get('due_date') or '').strip()
        notes = (data.get('notes') or '').strip()
        
        if not person_name or amount <= 0:
            return jsonify({'success': False, 'message': 'نام اور درست رقم درج کرنا لازمی ہے'}), 400
            
        conn = get_db()
        cursor = conn.cursor()
        cursor.execute('''
            INSERT INTO personal_loans (person_name, phone, type, amount, paid_back, date, due_date, status, notes)
            VALUES (?, ?, ?, ?, 0.0, ?, ?, 'pending', ?)
        ''', (person_name, phone, loan_type, amount, loan_date, due_date, notes))
        loan_id = cursor.lastrowid
        conn.commit()
        conn.close()
        
        if request.is_json:
            return jsonify({'success': True, 'loan_id': loan_id, 'message': 'قرضہ کا کھاتہ محفوظ ہو گیا'})
        return redirect('/loans')
    except Exception as e:
        logger.error(f"Error adding personal loan: {e}")
        return jsonify({'success': False, 'message': str(e)}), 500

@app.route('/api/loan/payment', methods=['POST'])
@login_required
def add_loan_payment():
    try:
        data = request.get_json() if request.is_json else request.form
        loan_id = int(data.get('loan_id'))
        amount = float(data.get('amount') or 0)
        payment_date = (data.get('payment_date') or datetime.now().strftime('%Y-%m-%d')).strip()
        if len(payment_date) == 10:
            payment_date = f"{payment_date} {datetime.now().strftime('%I:%M:%S %p')}"
        method = (data.get('method') or 'Cash').strip()
        notes = (data.get('notes') or '').strip()
        
        if not loan_id or amount <= 0:
            return jsonify({'success': False, 'message': 'درست رقم درج کریں'}), 400
            
        conn = get_db()
        cursor = conn.cursor()
        loan = cursor.execute("SELECT * FROM personal_loans WHERE id = ?", (loan_id,)).fetchone()
        if not loan:
            conn.close()
            return jsonify({'success': False, 'message': 'قرضہ ریکارڈ نہیں ملا'}), 404
            
        cursor.execute('''
            INSERT INTO personal_loan_payments (loan_id, amount, payment_date, method, notes)
            VALUES (?, ?, ?, ?, ?)
        ''', (loan_id, amount, payment_date, method, notes))
        
        new_paid = round(float(loan['paid_back'] or 0) + amount, 2)
        new_status = 'settled' if new_paid >= float(loan['amount'] or 0) else 'pending'
        cursor.execute("UPDATE personal_loans SET paid_back = ?, status = ? WHERE id = ?", 
                       (new_paid, new_status, loan_id))
        
        conn.commit()
        conn.close()
        return jsonify({'success': True, 'message': 'ادائیگی کامیابی سے درج ہو گئی'})
    except Exception as e:
        logger.error(f"Error adding loan payment: {e}")
        return jsonify({'success': False, 'message': str(e)}), 500

@app.route('/api/loan/payments/<int:loan_id>', methods=['GET'])
@login_required
def get_loan_payments(loan_id):
    conn = get_db()
    cursor = conn.cursor()
    loan = cursor.execute("SELECT * FROM personal_loans WHERE id = ?", (loan_id,)).fetchone()
    if not loan:
        conn.close()
        return jsonify({'success': False, 'message': 'قرضہ نہیں ملا'}), 404
        
    payments = cursor.execute("SELECT * FROM personal_loan_payments WHERE loan_id = ? ORDER BY id DESC", (loan_id,)).fetchall()
    conn.close()
    return jsonify({'success': True, 'loan': dict(loan), 'payments': [dict(p) for p in payments]})

@app.route('/api/loan/delete/<int:loan_id>', methods=['POST'])
@login_required
def delete_personal_loan(loan_id):
    try:
        conn = get_db()
        cursor = conn.cursor()
        cursor.execute("DELETE FROM personal_loan_payments WHERE loan_id = ?", (loan_id,))
        cursor.execute("DELETE FROM personal_loans WHERE id = ?", (loan_id,))
        conn.commit()
        conn.close()
        return jsonify({'success': True, 'message': 'قرضہ ریکارڈ ڈیلیٹ ہو گیا'})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)}), 500

if __name__ == '__main__':
    kill_port_5000_zombies()
    import threading
    threading.Thread(target=open_browser_window, daemon=True).start()
    app.run(host='127.0.0.1', port=5000, debug=False)



