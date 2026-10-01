import os
import shutil
import sqlite3
from datetime import datetime
from werkzeug.security import generate_password_hash

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_FILE = os.path.join(BASE_DIR, 'pos.db')

def reset_database():
    print("==================================================")
    print("      SMART POS - FACTORY RESET / CLEAN SCRIPT    ")
    print("==================================================")
    
    if os.path.exists(DB_FILE):
        # 1. Take a safe backup of current data first
        backup_name = f"pos_backup_before_clean_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db"
        backup_path = os.path.join(BASE_DIR, backup_name)
        shutil.copy2(DB_FILE, backup_path)
        print(f"[OK] Backup saved to: {backup_name}")

        # 2. Connect and truncate all business data
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()

        tables_to_clear = [
            'products',
            'product_units',
            'customers',
            'customer_payments',
            'customer_ledger',
            'suppliers',
            'supplier_payments',
            'purchases',
            'purchase_items',
            'sales',
            'sale_items',
            'expenses',
            'quotations',
            'personal_loans',
            'personal_loan_payments'
        ]

        for table in tables_to_clear:
            try:
                cursor.execute(f"DELETE FROM {table}")
                print(f"[OK] Cleared table: {table}")
            except sqlite3.OperationalError:
                pass

        # Reset all auto-increment counters back to 1
        try:
            cursor.execute("DELETE FROM sqlite_sequence")
            print("[OK] Reset all ID counters back to 1 (Fresh start)")
        except sqlite3.OperationalError:
            pass

        # Ensure default login users exist
        cursor.execute("DELETE FROM users")
        cursor.execute("INSERT INTO users (username, password, role, display_name) VALUES ('admin', ?, 'admin', 'Admin')",
                       (generate_password_hash('admin123'),))
        cursor.execute("INSERT INTO users (username, password, role, display_name) VALUES ('cashier', ?, 'cashier', 'Cashier')",
                       (generate_password_hash('1234'),))
        print("[OK] Default Logins Ready:")
        print("     - Admin: username='admin', password='admin123'")
        print("     - Cashier: username='cashier', password='1234'")

        conn.commit()
        cursor.execute("VACUUM")
        conn.close()

        print("\n==================================================")
        print(" [SUCCESS] Software is now 100% CLEAN & FRESH!")
        print(" Ready for project demonstration!")
        print("==================================================")
    else:
        print("[INFO] pos.db not found. Starting app.py will create a fresh one.")

if __name__ == '__main__':
    reset_database()
