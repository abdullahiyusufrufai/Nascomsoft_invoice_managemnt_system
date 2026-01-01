"""
Nascomsoft Invoice Manager System (Enterprise Edition v2.1)
Developed by: Senior Python Architect (AI)
Date: December 30, 2025
Updates: Restructured with Separate Tabs for Projects & Components
"""

import tkinter as tk
from tkinter import messagebox, ttk
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from datetime import datetime
import mysql.connector
from mysql.connector import Error
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle
import os
import sys
import textwrap
import csv
import subprocess  # Required for opening files on non-Windows systems
from tkinter import filedialog
import smtplib
import ssl
import mimetypes
from email.message import EmailMessage

# =============================================================================
# 1. SYSTEM CONFIGURATION
# =============================================================================

COMPANY_CONFIG = {
    "company_name": "NASCOMSOFT EMBEDED",  
    "address": "Anguwan Cashew, Off Dass Road,\nOpposite Elim Church, 740102,\nYelwa, Bauchi State.",
    "tin": "22843418-0001",
    "currency_symbol": "N",
    "vat_rate": 0.075,
    "bank_name": "First Bank",
    "account_number": "2037467351",
    "account_name": "Nascomsoft Embedded"
}

DB_SETTINGS = {
    'host': 'localhost',
    'user': 'root',
    'password': '',
    'database': 'nascomsoft_billing_db'
}

LOGO_FILENAME = "LOGO.png"  # Ensure this file exists in the same directory as the script
if not os.path.exists(LOGO_FILENAME):
    LOGO_FILENAME = None  # Set to None if the file is missing to avoid runtime errors

# Default SMTP/email settings (in-memory; configure via UI). The default from-email is the company address.
SMTP_SETTINGS = {
    'host': '',
    'port': 587,
    'username': '',
    'password': '',
    'use_tls': True,
    'from_email': 'info@nascomsoft.com'
}

# =============================================================================
# 2. DATABASE MANAGER (AUTO-MIGRATING)
# =============================================================================

class DatabaseManager:
    def __init__(self):
        self.conn = None
        self.cursor = None
        self.check_connection()

    def check_connection(self):
        try:
            conn = mysql.connector.connect(
                host=DB_SETTINGS['host'],
                user=DB_SETTINGS['user'],
                password=DB_SETTINGS['password']
            )
            if conn.is_connected():
                cursor = conn.cursor()
                cursor.execute(f"CREATE DATABASE IF NOT EXISTS {DB_SETTINGS['database']}")
                conn.close()
                return True
            return False
        except Error as e:
            # Do not terminate the entire application if DB is unavailable; run in offline mode.
            print(f"Database Warning: Cannot reach MySQL. Running in offline mode. Details: {e}")
            return False

    def get_connection(self):
        try:
            self.conn = mysql.connector.connect(
                host=DB_SETTINGS['host'],
                user=DB_SETTINGS['user'],
                password=DB_SETTINGS['password'],
                database=DB_SETTINGS['database']
            )
            self.cursor = self.conn.cursor(buffered=True)
            self.create_tables()
            return self.conn
        except Error as e:
            print(f"Database Error: Connection lost. Details: {e}")
            return None

    def create_tables(self):
        # Ensure connection is established
        if not self.conn or not self.conn.is_connected():
            self.get_connection()
        if not self.conn or not self.conn.is_connected():  # Ensure connection is established
            print("Database Error: Unable to establish a database connection.")
            return False

        # Create table if it doesn't exist
        query_invoices = """
        CREATE TABLE IF NOT EXISTS invoices (
            id INT AUTO_INCREMENT PRIMARY KEY,
            invoice_number VARCHAR(50) UNIQUE NOT NULL,
            client_name VARCHAR(100) NOT NULL,
            client_email VARCHAR(100),
            client_address VARCHAR(255),
            invoice_type VARCHAR(50),
            date_issued DATETIME DEFAULT CURRENT_TIMESTAMP,
            subtotal DECIMAL(15, 2),
            vat_amount DECIMAL(15, 2),
            shipping_cost DECIMAL(15, 2),
            wht_amount DECIMAL(15, 2),
            wht_rate DECIMAL(5, 2),
            grand_total DECIMAL(15, 2)
        )
        """
        self.cursor.execute(query_invoices)

        # SELF-HEALING: Check if client_email exists
        try:
            self.cursor.execute("SELECT client_email FROM invoices LIMIT 1")
            self.cursor.fetchone()
        except Error:
            self.cursor.execute("ALTER TABLE invoices ADD COLUMN client_email VARCHAR(100) AFTER client_name")
            self.conn.commit()
        
        # SELF-HEALING: Check if client_address exists
        try:
            self.cursor.execute("SELECT client_address FROM invoices LIMIT 1")
            self.cursor.fetchone()
        except Error:
            self.cursor.execute("ALTER TABLE invoices ADD COLUMN client_address VARCHAR(255) AFTER client_name")
            self.conn.commit()

        # SELF-HEALING: Check if invoice_type exists
        try:
            self.cursor.execute("SELECT invoice_type FROM invoices LIMIT 1")
            self.cursor.fetchone()
        except Error:
            self.cursor.execute("ALTER TABLE invoices ADD COLUMN invoice_type VARCHAR(50) AFTER client_address")
            self.conn.commit()

        # SELF-HEALING: Check if shipping_cost exists
        try:
            self.cursor.execute("SELECT shipping_cost FROM invoices LIMIT 1")
            self.cursor.fetchone()
        except Error:
            self.cursor.execute("ALTER TABLE invoices ADD COLUMN shipping_cost DECIMAL(15, 2) AFTER vat_amount")
            self.conn.commit()

        # SELF-HEALING: Check if wht_rate exists
        try:
            self.cursor.execute("SELECT wht_rate FROM invoices LIMIT 1")
            self.cursor.fetchone()
        except Error:
            self.cursor.execute("ALTER TABLE invoices ADD COLUMN wht_rate DECIMAL(5, 2) AFTER wht_amount")
            self.conn.commit()

        # SELF-HEALING: Remove net_payable column if it exists
        try:
            self.cursor.execute("SELECT net_payable FROM invoices LIMIT 1")
            self.cursor.fetchone()
            self.cursor.execute("ALTER TABLE invoices DROP COLUMN net_payable")
            self.conn.commit()
        except Error:
            pass

        self.conn.commit()

        # Create quotations table if it doesn't exist
        query_quotes = """
        CREATE TABLE IF NOT EXISTS quotations (
            id INT AUTO_INCREMENT PRIMARY KEY,
            quote_number VARCHAR(50) UNIQUE NOT NULL,
            client_name VARCHAR(100) NOT NULL,
            client_email VARCHAR(100),
            client_address VARCHAR(255),
            date_issued DATETIME DEFAULT CURRENT_TIMESTAMP,
            subtotal DECIMAL(15, 2),
            vat_amount DECIMAL(15, 2),
            shipping_cost DECIMAL(15, 2),
            grand_total DECIMAL(15, 2)
        )
        """
        try:
            self.cursor.execute(query_quotes)
            self.conn.commit()
        except Error as e:
            print(f"Error creating quotations table: {e}")
            # continue without stopping the app

        # SELF-HEALING: Check if client_email exists in quotations
        try:
            self.cursor.execute("SELECT client_email FROM quotations LIMIT 1")
            self.cursor.fetchone()
        except Error:
            try:
                self.cursor.execute("ALTER TABLE quotations ADD COLUMN client_email VARCHAR(100) AFTER client_name")
                self.conn.commit()
            except Error:
                pass
        
        # SELF-HEALING: ensure basic columns exist (no-op if table present)
        try:
            self.cursor.execute("SELECT quote_number FROM quotations LIMIT 1")
            self.cursor.fetchone()
        except Error:
            pass
        self.conn.commit()

    def save_invoice(self, data):
        if not self.conn or not self.conn.is_connected():
            self.get_connection()
            
        sql = """
        INSERT INTO invoices 
        (invoice_number, client_name, client_email, client_address, invoice_type, subtotal, vat_amount, shipping_cost, wht_amount, wht_rate, grand_total) 
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        """
        vals = (
            data['invoice_no'], data['client_name'], data.get('client_email', ''), data['client_address'], data['invoice_type'],
            data['subtotal'], data['vat'], data['shipping'], data['wht'], data['wht_rate'], data['grand_total']
        )
        try:
            self.cursor.execute(sql, vals)
            self.conn.commit()
            return True
        except Error as e:
            print(f"Save Error: Failed to save. Details: {e}")
            return False

    def generate_invoice_number(self):
        if not self.conn or not self.conn.is_connected():
            self.get_connection()
        if self.conn and self.conn.is_connected():
            self.cursor.execute("SELECT id FROM invoices ORDER BY id DESC LIMIT 1")
            result = self.cursor.fetchone()
            year = datetime.now().year
            next_id = 1
            if result:
                next_id = result[0] + 1
            return f"NSE-INV-{year}-{next_id:04d}"
        else:
            raise Error("Database connection could not be established.")

    def fetch_invoices(self, filters=None, page=1, page_size=25):
        """Return a list of invoices matching optional filters.
        filters: dict with keys: invoice_no, client_name, invoice_type, date_from, date_to
        """
        try:
            # Ensure connection
            if not self.conn or not self.conn.is_connected():
                self.get_connection()
            if not self.cursor:
                return []

            sql = "SELECT invoice_number, date_issued, client_name, client_email, invoice_type, subtotal, vat_amount, shipping_cost, wht_amount, wht_rate, grand_total FROM invoices"
            where = []
            params = []
            if filters:
                if filters.get('invoice_no'):
                    where.append("invoice_number LIKE %s")
                    params.append(f"%{filters['invoice_no']}%")
                if filters.get('client_name'):
                    where.append("client_name LIKE %s")
                    params.append(f"%{filters['client_name']}%")
                if filters.get('invoice_type') and filters.get('invoice_type') != 'All':
                    where.append("invoice_type = %s")
                    params.append(filters['invoice_type'])
                if filters.get('date_from'):
                    where.append("date_issued >= %s")
                    params.append(filters['date_from'])
                if filters.get('date_to'):
                    where.append("date_issued <= %s")
                    params.append(filters['date_to'])
            if where:
                sql += " WHERE " + " AND ".join(where)
            sql += " ORDER BY date_issued DESC LIMIT %s OFFSET %s"
            params.extend([page_size, (page - 1) * page_size])
            self.cursor.execute(sql, tuple(params))
            rows = self.cursor.fetchall()
            results = []
            for r in rows:
                # Format date_issued as string
                date_val = r[1]
                try:
                    date_str = date_val.strftime('%Y-%m-%d %H:%M:%S') if hasattr(date_val, 'strftime') else str(date_val)
                except Exception:
                    date_str = str(date_val)
                results.append({
                    'invoice_no': r[0],
                    'date_issued': date_str,
                    'client_name': r[2],
                    'client_email': r[3],
                    'invoice_type': r[4],
                    'subtotal': float(r[5]) if r[5] is not None else 0.0,
                    'vat': float(r[6]) if r[6] is not None else 0.0,
                    'shipping': float(r[7]) if r[7] is not None else 0.0,
                    'wht': float(r[8]) if r[8] is not None else 0.0,
                    'wht_rate': float(r[9]) if r[9] is not None else 0.0,
                    'grand_total': float(r[10]) if r[10] is not None else 0.0
                })
            return results
        except Exception as e:
            print(f"Fetch Error: {e}")
            return []

    def delete_invoice(self, invoice_number):
        try:
            if not self.conn or not self.conn.is_connected():
                self.get_connection()
            if not self.cursor:
                return False
            self.cursor.execute("DELETE FROM invoices WHERE invoice_number = %s", (invoice_number,))
            self.conn.commit()
            return True
        except Exception as e:
            print(f"Delete Invoice Error: {e}")
            return False

    def delete_quotation(self, quote_number):
        try:
            if not self.conn or not self.conn.is_connected():
                self.get_connection()
            if not self.cursor:
                return False
            self.cursor.execute("DELETE FROM quotations WHERE quote_number = %s", (quote_number,))
            self.conn.commit()
            return True
        except Exception as e:
            print(f"Delete Quote Error: {e}")
            return False

    def generate_quotation_number(self):
        if not self.conn or not self.conn.is_connected():
            self.get_connection()
        if self.conn and self.conn.is_connected():
            self.cursor.execute("SELECT id FROM quotations ORDER BY id DESC LIMIT 1")
            result = self.cursor.fetchone()
            year = datetime.now().year
            next_id = 1
            if result:
                next_id = result[0] + 1
            return f"NSE-QTN-{year}-{next_id:04d}"
        else:
            raise Error("Database connection could not be established.")

    def save_quotation(self, data):
        if not self.conn or not self.conn.is_connected():
            self.get_connection()
        sql = """
        INSERT INTO quotations (quote_number, client_name, client_email, client_address, subtotal, vat_amount, shipping_cost, grand_total)
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
        """
        vals = (
            data['quote_no'], data['client_name'], data.get('client_email', ''), data['client_address'], data['subtotal'], data['vat'], data['shipping'], data['grand_total']
        )
        try:
            self.cursor.execute(sql, vals)
            self.conn.commit()
            return True
        except Exception as e:
            print(f"Save Quote Error: {e}")
            return False

    def fetch_quotations(self, filters=None, page=1, page_size=25):
        try:
            if not self.conn or not self.conn.is_connected():
                self.get_connection()
            if not self.cursor:
                return []
            sql = "SELECT quote_number, date_issued, client_name, client_email, subtotal, vat_amount, shipping_cost, grand_total FROM quotations"
            where = []
            params = []
            if filters:
                if filters.get('invoice_no'):
                    where.append("quote_number LIKE %s")
                    params.append(f"%{filters['invoice_no']}%")
                if filters.get('client_name'):
                    where.append("client_name LIKE %s")
                    params.append(f"%{filters['client_name']}%")
            if where:
                sql += " WHERE " + " AND ".join(where)
            sql += " ORDER BY date_issued DESC LIMIT %s OFFSET %s"
            params.extend([page_size, (page - 1) * page_size])
            self.cursor.execute(sql, tuple(params))
            rows = self.cursor.fetchall()
            results = []
            for r in rows:
                date_val = r[1]
                try:
                    date_str = date_val.strftime('%Y-%m-%d %H:%M:%S') if hasattr(date_val, 'strftime') else str(date_val)
                except Exception:
                    date_str = str(date_val)
                results.append({
                    'invoice_no': r[0],
                    'date_issued': date_str,
                    'client_name': r[2],
                    'client_email': r[3],
                    'invoice_type': 'Quotation',
                    'subtotal': float(r[4]) if r[4] is not None else 0.0,
                    'vat': float(r[5]) if r[5] is not None else 0.0,
                    'shipping': float(r[6]) if r[6] is not None else 0.0,
                    'wht': 0.0,
                    'wht_rate': 0.0,
                    'grand_total': float(r[7]) if r[7] is not None else 0.0
                })
            return results
        except Exception as e:
            print(f"Fetch Quotes Error: {e}")
            return []

    def save_email_log(self, to_address, subject, attachment, status, error_message=None):
        """Persist an email delivery record to the database."""
        try:
            if not self.conn or not self.conn.is_connected():
                self.get_connection()
            sql = """
            INSERT INTO email_deliveries (to_address, subject, attachment, status, error_message)
            VALUES (%s, %s, %s, %s, %s)
            """
            vals = (to_address, subject, attachment or '', status, error_message or '')
            self.cursor.execute(sql, vals)
            self.conn.commit()
            return True
        except Exception as e:
            print(f"Email Log Save Error: {e}")
            return False

    def fetch_email_logs(self, limit=500):
        try:
            if not self.conn or not self.conn.is_connected():
                self.get_connection()
            self.cursor.execute("SELECT id, created_at, to_address, subject, attachment, status, error_message FROM email_deliveries ORDER BY created_at DESC LIMIT %s", (limit,))
            rows = self.cursor.fetchall()
            results = []
            for r in rows:
                results.append({
                    'id': r[0],
                    'created_at': r[1].strftime('%Y-%m-%d %H:%M:%S') if hasattr(r[1], 'strftime') else str(r[1]),
                    'to_address': r[2],
                    'subject': r[3],
                    'attachment': r[4],
                    'status': r[5],
                    'error_message': r[6]
                })
            return results
        except Exception as e:
            print(f"Fetch Email Logs Error: {e}")
            return []

# =============================================================================
# 3. PDF ENGINE
# =============================================================================

class InvoicePDF:
    def __init__(self, filename):
        self.filename = filename
        self.c = canvas.Canvas(filename, pagesize=A4)
        self.width, self.height = A4

    def draw_header(self, invoice_no, date_str, doc_type="INVOICE"):
        # Logo
        if LOGO_FILENAME and os.path.exists(LOGO_FILENAME):
            try:
                self.c.drawImage(LOGO_FILENAME, 30, self.height - 110, width=80, height=80, mask='auto')
            except Exception as e:
                print(f"Error loading logo: {e}")

        # Company Details
        self.c.setFont("Helvetica-Bold", 18)
        self.c.setFillColor(colors.HexColor("#0f3057"))
        self.c.drawRightString(self.width - 30, self.height - 50, COMPANY_CONFIG["company_name"])
        
        self.c.setFont("Helvetica", 10)
        self.c.setFillColor(colors.black)
        
        # Multi-line company address
        y_text = self.height - 70
        self.c.setFont("Helvetica", 10)
        for line in COMPANY_CONFIG["address"].split('\n'):
            self.c.drawRightString(self.width - 30, y_text, line)
            y_text -= 12
        
        y_pos = y_text - 10
        self.c.setFont("Helvetica-Bold", 10)
        self.c.drawRightString(self.width - 30, y_pos, f"TIN: {COMPANY_CONFIG['tin']}")

        # Document Banner
        self.c.setStrokeColor(colors.HexColor("#0f3057"))
        self.c.line(30, self.height - 130, self.width - 30, self.height - 130)
        
        self.c.setFont("Helvetica-Bold", 22)
        self.c.setFillColor(colors.HexColor("#e94560"))
        self.c.drawString(30, self.height - 160, doc_type)
        
        self.c.setFont("Helvetica-Bold", 12)
        self.c.setFillColor(colors.black)
        self.c.drawString(30, self.height - 180, f"{doc_type} #: {invoice_no}")
        self.c.drawString(30, self.height - 195, f"Date: {date_str}")

    def draw_client_info(self, name, address):
        self.c.setFont("Helvetica-Bold", 12)
        self.c.drawString(self.width - 250, self.height - 160, "BILL TO:")
        
        self.c.setFont("Helvetica-Bold", 12)
        self.c.drawString(self.width - 250, self.height - 180, name)
        
        # Render Client Address (Multi-line)
        self.c.setFont("Helvetica", 10)
        text_obj = self.c.beginText()
        text_obj.setTextOrigin(self.width - 250, self.height - 195)
        
        # Wrap long addresses so they don't run off page
        wrapped_address = textwrap.wrap(address, width=35) 
        for line in wrapped_address:
            text_obj.textLine(line)
        self.c.drawText(text_obj)

    def draw_items_table(self, items):
        data = [['S/N', 'Description', 'Type', 'Qty', 'Rate', 'Amount']]
        for item in items:
            data.append([
                item.get('sn', ''),
                item['desc'],
                item['type'],
                str(item['qty']),
                f"{COMPANY_CONFIG['currency_symbol']}{item['price']:,.2f}",
                f"{COMPANY_CONFIG['currency_symbol']}{item['total']:,.2f}"
            ])
            
        table = Table(data, colWidths=[50, 220, 70, 40, 90, 90])
        style = TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#0f3057")),
            ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
            ('ALIGN', (0,0), (-1,-1), 'LEFT'),
            ('ALIGN', (3,0), (-1,-1), 'CENTER'),
            ('ALIGN', (4,0), (-1,-1), 'RIGHT'),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0,0), (-1,0), 12),
            ('GRID', (0,0), (-1,-1), 1, colors.lightgrey),
        ])
        table.setStyle(style)
        
        w, h = table.wrap(self.width, self.height)
        self.y_position = self.height - 250 - h
        table.drawOn(self.c, 30, self.y_position)

    def draw_footer(self, totals):
        x_label = self.width - 200
        x_val = self.width - 35
        y = self.y_position - 30
        
        def print_line(label, val, is_bold=False, color=colors.black):
            self.c.setFillColor(color)
            font = "Helvetica-Bold" if is_bold else "Helvetica"
            self.c.setFont(font, 10 if not is_bold else 12)
            self.c.drawRightString(x_label, y, label)
            self.c.drawRightString(x_val, y, f"{COMPANY_CONFIG['currency_symbol']}{val:,.2f}")
        
        print_line("Subtotal:", totals['subtotal'])
        y -= 20
        print_line("VAT (7.5%):", totals['vat'])
        y -= 20
        print_line("Shipping Cost:", totals['shipping'])
        y -= 20
        self.c.setStrokeColor(colors.grey)
        self.c.line(x_label - 50, y + 15, x_val, y + 15)
        print_line("Grand Total:", totals['grand_total'], is_bold=True)
        y -= 20
        
        # Only show WHT if applicable
        if totals.get('wht_rate', 0) > 0:
            print_line(f"Less WHT ({totals['wht_rate']}%):", totals['wht'], color=colors.red)
            y -= 25

        # Decide whether there is enough space below to print payment and warranty
        required_space = 160  # approximate space needed for bank details + warranty
        new_page = False
        if y < required_space + 30:
            # Start a new page for the payment & warranty to avoid overlap
            self.c.showPage()
            new_page = True

        # Place Payment Details at bottom-left and Warranty at bottom-right
        # Use bottom coordinates (y from bottom); keep a small margin above page bottom
        y_bottom = 110
        y_bank = y_bottom
        y_warranty = y_bottom

        # Bank Details (left)
        self.c.setFillColor(colors.black)
        self.c.setFont("Helvetica-Bold", 10)
        self.c.drawString(30, y_bank, "PAYMENT DETAILS:")
        self.c.setFont("Helvetica", 9)
        self.c.drawString(30, y_bank - 15, f"Bank: {COMPANY_CONFIG['bank_name']}")
        self.c.drawString(30, y_bank - 28, f"Account Name: {COMPANY_CONFIG['account_name']}")
        self.c.drawString(30, y_bank - 41, f"Account Number: {COMPANY_CONFIG['account_number']}")

        # Delivery & Warranty Section (right-aligned)
        self.c.setFillColor(colors.HexColor("#0f3057"))
        self.c.setFont("Helvetica-Bold", 9)
        self.c.drawRightString(self.width - 30, y_warranty, "DELIVERY & WARRANTY:")
        
        self.c.setFont("Helvetica", 7.5)
        self.c.setFillColor(colors.black)
        warranty_text = [
            "Delivery: Orders dispatched within 48 hours of payment. Shipping notification will be sent.",
            "Replacements: Replacement guaranteed for defects reported within 48 hours.",
            "Note: No monetary refunds for faulty goods - direct item exchanges only."
        ]
        
        # Wrap warranty lines to a reasonable width and right-align them
        wrapped_lines = []
        for line in warranty_text:
            wrapped_lines.extend(textwrap.wrap(line, width=60))

        text_y = y_warranty - 12
        for line in wrapped_lines:
            self.c.drawRightString(self.width - 30, text_y, line)
            text_y -= 10

        # Footer note at bottom of page
        self.c.setFont("Helvetica-Oblique", 8)
        self.c.drawCentredString(self.width/2, 8, "Nascomsoft Embeded - Technology for all. Thank you for your patronage.")
        self.c.save()

# =============================================================================
# 4. GUI APP WITH TABS
# =============================================================================

class InvoiceApp(tb.Window):
    def __init__(self):
        super().__init__(themename="superhero")
        self.title("Nascomsoft Invoice Generator 2025")
        self.geometry("1400x900")
        self.resizable(True, True)
        
        self.db = DatabaseManager()
        self.current_tab = "component"  # Track current tab
        self.cart = []
        self.quote_cart = []
        # Dashboard pagination state
        self.dashboard_page = 1
        self.dashboard_page_size = 25
        
        self.setup_ui()
        self.refresh_invoice_number()
        # Initialize quotation number
        try:
            self.refresh_quote_number()
        except Exception:
            pass

    def setup_ui(self):
        # Header
        header = tb.Frame(self, bootstyle="secondary")
        header.pack(fill=X, padx=10, pady=10)
        tb.Label(header, text="NASCOMSOFT INVOICE SYSTEM", font=("Segoe UI", 20, "bold"), bootstyle="inverse-secondary").pack(pady=10)

        # Create Notebook (Tabs)
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=BOTH, expand=True, padx=20, pady=10)

        # Dashboard Tab (first)
        self.dashboard_frame = tb.Frame(self.notebook)
        self.notebook.add(self.dashboard_frame, text="Dashboard")
        self.setup_dashboard()

        # Component Tab
        self.component_frame = tb.Frame(self.notebook)
        self.notebook.add(self.component_frame, text="Component Invoice")
        self.setup_component_tab()

        # Project Tab
        self.project_frame = tb.Frame(self.notebook)
        self.notebook.add(self.project_frame, text="Project Invoice")
        self.setup_project_tab()

        # Quotation Tab
        self.quotation_frame = tb.Frame(self.notebook)
        self.notebook.add(self.quotation_frame, text="Quotation")
        self.setup_quotation()

        # Bind tab change event
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_changed)

        # Footer with buttons
        footer = tb.Frame(self, padding=20)
        footer.pack(fill=X)
        self.lbl_total = tb.Label(footer, text="Grand Total: N0.00", font=("Segoe UI", 16, "bold"), bootstyle="danger")
        self.lbl_total.pack(side=RIGHT, padx=20)
        tb.Button(footer, text="GENERATE INVOICE & SAVE", bootstyle="primary-outline", width=30, command=self.generate_invoice).pack(side=RIGHT)
        tb.Button(footer, text="Configure Email", bootstyle="secondary", command=self.configure_email_settings).pack(side=LEFT, padx=6)
        tb.Button(footer, text="SEND LAST FILE", bootstyle="info", command=self.send_last_file).pack(side=LEFT, padx=6)
        tb.Button(footer, text="Email Log", bootstyle="outline-info", command=self.show_email_log).pack(side=LEFT, padx=6)
        tb.Button(footer, text="Delete Selected", bootstyle="danger-outline", command=self.delete_selected_item).pack(side=LEFT, padx=10)
        tb.Button(footer, text="Clear List", bootstyle="secondary-link", command=self.clear_list).pack(side=LEFT)   

    def setup_project_tab(self):
        # Client Section
        details_frame = tb.Labelframe(self.project_frame, text="  Client Information  ", bootstyle="info", padding=20)
        details_frame.pack(fill=X, padx=20, pady=10)
        
        # Configure grid columns for consistent alignment
        details_frame.columnconfigure(1, minsize=200)
        details_frame.columnconfigure(3, minsize=300)
        
        # Row 0: Inv No & Name
        tb.Label(details_frame, text="Invoice No:", font=("Arial", 10)).grid(row=0, column=0, sticky=E, padx=10, pady=8)
        self.var_inv_no = tk.StringVar()
        tb.Entry(details_frame, textvariable=self.var_inv_no, state="readonly", width=18).grid(row=0, column=1, sticky=W, padx=10, pady=8)
        
        tb.Label(details_frame, text="Client Name:", font=("Arial", 10)).grid(row=0, column=2, sticky=E, padx=10, pady=8)
        self.var_client = tk.StringVar()
        tb.Entry(details_frame, textvariable=self.var_client, width=35).grid(row=0, column=3, sticky=W, padx=10, pady=8)

        tb.Label(details_frame, text="Client Email:", font=("Arial", 10)).grid(row=0, column=4, sticky=E, padx=10, pady=8)
        self.var_client_email = tk.StringVar()
        tb.Entry(details_frame, textvariable=self.var_client_email, width=30).grid(row=0, column=5, sticky=W, padx=10, pady=8)

        # Row 1: Address
        tb.Label(details_frame, text="Client Address:", font=("Arial", 10)).grid(row=1, column=0, sticky=NE, padx=10, pady=8)
        self.var_address = tk.Text(details_frame, height=2, width=35, wrap="word")
        self.var_address.grid(row=1, column=1, columnspan=5, sticky=W+N, padx=10, pady=8)

        # Auto-send toggle
        self.var_auto_send_invoice = tk.BooleanVar(value=False)
        tb.Checkbutton(details_frame, text="Send to client after generating", variable=self.var_auto_send_invoice).grid(row=2, column=2, columnspan=3, sticky=W, padx=10, pady=2) 

        # Row 2: WHT Rate (Projects Only)
        tb.Label(details_frame, text="WHT Rate (%):", font=("Arial", 10)).grid(row=2, column=0, sticky=E, padx=10, pady=8)
        self.var_wht = tk.DoubleVar(value=5.0)
        tb.Spinbox(details_frame, from_=0, to=10, increment=2.5, textvariable=self.var_wht, width=18).grid(row=2, column=1, sticky=W, padx=10, pady=8)
        tb.Label(details_frame, text="(Services: 5-10%)", font=("Arial", 8), bootstyle="info").grid(row=2, column=2, columnspan=2, sticky=W, padx=10, pady=8)

        # Row 3: Shipping Cost
        tb.Label(details_frame, text="Shipping Cost (N):", font=("Arial", 10)).grid(row=3, column=0, sticky=E, padx=10, pady=8)
        self.var_shipping = tk.DoubleVar(value=0.0)
        tb.Entry(details_frame, textvariable=self.var_shipping, width=18).grid(row=3, column=1, sticky=W, padx=10, pady=8)

        # Add Items Section
        item_frame = tb.Labelframe(self.project_frame, text="  Add Project Item  ", bootstyle="warning", padding=20)
        item_frame.pack(fill=X, padx=20, pady=10)
        
        # Standard columns (S/N is auto-generated)
        item_frame.columnconfigure(0, minsize=250)
        item_frame.columnconfigure(1, minsize=120)
        item_frame.columnconfigure(2, minsize=120)
        item_frame.columnconfigure(3, minsize=150)
        
        tb.Label(item_frame, text="Description:", font=("Arial", 10)).grid(row=0, column=0, sticky=W, padx=10, pady=8)
        self.var_project_desc = tk.StringVar()
        tb.Entry(item_frame, textvariable=self.var_project_desc, width=30).grid(row=1, column=0, sticky=W+E, padx=10, pady=8)
        
        tb.Label(item_frame, text="Qty / Fixed:", font=("Arial", 10)).grid(row=0, column=1, sticky=W, padx=10, pady=8)
        self.var_project_qty = tk.IntVar(value=1)
        tb.Spinbox(item_frame, from_=1, to=9999, textvariable=self.var_project_qty, width=15).grid(row=1, column=1, sticky=W+E, padx=10, pady=8)
        
        tb.Label(item_frame, text="Price (N):", font=("Arial", 10)).grid(row=0, column=2, sticky=W, padx=10, pady=8)
        self.var_project_price = tk.DoubleVar(value=0.0)
        tb.Entry(item_frame, textvariable=self.var_project_price, width=15).grid(row=1, column=2, sticky=W+E, padx=10, pady=8)
        
        tb.Button(item_frame, text="+ ADD ITEM", bootstyle="success", command=self.add_project_item).grid(row=1, column=3, sticky=W+E, padx=10, pady=8)  

        # Items List
        tree_frame = tb.Frame(self.project_frame)
        tree_frame.pack(fill=BOTH, expand=True, padx=20, pady=10)
        
        cols = ("sn", "desc", "qty", "price", "total")
        self.tree_project = ttk.Treeview(tree_frame, columns=cols, show="headings", height=8)
        self.tree_project.heading("sn", text="S/N")
        self.tree_project.heading("desc", text="Description")
        self.tree_project.heading("qty", text="Qty")
        self.tree_project.heading("price", text="Unit Price")
        self.tree_project.heading("total", text="Total")
        
        self.tree_project.column("sn", width=60, anchor=CENTER)
        self.tree_project.column("desc", width=420, anchor=W)
        self.tree_project.column("qty", width=80, anchor=CENTER)
        self.tree_project.column("price", width=120, anchor=E)
        self.tree_project.column("total", width=140, anchor=E)
        self.tree_project.pack(fill=BOTH, expand=True)

    def setup_component_tab(self):
        # Client Section
        details_frame = tb.Labelframe(self.component_frame, text="  Client Information  ", bootstyle="info", padding=20)
        details_frame.pack(fill=X, padx=20, pady=10)
        
        # Configure grid columns for consistent alignment
        details_frame.columnconfigure(1, minsize=200)
        details_frame.columnconfigure(3, minsize=300)
        
        # Row 0: Inv No & Name
        tb.Label(details_frame, text="Invoice No:", font=("Arial", 10)).grid(row=0, column=0, sticky=E, padx=10, pady=8)
        self.var_inv_no_comp = tk.StringVar()
        tb.Entry(details_frame, textvariable=self.var_inv_no_comp, state="readonly", width=18).grid(row=0, column=1, sticky=W, padx=10, pady=8)
        
        tb.Label(details_frame, text="Client Name:", font=("Arial", 10)).grid(row=0, column=2, sticky=E, padx=10, pady=8)
        self.var_client_comp = tk.StringVar()
        tb.Entry(details_frame, textvariable=self.var_client_comp, width=35).grid(row=0, column=3, sticky=W, padx=10, pady=8)

        tb.Label(details_frame, text="Client Email:", font=("Arial", 10)).grid(row=0, column=4, sticky=E, padx=10, pady=8)
        self.var_client_email_comp = tk.StringVar()
        tb.Entry(details_frame, textvariable=self.var_client_email_comp, width=30).grid(row=0, column=5, sticky=W, padx=10, pady=8)

        # Row 1: Address
        tb.Label(details_frame, text="Client Address:", font=("Arial", 10)).grid(row=1, column=0, sticky=NE, padx=10, pady=8)
        self.var_address_comp = tk.Text(details_frame, height=2, width=35, wrap="word")
        self.var_address_comp.grid(row=1, column=1, columnspan=5, sticky=W+N, padx=10, pady=8)

        # Auto-send toggle
        self.var_auto_send_invoice_comp = tk.BooleanVar(value=False)
        tb.Checkbutton(details_frame, text="Send to client after generating", variable=self.var_auto_send_invoice_comp).grid(row=2, column=2, columnspan=3, sticky=W, padx=10, pady=2) 

        # Row 2: Shipping Cost (Component)
        tb.Label(details_frame, text="Shipping Cost (N):", font=("Arial", 10)).grid(row=2, column=0, sticky=E, padx=10, pady=8)
        self.var_shipping_comp = tk.DoubleVar(value=0.0)
        tb.Entry(details_frame, textvariable=self.var_shipping_comp, width=18).grid(row=2, column=1, sticky=W, padx=10, pady=8)

        # Add Items Section (Components)
        item_frame = tb.Labelframe(self.component_frame, text="  Add Component Item  ", bootstyle="warning", padding=20)
        item_frame.pack(fill=X, padx=20, pady=10)
        
        item_frame.columnconfigure(0, minsize=250)
        item_frame.columnconfigure(1, minsize=120)
        item_frame.columnconfigure(2, minsize=120)
        item_frame.columnconfigure(3, minsize=150)
        
        tb.Label(item_frame, text="Component Name:", font=("Arial", 10)).grid(row=0, column=0, sticky=W, padx=10, pady=8)
        self.var_comp_desc = tk.StringVar()
        tb.Entry(item_frame, textvariable=self.var_comp_desc, width=30).grid(row=1, column=0, sticky=W+E, padx=10, pady=8)
        
        tb.Label(item_frame, text="Qty/Unit:", font=("Arial", 10)).grid(row=0, column=1, sticky=W, padx=10, pady=8)
        self.var_comp_qty = tk.IntVar(value=1)
        tb.Spinbox(item_frame, from_=1, to=9999, textvariable=self.var_comp_qty, width=15).grid(row=1, column=1, sticky=W+E, padx=10, pady=8)
        
        tb.Label(item_frame, text="Unit Price (N):", font=("Arial", 10)).grid(row=0, column=2, sticky=W, padx=10, pady=8)
        self.var_comp_price = tk.DoubleVar(value=0.0)
        tb.Entry(item_frame, textvariable=self.var_comp_price, width=15).grid(row=1, column=2, sticky=W+E, padx=10, pady=8)
        
        tb.Button(item_frame, text="+ ADD ITEM", bootstyle="success", command=self.add_component_item).grid(row=1, column=3, sticky=W+E, padx=10, pady=8)

        # Items List
        tree_frame = tb.Frame(self.component_frame)
        tree_frame.pack(fill=BOTH, expand=True, padx=20, pady=10)
        
        cols = ("sn", "desc", "qty", "price", "total")
        self.tree_comp = ttk.Treeview(tree_frame, columns=cols, show="headings", height=8)
        self.tree_comp.heading("sn", text="S/N")
        self.tree_comp.heading("desc", text="Component Name")
        self.tree_comp.heading("qty", text="Qty")
        self.tree_comp.heading("price", text="Unit Price")
        self.tree_comp.heading("total", text="Total")
        
        self.tree_comp.column("sn", width=60, anchor=CENTER)
        self.tree_comp.column("desc", width=420, anchor=W)
        self.tree_comp.column("qty", width=80, anchor=CENTER)
        self.tree_comp.column("price", width=120, anchor=E)
        self.tree_comp.column("total", width=140, anchor=E)
        self.tree_comp.pack(fill=BOTH, expand=True)

    def setup_quotation(self):
        # Quote Client Section
        details_frame = tb.Labelframe(self.quotation_frame, text="  Quotation Information  ", bootstyle="info", padding=20)
        details_frame.pack(fill=X, padx=20, pady=10)

        details_frame.columnconfigure(1, minsize=200)
        details_frame.columnconfigure(3, minsize=300)

        tb.Label(details_frame, text="Quote No:", font=("Arial", 10)).grid(row=0, column=0, sticky=E, padx=10, pady=8)
        self.var_quote_no = tk.StringVar()
        tb.Entry(details_frame, textvariable=self.var_quote_no, state="readonly", width=18).grid(row=0, column=1, sticky=W, padx=10, pady=8)

        tb.Label(details_frame, text="Client Name:", font=("Arial", 10)).grid(row=0, column=2, sticky=E, padx=10, pady=8)
        self.var_quote_client = tk.StringVar()
        tb.Entry(details_frame, textvariable=self.var_quote_client, width=35).grid(row=0, column=3, sticky=W, padx=10, pady=8)

        tb.Label(details_frame, text="Client Address:", font=("Arial", 10)).grid(row=1, column=0, sticky=NE, padx=10, pady=8)
        self.var_quote_address = tk.Text(details_frame, height=2, width=35, wrap="word")
        self.var_quote_address.grid(row=1, column=1, columnspan=3, sticky=W+N, padx=10, pady=8)

        tb.Label(details_frame, text="Shipping Cost (N):", font=("Arial", 10)).grid(row=2, column=0, sticky=E, padx=10, pady=8)
        self.var_quote_shipping = tk.DoubleVar(value=0.0)
        tb.Entry(details_frame, textvariable=self.var_quote_shipping, width=18).grid(row=2, column=1, sticky=W, padx=10, pady=8)

        # Add Items Section
        item_frame = tb.Labelframe(self.quotation_frame, text="  Add Quotation Item  ", bootstyle="warning", padding=20)
        item_frame.pack(fill=X, padx=20, pady=10)

        item_frame.columnconfigure(0, minsize=250)
        item_frame.columnconfigure(1, minsize=120)
        item_frame.columnconfigure(2, minsize=120)
        item_frame.columnconfigure(3, minsize=150)

        tb.Label(item_frame, text="Description:", font=("Arial", 10)).grid(row=0, column=0, sticky=W, padx=10, pady=8)
        self.var_quote_desc = tk.StringVar()
        tb.Entry(item_frame, textvariable=self.var_quote_desc, width=30).grid(row=1, column=0, sticky=W+E, padx=10, pady=8)

        tb.Label(item_frame, text="Qty:", font=("Arial", 10)).grid(row=0, column=1, sticky=W, padx=10, pady=8)
        self.var_quote_qty = tk.IntVar(value=1)
        tb.Spinbox(item_frame, from_=1, to=9999, textvariable=self.var_quote_qty, width=15).grid(row=1, column=1, sticky=W+E, padx=10, pady=8)

        tb.Label(item_frame, text="Unit Price (N):", font=("Arial", 10)).grid(row=0, column=2, sticky=W, padx=10, pady=8)
        self.var_quote_price = tk.DoubleVar(value=0.0)
        tb.Entry(item_frame, textvariable=self.var_quote_price, width=15).grid(row=1, column=2, sticky=W+E, padx=10, pady=8)

        tb.Button(item_frame, text="+ ADD ITEM", bootstyle="success", command=self.add_quote_item).grid(row=1, column=3, sticky=W+E, padx=10, pady=8)

        # Items List
        tree_frame = tb.Frame(self.quotation_frame)
        tree_frame.pack(fill=BOTH, expand=True, padx=20, pady=10)

        cols = ("sn", "desc", "qty", "price", "total")
        self.tree_quote = ttk.Treeview(tree_frame, columns=cols, show="headings", height=8)
        self.tree_quote.heading("sn", text="S/N")
        self.tree_quote.heading("desc", text="Description")
        self.tree_quote.heading("qty", text="Qty")
        self.tree_quote.heading("price", text="Unit Price")
        self.tree_quote.heading("total", text="Total")

        self.tree_quote.column("sn", width=60, anchor=CENTER)
        self.tree_quote.column("desc", width=420, anchor=W)
        self.tree_quote.column("qty", width=80, anchor=CENTER)
        self.tree_quote.column("price", width=120, anchor=E)
        self.tree_quote.column("total", width=140, anchor=E)
        self.tree_quote.pack(fill=BOTH, expand=True)

        # Quotation controls
        qfooter = tb.Frame(self.quotation_frame, padding=10)
        qfooter.pack(fill=X)
        self.lbl_quote_total = tb.Label(qfooter, text="Quote Total: N0.00", font=("Segoe UI", 12, "bold"), bootstyle="info")
        self.lbl_quote_total.pack(side=RIGHT, padx=10)
        tb.Button(qfooter, text="GENERATE QUOTATION & SAVE", bootstyle="primary-outline", command=self.generate_quotation).pack(side=RIGHT, padx=10)
        tb.Button(qfooter, text="SEND TO CLIENT", bootstyle="info-outline", command=self.send_quote_file).pack(side=RIGHT, padx=6)
        tb.Button(qfooter, text="Configure Email", bootstyle="secondary", command=self.configure_email_settings).pack(side=LEFT, padx=6)
        tb.Button(qfooter, text="Clear Quote", bootstyle="secondary-link", command=self.clear_quote).pack(side=LEFT) 

    def add_quote_item(self):
        desc = self.var_quote_desc.get().strip()
        price = self.var_quote_price.get()
        qty = self.var_quote_qty.get()
        
        if not desc or price <= 0:
            messagebox.showwarning("Error", "Check inputs.")
            return

        sn = str(len(self.quote_cart) + 1)
        total = price * qty
        self.quote_cart.append({"sn": sn, "desc": desc, "type": "Quotation", "qty": qty, "price": price, "total": total})
        self.tree_quote.insert("", "end", values=(sn, desc, qty, f"{price:,.2f}", f"{total:,.2f}"))
        self.calculate_quote_totals()

        self.var_quote_desc.set("")
        self.var_quote_price.set(0.0)
        self.var_quote_qty.set(1)

    def calculate_quote_totals(self):
        subtotal = sum(item['total'] for item in self.quote_cart)
        vat = subtotal * COMPANY_CONFIG['vat_rate']
        shipping = self.var_quote_shipping.get()
        grand_total = subtotal + vat + shipping
        self.lbl_quote_total.config(text=f"Quote Total: N{grand_total:,.2f}")
        return subtotal, vat, shipping, grand_total

    def clear_quote(self):
        self.quote_cart = []
        self.tree_quote.delete(*self.tree_quote.get_children())
        self.calculate_quote_totals()

    def refresh_quote_number(self):
        try:
            new_no = self.db.generate_quotation_number()
        except Exception as e:
            ts = datetime.now().strftime("%Y%m%d%H%M%S")
            new_no = f"NSE-QTN-FALLBACK-{ts}"
        self.var_quote_no.set(new_no)

    def generate_quotation(self):
        if not self.quote_cart:
            messagebox.showerror("Error", "Quotation is empty.")
            return

        client_name = self.var_quote_client.get().strip()
        client_addr = self.var_quote_address.get("1.0", tk.END).strip()
        if not client_name:
            messagebox.showerror("Error", "Client Name is required.")
            return

        subtotal, vat, shipping, grand_total = self.calculate_quote_totals()
        quote_no = self.var_quote_no.get()

        quote_data = {
            "quote_no": quote_no,
            "client_name": client_name,
            "client_email": self.var_quote_email.get().strip(),
            "client_address": client_addr,
            "subtotal": subtotal,
            "vat": vat,
            "shipping": shipping,
            "grand_total": grand_total
        }

        if self.db.save_quotation(quote_data):
            try:
                filename = f"Quotation_{quote_data['quote_no']}.pdf"
                pdf = InvoicePDF(filename)
                pdf.draw_header(quote_data['quote_no'], datetime.now().strftime("%d-%b-%Y"), doc_type="QUOTATION")
                pdf.draw_client_info(quote_data['client_name'], quote_data['client_address'])
                pdf.draw_items_table(self.quote_cart)
                pdf.draw_footer(quote_data)

                # remember last file for optional sending
                self.last_generated_file = filename

                messagebox.showinfo("Success", f"Quotation Saved!\nFilename: {filename}")
                try:
                    if os.name == 'nt':
                        os.startfile(filename)
                    else:
                        subprocess.run(['open' if sys.platform == 'darwin' else 'xdg-open', filename], check=False)
                except Exception as e:
                    messagebox.showerror("File Open Error", f"Could not open the file: {e}")

                # Auto-send if requested
                if self.var_auto_send_quote.get():
                    to_email = quote_data.get('client_email', '')
                    if to_email and self.is_valid_email(to_email):
                        subj = f"Quotation {quote_data['quote_no']}"
                        body = f"Please find attached the quotation {quote_data['quote_no']}."
                        success, err = self.send_email(to_email, subj, body, filename)
                        if success:
                            messagebox.showinfo("Email Sent", f"Quotation sent to {to_email}")
                        else:
                            messagebox.showerror("Email Error", f"Failed to send email: {err}")
                    else:
                        messagebox.showwarning("Email", "No valid client email provided.")

                self.clear_quote()
                self.var_quote_client.set("")
                self.var_quote_address.delete("1.0", tk.END)
                self.var_quote_email.set("")
                self.refresh_quote_number()
            except Exception as e:
                messagebox.showerror("PDF Error", f"An error occurred while generating the PDF: {e}")



    def setup_dashboard(self):
        # Dashboard Search / Controls
        top = tb.Frame(self.dashboard_frame, padding=10)
        top.pack(fill=X, padx=10, pady=8)

        tb.Label(top, text="Invoice #:", font=("Arial", 10)).grid(row=0, column=0, sticky=E, padx=8)
        self.var_dash_inv = tk.StringVar()
        tb.Entry(top, textvariable=self.var_dash_inv, width=18).grid(row=0, column=1, sticky=W, padx=8)

        tb.Label(top, text="Client:", font=("Arial", 10)).grid(row=0, column=2, sticky=E, padx=8)
        self.var_dash_client = tk.StringVar()
        tb.Entry(top, textvariable=self.var_dash_client, width=22).grid(row=0, column=3, sticky=W, padx=8)

        tb.Label(top, text="Type:", font=("Arial", 10)).grid(row=0, column=4, sticky=E, padx=8)
        self.var_dash_type = tk.StringVar(value="All")
        tb.Combobox(top, values=["All", "Project", "Component", "Quotation"], textvariable=self.var_dash_type, width=14, state="readonly").grid(row=0, column=5, sticky=W, padx=8)

        tb.Button(top, text="Search", bootstyle="primary", command=self.on_dashboard_search).grid(row=0, column=6, sticky=W, padx=6)
        tb.Button(top, text="Refresh", bootstyle="secondary", command=lambda: self.load_dashboard_data(1)).grid(row=0, column=7, sticky=W, padx=6)

        # Actions
        actions = tb.Frame(self.dashboard_frame, padding=8)
        actions.pack(fill=X, padx=10)
        tb.Button(actions, text="Open PDF", bootstyle="info-outline", command=self.open_selected_invoice_pdf).pack(side=LEFT, padx=6)
        tb.Button(actions, text="Export CSV", bootstyle="success-outline", command=self.export_dashboard_csv).pack(side=LEFT, padx=6)
        tb.Button(actions, text="Delete Invoice", bootstyle="danger-outline", command=self.delete_selected_invoice).pack(side=LEFT, padx=6)

        # Treeview
        tree_frame = tb.Frame(self.dashboard_frame)
        tree_frame.pack(fill=BOTH, expand=True, padx=10, pady=6)

        cols = ("invoice_no", "date", "client", "type", "subtotal", "vat", "shipping", "wht", "grand_total")
        self.dashboard_tree = ttk.Treeview(tree_frame, columns=cols, show="headings", height=18)
        headings = ["Invoice #", "Date", "Client", "Type", "Subtotal", "VAT", "Shipping", "WHT", "Grand Total"]
        widths = [120, 140, 260, 80, 90, 90, 90, 80, 110]
        for col, h, w in zip(cols, headings, widths):
            self.dashboard_tree.heading(col, text=h)
            self.dashboard_tree.column(col, width=w, anchor=W)
        # Adjust numeric alignment
        self.dashboard_tree.column("subtotal", anchor=E)
        self.dashboard_tree.column("vat", anchor=E)
        self.dashboard_tree.column("shipping", anchor=E)
        self.dashboard_tree.column("wht", anchor=E)
        self.dashboard_tree.column("grand_total", anchor=E)

        # Add striped rows tags
        self.dashboard_tree.tag_configure('oddrow', background='#f6f8fa')
        self.dashboard_tree.tag_configure('evenrow', background='white')

        self.dashboard_tree.pack(fill=BOTH, expand=True)
        # Double-click to open
        self.dashboard_tree.bind("<Double-1>", lambda e: self.open_selected_invoice_pdf())

        # Pagination controls
        footer = tb.Frame(self.dashboard_frame, padding=8)
        footer.pack(fill=X, padx=10, pady=6)
        tb.Button(footer, text="Prev", bootstyle="secondary", command=self.prev_dashboard_page).pack(side=LEFT, padx=6)
        self.lbl_dash_page = tb.Label(footer, text=f"Page {self.dashboard_page}")
        self.lbl_dash_page.pack(side=LEFT, padx=6)
        tb.Button(footer, text="Next", bootstyle="secondary", command=self.next_dashboard_page).pack(side=LEFT, padx=6)

        # Load initial data
        self.load_dashboard_data(self.dashboard_page)

    def load_dashboard_data(self, page=1):
        self.dashboard_page = page
        filters = {
            'invoice_no': self.var_dash_inv.get().strip() if hasattr(self, 'var_dash_inv') else '',
            'client_name': self.var_dash_client.get().strip() if hasattr(self, 'var_dash_client') else '',
            'invoice_type': self.var_dash_type.get().strip() if hasattr(self, 'var_dash_type') else 'All'
        }
        # Note: date_from/date_to not implemented yet
        # Decide whether to load invoices, quotations or both
        doc_type = filters.get('invoice_type', 'All')
        rows = []
        if doc_type in (None, 'All', 'Project', 'Component'):
            invoices = self.db.fetch_invoices(filters=filters, page=page, page_size=self.dashboard_page_size)
            rows.extend(invoices)
        if doc_type in (None, 'All', 'Quotation'):
            quotes = self.db.fetch_quotations(filters=filters, page=page, page_size=self.dashboard_page_size)
            rows.extend(quotes)

        # Clear tree
        for r in self.dashboard_tree.get_children():
            self.dashboard_tree.delete(r)
        # Insert rows
        for idx, inv in enumerate(rows):
            tag = 'evenrow' if idx % 2 == 0 else 'oddrow'
            self.dashboard_tree.insert('', 'end', values=(inv['invoice_no'], inv['date_issued'], inv['client_name'], inv['invoice_type'], f"{COMPANY_CONFIG['currency_symbol']}{inv['subtotal']:,.2f}", f"{COMPANY_CONFIG['currency_symbol']}{inv['vat']:,.2f}", f"{COMPANY_CONFIG['currency_symbol']}{inv['shipping']:,.2f}", f"{COMPANY_CONFIG['currency_symbol']}{inv['wht']:,.2f}", f"{COMPANY_CONFIG['currency_symbol']}{inv['grand_total']:,.2f}"), tags=(tag,))
        self.lbl_dash_page.config(text=f"Page {self.dashboard_page}")

    def on_dashboard_search(self):
        self.load_dashboard_data(1)

    def open_selected_invoice_pdf(self):
        sel = self.dashboard_tree.selection()
        if not sel:
            messagebox.showwarning("No selection", "Select an invoice or quotation to open its PDF.")
            return
        row = self.dashboard_tree.item(sel[0], 'values')
        inv_no = row[0]
        inv_type = row[3] if len(row) > 3 else 'Project'
        if inv_type == 'Quotation':
            filename = f"Quotation_{inv_no}.pdf"
        else:
            filename = f"Invoice_{inv_no}.pdf"
        if os.path.exists(filename):
            try:
                if os.name == 'nt':
                    os.startfile(filename)
                else:
                    subprocess.run(['open' if sys.platform == 'darwin' else 'xdg-open', filename], check=False)
            except Exception as e:
                messagebox.showerror("Open Error", f"Could not open PDF: {e}")
        else:
            messagebox.showwarning("Not found", f"PDF {filename} not found on disk.")

    def export_dashboard_csv(self):
        rows = [self.dashboard_tree.item(i, 'values') for i in self.dashboard_tree.get_children()]
        if not rows:
            messagebox.showinfo("No Data", "No rows to export.")
            return
        path = filedialog.asksaveasfilename(defaultextension='.csv', filetypes=[('CSV files','*.csv')], initialfile='invoices_export.csv')
        if not path:
            return
        try:
            with open(path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow([h for h in ["Invoice #","Date","Client","Type","Subtotal","VAT","Shipping","WHT","Grand Total"]])
                for row in rows:
                    writer.writerow(row)
            messagebox.showinfo("Exported", f"Exported {len(rows)} rows to {path}")
        except Exception as e:
            messagebox.showerror("Export Error", f"Could not export CSV: {e}")

    def delete_selected_invoice(self):
        sel = self.dashboard_tree.selection()
        if not sel:
            messagebox.showwarning("No selection", "Select an invoice to delete.")
            return
        row = self.dashboard_tree.item(sel[0], 'values')
        inv_no = row[0]
        inv_type = row[3] if len(row) > 3 else 'Project'
        if not messagebox.askyesno("Confirm Delete", f"Delete {inv_type} {inv_no} from database?"):
            return
        success = False
        if inv_type == 'Quotation':
            success = self.db.delete_quotation(inv_no)
            filename = f"Quotation_{inv_no}.pdf"
        else:
            success = self.db.delete_invoice(inv_no)
            filename = f"Invoice_{inv_no}.pdf"

        if success:
            # Attempt to delete the PDF file if present
            try:
                if os.path.exists(filename):
                    os.remove(filename)
            except Exception:
                pass
            messagebox.showinfo("Deleted", f"{inv_type} {inv_no} deleted.")
            self.load_dashboard_data(self.dashboard_page)
        else:
            messagebox.showerror("Delete Error", f"Could not delete {inv_type} from DB.")

    def prev_dashboard_page(self):
        if self.dashboard_page > 1:
            self.load_dashboard_data(self.dashboard_page - 1)

    def next_dashboard_page(self):
        # Simple next page - user can click to paginate
        self.load_dashboard_data(self.dashboard_page + 1)

    def on_tab_changed(self, event):
        selected = self.notebook.select()
        try:
            tab_text = self.notebook.tab(selected, "text")
        except Exception:
            tab_text = ""
        tab_text = tab_text.lower() if isinstance(tab_text, str) else ""
        if "component" in tab_text:
            self.current_tab = "component"
        elif "project" in tab_text:
            self.current_tab = "project"
        elif "quotation" in tab_text:
            self.current_tab = "quotation"
        elif "dashboard" in tab_text:
            self.current_tab = "dashboard"
        else:
            self.current_tab = "component"  # fallback

    def refresh_invoice_number(self):
        try:
            new_no = self.db.generate_invoice_number()
        except Exception as e:
            ts = datetime.now().strftime("%Y%m%d%H%M%S")
            new_no = f"NSE-INV-FALLBACK-{ts}"
        self.var_inv_no.set(new_no)
        self.var_inv_no_comp.set(new_no)

    def add_project_item(self):
        desc = self.var_project_desc.get().strip()
        price = self.var_project_price.get()
        qty = self.var_project_qty.get()
        
        if not desc or price <= 0:
            messagebox.showwarning("Error", "Check inputs.")
            return

        # S/N is auto-generated per project item
        sn = str(len([i for i in self.cart if i['type'] == 'Project']) + 1)

        total = price * qty
        self.cart.append({"sn": sn, "desc": desc, "type": "Project", "qty": qty, "price": price, "total": total})
        self.tree_project.insert("", "end", values=(sn, desc, qty, f"{price:,.2f}", f"{total:,.2f}"))
        self.calculate_totals()
        
        self.var_project_desc.set("")
        self.var_project_price.set(0.0)
        self.var_project_qty.set(1)

    def add_component_item(self):
        desc = self.var_comp_desc.get().strip()
        price = self.var_comp_price.get()
        qty = self.var_comp_qty.get()
        
        if not desc or price <= 0:
            messagebox.showwarning("Error", "Check inputs.")
            return

        # S/N is auto-generated per component item
        sn = str(len([i for i in self.cart if i['type'] == 'Component']) + 1)

        total = price * qty
        self.cart.append({"sn": sn, "desc": desc, "type": "Component", "qty": qty, "price": price, "total": total})
        self.tree_comp.insert("", "end", values=(sn, desc, qty, f"{price:,.2f}", f"{total:,.2f}"))
        self.calculate_totals()
        
        self.var_comp_desc.set("")
        self.var_comp_price.set(0.0)
        self.var_comp_qty.set(1)

    def calculate_totals(self):
        subtotal = sum(item['total'] for item in self.cart)
        vat = subtotal * COMPANY_CONFIG['vat_rate']
        
        # Get shipping based on current tab
        if self.current_tab == "project":
            shipping = self.var_shipping.get()
        else:
            shipping = self.var_shipping_comp.get()
        
        grand_total = subtotal + vat + shipping
        self.lbl_total.config(text=f"Total: N{grand_total:,.2f}")
        return subtotal, vat, shipping, grand_total

    def clear_list(self):
        self.cart = []
        self.tree_project.delete(*self.tree_project.get_children())
        self.tree_comp.delete(*self.tree_comp.get_children())
        self.calculate_totals()

    def delete_selected_item(self):
        # Determine active tree based on tab
        if self.current_tab == "project":
            tree = self.tree_project
            item_type = "Project"
        else:
            tree = self.tree_comp
            item_type = "Component"

        selected = tree.selection()
        if not selected:
            messagebox.showwarning("No selection", "Select an item to delete.")
            return

        if not messagebox.askyesno("Confirm Delete", "Delete selected item(s)?"):
            return

        # Remove selected tree items and matching cart entries
        for sel in selected:
            vals = tree.item(sel, "values")
            sn = str(vals[0]) if vals else None
            tree.delete(sel)
            # remove first matching cart entry by type and sn
            for i, it in enumerate(self.cart):
                if it.get('type') == item_type and str(it.get('sn')) == sn:
                    del self.cart[i]
                    break

        # Reindex S/N for remaining items of this type in cart
        cart_indices = [i for i, it in enumerate(self.cart) if it.get('type') == item_type]
        for idx, cart_idx in enumerate(cart_indices):
            new_sn = str(idx + 1)
            self.cart[cart_idx]['sn'] = new_sn

        # Update tree view S/N column to match new numbering
        children = tree.get_children()
        for idx, child in enumerate(children):
            vals = list(tree.item(child, "values"))
            vals[0] = str(idx + 1)
            tree.item(child, values=tuple(vals))

        self.calculate_totals()

    # ------------------- Email / SMTP helpers -------------------
    def is_valid_email(self, email):
        return bool(email and "@" in email and "." in email)

    def send_email(self, to_address, subject, body, attachment_path):
        """Send an email with the given attachment. Returns (True, '') on success, (False, error_message) on failure."""
        if not SMTP_SETTINGS.get('host'):
            return False, "SMTP is not configured. Please configure email settings first."
        try:
            msg = EmailMessage()
            msg['Subject'] = subject
            msg['From'] = SMTP_SETTINGS.get('from_email') or SMTP_SETTINGS.get('username')
            msg['To'] = to_address
            msg.set_content(body)

            # attach file
            if attachment_path and os.path.exists(attachment_path):
                with open(attachment_path, 'rb') as f:
                    data = f.read()
                ctype, encoding = mimetypes.guess_type(attachment_path)
                if ctype:
                    maintype, subtype = ctype.split('/', 1)
                else:
                    maintype, subtype = 'application', 'octet-stream'
                msg.add_attachment(data, maintype=maintype, subtype=subtype, filename=os.path.basename(attachment_path))

            context = ssl.create_default_context()
            with smtplib.SMTP(SMTP_SETTINGS.get('host'), SMTP_SETTINGS.get('port')) as server:
                if SMTP_SETTINGS.get('use_tls'):
                    server.starttls(context=context)
                if SMTP_SETTINGS.get('username'):
                    server.login(SMTP_SETTINGS.get('username'), SMTP_SETTINGS.get('password'))
                server.send_message(msg)
            # Log success in DB if available
            try:
                if hasattr(self, 'db') and self.db:
                    try:
                        self.db.save_email_log(to_address, subject, attachment_path, 'SENT', '')
                    except Exception:
                        pass
            except Exception:
                pass
            return True, ''
        except Exception as e:
            # Log failure in DB if available
            try:
                if hasattr(self, 'db') and self.db:
                    try:
                        self.db.save_email_log(to_address, subject, attachment_path, 'FAILED', str(e))
                    except Exception:
                        pass
            except Exception:
                pass
            return False, str(e) 

    def configure_email_settings(self):
        # Simple dialog to configure SMTP settings
        dlg = tk.Toplevel(self)
        dlg.title("Configure Email Settings")
        dlg.transient(self)
        dlg.grab_set()

        tk.Label(dlg, text="SMTP Host:").grid(row=0, column=0, sticky=E, padx=6, pady=6)
        host_var = tk.StringVar(value=SMTP_SETTINGS.get('host', ''))
        tk.Entry(dlg, textvariable=host_var, width=30).grid(row=0, column=1, padx=6, pady=6)

        tk.Label(dlg, text="Port:").grid(row=1, column=0, sticky=E, padx=6, pady=6)
        port_var = tk.IntVar(value=SMTP_SETTINGS.get('port', 587))
        tk.Entry(dlg, textvariable=port_var, width=10).grid(row=1, column=1, padx=6, pady=6, sticky=W)

        tk.Label(dlg, text="Username:").grid(row=2, column=0, sticky=E, padx=6, pady=6)
        user_var = tk.StringVar(value=SMTP_SETTINGS.get('username', ''))
        tk.Entry(dlg, textvariable=user_var, width=30).grid(row=2, column=1, padx=6, pady=6)

        tk.Label(dlg, text="Password:").grid(row=3, column=0, sticky=E, padx=6, pady=6)
        pass_var = tk.StringVar(value=SMTP_SETTINGS.get('password', ''))
        tk.Entry(dlg, textvariable=pass_var, width=30, show='*').grid(row=3, column=1, padx=6, pady=6)

        use_tls_var = tk.BooleanVar(value=SMTP_SETTINGS.get('use_tls', True))
        tk.Checkbutton(dlg, text="Use TLS", variable=use_tls_var).grid(row=4, column=1, sticky=W, padx=6, pady=6)

        tk.Label(dlg, text="From Email:").grid(row=5, column=0, sticky=E, padx=6, pady=6)
        from_var = tk.StringVar(value=SMTP_SETTINGS.get('from_email', ''))
        tk.Entry(dlg, textvariable=from_var, width=30).grid(row=5, column=1, padx=6, pady=6)

        def save_settings():
            SMTP_SETTINGS['host'] = host_var.get().strip()
            SMTP_SETTINGS['port'] = int(port_var.get())
            SMTP_SETTINGS['username'] = user_var.get().strip()
            SMTP_SETTINGS['password'] = pass_var.get()
            SMTP_SETTINGS['use_tls'] = use_tls_var.get()
            SMTP_SETTINGS['from_email'] = from_var.get().strip()
            dlg.destroy()
            messagebox.showinfo("Saved", "SMTP settings saved (in memory).")

        def send_test():
            tmp_to = user_var.get().strip() or from_var.get().strip()
            if not tmp_to:
                messagebox.showwarning("Test Email", "Provide a recipient or username to test.")
                return
            SMTP_SETTINGS['host'] = host_var.get().strip()
            SMTP_SETTINGS['port'] = int(port_var.get())
            SMTP_SETTINGS['username'] = user_var.get().strip()
            SMTP_SETTINGS['password'] = pass_var.get()
            SMTP_SETTINGS['use_tls'] = use_tls_var.get()
            SMTP_SETTINGS['from_email'] = from_var.get().strip()
            success, err = self.send_email(tmp_to, "Test Email from Nascomsoft", "This is a test email.", None)
            if success:
                messagebox.showinfo("Success", "Test email sent successfully.")
            else:
                messagebox.showerror("Error", f"Test email failed: {err}")

        tk.Button(dlg, text="Save", command=save_settings).grid(row=6, column=0, padx=6, pady=8)
        tk.Button(dlg, text="Send Test", command=send_test).grid(row=6, column=1, padx=6, pady=8, sticky=W)

    def send_last_file(self):
        if not getattr(self, 'last_generated_file', None):
            messagebox.showwarning("No file", "No generated file to send.")
            return
        # try to get current tab's client email, fallback to asking
        if self.current_tab == 'project':
            to_email = self.var_client_email.get().strip()
        elif self.current_tab == 'component':
            to_email = self.var_client_email_comp.get().strip()
        else:
            to_email = ''
        if not to_email:
            messagebox.showwarning("No recipient", "No client email available for the current tab.")
            return
        if not self.is_valid_email(to_email):
            messagebox.showwarning("Invalid Email", "The client email address appears invalid.")
            return
        subj = f"Document {os.path.basename(self.last_generated_file)}"
        body = "Please find attached the requested document."
        success, err = self.send_email(to_email, subj, body, self.last_generated_file)
        if success:
            messagebox.showinfo("Email Sent", f"File sent to {to_email}")
        else:
            messagebox.showerror("Email Error", f"Failed to send email: {err}")

    def send_quote_file(self):
        if not getattr(self, 'last_generated_file', None):
            messagebox.showwarning("No file", "No generated file to send.")
            return
        to_email = self.var_quote_email.get().strip()
        if not to_email or not self.is_valid_email(to_email):
            messagebox.showwarning("Invalid Email", "Provide a valid client email to send the quotation.")
            return
        subj = f"Quotation {self.var_quote_no.get()}"
        body = f"Please find attached quotation {self.var_quote_no.get()}"
        success, err = self.send_email(to_email, subj, body, self.last_generated_file)
        if success:
            messagebox.showinfo("Email Sent", f"Quotation sent to {to_email}")
        else:
            messagebox.showerror("Email Error", f"Failed to send email: {err}")

    def show_email_log(self):
        dlg = tk.Toplevel(self)
        dlg.title("Email Delivery Log")
        dlg.geometry("900x400")
        dlg.transient(self)
        dlg.grab_set()

        frame = tb.Frame(dlg, padding=10)
        frame.pack(fill=BOTH, expand=True)

        cols = ("id", "created_at", "to", "subject", "attachment", "status", "error")
        tree = ttk.Treeview(frame, columns=cols, show='headings')
        for c, title in zip(cols, ["ID", "Time", "To", "Subject", "Attachment", "Status", "Error"]):
            tree.heading(c, text=title)
            tree.column(c, width=120 if c in ("subject","error") else 100, anchor=W)
        tree.pack(fill=BOTH, expand=True, padx=6, pady=6)

        # Load logs
        logs = []
        try:
            if hasattr(self, 'db') and self.db:
                logs = self.db.fetch_email_logs(500)
        except Exception:
            logs = []

        for l in logs:
            tree.insert('', 'end', values=(l['id'], l['created_at'], l['to_address'], l['subject'], l['attachment'], l['status'], l['error_message']))

        btn_frame = tb.Frame(dlg, padding=6)
        btn_frame.pack(fill=X)
        def refresh():
            for it in tree.get_children():
                tree.delete(it)
            try:
                logs = self.db.fetch_email_logs(500)
            except Exception:
                logs = []
            for l in logs:
                tree.insert('', 'end', values=(l['id'], l['created_at'], l['to_address'], l['subject'], l['attachment'], l['status'], l['error_message']))

        def export_csv():
            path = filedialog.asksaveasfilename(defaultextension='.csv', filetypes=[('CSV files','*.csv')], title='Export Email Log')
            if not path:
                return
            try:
                with open(path, 'w', newline='', encoding='utf-8') as f:
                    writer = csv.writer(f)
                    writer.writerow(["ID","Time","To","Subject","Attachment","Status","Error"])
                    for l in (self.db.fetch_email_logs(10000) if hasattr(self,'db') and self.db else logs):
                        writer.writerow([l['id'], l['created_at'], l['to_address'], l['subject'], l['attachment'], l['status'], l['error_message']])
                messagebox.showinfo("Exported", f"Email log exported to {path}")
            except Exception as e:
                messagebox.showerror("Export Error", f"Failed to export: {e}")

        tb.Button(btn_frame, text="Refresh", command=refresh, bootstyle='secondary').pack(side=LEFT, padx=6)
        tb.Button(btn_frame, text="Export CSV", command=export_csv, bootstyle='success').pack(side=LEFT, padx=6)
        tb.Button(btn_frame, text="Close", command=dlg.destroy, bootstyle='danger-outline').pack(side=RIGHT, padx=6)


    def generate_invoice(self):
        if not self.cart:
            messagebox.showerror("Error", "Invoice is empty.")
            return
        
        # Get data based on current tab
        if self.current_tab == "project":
            client_name = self.var_client.get().strip()
            client_addr = self.var_address.get("1.0", tk.END).strip()
            wht_rate = self.var_wht.get()
            invoice_type = "Project"
        else:
            client_name = self.var_client_comp.get().strip()
            client_addr = self.var_address_comp.get("1.0", tk.END).strip()
            wht_rate = 0  # No WHT for components
            invoice_type = "Component"
        
        # Validate client name from the active tab only
        if not client_name:
            messagebox.showerror("Error", "Client Name is required.")
            return

        subtotal, vat, shipping, grand_total = self.calculate_totals()
        wht_amount = grand_total * (wht_rate / 100)
        
        # Use the invoice number field for the active tab
        invoice_no = self.var_inv_no.get() if self.current_tab == "project" else self.var_inv_no_comp.get()

        # capture client email depending on tab
        client_email = self.var_client_email.get().strip() if self.current_tab == "project" else self.var_client_email_comp.get().strip()

        invoice_data = {
            "invoice_no": invoice_no,
            "client_name": client_name,
            "client_email": client_email,
            "client_address": client_addr,
            "invoice_type": invoice_type,
            "subtotal": subtotal,
            "vat": vat,
            "shipping": shipping,
            "grand_total": grand_total,
            "wht_rate": wht_rate,
            "wht": wht_amount
        }

        if self.db.save_invoice(invoice_data):
            try:
                filename = f"Invoice_{invoice_data['invoice_no']}.pdf"
                pdf = InvoicePDF(filename)
                pdf.draw_header(invoice_data['invoice_no'], datetime.now().strftime("%d-%b-%Y"))
                pdf.draw_client_info(invoice_data['client_name'], invoice_data['client_address'])
                pdf.draw_items_table(self.cart)
                pdf.draw_footer(invoice_data)
                
                # remember last generated file
                self.last_generated_file = filename

                messagebox.showinfo("Success", f"Invoice Saved!\nFilename: {filename}")
                try:
                    if os.name == 'nt':
                        os.startfile(filename)
                    else:
                        subprocess.run(['open' if sys.platform == 'darwin' else 'xdg-open', filename], check=False)
                except Exception as e:
                    messagebox.showerror("File Open Error", f"Could not open the file: {e}")
                
                # Auto-send if requested
                send_flag = self.var_auto_send_invoice.get() if self.current_tab == "project" else self.var_auto_send_invoice_comp.get()
                if send_flag and client_email:
                    if self.is_valid_email(client_email):
                        subj = f"Invoice {invoice_data['invoice_no']}"
                        body = f"Please find attached the invoice {invoice_data['invoice_no']}"
                        success, err = self.send_email(client_email, subj, body, filename)
                        if success:
                            messagebox.showinfo("Email Sent", f"Invoice sent to {client_email}")
                        else:
                            messagebox.showerror("Email Error", f"Failed to send email: {err}")
                    else:
                        messagebox.showwarning("Email", "No valid client email provided.")

                self.clear_list()
                if self.current_tab == "project":
                    self.var_client.set("")
                    self.var_address.delete("1.0", tk.END)
                    self.var_client_email.set("")
                else:
                    self.var_client_comp.set("")
                    self.var_address_comp.delete("1.0", tk.END)
                    self.var_client_email_comp.set("")
                self.refresh_invoice_number()
            except Exception as e:
                messagebox.showerror("PDF Error", f"An error occurred while generating the PDF: {e}")

if __name__ == "__main__":
    app = InvoiceApp()
    app.mainloop()
