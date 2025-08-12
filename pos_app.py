from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file
from openpyxl import load_workbook, Workbook
from datetime import datetime
import os
import shutil
from docx import Document

app = Flask(__name__, template_folder='templates')
app.secret_key = 'change_this_to_a_secure_random_value'  # change for production

@app.context_processor
def inject_datetime():
    return {'datetime': datetime}

# === Files ===
USER_FILE = 'users.xlsx'
PRODUCT_FILE = 'products.xlsx'
SALES_FILE = 'sales_log.xlsx'
OIL_FILE = 'oils.xlsx'
WHEEL_FILE = 'wheels.xlsx'
CREDIT_FILE_PREFIX = 'debts_'
MEDGULF_FILE_PREFIX = 'medgulf_'
REPORTS_DIR = 'reports'
ARCHIVE_DIR = 'archive'

# --------------------------
# Helpers & file initialization
# --------------------------
def ensure_files():
    # Create directories if they don't exist
    os.makedirs(REPORTS_DIR, exist_ok=True)
    os.makedirs(ARCHIVE_DIR, exist_ok=True)
    
    # User file setup
    if not os.path.exists(USER_FILE):
        wb = Workbook()
        ws = wb.active
        ws.append(['Username', 'Password', 'Role'])
        ws.append(['admin', 'admin123', 'admin'])
        ws.append(['cashier', '1234', 'cashier'])
        wb.save(USER_FILE)

    # Product files with buying price (only visible to admin)
    for file, headers in [
        (PRODUCT_FILE, ['Product', 'Buy Price', 'Sell Price', 'Stock']),
        (OIL_FILE, ['Oil', 'Buy Price', 'Sell Price', 'Stock']), 
        (WHEEL_FILE, ['Wheel', 'Buy Price', 'Sell Price', 'Stock'])
    ]:
        if not os.path.exists(file):
            wb = Workbook()
            ws = wb.active
            ws.append(headers)
            wb.save(file)

    # Monthly transaction files
    current_month = datetime.now().strftime("%Y-%m")
    for prefix in [CREDIT_FILE_PREFIX, MEDGULF_FILE_PREFIX]:
        filename = f"{prefix}{current_month}.xlsx"
        if not os.path.exists(filename):
            wb = Workbook()
            ws = wb.active
            ws.append(['Customer Name', 'Product', 'Price', 'Quantity', 'Total', 'DateTime', 'ReceiptID'])
            wb.save(filename)

    # Daily sales file
    if not os.path.exists(SALES_FILE):
        wb = Workbook()
        ws = wb.active
        ws.append(['Product', 'Price', 'Quantity', 'Total', 'DateTime', 'ReceiptID'])
        wb.save(SALES_FILE)

def get_monthly_file(prefix):
    """Get current month's transaction file"""
    return f"{prefix}{datetime.now().strftime('%Y-%m')}.xlsx"

def archive_old_files():
    """Move old monthly files to archive"""
    current_month = datetime.now().strftime("%Y-%m")
    for filename in os.listdir('.'):
        if filename.startswith((CREDIT_FILE_PREFIX, MEDGULF_FILE_PREFIX)):
            try:
                file_month = filename.split('_')[-1].split('.')[0]
            except Exception:
                continue
            if file_month != current_month:
                try:
                    shutil.move(filename, os.path.join(ARCHIVE_DIR, filename))
                except Exception as e:
                    app.logger.error(f"Error archiving {filename}: {str(e)}")

def read_table(file_path):
    """Read data from Excel file"""
    if not os.path.exists(file_path):
        return []
    wb = load_workbook(file_path)
    ws = wb.active
    return [list(row) for row in ws.iter_rows(min_row=2, values_only=True) if row and row[0] is not None]

# --------------------------
# Authentication
# --------------------------
def validate_login(username, password):
    if not os.path.exists(USER_FILE):
        return None
    wb = load_workbook(USER_FILE)
    ws = wb.active
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row and str(row[0]).lower() == username.lower() and str(row[1]) == password:
            return row[2] if len(row) > 2 and row[2] else 'cashier'
    return None

# --------------------------
# Inventory Management
# --------------------------
def get_products():
    """Get products with buying price (admin only)"""
    out = []
    for r in read_table(PRODUCT_FILE):
        try:
            item = {
                'name': str(r[0]),
                'price': float(r[2]),  # Selling price
                'stock': int(r[3])
            }
            if session.get('role') == 'admin':
                item['buy_price'] = float(r[1])  # Buying price
            out.append(item)
        except Exception:
            continue
    return out

def get_oils():
    """Get oils with buying price (admin only)"""
    out = []
    for r in read_table(OIL_FILE):
        try:
            item = {
                'name': str(r[0]),
                'price': float(r[2]),
                'stock': int(r[3])
            }
            if session.get('role') == 'admin':
                item['buy_price'] = float(r[1])
            out.append(item)
        except Exception:
            continue
    return out

def get_wheels():
    """Get wheels with buying price (admin only)"""
    out = []
    for r in read_table(WHEEL_FILE):
        try:
            item = {
                'name': str(r[0]),
                'price': float(r[2]),
                'stock': int(r[3])
            }
            if session.get('role') == 'admin':
                item['buy_price'] = float(r[1])
            out.append(item)
        except Exception:
            continue
    return out

def get_stock_from_file(file_path, item_name):
    """Check current stock level"""
    if not os.path.exists(file_path): 
        return 0
    wb = load_workbook(file_path)
    ws = wb.active
    for row in ws.iter_rows(min_row=2):
        if row[0].value == item_name:
            try:
                return int(row[3].value or 0)
            except Exception:
                return 0
    return 0

def pending_in_session(item_name, item_type='product'):
    """Check for pending items in current session (to avoid overselling in UI)"""
    pending = 0
    items = session.get('receipt_items', [])
    for it in items:
        name = it.get('name', '')
        qty = int(it.get('quantity', 0))
        if item_type == 'oil' and name.startswith('Oil Change (') and name.endswith(')'):
            inner = name.replace('Oil Change (', '').replace(')', '')
            if inner == item_name:
                pending += qty
        elif item_type == 'wheel' and name.startswith('Wheel Change (') and name.endswith(')'):
            inner = name.replace('Wheel Change (', '').replace(')', '')
            if inner == item_name:
                pending += qty
        elif item_type == 'product' and not ('Oil Change (' in name or 'Wheel Change (' in name):
            if name == item_name:
                pending += qty
    return pending

def get_available_stock(file_path, item_name, item_type='product'):
    """Calculate available stock considering pending items"""
    base = get_stock_from_file(file_path, item_name)
    pending = pending_in_session(item_name, item_type)
    return base - pending

def update_stock(file_path, product_name, quantity):
    """Update stock after sale (subtract quantity). quantity should be positive integer (we subtract inside)"""
    if not os.path.exists(file_path): 
        return
    wb = load_workbook(file_path)
    ws = wb.active
    changed = False
    for row in ws.iter_rows(min_row=2):
        if row[0].value == product_name:
            try:
                curr = int(row[3].value or 0)
                row[3].value = max(0, curr - quantity)
                changed = True
                break
            except Exception:
                continue
    if changed:
        wb.save(file_path)

# --------------------------
# Transaction Processing
# --------------------------
def is_service_or_used(item_name):
    """Return True if item is service or used part (Arabic variant included)"""
    if item_name is None:
        return False
    name = str(item_name)
    return (name.startswith('Service') or name.startswith('Used Part') or name.startswith('قطعة مستعملة') or name.startswith('Used Part:'))

def log_sale(receipt_items, receipt_id):
    """Record cash sale"""
    wb = load_workbook(SALES_FILE)
    ws = wb.active
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for item in receipt_items:
        product, price, qty = item['name'], float(item['price']), int(item['quantity'])
        total = price * qty
        ws.append([product, price, qty, total, now, receipt_id])
        
        # Only update stock if it's a regular product (not service or used part)
        if not is_service_or_used(product):
            if product.startswith('Oil Change (') and product.endswith(')'):
                oil = product.replace('Oil Change (', '').replace(')', '')
                update_stock(OIL_FILE, oil, qty)
            elif product.startswith('Wheel Change (') and product.endswith(')'):
                wheel = product.replace('Wheel Change (', '').replace(')', '')
                update_stock(WHEEL_FILE, wheel, qty)
            else:
                # normal product
                update_stock(PRODUCT_FILE, product, qty)
    wb.save(SALES_FILE)

def log_credit(receipt_items, receipt_id, customer_name):
    """Record credit sale"""
    file_path = get_monthly_file(CREDIT_FILE_PREFIX)
    if not os.path.exists(file_path):
        ensure_files()
        
    wb = load_workbook(file_path)
    ws = wb.active
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for item in receipt_items:
        product, price, qty = item['name'], float(item['price']), int(item['quantity'])
        total = price * qty
        ws.append([customer_name, product, price, qty, total, now, receipt_id])
        
        # Only update stock if it's a regular product (not service or used part)
        if not is_service_or_used(product):
            if product.startswith('Oil Change (') and product.endswith(')'):
                oil = product.replace('Oil Change (', '').replace(')', '')
                update_stock(OIL_FILE, oil, qty)
            elif product.startswith('Wheel Change (') and product.endswith(')'):
                wheel = product.replace('Wheel Change (', '').replace(')', '')
                update_stock(WHEEL_FILE, wheel, qty)
            else:
                update_stock(PRODUCT_FILE, product, qty)
    wb.save(file_path)

def log_medgulf(receipt_items, receipt_id, customer_name):
    """Record MedGulf sale"""
    file_path = get_monthly_file(MEDGULF_FILE_PREFIX)
    if not os.path.exists(file_path):
        ensure_files()
        
    wb = load_workbook(file_path)
    ws = wb.active
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for item in receipt_items:
        product, price, qty = item['name'], float(item['price']), int(item['quantity'])
        total = price * qty
        ws.append([customer_name, product, price, qty, total, now, receipt_id])
        
        # Only update stock if it's a regular product (not service or used part)
        if not is_service_or_used(product):
            if product.startswith('Oil Change (') and product.endswith(')'):
                oil = product.replace('Oil Change (', '').replace(')', '')
                update_stock(OIL_FILE, oil, qty)
            elif product.startswith('Wheel Change (') and product.endswith(')'):
                wheel = product.replace('Wheel Change (', '').replace(')', '')
                update_stock(WHEEL_FILE, wheel, qty)
            else:
                update_stock(PRODUCT_FILE, product, qty)
    wb.save(file_path)

# --------------------------
# Report Generation
# --------------------------
def generate_daily_word_report():
    """Generate daily sales report"""
    today = datetime.now().strftime("%Y-%m-%d")
    doc = Document()
    doc.add_heading(f"Salimco Motorcycle Shop - Daily Report - {today}", level=1)

    # Normal sales
    doc.add_heading("=== Normal Sales ===", level=2)
    sales = [row for row in read_table(SALES_FILE) if str(row[4]).startswith(today)]
    if sales:
        tbl = doc.add_table(rows=1, cols=6)
        hdr = tbl.rows[0].cells
        hdr[0].text='Product'; hdr[1].text='Price'; hdr[2].text='Qty'; hdr[3].text='Total'; hdr[4].text='DateTime'; hdr[5].text='ReceiptID'
        subtotal = 0.0
        for r in sales:
            row = tbl.add_row().cells
            row[0].text = str(r[0])
            try:
                row[1].text = f"{float(r[1]):.2f}"
            except Exception:
                row[1].text = str(r[1])
            row[2].text = str(int(r[2]))
            try:
                row[3].text = f"{float(r[3]):.2f}"
            except Exception:
                row[3].text = str(r[3])
            row[4].text = str(r[4])
            row[5].text = str(r[5])
            try:
                subtotal += float(r[3])
            except Exception:
                pass
        doc.add_paragraph(f"Subtotal (Normal): {subtotal:.2f}")
    else:
        doc.add_paragraph("No normal sales for today.")

    # Debts
    doc.add_heading("=== Debt Transactions ===", level=2)
    debts = [row for row in read_table(get_monthly_file(CREDIT_FILE_PREFIX)) if str(row[5]).startswith(today)]
    if debts:
        tbl = doc.add_table(rows=1, cols=7)
        hdr = tbl.rows[0].cells
        hdr[0].text='Customer'; hdr[1].text='Product'; hdr[2].text='Price'; hdr[3].text='Qty'; hdr[4].text='Total'; hdr[5].text='DateTime'; hdr[6].text='ReceiptID'
        subtotal = 0.0
        for r in debts:
            row = tbl.add_row().cells
            row[0].text = str(r[0]); row[1].text = str(r[1])
            try:
                row[2].text = f"{float(r[2]):.2f}"
            except Exception:
                row[2].text = str(r[2])
            row[3].text = str(int(r[3])); 
            try:
                row[4].text = f"{float(r[4]):.2f}"
            except Exception:
                row[4].text = str(r[4])
            row[5].text = str(r[5]); row[6].text = str(r[6])
            try:
                subtotal += float(r[4])
            except Exception:
                pass
        doc.add_paragraph(f"Subtotal (Debts): {subtotal:.2f}")
    else:
        doc.add_paragraph("No debt transactions for today.")

    # MedGulf
    doc.add_heading("=== MedGulf Transactions ===", level=2)
    med = [row for row in read_table(get_monthly_file(MEDGULF_FILE_PREFIX)) if str(row[5]).startswith(today)]
    if med:
        tbl = doc.add_table(rows=1, cols=7)
        hdr = tbl.rows[0].cells
        hdr[0].text='Customer'; hdr[1].text='Product'; hdr[2].text='Price'; hdr[3].text='Qty'; hdr[4].text='Total'; hdr[5].text='DateTime'; hdr[6].text='ReceiptID'
        subtotal = 0.0
        for r in med:
            row = tbl.add_row().cells
            row[0].text = str(r[0]); row[1].text = str(r[1])
            try:
                row[2].text = f"{float(r[2]):.2f}"
            except Exception:
                row[2].text = str(r[2])
            row[3].text = str(int(r[3])); 
            try:
                row[4].text = f"{float(r[4]):.2f}"
            except Exception:
                row[4].text = str(r[4])
            row[5].text = str(r[5]); row[6].text = str(r[6])
            try:
                subtotal += float(r[4])
            except Exception:
                pass
        doc.add_paragraph(f"Subtotal (MedGulf): {subtotal:.2f}")
    else:
        doc.add_paragraph("No MedGulf transactions for today.")

    # Services & Used Parts - appear in the Normal Sales already, but add a categorized summary section for clarity
    doc.add_heading("=== Services and Used Parts Summary ===", level=2)
    services = []
    used_parts = []
    for r in read_table(SALES_FILE):
        # row structure: [Product, Price, Quantity, Total, DateTime, ReceiptID]
        try:
            dt = str(r[4])
        except Exception:
            dt = ''
        if not dt.startswith(today):
            continue
        pname = str(r[0])
        if is_service_or_used(pname):
            # service names start with 'Service' prefix; used parts start with 'Used Part'
            if pname.lower().startswith('service'):
                services.append(r)
            else:
                used_parts.append(r)
    if services:
        tbl = doc.add_table(rows=1, cols=6)
        hdr = tbl.rows[0].cells
        hdr[0].text='Product'; hdr[1].text='Price'; hdr[2].text='Qty'; hdr[3].text='Total'; hdr[4].text='DateTime'; hdr[5].text='ReceiptID'
        for r in services:
            row = tbl.add_row().cells
            row[0].text = str(r[0])
            try: row[1].text = f"{float(r[1]):.2f}"
            except: row[1].text = str(r[1])
            row[2].text = str(int(r[2]))
            try: row[3].text = f"{float(r[3]):.2f}"
            except: row[3].text = str(r[3])
            row[4].text = str(r[4]); row[5].text = str(r[5])
    else:
        doc.add_paragraph("No services for today.")

    if used_parts:
        tbl = doc.add_table(rows=1, cols=6)
        hdr = tbl.rows[0].cells
        hdr[0].text='Product'; hdr[1].text='Price'; hdr[2].text='Qty'; hdr[3].text='Total'; hdr[4].text='DateTime'; hdr[5].text='ReceiptID'
        for r in used_parts:
            row = tbl.add_row().cells
            row[0].text = str(r[0])
            try: row[1].text = f"{float(r[1]):.2f}"
            except: row[1].text = str(r[1])
            row[2].text = str(int(r[2]))
            try: row[3].text = f"{float(r[3]):.2f}"
            except: row[3].text = str(r[3])
            row[4].text = str(r[4]); row[5].text = str(r[5])
    else:
        doc.add_paragraph("No used parts for today.")

    filename = f"Daily_Report_{today}.docx"
    filepath = os.path.join(REPORTS_DIR, filename)
    doc.save(filepath)
    return filepath

def generate_debts_word_report():
    """Generate detailed monthly debts report with customer subtotals"""
    doc = Document()
    month = datetime.now().strftime("%Y-%m")
    doc.add_heading(f"Salimco - Monthly Debts Report - {month}", level=1)
    
    # Get all transactions for the month
    transactions = read_table(get_monthly_file(CREDIT_FILE_PREFIX))
    
    if not transactions:
        doc.add_paragraph("No debt transactions for this month.")
        filepath = os.path.join(REPORTS_DIR, f"Debts_Report_{month}.docx")
        doc.save(filepath)
        return filepath
    
    # Calculate customer totals
    customer_totals = {}
    for t in transactions:
        customer = t[0]
        try:
            amount = float(t[4])
        except Exception:
            amount = 0.0
        customer_totals[customer] = customer_totals.get(customer, 0) + amount
    
    # Add customer summary
    doc.add_heading("Customer Debt Summary", level=2)
    summary_table = doc.add_table(rows=1, cols=2)
    hdr = summary_table.rows[0].cells
    hdr[0].text = 'Customer'
    hdr[1].text = 'Total Owed'
    
    for customer, total in sorted(customer_totals.items()):
        row = summary_table.add_row().cells
        row[0].text = customer
        row[1].text = f"{total:.2f}"
    
    doc.add_paragraph(f"\nGrand Total: {sum(customer_totals.values()):.2f}")
    
    # Add transaction details
    doc.add_heading("Transaction Details", level=2)
    details_table = doc.add_table(rows=1, cols=7)
    hdr = details_table.rows[0].cells
    hdr[0].text = 'Customer'
    hdr[1].text = 'Product'
    hdr[2].text = 'Price'
    hdr[3].text = 'Qty'
    hdr[4].text = 'Total'
    hdr[5].text = 'Date'
    hdr[6].text = 'Receipt'
    
    for t in transactions:
        row = details_table.add_row().cells
        row[0].text = str(t[0])
        row[1].text = str(t[1])
        try:
            row[2].text = f"{float(t[2]):.2f}"
        except Exception:
            row[2].text = str(t[2])
        row[3].text = str(int(t[3]))
        try:
            row[4].text = f"{float(t[4]):.2f}"
        except Exception:
            row[4].text = str(t[4])
        row[5].text = str(t[5])
        row[6].text = str(t[6])
    
    filename = f"Debts_Report_{month}.docx"
    filepath = os.path.join(REPORTS_DIR, filename)
    doc.save(filepath)
    return filepath

def generate_medgulf_word_report():
    """Generate monthly MedGulf report"""
    month = datetime.now().strftime("%Y-%m")
    doc = Document()
    doc.add_heading(f"Salimco - MedGulf Report - {month}", level=1)
    
    med = read_table(get_monthly_file(MEDGULF_FILE_PREFIX))
    if med:
        tbl = doc.add_table(rows=1, cols=7)
        hdr = tbl.rows[0].cells
        hdr[0].text='Customer'; hdr[1].text='Product'; hdr[2].text='Price'; hdr[3].text='Qty'; hdr[4].text='Total'; hdr[5].text='DateTime'; hdr[6].text='ReceiptID'
        subtotal = 0.0
        for r in med:
            row = tbl.add_row().cells
            row[0].text = str(r[0]); row[1].text = str(r[1])
            try: row[2].text = f"{float(r[2]):.2f}"
            except: row[2].text = str(r[2])
            row[3].text = str(int(r[3])); 
            try: row[4].text = f"{float(r[4]):.2f}"
            except: row[4].text = str(r[4])
            row[5].text = str(r[5]); row[6].text = str(r[6])
            try:
                subtotal += float(r[4])
            except Exception:
                pass
        doc.add_paragraph(f"Subtotal (MedGulf): {subtotal:.2f}")
    else:
        doc.add_paragraph("No MedGulf transactions for this month.")
    
    filename = f"MedGulf_Report_{month}.docx"
    filepath = os.path.join(REPORTS_DIR, filename)
    doc.save(filepath)
    return filepath

# --------------------------
# Application Routes
# --------------------------
@app.route('/', methods=['GET', 'POST'])
def login():
    ensure_files()
    if request.method == 'POST':
        username = request.form.get('username','')
        password = request.form.get('password','')
        role = validate_login(username, password)
        if role:
            session['username'] = username
            session['role'] = role
            return redirect(url_for('pos'))
        flash('Invalid username or password', 'danger')
    return render_template('login.html', shop_name='Salimco Motorcycle Shop')

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

@app.route('/pos', methods=['GET', 'POST'])
def pos():
    if 'username' not in session:
        return redirect(url_for('login'))

    if 'receipt_items' not in session:
        session['receipt_items'] = []
        session['receipt_id'] = datetime.now().strftime("%Y%m%d%H%M%S")

    if request.method == 'POST':
        action = request.form.get('action')

        if action == 'add_product':
            product_name = request.form.get('product_name')
            try:
                qty = int(request.form.get('quantity',1))
            except:
                qty = 1
            product = next((p for p in get_products() if p['name'] == product_name), None)
            if product:
                available = get_available_stock(PRODUCT_FILE, product_name, 'product')
                if qty > available:
                    flash(f"Only {available} available in stock", 'warning')
                else:
                    session['receipt_items'].append({'name': product_name, 'price': product['price'], 'quantity': qty})
                    session.modified = True
                    flash('Added product to receipt', 'success')

        elif action == 'add_oil':
            oil_name = request.form.get('oil_name')
            try:
                qty = int(request.form.get('quantity',1))
            except:
                qty = 1
            oil = next((o for o in get_oils() if o['name'] == oil_name), None)
            if oil:
                available = get_available_stock(OIL_FILE, oil_name, 'oil')
                if qty > available:
                    flash(f"Only {available} oil in stock", 'warning')
                else:
                    item_name = f"Oil Change ({oil_name})"
                    session['receipt_items'].append({'name': item_name, 'price': oil['price'], 'quantity': qty})
                    session.modified = True
                    flash('Added oil change to receipt', 'success')

        elif action == 'add_wheel':
            wheel_name = request.form.get('wheel_name')
            try:
                qty = int(request.form.get('quantity',1))
            except:
                qty = 1
            wheel = next((w for w in get_wheels() if w['name'] == wheel_name), None)
            if wheel:
                available = get_available_stock(WHEEL_FILE, wheel_name, 'wheel')
                if qty > available:
                    flash(f"Only {available} wheel in stock", 'warning')
                else:
                    item_name = f"Wheel Change ({wheel_name})"
                    session['receipt_items'].append({'name': item_name, 'price': wheel['price'], 'quantity': qty})
                    session.modified = True
                    flash('Added wheel change to receipt', 'success')

        elif action == 'add_service':
            service_name = request.form.get('service_name', 'Service Charge').strip()
            try:
                price = float(request.form.get('service_price', 0))
                if price <= 0:
                    flash('Service price must be positive', 'warning')
                else:
                    # mark services with prefix 'Service' - won't affect stock
                    session['receipt_items'].append({'name': f"Service: {service_name}", 'price': price, 'quantity': 1})
                    session.modified = True
                    flash('Added service charge to receipt', 'success')
            except ValueError:
                flash('Invalid service price', 'danger')

        elif action == 'add_used_part':
            part_name = request.form.get('part_name', 'قطعة مستعملة').strip()
            try:
                price = float(request.form.get('part_price', 0))
                if price <= 0:
                    flash('Part price must be positive', 'warning')
                else:
                    # mark used parts as 'Used Part: ...' - won't affect stock
                    session['receipt_items'].append({'name': f"Used Part: {part_name}", 'price': price, 'quantity': 1})
                    session.modified = True
                    flash('Added used part to receipt', 'success')
            except ValueError:
                flash('Invalid part price', 'danger')

        elif action == 'finalize_cash':
            if not session.get('receipt_items'):
                flash('Receipt empty', 'warning')
            else:
                rid = session.get('receipt_id')
                log_sale(session['receipt_items'], rid)
                session['receipt_items'] = []
                session['receipt_id'] = datetime.now().strftime("%Y%m%d%H%M%S")
                session.modified = True
                flash(f"Saved Receipt #{rid} (Cash)", 'success')

        elif action == 'finalize_credit':
            if not session.get('receipt_items'):
                flash('Receipt empty', 'warning')
            else:
                customer_name = request.form.get('customer_name','').strip()
                if not customer_name:
                    flash('Customer name required for credit', 'warning')
                else:
                    rid = session.get('receipt_id')
                    log_credit(session['receipt_items'], rid, customer_name)
                    session['receipt_items'] = []
                    session['receipt_id'] = datetime.now().strftime("%Y%m%d%H%M%S")
                    session.modified = True
                    flash(f"Saved Receipt #{rid} (Credit)", 'success')

        elif action == 'finalize_medgulf':
            if not session.get('receipt_items'):
                flash('Receipt empty', 'warning')
            else:
                customer_name = request.form.get('customer_name','').strip()
                if not customer_name:
                    flash('Customer name required for MedGulf', 'warning')
                else:
                    rid = session.get('receipt_id')
                    log_medgulf(session['receipt_items'], rid, customer_name)
                    session['receipt_items'] = []
                    session['receipt_id'] = datetime.now().strftime("%Y%m%d%H%M%S")
                    session.modified = True
                    flash(f"Saved Receipt #{rid} (MedGulf)", 'success')

        return redirect(url_for('pos'))

    total = sum(float(i['price']) * int(i['quantity']) for i in session.get('receipt_items', []))
    return render_template('pos.html',
                           products=get_products(),
                           oils=get_oils(),
                           wheels=get_wheels(),
                           receipt_items=session.get('receipt_items', []),
                           total=total,
                           receipt_id=session.get('receipt_id'),
                           role=session.get('role'),
                           shop_name='Salimco Motorcycle Shop')

@app.route('/remove_from_cart/<int:index>', methods=['POST'])
def remove_from_cart(index):
    items = session.get('receipt_items', [])
    if 0 <= index < len(items):
        removed = items.pop(index)
        session['receipt_items'] = items
        session.modified = True
        flash(f"Removed {removed['name']} x{removed['quantity']}", 'info')
    else:
        flash('Invalid item index', 'danger')
    return redirect(url_for('pos'))

@app.route('/inventory', methods=['GET', 'POST'])
def inventory():
    if 'username' not in session or session.get('role') != 'admin':
        flash('Access denied', 'danger')
        return redirect(url_for('pos'))

    q = request.args.get('q', '').strip().lower()

    if request.method == 'POST':
        action = request.form.get('action')
        if action == 'add_product':
            name = request.form.get('product_name','').strip()
            try:
                buy_price = float(request.form.get('buy_price',0))
            except:
                buy_price = 0.0
            try:
                sell_price = float(request.form.get('sell_price',0))
            except:
                sell_price = 0.0
            try:
                stock = int(request.form.get('stock',0))
            except:
                stock = 0
            wb = load_workbook(PRODUCT_FILE)
            ws = wb.active
            ws.append([name, buy_price, sell_price, stock])
            wb.save(PRODUCT_FILE)
            flash(f"Product '{name}' added", 'success')

        elif action == 'add_oil':
            name = request.form.get('oil_name','').strip()
            try:
                buy_price = float(request.form.get('buy_price',0))
            except:
                buy_price = 0.0
            try:
                sell_price = float(request.form.get('sell_price',0))
            except:
                sell_price = 0.0
            try:
                stock = int(request.form.get('stock',0))
            except:
                stock = 0
            wb = load_workbook(OIL_FILE)
            ws = wb.active
            ws.append([name, buy_price, sell_price, stock])
            wb.save(OIL_FILE)
            flash(f"Oil '{name}' added", 'success')

        elif action == 'add_wheel':
            name = request.form.get('wheel_name','').strip()
            try:
                buy_price = float(request.form.get('buy_price',0))
            except:
                buy_price = 0.0
            try:
                sell_price = float(request.form.get('sell_price',0))
            except:
                sell_price = 0.0
            try:
                stock = int(request.form.get('stock',0))
            except:
                stock = 0
            wb = load_workbook(WHEEL_FILE)
            ws = wb.active
            ws.append([name, buy_price, sell_price, stock])
            wb.save(WHEEL_FILE)
            flash(f"Wheel '{name}' added", 'success')

        return redirect(url_for('inventory'))

    products = get_products()
    oils = get_oils()
    wheels = get_wheels()

    if q:
        products = [p for p in products if q in p['name'].lower()]
        oils = [o for o in oils if q in o['name'].lower()]
        wheels = [w for w in wheels if q in w['name'].lower()]

    return render_template('inventory.html',
                           products=products,
                           oils=oils,
                           wheels=wheels,
                           q=q,
                           shop_name='Salimco Motorcycle Shop')

@app.route('/report/daily')
def report_daily():
    try:
        filename = generate_daily_word_report()
        return send_file(filename, as_attachment=True)
    except Exception as e:
        flash(f"Error generating daily report: {str(e)}", "danger")
        return redirect(url_for('pos'))

@app.route('/report/debts')
def report_debts():
    try:
        archive_old_files()
        filename = generate_debts_word_report()
        return send_file(filename, as_attachment=True)
    except Exception as e:
        flash(f"Error generating debts report: {str(e)}", "danger")
        return redirect(url_for('pos'))

@app.route('/report/medgulf')
def report_medgulf():
    try:
        archive_old_files()
        filename = generate_medgulf_word_report()
        return send_file(filename, as_attachment=True)
    except Exception as e:
        flash(f"Error generating MedGulf report: {str(e)}", "danger")
        return redirect(url_for('pos'))

if __name__ == '__main__':
    ensure_files()
    archive_old_files()
    app.run(host='0.0.0.0', port=5000, debug=False)
