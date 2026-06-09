from flask import Flask, request, jsonify, render_template, send_from_directory, send_file, session
import sqlite3, os, io
from datetime import datetime
from werkzeug.utils import secure_filename
from functools import wraps
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

app = Flask(__name__)
app.secret_key = 'dozakaz-secret-key-2026'
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024

ALLOWED = {'.xlsx', '.xls', '.csv'}

USERS = {
    'admin': {'password': '123456', 'role': 'admin', 'branch': None},
    'ALAYSKIY': {'password': 'LI-NING1', 'role': 'user', 'branch': 'ALAYSKIY'},
    'ATLAS CHIMGAN': {'password': 'LI-NING1', 'role': 'user', 'branch': 'ATLAS CHIMGAN'},
    'ECO PARK': {'password': 'LI-NING1', 'role': 'user', 'branch': 'ECO PARK'},
    'Family park': {'password': 'LI-NING1', 'role': 'user', 'branch': 'Family park'},
    'HIGH TOWN PLAZA': {'password': 'LI-NING1', 'role': 'user', 'branch': 'HIGH TOWN PLAZA'},
    'M. BARAKA': {'password': 'LI-NING1', 'role': 'user', 'branch': 'M. BARAKA'},
    'MAGIC CITY': {'password': 'LI-NING1', 'role': 'user', 'branch': 'MAGIC CITY'},
    'MALIKA': {'password': 'LI-NING1', 'role': 'user', 'branch': 'MALIKA'},
    'NOVZA': {'password': 'LI-NING1', 'role': 'user', 'branch': 'NOVZA'},
    'Scopus Mall': {'password': 'LI-NING1', 'role': 'user', 'branch': 'Scopus Mall'},
    'Shota Rustavely': {'password': 'LI-NING1', 'role': 'user', 'branch': 'Shota Rustavely'},
    'TASHKENT CITY MALL': {'password': 'LI-NING1', 'role': 'user', 'branch': 'TASHKENT CITY MALL'},
    'UZBEGIM ANDIJAN': {'password': 'LI-NING1', 'role': 'user', 'branch': 'UZBEGIM ANDIJAN'},
    'Yunusabad gallery': {'password': 'LI-NING1', 'role': 'user', 'branch': 'Yunusabad gallery'},
}

BRANCHES = ['ALAYSKIY','ATLAS CHIMGAN','ECO PARK','Family park','HIGH TOWN PLAZA',
            'M. BARAKA','MAGIC CITY','MALIKA','NOVZA','Scopus Mall',
            'Shota Rustavely','TASHKENT CITY MALL','UZBEGIM ANDIJAN','Yunusabad gallery']

def get_db():
    db = sqlite3.connect('dozakaz.db')
    db.row_factory = sqlite3.Row
    return db

def init_db():
    db = get_db()
    db.execute('''CREATE TABLE IF NOT EXISTS orders (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        branch TEXT NOT NULL,
        responsible TEXT NOT NULL,
        date TEXT NOT NULL,
        priority TEXT DEFAULT 'Обычный',
        note TEXT DEFAULT '',
        filename TEXT,
        original_name TEXT,
        status TEXT DEFAULT 'Новая',
        created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )''')
    db.execute('''CREATE TABLE IF NOT EXISTS order_items (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        order_id INTEGER,
        article TEXT,
        name TEXT,
        size TEXT,
        qty INTEGER,
        wms_stock INTEGER DEFAULT 0,
        FOREIGN KEY(order_id) REFERENCES orders(id)
    )''')
    db.execute('''CREATE TABLE IF NOT EXISTS catalog (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        article TEXT,
        name TEXT,
        size TEXT,
        wms_stock INTEGER DEFAULT 0,
        abc TEXT DEFAULT 'C',
        sold INTEGER DEFAULT 0,
        updated_at TEXT DEFAULT CURRENT_TIMESTAMP
    )''')
    db.execute('''CREATE TABLE IF NOT EXISTS branch_stock (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        article TEXT,
        size TEXT,
        branch TEXT,
        qty INTEGER DEFAULT 0
    )''')
    db.execute('''CREATE TABLE IF NOT EXISTS top_products (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        article TEXT,
        category TEXT,
        sold INTEGER,
        stock INTEGER,
        updated_at TEXT DEFAULT CURRENT_TIMESTAMP
    )''')
    db.commit()
    db.close()

os.makedirs('uploads', exist_ok=True)
init_db()

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'username' not in session:
            return jsonify({'error': 'Unauthorized'}), 401
        return f(*args, **kwargs)
    return decorated

def admin_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'username' not in session:
            return jsonify({'error': 'Unauthorized'}), 401
        if session.get('role') != 'admin':
            return jsonify({'error': 'Forbidden'}), 403
        return f(*args, **kwargs)
    return decorated

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/login', methods=['POST'])
def login():
    data = request.get_json()
    username = data.get('username', '').strip()
    password = data.get('password', '').strip()
    user = USERS.get(username)
    if not user or user['password'] != password:
        return jsonify({'error': 'Неверный логин или пароль'}), 401
    session['username'] = username
    session['role'] = user['role']
    session['branch'] = user['branch']
    return jsonify({'role': user['role'], 'branch': user['branch'], 'username': username})

@app.route('/api/logout', methods=['POST'])
def logout():
    session.clear()
    return jsonify({'ok': True})

@app.route('/api/me')
def me():
    if 'username' not in session:
        return jsonify({'logged_in': False})
    return jsonify({'logged_in': True, 'role': session['role'], 'branch': session['branch'], 'username': session['username']})

@app.route('/api/orders', methods=['GET'])
@login_required
def get_orders():
    branch = request.args.get('branch', '')
    status = request.args.get('status', '')
    responsible = request.args.get('responsible', '')
    date_from = request.args.get('date_from', '')
    date_to = request.args.get('date_to', '')
    db = get_db()
    q = 'SELECT * FROM orders WHERE 1=1'
    params = []
    if session['role'] == 'user':
        q += ' AND branch=?'; params.append(session['branch'])
    else:
        if branch:
            q += ' AND branch=?'; params.append(branch)
    if status:
        q += ' AND status=?'; params.append(status)
    if responsible:
        q += ' AND responsible LIKE ?'; params.append(f'%{responsible}%')
    if date_from:
        q += ' AND date >= ?'; params.append(date_from)
    if date_to:
        q += ' AND date <= ?'; params.append(date_to)
    q += ' ORDER BY created_at DESC'
    rows = db.execute(q, params).fetchall()
    db.close()
    return jsonify([dict(r) for r in rows])

@app.route('/api/orders/<int:oid>/items', methods=['GET'])
@login_required
def get_order_items(oid):
    db = get_db()
    rows = db.execute('SELECT * FROM order_items WHERE order_id=?', (oid,)).fetchall()
    db.close()
    return jsonify([dict(r) for r in rows])

def generate_order_excel(branch, items):
    """Generate Excel file in LI-NING order format and return as BytesIO"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Заказ'

    # Column widths
    ws.column_dimensions['A'].width = 6
    ws.column_dimensions['B'].width = 18
    ws.column_dimensions['C'].width = 18
    ws.column_dimensions['D'].width = 10

    # Row 1 empty
    ws.row_dimensions[1].height = 6

    # Row 2: Branch header (merged B2:D2, yellow background)
    ws.merge_cells('B2:D2')
    cell = ws['B2']
    cell.value = f'Филиал {branch}'
    cell.font = Font(name='Arial', bold=True, size=11)
    cell.fill = PatternFill('solid', fgColor='FFFF00')
    cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[2].height = 20

    # Row 3: Headers
    thin = Side(style='thin')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    headers = [('A3', '№'), ('B3', 'Артикул'), ('C3', 'Характеристика'), ('D3', 'Кол-во')]
    for coord, val in headers:
        c = ws[coord]
        c.value = val
        c.font = Font(name='Arial', bold=True, size=10)
        c.alignment = Alignment(horizontal='center', vertical='center')
        c.border = border
    ws.row_dimensions[3].height = 16

    # Data rows starting at row 4
    for i, item in enumerate(items, 1):
        row = i + 3
        ws.cell(row=row, column=1, value=i).alignment = Alignment(horizontal='center')
        ws.cell(row=row, column=2, value=item.get('article', ''))
        ws.cell(row=row, column=3, value=item.get('size', '')).alignment = Alignment(horizontal='center')
        ws.cell(row=row, column=4, value=item.get('qty', 1)).alignment = Alignment(horizontal='center')
        for col in range(1, 5):
            ws.cell(row=row, column=col).border = border
            ws.cell(row=row, column=col).font = Font(name='Arial', size=10)
        ws.row_dimensions[row].height = 15

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


@app.route('/api/orders', methods=['POST'])
@login_required
def create_order():
    branch = session['branch'] if session['role'] == 'user' else request.form.get('branch', '').strip()
    responsible = request.form.get('responsible', '').strip()
    date = request.form.get('date', datetime.now().strftime('%Y-%m-%d'))
    priority = request.form.get('priority', 'Обычный')
    note = request.form.get('note', '').strip()
    items_json = request.form.get('items', '[]')

    if not branch or not responsible:
        return jsonify({'error': 'Заполните обязательные поля'}), 400

    import json
    items = json.loads(items_json)

    filename = None
    original_name = None
    f = request.files.get('file')
    if f and f.filename:
        ext = os.path.splitext(f.filename)[1].lower()
        if ext not in ALLOWED:
            return jsonify({'error': 'Только .xlsx, .xls, .csv'}), 400
        original_name = f.filename
        filename = f'{datetime.now().strftime("%Y%m%d_%H%M%S")}_{secure_filename(f.filename)}'
        f.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))

    # Auto-generate Excel from catalog items if no file uploaded
    if not filename and items:
        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        branch_slug = branch.replace(' ', '_').replace('.', '')
        original_name = f'ДОЗАКАЗ_{branch_slug}_{ts}.xlsx'
        filename = f'{ts}_{secure_filename(original_name)}'
        excel_buf = generate_order_excel(branch, items)
        with open(os.path.join(app.config['UPLOAD_FOLDER'], filename), 'wb') as out:
            out.write(excel_buf.read())

    db = get_db()
    cur = db.execute(
        'INSERT INTO orders (branch,responsible,date,priority,note,filename,original_name) VALUES (?,?,?,?,?,?,?)',
        (branch, responsible, date, priority, note, filename, original_name)
    )
    order_id = cur.lastrowid
    for item in items:
        db.execute('INSERT INTO order_items (order_id,article,name,size,qty,wms_stock) VALUES (?,?,?,?,?,?)',
                   (order_id, item.get('article',''), item.get('name',''), item.get('size',''), item.get('qty',1), item.get('wms_stock',0)))
    db.commit()
    row = db.execute('SELECT * FROM orders WHERE id=?', (order_id,)).fetchone()
    db.close()
    return jsonify(dict(row)), 201

@app.route('/api/orders/<int:oid>/status', methods=['PATCH'])
@admin_required
def update_status(oid):
    data = request.get_json()
    status = data.get('status')
    if status not in ('Новая', 'В работе', 'Выполнена'):
        return jsonify({'error': 'Invalid status'}), 400
    db = get_db()
    db.execute('UPDATE orders SET status=? WHERE id=?', (status, oid))
    db.commit()
    row = db.execute('SELECT * FROM orders WHERE id=?', (oid,)).fetchone()
    db.close()
    return jsonify(dict(row))

@app.route('/api/orders/<int:oid>/excel')
@login_required
def download_order_excel(oid):
    db = get_db()
    row = db.execute('SELECT * FROM orders WHERE id=?', (oid,)).fetchone()
    if not row:
        db.close()
        return 'Not found', 404
    items = db.execute('SELECT * FROM order_items WHERE order_id=?', (oid,)).fetchall()
    db.close()
    items_list = [dict(i) for i in items]
    branch = row['branch']
    buf = generate_order_excel(branch, items_list)
    ts = row['date'].replace('-', '') if row['date'] else datetime.now().strftime('%Y%m%d')
    fname = f"ДОЗАКАЗ_{branch.replace(' ', '_')}_{ts}.xlsx"
    return send_file(buf, as_attachment=True, download_name=fname,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/api/download/<int:oid>')
@login_required
def download(oid):
    db = get_db()
    row = db.execute('SELECT * FROM orders WHERE id=?', (oid,)).fetchone()
    db.close()
    if not row or not row['filename']:
        return 'Not found', 404
    return send_from_directory(app.config['UPLOAD_FOLDER'], row['filename'],
                               as_attachment=True, download_name=row['original_name'])

@app.route('/api/stats')
@login_required
def stats():
    db = get_db()
    if session['role'] == 'admin':
        total = db.execute('SELECT COUNT(*) FROM orders').fetchone()[0]
        new = db.execute("SELECT COUNT(*) FROM orders WHERE status='Новая'").fetchone()[0]
        inprog = db.execute("SELECT COUNT(*) FROM orders WHERE status='В работе'").fetchone()[0]
        done = db.execute("SELECT COUNT(*) FROM orders WHERE status='Выполнена'").fetchone()[0]
    else:
        b = session['branch']
        total = db.execute('SELECT COUNT(*) FROM orders WHERE branch=?', (b,)).fetchone()[0]
        new = db.execute("SELECT COUNT(*) FROM orders WHERE branch=? AND status='Новая'", (b,)).fetchone()[0]
        inprog = db.execute("SELECT COUNT(*) FROM orders WHERE branch=? AND status='В работе'", (b,)).fetchone()[0]
        done = db.execute("SELECT COUNT(*) FROM orders WHERE branch=? AND status='Выполнена'", (b,)).fetchone()[0]
    db.close()
    return jsonify({'total': total, 'new': new, 'in_progress': inprog, 'done': done})

@app.route('/api/catalog/upload', methods=['POST'])
@admin_required
def upload_catalog():
    f = request.files.get('file')
    if not f:
        return jsonify({'error': 'Файл не найден'}), 400
    try:
        import openpyxl
        wb = openpyxl.load_workbook(io.BytesIO(f.read()), data_only=True)
        ws = wb.active

        rows = list(ws.iter_rows(values_only=True))

        # Find header row (row with 'Артикул')
        header_idx = None
        branch_cols = {}
        wms_col = None

        for i, row in enumerate(rows):
            row_vals = [str(c).strip() if c else '' for c in row]
            if 'Артикул' in row_vals:
                header_idx = i
                # Next row has 'Доступно' markers
                avail_row = [str(c).strip() if c else '' for c in rows[i+1]]
                for j, val in enumerate(row_vals):
                    if val in BRANCHES:
                        branch_cols[val] = j
                    if val == 'Склад WMS':
                        wms_col = j
                break

        if header_idx is None:
            return jsonify({'error': 'Не найден заголовок таблицы'}), 400

        # Find name column (Номенклатура, Характеристика)
        header_row = [str(c).strip() if c else '' for c in rows[header_idx]]
        name_col = None
        art_col = 0
        for j, val in enumerate(header_row):
            if 'Номенклатура' in val or 'Характеристика' in val:
                name_col = j
                break

        db = get_db()
        db.execute('DELETE FROM catalog')
        db.execute('DELETE FROM branch_stock')

        catalog_items = []
        branch_items = []

        for row in rows[header_idx+2:]:
            if not row or not row[art_col]:
                continue
            art = str(row[art_col]).strip()
            if not art or art == 'None':
                continue

            name_full = str(row[name_col]).strip() if name_col and row[name_col] else ''
            # Parse name and size: "AAFW021-3 Basketball Woven S/S Top Navy Blazer, L"
            name = name_full
            size = ''
            if ', ' in name_full:
                parts = name_full.rsplit(', ', 1)
                name = parts[0].strip()
                size = parts[1].strip()
            # Remove article from name
            if name.startswith(art.rstrip('A')):
                name = name[len(art.rstrip('A')):].strip()

            wms = int(float(str(row[wms_col]))) if wms_col and row[wms_col] and str(row[wms_col]) not in ('None','nan') else 0

            catalog_items.append((art, name, size, wms))

            for branch, col in branch_cols.items():
                qty = row[col]
                if qty and str(qty) not in ('None', 'nan'):
                    try:
                        q = int(float(str(qty)))
                        if q > 0:
                            branch_items.append((art, size, branch, q))
                    except:
                        pass

        db.executemany('INSERT INTO catalog (article, name, size, wms_stock) VALUES (?,?,?,?)', catalog_items)
        db.executemany('INSERT INTO branch_stock (article, size, branch, qty) VALUES (?,?,?,?)', branch_items)
        db.commit()
        db.close()
        return jsonify({'ok': True, 'count': len(catalog_items)})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/catalog', methods=['GET'])
@login_required
def get_catalog():
    search = request.args.get('search', '')
    page = int(request.args.get('page', 1))
    per_page = 50
    branch = session.get('branch') or request.args.get('branch', '')

    db = get_db()

    # Get unique articles with their sizes
    q = '''SELECT DISTINCT article, name FROM catalog WHERE 1=1'''
    params = []
    if search:
        q += ' AND (article LIKE ? OR name LIKE ?)'
        params.extend([f'%{search}%', f'%{search}%'])
    q += ' ORDER BY article'
    q += f' LIMIT {per_page} OFFSET {(page-1)*per_page}'

    articles = db.execute(q, params).fetchall()

    result = []
    for art_row in articles:
        art = art_row['article']
        name = art_row['name']

        # Get all sizes with WMS stock
        sizes = db.execute('SELECT size, wms_stock FROM catalog WHERE article=? ORDER BY size', (art,)).fetchall()

        # Get branch stock if branch provided
        branch_stock = {}
        if branch:
            bs = db.execute('SELECT size, qty FROM branch_stock WHERE article=? AND branch=?', (art, branch)).fetchall()
            branch_stock = {r['size']: r['qty'] for r in bs}

        # Get ABC
        abc_row = db.execute('SELECT abc, sold FROM catalog WHERE article=? LIMIT 1', (art,)).fetchone()
        abc = abc_row['abc'] if abc_row else 'C'
        sold = abc_row['sold'] if abc_row else 0

        total_wms = sum(r['wms_stock'] for r in sizes)

        result.append({
            'article': art,
            'name': name,
            'abc': abc,
            'sold': sold,
            'total_wms': total_wms,
            'sizes': [{'size': r['size'], 'wms': r['wms_stock'], 'branch': branch_stock.get(r['size'], 0)} for r in sizes]
        })

    db.close()
    return jsonify(result)

@app.route('/api/catalog/abc', methods=['POST'])
@admin_required
def update_abc():
    '''Update ABC from top products file'''
    f = request.files.get('file')
    if not f:
        return jsonify({'error': 'Файл не найден'}), 400
    try:
        import openpyxl
        wb = openpyxl.load_workbook(io.BytesIO(f.read()), data_only=True)
        db = get_db()
        count = 0
        for ws in wb.worksheets:
            rows = list(ws.iter_rows(values_only=True))
            art_col = sold_col = None
            for i, row in enumerate(rows):
                row_s = [str(c).strip() if c else '' for c in row]
                if any(x in row_s for x in ['ITEM NO.', 'Одежда', 'Обувь']):
                    for j, v in enumerate(row_s):
                        if v in ('ITEM NO.', 'Одежда', 'Обувь'): art_col = j
                        if v == 'Продано': sold_col = j
                    for row2 in rows[i+1:]:
                        if not row2 or not row2[art_col]: continue
                        art = str(row2[art_col]).strip()
                        if art in ('Одежда','Обувь','Аксессуары','nan',''): continue
                        sold = int(float(str(row2[sold_col]))) if sold_col and row2[sold_col] else 0
                        abc = 'A' if sold >= 25 else 'B' if sold >= 15 else 'C'
                        db.execute('UPDATE catalog SET abc=?, sold=? WHERE article LIKE ?', (abc, sold, art+'%'))
                        count += 1
                    break
        db.commit()
        db.close()
        return jsonify({'ok': True, 'updated': count})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/products', methods=['GET'])
@login_required
def get_products():
    db = get_db()
    rows = db.execute('SELECT * FROM top_products ORDER BY sold DESC').fetchall()
    db.close()
    return jsonify([dict(r) for r in rows])

@app.route('/api/products/upload', methods=['POST'])
@admin_required
def upload_products():
    f = request.files.get('file')
    if not f:
        return jsonify({'error': 'Файл не найден'}), 400
    try:
        import openpyxl
        wb = openpyxl.load_workbook(io.BytesIO(f.read()), data_only=True)
        products = []
        for ws in wb.worksheets:
            rows = list(ws.iter_rows(values_only=True))
            art_col = sold_col = stock_col = cat_col = None
            header_row_idx = None
            for i, row in enumerate(rows):
                row_str = [str(c).strip() if c else '' for c in row]
                if any(x in row_str for x in ['ITEM NO.', 'Одежда', 'Обувь']):
                    header_row_idx = i
                    for j, cell in enumerate(row_str):
                        if cell in ('ITEM NO.', 'Одежда', 'Обувь', 'Аксессуары'): art_col = j
                        if cell == 'Продано': sold_col = j
                        if cell == 'Остаток': stock_col = j
                        if j == 1: cat_col = j
                    break
            if header_row_idx is not None and art_col is not None:
                current_cat = 'APP'
                for row in rows[header_row_idx+1:]:
                    if not row or not row[art_col]: continue
                    art = str(row[art_col]).strip()
                    if art in ('Одежда', 'Обувь', 'Аксессуары', 'nan', ''):
                        if art == 'Обувь': current_cat = 'FTW'
                        elif art == 'Аксессуары': current_cat = 'ACC'
                        continue
                    cat = str(row[cat_col]).strip() if cat_col and row[cat_col] else current_cat
                    sold = int(float(str(row[sold_col]))) if sold_col and row[sold_col] and str(row[sold_col]) != 'None' else 0
                    stock = int(float(str(row[stock_col]))) if stock_col and row[stock_col] and str(row[stock_col]) != 'None' else 0
                    if art and art != 'None' and len(art) > 3:
                        products.append((art, cat, sold, stock))
        db = get_db()
        db.execute('DELETE FROM top_products')
        for p in products:
            db.execute('INSERT INTO top_products (article, category, sold, stock) VALUES (?,?,?,?)', p)
        db.commit()
        db.close()
        return jsonify({'ok': True, 'count': len(products)})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/analytics/orders')
@admin_required
def analytics_orders():
    db = get_db()
    top_articles = db.execute('''
        SELECT article, SUM(qty) as total_qty, COUNT(DISTINCT order_id) as order_count
        FROM order_items WHERE article != '' AND article != 'None'
        GROUP BY article ORDER BY total_qty DESC LIMIT 20
    ''').fetchall()
    branch_totals = db.execute('''
        SELECT o.branch, SUM(oi.qty) as total_qty, COUNT(DISTINCT o.id) as order_count
        FROM orders o JOIN order_items oi ON o.id = oi.order_id
        GROUP BY o.branch ORDER BY total_qty DESC
    ''').fetchall()
    db.close()
    return jsonify({'top_articles': [dict(r) for r in top_articles], 'branch_totals': [dict(r) for r in branch_totals]})

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)