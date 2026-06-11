from flask import Flask, request, jsonify, render_template, send_from_directory, send_file, session, Response
import os, io, re
from datetime import datetime
from werkzeug.utils import secure_filename
from functools import wraps
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import psycopg2
from psycopg2.extras import RealDictCursor

app = Flask(__name__)
app.secret_key = 'dozakaz-secret-key-2026'
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024

ALLOWED = {'.xlsx', '.xls', '.csv'}
DATABASE_URL = os.environ.get('DATABASE_URL', '')

USERS = {
    'admin': {'password': '123456', 'role': 'admin', 'branch': None, 'branches': None},
    'ANNA_V': {'password': 'region1', 'role': 'regional', 'branch': None, 'branches': ['HIGH TOWN PLAZA','MAGIC CITY','NOVZA','Scopus Mall']},
    'ANNA_G': {'password': 'region2', 'role': 'regional', 'branch': None, 'branches': ['ATLAS CHIMGAN','ECO PARK','HIGH TOWN PLAZA','MALIKA','Shota Rustavely','Yunusabad gallery']},
    'ALAYSKIY': {'password': 'LI-NING1', 'role': 'user', 'branch': 'ALAYSKIY', 'branches': None},
    'ATLAS CHIMGAN': {'password': 'LI-NING1', 'role': 'user', 'branch': 'ATLAS CHIMGAN', 'branches': None},
    'ECO PARK': {'password': 'LI-NING1', 'role': 'user', 'branch': 'ECO PARK', 'branches': None},
    'Family park': {'password': 'LI-NING1', 'role': 'user', 'branch': 'Family park', 'branches': None},
    'HIGH TOWN PLAZA': {'password': 'LI-NING1', 'role': 'user', 'branch': 'HIGH TOWN PLAZA', 'branches': None},
    'M. BARAKA': {'password': 'LI-NING1', 'role': 'user', 'branch': 'M. BARAKA', 'branches': None},
    'MAGIC CITY': {'password': 'LI-NING1', 'role': 'user', 'branch': 'MAGIC CITY', 'branches': None},
    'MALIKA': {'password': 'LI-NING1', 'role': 'user', 'branch': 'MALIKA', 'branches': None},
    'NOVZA': {'password': 'LI-NING1', 'role': 'user', 'branch': 'NOVZA', 'branches': None},
    'Scopus Mall': {'password': 'LI-NING1', 'role': 'user', 'branch': 'Scopus Mall', 'branches': None},
    'Shota Rustavely': {'password': 'LI-NING1', 'role': 'user', 'branch': 'Shota Rustavely', 'branches': None},
    'TASHKENT CITY MALL': {'password': 'LI-NING1', 'role': 'user', 'branch': 'TASHKENT CITY MALL', 'branches': None},
    'UZBEGIM ANDIJAN': {'password': 'LI-NING1', 'role': 'user', 'branch': 'UZBEGIM ANDIJAN', 'branches': None},
    'Yunusabad gallery': {'password': 'LI-NING1', 'role': 'user', 'branch': 'Yunusabad gallery', 'branches': None},
}

BRANCHES = ['ALAYSKIY','ATLAS CHIMGAN','ECO PARK','Family park','HIGH TOWN PLAZA',
            'M. BARAKA','MAGIC CITY','MALIKA','NOVZA','Scopus Mall',
            'Shota Rustavely','TASHKENT CITY MALL','UZBEGIM ANDIJAN','Yunusabad gallery']

YADISK_TOKEN = os.environ.get('YADISK_TOKEN', '35a4110c9f5a4729ba8a54cf978276f4')
YADISK_PUBLIC_KEY = 'https://disk.yandex.com/d/-plm2CMx-kHNuA'
YADISK_FOLDER = '/06-White Background Pics'
_photo_cache = {}
_photo_url_cache = {}

def get_db():
    conn = psycopg2.connect(DATABASE_URL, cursor_factory=RealDictCursor)
    return conn

def init_db():
    conn = get_db()
    cur = conn.cursor()
    cur.execute('''CREATE TABLE IF NOT EXISTS orders (
        id SERIAL PRIMARY KEY,
        branch TEXT NOT NULL,
        responsible TEXT NOT NULL,
        date TEXT NOT NULL,
        priority TEXT DEFAULT 'Обычный',
        note TEXT DEFAULT '',
        filename TEXT,
        original_name TEXT,
        status TEXT DEFAULT 'Новая',
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )''')
    cur.execute('''CREATE TABLE IF NOT EXISTS order_items (
        id SERIAL PRIMARY KEY,
        order_id INTEGER REFERENCES orders(id),
        article TEXT,
        name TEXT,
        size TEXT,
        qty INTEGER,
        wms_stock INTEGER DEFAULT 0
    )''')
    cur.execute('''CREATE TABLE IF NOT EXISTS catalog (
        id SERIAL PRIMARY KEY,
        article TEXT,
        name TEXT,
        size TEXT,
        wms_stock INTEGER DEFAULT 0,
        abc TEXT DEFAULT 'C',
        sold INTEGER DEFAULT 0,
        season TEXT DEFAULT '',
        category TEXT DEFAULT '',
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )''')
    cur.execute('''CREATE TABLE IF NOT EXISTS branch_stock (
        id SERIAL PRIMARY KEY,
        article TEXT,
        size TEXT,
        branch TEXT,
        qty INTEGER DEFAULT 0
    )''')
    cur.execute('''CREATE TABLE IF NOT EXISTS top_products (
        id SERIAL PRIMARY KEY,
        article TEXT,
        category TEXT,
        sold INTEGER,
        stock INTEGER,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )''')
    cur.execute('''CREATE TABLE IF NOT EXISTS sales (
        id SERIAL PRIMARY KEY,
        season TEXT,
        article TEXT,
        category TEXT,
        branch TEXT,
        sale_date DATE,
        qty INTEGER DEFAULT 1,
        price NUMERIC DEFAULT 0,
        amount NUMERIC DEFAULT 0,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )''')
    conn.commit()
    cur.close()
    conn.close()

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
    session['branches'] = user.get('branches')
    return jsonify({'role': user['role'], 'branch': user['branch'], 'branches': user.get('branches'), 'username': username})

@app.route('/api/logout', methods=['POST'])
def logout():
    session.clear()
    return jsonify({'ok': True})

@app.route('/api/me')
def me():
    if 'username' not in session:
        return jsonify({'logged_in': False})
    return jsonify({'logged_in': True, 'role': session['role'], 'branch': session['branch'], 'branches': session.get('branches'), 'username': session['username']})

@app.route('/api/orders', methods=['GET'])
@login_required
def get_orders():
    branch = request.args.get('branch', '')
    status = request.args.get('status', '')
    responsible = request.args.get('responsible', '')
    date_from = request.args.get('date_from', '')
    date_to = request.args.get('date_to', '')
    conn = get_db(); cur = conn.cursor()
    q = 'SELECT * FROM orders WHERE 1=1'
    params = []
    if session['role'] == 'user':
        q += ' AND branch=%s'; params.append(session['branch'])
    elif session['role'] == 'regional':
        bs = session.get('branches', [])
        if bs:
            q += ' AND branch = ANY(%s)'; params.append(bs)
        if branch and branch in bs:
            q += ' AND branch=%s'; params.append(branch)
    else:
        if branch: q += ' AND branch=%s'; params.append(branch)
    if status: q += ' AND status=%s'; params.append(status)
    if responsible: q += ' AND responsible ILIKE %s'; params.append(f'%{responsible}%')
    if date_from: q += ' AND date >= %s'; params.append(date_from)
    if date_to: q += ' AND date <= %s'; params.append(date_to)
    q += ' ORDER BY created_at DESC'
    cur.execute(q, params)
    rows = cur.fetchall()
    cur.close(); conn.close()
    return jsonify([dict(r) for r in rows])

@app.route('/api/orders/<int:oid>/items', methods=['GET'])
@login_required
def get_order_items(oid):
    conn = get_db(); cur = conn.cursor()
    cur.execute('SELECT * FROM order_items WHERE order_id=%s', (oid,))
    rows = cur.fetchall()
    cur.close(); conn.close()
    return jsonify([dict(r) for r in rows])

def generate_order_excel(branch, items):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Заказ'
    ws.column_dimensions['A'].width = 6
    ws.column_dimensions['B'].width = 18
    ws.column_dimensions['C'].width = 18
    ws.column_dimensions['D'].width = 10
    ws.row_dimensions[1].height = 6
    ws.merge_cells('B2:D2')
    cell = ws['B2']
    cell.value = f'Филиал {branch}'
    cell.font = Font(name='Arial', bold=True, size=11)
    cell.fill = PatternFill('solid', fgColor='FFFF00')
    cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[2].height = 20
    thin = Side(style='thin')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    headers = [('A3','№'),('B3','Артикул'),('C3','Характеристика'),('D3','Кол-во')]
    for coord, val in headers:
        c = ws[coord]
        c.value = val
        c.font = Font(name='Arial', bold=True, size=10)
        c.alignment = Alignment(horizontal='center', vertical='center')
        c.border = border
    ws.row_dimensions[3].height = 16
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
    import json
    branch = session['branch'] if session['role'] == 'user' else request.form.get('branch', '').strip()
    responsible = request.form.get('responsible', '').strip()
    date = request.form.get('date', datetime.now().strftime('%Y-%m-%d'))
    priority = request.form.get('priority', 'Обычный')
    note = request.form.get('note', '').strip()
    items_json = request.form.get('items', '[]')
    if not branch or not responsible:
        return jsonify({'error': 'Заполните обязательные поля'}), 400
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
    if not filename and items:
        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        branch_slug = branch.replace(' ','_').replace('.','')
        original_name = f'ДОЗАКАЗ_{branch_slug}_{ts}.xlsx'
        filename = f'{ts}_{secure_filename(original_name)}'
        excel_buf = generate_order_excel(branch, items)
        with open(os.path.join(app.config['UPLOAD_FOLDER'], filename), 'wb') as out:
            out.write(excel_buf.read())
    conn = get_db(); cur = conn.cursor()
    cur.execute(
        'INSERT INTO orders (branch,responsible,date,priority,note,filename,original_name) VALUES (%s,%s,%s,%s,%s,%s,%s) RETURNING id',
        (branch, responsible, date, priority, note, filename, original_name)
    )
    order_id = cur.fetchone()['id']
    for item in items:
        cur.execute('INSERT INTO order_items (order_id,article,name,size,qty,wms_stock) VALUES (%s,%s,%s,%s,%s,%s)',
                   (order_id, item.get('article',''), item.get('name',''), item.get('size',''), item.get('qty',1), item.get('wms_stock',0)))
    conn.commit()
    cur.execute('SELECT * FROM orders WHERE id=%s', (order_id,))
    row = cur.fetchone()
    cur.close(); conn.close()
    return jsonify(dict(row)), 201

@app.route('/api/orders/<int:oid>/status', methods=['PATCH'])
@login_required
def update_status(oid):
    if session.get('role') not in ('admin', 'regional'):
        return jsonify({'error': 'Forbidden'}), 403
    data = request.get_json()
    status = data.get('status')
    if status not in ('Новая', 'В работе', 'Выполнена'):
        return jsonify({'error': 'Invalid status'}), 400
    conn = get_db(); cur = conn.cursor()
    cur.execute('UPDATE orders SET status=%s WHERE id=%s', (status, oid))
    conn.commit()
    cur.execute('SELECT * FROM orders WHERE id=%s', (oid,))
    row = cur.fetchone()
    cur.close(); conn.close()
    return jsonify(dict(row))

@app.route('/api/orders/<int:oid>/excel')
@login_required
def download_order_excel(oid):
    conn = get_db(); cur = conn.cursor()
    cur.execute('SELECT * FROM orders WHERE id=%s', (oid,))
    row = cur.fetchone()
    if not row:
        cur.close(); conn.close()
        return 'Not found', 404
    cur.execute('SELECT * FROM order_items WHERE order_id=%s', (oid,))
    items = cur.fetchall()
    cur.close(); conn.close()
    buf = generate_order_excel(row['branch'], [dict(i) for i in items])
    ts = row['date'].replace('-','') if row['date'] else datetime.now().strftime('%Y%m%d')
    fname = f"ДОЗАКАЗ_{row['branch'].replace(' ','_')}_{ts}.xlsx"
    return send_file(buf, as_attachment=True, download_name=fname,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@app.route('/api/download/<int:oid>')
@login_required
def download(oid):
    conn = get_db(); cur = conn.cursor()
    cur.execute('SELECT * FROM orders WHERE id=%s', (oid,))
    row = cur.fetchone()
    cur.close(); conn.close()
    if not row or not row['filename']:
        return 'Not found', 404
    return send_from_directory(app.config['UPLOAD_FOLDER'], row['filename'],
                               as_attachment=True, download_name=row['original_name'])

@app.route('/api/stats')
@login_required
def stats():
    conn = get_db(); cur = conn.cursor()
    if session['role'] == 'admin':
        cur.execute('SELECT COUNT(*) as c FROM orders'); total = cur.fetchone()['c']
        cur.execute("SELECT COUNT(*) as c FROM orders WHERE status='Новая'"); new = cur.fetchone()['c']
        cur.execute("SELECT COUNT(*) as c FROM orders WHERE status='В работе'"); inprog = cur.fetchone()['c']
        cur.execute("SELECT COUNT(*) as c FROM orders WHERE status='Выполнена'"); done = cur.fetchone()['c']
    elif session['role'] == 'regional':
        bs = session.get('branches', [])
        cur.execute('SELECT COUNT(*) as c FROM orders WHERE branch = ANY(%s)', (bs,)); total = cur.fetchone()['c']
        cur.execute("SELECT COUNT(*) as c FROM orders WHERE branch = ANY(%s) AND status='Новая'", (bs,)); new = cur.fetchone()['c']
        cur.execute("SELECT COUNT(*) as c FROM orders WHERE branch = ANY(%s) AND status='В работе'", (bs,)); inprog = cur.fetchone()['c']
        cur.execute("SELECT COUNT(*) as c FROM orders WHERE branch = ANY(%s) AND status='Выполнена'", (bs,)); done = cur.fetchone()['c']
    else:
        b = session['branch']
        cur.execute('SELECT COUNT(*) as c FROM orders WHERE branch=%s', (b,)); total = cur.fetchone()['c']
        cur.execute("SELECT COUNT(*) as c FROM orders WHERE branch=%s AND status='Новая'", (b,)); new = cur.fetchone()['c']
        cur.execute("SELECT COUNT(*) as c FROM orders WHERE branch=%s AND status='В работе'", (b,)); inprog = cur.fetchone()['c']
        cur.execute("SELECT COUNT(*) as c FROM orders WHERE branch=%s AND status='Выполнена'", (b,)); done = cur.fetchone()['c']
    cur.close(); conn.close()
    return jsonify({'total': total, 'new': new, 'in_progress': inprog, 'done': done})

@app.route('/api/catalog/upload', methods=['POST'])
@admin_required
def upload_catalog():
    f = request.files.get('file')
    if not f:
        return jsonify({'error': 'Файл не найден'}), 400
    try:
        wb = openpyxl.load_workbook(io.BytesIO(f.read()), data_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        header_idx = None
        branch_cols = {}
        wms_col = None
        for i, row in enumerate(rows):
            row_vals = [str(c).strip() if c else '' for c in row]
            if 'Артикул' in row_vals:
                header_idx = i
                for j, val in enumerate(row_vals):
                    val_norm = val.replace('\u0421','C').replace('\u0441','c').replace('\u0415','E').replace('\u0435','e')
                    matched = None
                    for branch in BRANCHES:
                        branch_norm = branch.replace('\u0421','C').replace('\u0441','c').replace('\u0415','E').replace('\u0435','e')
                        if val_norm == branch_norm or val == branch:
                            matched = branch; break
                    if matched: branch_cols[matched] = j
                    if val == 'Склад WMS': wms_col = j
                break
        if header_idx is None:
            return jsonify({'error': 'Не найден заголовок таблицы'}), 400
        header_row = [str(c).strip() if c else '' for c in rows[header_idx]]
        name_col = size_col = season_col = category_col = None
        art_col = 0
        for j, val in enumerate(header_row):
            if val == 'Характеристика' and size_col is None: size_col = j
            if 'Номенклатура' in val and ',' in val and name_col is None: name_col = j
            if 'Сезон' in val: season_col = j
            if 'Вид' in val and category_col is None: category_col = j
        if name_col is None:
            for j, val in enumerate(header_row):
                if 'Номенклатура' in val: name_col = j; break
        data_start = header_idx + 1
        for i in range(header_idx+1, min(header_idx+3, len(rows))):
            rv = [str(c).strip() if c else '' for c in rows[i]]
            if any('Доступно' in v for v in rv):
                data_start = i + 1; break
        conn = get_db(); cur = conn.cursor()
        cur.execute('DELETE FROM catalog')
        cur.execute('DELETE FROM branch_stock')
        catalog_items = []
        branch_items = []
        for row in rows[data_start:]:
            if not row or not row[art_col]: continue
            art = str(row[art_col]).strip()
            if not art or art in ('None','nan'): continue
            size = ''
            if size_col is not None and row[size_col]:
                size = str(row[size_col]).strip()
            name = ''
            if name_col is not None and row[name_col]:
                name_full = str(row[name_col]).strip()
                if not size and ', ' in name_full:
                    parts = name_full.rsplit(', ', 1)
                    name_full = parts[0].strip()
                    size = parts[1].strip()
                art_base = art.rstrip('A')
                if name_full.startswith(art_base):
                    name_full = name_full[len(art_base):].strip()
                if size and name_full.endswith(f', {size}'):
                    name_full = name_full[:-len(f', {size}')].strip()
                name = name_full
            season = str(row[season_col]).strip() if season_col is not None and row[season_col] and str(row[season_col]) not in ('None','nan') else ''
            category = str(row[category_col]).strip() if category_col is not None and row[category_col] and str(row[category_col]) not in ('None','nan') else ''
            wms = 0
            if wms_col is not None and row[wms_col] and str(row[wms_col]) not in ('None','nan'):
                try: wms = int(float(str(row[wms_col])))
                except: pass
            catalog_items.append((art, name, size, wms, season, category))
            for branch, col in branch_cols.items():
                qty = row[col]
                if qty and str(qty) not in ('None','nan'):
                    try:
                        q = int(float(str(qty)))
                        if q > 0: branch_items.append((art, size, branch, q))
                    except: pass
        cur.executemany('INSERT INTO catalog (article,name,size,wms_stock,season,category) VALUES (%s,%s,%s,%s,%s,%s)', catalog_items)
        cur.executemany('INSERT INTO branch_stock (article,size,branch,qty) VALUES (%s,%s,%s,%s)', branch_items)
        conn.commit(); cur.close(); conn.close()
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
    conn = get_db(); cur = conn.cursor()
    q = 'SELECT article, MIN(name) as name FROM catalog WHERE 1=1'
    params = []
    if search:
        q += ' AND (article ILIKE %s OR name ILIKE %s)'
        params.extend([f'%{search}%', f'%{search}%'])
    q += ' GROUP BY article ORDER BY article'
    q += f' LIMIT {per_page} OFFSET {(page-1)*per_page}'
    cur.execute(q, params)
    articles = cur.fetchall()
    result = []
    for art_row in articles:
        art = art_row['article']
        name = art_row['name']
        cur.execute('SELECT size, wms_stock FROM catalog WHERE article=%s ORDER BY size', (art,))
        sizes = cur.fetchall()
        branch_stock = {}
        if branch:
            cur.execute('SELECT size, qty FROM branch_stock WHERE article=%s AND branch=%s', (art, branch))
            branch_stock = {r['size']: r['qty'] for r in cur.fetchall()}
        cur.execute('SELECT abc, sold, season, category FROM catalog WHERE article=%s LIMIT 1', (art,))
        abc_row = cur.fetchone()
        abc = abc_row['abc'] if abc_row else 'C'
        sold = abc_row['sold'] if abc_row else 0
        season = abc_row['season'] if abc_row and abc_row['season'] else ''
        category = abc_row['category'] if abc_row and abc_row['category'] else ''
        total_wms = sum(r['wms_stock'] for r in sizes)
        result.append({
            'article': art, 'name': name, 'abc': abc, 'sold': sold,
            'season': season, 'category': category, 'total_wms': total_wms,
            'sizes': [{'size': r['size'], 'wms': r['wms_stock'], 'branch': branch_stock.get(r['size'], 0)} for r in sizes]
        })
    cur.close(); conn.close()
    return jsonify(result)

@app.route('/api/catalog/abc', methods=['POST'])
@admin_required
def update_abc():
    f = request.files.get('file')
    if not f:
        return jsonify({'error': 'Файл не найден'}), 400
    try:
        wb = openpyxl.load_workbook(io.BytesIO(f.read()), data_only=True)
        conn = get_db(); cur = conn.cursor()
        count = 0
        for ws in wb.worksheets:
            rows = list(ws.iter_rows(values_only=True))
            art_col = sold_col = None
            for i, row in enumerate(rows):
                row_s = [str(c).strip() if c else '' for c in row]
                if any(x in row_s for x in ['ITEM NO.','Одежда','Обувь']):
                    for j, v in enumerate(row_s):
                        if v in ('ITEM NO.','Одежда','Обувь'): art_col = j
                        if v == 'Продано': sold_col = j
                    for row2 in rows[i+1:]:
                        if not row2 or not row2[art_col]: continue
                        art = str(row2[art_col]).strip()
                        if art in ('Одежда','Обувь','Аксессуары','nan',''): continue
                        sold = int(float(str(row2[sold_col]))) if sold_col and row2[sold_col] else 0
                        abc = 'A' if sold >= 25 else 'B' if sold >= 15 else 'C'
                        cur.execute('UPDATE catalog SET abc=%s, sold=%s WHERE article LIKE %s', (abc, sold, art+'%'))
                        count += 1
                    break
        conn.commit(); cur.close(); conn.close()
        return jsonify({'ok': True, 'updated': count})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/catalog/abc-sync', methods=['POST'])
@admin_required
def abc_sync():
    data = request.get_json()
    updates = data.get('updates', [])
    if not updates:
        return jsonify({'error': 'Нет данных'}), 400
    conn = get_db(); cur = conn.cursor()
    count = 0
    for u in updates:
        art = u.get('article', '')
        abc = u.get('abc', 'C')
        sold = u.get('sold', 0)
        cur.execute('UPDATE catalog SET abc=%s, sold=%s WHERE article=%s', (abc, sold, art))
        count += cur.rowcount
        cur.execute('UPDATE catalog SET abc=%s, sold=%s WHERE article=%s', (abc, sold, art + 'A'))
        count += cur.rowcount
    conn.commit(); cur.close(); conn.close()
    return jsonify({'ok': True, 'updated': count})

@app.route('/api/products', methods=['GET'])
@login_required
def get_products():
    conn = get_db(); cur = conn.cursor()
    cur.execute('SELECT * FROM top_products ORDER BY sold DESC')
    rows = cur.fetchall()
    cur.close(); conn.close()
    return jsonify([dict(r) for r in rows])

@app.route('/api/products/upload', methods=['POST'])
@admin_required
def upload_products():
    f = request.files.get('file')
    if not f:
        return jsonify({'error': 'Файл не найден'}), 400
    try:
        wb = openpyxl.load_workbook(io.BytesIO(f.read()), data_only=True)
        products = []
        for ws in wb.worksheets:
            rows = list(ws.iter_rows(values_only=True))
            art_col = sold_col = stock_col = cat_col = None
            header_row_idx = None
            for i, row in enumerate(rows):
                row_str = [str(c).strip() if c else '' for c in row]
                if any(x in row_str for x in ['ITEM NO.','Одежда','Обувь']):
                    header_row_idx = i
                    for j, cell in enumerate(row_str):
                        if cell in ('ITEM NO.','Одежда','Обувь','Аксессуары'): art_col = j
                        if cell == 'Продано': sold_col = j
                        if cell == 'Остаток': stock_col = j
                        if j == 1: cat_col = j
                    break
            if header_row_idx is not None and art_col is not None:
                current_cat = 'APP'
                for row in rows[header_row_idx+1:]:
                    if not row or not row[art_col]: continue
                    art = str(row[art_col]).strip()
                    if art in ('Одежда','Обувь','Аксессуары','nan',''):
                        if art == 'Обувь': current_cat = 'FTW'
                        elif art == 'Аксессуары': current_cat = 'ACC'
                        continue
                    cat = str(row[cat_col]).strip() if cat_col and row[cat_col] else current_cat
                    sold = int(float(str(row[sold_col]))) if sold_col and row[sold_col] and str(row[sold_col]) != 'None' else 0
                    stock = int(float(str(row[stock_col]))) if stock_col and row[stock_col] and str(row[stock_col]) != 'None' else 0
                    if art and art != 'None' and len(art) > 3:
                        products.append((art, cat, sold, stock))
        conn = get_db(); cur = conn.cursor()
        cur.execute('DELETE FROM top_products')
        cur.executemany('INSERT INTO top_products (article,category,sold,stock) VALUES (%s,%s,%s,%s)', products)
        conn.commit(); cur.close(); conn.close()
        return jsonify({'ok': True, 'count': len(products)})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/analytics/orders')
@login_required
def analytics_orders():
    if session.get('role') not in ('admin', 'regional'):
        return jsonify({'error': 'Forbidden'}), 403
    conn = get_db(); cur = conn.cursor()
    if session['role'] == 'regional':
        bs = session.get('branches', [])
        cur.execute('''SELECT oi.article, SUM(oi.qty) as total_qty, COUNT(DISTINCT oi.order_id) as order_count
            FROM order_items oi JOIN orders o ON o.id = oi.order_id
            WHERE o.branch = ANY(%s) AND oi.article != %s AND oi.article != %s
            GROUP BY oi.article ORDER BY total_qty DESC LIMIT 20''', (bs, '', 'None'))
        top_articles = cur.fetchall()
        cur.execute('''SELECT o.branch, SUM(oi.qty) as total_qty, COUNT(DISTINCT o.id) as order_count
            FROM orders o JOIN order_items oi ON o.id = oi.order_id
            WHERE o.branch = ANY(%s)
            GROUP BY o.branch ORDER BY total_qty DESC''', (bs,))
        branch_totals = cur.fetchall()
    else:
        cur.execute('''SELECT article, SUM(qty) as total_qty, COUNT(DISTINCT order_id) as order_count
            FROM order_items WHERE article != '' AND article != 'None'
            GROUP BY article ORDER BY total_qty DESC LIMIT 20''')
        top_articles = cur.fetchall()
        cur.execute('''SELECT o.branch, SUM(oi.qty) as total_qty, COUNT(DISTINCT o.id) as order_count
            FROM orders o JOIN order_items oi ON o.id = oi.order_id
            GROUP BY o.branch ORDER BY total_qty DESC''')
        branch_totals = cur.fetchall()
    cur.close(); conn.close()
    return jsonify({'top_articles': [dict(r) for r in top_articles], 'branch_totals': [dict(r) for r in branch_totals]})

# ===== SALES =====

@app.route('/api/sales/upload', methods=['POST'])
@admin_required
def upload_sales():
    f = request.files.get('file')
    if not f:
        return jsonify({'error': 'Файл не найден'}), 400
    try:
        wb = openpyxl.load_workbook(io.BytesIO(f.read()), data_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        header_idx = None
        for i, row in enumerate(rows):
            row_s = [str(c).strip() if c else '' for c in row]
            if 'Арт' in row_s or 'Артикул' in row_s:
                header_idx = i; break
        if header_idx is None:
            return jsonify({'error': 'Не найден заголовок'}), 400
        h1 = [str(c).strip() if c else '' for c in rows[header_idx]]
        h2 = [str(c).strip() if c else '' for c in rows[header_idx+1]] if header_idx+1 < len(rows) else []
        season_col = art_col = cat_col = branch_col = ref_col = qty_col = price_col = amount_col = None
        for j, v in enumerate(h1):
            if v == 'Магазин': branch_col = j
            if v == 'Ссылка': ref_col = j
            if v == 'Количество': qty_col = j
            if v == 'Цена': price_col = j
            if v == 'Сумма': amount_col = j
        for j, v in enumerate(h2):
            if 'Сезон' in v: season_col = j
            if 'Артикул' in v: art_col = j
            if 'Вид' in v: cat_col = j
        replace = request.form.get('replace', 'false') == 'true'
        conn = get_db(); cur = conn.cursor()
        if replace:
            cur.execute('DELETE FROM sales')
        inserted = 0
        for row in rows[header_idx+2:]:
            if not row: continue
            art = str(row[art_col]).strip() if art_col is not None and row[art_col] else ''
            if not art or art in ('None','nan',''): continue
            season = str(row[season_col]).strip() if season_col is not None and row[season_col] else ''
            cat = str(row[cat_col]).strip() if cat_col is not None and row[cat_col] else ''
            branch = str(row[branch_col]).strip() if branch_col is not None and row[branch_col] else ''
            ref = str(row[ref_col]).strip() if ref_col is not None and row[ref_col] else ''
            try: qty = int(float(str(row[qty_col]))) if qty_col is not None and row[qty_col] and str(row[qty_col]) not in ('None','nan') else 1
            except: qty = 1
            try: price = float(str(row[price_col])) if price_col is not None and row[price_col] and str(row[price_col]) not in ('None','nan') else 0
            except: price = 0
            try: amount = float(str(row[amount_col])) if amount_col is not None and row[amount_col] and str(row[amount_col]) not in ('None','nan') else 0
            except: amount = 0
            sale_date = None
            m = re.search(r'от (\d{2})\.(\d{2})\.(\d{4})', ref)
            if m:
                try: sale_date = f"{m.group(3)}-{m.group(2)}-{m.group(1)}"
                except: pass
            cur.execute(
                'INSERT INTO sales (season,article,category,branch,sale_date,qty,price,amount) VALUES (%s,%s,%s,%s,%s,%s,%s,%s)',
                (season, art.rstrip('A'), cat, branch, sale_date, qty, price, amount)
            )
            inserted += 1
        conn.commit(); cur.close(); conn.close()
        return jsonify({'ok': True, 'inserted': inserted})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/sales/analytics', methods=['GET'])
@login_required
def sales_analytics():
    period = request.args.get('period', 'month')
    category = request.args.get('category', '')
    branch = request.args.get('branch', '')
    limit = int(request.args.get('limit', 20))
    conn = get_db(); cur = conn.cursor()
    extra = ''
    params = []
    if period == 'week':
        extra += " AND sale_date >= CURRENT_DATE - INTERVAL '7 days'"
    elif period == 'month':
        extra += " AND sale_date >= CURRENT_DATE - INTERVAL '30 days'"
    elif period == 'season':
        extra += " AND sale_date >= CURRENT_DATE - INTERVAL '90 days'"
    if session['role'] == 'regional':
        bs = session.get('branches', [])
        extra += ' AND branch = ANY(%s)'; params.append(bs)
    if category:
        extra += ' AND category = %s'; params.append(category)
    if branch:
        extra += ' AND branch = %s'; params.append(branch)
    cur.execute(f'''SELECT article, category,
        SUM(qty) as total_qty, SUM(amount) as total_amount, COUNT(*) as tx_count
        FROM sales WHERE 1=1 {extra}
        GROUP BY article, category ORDER BY total_qty DESC LIMIT %s
    ''', params + [limit])
    top_articles = cur.fetchall()
    cur.execute(f'''SELECT branch, SUM(qty) as total_qty, SUM(amount) as total_amount
        FROM sales WHERE 1=1 {extra} GROUP BY branch ORDER BY total_qty DESC
    ''', params)
    by_branch = cur.fetchall()
    cur.execute(f'''SELECT category, SUM(qty) as total_qty, SUM(amount) as total_amount
        FROM sales WHERE 1=1 {extra} GROUP BY category ORDER BY total_qty DESC
    ''', params)
    by_category = cur.fetchall()
    cur.execute(f'''SELECT SUM(qty) as total_qty, SUM(amount) as total_amount, COUNT(*) as tx_count
        FROM sales WHERE 1=1 {extra}
    ''', params)
    totals = cur.fetchone()
    cur.execute('SELECT MIN(sale_date) as min_date, MAX(sale_date) as max_date FROM sales')
    dates = cur.fetchone()
    cur.close(); conn.close()
    return jsonify({
        'top_articles': [dict(r) for r in top_articles],
        'by_branch': [dict(r) for r in by_branch],
        'by_category': [dict(r) for r in by_category],
        'totals': dict(totals) if totals else {},
        'dates': {'min': str(dates['min_date']) if dates and dates['min_date'] else None,
                  'max': str(dates['max_date']) if dates and dates['max_date'] else None}
    })

@app.route('/api/sales/info')
@admin_required
def sales_info():
    conn = get_db(); cur = conn.cursor()
    cur.execute('SELECT COUNT(*) as c FROM sales'); total = cur.fetchone()['c']
    cur.execute('SELECT MIN(sale_date) as mn, MAX(sale_date) as mx FROM sales'); dates = cur.fetchone()
    cur.close(); conn.close()
    return jsonify({'total': total, 'min_date': str(dates['mn']) if dates['mn'] else None, 'max_date': str(dates['mx']) if dates['mx'] else None})

# ===== PHOTOS =====

@app.route('/api/photos/<article>')
def get_photos(article):
    import urllib.request, urllib.parse, json as _json
    art_base = article.rstrip('A')
    if art_base in _photo_cache:
        return jsonify(_photo_cache[art_base])
    folder_path = f"{YADISK_FOLDER}/{art_base}"
    url = ("https://cloud-api.yandex.net/v1/disk/public/resources?"
        + urllib.parse.urlencode({'public_key': YADISK_PUBLIC_KEY, 'path': folder_path, 'limit': 20, 'preview_size': '400x400'}))
    try:
        req = urllib.request.Request(url, headers={'Authorization': f'OAuth {YADISK_TOKEN}'})
        with urllib.request.urlopen(req, timeout=8) as resp:
            data = _json.loads(resp.read())
        items = data.get('_embedded', {}).get('items', [])
        photos = []
        for item in items:
            name = item.get('name', '')
            if name.lower().endswith(('.jpg','.jpeg','.png','.webp')):
                dl_url = ("https://cloud-api.yandex.net/v1/disk/public/resources/download?"
                    + urllib.parse.urlencode({'public_key': YADISK_PUBLIC_KEY, 'path': f"{folder_path}/{name}"}))
                req2 = urllib.request.Request(dl_url, headers={'Authorization': f'OAuth {YADISK_TOKEN}'})
                with urllib.request.urlopen(req2, timeout=8) as r2:
                    dl_data = _json.loads(r2.read())
                href = dl_data.get('href', '')
                if href:
                    proxy_key = f"{art_base}/{name}"
                    _photo_url_cache[proxy_key] = href
                    photos.append(f"/api/photo-proxy/{urllib.parse.quote(art_base)}/{urllib.parse.quote(name)}")
        result = {'photos': photos}
        _photo_cache[art_base] = result
        return jsonify(result)
    except Exception as e:
        return jsonify({'photos': [], 'error': str(e)})

@app.route('/api/photo-proxy/<art_base>/<filename>')
def photo_proxy(art_base, filename):
    import urllib.request, urllib.parse, json as _json
    key = f"{art_base}/{filename}"
    href = _photo_url_cache.get(key)
    if not href:
        folder_path = f"{YADISK_FOLDER}/{art_base}"
        dl_url = ("https://cloud-api.yandex.net/v1/disk/public/resources/download?"
            + urllib.parse.urlencode({'public_key': YADISK_PUBLIC_KEY, 'path': f"{folder_path}/{filename}"}))
        try:
            req = urllib.request.Request(dl_url, headers={'Authorization': f'OAuth {YADISK_TOKEN}'})
            with urllib.request.urlopen(req, timeout=8) as r:
                href = _json.loads(r.read()).get('href', '')
            _photo_url_cache[key] = href
        except:
            return 'Not found', 404
    try:
        req = urllib.request.Request(href)
        with urllib.request.urlopen(req, timeout=10) as r:
            data = r.read()
            ct = r.headers.get('Content-Type', 'image/jpeg')
        resp = Response(data, mimetype=ct)
        resp.headers['Cache-Control'] = 'public, max-age=86400'
        return resp
    except:
        return 'Error', 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
