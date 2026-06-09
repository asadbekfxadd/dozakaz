from flask import Flask, request, jsonify, render_template, send_from_directory, session
import sqlite3, os
from datetime import datetime
from werkzeug.utils import secure_filename
from functools import wraps

app = Flask(__name__)
app.secret_key = 'dozakaz-secret-key-2026'
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

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
        size TEXT,
        qty INTEGER,
        FOREIGN KEY(order_id) REFERENCES orders(id)
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

@app.route('/api/orders', methods=['POST'])
@login_required
def create_order():
    branch = session['branch'] if session['role'] == 'user' else request.form.get('branch', '').strip()
    responsible = request.form.get('responsible', '').strip()
    date = request.form.get('date', datetime.now().strftime('%Y-%m-%d'))
    priority = request.form.get('priority', 'Обычный')
    note = request.form.get('note', '').strip()
    if not branch or not responsible:
        return jsonify({'error': 'Заполните обязательные поля'}), 400
    filename = None
    original_name = None
    items = []
    f = request.files.get('file')
    if f and f.filename:
        ext = os.path.splitext(f.filename)[1].lower()
        if ext not in ALLOWED:
            return jsonify({'error': 'Только .xlsx, .xls, .csv'}), 400
        original_name = f.filename
        filename = f'{datetime.now().strftime("%Y%m%d_%H%M%S")}_{secure_filename(f.filename)}'
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        f.save(filepath)
        # Parse items from Excel
        try:
            import openpyxl
            wb = openpyxl.load_workbook(filepath, data_only=True)
            ws = None
            for sname in wb.sheetnames:
                if 'WMS' in sname.upper() or 'СКЛАД' in sname.upper():
                    ws = wb[sname]
                    break
            if not ws:
                ws = wb.active
            header_row = None
            art_col = qty_col = size_col = None
            for i, row in enumerate(ws.iter_rows(values_only=True), 1):
                row_str = [str(c).strip() if c else '' for c in row]
                if 'Артикул' in row_str:
                    header_row = i
                    art_col = row_str.index('Артикул')
                    if 'Кол-во' in row_str:
                        qty_col = row_str.index('Кол-во')
                    if 'Характеристика' in row_str:
                        size_col = row_str.index('Характеристика')
                    break
            if header_row and art_col is not None and qty_col is not None:
                for row in ws.iter_rows(min_row=header_row+1, values_only=True):
                    art = str(row[art_col]).strip() if row[art_col] else ''
                    qty = row[qty_col]
                    size = str(row[size_col]).strip() if size_col is not None and row[size_col] else ''
                    if art and art != 'None' and qty and str(qty) != 'None':
                        try:
                            items.append({'article': art, 'size': size, 'qty': int(float(str(qty)))})
                        except:
                            pass
        except Exception as e:
            pass

    db = get_db()
    cur = db.execute(
        'INSERT INTO orders (branch,responsible,date,priority,note,filename,original_name) VALUES (?,?,?,?,?,?,?)',
        (branch, responsible, date, priority, note, filename, original_name)
    )
    order_id = cur.lastrowid
    for item in items:
        db.execute('INSERT INTO order_items (order_id,article,size,qty) VALUES (?,?,?,?)',
                   (order_id, item['article'], item['size'], item['qty']))
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

@app.route('/api/analytics/orders')
@admin_required
def analytics_orders():
    db = get_db()
    # Топ артикулов по заявкам
    top_articles = db.execute('''
        SELECT article, SUM(qty) as total_qty, COUNT(DISTINCT order_id) as order_count
        FROM order_items
        WHERE article != '' AND article != 'None'
        GROUP BY article
        ORDER BY total_qty DESC
        LIMIT 20
    ''').fetchall()
    # По филиалам
    branch_totals = db.execute('''
        SELECT o.branch, SUM(oi.qty) as total_qty, COUNT(DISTINCT o.id) as order_count
        FROM orders o
        JOIN order_items oi ON o.id = oi.order_id
        GROUP BY o.branch
        ORDER BY total_qty DESC
    ''').fetchall()
    db.close()
    return jsonify({
        'top_articles': [dict(r) for r in top_articles],
        'branch_totals': [dict(r) for r in branch_totals]
    })

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
        import io
        wb = openpyxl.load_workbook(io.BytesIO(f.read()), data_only=True)
        products = []
        for sname in wb.sheetnames:
            ws = wb[sname]
            art_col = sold_col = stock_col = cat_col = None
            header_row = None
            for i, row in enumerate(ws.iter_rows(values_only=True), 1):
                row_str = [str(c).strip() if c else '' for c in row]
                if any(x in row_str for x in ['ITEM NO.', 'Одежда', 'Обувь']):
                    header_row = i
                    for j, cell in enumerate(row_str):
                        if cell in ('ITEM NO.', 'Одежда', 'Обувь', 'Аксессуары'):
                            art_col = j
                        if cell == 'Продано':
                            sold_col = j
                        if cell == 'Остаток':
                            stock_col = j
                        if j == 1:
                            cat_col = j
                    break
            if header_row and art_col is not None:
                current_cat = 'APP'
                for row in ws.iter_rows(min_row=header_row+1, values_only=True):
                    if not row or not row[art_col]:
                        continue
                    art = str(row[art_col]).strip()
                    if art in ('Одежда', 'Обувь', 'Аксессуары', 'nan', ''):
                        if art == 'Обувь': current_cat = 'FTW'
                        elif art == 'Аксессуары': current_cat = 'ACC'
                        continue
                    cat = str(row[cat_col]).strip() if cat_col is not None and row[cat_col] else current_cat
                    sold = int(float(str(row[sold_col]))) if sold_col is not None and row[sold_col] and str(row[sold_col]) != 'None' else 0
                    stock = int(float(str(row[stock_col]))) if stock_col is not None and row[stock_col] and str(row[stock_col]) != 'None' else 0
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

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)