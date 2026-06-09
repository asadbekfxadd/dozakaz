from flask import Flask, request, jsonify, render_template, send_from_directory, session, redirect
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
    f = request.files.get('file')
    if f and f.filename:
        ext = os.path.splitext(f.filename)[1].lower()
        if ext not in ALLOWED:
            return jsonify({'error': 'Только .xlsx, .xls, .csv'}), 400
        original_name = f.filename
        filename = f'{datetime.now().strftime("%Y%m%d_%H%M%S")}_{secure_filename(f.filename)}'
        f.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
    db = get_db()
    cur = db.execute(
        'INSERT INTO orders (branch,responsible,date,priority,note,filename,original_name) VALUES (?,?,?,?,?,?,?)',
        (branch, responsible, date, priority, note, filename, original_name)
    )
    order_id = cur.lastrowid
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

@app.route('/api/orders/<int:oid>/delete', methods=['DELETE'])
@admin_required
def delete_order(oid):
    db = get_db()
    row = db.execute('SELECT filename FROM orders WHERE id=?', (oid,)).fetchone()
    if row and row['filename']:
        path = os.path.join(app.config['UPLOAD_FOLDER'], row['filename'])
        if os.path.exists(path):
            os.remove(path)
    db.execute('DELETE FROM orders WHERE id=?', (oid,))
    db.commit()
    db.close()
    return jsonify({'ok': True})

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

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)