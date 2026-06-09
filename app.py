from flask import Flask, request, jsonify, send_file, render_template, send_from_directory
import sqlite3, os, json
from datetime import datetime
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

ALLOWED = {'.xlsx', '.xls', '.csv'}

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

init_db()

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/orders', methods=['GET'])
def get_orders():
    branch = request.args.get('branch', '')
    status = request.args.get('status', '')
    db = get_db()
    q = 'SELECT * FROM orders WHERE 1=1'
    params = []
    if branch:
        q += ' AND branch=?'; params.append(branch)
    if status:
        q += ' AND status=?'; params.append(status)
    q += ' ORDER BY created_at DESC'
    rows = db.execute(q, params).fetchall()
    db.close()
    return jsonify([dict(r) for r in rows])

@app.route('/api/orders', methods=['POST'])
def create_order():
    branch = request.form.get('branch', '').strip()
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
def download(oid):
    db = get_db()
    row = db.execute('SELECT * FROM orders WHERE id=?', (oid,)).fetchone()
    db.close()
    if not row or not row['filename']:
        return 'Not found', 404
    return send_from_directory(app.config['UPLOAD_FOLDER'], row['filename'],
                               as_attachment=True, download_name=row['original_name'])

@app.route('/api/stats')
def stats():
    db = get_db()
    total = db.execute('SELECT COUNT(*) FROM orders').fetchone()[0]
    new = db.execute("SELECT COUNT(*) FROM orders WHERE status='Новая'").fetchone()[0]
    inprog = db.execute("SELECT COUNT(*) FROM orders WHERE status='В работе'").fetchone()[0]
    done = db.execute("SELECT COUNT(*) FROM orders WHERE status='Выполнена'").fetchone()[0]
    db.close()
    return jsonify({'total': total, 'new': new, 'in_progress': inprog, 'done': done})

if __name__ == '__main__':
    app.run(debug=True, port=5000)
