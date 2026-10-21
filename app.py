import sqlite3
import os
import sys
import socket
import webbrowser
import threading
from datetime import datetime
from io import BytesIO
from flask import Flask, request, redirect, url_for, render_template_string, send_file, jsonify
from openpyxl import Workbook
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, 'journal.db')
app = Flask(__name__)
def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS requests (
        id INTEGER PRIMARY KEY AUTOINCREMENT, created_at TEXT NOT NULL, created_by TEXT NOT NULL,
        phone TEXT, object_name TEXT NOT NULL, safety TEXT, description TEXT, comment TEXT,
        priority TEXT, status TEXT, assignee TEXT, updated_at TEXT)''')
    c.execute('''CREATE TRIGGER IF NOT EXISTS prevent_created_at_update BEFORE UPDATE ON requests
        FOR EACH ROW WHEN OLD.created_at <> NEW.created_at BEGIN SELECT RAISE(ABORT, 'immutable'); END''')
    c.execute('''CREATE TRIGGER IF NOT EXISTS prevent_delete BEFORE DELETE ON requests
        FOR EACH ROW BEGIN SELECT RAISE(ABORT, 'delete forbidden'); END''')
    conn.commit()
    conn.close()
def get_local_ip():
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try: s.connect(("8.8.8.8", 80)); return s.getsockname()[0]
    except: return "127.0.0.1"
    finally: s.close()
BASE_HTML = """
<!DOCTYPE html><html lang="ru"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Журнал заявок</title>
<style>
body{font-family:Arial,sans-serif;background:#f4f6f8;margin:0;color:#222;padding:20px}
.card{background:#fff;padding:20px;border-radius:8px;box-shadow:0 2px 5px rgba(0,0,0,0.1);max-width:1400px;margin:0 auto}
table{width:100%;border-collapse:collapse;margin-top:15px;font-size:13px;table-layout:fixed}
th,td{padding:8px;border-bottom:1px solid #eee;text-align:left;word-wrap:break-word}
th{background:#f8f9fa} input,select,textarea{width:100%;padding:8px;margin:5px 0;box-sizing:border-box}
.btn{background:#1e3a5f;color:#fff;padding:10px;border:none;border-radius:4px;cursor:pointer;width:100%;margin-top:10px}
.high{background-color:#fff0f0;color:#b03030}
</style></head><body><div class="card"><h2>Журнал заявок маркшейдерской службы</h2>
<form method="POST" action="/create">
<div style="display:flex;gap:10px"><div style="flex:1"><label>Подал:</label><input name="author" required>
<label>Телефон:</label><input name="phone"></div><div style="flex:1"><label>Объект:</label><input name="object" required>
<label>Обезопашено:</label><select name="safety"><option>Да</option><option>Нет</option></select></div>
<div style="flex:1"><label>Приоритет:</label><select name="priority"><option>Обычный</option><option>Высокий</option><option>Срочный</option></select>
<label>Комментарий:</label><input name="comment"></div></div>
<label>Описание:</label><textarea name="desc"></textarea>
<button type="submit" class="btn">Подать заявку</button></form><hr>
<table><tr><th style="width:5%">№</th><th style="width:12%">Дата</th><th style="width:10%">Подал</th><th style="width:10%">Тел</th><th style="width:10%">Объект</th><th style="width:8%">Обезоп.</th><th style="width:8%">Приоритет</th><th style="width:15%">Описание</th><th style="width:12%">Коммент.</th><th style="width:10%">Действие</th></tr>
{% for r in requests %}<tr class="{% if r[8] in ['Высокий','Срочный'] %}high{% endif %}">
<td>{{r[0]}}</td><td><b>{{r[1]}}</b></td><td>{{r[2]}}</td><td>{{r[3]}}</td><td>{{r[4]}}</td><td>{{r[5]}}</td><td>{{r[8]}}</td><td>{{r[6]}}</td><td>{{r[9]}}</td>
<td><a href="/edit/{{r[0]}}">Ред.</a></td></tr>{% endfor %}</table>
<a href="/export" style="display:inline-block;margin-top:15px;background:#2a6e2a;color:#fff;padding:10px;text-decoration:none;border-radius:4px">Экспорт в Excel</a>
</div></body></html>"""
EDIT_HTML = BASE_HTML.replace('<form method="POST" action="/create">', '<form method="POST" action="/update/{{r[0]}}">').replace('Подать заявку', 'Сохранить изменения')
@app.route('/')
def index():
    conn = sqlite3.connect(DB_PATH)
    rows = conn.execute("SELECT id,created_at,created_by,phone,object_name,safety,description,assignee,priority,comment FROM requests ORDER BY id DESC").fetchall()
    conn.close()
    return render_template_string(BASE_HTML, requests=rows)
@app.route('/create', methods=['POST'])
def create():
    now = datetime.now().strftime("%d.%m.%Y %H:%M:%S")
    conn = sqlite3.connect(DB_PATH)
    conn.execute("INSERT INTO requests (created_at,created_by,phone,object_name,safety,priority,description,comment,updated_at) VALUES (?,?,?,?,?,?,?,?,?)",
                 (now, request.form['author'], request.form.get('phone',''), request.form['object'], request.form['safety'], request.form['priority'], request.form.get('desc',''), request.form.get('comment',''), now))
    conn.commit(); conn.close()
    return redirect(url_for('index'))
@app.route('/edit/<int:id>')
def edit(id):
    conn = sqlite3.connect(DB_PATH)
    r = conn.execute("SELECT id,created_at,created_by,phone,object_name,safety,description,assignee,priority,comment FROM requests WHERE id=?", (id,)).fetchone()
    conn.close()
    return render_template_string(EDIT_HTML.replace('action="/create"', f'action="/update/{id}"'), requests=[r])
@app.route('/update/<int:id>', methods=['POST'])
def update(id):
    now = datetime.now().strftime("%d.%m.%Y %H:%M:%S")
    conn = sqlite3.connect(DB_PATH)
    conn.execute("UPDATE requests SET created_by=?, phone=?, object_name=?, safety=?, priority=?, description=?, comment=?, updated_at=? WHERE id=?",
                 (request.form['author'], request.form.get('phone',''), request.form['object'], request.form['safety'], request.form['priority'], request.form.get('desc',''), request.form.get('comment',''), now, id))
    conn.commit(); conn.close()
    return redirect(url_for('index'))
@app.route('/export')
def export():
    conn = sqlite3.connect(DB_PATH)
    rows = conn.execute("SELECT id,created_at,created_by,phone,object_name,safety,priority,description,comment FROM requests ORDER BY id").fetchall()
    conn.close()
    wb = Workbook(); ws = wb.active
    ws.append(["№","Дата","Подал","Телефон","Объект","Обезопашено","Приоритет","Описание","Комментарий"])
    for r in rows: ws.append(list(r))
    buf = BytesIO(); wb.save(buf); buf.seek(0)
    return send_file(buf, as_attachment=True, download_name="Журнал.xlsx")
def open_browser():
    import time; time.sleep(1.5); webbrowser.open("http://127.0.0.1:5000")
if __name__ == '__main__':
    init_db()
    print(f"Сервер запущен: http://{get_local_ip()}:5000")
    threading.Thread(target=open_browser, daemon=True).start()
    app.run(host='0.0.0.0', port=5000, debug=False, use_reloader=False)
