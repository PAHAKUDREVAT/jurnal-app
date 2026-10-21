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
IP_FILE = os.path.join(BASE_DIR, 'IP_адрес_для_входа.txt')
app = Flask(__name__)
def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS requests (
        id INTEGER PRIMARY KEY AUTOINCREMENT, created_at TEXT NOT NULL, created_by TEXT NOT NULL,
        phone TEXT, horizon TEXT, area TEXT, safety TEXT, description TEXT, executor TEXT, 
        priority TEXT, status TEXT, comment TEXT, updated_at TEXT)''')
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
def write_ip_file(ip):
    """Создает текстовый файл с адресом, чтобы сотрудники знали куда заходить"""
    try:
        with open(IP_FILE, "w", encoding="utf-8") as f:
            f.write("=== ЖУРНАЛ ЗАЯВОК МАРКШЕЙДЕРСКОЙ СЛУЖБЫ ===\n\n")
            f.write(f"Адрес для входа с других компьютеров и телефонов:\n")
            f.write(f"http://{ip}:5000\n\n")
            f.write("ВНИМАНИЕ:\n")
            f.write("1. Программа уже работает в фоне (ищите значок в трее возле часов).\n")
            f.write("2. НЕ ЗАКРЫВАЙТЕ программу, пока работаете с журналом.\n")
            f.write("3. Остальные сотрудники просто вводят этот адрес в браузере.\n")
    except Exception:
        pass
# --- ШАБЛОНЫ ---
BASE_HTML = """
<!DOCTYPE html><html lang="ru"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Журнал заявок</title>
<style>
body{font-family:Arial,sans-serif;background:#f4f6f8;margin:0;color:#222;padding:20px}
.card{background:#fff;padding:20px;border-radius:8px;box-shadow:0 2px 5px rgba(0,0,0,0.1);max-width:1600px;margin:0 auto}
table{width:100%;border-collapse:collapse;margin-top:15px;font-size:12px;table-layout:fixed}
th,td{padding:6px;border-bottom:1px solid #eee;text-align:left;word-wrap:break-word;vertical-align:top;}
th{background:#f8f9fa} input,select,textarea{width:100%;padding:8px;margin:5px 0;box-sizing:border-box;font-family:inherit;font-size:13px;}
.btn{background:#1e3a5f;color:#fff;padding:10px;border:none;border-radius:4px;cursor:pointer;width:100%;margin-top:10px;font-weight:bold;}
.btn:hover{background:#16304e;}
.high{background-color:#ffe6e6;color:#b03030;font-weight:bold;}
.col-form{display:flex;gap:10px;margin-bottom:10px;}
.col{flex:1;}
label{font-size:12px;color:#555;font-weight:bold;}
a{color:#1e3a5f;}
</style></head><body><div class="card">
<h2 style="margin-top:0;">Журнал заявок маркшейдерской службы</h2>
{% block content %}{% endblock %}
</div></body></html>
"""
FORM_HTML = """
<form method="POST" action="{{ action }}">
<div class="col-form">
    <div class="col">
        <label>Подал (ФИО):</label>
        <input name="author" required value="{{ r.created_by if r else '' }}">
        <label>Телефон:</label>
        <input name="phone" value="{{ r.phone if r else '' }}">
        <label>Дата подачи (не меняется):</label>
        <input value="{{ r.created_at if r else now }}" readonly style="background:#eee;">
    </div>
    <div class="col">
        <label>Горизонт:</label>
        <input name="horizon" required value="{{ r.horizon if r else '' }}">
        <label>Участок:</label>
        <input name="area" value="{{ r.area if r else '' }}">
        <label>Безопасность места работ:</label>
        <select name="safety">
            <option value="Да" {% if r and r.safety == 'Да' %}selected{% endif %}>Да (обезопасено)</option>
            <option value="Нет" {% if r and r.safety == 'Нет' %}selected{% endif %}>Нет (опасно)</option>
            <option value="Не требуется" {% if r and r.safety == 'Не требуется' %}selected{% endif %}>Не требуется</option>
        </select>
    </div>
    <div class="col">
        <label>Приоритет:</label>
        <select name="priority">
            <option value="Низкий" {% if r and r.priority == 'Низкий' %}selected{% endif %}>Низкий</option>
            <option value="Обычный" {% if not r or r.priority == 'Обычный' %}selected{% endif %}>Обычный</option>
            <option value="Высокий" {% if r and r.priority == 'Высокий' %}selected{% endif %}>Высокий</option>
            <option value="Срочный" {% if r and r.priority == 'Срочный' %}selected{% endif %}>Срочный</option>
        </select>
        <label>Исполнитель (кто выполнил):</label>
        <input name="executor" value="{{ r.executor if r else '' }}">
        <label>Статус (меняет маркшейдер):</label>
        <select name="status">
            <option value="Новая" {% if not r or r.status == 'Новая' %}selected{% endif %}>Новая</option>
            <option value="Принята" {% if r and r.status == 'Принята' %}selected{% endif %}>Принята</option>
            <option value="В работе" {% if r and r.status == 'В работе' %}selected{% endif %}>В работе</option>
            <option value="Выполнена" {% if r and r.status == 'Выполнена' %}selected{% endif %}>Выполнена</option>
            <option value="Отклонена" {% if r and r.status == 'Отклонена' %}selected{% endif %}>Отклонена</option>
        </select>
    </div>
</div>
<label>Описание задачи:</label>
<textarea name="desc" rows="2">{{ r.description if r else '' }}</textarea>
<label>Комментарий:</label>
<textarea name="comment" rows="1">{{ r.comment if r else '' }}</textarea>
<button type="submit" class="btn">{{ btn_text }}</button>
{% if r %}<a href="/" style="display:block;text-align:center;margin-top:10px;font-size:14px;">Отмена</a>{% endif %}
</form>
"""
INDEX_HTML = BASE_HTML.replace("{% block content %}{% endblock %}", """
""" + FORM_HTML + """
<hr style="margin:20px 0; border:0; border-top:1px solid #ddd;">
<table>
<tr>
    <th style="width:3%">№</th>
    <th style="width:8%">Дата и время</th>
    <th style="width:8%">Подал</th>
    <th style="width:7%">Телефон</th>
    <th style="width:8%">Горизонт</th>
    <th style="width:8%">Участок</th>
    <th style="width:8%">Безопасность</th>
    <th style="width:7%">Приоритет</th>
    <th style="width:8%">Статус</th>
    <th style="width:8%">Исполнитель</th>
    <th style="width:12%">Описание</th>
    <th style="width:8%">Коммент.</th>
    <th style="width:7%">Действие</th>
</tr>
{% for req in requests %}
<!-- Если приоритет высокий И статус НЕ выполнен/отклонен, то красим в красный -->
<tr class="{% if req[10] in ['Высокий','Срочный'] and req[11] not in ['Выполнена', 'Отклонена'] %}high{% endif %}">
    <td>{{req[0]}}</td>
    <td>{{req[1]}}</td>
    <td>{{req[2]}}</td>
    <td>{{req[3]}}</td>
    <td>{{req[4]}}</td>
    <td>{{req[5]}}</td>
    <td>{{req[6]}}</td>
    <td>{{req[10]}}</td>
    <td><b>{{req[11]}}</b></td>
    <td>{{req[8]}}</td>
    <td>{{req[7]}}</td>
    <td>{{req[12]}}</td>
    <td><a href="/edit/{{req[0]}}">Ред.</a></td>
</tr>
{% endfor %}
</table>
<div style="margin-top:20px;">
    <a href="/export" style="background:#2a6e2a;color:#fff;padding:10px 15px;text-decoration:none;border-radius:4px;font-weight:bold;">Скачать Excel (для Р7)</a>
</div>
""")
EDIT_PAGE = BASE_HTML.replace("{% block content %}{% endblock %}", """
<h3>Редактирование заявки №{{ r[0] }}</h3>
""" + FORM_HTML.replace("{{ action }}", "/update/{{ r[0] }}"))
# --- МАРШРУТЫ ---
@app.route('/')
def index():
    conn = sqlite3.connect(DB_PATH)
    rows = conn.execute("SELECT id,created_at,created_by,phone,horizon,area,safety,description,executor,0,priority,status,comment FROM requests ORDER BY id DESC").fetchall()
    conn.close()
    now = datetime.now().strftime("%d.%m.%Y %H:%M:%S")
    html = INDEX_HTML.replace("{{ action }}", "/create").replace("{{ btn_text }}", "Подать заявку").replace("{{ now }}", now)
    html = html.replace('value=""', 'value=""').replace('{% if r %}', '{% if False %}') 
    return render_template_string(html, requests=rows, r=None, now=now)
@app.route('/create', methods=['POST'])
def create():
    now = datetime.now().strftime("%d.%m.%Y %H:%M:%S")
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""INSERT INTO requests (created_at,created_by,phone,horizon,area,safety,priority,description,comment,executor,status,updated_at) 
                 VALUES (?,?,?,?,?,?,?,?,?,?,?,?)""",
                 (now, request.form['author'], request.form.get('phone',''), request.form['horizon'], request.form.get('area',''),
                  request.form['safety'], request.form['priority'], request.form.get('desc',''), request.form.get('comment',''),
                  request.form.get('executor',''), 'Новая', now))
    conn.commit(); conn.close()
    return redirect(url_for('index'))
@app.route('/edit/<int:id>')
def edit(id):
    conn = sqlite3.connect(DB_PATH)
    r = conn.execute("SELECT id,created_at,created_by,phone,horizon,area,safety,description,executor,0,priority,status,comment FROM requests WHERE id=?", (id,)).fetchone()
    conn.close()
    
    class Req: pass
    req = Req()
    req.id, req.created_at, req.created_by, req.phone = r[0], r[1], r[2], r[3]
    req.horizon, req.area, req.safety, req.description = r[4], r[5], r[6], r[7]
    req.executor, req.priority, req.status, req.comment = r[8], r[10], r[11], r[12]
    
    html = EDIT_PAGE.replace("{{ action }}", f"/update/{id}").replace("{{ btn_text }}", "Сохранить изменения")
    return render_template_string(html, r=req, now=req.created_at)
@app.route('/update/<int:id>', methods=['POST'])
def update(id):
    now = datetime.now().strftime("%d.%m.%Y %H:%M:%S")
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""UPDATE requests SET created_by=?, phone=?, horizon=?, area=?, safety=?, 
                 priority=?, description=?, comment=?, executor=?, status=?, updated_at=? WHERE id=?""",
                 (request.form['author'], request.form.get('phone',''), request.form['horizon'], request.form.get('area',''),
                  request.form['safety'], request.form['priority'], request.form.get('desc',''), request.form.get('comment',''),
                  request.form.get('executor',''), request.form['status'], now, id))
    conn.commit(); conn.close()
    return redirect(url_for('index'))
@app.route('/export')
def export():
    conn = sqlite3.connect(DB_PATH)
    rows = conn.execute("SELECT id,created_at,created_by,phone,horizon,area,safety,priority,status,executor,description,comment FROM requests ORDER BY id").fetchall()
    conn.close()
    wb = Workbook(); ws = wb.active
    ws.append(["№","Дата","Подал","Телефон","Горизонт","Участок","Безопасность","Приоритет","Статус","Исполнитель","Описание","Комментарий"])
    for r in rows: ws.append(list(r))
    buf = BytesIO(); wb.save(buf); buf.seek(0)
    return send_file(buf, as_attachment=True, download_name="Журнал.xlsx")
def open_browser():
    import time; time.sleep(1.5); webbrowser.open("http://127.0.0.1:5000")
if __name__ == '__main__':
    init_db()
    ip = get_local_ip()
    write_ip_file(ip)  # Создаем текстовый файл с IP-адресом
    print(f"Сервер запущен: http://{ip}:5000")
    threading.Thread(target=open_browser, daemon=True).start()
    app.run(host='0.0.0.0', port=5000, debug=False, use_reloader=False)
