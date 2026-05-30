import os
import json
import bcrypt
import secrets
from datetime import datetime
from functools import wraps
from flask import (Flask, render_template, request, jsonify,
                   send_file, session, redirect, url_for)
import openpyxl
from generate_payslip import build_docx

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', secrets.token_hex(32))

_BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH  = os.environ.get("EXCEL_PATH", os.path.join(_BASE_DIR, "data-pribadi-pegawai.xlsx"))
OUTPUT_DIR  = os.path.join(_BASE_DIR, "output")
USERS_FILE  = os.path.join(_BASE_DIR, "users.json")
os.makedirs(OUTPUT_DIR, exist_ok=True)


# ── Auth helpers ──────────────────────────────────────────────────────────

def load_users():
    if not os.path.isfile(USERS_FILE):
        return {}
    with open(USERS_FILE, 'r') as f:
        return json.load(f)

def check_password(username, password):
    users = load_users()
    if username not in users:
        return False
    stored_hash = users[username]["password"].encode('utf-8')
    return bcrypt.checkpw(password.encode('utf-8'), stored_hash)

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get('logged_in'):
            if request.is_json:
                return jsonify({"ok": False, "error": "Unauthorized"}), 401
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated


# ── Auth routes ───────────────────────────────────────────────────────────

@app.route('/login', methods=['GET', 'POST'])
def login():
    if session.get('logged_in'):
        return redirect(url_for('index'))
    error = None
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '')
        if check_password(username, password):
            session['logged_in'] = True
            session['username']  = username
            return redirect(url_for('index'))
        else:
            error = 'Username atau password salah.'
    return render_template('login.html', error=error)

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))


# ── App routes ────────────────────────────────────────────────────────────

@app.route('/')
@login_required
def index():
    return render_template('index.html', username=session.get('username'))


@app.route('/api/employees')
@login_required
def api_employees():
    try:
        emps   = load_employees()
        result = [{
            "nama":                e["Nama"],
            "jabatan":             e["Jabatan"],
            "gaji_pokok":          e["Gaji Pokok"],
            "uang_makan_per_hari": e["uang makan 1 hari"],
            "komisi_coating":      e["komisi coating detailing"],
            "komisi_detailing":    e["komisi detailing only"],
            "komisi_maintenance":  e["komisi maintenance"],
        } for e in emps]
        return jsonify({"ok": True, "employees": result})
    except Exception as ex:
        return jsonify({"ok": False, "error": str(ex)}), 500


@app.route('/api/generate', methods=['POST'])
@login_required
def api_generate():
    data = request.get_json()
    try:
        emps    = load_employees()
        emp_map = {e["Nama"]: e for e in emps}
        nama    = data["nama"]
        if nama not in emp_map:
            return jsonify({"ok": False, "error": f"Employee '{nama}' not found"}), 404

        emp             = emp_map[nama]
        period          = data.get("period", current_period())
        num_coating     = int(data.get("num_coating", 0))
        num_detailing   = int(data.get("num_detailing", 0))
        num_maintenance = int(data.get("num_maintenance", 0))
        uang_makan_days = int(data.get("uang_makan_days", 0))
        thr             = float(data.get("thr", 0))
        kasbon          = float(data.get("kasbon", 0))

        safe_name = nama.replace(" ", "_")
        filename  = f"Payslip_{safe_name}_{period.replace(' ', '_')}.docx"
        out_path  = os.path.join(OUTPUT_DIR, filename)

        build_docx(emp, period, num_coating, num_detailing, num_maintenance,
                   uang_makan_days, thr, kasbon, out_path)

        return jsonify({"ok": True, "filename": filename})
    except Exception as ex:
        import traceback
        return jsonify({"ok": False, "error": str(ex),
                        "trace": traceback.format_exc()}), 500


@app.route('/api/download/<filename>')
@login_required
def api_download(filename):
    path = os.path.join(OUTPUT_DIR, filename)
    if not os.path.isfile(path):
        return "File not found", 404
    return send_file(path, as_attachment=True, download_name=filename)


@app.route('/api/period')
@login_required
def api_period():
    return jsonify({"period": current_period()})


# ── Employee loader ───────────────────────────────────────────────────────

def load_employees():
    wb     = openpyxl.load_workbook(EXCEL_PATH, read_only=True, data_only=True)
    ws     = wb["Data"]
    rows   = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    header = [str(c).strip() if c else "" for c in rows[0]]
    skip   = {"rekening listrik", "rekening pdam", "indihome akun"}
    result = []
    for row in rows[1:]:
        if not row[0] or str(row[0]).strip() in ("", " "):
            continue
        name = str(row[0]).strip()
        if name.lower() in skip:
            continue
        emp = dict(zip(header, row))
        for field in ["Gaji Pokok", "uang makan 1 hari",
                      "komisi coating detailing", "komisi detailing only",
                      "komisi maintenance"]:
            try:
                emp[field] = float(emp.get(field) or 0)
            except (TypeError, ValueError):
                emp[field] = 0.0
        emp["Jabatan"] = str(emp.get("Jabatan") or "").strip()
        result.append(emp)
    wb.close()
    return result


def current_period():
    now = datetime.now()
    months_id = ["Januari","Februari","Maret","April","Mei","Juni",
                 "Juli","Agustus","September","Oktober","November","Desember"]
    return f"{months_id[now.month - 1]} {now.year}"


if __name__ == "__main__":
    if not os.path.isfile(USERS_FILE):
        print("[WARNING] users.json not found — run create_user.py first to add a user.")
    print(f"[Payslip App] Excel:   {os.path.abspath(EXCEL_PATH)}")
    print(f"[Payslip App] Output:  {OUTPUT_DIR}")
    port = int(os.environ.get('PORT', 5000))
    debug = os.environ.get('FLASK_ENV') == 'development'
    print(f"[Payslip App] Starting on port {port} (debug={debug})")
    app.run(debug=debug, port=port, host='0.0.0.0')
