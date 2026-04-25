import os
import json
import uuid
import hashlib
import datetime
from flask import Flask, render_template, request, redirect, url_for, session, send_file, jsonify, flash
from werkzeug.utils import secure_filename
from werkzeug.security import generate_password_hash, check_password_hash
from docx_formatter import format_document

app = Flask(__name__)
app.secret_key = 'docuneat-secret-key-2024'
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

USERS_FILE = 'users.json'
HISTORY_FILE = 'history.json'

def load_users():
    if not os.path.exists(USERS_FILE):
        default = {'admin': generate_password_hash('admin123')}
        with open(USERS_FILE, 'w') as f:
            json.dump(default, f)
        return default
    with open(USERS_FILE) as f:
        return json.load(f)

def save_users(users):
    with open(USERS_FILE, 'w') as f:
        json.dump(users, f)

def load_history():
    if not os.path.exists(HISTORY_FILE):
        return []
    with open(HISTORY_FILE) as f:
        return json.load(f)

def save_history(history):
    with open(HISTORY_FILE, 'w') as f:
        json.dump(history, f, indent=2)

def add_history(user, original_name, output_name, options):
    history = load_history()
    history.insert(0, {
        'id': str(uuid.uuid4()),
        'user': user,
        'original': original_name,
        'output': output_name,
        'options': options,
        'timestamp': datetime.datetime.now().strftime('%d %b %Y, %H:%M')
    })
    save_history(history)

@app.route('/')
def index():
    if 'user' not in session:
        return redirect(url_for('login'))
    return redirect(url_for('index'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '')
        users = load_users()
        if username in users and check_password_hash(users[username], password):
            session['user'] = username
            return redirect(url_for('index'))
        return render_template('login.html', error='Username atau password salah.')
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

@app.route('/index')
def index():
    if 'user' not in session:
        return redirect(url_for('login'))
    history = load_history()
    user_history = [h for h in history if h['user'] == session['user']]
    return render_template('index.html', user=session['user'], history=user_history)

@app.route('/format', methods=['POST'])
def format_file():
    if 'user' not in session:
        return jsonify({'error': 'Unauthorized'}), 401
    
    file = request.files.get('file')
    if not file or not file.filename.endswith('.docx'):
        return jsonify({'error': 'Upload file .docx yang valid'}), 400

    options = {
        'font': request.form.get('font', 'Times New Roman'),
        'font_size': int(request.form.get('font_size', 12)),
        'line_spacing': request.form.get('line_spacing', '1.5'),
        'margin': request.form.get('margin', 'normal'),
        'fix_headings': request.form.get('fix_headings') == 'true',
        'fix_spacing': request.form.get('fix_spacing') == 'true',
        'fix_alignment': request.form.get('fix_alignment') == 'true',
        'add_page_numbers': request.form.get('add_page_numbers') == 'true',
        'page_number_position': request.form.get('page_number_position', 'bottom_center'),
        'remove_extra_blank': request.form.get('remove_extra_blank') == 'true',
    }

    original_name = secure_filename(file.filename)
    uid = uuid.uuid4().hex[:8]
    input_path = os.path.join(UPLOAD_FOLDER, f"{uid}_{original_name}")
    output_name = f"RAPIH_{uid}_{original_name}"
    output_path = os.path.join(OUTPUT_FOLDER, output_name)

    file.save(input_path)

    try:
        stats = format_document(input_path, output_path, options)
        add_history(session['user'], original_name, output_name, options)
        os.remove(input_path)
        return jsonify({
            'success': True,
            'output': output_name,
            'stats': stats
        })
    except Exception as e:
        if os.path.exists(input_path):
            os.remove(input_path)
        return jsonify({'error': f'Gagal memformat: {str(e)}'}), 500

@app.route('/download/<filename>')
def download(filename):
    if 'user' not in session:
        return redirect(url_for('login'))
    path = os.path.join(OUTPUT_FOLDER, secure_filename(filename))
    if not os.path.exists(path):
        flash('File tidak ditemukan.')
        return redirect(url_for('index'))
    return send_file(path, as_attachment=True)

@app.route('/history/delete/<item_id>', methods=['POST'])
def delete_history(item_id):
    if 'user' not in session:
        return jsonify({'error': 'Unauthorized'}), 401
    history = load_history()
    history = [h for h in history if not (h['id'] == item_id and h['user'] == session['user'])]
    save_history(history)
    return jsonify({'success': True})

@app.route('/history/delete-all', methods=['POST'])
def delete_all_history():
    if 'user' not in session:
        return jsonify({'error': 'Unauthorized'}), 401
    history = load_history()
    history = [h for h in history if h['user'] != session['user']]
    save_history(history)
    return jsonify({'success': True})

# Biarkan ini jika kamu masih ingin menjalankan di laptop sendiri (lokal)
if __name__ == '__main__':
    app.run(debug=True)

# WAJIB tambahkan ini di baris paling bawah (di luar blok if) agar terbaca oleh Vercel
app = app

import os

# Mencari lokasi folder api
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Gunakan ini untuk memanggil file JSON
USERS_FILE = os.path.join(BASE_DIR, 'users.json')
HISTORY_FILE = os.path.join(BASE_DIR, 'history.json')

app = Flask(__name__, template_folder='../templates')



