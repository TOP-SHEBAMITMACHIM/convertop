from flask import Flask, render_template, request, send_file, jsonify
import os
import sys
import subprocess
import json
from PIL import Image
import img2pdf
from tkinter import filedialog
import tkinter as tk
import webview
import tempfile
from threading import Timer

from pdf2docx import Converter
from docx2pdf import convert as docx2pdf_convert
import pythoncom
import win32com.client

if getattr(sys, 'frozen', False):
    template_folder = os.path.join(sys._MEIPASS, 'templates')
    static_folder = os.path.join(sys._MEIPASS, 'static')
    app = Flask(__name__, template_folder=template_folder, static_folder=static_folder)
    exe_dir = os.path.dirname(sys.executable)
else:
    app = Flask(__name__)
    exe_dir = os.path.dirname(os.path.abspath(__file__))

UPLOAD_FOLDER = os.path.join(tempfile.gettempdir(), 'Convertop_Workspace')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# --- מערכת שמירת הגדרות חכמה ---
CONFIG_PATH = os.path.join(os.path.expanduser("~"), "convertop_config.json")

def load_config():
    default_config = {
        "target_dir": "",
        "theme": "system",
        "sound": "none",
        "doc_engine": "msword"
    }
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
                loaded = json.load(f)
                default_config.update(loaded)
        except:
            pass
    return default_config

def save_config(config):
    try:
        with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=4)
    except:
        pass

app_settings = load_config()

def ms_word_convert(input_path, output_path, to_pdf=True):
    pythoncom.CoInitialize()
    word = None
    doc = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0 
        
        doc = word.Documents.Open(input_path, ConfirmConversions=False)
        if to_pdf:
            doc.SaveAs(output_path, FileFormat=17, UseISO19005_1=True)
        else:
            doc.SaveAs(output_path, FileFormat=16)
    finally:
        if doc: doc.Close(False)
        if word: word.Quit()
        pythoncom.CoUninitialize()

@app.route('/')
def index():
    short_name = os.path.basename(app_settings["target_dir"]) if app_settings["target_dir"] else ""
    return render_template('index.html', settings=app_settings, short_name=short_name)

@app.route('/save_setting', methods=['POST'])
def save_setting():
    """ מקבל שמירות חדשות מהממשק ושומר בקובץ """
    data = request.json
    if 'key' in data and 'value' in data:
        app_settings[data['key']] = data['value']
        save_config(app_settings)
    return jsonify({"status": "success"})

@app.route('/select_destination', methods=['POST'])
def select_destination():
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    folder = filedialog.askdirectory()
    root.destroy()
    
    if folder:
        app_settings["target_dir"] = folder
        save_config(app_settings) # שומר מיד
        return jsonify({"status": "success", "folder": folder, "short_name": os.path.basename(folder)})
    return jsonify({"status": "cancelled"})

@app.route('/reset_destination', methods=['POST'])
def reset_destination():
    app_settings["target_dir"] = ""
    save_config(app_settings) # שומר מיד
    return jsonify({"status": "success"})

@app.route('/convert', methods=['POST'])
def convert():
    if 'file' not in request.files:
        return jsonify({'error': 'לא נבחר קובץ'}), 400
    
    file = request.files['file']
    target_format = request.form.get('format', 'mp4')
    width = request.form.get('width', '')
    doc_engine = request.form.get('doc_engine', 'msword')

    if file.filename == '':
        return jsonify({'error': 'לא נבחר קובץ'}), 400

    base_name, original_ext = os.path.splitext(file.filename)
    original_ext = original_ext.lower()
    input_path = os.path.join(UPLOAD_FOLDER, file.filename)
    output_filename = f"{base_name}.{target_format}"
    
    save_locally = False
    if app_settings["target_dir"]:
        output_path = os.path.join(app_settings["target_dir"], output_filename)
        save_locally = True
    else:
        output_path = os.path.join(UPLOAD_FOLDER, output_filename)
    
    file.save(input_path)

    try:
        if target_format in ["mp4", "avi", "mkv", "mp3", "wav", "flac"]:
            ffmpeg_path = os.path.join(exe_dir, 'ffmpeg.exe')
            if not os.path.exists(ffmpeg_path):
                raise Exception("קובץ ffmpeg.exe חסר! אנא שים אותו בתיקיית התוכנה הראשית.")
            subprocess.run(f'"{ffmpeg_path}" -i "{input_path}" "{output_path}" -y', 
                           shell=True, check=True, creationflags=0x08000000)
                           
        elif target_format == "docx":
            if original_ext != ".pdf":
                raise Exception("ניתן להמיר ל-DOCX רק קבצי PDF.")
            if doc_engine == "msword":
                ms_word_convert(input_path, output_path, to_pdf=False)
            else:
                cv = Converter(input_path)
                cv.convert(output_path)
                cv.close()
            
        elif target_format == "pdf":
            if original_ext in [".docx", ".doc"]:
                if doc_engine == "msword":
                    ms_word_convert(input_path, output_path, to_pdf=True)
                else:
                    docx2pdf_convert(input_path, output_path)
            else:
                with open(output_path, "wb") as f:
                    f.write(img2pdf.convert(input_path))
                    
        elif target_format == "ico":
            img = Image.open(input_path)
            img.save(output_path, format='ICO', sizes=[(256, 256)])
            
        else:
            img = Image.open(input_path)
            if width and width.isdigit():
                w = int(width)
                h_size = int((w / float(img.size[0])) * float(img.size[1]))
                img = img.resize((w, h_size), Image.Resampling.LANCZOS)
            if target_format == "jpg" and img.mode in ("RGBA", "P"):
                img = img.convert("RGB")
            img.save(output_path)

        if os.path.exists(input_path):
            os.remove(input_path)

        if save_locally:
            return jsonify({'status': 'saved_locally', 'path': output_path})
        else:
            return send_file(output_path, as_attachment=True)

    except Exception as e:
        if os.path.exists(input_path):
            os.remove(input_path)
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    window = webview.create_window('CONVERTOP - Studio Edition', app, width=1150, height=750, min_size=(900, 600))
    webview.start()