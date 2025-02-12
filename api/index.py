from flask import Flask, request, send_file
import os
import json
import pandas as pd
from werkzeug.utils import secure_filename
from docx import Document
import PyPDF2

app = Flask(__name__)

UPLOAD_FOLDER = "/tmp"  # Use /tmp for temporary storage in serverless environment
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def extract_json(file_path):
    ext = os.path.splitext(file_path)[-1].lower()
    
    with open(file_path, "rb") as file:
        if ext == ".txt":
            return json.load(file)
        elif ext == ".docx":
            doc = Document(file)
            text = "\n".join([para.text for para in doc.paragraphs])
            return json.loads(text)
        elif ext == ".pdf":
            reader = PyPDF2.PdfReader(file)
            text = "\n".join([page.extract_text() for page in reader.pages if page.extract_text()])
            return json.loads(text)
        else:
            raise ValueError("Unsupported file type")

def json_to_excel(json_data, output_path):
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        for sheet_name, records in json_data.items():
            df = pd.DataFrame(records if isinstance(records, list) else [records])
            df = df.applymap(lambda x: json.dumps(x) if isinstance(x, (dict, list)) else x)
            df.to_excel(writer, sheet_name=sheet_name, index=False)

@app.route("/api/upload", methods=["POST"])
def upload_file():
    if 'file' not in request.files:
        return {"error": "No file uploaded"}, 400
    
    file = request.files['file']
    if file.filename == "":
        return {"error": "No selected file"}, 400
    
    filename = secure_filename(file.filename)
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(file_path)
    
    try:
        json_data = extract_json(file_path)
        excel_path = file_path.replace(os.path.splitext(file_path)[1], ".xlsx")
        json_to_excel(json_data, excel_path)
        return send_file(excel_path, as_attachment=True)
    except Exception as e:
        return {"error": str(e)}, 500

# Vercel handler
def handler(request, context):
    return app(request.environ, context)
