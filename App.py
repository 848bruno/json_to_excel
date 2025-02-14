from flask import Flask, request, render_template, send_file
import os
import json
import pandas as pd
from werkzeug.utils import secure_filename
from docx import Document
import PyPDF2

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def extract_json_from_txt(file_path):
    with open(file_path, "r", encoding="utf-8") as file:
        return json.load(file)

def extract_json_from_docx(file_path):
    doc = Document(file_path)
    text = "\n".join([para.text for para in doc.paragraphs])
    return json.loads(text)

def extract_json_from_pdf(file_path):
    with open(file_path, "rb") as file:
        reader = PyPDF2.PdfReader(file)
        text = "\n".join([page.extract_text() for page in reader.pages if page.extract_text()])
    return json.loads(text)

def json_to_excel(json_data, output_path):
    """
    Converts JSON data to an Excel file.
    Supports different JSON formats:
      - A single top-level key with a list value -> one sheet.
      - A single top-level key with a dict value where each value is a list/dict -> multiple sheets.
      - Multiple top-level keys -> each becomes a sheet.
      - A top-level list -> one sheet (named 'Sheet1').
    """
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        if isinstance(json_data, dict):
            if len(json_data) == 1:
                # Only one top-level key
                main_key = list(json_data.keys())[0]
                inner_data = json_data[main_key]
                if isinstance(inner_data, list):
                    # Case: {"data": [ ... ]}
                    df = pd.DataFrame(inner_data)
                    df = df.astype(str)
                    df.to_excel(writer, sheet_name=main_key[:31], index=False)
                elif isinstance(inner_data, dict):
                    # Check if inner_data's values are all lists or dicts (i.e., multi-sheet)
                    if all(isinstance(v, (list, dict)) for v in inner_data.values()):
                        for sheet_name, records in inner_data.items():
                            if isinstance(records, list):
                                df = pd.DataFrame(records)
                            elif isinstance(records, dict):
                                df = pd.DataFrame([records])
                            else:
                                df = pd.DataFrame([{'value': records}])
                            df = df.astype(str)
                            df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
                    else:
                        # Otherwise treat as single sheet with one record
                        df = pd.DataFrame([inner_data])
                        df = df.astype(str)
                        df.to_excel(writer, sheet_name=main_key[:31], index=False)
                else:
                    # inner_data is neither list nor dict
                    df = pd.DataFrame([{'value': inner_data}])
                    df = df.astype(str)
                    df.to_excel(writer, sheet_name=main_key[:31], index=False)
            else:
                # Multiple top-level keys; each key becomes a sheet.
                for sheet_name, records in json_data.items():
                    if isinstance(records, list):
                        df = pd.DataFrame(records)
                    elif isinstance(records, dict):
                        df = pd.DataFrame([records])
                    else:
                        df = pd.DataFrame([{'value': records}])
                    df = df.astype(str)
                    df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
        elif isinstance(json_data, list):
            # Top-level list; create a default sheet
            df = pd.DataFrame(json_data)
            df = df.astype(str)
            df.to_excel(writer, sheet_name='Sheet1', index=False)
        else:
            # Fallback: wrap in list and output in one sheet.
            df = pd.DataFrame([{'value': json_data}])
            df = df.astype(str)
            df.to_excel(writer, sheet_name='Sheet1', index=False)

def process_file(file_path):
    ext = os.path.splitext(file_path)[-1].lower()
    if ext == ".txt":
        json_data = extract_json_from_txt(file_path)
    elif ext == ".docx":
        json_data = extract_json_from_docx(file_path)
    elif ext == ".pdf":
        json_data = extract_json_from_pdf(file_path)
    else:
        raise ValueError("Unsupported file type")
    
    # Determine output Excel file name.
    if isinstance(json_data, dict) and len(json_data) == 1:
        main_key = list(json_data.keys())[0]
        output_excel = os.path.join(os.path.dirname(file_path), f"{main_key}.xlsx")
    else:
        output_excel = file_path.replace(ext, ".xlsx")
    
    json_to_excel(json_data, output_excel)
    return output_excel

@app.route("/", methods=["GET", "POST"])
def upload_file():
    if request.method == "POST":
        if 'file' not in request.files:
            return "No file uploaded"
        
        file = request.files['file']
        if file.filename == "":
            return "No selected file"
        
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)
        
        try:
            excel_path = process_file(file_path)
            return send_file(excel_path, as_attachment=True)
        except Exception as e:
            return f"Error: {str(e)}"
    
    return render_template("upload.html")

if __name__ == "__main__":
    app.run(debug=True)
