JSON to Excel Converter
This is a Flask web application that converts JSON data—extracted from TXT, DOCX, or PDF files—into a structured Excel file. The app automatically generates multiple Excel sheets based on the JSON structure and features a modern, Excel-inspired UI with a green monochrome theme and a subtle watermark design.

Features
File Upload: Accepts TXT, DOCX, and PDF files containing JSON data.
JSON Extraction: Extracts JSON data from the uploaded files.
Excel Generation: Converts the extracted JSON into an Excel file with multiple sheets (if needed).
Modern UI: A clean and attractive interface with an Excel-like green color palette and a decorative watermark.
Client-Friendly: Processes files on the server using Python Flask, Pandas, python-docx, and PyPDF2.
Technologies Used
Python Flask – Web framework for building the application.
Pandas – Data manipulation and Excel file creation.
python-docx – Extracts text from DOCX files.
PyPDF2 – Extracts text from PDF files.
Openpyxl – Used internally by Pandas to write Excel files.
Werkzeug – For secure file handling.
Installation
Prerequisites
Python 3.x installed
pip
Setup Instructions
Clone the Repository

bash
Copy
Edit
git clone https://github.com/yourusername/json-to-excel-converter.git
cd json-to-excel-converter
Create and Activate a Virtual Environment

bash
Copy
Edit
python -m venv venv
# On Windows:
venv\Scripts\activate
# On macOS/Linux:
source venv/bin/activate
Install Dependencies

bash
Copy
Edit
pip install -r requirements.txt
Run the Application

bash
Copy
Edit
python app.py
Open in Browser

Visit http://127.0.0.1:5000 in your web browser to use the app.

Usage
Upload a File:
On the homepage, select a TXT, DOCX, or PDF file that contains JSON data.

Conversion Process:
The app parses the file, extracts the JSON data, and converts it into an Excel file.

Download:
Once processing is complete, the generated Excel file is automatically served for download.

Project Structure
bash
Copy
Edit
json-to-excel-converter/
├── app.py                # Main Flask application
├── requirements.txt      # List of Python dependencies
├── templates/
│   └── upload.html       # HTML template for the file upload UI
└── uploads/              # Directory for storing uploaded files (created automatically)
Example JSON Data
Here’s an example JSON to test the application:

json
Copy
Edit
{
  "data": [
    {
      "id": 1,
      "input": {
        "username": "user1",
        "password": "password123",
        "email": "user1@example.com"
      },
      "expected_output": {
        "status": "success",
        "message": "User created successfully."
      },
      "type": "training"
    },
    {
      "id": 2,
      "input": {
        "username": "user2",
        "password": "wrongpassword",
        "email": "user2@example.com"
      },
      "expected_output": {
        "status": "error",
        "message": "Invalid credentials."
      },
      "type": "training"
    }
  ]
}
License
This project is licensed under the MIT License.

Acknowledgments
Thanks to the developers behind Flask, Pandas, python-docx, and PyPDF2 for their amazing libraries.
Inspired by the clean, modern design of Excel, with a custom green monochrome aesthetic and unique watermark.