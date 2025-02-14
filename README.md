JSON to Excel Converter
A modern, client‑side web application that converts JSON data extracted from TXT, DOCX, or PDF files into a structured Excel file. Built using React, Vite, and Tailwind CSS, this tool leverages modern libraries such as XLSX, Mammoth, and pdfjs‑dist to deliver a seamless and engaging user experience—all entirely in the browser.

Features
File Parsing: Extract JSON data from text, Word, and PDF files.
Excel Generation: Automatically converts JSON data into a multi‑sheet Excel file using the XLSX library.
Dynamic UI: Features a sleek, Excel-inspired green mono color palette.
Interactive Background: The background geometry responds to cursor movement, creating a parallax effect.
Animated Watermark: Multi‑line watermark text scrolls across the background for a unique, branded look.
Client‑Side Only: All processing happens in the user's browser—no server required.
Getting Started
Prerequisites
Node.js (v14 or above)
Yarn (or npm, if preferred)
Installation
Clone the Repository

bash
Copy
Edit
git clone https://github.com/yourusername/json-to-excel-converter.git
cd json-to-excel-converter
Install Dependencies

bash
Copy
Edit
yarn
Run the Development Server

bash
Copy
Edit
yarn dev
The app should open automatically in your browser (default: http://localhost:3000). If it doesn't, navigate to the URL manually.

Building for Production
To generate an optimized production build, run:

bash
Copy
Edit
yarn build
The output will be placed in the dist folder, which you can then deploy to any static hosting provider (e.g., GitHub Pages, Netlify, Vercel).