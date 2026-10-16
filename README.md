# 📄 DocXL — Intelligent PDF to Excel Converter

DocXL is a web application designed to extract structured tabular data from PDF documents and convert it into Excel spreadsheets. It supports text-based PDF extraction and OCR-based processing for scanned documents.

The goal of DocXL is to simplify document processing by transforming unstructured PDF tables into organized, usable spreadsheet data.

---

## ✨ Features

- 📄 Upload PDF documents.
- 🔍 Extract tables from text-based PDFs.
- 🧠 OCR support for scanned PDF documents.
- 📊 Convert extracted tables into Excel spreadsheets.
- 📥 Download generated Excel files.
- 🔐 User registration and email/password authentication.
- 🖥️ Modern React-based user interface.
- ⚡ FastAPI backend for document processing.
- 📚 Interactive API documentation through Swagger UI.
- 🎨 Animated DocXL introduction page and hero section.

---

## 🛠️ Technology Stack

### Frontend
- React
- TypeScript
- Vite
- CSS

### Backend
- Python
- FastAPI
- Uvicorn
- PyMuPDF
- pdfplumber
- PaddleOCR
- OpenCV
- pytesseract
- Pillow
- pandas
- openpyxl

### Development Tools
- Visual Studio Code
- Git
- GitHub
- npm
- Python virtual environments

---

### Clone the Repository
git clone https://github.com/pujitha10-vavilala/DocXL.git
cd DocXL

## Set Up the Backend
cd Backend
python -m venv venv
.\venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
Start the Backend
From the Backend directory, run:
python -m uvicorn app.main:app --reload

The backend should be available at:
http://127.0.0.1:8000

## Set Up the Frontend
Open a second terminal.
Navigate to the frontend:
cd C:\Users\ftt\Documents\DocXL\Frontend

Install the dependencies:
npm install

6. Configure the Frontend
Check Frontend/.env.example for the API URL variable expected by the application.
Create a local frontend .env file if necessary and configure the backend URL.
For example, if the application uses VITE_API_URL:
VITE_API_URL=http://127.0.0.1:8000

Use the exact variable name expected by the frontend source code.
7. Run the Frontend
Start the Vite development server:
npm run dev

Open the local URL printed in the terminal. It is usually:
http://localhost:5173

Keep both frontend and backend terminals running while using the application.

---

### 📄 How to Use DocXL
1. Open DocXL in your browser.
2. Explore the animated introduction and enter the main application.
3. Register or log in using the authentication options available in the application.
4. Upload a PDF document.
5. Start the document extraction process.
6. Review the extracted table data.
7. Export the extracted data to Excel.
8. Download the generated spreadsheet.

## 📁 Project Structure

```text
DocXL/
│
├── Backend/
│   ├── app/
│   │   └── main.py
│   ├── uploads/
│   ├── generated/
│   ├── requirements.txt
│   ├── .env.example
│   └── README.md
│
├── Frontend/
│   ├── src/
│   │   ├── App.tsx
│   │   ├── main.tsx
│   │   └── styles.css
│   ├── index.html
│   ├── package.json
│   ├── package-lock.json
│   ├── vite.config.ts
│   ├── tsconfig.json
│   ├── .env.example
│   └── README.md
│
├── sample_documents/
├── tests/
├── .gitignore
└── README.md

