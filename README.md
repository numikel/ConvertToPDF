# 🧾 ConvertToPdf

**ConvertToPdf** is a Windows-only Python script that automatically converts Microsoft Word, Excel, and PowerPoint documents in a folder to PDF format using installed Office applications.

## 🚀 Features

- ✅ Converts `.doc`, `.docx`, `.ppt`, `.pptx`, `.xls`, `.xlsx` to `.pdf`
- ✅ Uses Microsoft Office COM automation (no third-party libraries)
- ✅ Processes all files in a selected folder
- ✅ Displays progress using a progress bar (`tqdm`)
- ✅ Logs any errors to a `logs.txt` file

## ⚠️ Requirements

This script **requires Windows** and **Microsoft Office installed** (Word, Excel, PowerPoint).

Install required Python package:

```bash
pip install -r requirements.txt
```

## ▶️ How to Use
Run the script:

```bash
python convert_to_pdf.py
```

### You will be asked to:
Enter the path to the folder containing Office files

### The script will:
1) Find all supported Office files in the folder
2) Convert them one by one into PDFs
3) Save the PDFs in the same folder as the originals

## Project Structure
```bash
convert_to_pdf/
├── convert_to_pdf.py       # Main script
├── requirements.txt        # Python dependencies
├── README.md               # This file
└── .gitignore              # Ignore .pdf, logs, temp files
```

## 📌 To-do
1) Add support for subdirectories
2) GUI version using Tkinter or PyQt
3) Better error handling for corrupted files or permission issues

## Limitations
1) Only works on Windows
2) Requires local installation of MS Office (365, 2016, etc.)
3) Files must not be open in other programs during conversion

## 👤 Author
Made with ❤️ by Michał Kamiński

## 🧾 License
This project is licensed under the MIT License. Feel free to use, modify, and share it.