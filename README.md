# Geam Digital - Multi-Ward Voter List & PDF to Excel Studio (2026 Edition) 📊

**Geam Digital** is an enterprise-grade document conversion platform specifically engineered to process Indian Election Commission Electoral Rolls (मतदाता सूची PDFs) as well as general tabular business documents into structured, formatted, formula-ready Excel (`.xlsx`) spreadsheets.

---

## 🗳️ Core Feature: Multi-Ward Voter List to Excel (मतदाता सूची ➔ 11-Column Multi-Sheet Excel)

Directly upload single or multiple Electoral Roll PDFs (e.g. all 9 wards of a Panchayat). The system automatically parses each ward and builds a comprehensive multi-sheet Excel workbook.

### 📋 11-Column Standard Format (ROW-Wise Sequence):
1. **भाग संख्या** (Part / Ward No.)
2. **क्रम संख्या** (Serial No. from 1 to Last without sequence gaps)
3. **नाम** (Voter Name in clean Devanagari Hindi via dual-engine OCR)
4. **पिता/पति का नाम** (Relative Name in clean Devanagari Hindi)
5. **आयु** (Age as numeric integer, 100% captured with zero missing values)
6. **मोबाइल नो** (**Header contains "मोबाइल नो", data cells kept clean and blank for field teams**)
7. **वोटर ID** (EPIC Number e.g. `NQU1571975`, `HRJ1220979`)
8. **हाउस नंबर** (House / Makan Number)
9. **एड्रेस** (Individual voter address: kept clean/empty if not explicitly printed on the individual voter card)
10. **बूथ का पता** (Polling Station / Booth Address)
11. **Status** (**Active** / **Deleted** with smart soft color coding)

---

### 📑 Multi-Sheet Structure for Panchayat Batch Upload:
- **Sheet 1 (`वार्ड_सारांश`)**: Summary of all wards including Ward No, Polling Station, Total Voters, Active Count, Deleted Count, and a highlighted **Grand Total (कुल योग)** row.
- **Sheet 2 (`समस्त_मतदाता_सूची`)**: Master sequence list starting from Ward 1, then Ward 2, Ward 3... in ascending order with all 11 columns.
- **Sheets 3+ (`वार्ड_1`, `वार्ड_2`...)**: Dedicated worksheet for each individual ward.

---

### 🎯 Key Highlights:
- **100% Name & Age Accuracy**: Zero missing names and zero missing ages. Advanced line-clustering prevents grid-boundary splitting issues.
- **Watermark & Deletion Detection**: Automatically recognizes shifted, deceased, or deleted voters marked with "DELETED" stamps or listed in deletion supplements.
- **Geam Digital Pro Theme**: Professional navy headers (`#0F2942`), subtle alternating zebra shading, frozen header rows, and automatic autofilters.
- **Auto-Download Support**: Enable the checkbox to automatically trigger the `.xlsx` download in your browser immediately upon extraction completion.

---

## 🚀 Getting Started

### Prerequisites:
- Python 3.10+
- Tesseract OCR (with `hin` Hindi language pack)

### Installation:
```bash
git clone https://github.com/gauravshahvt/Geam-Digital.git
cd Geam-Digital
pip install -r requirements.txt
```

### Running the Application:
- **Windows One-Click**: Double-click `run.bat`
- **Command Line**:
  ```bash
  python app.py
  ```
- Open **`http://127.0.0.1:5000`** in your browser.

---

## 🏢 Brand & Licensing
Developed for **Geam Digital**.  
All rights reserved © 2026 Geam Digital.
