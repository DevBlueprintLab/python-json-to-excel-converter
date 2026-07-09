# 📄 JSON to Excel Converter

![Python Version](https://img.shields.io/badge/Python-3.x-blue?logo=python&logoColor=white)
![Dependencies](https://img.shields.io/badge/Dependencies-json%20%7C%20openpyxl-orange)
![Build Status](https://img.shields.io/badge/Build-Passing-brightgreen)
![License](https://img.shields.io/badge/License-MIT-green)

A Python CLI automation tool that converts structured JSON data into clean, formatted Excel workbooks.

The program validates JSON input, transforms lists of dictionaries into Excel worksheets, applies professional formatting, and saves the converted workbook to a dedicated output folder without modifying the original file.

---

## 🖥️ Pipeline Lifecycle & Live Demo

### Ingestion ➔ Processing ➔ Output

<p align="center">
  <img src="images/json-input.png" width="380" alt="JSON Input File" />
  <img src="images/terminal-output.png" width="380" alt="CLI Conversion Process" />
</p>

<p align="center">
  <img src="images/excel-output.png" width="765" alt="Generated Excel Workbook" />
</p>

---
## 🧠 Core Features & Architecture

* 📄 **JSON to Excel Conversion:** Converts JSON files containing lists of dictionaries into Excel workbooks.
* ✅ **Input Validation:** Verifies that the JSON file is valid before processing and detects unsupported structures.
* 📋 **Schema Verification:** Ensures every record contains a consistent set of fields before generating the workbook.
* 🎨 **Workbook Formatting:** Applies professional header styling, alignment, borders, and formatting for improved readability.
* 📌 **Excel Enhancements:** Automatically freezes the header row, enables filters, and adjusts column widths.
* 📂 **Safe Output Management:** Preserves the original JSON file and saves the converted workbook in a dedicated output folder.

---

## 🛠️ Tech Stack & Requirements

* **Core Language:** Python 3.x
* **JSON Processing:** `json`
* **Spreadsheet Engine:** `openpyxl`
* **File Handling:** `pathlib`

---
## ⚡ Quick Start & Usage

### 1. Clone the repository

```bash
git clone https://github.com/DevBlueprintLab/python-json-to-excel-converter.git
cd python-json-to-excel-converter
```

### 2. Install dependencies

```bash
pip install openpyxl
```

### 3. Run the tool

```bash
python json_to_excel.py
```
