# 📊 Banco Pichincha - Excel Cleaner & Formatter (100% Offline)

[🇺🇸 English](README.md) | [🇪🇸 Español](README_ES.md)

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20macOS%20%7C%20Linux-brightgreen.svg)]()
[![Offline](https://img.shields.io/badge/Security-100%25%20Offline%20%26%20Local-success.svg)]()
[![Release](https://img.shields.io/badge/Release-v1.0.0-blue.svg)](https://github.com/MaizaJoel/pichincha-excel-cleaner/releases)

A lightweight, standalone desktop application that converts messy, multi-row **Banco Pichincha** statement exports into clean, filterable, officially structured **Excel Tables (`ListObject`)** with active filter headers, date formatting (`dd-mm-yyyy hh:mm`), and currency formatting.

---

> ### ⚡ Vibecoding & Transparency Notice
> **This project was created through *vibecoding* with AI assistance.**
> - 🛡️ **100% Local & Safe**: Runs entirely on your local machine with **zero internet connection, zero telemetry, and zero cloud processing**. Your financial data never leaves your computer.
> - ⚠️ **Human Verification Recommended**: While the extraction and formatting engine has been thoroughly tested, bank statements may occasionally have non-standard rows or unexpected formatting. Always double-check your total balances against your official bank statement. If you find an unparsed row, please open an Issue!

---

## 🎯 Purpose & Problem Solved

When you download your transaction history from Banco Pichincha:
1. Each transaction is split across **alternating rows** (document numbers and beneficiaries appear on the spacer row below the transaction date and description).
2. Dates and monetary amounts use formats that cannot be directly sorted or calculated with formulas like `=SUM()`.
3. The report is not an official Excel Table, meaning you cannot easily filter by columns or sort by date/amount.

**This application automatically:**
- Joins all metadata rows with their corresponding transactions.
- Removes blank and spacer rows.
- Formats dates as **`dd-mm-yyyy hh:mm`**.
- Formats document numbers as **Text (`@`)** so leading zeros are preserved.
- Formats amounts and balances as **Currency (`$#,##0.00;[Red]-$#,##0.00;"-"`)**.
- Embeds a native **Excel Table (`ListObject`)** with active filter buttons on all headers and a frozen top row.

---

## 🔒 100% Local, Offline & Private

| Guarantee | Description |
| :--- | :--- |
| **Zero Network Calls** | No HTTP/HTTPS, WebSockets, or analytics are used. The app works without internet. |
| **No Cloud Processing** | Parsing and Excel generation happen purely inside your computer's memory using `openpyxl`. |
| **No Credentials Required** | You do not enter any banking passwords or API keys. |
| **No Database Leftovers** | Files are read from disk, processed, and saved directly to your destination path. |

---

## 🚀 Getting Started

### Option 1: Download Standalone Executable (Windows `.exe`)
No Python or software installation required!
1. Go to the [Releases](https://github.com/MaizaJoel/pichincha-excel-cleaner/releases/latest) page.
2. Download **`PichinchaExcelCleaner.exe`**.
3. Double-click to run!

---

### Option 2: Run from Python Source (Windows / macOS / Linux)

#### 1. Clone Repository
```bash
git clone https://github.com/MaizaJoel/pichincha-excel-cleaner.git
cd pichincha-excel-cleaner
```

#### 2. Install Requirements
```bash
pip install -r requirements.txt
```

#### 3. Launch the Application
```bash
python app_gui.py
```

---

### Option 3: Compile Your Own Windows `.EXE`
To build the single-file executable yourself:
```bash
python build_exe.py
```
The output file will be generated in `dist/PichinchaExcelCleaner.exe`.

---

## 🧪 Automated Testing

Run the test suite to verify the parser and Excel generation:
```bash
pytest test_engine.py -v
```

---

## 📱 Cross-Platform Roadmap (Android, iOS, macOS)

For details on extending this engine to mobile and multi-platform environments using Flutter or Flet, check the [Cross-Platform Guide](CROSS_PLATFORM_GUIDE.md).

---

## 📄 License

This project is licensed under the [MIT License](LICENSE).
