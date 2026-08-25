# Multiplatform Architecture & Deployment Guide

This document details the research, architecture, programming languages, and step-by-step roadmap to make the **Banco Pichincha Excel Cleaner** run 100% offline across **Windows, Android, iOS, and macOS**.

---

## 1. Guiding Principles & Constraints

1. **100% Local & Offline**:
   - Zero internet connections, zero telemetry, zero cloud processing.
   - Bank files stay on the local device storage.
2. **Not Web-Based**:
   - Must run as a native desktop window (Windows/macOS) or mobile app (Android/iOS) without requiring the user to open a web browser or start a local web server.
3. **High Reusability**:
   - The business logic (extracting the 2-row Pichincha layout and generating an official formatted Excel Table) is cleanly decoupled from the UI.

---

## 2. Technology & Framework Comparison

| Framework | Supported Targets | Language | Local Excel Library | Verdict |
| :--- | :--- | :--- | :--- | :--- |
| **Python (Tkinter / PyInstaller)** | Windows, macOS, Linux | Python | `openpyxl` | **Best for immediate Windows .EXE delivery** (already built in `pichincha_cleaner_app/`). Single portable `.exe`, zero dependencies. |
| **Flutter** | Windows, Android, iOS, macOS, Linux, Web | Dart | `excel` / `syncfusion_flutter_xlsio` | **Best overall choice for full mobile (Android/iOS) + desktop parity**. Single Dart codebase compiles to native `.exe`, `.apk`, and `.ipa`. |
| **Flet (Python + Flutter)** | Windows, macOS, Linux, Android, iOS | Python | `openpyxl` | Excellent middle-ground: Keeps the existing Python engine while using Flutter's multi-platform rendering engine. |
| **Tauri v2** | Windows, macOS, Linux, Android, iOS | Rust + Native WebView | `calamine` + `rust_xlsxwriter` | Blazing fast, ultra-compact binary (<8 MB), native window without running a web browser. |
| **.NET MAUI** | Windows, macOS, Android, iOS | C# | `ClosedXML` / `EPPlus` | Enterprise grade, but larger binary size and complex build toolchain. |

---

## 3. Platform-by-Platform Deployment Roadmap

### A. Windows (Available Now in `pichincha_cleaner_app/`)
- **Status**: Production Ready.
- **How it works**: Uses `app_gui.py` and `engine.py`.
- **Compile to `.exe`**:
  ```powershell
  python build_exe.py
  ```
- **Output**: `dist/PichinchaExcelCleaner.exe` (Single file, double-click to run on any Windows 10/11 machine without Python installed).

---

### B. macOS (.app / .dmg)

#### Option 1: Using the Existing Python App
- The Python code in `pichincha_cleaner_app/app_gui.py` runs natively on macOS.
- To package as a standalone macOS `.app` or `.dmg`:
  ```bash
  pip install pyinstaller
  pyinstaller --onefile --windowed --name "PichinchaExcelCleaner" app_gui.py
  ```
  *(Run on a Mac machine to produce a native signed `.app` bundle).*

---

### C. Android & iOS (Mobile)

To run on mobile devices (smartphones and tablets), users can pick a file from their local file manager or bank download folder, process it on-device, and share or open it in Microsoft Excel mobile.

#### Recommended Mobile Path: Flutter (Dart)

Flutter provides the best native experience for Android (`.apk`) and iOS (`.ipa`) with 100% offline file processing.

##### 1. Project Setup
```bash
flutter create pichincha_cleaner_mobile
cd pichincha_cleaner_mobile
flutter pub add file_picker excel path_provider open_file
```

##### 2. Core Dart Engine (`lib/engine.dart`)
```dart
import 'dart:io';
import 'package:excel/excel.dart';

class PichinchaCleaner {
  static Future<List<int>> cleanPichinchaExcel(List<int> bytes) async {
    var excel = Excel.decodeBytes(bytes);
    var outputExcel = Excel.createExcel();
    Sheet outputSheet = outputExcel['Movimientos'];

    // Headers
    outputSheet.appendRow([
      TextCellValue('Fecha'),
      TextCellValue('Concepto'),
      TextCellValue('Nro. Documento'),
      TextCellValue('Tipo'),
      TextCellValue('Beneficiario'),
      TextCellValue('Monto'),
      TextCellValue('Saldo'),
    ]);

    for (var table in excel.tables.keys) {
      var rows = excel.tables[table]!.rows;
      // Parsing logic: Find header row, merge row N with row N+1 metadata
      for (int i = 0; i < rows.length; i++) {
        var row = rows[i];
        // Parse date, concept, amount, and look ahead for document/beneficiary
        // Append cleaned row to outputSheet...
      }
    }

    return outputExcel.save()!;
  }
}
```

##### 3. Build Commands
- **Android APK**: `flutter build apk --release` (Generates standalone APK for direct installation).
- **iOS App**: `flutter build ipa --release` (Generates iOS archive for App Store or TestFlight).

---

## 4. Architecture & Security Checklist (100% Offline)

| Requirement | Implementation in `pichincha_cleaner_app` |
| :--- | :--- |
| **No Network Sockets** | No HTTP/HTTPS, WebSockets, or remote API calls are imported or initialized. |
| **Local File IO Only** | Files are read from memory/disk and saved directly to the user's selected location. |
| **Data Privacy** | No banking passwords or credentials are ever handled. Bank account numbers and transaction metadata never leave the host device. |
| **Zero Cloud Dependencies** | The Excel engine runs 100% inside `openpyxl`, requiring no Node.js or cloud services. |

---

## 5. Summary of Deliverables

1. **`engine.py`**: Pure Python core logic with custom date format `dd-mm-yyyy hh:mm`, document text format `@`, currency format `$#,##0.00;[Red]-$#,##0.00;"-"`, and native Excel Table with auto-filters.
2. **`app_gui.py`**: Native standalone desktop GUI window with instant metrics, data preview, and one-click export.
3. **`test_engine.py`**: Automated test suite validating parsing accuracy and Excel formatting.
4. **`build_exe.py`**: Automated PyInstaller packaging script generating `dist/PichinchaExcelCleaner.exe`.
