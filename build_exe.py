"""Build script to compile the standalone Windows .EXE for Pichincha Excel Cleaner.

Run this script to produce:
pichincha_cleaner_app/dist/PichinchaExcelCleaner.exe

Usage:
    python build_exe.py
"""

from __future__ import annotations

from pathlib import Path
import subprocess
import sys


def main():
    root = Path(__file__).resolve().parent
    entry_point = root / "app_gui.py"
    dist_dir = root / "dist"
    build_dir = root / "build"

    print("=" * 60)
    print(" Compilando Banco Pichincha Excel Cleaner -> Windows .EXE")
    print("=" * 60)

    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--onefile",
        "--windowed",
        "--name",
        "PichinchaExcelCleaner",
        "--distpath",
        str(dist_dir),
        "--workpath",
        str(build_dir),
        "--specpath",
        str(root),
        "--hidden-import",
        "openpyxl",
        "--hidden-import",
        "openpyxl.worksheet.table",
        "--hidden-import",
        "openpyxl.styles",
        str(entry_point),
    ]

    print(f"\n[1/2] Ejecutando PyInstaller: {' '.join(cmd)}\n")
    result = subprocess.run(cmd, cwd=root)

    if result.returncode == 0:
        exe_path = dist_dir / "PichinchaExcelCleaner.exe"
        size_mb = exe_path.stat().st_size / (1024 * 1024) if exe_path.exists() else 0
        print("\n" + "=" * 60)
        print("  ¡COMPILACIÓN EXITOSA!")
        print(f"  Ejecutable generado: {exe_path}")
        print(f"  Tamaño: {size_mb:.2f} MB")
        print("  100% Autónomo: No requiere Python ni dependencias en el equipo final.")
        print("=" * 60)
    else:
        print("\n[ERROR] Falló la compilación de PyInstaller.")
        sys.exit(result.returncode)


if __name__ == "__main__":
    main()
