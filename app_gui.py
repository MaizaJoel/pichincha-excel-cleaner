"""Banco Pichincha Excel Cleaner - Desktop GUI Application.

A standalone, 100% local, offline desktop window application for Windows, macOS, and Linux.
Processes messy Banco Pichincha Excel exports and outputs structured Excel Tables with
headers, date formatting (dd-mm-yyyy hh:mm), currency formatting, and active filter buttons.
"""

from __future__ import annotations

import os
from pathlib import Path
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Ensure engine is importable when running standalone or frozen
current_dir = Path(__file__).resolve().parent
if str(current_dir) not in sys.path:
    sys.path.insert(0, str(current_dir))

from engine import (
    CLEAN_HEADERS,
    extract_pichincha_transactions,
    generate_clean_excel,
)


class PichinchaCleanerGUI(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Banco Pichincha - Limpiador de Excel Filtrable (100% Offline)")
        self.geometry("960x680")
        self.minsize(820, 560)

        # Application state
        self.current_file_path: Path | None = None
        self.cleaned_records: list[dict] = []
        self.summary_stats: dict = {}
        self.cleaned_excel_bytes: bytes | None = None

        self._setup_styles()
        self._build_ui()

    def _setup_styles(self):
        self.style = ttk.Style(self)
        # Use clam or native theme for clean modern look
        available_themes = self.style.theme_names()
        if "vista" in available_themes:
            self.style.theme_use("vista")
        elif "clam" in available_themes:
            self.style.theme_use("clam")

        self.configure(bg="#F4F6F9")

        self.style.configure(".", background="#F4F6F9", font=("Segoe UI", 10))
        self.style.configure("Header.TFrame", background="#1F4E78")
        self.style.configure("HeaderTitle.TLabel", background="#1F4E78", foreground="#FFFFFF", font=("Segoe UI", 16, "bold"))
        self.style.configure("HeaderSubtitle.TLabel", background="#1F4E78", foreground="#D2E3F3", font=("Segoe UI", 9))
        self.style.configure("OfflineBadge.TLabel", background="#153654", foreground="#78C257", font=("Segoe UI", 9, "bold"), padding=6)

        self.style.configure("Card.TFrame", background="#FFFFFF", relief="solid", borderwidth=1)
        self.style.configure("CardTitle.TLabel", background="#FFFFFF", foreground="#6C757D", font=("Segoe UI", 9))
        self.style.configure("CardValue.TLabel", background="#FFFFFF", foreground="#212529", font=("Segoe UI", 14, "bold"))
        self.style.configure("CardIncome.TLabel", background="#FFFFFF", foreground="#28A745", font=("Segoe UI", 14, "bold"))
        self.style.configure("CardExpense.TLabel", background="#FFFFFF", foreground="#DC3545", font=("Segoe UI", 14, "bold"))

        self.style.configure("Primary.TButton", font=("Segoe UI", 10, "bold"), padding=8)
        self.style.configure("Success.TButton", font=("Segoe UI", 11, "bold"), padding=10)

        # Treeview styling
        self.style.configure(
            "Treeview",
            background="#FFFFFF",
            foreground="#212529",
            fieldbackground="#FFFFFF",
            rowheight=26,
            font=("Segoe UI", 9),
        )
        self.style.configure(
            "Treeview.Heading",
            background="#E9ECEF",
            foreground="#1F4E78",
            font=("Segoe UI", 9, "bold"),
            padding=4,
        )
        self.style.map("Treeview", background=[("selected", "#0D6EFD")], foreground=[("selected", "#FFFFFF")])

    def _build_ui(self):
        # 1. Top Header Banner
        header_frame = ttk.Frame(self, style="Header.TFrame", padding=(18, 14))
        header_frame.pack(fill=tk.X, side=tk.TOP)

        title_box = ttk.Frame(header_frame, style="Header.TFrame")
        title_box.pack(side=tk.LEFT, fill=tk.Y)

        title_lbl = ttk.Label(
            title_box,
            text="Transformador de Movimientos Banco Pichincha",
            style="HeaderTitle.TLabel",
        )
        title_lbl.pack(anchor="w")

        sub_lbl = ttk.Label(
            title_box,
            text="Convierte reportes complejos de Banco Pichincha en Tablas Excel oficiales con filtros, fechas y montos.",
            style="HeaderSubtitle.TLabel",
        )
        sub_lbl.pack(anchor="w", pady=(2, 0))

        badge = ttk.Label(
            header_frame,
            text="🔒 100% Local & Offline (Sin Nube)",
            style="OfflineBadge.TLabel",
        )
        badge.pack(side=tk.RIGHT, anchor="center")

        # 2. Main Content Container
        main_frame = ttk.Frame(self, padding=16)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 2.1 File Selection Card
        file_card = ttk.Frame(main_frame, style="Card.TFrame", padding=14)
        file_card.pack(fill=tk.X, pady=(0, 12))

        fc_top = ttk.Frame(file_card, style="Card.TFrame")
        fc_top.pack(fill=tk.X)

        self.btn_select_file = ttk.Button(
            fc_top,
            text="📁 Seleccionar Archivo XLSX...",
            style="Primary.TButton",
            command=self._on_select_file,
        )
        self.btn_select_file.pack(side=tk.LEFT, padx=(0, 12))

        self.lbl_selected_file = ttk.Label(
            fc_top,
            text="Ningún archivo seleccionado. Haz clic para elegir un reporte XLSX de Banco Pichincha.",
            background="#FFFFFF",
            foreground="#6C757D",
            font=("Segoe UI", 10, "italic"),
        )
        self.lbl_selected_file.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # 2.2 Stats Cards Row
        self.stats_row = ttk.Frame(main_frame)
        self.stats_row.pack(fill=tk.X, pady=(0, 12))

        self.card_tx = self._create_metric_card(self.stats_row, "TRANSACCIONES", "0", "CardValue.TLabel")
        self.card_tx.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 6))

        self.card_inc = self._create_metric_card(self.stats_row, "TOTAL INGRESOS (+)", "$0.00", "CardIncome.TLabel")
        self.card_inc.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=3)

        self.card_exp = self._create_metric_card(self.stats_row, "TOTAL GASTOS (-)", "$0.00", "CardExpense.TLabel")
        self.card_exp.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=3)

        self.card_net = self._create_metric_card(self.stats_row, "FLUJO NETO", "$0.00", "CardValue.TLabel")
        self.card_net.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(6, 0))

        # 2.3 Preview Table Area
        preview_container = ttk.Frame(main_frame, style="Card.TFrame", padding=10)
        preview_container.pack(fill=tk.BOTH, expand=True, pady=(0, 12))

        preview_header = ttk.Frame(preview_container, style="Card.TFrame")
        preview_header.pack(fill=tk.X, pady=(0, 6))

        ttk.Label(
            preview_header,
            text="Vista Previa de la Tabla Limpia (Primeras filas procesadas)",
            background="#FFFFFF",
            font=("Segoe UI", 10, "bold"),
            foreground="#1F4E78",
        ).pack(side=tk.LEFT)

        self.lbl_row_count = ttk.Label(
            preview_header,
            text="",
            background="#FFFFFF",
            foreground="#6C757D",
            font=("Segoe UI", 9),
        )
        self.lbl_row_count.pack(side=tk.RIGHT)

        # Table with scrollbars
        table_frame = ttk.Frame(preview_container)
        table_frame.pack(fill=tk.BOTH, expand=True)

        cols = ("fecha", "concepto", "doc", "tipo", "beneficiario", "monto", "saldo")
        self.tree = ttk.Treeview(table_frame, columns=cols, show="headings", selectmode="browse")

        col_configs = [
            ("fecha", "Fecha", 130, "center"),
            ("concepto", "Concepto", 260, "w"),
            ("doc", "Nro. Documento", 110, "center"),
            ("tipo", "Tipo", 80, "center"),
            ("beneficiario", "Beneficiario", 140, "w"),
            ("monto", "Monto ($)", 100, "e"),
            ("saldo", "Saldo ($)", 100, "e"),
        ]

        for cid, heading, width, align in col_configs:
            self.tree.heading(cid, text=heading)
            self.tree.column(cid, width=width, anchor=align)

        v_scroll = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        h_scroll = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        v_scroll.grid(row=0, column=1, sticky="ns")
        h_scroll.grid(row=1, column=0, sticky="ew")

        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        # 2.4 Action & Export Bar
        action_bar = ttk.Frame(main_frame)
        action_bar.pack(fill=tk.X, side=tk.BOTTOM)

        self.var_auto_open = tk.BooleanVar(value=True)
        self.chk_auto_open = ttk.Checkbutton(
            action_bar,
            text="Abrir en Microsoft Excel al guardar",
            variable=self.var_auto_open,
        )
        self.chk_auto_open.pack(side=tk.LEFT, anchor="center")

        self.btn_export = ttk.Button(
            action_bar,
            text="💾 Guardar Excel Limpio y Filtrable (.xlsx)",
            style="Success.TButton",
            state=tk.DISABLED,
            command=self._on_export_excel,
        )
        self.btn_export.pack(side=tk.RIGHT)

        # Status Bar
        self.status_bar = ttk.Label(
            self,
            text="Listo. Selecciona un archivo XLSX de Banco Pichincha para comenzar.",
            relief="sunken",
            anchor="w",
            font=("Segoe UI", 9),
            padding=(8, 4),
        )
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def _create_metric_card(self, parent: ttk.Frame, title: str, initial_value: str, value_style: str) -> ttk.Frame:
        card = ttk.Frame(parent, style="Card.TFrame", padding=(12, 10))
        lbl_title = ttk.Label(card, text=title, style="CardTitle.TLabel")
        lbl_title.pack(anchor="w")
        lbl_val = ttk.Label(card, text=initial_value, style=value_style)
        lbl_val.pack(anchor="w", pady=(2, 0))
        card.value_label = lbl_val  # type: ignore
        return card

    def _update_status(self, text: str):
        self.status_bar.config(text=text)
        self.update_idletasks()

    def _on_select_file(self):
        file_path = filedialog.askopenfilename(
            title="Seleccionar reporte de Banco Pichincha",
            filetypes=[("Archivos Excel (*.xlsx)", "*.xlsx"), ("Todos los archivos", "*.*")],
        )
        if not file_path:
            return

        self._process_selected_file(Path(file_path))

    def _process_selected_file(self, path: Path):
        self.current_file_path = path
        self.lbl_selected_file.config(
            text=f"📄 {path.name} ({path.stat().st_size / 1024:.1f} KB)",
            foreground="#212529",
            font=("Segoe UI", 10, "bold"),
        )
        self._update_status(f"Procesando {path.name}...")

        try:
            records, summaries = extract_pichincha_transactions(path)
            self.cleaned_records = records
            stats = summaries[0] if summaries else {}
            self.summary_stats = stats

            # Generate in-memory Excel table
            self.cleaned_excel_bytes = generate_clean_excel(records)

            # Update Metric Cards
            tx_count = stats.get("transactions", len(records))
            inc_total = stats.get("total_income", 0.0)
            exp_total = stats.get("total_expenses", 0.0)
            net_total = stats.get("net_flow", inc_total - exp_total)
            removed = stats.get("rows_removed", 0)

            self.card_tx.value_label.config(text=f"{tx_count:,}")  # type: ignore
            self.card_inc.value_label.config(text=f"${inc_total:,.2f}")  # type: ignore
            self.card_exp.value_label.config(text=f"-${exp_total:,.2f}")  # type: ignore
            self.card_net.value_label.config(text=f"${net_total:,.2f}")  # type: ignore

            # Update Table Preview
            for row in self.tree.get_children():
                self.tree.delete(row)

            for rec in records[:150]:  # preview up to 150 items for speed
                dt_str = rec["Fecha"].strftime("%d-%m-%Y %H:%M") if rec.get("Fecha") else ""
                monto_val = rec.get("Monto", 0.0)
                monto_str = f"${monto_val:,.2f}" if monto_val >= 0 else f"-${abs(monto_val):,.2f}"
                saldo_val = rec.get("Saldo")
                saldo_str = f"${saldo_val:,.2f}" if saldo_val is not None else ""

                self.tree.insert(
                    "",
                    tk.END,
                    values=(
                        dt_str,
                        rec.get("Concepto", ""),
                        rec.get("Nro. Documento", ""),
                        rec.get("Tipo", ""),
                        rec.get("Beneficiario", ""),
                        monto_str,
                        saldo_str,
                    ),
                )

            self.lbl_row_count.config(
                text=f"{len(records)} transacciones limpias | {removed} filas espaciadoras eliminadas"
            )
            self.btn_export.config(state=tk.NORMAL)
            self._update_status(f"Éxito: {len(records)} transacciones extraídas. Listo para guardar Excel.")

        except Exception as ex:
            self.btn_export.config(state=tk.DISABLED)
            self._update_status(f"Error: {str(ex)}")
            messagebox.showerror(
                "Error al procesar archivo",
                f"No se pudo extraer los datos de Banco Pichincha:\n\n{str(ex)}\n\n"
                "Asegúrate de que sea un archivo XLSX emitido por Banco Pichincha.",
            )

    def _on_export_excel(self):
        if not self.cleaned_excel_bytes or not self.current_file_path:
            return

        default_name = f"{self.current_file_path.stem}_filtrable.xlsx"
        save_path = filedialog.asksaveasfilename(
            title="Guardar archivo Excel limpio",
            initialdir=str(self.current_file_path.parent),
            initialfile=default_name,
            defaultextension=".xlsx",
            filetypes=[("Libro de Excel (*.xlsx)", "*.xlsx")],
        )

        if not save_path:
            return

        try:
            target = Path(save_path)
            target.write_bytes(self.cleaned_excel_bytes)
            self._update_status(f"Archivo guardado exitosamente: {target.name}")

            if self.var_auto_open.get():
                try:
                    os.startfile(str(target))
                except Exception:
                    pass

            messagebox.showinfo(
                "Excel Generado",
                f"¡Archivo generado exitosamente!\n\n"
                f"Ubicación: {target}\n"
                f"Transacciones: {len(self.cleaned_records)}\n\n"
                "El archivo ya contiene la tabla con filtros activos y formatos listos.",
            )

        except Exception as ex:
            messagebox.showerror("Error al guardar", f"No se pudo escribir el archivo:\n\n{str(ex)}")


def main():
    app = PichinchaCleanerGUI()
    # Center window on screen
    app.eval("tk::PlaceWindow . center")
    app.mainloop()


if __name__ == "__main__":
    main()
