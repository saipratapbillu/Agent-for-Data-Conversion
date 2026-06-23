#!/usr/bin/env python3
"""
Excel to CSV Converter Agent
Converts all Excel files to CSV format with a user-friendly folder selection UI.
"""

import os
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from pathlib import Path
from datetime import datetime
from config import (
    EXCEL_SOURCE_PATH,
    DEFAULT_OUTPUT_PATH,
    SERVICE_ACCOUNT_KEY_PATH,
    AUTO_SAVE_ENABLED,
    CREATE_LOG_FILE,
)


class ExcelToCSVAgent:
    def __init__(self, root, default_source=None, default_destination=None, service_account_key=None):
        self.root = root
        self.root.title("Excel to CSV Converter Agent")
        self.root.geometry("760x560")
        self.root.resizable(False, False)

        self.root.configure(bg="#f0f0f0")

        self.default_source = default_source or EXCEL_SOURCE_PATH or str(Path.home())
        self.default_destination = default_destination or DEFAULT_OUTPUT_PATH or str(Path.home())
        self.service_account_key = service_account_key or SERVICE_ACCOUNT_KEY_PATH
        self.is_authenticated = self.service_account_key is not None

        self.source_folder = self.default_source
        self.destination_folder = self.default_destination
        self.sheet_name = None
        self.create_log = CREATE_LOG_FILE

        self.setup_ui()

    def setup_ui(self):
        title_label = tk.Label(
            self.root,
            text="Excel to CSV Batch Converter",
            font=("Segoe UI", 20, "bold"),
            bg="#f0f0f0",
            fg="#2c3e50",
        )
        title_label.pack(pady=(15, 5))

        subtitle_label = tk.Label(
            self.root,
            text="Select source and destination paths, then convert your Excel files to CSV.",
            font=("Segoe UI", 10),
            bg="#f0f0f0",
            fg="#4d4d4d",
        )
        subtitle_label.pack(pady=(0, 15))

        status_text = "Authenticated" if self.is_authenticated else "Not authenticated"
        status_color = "#2e7d32" if self.is_authenticated else "#f57c00"
        status_frame = tk.Frame(self.root, bg="#ffffff", relief=tk.RIDGE, bd=2)
        status_frame.pack(padx=18, pady=5, fill=tk.X)

        tk.Label(
            status_frame,
            text=f"Service Account: {status_text}",
            font=("Segoe UI", 10, "bold"),
            bg="#ffffff",
            fg=status_color,
        ).pack(anchor="w", padx=12, pady=10)

        form_frame = tk.Frame(self.root, bg="#ffffff", relief=tk.RIDGE, bd=2)
        form_frame.pack(padx=18, pady=10, fill=tk.X)

        tk.Label(
            form_frame,
            text="Source Folder:",
            font=("Segoe UI", 10),
            bg="#ffffff",
        ).grid(row=0, column=0, sticky="w", padx=12, pady=(12, 4))

        self.source_display = tk.Label(
            form_frame,
            text=self.source_folder,
            font=("Segoe UI", 9),
            bg="#f7f7f7",
            fg="#4d4d4d",
            wraplength=610,
            justify="left",
            anchor="w",
            relief=tk.SUNKEN,
            bd=1,
            padx=6,
            pady=6,
        )
        self.source_display.grid(row=1, column=0, columnspan=2, sticky="ew", padx=12)

        tk.Button(
            form_frame,
            text="Browse Source",
            command=self.select_source_folder,
            font=("Segoe UI", 10),
            bg="#1565c0",
            fg="white",
            activebackground="#0d47a1",
            padx=10,
            pady=6,
            relief=tk.RAISED,
            bd=2,
        ).grid(row=1, column=2, sticky="e", padx=12)

        tk.Label(
            form_frame,
            text="Destination Folder:",
            font=("Segoe UI", 10),
            bg="#ffffff",
        ).grid(row=2, column=0, sticky="w", padx=12, pady=(12, 4))

        self.destination_display = tk.Label(
            form_frame,
            text=self.destination_folder,
            font=("Segoe UI", 9),
            bg="#f7f7f7",
            fg="#4d4d4d",
            wraplength=610,
            justify="left",
            anchor="w",
            relief=tk.SUNKEN,
            bd=1,
            padx=6,
            pady=6,
        )
        self.destination_display.grid(row=3, column=0, columnspan=2, sticky="ew", padx=12)

        tk.Button(
            form_frame,
            text="Browse Destination",
            command=self.select_destination_folder,
            font=("Segoe UI", 10),
            bg="#2e7d32",
            fg="white",
            activebackground="#1b5e20",
            padx=10,
            pady=6,
            relief=tk.RAISED,
            bd=2,
        ).grid(row=3, column=2, sticky="e", padx=12)

        tk.Label(
            form_frame,
            text="Optional sheet name (leave empty for first sheet):",
            font=("Segoe UI", 10),
            bg="#ffffff",
        ).grid(row=4, column=0, sticky="w", padx=12, pady=(12, 4))

        self.sheet_entry = tk.Entry(form_frame, font=("Segoe UI", 10), bd=1, relief=tk.SUNKEN)
        self.sheet_entry.grid(row=5, column=0, columnspan=3, sticky="ew", padx=12, pady=(0, 12))

        self.auto_save_var = tk.BooleanVar(value=AUTO_SAVE_ENABLED)
        self.auto_save_var.set(AUTO_SAVE_ENABLED)
        auto_save_check = tk.Checkbutton(
            form_frame,
            text="Keep existing CSV files in destination folder",
            font=("Segoe UI", 10),
            bg="#ffffff",
            variable=self.auto_save_var,
            onvalue=False,
            offvalue=True,
            command=self.update_save_option,
        )
        auto_save_check.grid(row=6, column=0, columnspan=3, sticky="w", padx=12, pady=(0, 12))

        self.result_box = tk.Text(
            self.root,
            height=12,
            font=("Segoe UI", 10),
            bg="#ffffff",
            fg="#333",
            bd=2,
            relief=tk.SUNKEN,
            wrap=tk.WORD,
        )
        self.result_box.pack(padx=18, pady=(0, 8), fill=tk.BOTH, expand=True)
        self.result_box.configure(state=tk.DISABLED)

        footer_frame = tk.Frame(self.root, bg="#f0f0f0")
        footer_frame.pack(padx=18, pady=(0, 10), fill=tk.X)

        convert_btn = tk.Button(
            footer_frame,
            text="Convert Source Folder",
            command=self.convert_folder,
            font=("Segoe UI", 11, "bold"),
            bg="#1976d2",
            fg="white",
            padx=16,
            pady=10,
            relief=tk.RAISED,
            bd=2,
        )
        convert_btn.pack(side=tk.LEFT, padx=(0, 12))

        self.open_folder_btn = tk.Button(
            footer_frame,
            text="Open Output Folder",
            command=self.open_output_folder,
            font=("Segoe UI", 11, "bold"),
            bg="#388e3c",
            fg="white",
            padx=14,
            pady=10,
            relief=tk.RAISED,
            bd=2,
            state=tk.DISABLED,
        )
        self.open_folder_btn.pack(side=tk.LEFT)

        self.status_label = tk.Label(
            self.root,
            text="Ready to convert. Use the default paths or select new folders.",
            font=("Segoe UI", 9),
            bg="#e0e0e0",
            fg="#111",
            anchor="w",
            padx=10,
            pady=6,
        )
        self.status_label.pack(fill=tk.X, side=tk.BOTTOM)

    def update_save_option(self):
        self.delete_existing = not self.auto_save_var.get()
        self.status_label.config(text="Conversion option updated.")

    def select_source_folder(self):
        selected = filedialog.askdirectory(
            title="Select source folder containing Excel files",
            initialdir=self.source_folder,
        )
        if selected:
            self.source_folder = selected
            self.source_display.config(text=self.source_folder)
            self.status_label.config(text="Source folder updated.")

    def select_destination_folder(self):
        selected = filedialog.askdirectory(
            title="Select destination folder for CSV files",
            initialdir=self.destination_folder,
        )
        if selected:
            self.destination_folder = selected
            self.destination_display.config(text=self.destination_folder)
            self.status_label.config(text="Destination folder updated.")

    def append_result(self, message):
        self.result_box.configure(state=tk.NORMAL)
        self.result_box.insert(tk.END, message + "\n")
        self.result_box.see(tk.END)
        self.result_box.configure(state=tk.DISABLED)

    def find_excel_files(self):
        path = Path(self.source_folder)
        excel_files = []
        for ext in ['*.xlsx', '*.xls']:
            excel_files.extend([f for f in path.rglob(ext) if not f.name.startswith('~$')])
        return sorted(excel_files)

    def delete_existing_csvs(self):
        output = Path(self.destination_folder)
        csv_files = list(output.glob('*.csv'))
        if not csv_files:
            self.append_result('No existing CSV files found in destination.')
            return

        self.append_result(f'Deleting {len(csv_files)} existing CSV file(s) ...')
        deleted = 0
        for csv_file in csv_files:
            try:
                csv_file.unlink()
                deleted += 1
            except Exception as e:
                self.append_result(f'  ✖ Failed to delete {csv_file.name}: {e}')
        self.append_result(f'  ✓ Deleted {deleted} CSV files.')

    def convert_folder(self):
        self.open_folder_btn.config(state=tk.DISABLED)
        self.result_box.configure(state=tk.NORMAL)
        self.result_box.delete(1.0, tk.END)
        self.result_box.configure(state=tk.DISABLED)

        source = Path(self.source_folder)
        destination = Path(self.destination_folder)

        if not source.exists():
            messagebox.showerror("Source Missing", "The selected source folder does not exist.")
            return

        if not destination.exists():
            try:
                destination.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                messagebox.showerror("Destination Error", f"Could not create destination folder:\n{e}")
                return

        excel_files = self.find_excel_files()
        if not excel_files:
            messagebox.showwarning("No Excel Files", "No Excel (.xlsx or .xls) files were found in the source folder.")
            return

        if self.delete_existing:
            self.delete_existing_csvs()

        self.append_result(f'Converting {len(excel_files)} Excel file(s) from: {source}')
        self.append_result(f'Target destination: {destination}')
        self.append_result('-' * 72)

        failures = 0
        for idx, excel_file in enumerate(excel_files, start=1):
            try:
                sheet_name = self.sheet_entry.get().strip() if self.sheet_entry.get().strip() else 0
                df = pd.read_excel(excel_file, sheet_name=sheet_name)
                output_file = destination / f"{excel_file.stem}.csv"

                counter = 1
                while output_file.exists():
                    output_file = destination / f"{excel_file.stem}_{counter}.csv"
                    counter += 1

                df.to_csv(output_file, index=False, encoding='utf-8')
                self.append_result(f'[{idx}/{len(excel_files)}] ✓ Converted: {excel_file.name}')
            except Exception as e:
                failures += 1
                self.append_result(f'[{idx}/{len(excel_files)}] ✖ Failed: {excel_file.name} ({e})')

        self.append_result('-' * 72)
        self.append_result(f'Completed with {failures} failure(s).')

        self.log_conversion(source, destination, len(excel_files), failures)
        self.status_label.config(text="Conversion finished. You may open the output folder.")
        self.open_folder_btn.config(state=tk.NORMAL)

        messagebox.showinfo(
            "Conversion Complete",
            f"Batch conversion finished.\n\nDestination folder:\n{destination}"
        )

    def open_output_folder(self):
        try:
            os.startfile(self.destination_folder)
        except Exception as e:
            messagebox.showerror("Open Folder Failed", f"Could not open output folder:\n{e}")

    def log_conversion(self, source, destination, total_files, failures):
        if not self.create_log:
            return

        log_file = Path(self.destination_folder) / 'conversion_log.txt'
        try:
            timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            with open(log_file, 'a', encoding='utf-8') as f:
                f.write(
                    f'[{timestamp}] source={source} destination={destination} total={total_files} failures={failures}\n'
                )
            self.append_result(f'Log saved to {log_file}')
        except Exception as e:
            self.append_result(f'Warning: Could not save log file: {e}')


def main(default_source=None, default_destination=None, service_account_key=None):
    root = tk.Tk()
    app = ExcelToCSVAgent(
        root,
        default_source=default_source,
        default_destination=default_destination,
        service_account_key=service_account_key,
    )
    root.mainloop()


if __name__ == '__main__':
    import sys

    default_source = None
    default_destination = None
    service_account_key = None

    if len(sys.argv) > 1:
        default_source = sys.argv[1]
    if len(sys.argv) > 2:
        default_destination = sys.argv[2]
    if len(sys.argv) > 3:
        service_account_key = sys.argv[3]

    main(default_source=default_source, default_destination=default_destination, service_account_key=service_account_key)
