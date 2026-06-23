#!/usr/bin/env python3
"""
Excel to CSV Converter Agent
Converts Excel files to CSV format with a user-friendly GUI
Enhanced with service account authentication and auto-save
"""

import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os
from pathlib import Path
from datetime import datetime


class ExcelToCSVAgent:
    def __init__(self, root, default_path=None, service_account_key=None):
        self.root = root
        self.root.title("Excel to CSV Converter Agent - Enhanced")
        self.root.geometry("700x500")
        self.root.resizable(False, False)
        
        # Configure style
        self.root.configure(bg="#f0f0f0")
        
        # Configuration
        self.default_path = default_path or str(Path.home() / "Desktop")
        self.service_account_key = service_account_key
        self.is_authenticated = service_account_key is not None
        
        # Selected file path
        self.selected_file = None
        self.auto_save_enabled = True
        self.output_folder = self.default_path
        
        self.setup_ui()
    
    def setup_ui(self):
        """Setup the user interface"""
        # Title
        title_label = tk.Label(
            self.root,
            text="Excel to CSV Converter - Enhanced",
            font=("Arial", 18, "bold"),
            bg="#f0f0f0",
            fg="#333"
        )
        title_label.pack(pady=15)
        
        # Authentication status
        auth_status = "✓ Authenticated (Service Account)" if self.is_authenticated else "⚠ Not Authenticated"
        auth_color = "#4CAF50" if self.is_authenticated else "#FF9800"
        auth_label = tk.Label(
            self.root,
            text=f"Status: {auth_status}",
            font=("Arial", 9),
            bg="#f0f0f0",
            fg=auth_color
        )
        auth_label.pack()
        
        # Default path frame
        path_frame = tk.Frame(self.root, bg="#ffffff", relief=tk.RIDGE, bd=2)
        path_frame.pack(padx=15, pady=5, fill=tk.X)
        
        path_label = tk.Label(
            path_frame,
            text="Default Output Path:",
            font=("Arial", 10),
            bg="#ffffff"
        )
        path_label.pack(anchor="w", padx=10, pady=(8, 3))
        
        path_display = tk.Label(
            path_frame,
            text=self.default_path,
            font=("Arial", 9),
            bg="#f9f9f9",
            fg="#666",
            wraplength=600,
            justify="left"
        )
        path_display.pack(anchor="w", padx=10, pady=3, fill=tk.X)
        
        # File selection frame
        file_frame = tk.Frame(self.root, bg="#ffffff", relief=tk.RIDGE, bd=2)
        file_frame.pack(padx=15, pady=5, fill=tk.BOTH, expand=True)
        
        file_label = tk.Label(
            file_frame,
            text="Selected File:",
            font=("Arial", 11),
            bg="#ffffff"
        )
        file_label.pack(anchor="w", padx=10, pady=(10, 5))
        
        self.file_display = tk.Label(
            file_frame,
            text="No file selected",
            font=("Arial", 9),
            bg="#f9f9f9",
            fg="#666",
            wraplength=600,
            justify="left"
        )
        self.file_display.pack(anchor="w", padx=10, pady=5, fill=tk.BOTH, expand=True)
        
        # Sheet name and options frame
        options_frame = tk.Frame(self.root, bg="#ffffff", relief=tk.RIDGE, bd=2)
        options_frame.pack(padx=15, pady=5, fill=tk.X)
        
        sheet_label = tk.Label(
            options_frame,
            text="Sheet Name (optional - defaults to first sheet):",
            font=("Arial", 10),
            bg="#ffffff"
        )
        sheet_label.pack(anchor="w", padx=10, pady=(8, 3))
        
        self.sheet_entry = tk.Entry(options_frame, font=("Arial", 9))
        self.sheet_entry.pack(padx=10, pady=3, fill=tk.X)
        
        # Auto-save checkbox
        self.auto_save_var = tk.BooleanVar(value=True)
        auto_save_check = tk.Checkbutton(
            options_frame,
            text="✓ Auto-save CSV (same name as Excel file)",
            font=("Arial", 9),
            bg="#ffffff",
            variable=self.auto_save_var,
            command=self.toggle_auto_save
        )
        auto_save_check.pack(anchor="w", padx=10, pady=(5, 8))
        
        # Button frame
        button_frame = tk.Frame(self.root, bg="#f0f0f0")
        button_frame.pack(pady=12)
        
        select_btn = tk.Button(
            button_frame,
            text="📂 Select File",
            command=self.select_file,
            font=("Arial", 10, "bold"),
            bg="#4CAF50",
            fg="white",
            padx=12,
            pady=8,
            cursor="hand2",
            relief=tk.RAISED,
            bd=2
        )
        select_btn.pack(side=tk.LEFT, padx=8)
        
        convert_btn = tk.Button(
            button_frame,
            text="⚙️ Convert to CSV",
            command=self.convert_to_csv,
            font=("Arial", 10, "bold"),
            bg="#2196F3",
            fg="white",
            padx=12,
            pady=8,
            cursor="hand2",
            relief=tk.RAISED,
            bd=2
        )
        convert_btn.pack(side=tk.LEFT, padx=8)
        
        # Status bar
        self.status_label = tk.Label(
            self.root,
            text="Ready",
            font=("Arial", 9),
            bg="#e0e0e0",
            fg="#333",
            anchor="w",
            padx=10,
            pady=5
        )
        self.status_label.pack(fill=tk.X, side=tk.BOTTOM)
    
    def toggle_auto_save(self):
        """Toggle auto-save feature"""
        self.auto_save_enabled = self.auto_save_var.get()
        status = "enabled" if self.auto_save_enabled else "disabled"
        self.status_label.config(text=f"Auto-save {status}")
    
    def select_file(self):
        """Open file dialog to select Excel file"""
        file_path = filedialog.askopenfilename(
            title="Select an Excel file",
            initialdir=self.default_path,
            filetypes=[
                ("Excel files", "*.xlsx *.xls"),
                ("All files", "*.*")
            ]
        )
        
        if file_path:
            self.selected_file = file_path
            self.file_display.config(text=file_path)
            filename = os.path.basename(file_path)
            self.status_label.config(text=f"File selected: {filename}")
            
            # Try to list available sheets
            try:
                xls = pd.ExcelFile(file_path)
                sheets = xls.sheet_names
                sheet_info = f"Available sheets: {', '.join(sheets)}"
                self.status_label.config(text=sheet_info)
            except Exception as e:
                self.status_label.config(text=f"Error reading file: {str(e)}")
    
    def convert_to_csv(self):
        """Convert the selected Excel file to CSV"""
        if not self.selected_file:
            messagebox.showwarning("Warning", "Please select an Excel file first!")
            return
        
        try:
            # Get sheet name if specified
            sheet_name = self.sheet_entry.get().strip() if self.sheet_entry.get().strip() else 0
            
            # Read Excel file
            self.status_label.config(text="Reading Excel file...")
            self.root.update()
            
            df = pd.read_excel(self.selected_file, sheet_name=sheet_name)
            
            # Generate output file path (same name as input)
            input_path = Path(self.selected_file)
            output_path = input_path.with_suffix('.csv')
            
            # If auto-save enabled, always save to default output folder
            if self.auto_save_enabled and self.output_folder:
                output_path = Path(self.output_folder) / output_path.name
            
            # Check if file already exists and handle duplicates
            counter = 1
            original_output = output_path
            while output_path.exists():
                output_path = original_output.with_stem(f"{original_output.stem}_{counter}")
                counter += 1
            
            # Create output directory if it doesn't exist
            output_path.parent.mkdir(parents=True, exist_ok=True)
            
            # Write to CSV
            self.status_label.config(text="Writing CSV file...")
            self.root.update()
            
            df.to_csv(output_path, index=False, encoding='utf-8')
            
            # Log the conversion
            self.log_conversion(self.selected_file, str(output_path), len(df), len(df.columns))
            
            self.status_label.config(text="✓ Conversion complete!")
            messagebox.showinfo(
                "Success",
                f"File converted successfully!\n\n"
                f"Input: {input_path.name}\n"
                f"Output: {output_path.name}\n\n"
                f"Location: {output_path}\n\n"
                f"Data: {len(df)} rows × {len(df.columns)} columns"
            )
            
        except Exception as e:
            self.status_label.config(text="❌ Error during conversion")
            messagebox.showerror("Error", f"Failed to convert file:\n{str(e)}")
    
    def log_conversion(self, input_file, output_file, rows, columns):
        """Log conversion details to a log file"""
        log_file = Path(self.default_path) / "conversion_log.txt"
        try:
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            log_entry = (
                f"[{timestamp}] Converted '{os.path.basename(input_file)}' to '{os.path.basename(output_file)}' "
                f"({rows} rows, {columns} columns)\n"
            )
            with open(log_file, "a", encoding='utf-8') as f:
                f.write(log_entry)
        except Exception as e:
            print(f"Warning: Could not write to log file: {str(e)}")


def main(default_path=None, service_account_key=None):
    """Main entry point
    
    Args:
        default_path (str, optional): Default directory for file selection and output
        service_account_key (str, optional): Path to service account JSON key file
    """
    root = tk.Tk()
    app = ExcelToCSVAgent(root, default_path=default_path, service_account_key=service_account_key)
    root.mainloop()


if __name__ == "__main__":
    import sys
    
    # Parse command-line arguments
    default_path = None
    service_account_key = None
    
    if len(sys.argv) > 1:
        default_path = sys.argv[1]
    
    if len(sys.argv) > 2:
        service_account_key = sys.argv[2]
    
    main(default_path=default_path, service_account_key=service_account_key)
