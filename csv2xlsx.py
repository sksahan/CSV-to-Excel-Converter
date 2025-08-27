#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Aug 27 08:27:30 2025

@author: sahan
"""

import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
import os

class CSVToExcelConverter(tk.Tk):
    """
    A GUI application to convert CSV files to Excel files.
    """
    def __init__(self):
        super().__init__()

        # Set up the main window
        self.title("CSV to Excel Converter")
        self.geometry("600x400")
        self.configure(bg="#2c3e50")
        self.resizable(False, False)

        # Apply a modern style to the widgets
        self.style = ttk.Style(self)
        self.style.theme_use('clam')
        self.style.configure("TFrame", background="#2c3e50")
        self.style.configure("TLabel", background="#2c3e50", foreground="#ecf0f1", font=("Helvetica", 12))
        self.style.configure("TButton", background="#3498db", foreground="#ecf0f1", font=("Helvetica", 10, "bold"))
        self.style.map("TButton", background=[('active', '#2980b9')])
        self.style.configure("TCombobox", fieldbackground="#ecf0f1", background="#34495e", foreground="#2c3e50")

        # Create main frame
        self.main_frame = ttk.Frame(self, padding="20")
        self.main_frame.pack(expand=True, fill="both")

        # Folder selection widgets
        self.folder_label = ttk.Label(self.main_frame, text="Select Folder:")
        self.folder_label.pack(pady=(10, 5))

        self.folder_frame = ttk.Frame(self.main_frame)
        self.folder_frame.pack(fill="x", pady=5)

        self.folder_path = tk.StringVar()
        self.folder_entry = ttk.Entry(self.folder_frame, textvariable=self.folder_path, width=60)
        self.folder_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))

        self.browse_button = ttk.Button(self.folder_frame, text="Browse", command=self.browse_folder)
        self.browse_button.pack(side="right")

        # Separator selection widgets
        self.separator_label = ttk.Label(self.main_frame, text="Select CSV Separator:")
        self.separator_label.pack(pady=(10, 5))
        
        # Updated options for auto-detection and tab only
        self.separator_options = {
            "Semicolon (;)": ";",
            "Comma (,)": ",",
            "All (auto-detect)": "auto" # Special value to trigger auto-detection logic
        }
        self.separator_var = tk.StringVar(value="All (auto-detect)") # Set default value to the auto-detect option
        self.separator_combobox = ttk.Combobox(
            self.main_frame,
            textvariable=self.separator_var,
            values=list(self.separator_options.keys()),
            state="readonly",
            width=25
        )
        self.separator_combobox.pack(pady=5)

        # Convert button
        self.convert_button = ttk.Button(self.main_frame, text="Convert Files", command=self.convert_files)
        self.convert_button.pack(pady=20)

        # Status text area
        self.status_label = ttk.Label(self.main_frame, text="Status:", font=("Helvetica", 12, "bold"))
        self.status_label.pack(pady=(10, 5))
        
        self.status_text = tk.Text(self.main_frame, height=10, width=60, bg="#34495e", fg="#ecf0f1", relief="flat")
        self.status_text.pack(expand=True, fill="both")
        self.status_text.insert(tk.END, "Please select a folder and click 'Convert Files'.\n")

    def browse_folder(self):
        """
        Opens a directory selection dialog and updates the folder path entry.
        """
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.folder_path.set(folder_selected)
            self.log_status(f"Folder selected: {folder_selected}")

    def log_status(self, message):
        """
        Appends a message to the status text area.
        """
        self.status_text.insert(tk.END, f"{message}\n")
        self.status_text.see(tk.END) # Scroll to the end of the text widget
        self.update_idletasks()

    def convert_files(self):
        """
        Walks through the selected folder, converts CSV files to Excel,
        and logs the process in the status area.
        """
        folder_path = self.folder_path.get()
        if not folder_path:
            messagebox.showerror("Error", "Please select a folder first.")
            return

        selected_separator_key = self.separator_var.get()
        separator = self.separator_options[selected_separator_key]
        
        self.log_status("\n--- Starting Conversion ---")
        self.convert_button.config(state="disabled")

        total_files = 0
        converted_files = 0
        failed_files = 0
        
        # Use os.walk to traverse the directory tree
        for root, _, files in os.walk(folder_path):
            for file_name in files:
                if file_name.lower().endswith('.csv'):
                    total_files += 1
                    csv_path = os.path.join(root, file_name)
                    excel_path = os.path.join(root, file_name.replace('.csv', '.xlsx'))
                    
                    self.log_status(f"Converting '{file_name}'...")
                    
                    try:
                        # Check if auto-detection is selected
                        if separator == "auto":
                            detected_sep = None
                            separators_to_try = [";", ","]
                            
                            for s in separators_to_try:
                                try:
                                    # Try reading the file with each separator and drop empty columns
                                    df = pd.read_csv(csv_path, sep=s, engine='python', on_bad_lines='skip')
                                    detected_sep = s
                                    self.log_status(f"  --> Detected separator: '{s}'")
                                    break # Exit the inner loop once a separator is found
                                except pd.errors.ParserError:
                                    continue # Try the next separator if parsing fails
                            
                            if detected_sep is None:
                                # If no separator works after trying all of them
                                raise ValueError("Could not auto-detect a valid separator.")
                        else:
                            # Use the user-selected separator
                            df = pd.read_csv(csv_path, sep=separator, engine='python', on_bad_lines='skip')

                        # Remove leading/trailing spaces from column names
                        df.columns = df.columns.str.strip()

                        # Drop any unnamed columns that are entirely blank
                        df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
                        
                        # Use the default openpyxl engine
                        df.to_excel(excel_path, index=False, engine='openpyxl')
                        
                        converted_files += 1
                        self.log_status(f"  --> Success! Saved as '{os.path.basename(excel_path)}'")
                    except Exception as e:
                        failed_files += 1
                        self.log_status(f"  --> Failed to convert '{file_name}': {e}")
        
        self.log_status("\n--- Conversion Complete ---")
        self.log_status(f"Total CSV files found: {total_files}")
        self.log_status(f"Successfully converted: {converted_files}")
        self.log_status(f"Failed conversions: {failed_files}")
        self.convert_button.config(state="normal")

        # Add a success message box for user confirmation
        if converted_files > 0:
            messagebox.showinfo("Conversion Complete", f"Successfully converted {converted_files} out of {total_files} files.")
        else:
            messagebox.showinfo("Conversion Complete", "No CSV files were converted.")

if __name__ == "__main__":
    app = CSVToExcelConverter()
    app.mainloop()
