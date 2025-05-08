import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import sys

# Attempt to import processing logic, handle potential import errors
try:
    # Ensure the script directory is in the Python path
    script_dir = os.path.dirname(os.path.abspath(__file__))
    if script_dir not in sys.path:
        sys.path.append(script_dir)

    from processing_logic import read_pos_csv, process_excel, generate_output_filename
except ImportError as e:
    messagebox.showerror("Import Error", f"Failed to import processing_logic.py. Make sure it's in the same directory.\nError: {e}")
    sys.exit(1) # Exit if core logic is missing
except Exception as e:
    messagebox.showerror("Error", f"An unexpected error occurred during import:\n{e}")
    sys.exit(1)

class ExcelProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Processor")
        # self.root.geometry("500x250") # Optional: Set initial size

        self.csv_file_path = tk.StringVar(value="No file selected")
        self.xlsx_file_path = tk.StringVar(value="No file selected")
        self.selected_week = tk.StringVar(value="Week 1") # Default week
        self.status_message = tk.StringVar(value="")

        # --- GUI Layout ---
        frame = ttk.Frame(root, padding="10")
        frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)

        # CSV File Selection
        ttk.Label(frame, text="POS CSV File:").grid(row=0, column=0, sticky=tk.W, pady=2)
        ttk.Button(frame, text="Browse...", command=self.select_csv_file).grid(row=0, column=1, sticky=tk.W, padx=5)
        ttk.Label(frame, textvariable=self.csv_file_path, wraplength=350).grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=2)

        # XLSX Template Selection
        ttk.Label(frame, text="Excel Template File:").grid(row=2, column=0, sticky=tk.W, pady=2)
        ttk.Button(frame, text="Browse...", command=self.select_xlsx_file).grid(row=2, column=1, sticky=tk.W, padx=5)
        ttk.Label(frame, textvariable=self.xlsx_file_path, wraplength=350).grid(row=3, column=0, columnspan=2, sticky=tk.W, pady=2)

        # Week Selection
        ttk.Label(frame, text="Select Week:").grid(row=4, column=0, sticky=tk.W, pady=5)
        week_frame = ttk.Frame(frame)
        week_frame.grid(row=4, column=1, sticky=tk.W)
        ttk.Radiobutton(week_frame, text="Week 1", variable=self.selected_week, value="Week 1").pack(side=tk.LEFT)
        ttk.Radiobutton(week_frame, text="Week 2", variable=self.selected_week, value="Week 2").pack(side=tk.LEFT, padx=5)

        # Run Button
        self.run_button = ttk.Button(frame, text="Run Processing", command=self.run_processing)
        self.run_button.grid(row=5, column=0, columnspan=2, pady=15)

        # Status Label
        ttk.Label(frame, textvariable=self.status_message, foreground="grey").grid(row=6, column=0, columnspan=2, sticky=tk.W, pady=2)

        # Configure resizing behavior
        frame.columnconfigure(0, weight=1)


    def select_csv_file(self):
        file_path = filedialog.askopenfilename(
            title="Select POS CSV File",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if file_path:
            self.csv_file_path.set(file_path)
            self.status_message.set("") # Clear status on new selection

    def select_xlsx_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Excel Template File",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file_path:
            self.xlsx_file_path.set(file_path)
            self.status_message.set("") # Clear status on new selection

    def run_processing(self):
        csv_path = self.csv_file_path.get()
        xlsx_path = self.xlsx_file_path.get()
        week = self.selected_week.get()

        # --- Validation ---
        if csv_path == "No file selected" or not os.path.exists(csv_path):
            messagebox.showerror("Error", "Please select a valid POS CSV file.")
            return
        if xlsx_path == "No file selected" or not os.path.exists(xlsx_path):
            messagebox.showerror("Error", "Please select a valid Excel template file.")
            return

        # --- Disable Button & Set Status ---
        self.run_button.config(state=tk.DISABLED)
        self.status_message.set("Processing...")
        self.root.update_idletasks() # Ensure status update is visible

        try:
            # --- Generate Output Filename ---
            # --- Generate Output Filename and Path ---
            base_output_filename = generate_output_filename(xlsx_path)
            # Get the user's Downloads folder
            downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
            # Ensure the Downloads folder exists, create if not (optional, but good practice)
            if not os.path.exists(downloads_folder):
                os.makedirs(downloads_folder) # Create if it doesn't exist
            output_path = os.path.join(downloads_folder, base_output_filename)

            # --- Read CSV ---
            self.status_message.set("Reading CSV...")
            self.root.update_idletasks() # Update GUI
            # read_pos_csv returns DataFrame on success, None on failure
            pos_data = read_pos_csv(csv_path)

            if pos_data is None:
                # Error message already logged by read_pos_csv
                messagebox.showerror("CSV Read Error", f"Failed to read or process CSV file: {os.path.basename(csv_path)}. Check logs or file format.")
                self.status_message.set("CSV read failed.")
                self.run_button.config(state=tk.NORMAL)
                return

            # --- Process Excel ---
            self.status_message.set("Processing Excel...")
            self.root.update_idletasks()
            success_excel, message = process_excel(xlsx_path, pos_data, week, output_path) # Pass the DataFrame

            if success_excel:
                messagebox.showinfo("Success", f"Processing complete!\nOutput saved to:\n{output_path}")
                self.status_message.set("Complete!")
            else:
                messagebox.showerror("Processing Error", f"Failed to process Excel file:\n{message}")
                self.status_message.set("Excel processing failed.")

        except FileNotFoundError as e:
             messagebox.showerror("File Error", f"File not found during processing:\n{e}")
             self.status_message.set("Error: File not found.")
        except PermissionError as e:
             messagebox.showerror("Permission Error", f"Permission denied. Cannot read/write file:\n{e}")
             self.status_message.set("Error: Permission denied.")
        except KeyError as e:
             messagebox.showerror("Data Error", f"Missing expected column or data in input files: {e}")
             self.status_message.set("Error: Data mismatch.")
        except Exception as e:
            # Catch any other unexpected errors from the processing logic or GUI
            messagebox.showerror("Unexpected Error", f"An unexpected error occurred:\n{e}")
            self.status_message.set("An unexpected error occurred.")
        finally:
            # --- Re-enable Button ---
            self.run_button.config(state=tk.NORMAL)


if __name__ == "__main__":
    # Check if processing_logic functions are available before starting GUI
    if 'read_pos_csv' not in globals() or \
       'process_excel' not in globals() or \
       'generate_output_filename' not in globals():
        # Error already shown by initial import attempt, just exit cleanly
        print("Exiting due to missing processing logic functions.")
    else:
        root = tk.Tk()
        app = ExcelProcessorApp(root)
        root.mainloop()