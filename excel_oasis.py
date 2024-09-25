import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import openpyxl

# Function to browse and open Excel file 1 (data extraction)
def browse_file1():
    global file1_path
    file1_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file1_path:
        file1_label.config(text=f"File 1 Loaded: {file1_path.split('/')[-1]}")
    else:
        messagebox.showwarning("Warning", "No file selected")

# Function to browse and open Excel file 2 (formatting)
def browse_file2():
    global file2_path
    file2_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file2_path:
        file2_label.config(text=f"File 2 Loaded: {file2_path.split('/')[-1]}")
    else:
        messagebox.showwarning("Warning", "No file selected")

# Function to apply data from File 1 to File 2 and save the new Excel file
def apply_format():
    if not file1_path or not file2_path:
        messagebox.showwarning("Warning", "Please load both File 1 and File 2")
        return
    
    try:
        # Load both files into DataFrames
        data_df = pd.read_excel(file1_path)
        formatting_wb = openpyxl.load_workbook(file2_path)
        formatting_sheet = formatting_wb.active  # Assuming formatting is in the first sheet

        # Apply data from File 1 to File 2's format (assuming column names match)
        for col_idx, column in enumerate(data_df.columns, start=1):
            for row_idx, value in enumerate(data_df[column], start=2):  # Start from row 2 to avoid overwriting headers
                formatting_sheet.cell(row=row_idx, column=col_idx).value = value

        # Save the new formatted Excel file
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if save_path:
            formatting_wb.save(save_path)
            messagebox.showinfo("Success", f"File saved: {save_path}")
        else:
            messagebox.showwarning("Warning", "No save location selected")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

# Setting up GUI
root = tk.Tk()
root.title("Excel File Formatter")
root.geometry("600x400")

# File 1 upload (data extraction file)
file1_button = tk.Button(root, text="Browse File 1 (Data Extraction)", command=browse_file1)
file1_button.pack(pady=10)
file1_label = tk.Label(root, text="File 1 Not Loaded", fg="red")
file1_label.pack()

# File 2 upload (formatting file)
file2_button = tk.Button(root, text="Browse File 2 (Formatting File)", command=browse_file2)
file2_button.pack(pady=10)
file2_label = tk.Label(root, text="File 2 Not Loaded", fg="red")
file2_label.pack()

# Apply formatting and save
apply_button = tk.Button(root, text="Apply Format and Save", command=apply_format)
apply_button.pack(pady=20)

root.mainloop()
