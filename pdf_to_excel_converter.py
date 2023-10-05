import tkinter as tk
from tkinter import filedialog, messagebox
import tabula
import pandas as pd
import os

def pdf_to_excel_tabula(pdf_path, excel_path):
    tables = tabula.read_pdf(pdf_path, pages='all', multiple_tables=True)
    with pd.ExcelWriter(excel_path) as writer:
        for i, table in enumerate(tables):
            table.to_excel(writer, sheet_name=f"Table_{i+1}", index=False)

def convert_pdf_to_excel():
    pdf_path = filedialog.askopenfilename(title="Select PDF File", filetypes=[("PDF Files", "*.pdf")])
    if not pdf_path:
        return

    excel_path = os.path.splitext(pdf_path)[0] + ".xlsx"
    try:
        pdf_to_excel_tabula(pdf_path, excel_path)
        messagebox.showinfo("Success", f"File converted successfully!\nSaved at: {excel_path}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

app = tk.Tk()
app.title("PDF to Excel Converter")

frame = tk.Frame(app, padx=20, pady=20)
frame.pack(padx=10, pady=10)

label = tk.Label(frame, text="PDF to Excel Converter", font=("Arial", 16))
label.pack(pady=20)

convert_btn = tk.Button(frame, text="Convert PDF to Excel", command=convert_pdf_to_excel)
convert_btn.pack()

app.mainloop()
