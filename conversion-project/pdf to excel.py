import os
import pdfplumber
from openpyxl import Workbook
from tkinter import filedialog, messagebox, Tk, Button, Label

def select_input_folder():
    folder = filedialog.askdirectory()
    if folder:
        input_label.config(text=folder)
        global input_folder
        input_folder = folder

def select_output_folder():
    folder = filedialog.askdirectory()
    if folder:
        output_label.config(text=folder)
        global output_folder
        output_folder = folder

def convert_pdf_to_excel(pdf_path, excel_path):
    with pdfplumber.open(pdf_path) as pdf:
        wb = Workbook()
        ws = wb.active

        for page in pdf.pages:
            table = page.extract_table()
            if table:
                for row in table:
                    ws.append(row)

        wb.save(excel_path)

def convert_files():
    if not input_folder or not output_folder:
        messagebox.showerror("Error", "Please select both input and output folders.")
        return

    pdf_files = [f for f in os.listdir(input_folder) if f.endswith('.pdf')]
    if not pdf_files:
        messagebox.showinfo("No PDFs Found", "No PDF files found in the input folder.")
        return

    for file in pdf_files:
        pdf_path = os.path.join(input_folder, file)
        excel_name = file.replace('.pdf', '.xlsx')
        excel_path = os.path.join(output_folder, excel_name)

        try:
            convert_pdf_to_excel(pdf_path, excel_path)
        except Exception as e:
            print(f"Failed to convert {file}: {e}")

    messagebox.showinfo("Done", "All files converted to Excel!")

# GUI
app = Tk()
app.title("Bulk PDF to Excel Converter")
app.geometry("400x250")

input_folder = ''
output_folder = ''

Label(app, text="Select Input Folder:").pack(pady=(10, 0))
Button(app, text="Browse", command=select_input_folder).pack()
input_label = Label(app, text="No folder selected", fg="gray")
input_label.pack()

Label(app, text="Select Output Folder:").pack(pady=(10, 0))
Button(app, text="Browse", command=select_output_folder).pack()
output_label = Label(app, text="No folder selected", fg="gray")
output_label.pack()

Button(app, text="Convert PDFs to Excel", command=convert_files, bg="blue", fg="white").pack(pady=20)

app.mainloop()
