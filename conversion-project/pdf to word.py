import os
from tkinter import filedialog, messagebox, Tk, Button, Label
from pdf2docx import Converter

def select_input_folder():
    folder = filedialog.askdirectory()
    if folder:
        input_folder_label.config(text=folder)
        global input_folder
        input_folder = folder

def select_output_folder():
    folder = filedialog.askdirectory()
    if folder:
        output_folder_label.config(text=folder)
        global output_folder
        output_folder = folder

def convert_pdfs():
    if not input_folder or not output_folder:
        messagebox.showerror("Error", "Please select both input and output folders.")
        return

    files = [f for f in os.listdir(input_folder) if f.endswith('.pdf')]
    if not files:
        messagebox.showinfo("No PDFs Found", "No PDF files found in the selected input folder.")
        return

    for filename in files:
        pdf_path = os.path.join(input_folder, filename)
        docx_name = filename.replace('.pdf', '.docx')
        docx_path = os.path.join(output_folder, docx_name)

        try:
            cv = Converter(pdf_path)
            cv.convert(docx_path, start=0, end=None)
            cv.close()
        except Exception as e:
            print(f"Error converting {filename}: {e}")

    messagebox.showinfo("Done", "All files converted successfully!")

# GUI Setup
app = Tk()
app.title("Bulk PDF to Word Converter")
app.geometry("400x250")

input_folder = ''
output_folder = ''

Label(app, text="Select Input Folder:").pack(pady=(10, 0))
Button(app, text="Browse", command=select_input_folder).pack()
input_folder_label = Label(app, text="No folder selected", fg="gray")
input_folder_label.pack()

Label(app, text="Select Output Folder:").pack(pady=(10, 0))
Button(app, text="Browse", command=select_output_folder).pack()
output_folder_label = Label(app, text="No folder selected", fg="gray")
output_folder_label.pack()

Button(app, text="Convert PDFs to Word", command=convert_pdfs, bg="green", fg="white").pack(pady=20)

app.mainloop()
