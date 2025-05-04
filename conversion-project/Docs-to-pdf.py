import os
from tkinter import filedialog, messagebox, Tk, Button, Label
from docx2pdf import convert

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

def convert_docs():
    if not input_folder or not output_folder:
        messagebox.showerror("Error", "Please select both folders.")
        return

    docx_files = [f for f in os.listdir(input_folder) if f.endswith(".docx")]

    if not docx_files:
        messagebox.showinfo("No DOCX Found", "No Word (.docx) files found.")
        return

    for file in docx_files:
        input_path = os.path.join(input_folder, file)
        output_path = os.path.join(output_folder, file.replace(".docx", ".pdf"))

        try:
            convert(input_path, output_path)
        except Exception as e:
            print(f"Error converting {file}: {e}")

    messagebox.showinfo("Success", "All DOCX files converted to PDF!")

# GUI setup
app = Tk()
app.title("Bulk DOCX to PDF Converter")
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

Button(app, text="Convert DOCX to PDF", command=convert_docs, bg="purple", fg="white").pack(pady=20)

app.mainloop()
