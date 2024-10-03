import os
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox


# Function to select the file to convert
def select_file():
    filetypes = (("PowerPoint files", "*.pptx"), ("Word documents", "*.docx"))
    filename = filedialog.askopenfilename(
        title="Open file", initialdir="/", filetypes=filetypes
    )
    file_entry.delete(0, tk.END)
    file_entry.insert(0, filename)


# Function to convert the selected file to PDF
def convert_to_pdf():
    filepath = file_entry.get()
    if not filepath:
        messagebox.showerror("Error", "Please select a file to convert.")
        return

    if not filepath.lower().endswith(('.pptx', '.docx')):
        messagebox.showerror("Error", "Please select a .pptx or .docx file.")
        return

    save_path = filedialog.asksaveasfilename(
        defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")]
    )
    if not save_path:
        return

    try:
        convert_file_to_pdf(filepath, save_path)
        messagebox.showinfo("Success", f"Converted '{os.path.basename(filepath)}' to PDF successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to convert file to PDF. Error: {e}")


# Function to convert file to PDF using unoconv
def convert_file_to_pdf(input_path, output_path):
    try:
        subprocess.run(['unoconv', '-f', 'pdf', '-o', output_path, input_path], check=True)
    except subprocess.CalledProcessError as e:
        raise Exception(f"Conversion failed: {e}")


# Setup 
root = tk.Tk()
root.title("File Converter")

frame = tk.Frame(root)
frame.pack(padx=10, pady=10)

file_label = tk.Label(frame, text="Select File:")
file_label.grid(row=0, column=0, padx=5, pady=5)

file_entry = tk.Entry(frame, width=40)
file_entry.grid(row=0, column=1, padx=5, pady=5)

browse_button = tk.Button(frame, text="Browse", command=select_file)
browse_button.grid(row=0, column=2, padx=5, pady=5)

convert_button = tk.Button(root, text="Convert to PDF", command=convert_to_pdf)
convert_button.pack(pady=20)

root.mainloop()
