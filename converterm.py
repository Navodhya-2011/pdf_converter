import os
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox

# Function to select files to convert
def select_files():
    filetypes = (("PowerPoint files", "*.pptx"), ("Word documents", "*.docx"))
    filenames = filedialog.askopenfilenames(
        title="Open files", initialdir="/", filetypes=filetypes
    )
    file_listbox.delete(0, tk.END)  # Clear the listbox
    for filename in filenames:
        file_listbox.insert(tk.END, filename)

# Function to convert the selected files to PDF
def convert_to_pdf():
    filepaths = file_listbox.get(0, tk.END)
    if not filepaths:
        messagebox.showerror("Error", "Please select at least one file to convert.")
        return

    # Ask for directory to save converted PDFs
    save_directory = filedialog.askdirectory(title="Select directory to save PDFs")
    if not save_directory:
        return

    for filepath in filepaths:
        try:
            filename = os.path.basename(filepath)  # Get the file name with extension
            # Create the output PDF path with the same base name
            pdf_path = os.path.join(save_directory, os.path.splitext(filename)[0] + ".pdf")
            convert_file_to_pdf(filepath, pdf_path)
            messagebox.showinfo("Success", f"Converted '{filename}' to PDF successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to convert '{filename}' to PDF. Error: {e}")

# Function to convert a file to PDF using unoconv
def convert_file_to_pdf(input_path, output_path):
    try:
        subprocess.run(['unoconv', '-f', 'pdf', '-o', output_path, input_path], check=True)
    except subprocess.CalledProcessError as e:
        raise Exception(f"Conversion failed: {e}")

# Setup GUI .
root = tk.Tk()
root.title("File Converter")

frame = tk.Frame(root)
frame.pack(padx=10, pady=10)

file_label = tk.Label(frame, text="Selected Files:")
file_label.grid(row=0, column=0, padx=5, pady=5)

file_listbox = tk.Listbox(frame, selectmode=tk.MULTIPLE, width=50)
file_listbox.grid(row=1, column=0, columnspan=2, padx=5, pady=5)

browse_button = tk.Button(frame, text="Browse", command=select_files)
browse_button.grid(row=2, column=0, padx=5, pady=5)

convert_button = tk.Button(root, text="Convert to PDF", command=convert_to_pdf)
convert_button.pack(pady=20)

root.mainloop()
