import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinterdnd2 import DND_FILES, TkinterDnD
from PyPDF2 import PdfReader
import os
import time

def open_file():
    filepath = filedialog.askopenfilename(filetypes=[("PDF files","*.pdf"), ("All Files"),("*.*")])
    if filepath:
        selected_file.set(filepath)
        
def drop(event):
    file_patch = event.data
    if file_patch.starwith("{") and file_patch.endswith("}"):
        file_patch = file_patch[1:-1]
    selected_file.set(file_patch)

def choose_directory():
    directory = filedialog.askdirectory()
    output_dir.set(directory)

def start_conversion():
    try:
        pd_path = selected_file.get()
        if not pdf_path:
            messagebox.showerror("Error","PDF file not selected")
            return
        if not output_file.get():
            raise ValueError("output file name cannot be empty")
        output_dir_patch = output_dir.get()
        
        if not output_dir_patch:
            messagebox.showerror("Error","No output directory selected")
    except Exception as e:
        messagebox.showerror("Error",str(e))
    
#
#
# MAIN WINDOW
#
#
#
root = TkinterDnD.Tk()
root.title("PDF converter")
root.geometry("500x270")

root.drop_target_register(DND_FILES)
root.dnd_bind("<<Drop>>", drop)

output_dir = tk.StringVar
output_file = tk.StringVar
selected_file = tk.StringVar

tk.Label(root, text="Output directory").pack()
tk.Entry(root, text=output_dir).pack()
tk.Button(root, text="Browse", command=choose_directory).pack()

tk.Label(root, text="Output file name").pack()
tk.Entry(root, textvariable=output_file).pack()

tk.Label(root, text="Select or drop PDF file").pack()
tk.Entry(root, textvariable=selected_file, state="readonly").pack()

open_button = tk.Button(root, text="Open PDF", command=open_file)
open_button.pack()

convert_button = tk.Button(root, text="Convert", command=start_conversion)

root.mainloop()