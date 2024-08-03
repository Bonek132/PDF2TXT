import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinterdnd2 import DND_FILES, TkinterDnD
from PyPDF2 import PdfReader
import os
import time

def open_file():
    filepath = filedialog.askopenfilename(filetypes=[("PDF files","*.pdf"), ("All Files","*.*")])
    if filepath:
        selected_file.set(filepath)
        
def drop(event):
    file_path = event.data
    if file_path.startswith("{") and file_path.endswith("}"):
        file_path = file_path[1:-1]
    selected_file.set(file_path)

def choose_directory():
    directory = filedialog.askdirectory()
    output_dir.set(directory)

def start_conversion():
    try:
        pdf_path = selected_file.get()
        if not pdf_path:
            messagebox.showerror("Error", "PDF file not selected")
            return
        if not output_file.get():
            raise ValueError("Output file name cannot be empty")
        output_dir_path = output_dir.get()
        if not output_dir_path:
            messagebox.showerror("Error", "No output directory selected")
            return
        
        progress['value'] = 0
        output_path = os.path.join(output_dir_path, f"{output_file.get()}.txt")
        with open(pdf_path, "rb") as pdf_file:
            reader = PdfReader(pdf_file)
            total_pages = len(reader.pages)
            
            with open(output_path, "w", encoding="utf-8") as text_file:
                for i, page in enumerate(reader.pages):
                    progress["value"] = ((i + 1) / total_pages) * 100
                    root.update_idletasks()
                    time.sleep(0.5)
                    text = page.extract_text()
                    
                    if text:
                        text_file.write(f"Page {i + 1}\n{'=' * 20}\n")
                        text_file.write(text)
                        text_file.write("\n\n")
        
        messagebox.showinfo("Success", "File converted successfully!")
        
    except Exception as e:
        messagebox.showerror("Error", str(e))

# MAIN WINDOW

root = TkinterDnD.Tk()
root.title("PDF 2 TXT converter")
root.geometry("500x400")

root.drop_target_register(DND_FILES)
root.dnd_bind("<<Drop>>", drop)

output_dir = tk.StringVar()
output_file = tk.StringVar()
selected_file = tk.StringVar()

tk.Label(root, text="Output directory").pack(pady=5)  # Zwiększony odstęp pionowy
tk.Entry(root, textvariable=output_dir).pack(padx=10, pady=5)  # Zwiększone odstępy poziome i pionowe
tk.Button(root, text="Browse", command=choose_directory).pack(pady=5)  # Zwiększony odstęp pionowy

tk.Label(root, text="Output file name").pack(pady=5)  # Zwiększony odstęp pionowy
tk.Entry(root, textvariable=output_file).pack(padx=10, pady=5)  # Zwiększone odstępy poziome i pionowe

tk.Label(root, text="Select or drop PDF file").pack(pady=5)  # Zwiększony odstęp pionowy
tk.Entry(root, textvariable=selected_file, state="readonly").pack(padx=10, pady=5)  # Zwiększone odstępy poziome i pionowe

open_button = tk.Button(root, text="Open PDF", command=open_file)
open_button.pack(pady=5)  # Zwiększony odstęp pionowy

convert_button = tk.Button(root, text="Convert", command=start_conversion)
convert_button.pack(pady=5)  # Zwiększony odstęp pionowy

progress = ttk.Progressbar(root, orient="horizontal", length=200, mode="determinate")
progress.pack(pady=10)  # Zwiększony odstęp pionowy

root.mainloop()
