import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
import os
import threading
import xlwt


def convert_file(file_path):
    """Read CSV and write as real XLS (Excel 97-2003)."""
    filename = os.path.basename(file_path)
    input_dir = os.path.dirname(file_path)
    name_without_ext = os.path.splitext(filename)[0]

    output_dir = os.path.join(input_dir, "Converted")
    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, name_without_ext + ".xls")

    df = pd.read_csv(file_path, encoding="utf-8-sig")

    if df.empty:
        raise ValueError("CSV file is empty")

    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet("Sheet1")

    # Write header
    for col_idx, col_name in enumerate(df.columns):
        sheet.write(0, col_idx, str(col_name))

    # Write data
    for row_idx, row in df.iterrows():
        for col_idx, value in enumerate(row):
            try:
                if pd.isna(value):
                    sheet.write(row_idx + 1, col_idx, "")
                elif isinstance(value, (int, float)):
                    sheet.write(row_idx + 1, col_idx, value)
                else:
                    sheet.write(row_idx + 1, col_idx, str(value))
            except Exception:
                sheet.write(row_idx + 1, col_idx, str(value))

    workbook.save(output_path)
    return output_path


def run_conversion(file_paths, status_text, progress_var, convert_btn, root):
    total = len(file_paths)
    success_count = 0
    fail_count = 0

    for i, file_path in enumerate(file_paths):
        filename = os.path.basename(file_path)

        def log(msg):
            status_text.config(state="normal")
            status_text.insert(tk.END, msg + "\n")
            status_text.see(tk.END)
            status_text.config(state="disabled")
            root.update_idletasks()

        try:
            output_path = convert_file(file_path)
            log(f"[OK]  {filename}  ->  Converted/{os.path.basename(output_path)}")
            success_count += 1
        except Exception as e:
            log(f"[FAIL]  {filename}  ->  {str(e)}")
            fail_count += 1

        progress_var.set((i + 1) / total * 100)
        root.update_idletasks()

    status_text.config(state="normal")
    status_text.insert(tk.END, f"\nDone: {success_count} converted, {fail_count} failed.\n")
    status_text.see(tk.END)
    status_text.config(state="disabled")

    convert_btn.config(state="normal")
    messagebox.showinfo(
        "Conversion Complete",
        f"{success_count} file(s) converted successfully.\n"
        f"Saved to 'Converted' folder next to original files."
        + (f"\n{fail_count} file(s) failed." if fail_count else ""),
    )


def select_files():
    paths = filedialog.askopenfilenames(
        title="Select CSV files to convert",
        filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
    )
    if paths:
        selected_files.clear()
        selected_files.extend(paths)
        file_listbox.delete(0, tk.END)
        for p in paths:
            file_listbox.insert(tk.END, os.path.basename(p))
        files_label.config(text=f"{len(paths)} file(s) selected")


def start_conversion():
    if not selected_files:
        messagebox.showwarning("No Files", "Please select at least one file.")
        return

    status_text.config(state="normal")
    status_text.delete("1.0", tk.END)
    status_text.config(state="disabled")
    progress_var.set(0)
    convert_btn.config(state="disabled")

    thread = threading.Thread(
        target=run_conversion,
        args=(list(selected_files), status_text, progress_var, convert_btn, root),
        daemon=True,
    )
    thread.start()


# ── UI Setup ──────────────────────────────────────────────────────────────────

root = tk.Tk()
root.title("CSV to XLS Converter")
root.geometry("620x480")
root.resizable(True, True)
root.configure(bg="#f0f0f0")

selected_files = []

# Title
tk.Label(
    root, text="CSV to XLS Converter", font=("Segoe UI", 14, "bold"), bg="#f0f0f0"
).pack(pady=(12, 4))

tk.Label(
    root,
    text="Converts CSV files into real Excel 97-2003 (.xls) files.",
    font=("Segoe UI", 9),
    bg="#f0f0f0",
    fg="#555",
).pack()

# File selection frame
frame_select = tk.LabelFrame(root, text="Files", font=("Segoe UI", 9), bg="#f0f0f0", padx=8, pady=6)
frame_select.pack(fill="x", padx=12, pady=(10, 4))

top_row = tk.Frame(frame_select, bg="#f0f0f0")
top_row.pack(fill="x")

files_label = tk.Label(top_row, text="No files selected", bg="#f0f0f0", fg="#333", font=("Segoe UI", 9))
files_label.pack(side="left")

tk.Button(
    top_row, text="Browse...", command=select_files, font=("Segoe UI", 9), padx=8
).pack(side="right")

file_listbox = tk.Listbox(frame_select, height=5, font=("Segoe UI", 9), selectmode=tk.EXTENDED)
file_listbox.pack(fill="x", pady=(4, 0))

scrollbar = tk.Scrollbar(frame_select, orient="vertical", command=file_listbox.yview)
file_listbox.config(yscrollcommand=scrollbar.set)

# Convert button
convert_btn = tk.Button(
    root,
    text="Convert All",
    command=start_conversion,
    font=("Segoe UI", 11, "bold"),
    bg="#0078d4",
    fg="white",
    relief="flat",
    padx=12,
    pady=8,
    cursor="hand2",
)
convert_btn.pack(fill="x", padx=12, pady=6)

# Progress bar
progress_var = tk.DoubleVar()
ttk.Progressbar(root, variable=progress_var, maximum=100).pack(fill="x", padx=12, pady=(0, 4))

# Status log
frame_status = tk.LabelFrame(root, text="Status", font=("Segoe UI", 9), bg="#f0f0f0", padx=8, pady=6)
frame_status.pack(fill="both", expand=True, padx=12, pady=(0, 12))

status_text = tk.Text(
    frame_status, height=8, font=("Consolas", 9), state="disabled", bg="#1e1e1e", fg="#d4d4d4"
)
status_text.pack(fill="both", expand=True)

root.mainloop()
