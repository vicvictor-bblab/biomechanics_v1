import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd


def integrate_files():
    """Combine multiple CSV/Excel files into one CSV file."""
    root = tk.Tk()
    root.withdraw()

    file_paths = filedialog.askopenfilenames(
        title="Select files to integrate",
        filetypes=[
            ("CSV and Excel", ("*.csv", "*.xlsx", "*.xls")),
            ("All files", "*.*"),
        ],
    )
    if not file_paths:
        messagebox.showinfo("Integration cancelled", "No files were selected.")
        root.destroy()
        return

    frames = []
    for path in file_paths:
        try:
            if path.lower().endswith((".xlsx", ".xls")):
                df = pd.read_excel(path)
            else:
                df = pd.read_csv(path)
            frames.append(df)
        except Exception as e:
            messagebox.showerror("Read error", f"Failed to read {path}:\n{e}")
            root.destroy()
            return

    try:
        combined = pd.concat(frames, ignore_index=True)
    except Exception as e:
        messagebox.showerror("Combine error", f"Failed to combine files:\n{e}")
        root.destroy()
        return

    save_path = filedialog.asksaveasfilename(
        title="Save merged CSV",
        defaultextension=".csv",
        filetypes=[("CSV", "*.csv")],
    )
    if not save_path:
        messagebox.showinfo("Save cancelled", "Integration result was not saved.")
        root.destroy()
        return

    try:
        combined.to_csv(save_path, index=False)
        messagebox.showinfo("Success", f"Integrated file saved to:\n{save_path}")
    except Exception as e:
        messagebox.showerror("Save error", f"Failed to save file:\n{e}")
    finally:
        root.destroy()


if __name__ == "__main__":
    integrate_files()
