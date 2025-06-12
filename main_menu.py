import tkinter as tk
from tkinter import messagebox

from bio_graph_app import BioGraphApp
import integrate_files


def main():
    """Prompt the user for file integration or GUI launch."""
    # Create a temporary root window for the dialog
    root = tk.Tk()
    root.withdraw()

    answer = messagebox.askquestion(
        "Select Mode",
        "Do you want to integrate files?\n\nYes: integrate files\nNo: launch GUI",
    )
    root.destroy()

    if answer == 'yes':
        integrate_files.integrate_files()
    else:
        gui_root = tk.Tk()
        app = BioGraphApp(gui_root)
        gui_root.mainloop()


if __name__ == "__main__":
    main()
