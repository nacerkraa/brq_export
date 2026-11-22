"""Application entry point."""
import os
import tkinter as tk
from config import OUTPUT_FOLDER, WORK_FOLDER
from gui import BrqApp


def main():
    os.makedirs(WORK_FOLDER, exist_ok=True)
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    root = tk.Tk()
    BrqApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
