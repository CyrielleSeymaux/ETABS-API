"""
===============================================================================
 Script Title   : NCK_ui.py
 Description    : Modern graphical user interface (GUI) for launching and 
                  monitoring numerical analysis scripts (FEM, API-based),
                  with a sleek and intuitive dark design.

 Author         : Cyrielle Seymaux
 Date Created   : 2025-06-23
 Version        : 1.0.0
 Python Version : 3.10+
 Dependencies   : ttkbootstrap, tkinter (standard), openpyxl (for Excel support)

 Usage          :
     This interface was developed for NCK Inc. to provide engineers with a fast, 
     user-friendly, and centralized tool to:
     - run Python-based numerical analyses connected to CSI software APIs 
       (ETABS, SAP2000, SAFE)
     - read from and write to internal Excel spreadsheets used in structural 
       design calculations (e.g. stiffness, loads, modal results)

 Installation & Execution:
     > pip install ttkbootstrap openpyxl
     > python badass_ui.py

 Notes:
     - Easily extendable with additional modules (tabs, plotting, model loading)
     - Integrates with NCK Inc.'s internal design workflows and automation scripts
     - Excel I/O support allows automatic updates and data retrieval from 
       design spreadsheets
===============================================================================
"""

# ========================== Imports & Dependencies ========================== #
import ttkbootstrap as tb
from ttkbootstrap.constants import *
import tkinter as tk
from tkinter.scrolledtext import ScrolledText
import threading
import time

# ============================ Interface Class =============================== #
class NCKUI:
    def __init__(self, root):
        self.root = root
        self.root.title("🔥 FEM Analysis Console - NCK Inc.")
        self.root.geometry("700x500")
        self.style = tb.Style("darkly")

        self._create_title()
        self._create_input()
        self._create_buttons()
        self._create_console()
        self.log("System ready. Awaiting your commands, engineer...")

    # -------------------------- UI Sections --------------------------------- #
    def _create_title(self):
        title = tb.Label(self.root, text="🚀 FEM Model Control Center",
                         font=("Segoe UI", 20, "bold"), bootstyle="info")
        title.pack(pady=15)

    def _create_input(self):
        self.entry = tb.Entry(self.root, font=("Segoe UI", 12), width=50)
        self.entry.pack(pady=10)

    def _create_buttons(self):
        btn_frame = tb.Frame(self.root)
        btn_frame.pack(pady=10)

        self.run_btn = tb.Button(btn_frame, text="Run Analysis",
                                 command=self.run_analysis, bootstyle="success-outline")
        self.run_btn.pack(side=LEFT, padx=10)

        self.reset_btn = tb.Button(btn_frame, text="Reset",
                                   command=self.reset_fields, bootstyle="danger-outline")
        self.reset_btn.pack(side=LEFT, padx=10)

    def _create_console(self):
        self.console = ScrolledText(self.root, height=15, font=("Consolas", 10),
                                    bg="#1e1e1e", fg="#00FF00", insertbackground="white")
        self.console.pack(fill=BOTH, expand=True, padx=20, pady=10)

    # -------------------------- Event Handlers ------------------------------ #
    def run_analysis(self):
        command = self.entry.get()
        self.log(f"> {command}")
        self.entry.delete(0, tk.END)
        threading.Thread(target=self._fake_processing, daemon=True).start()

    def _fake_processing(self):
        self.log("Processing...")
        for i in range(5):
            self.log(f"Running iteration {i+1}/5...")
            time.sleep(0.5)
        self.log("✅ Analysis complete!")

    def reset_fields(self):
        self.entry.delete(0, tk.END)
        self.console.delete(1.0, tk.END)
        self.log("Console reset.")

    # ---------------------------- Logger ------------------------------------ #
    def log(self, message):
        self.console.insert(tk.END, message + "\n")
        self.console.see(tk.END)

# ============================== Main Script ================================ #
if __name__ == "__main__":
    root = tb.Window(themename="darkly")
    app = NCKUI(root)
    root.mainloop()


errwwte