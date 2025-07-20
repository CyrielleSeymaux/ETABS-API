"""
===============================================================================
 Script Title   : nck_ui.py
 Description    : Professional GUI for NCK Inc. to automate ETABS/SAP2000 workflows,
                  interact with Excel spreadsheets, and run analysis scripts.
 Author         : Cyrielle Seymaux
 Date Created   : 2025-06-23
 Version        : 1.0.0
 Python Version : 3.10+
 Dependencies   : customtkinter, openpyxl, Pillow (for image support)
===============================================================================
"""

import customtkinter as ctk
import tkinter.filedialog as fd
from PIL import Image, ImageTk
import os

# Color palette extracted from NCK logo
PRIMARY_COLOR = "#00A3E0"   # Cyan blue
SECONDARY_COLOR = "#B0B0B0"  # Medium gray (increased visibility)
DARK_COLOR = "#2D2D2D"       # Less dark background for contrast
LIGHT_COLOR = "#FFFFFF"      # Clear white for legibility
DARK_GRAY = "#3A3A3A"        # Sidebar gray

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


class NCKApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # === Window Configuration ===
        self.title("NCK FEM Tool - Interface")
        self.geometry("900x600")
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # === Sidebar Frame ===
        self.sidebar = ctk.CTkFrame(self, width=200, fg_color=DARK_GRAY)
        self.sidebar.grid(row=0, column=0, sticky="nswe")
        self.sidebar.grid_rowconfigure(7, weight=1)

        # === Add NCK logo ===
        logo_path = os.path.join(os.path.dirname(__file__), "nck_logo.png")
        if os.path.exists(logo_path):
            image = Image.open(logo_path)
            image = image.resize((80, 80), Image.LANCZOS)
            self.logo_img = ctk.CTkImage(light_image=image, dark_image=image, size=(80, 80))
            logo_label = ctk.CTkLabel(self.sidebar, image=self.logo_img, text="")
            logo_label.grid(row=0, column=0, padx=10, pady=(15, 5))

        label_title = ctk.CTkLabel(self.sidebar, text="NCK TOOLKIT",
                                   font=ctk.CTkFont(size=18, weight="bold"),
                                   text_color=PRIMARY_COLOR)
        label_title.grid(row=1, column=0, padx=20, pady=(0, 10))

        self.home_btn = ctk.CTkButton(self.sidebar, text="Home", command=self.show_home, fg_color=PRIMARY_COLOR)
        self.home_btn.grid(row=2, column=0, padx=20, pady=10)

        self.open_excel_btn = ctk.CTkButton(self.sidebar, text="Open Excel File", command=self.open_excel, fg_color=PRIMARY_COLOR)
        self.open_excel_btn.grid(row=3, column=0, padx=20, pady=10)

        self.run_script_btn = ctk.CTkButton(self.sidebar, text="Run Analysis", command=self.run_analysis, fg_color=PRIMARY_COLOR)
        self.run_script_btn.grid(row=4, column=0, padx=20, pady=10)

        self.reset_btn = ctk.CTkButton(self.sidebar, text="Reset Console", command=self.reset_console, fg_color="#CC0000")
        self.reset_btn.grid(row=5, column=0, padx=20, pady=10)

        # === Main Content ===
        self.main = ctk.CTkFrame(self, corner_radius=10, fg_color=DARK_COLOR)
        self.main.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")
        self.main.grid_rowconfigure(1, weight=1)
        self.main.grid_columnconfigure(0, weight=1)

        self.show_home()

    def show_home(self):
        for widget in self.main.winfo_children():
            widget.destroy()

        label_title = ctk.CTkLabel(self.main, text="Welcome to the NCK FEM Automation Tool",
                                   font=ctk.CTkFont(size=20, weight="bold"), text_color=PRIMARY_COLOR)
        label_title.grid(row=0, column=0, pady=(30, 10), padx=10)

        label_sub = ctk.CTkLabel(self.main, text="Use the sidebar to start your tasks:",
                                 font=ctk.CTkFont(size=14), text_color=LIGHT_COLOR)
        label_sub.grid(row=1, column=0, pady=(0, 20), padx=10)

        bullet_points = [
            "• Open and read Excel design files",
            "• Launch ETABS/SAP2000 analysis",
            "• Review logs and results",
            "• Export formatted outputs"
        ]

        for idx, point in enumerate(bullet_points):
            bullet_label = ctk.CTkLabel(self.main, text=point,
                                        font=ctk.CTkFont(size=13), text_color=SECONDARY_COLOR)
            bullet_label.grid(row=2+idx, column=0, sticky="w", padx=30, pady=2)

    def open_excel(self):
        filepath = fd.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if filepath:
            self.log(f"Opened Excel file: {os.path.basename(filepath)}")
            # Excel logic goes here

    def run_analysis(self):
        self.log("Launching analysis script...")
        # Replace with your script execution logic
        self.log("✅ Analysis complete.")

    def reset_console(self):
        if hasattr(self, 'console'):
            self.console.delete("1.0", "end")
            self.log("Console cleared.")

    def log(self, message):
        if hasattr(self, 'console'):
            self.console.insert("end", message + "\n")
            self.console.see("end")


if __name__ == "__main__":
    app = NCKApp()
    app.mainloop()



aadda