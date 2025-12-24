"""
=======================================================================================================================
 Script Title   : nck_ui.py
 Description    : Professional GUI for NCK Inc. to automate ETABS/SAP2000 workflows,
                  interact with Excel spreadsheets, and run analysis scripts.
 Author         : Cyrielle Seymaux
 Date Created   : 2025-06-23
 Version        : 1.0.0
 Python Version : 3.10+
 Dependencies   : customtkinter, openpyxl, Pillow, tkinter
 Environment    : GITHUB_OWNER, GITHUB_REPO, GITHUB_TOKEN (Personal Access Token)
                  - Set these environment variables to enable GitHub issue creation.
                  - The PAT should have 'repo' scope for private repos or 'public_repo' for public ones.
=======================================================================================================================
"""

from __future__ import annotations

import os
import webbrowser

# Backend configuration (optional)
# If set, the app will send issues to this backend (which posts to GitHub with your secret PAT)
# Example: https://your-app.vercel.app/api/create-issue
BACKEND_URL = os.getenv("GITHUB_BACKEND_URL")
APP_KEY = os.getenv("APP_KEY")  # optional shared secret header
import tkinter as tk
import tkinter.filedialog as fd

import customtkinter as ctk
import requests
from PIL import Image

# Color palette extracted from NCK logo
PRIMARY_COLOR = "#00A3E0"   # Cyan blue from NCK logo
SECONDARY_COLOR = "#B0B0B0"  # Medium gray
DARK_COLOR = "#2D2D2D"       # Dark background
LIGHT_COLOR = "#FFFFFF"      # White text
DARK_GRAY = "#3A3A3A"        # Sidebar gray

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


class NCKApp(ctk.CTk):
    
    def __init__(self) -> None:
        super().__init__()

        self._setup_window()
        self._setup_sidebar()
        self._setup_main_content_and_tabs()

        # Track last opened Excel file path (optional)
        self.last_excel_path: str | None = None

    def _setup_window(self) -> None:
        # === Window Configuration ===
        self.title("NCK FEM Tool - Interface")
        self.geometry("1000x650")
        self.minsize(900, 580)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

    def _setup_sidebar(self) -> None:
        # === Sidebar (left) ===
        self.sidebar = ctk.CTkFrame(self, width=220, fg_color=DARK_GRAY)
        self.sidebar.grid(row=0, column=0, sticky="nswe")
        for r in range(10):
            self.sidebar.grid_rowconfigure(r, weight=0)
        self.sidebar.grid_rowconfigure(9, weight=1)  # push bottom items down

        # === NCK Logo ===
        logo_path = os.path.join(os.path.dirname(__file__), "nck_logo.png")
        if os.path.exists(logo_path):
            try:
                image = Image.open(logo_path).resize((84, 84), Image.LANCZOS)
                self.logo_img = ctk.CTkImage(light_image=image, dark_image=image, size=(84, 84))
                ctk.CTkLabel(self.sidebar, image=self.logo_img, text="").grid(row=0, column=0, padx=10, pady=(16, 8))
            except Exception:
                pass

        ctk.CTkLabel(
            self.sidebar,
            text="NCK TOOLKIT",
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color=PRIMARY_COLOR,
        ).grid(row=1, column=0, padx=20, pady=(0, 14))

        self.home_btn = ctk.CTkButton(self.sidebar, text="Home", command=self.show_home, fg_color=PRIMARY_COLOR)
        self.home_btn.grid(row=2, column=0, padx=20, pady=8)

        self.open_excel_btn = ctk.CTkButton(self.sidebar, text="Open Excel File", command=self.open_excel, fg_color=PRIMARY_COLOR)
        self.open_excel_btn.grid(row=3, column=0, padx=20, pady=8)

        self.run_script_btn = ctk.CTkButton(self.sidebar, text="Run Analysis", command=self.run_analysis, fg_color=PRIMARY_COLOR)
        self.run_script_btn.grid(row=4, column=0, padx=20, pady=8)

        self.support_btn = ctk.CTkButton(self.sidebar, text="Support", command=self.show_support, fg_color=PRIMARY_COLOR)
        self.support_btn.grid(row=5, column=0, padx=20, pady=8)

        self.reset_btn = ctk.CTkButton(self.sidebar, text="Reset Console", command=self.reset_console, fg_color="#CC0000")
        self.reset_btn.grid(row=6, column=0, padx=20, pady=8)

        # Spacer
        ctk.CTkLabel(self.sidebar, text="").grid(row=7, column=0, pady=4)

        # Footer-like info or version
        ctk.CTkLabel(
            self.sidebar,
            text="v1.1.0",
            font=ctk.CTkFont(size=12, weight="normal"),
            text_color=SECONDARY_COLOR,
        ).grid(row=8, column=0, padx=20, pady=(4, 16))

    def _setup_main_content_and_tabs(self) -> None:
        # === Main Content (right) ===
        self.main = ctk.CTkFrame(self, corner_radius=10, fg_color=DARK_COLOR)
        self.main.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")
        self.main.grid_rowconfigure(0, weight=1)
        self.main.grid_columnconfigure(0, weight=1)

        # Tabview for pages: Home / Support / Logs
        self.tabview = ctk.CTkTabview(self.main)
        self.tabview.grid(row=0, column=0, sticky="nsew")

        self.tab_home = self.tabview.add("Home")
        self.tab_support = self.tabview.add("Support")
        self.tab_logs = self.tabview.add("Logs")

        # Build each tab
        self._build_home_tab()
        self._build_support_tab()
        self._build_logs_tab()

    # ---------------------------------------------------------------------
    #  Tabs Builders
    # ---------------------------------------------------------------------
    def _build_home_tab(self) -> None:
        self.tab_home.grid_rowconfigure(0, weight=0)
        self.tab_home.grid_rowconfigure(1, weight=0)
        self.tab_home.grid_rowconfigure(2, weight=1)
        self.tab_home.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            self.tab_home,
            text="Welcome to the NCK FEM Automation Tool",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color=PRIMARY_COLOR,
        ).grid(row=0, column=0, pady=(28, 8), padx=10)

        ctk.CTkLabel(
            self.tab_home,
            text="Use the sidebar to start your tasks:",
            font=ctk.CTkFont(size=14),
            text_color=LIGHT_COLOR,
        ).grid(row=1, column=0, pady=(0, 18), padx=10)

        bullet_points = [
            "• Open and read Excel design files",
            "• Launch ETABS/SAP2000 analysis",
            "• Create GitHub support tickets",
            "• Review logs and results",
            "• Export formatted outputs (à venir)",
        ]

        bullets_frame = ctk.CTkFrame(self.tab_home, fg_color="transparent")
        bullets_frame.grid(row=2, column=0, sticky="n", pady=(0, 10))
        for idx, point in enumerate(bullet_points):
            ctk.CTkLabel(
                bullets_frame,
                text=point,
                font=ctk.CTkFont(size=13),
                text_color=SECONDARY_COLOR,
            ).grid(row=idx, column=0, sticky="w", padx=10, pady=2)

    def _build_support_tab(self) -> None:
        # Grid config
        for r in range(8):
            self.tab_support.grid_rowconfigure(r, weight=0)
        self.tab_support.grid_rowconfigure(6, weight=1)  # body text expands
        self.tab_support.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            self.tab_support,
            text="Créer un ticket GitHub",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color=PRIMARY_COLOR,
        ).grid(row=0, column=0, sticky="w", padx=20, pady=(20, 8))

        # Owner/Repo (UI fields; defaults from env)
        owner_default = os.getenv("GITHUB_OWNER", "")
        repo_default = os.getenv("GITHUB_REPO", "")
        self.owner_entry = ctk.CTkEntry(self.tab_support, placeholder_text="Owner (ex: CyrielleSeymaux)")
        if owner_default:
            self.owner_entry.insert(0, owner_default)
        self.owner_entry.grid(row=1, column=0, sticky="ew", padx=20, pady=4)

        self.repo_entry = ctk.CTkEntry(self.tab_support, placeholder_text="Repo (ex: ETABS-API)")
        if repo_default:
            self.repo_entry.insert(0, repo_default)
        self.repo_entry.grid(row=2, column=0, sticky="ew", padx=20, pady=4)

        # Title entry
        self.issue_title_entry = ctk.CTkEntry(self.tab_support, placeholder_text="Titre (ex: ETABS API - crash à l'ouverture)")
        self.issue_title_entry.grid(row=3, column=0, sticky="ew", padx=20, pady=6)

        # Labels entry
        self.issue_labels_entry = ctk.CTkEntry(self.tab_support, placeholder_text="Labels (ex: bug, etabs, urgent)")
        self.issue_labels_entry.grid(row=2, column=0, sticky="ew", padx=20, pady=6)

        # Description label + body
        ctk.CTkLabel(self.tab_support, text="Description", text_color=LIGHT_COLOR).grid(row=4, column=0, sticky="w", padx=20, pady=(10, 0))

        self.issue_body_text = ctk.CTkTextbox(self.tab_support, height=260)
        self.issue_body_text.insert(
            "end",
            "Étapes pour reproduire :\n1. ...\n2. ...\n\nRésultat observé :\n\nRésultat attendu :\n\nContexte : (version ETABS/SAP2000, OS, version script, etc.)\n",
        )
        self.issue_body_text.grid(row=4, column=0, sticky="nsew", padx=20, pady=6)

        # Include log checkbox
        self.include_log_var = tk.BooleanVar(value=True)
        ctk.CTkCheckBox(
            self.tab_support,
            text="Inclure le log (console) si disponible",
            variable=self.include_log_var,
        ).grid(row=5, column=0, sticky="w", padx=20, pady=6)

        # Buttons frame
        btn_frame = ctk.CTkFrame(self.tab_support, fg_color="transparent")
        btn_frame.grid(row=7, column=0, sticky="e", padx=20, pady=(8, 14))

        ctk.CTkButton(btn_frame, text="Créer le ticket", command=self.submit_support, fg_color=PRIMARY_COLOR).grid(row=0, column=0, padx=(0, 10))
        ctk.CTkButton(btn_frame, text="Réinitialiser", command=self._reset_support_form, fg_color=SECONDARY_COLOR, text_color=DARK_COLOR).grid(row=0, column=1, padx=(0,10))
        ctk.CTkButton(btn_frame, text="Ouvrir Issues", command=self._open_issues_page, fg_color=PRIMARY_COLOR).grid(row=0, column=2)

    def _build_logs_tab(self) -> None:
        self.tab_logs.grid_rowconfigure(0, weight=1)
        self.tab_logs.grid_columnconfigure(0, weight=1)

        # Console / Log textbox
        self.console = ctk.CTkTextbox(self.tab_logs, height=480)
        self.console.grid(row=0, column=0, sticky="nsew", padx=16, pady=16)

        # Hint on env vars
        hint = (
            "Astuce: définissez les variables d'environnement GITHUB_OWNER, GITHUB_REPO, GITHUB_TOKEN\n"
            "pour activer la création d'issues. Le PAT doit avoir le scope 'repo' (ou 'public_repo' pour un repo public)."
        )
        ctk.CTkLabel(self.tab_logs, text=hint, text_color=SECONDARY_COLOR, wraplength=760, justify="left").grid(row=1, column=0, sticky="w", padx=16, pady=(0, 12))

    # ---------------------------------------------------------------------
    #  Navigation helpers
    # ---------------------------------------------------------------------
    def show_home(self) -> None:
        """
        Switches the current tab view to the 'Home' tab.

        This method sets the active tab in the tab view to 'Home', updating the user interface accordingly.
        """
        self.tabview.set("Home")

    def show_support(self) -> None:
        self.tabview.set("Support")

    # ---------------------------------------------------------------------
    #  Actions
    # ---------------------------------------------------------------------
    def open_excel(self) -> None:
        filepath = fd.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if filepath:
            self.last_excel_path = filepath
            self.log(f"Opened Excel file: {os.path.basename(filepath)}")
            # TODO: add Excel logic here

    def run_analysis(self) -> None:
        self.log("Launching analysis script...")
        # TODO: replace with your script execution logic
        # Example: call ETABS/SAP2000 automation here
        self.log("✅ Analysis complete.")

    def reset_console(self) -> None:
        if hasattr(self, "console"):
            self.console.delete("1.0", "end")
            self.log("Console cleared.")

    # ---------------------------------------------------------------------
    #  Support (GitHub Issues)
    # ---------------------------------------------------------------------
    def submit_support(self) -> None:
        # Prefer UI values; fall back to env
        owner = (self.owner_entry.get() or os.getenv("GITHUB_OWNER") or "").strip()
        repo = (self.repo_entry.get() or os.getenv("GITHUB_REPO") or "").strip()
        token = os.getenv("GITHUB_TOKEN")

        if not all([owner, repo, token]):
            self._notify("⚠️ Renseignez Owner/Repo et configurez GITHUB_TOKEN (PAT) en variable d'environnement.", is_error=True)
            self.tabview.set("Logs")
            return

        title = (self.issue_title_entry.get() or "").strip()
        if not title:
            self._notify("⚠️ Le titre du ticket est requis.", is_error=True)
            return

        labels_text = (self.issue_labels_entry.get() or "").strip()
        labels = [lbl.strip() for lbl in labels_text.split(",") if lbl.strip()] if labels_text else []

        body = (self.issue_body_text.get("1.0", "end") or "").strip()

        # Append console logs if requested (robust to encoding / empty text)
        body = self._append_console_snippet(body)

        try:
            issue_url = self.create_github_issue(owner, repo, token, title, body, labels)
            self._notify(f"✅ Ticket créé : {issue_url}")
            webbrowser.open_new_tab(issue_url)
            self._reset_support_form()
        except Exception as e:
            self._notify(f"❌ Échec de création du ticket : {e}", is_error=True)

    @staticmethod
    def create_github_issue(owner: str, repo: str, token: str, title: str, body: str, labels: list[str] | None = None) -> str:
        url = f"https://api.github.com/repos/{owner}/{repo}/issues"
        headers = {
            # Classic PAT: 'token'; Fine-grained PAT also works with 'Bearer'
            "Authorization": f"token {token}",
            "Accept": "application/vnd.github+json",
            "X-GitHub-Api-Version": "2022-11-28",
        }
        payload: dict = {"title": title, "body": body}
        if labels:
            payload["labels"] = labels

        resp = requests.post(url, headers=headers, json=payload, timeout=20)

        # Helpful error mapping
        if resp.status_code == 401:
            raise RuntimeError("401 Unauthorized - PAT invalide ou manquant. Vérifiez GITHUB_TOKEN et ses scopes.")
        if resp.status_code == 403:
            # Could be SSO not authorized, rate limit, or missing permissions
            try:
                j = resp.json()
                if "SSO" in j.get("message", ""):
                    raise RuntimeError("403 Forbidden - SSO non autorisé pour l'organisation. Autorisez le token.")
            except Exception:
                pass
            raise RuntimeError("403 Forbidden - Permissions insuffisantes ou rate limit.")
        if resp.status_code == 404:
            raise RuntimeError("404 Not Found - Owner/Repo incorrect (ou pas d'accès).")
        if resp.status_code == 410:
            raise RuntimeError("410 Gone - Les Issues sont désactivées pour ce repo (Settings > Features > Issues).")
        if resp.status_code not in (200, 201):
            try:
                j = resp.json()
                raise RuntimeError(f"{resp.status_code} - {j.get('message')} ({j})")
            except Exception:
                resp.raise_for_status()

        issue = resp.json()
        return issue.get("html_url", f"https://github.com/{owner}/{repo}/issues")

    def _reset_support_form(self) -> None:
        self.issue_title_entry.delete(0, "end")
        self.issue_labels_entry.delete(0, "end")
        self.issue_body_text.delete("1.0", "end")
        self.issue_body_text.insert(
            "end",
            "Étapes pour reproduire :\n1. ...\n2. ...\n\nRésultat observé :\n\nRésultat attendu :\n\nContexte : (version ETABS/SAP2000, OS, version script, etc.)\n",
        )
        self.include_log_var.set(True)

    # ---------------------------------------------------------------------
    #  Utilities
    # ---------------------------------------------------------------------
    def _append_console_snippet(self, body: str) -> str:
        if not getattr(self, "include_log_var", None) or not self.include_log_var.get():
            return body
        if not hasattr(self, "console"):
            return body
        try:
            log_content = self.console.get("1.0", "end-1c") or ""
        except Exception:
            return body
        log_snippet = log_content[-4000:]
        if not log_snippet:
            return body
        body += "\n\n---\n**Console log**\n```\n" + log_snippet + "\n```"
        return body

    def log(self, message: str) -> None:
        if hasattr(self, "console") and self.console is not None:
            self.console.insert("end", message + "\n")
            self.console.see("end")
        print(message)

    def _notify(self, msg: str, is_error: bool = False) -> None:
        # Prefer console when available
        self.log(msg)
        # Minimal popup fallback for visibility
        popup = ctk.CTkToplevel(self)
        popup.title("Notification")
        popup.geometry("520x180")
        popup.transient(self)
        popup.grab_set()

        ctk.CTkLabel(
            popup,
            text=msg,
            text_color=("#FF6B6B" if is_error else PRIMARY_COLOR),
            wraplength=460,
            justify="left",
        ).pack(padx=20, pady=(20, 12))
        ctk.CTkButton(popup, text="OK", command=popup.destroy).pack(pady=(0, 16))

    def _open_issues_page(self) -> None:
        owner = (self.owner_entry.get() or os.getenv("GITHUB_OWNER") or "").strip()
        repo = (self.repo_entry.get() or os.getenv("GITHUB_REPO") or "").strip()
        if owner and repo:
            webbrowser.open_new_tab(f"https://github.com/{owner}/{repo}/issues")
        else:
            self._notify("⚠️ Renseignez Owner et Repo pour ouvrir la page Issues.", is_error=True)


if __name__ == "__main__":
    app = NCKApp()
    app.mainloop()


aadda