"""
===============================================================================
 Script Title   : NCK_OpenAI.py
 Description    : Test Script for integrating OpenAI API with NCK Inc.'s
                  structural engineering tools and workflows.

 Author         : Cyrielle Seymaux
 Date Created   : 2026-01-01
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
     > python NCK_OpenAI.py

 Notes:
     - Easily extendable with additional modules (tabs, plotting, model loading)
     - Integrates with NCK Inc.'s internal design workflows and automation scripts
     - Excel I/O support allows automatic updates and data retrieval from 
       design spreadsheets
===============================================================================
"""

# ========================== Imports & Dependencies ========================== #
from openai import OpenAI
import os
api_key = os.getenv("OPENAI_API_KEY")
print("OPENAI_API_KEY visible ?", "YES" if api_key else "NO")

if not api_key:
    raise RuntimeError(
        "OPENAI_API_KEY is not set in this environment. "
        "Set it in Windows (setx OPENAI_API_KEY \"sk-...\") then restart VS Code."
    )

# --------------------------- Test Sections ---------------------------------- #
client = OpenAI(api_key=api_key)

response = client.responses.create(
    model="gpt-4.1-mini",
    input="Write a one-sentence bedtime story about a unicorn."
)

print(response.output_text)

etwetweq