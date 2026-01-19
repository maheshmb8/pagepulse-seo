"""
PagePulse – SEO Metrics Automation Tool

NOTE:
This repository demonstrates engineering patterns for automating
SEO data retrieval and reporting.

External services (Google Search Console), credentials, and
authentication flows are intentionally excluded and must be
provided externally.
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import DateEntry
from datetime import datetime
import pandas as pd
import os
import getpass
import threading
import time
from google.oauth2 import service_account
from googleapiclient.discovery import build
from oauth2client.service_account import ServiceAccountCredentials
import gspread

# Cell 2 – Paths & Constants
def resource_path(relative_path):
    """ Get absolute path to resource, works in dev & PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

ICON_PATH = resource_path('url_icon.ico')

current_version = 4
local_version = 4

if local_version != current_version:
    messagebox.showerror(
        "Version Mismatch",
        f"This app version ({local_version}) is outdated.\n"
        f"Please download the latest version ({current_version})."
    )
    raise SystemExit

APP_PASSWORD = os.getenv("PAGEPULSE_APP_PASSWORD")
SITE_URL = os.getenv("TARGET_SITE_URL", "https://example.com/")


auth_credentials = 'google_search_console_api_key'

# Credentials are expected to be provided externally
# search_console_service = build('searchconsole', 'v1', credentials=auth_credentials)

input_file_path = None
output_folder_path = os.path.join(os.path.expanduser("~"), "Downloads")

# Cell 4 – File selection function
def browse_input_file():
    global input_file_path
    path = filedialog.askopenfilename(
        filetypes=[("Excel/CSV files", "*.xlsx *.xls *.csv")],
        title="Select File with Page URLs"
    )
    if path:
        input_file_path = path
        try:
            input_label.config(text=f"Selected file:\n{path}")
        except Exception:
            pass

# Cell 5 – Data fetching function
def fetch_data_for_range(start_date, end_date, progress_bar=None, progress_label=None, time_label=None):
    try:
        if not input_file_path:
            messagebox.showerror("Missing Info", "Please select an input file.")
            return

        if input_file_path.endswith(".csv"):
            df_inputs = pd.read_csv(input_file_path, usecols=[0])
        else:
            df_inputs = pd.read_excel(input_file_path, usecols=[0])

        df_inputs.columns = ["input_value"]
        inputs = df_inputs["input_value"].dropna().tolist()
        total_inputs = len(inputs)

        if not inputs:
            messagebox.showerror("No Inputs", "The selected file contains no URLs or queries.")
            return

        timestamp = datetime.now().strftime("_%Y%m%d_%H%M%S")
        output_file = os.path.join(
            output_folder_path,
            f"search_console_data_{start_date}_to_{end_date}{timestamp}.xlsx"
        )

        all_rows = []
        nth = 0
        start_time = time.time()

        for idx, value in enumerate(inputs, start=1):
            value = str(value).strip()

            if value.lower().startswith("http"):
                filters = [{"dimension": "page", "operator": "equals", "expression": value}]
                dimensions = ["page"]
                detected_type = "URL"
            else:
                filters = [{"dimension": "query", "operator": "contains", "expression": value}]
                dimensions = ["query", "page"]
                detected_type = "Query"

            request = {
                "startDate": start_date,
                "endDate": end_date,
                "dimensions": dimensions,
                "dimensionFilterGroups": [{"filters": filters}],
                "rowLimit": 25000
            }

            nth += 1
            if nth > 900:
                time.sleep(65)
                nth = 0

            response = search_console_service.searchanalytics().query(
                siteUrl=SITE_URL, body=request).execute()
            rows = response.get("rows", [])

            for row in rows:
                if detected_type == "URL":
                    all_rows.append({
                        "Input": value, "DetectedType": detected_type, "Query": "",
                        "Page": row["keys"][0], "Clicks": row.get("clicks", 0),
                        "Impressions": row.get("impressions", 0),
                        "CTR": row.get("ctr", 0), "Position": row.get("position", 0)
                    })
                else:
                    all_rows.append({
                        "Input": value, "DetectedType": detected_type,
                        "Query": row["keys"][0], "Page": row["keys"][1],
                        "Clicks": row.get("clicks", 0), "Impressions": row.get("impressions", 0),
                        "CTR": row.get("ctr", 0), "Position": row.get("position", 0)
                    })

            # Progress updates
            if progress_bar and progress_label and time_label:
                percent = int((idx / total_inputs) * 100)
                elapsed = int(time.time() - start_time)
                avg_time = elapsed / idx
                eta = int(avg_time * (total_inputs - idx))

                progress_bar["value"] = percent
                progress_label.config(text=f"Processed {idx}/{total_inputs} ({percent}%)")
                time_label.config(text=f"Elapsed: {elapsed}s | ETA: {eta}s")
                progress_bar.update()
                progress_label.update()
                time_label.update()

        date_range = f"{start_date} to {end_date}"

        if all_rows:
            df = pd.DataFrame(all_rows)
            df["CTR"] = (df["CTR"] * 100).round(2).astype(str) + "%"
            df["date_range"] = date_range
            df = df[["Input", "DetectedType", "Query", "Page",
                     "Clicks", "Impressions", "CTR", "Position", "date_range"]]

            df.to_excel(output_file, index=False)
            messagebox.showinfo("Success", f"Data exported to:\n{output_file}")
        else:
            messagebox.showinfo("Done", "No data found for selected date range.")

    except Exception as e:
        messagebox.showerror("Error", f"Something went wrong:\n{str(e)}")

# Cell 6 – GUI Windows
def show_main_window():
    try:
        login_win.destroy()
    except Exception:
        pass

    root = tk.Tk()
    root.title("PagePulse")
    root.geometry("760x600")
    root.configure(bg="#eef2f8")

    try:
        root.iconbitmap(ICON_PATH)
    except Exception:
        pass

    HEADER_GRADIENT_START = "#635BFF"
    HEADER_GRADIENT_END = "#6FA3FF"
    CARD_BG = "white"
    PRIMARY_GRAD = ("#36b37e", "#2b8a5e")
    BUTTON_HOVER = "#2aa06d"
    TEXT_COLOR = "#263238"

    style = ttk.Style()
    style.theme_use("clam")
    style.configure("TLabel", background="#eef2f8", foreground=TEXT_COLOR)
    style.configure("Card.TFrame", background=CARD_BG)
    style.configure("Small.TLabel", font=("Segoe UI", 9), background=CARD_BG, foreground="#555")
    style.configure("Title.TLabel", font=("Segoe UI", 18, "bold"), background=CARD_BG, foreground="#111")
    style.configure("Desc.TLabel", font=("Segoe UI", 10), background=CARD_BG, foreground="#444")

    header = tk.Canvas(root, height=100, width=760, bd=0, highlightthickness=0)
    header.pack(fill="x", padx=0, pady=(0, 10))

    width = 760
    for i in range(width):
        r1, g1, b1 = root.winfo_rgb(HEADER_GRADIENT_START)
        r2, g2, b2 = root.winfo_rgb(HEADER_GRADIENT_END)
        r = int(r1 + (r2 - r1) * i / width)
        g = int(g1 + (g2 - g1) * i / width)
        b = int(b1 + (b2 - b1) * i / width)
        color = f"#{r>>8:02x}{g>>8:02x}{b>>8:02x}"
        header.create_line(i, 0, i, 100, fill=color)

    header.create_text(40, 50, anchor="w", text="🔗", font=("Segoe UI", 28), fill="white")
    header.create_text(110, 50, anchor="w", text="PagePulse", font=("Segoe UI", 26, "bold"), fill="white")

    card_outer = tk.Frame(root, bg="#eef2f8")
    card_outer.pack(fill="both", expand=True, padx=30)

    card = tk.Frame(card_outer, bg=CARD_BG, bd=0, relief="flat")
    card.pack(fill="both", expand=True, padx=0, pady=0)
    card.config(highlightbackground="#e1e7f0", highlightthickness=1)

    desc = ttk.Label(card, text="Export Google Search Console data for chosen URLs and dates, saved to your Downloads folder.",
                     style="Desc.TLabel", wraplength=660, justify="left")
    desc.pack(padx=30, pady=(20, 10), anchor="w")

    date_frame = tk.Frame(card, bg=CARD_BG)
    date_frame.pack(fill="x", padx=30, pady=(5, 10))

    start_label = ttk.Label(date_frame, text="Start Date", style="Small.TLabel")
    start_label.grid(row=0, column=0, sticky="w", padx=(0,5), pady=(4,4))
    start_cal = DateEntry(date_frame, width=19, background='darkblue', foreground='white', date_pattern='yyyy-mm-dd')
    start_cal.grid(row=0, column=1, sticky="w", padx=(5,20), pady=(4,4))

    end_label = ttk.Label(date_frame, text="End Date", style="Small.TLabel")
    end_label.grid(row=1, column=0, sticky="w", padx=(0,5), pady=(4,4))
    end_cal = DateEntry(date_frame, width=19, background='darkblue', foreground='white', date_pattern='yyyy-mm-dd')
    end_cal.grid(row=1, column=1, sticky="w", padx=(5,20), pady=(4,4))

    file_frame = tk.Frame(card, bg=CARD_BG)
    file_frame.pack(fill="x", padx=30, pady=(10, 6))

    def browse_btn_style_enter(e):
        browse_btn.config(bg="#e8e7ff")
    def browse_btn_style_leave(e):
        browse_btn.config(bg="#f3f3ff")

    browse_btn = tk.Button(file_frame, text="📂  Browse Input File", command=browse_input_file,
                           bg="#f3f3ff", fg="#222", bd=0, relief="flat", padx=12, pady=8,
                           font=("Segoe UI", 10, "bold"))
    browse_btn.pack(side="left")
    browse_btn.bind("<Enter>", browse_btn_style_enter)
    browse_btn.bind("<Leave>", browse_btn_style_leave)

    global input_label
    input_label = ttk.Label(file_frame, text="No input file selected", style="Small.TLabel", wraplength=520, justify="left")
    input_label.pack(side="left", padx=12)

    out_info = ttk.Label(card, text=f"Output will be saved to your Downloads folder:\n{output_folder_path}",
                         style="Small.TLabel", wraplength=660, justify="left")
    out_info.pack(padx=30, pady=(10, 8), anchor="w")

    # --- Progress Bar + Labels ---
    progress_bar = ttk.Progressbar(card, length=600, mode="determinate")
    progress_bar.pack(pady=(5, 2))
    progress_label = ttk.Label(card, text="", style="Small.TLabel")
    progress_label.pack()
    time_label = ttk.Label(card, text="", style="Small.TLabel")
    time_label.pack()

    run_frame = tk.Frame(card, bg=CARD_BG)
    run_frame.pack(fill="x", padx=30, pady=(10, 20))

    def run_task(start_date, end_date):
        try:
            fetch_data_for_range(start_date, end_date, progress_bar, progress_label, time_label)
        finally:
            run_btn.config(state="normal", text="▶ Run Report")

    def on_run():
        start_date = start_cal.get_date().strftime("%Y-%m-%d")
        end_date = end_cal.get_date().strftime("%Y-%m-%d")
        if start_date > end_date:
            messagebox.showerror("Invalid Dates", "Start date cannot be after end date.")
            return
        run_btn.config(state="disabled", text="⏳ Running...")
        threading.Thread(target=run_task, args=(start_date, end_date), daemon=True).start()

    def run_enter(e):
        run_btn.config(bg=BUTTON_HOVER)
    def run_leave(e):
        run_btn.config(bg=PRIMARY_GRAD[0])

    run_btn = tk.Button(run_frame, text="▶ Run Report", bg=PRIMARY_GRAD[0], fg="white",
                        font=("Segoe UI", 12, "bold"), bd=0, relief="flat", padx=20, pady=10, command=on_run)
    run_btn.pack(pady=5)
    run_btn.bind("<Enter>", run_enter)
    run_btn.bind("<Leave>", run_leave)

    footer = ttk.Label(card, text="Created by (Team Name)    Support: (Team_email@domain.com) | (alternate_email@domain.com)",
                       style="Small.TLabel", wraplength=660, justify="center")
    footer.pack(side="bottom", pady=(10, 16))

    root.mainloop()

# Cell 7 – Login Window
login_win = tk.Tk()
login_win.title("PagePulse Login")
login_win.geometry("360x170")
login_win.configure(bg="#eef2f8")

try:
    login_win.iconbitmap(ICON_PATH)
except Exception:
    pass

ttk.Label(login_win, text="Enter Password:", background="#eef2f8", font=("Segoe UI", 10)).pack(pady=(18,6))
password_entry = ttk.Entry(login_win, show="*", width=26)
password_entry.pack(pady=(0,10))

def attempt_login():
    if password_entry.get() == APP_PASSWORD:
        show_main_window()
    else:
        messagebox.showerror("Access Denied", "Incorrect password!")

login_btn = tk.Button(login_win, text="Login", bg="#36b37e", fg="white",
                      font=("Segoe UI", 10, "bold"), relief="flat", command=attempt_login, padx=10, pady=6)
login_btn.pack(pady=(2,10))

login_win.mainloop()
