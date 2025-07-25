import time, re, tkinter as tk
from tkinter import messagebox, filedialog
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import openpyxl
import pandas as pd
from datetime import datetime
import os
from collections import Counter
import ttkbootstrap as ttkb
from ttkbootstrap.constants import *

# ──────── Helpers ────────
def extract_hour(dt):
    m = re.search(r'\b(\d{1,2}:\d{2}\s*[AP]M)\b', dt)
    return m.group(1) if m else "Time not found"

def parse_datetime(dt):
    dt = re.split(r'\s*\(', dt)[0].strip()
    try:
        return datetime.strptime(dt, "%d/%m/%Y %I:%M %p")
    except ValueError:
        return None

def calc_duration(start, end):
    d1, d2 = parse_datetime(start), parse_datetime(end)
    if not d1 or not d2:
        return "n/a"
    sec = int(abs((d2 - d1).total_seconds()))
    mins = sec // 60 
    hrs, mins = divmod(mins, 60)
    return f"{hrs} h {mins} min" if mins else f"{hrs} hours"

def determine_task_type_and_tool(description):
    desc = str(description).lower()
    if "please extract bilingual file" in desc and "cloud.memsource.com" in desc:
        return "File extraction", "Memsource"
    if "could you please run qa using the tb and convert the file back" in desc:
        return "Convert back", "Trados"
    if "could you please convert the file to sdlppx and lock pm and cm segments" in desc:
        return "Preparation", "Trados"
    if "could you please provide the source and the target files" in desc:
        return "File extraction", "Automation tool by LE"
    if "could you please provide the log analysis" in desc:
        return "Analysis Report", "Automation tool by LE"
    if "send me" in desc and "https://logs08.sdlproducts.com" in desc:
        return "File preparation", "Trados"
    if any(x in desc for x in ["change name", "change username", "change tcs", "change author"]):
        return "Change author name from comment", "Automation tool by LE"
    if "extract files" in desc:
        if "cloud.memsource.com" in desc:
            return "File extraction", "Memsource"
        elif "memoq" in desc:
            return "File extraction", "MemoQ"
        return "File extraction", ""
    if "target file" in desc and "memoqxliff" in desc:
        return "File extraction", "MemoQ"
    if "deliver groupshare" in desc:
        return "Finalization", "Trados"
    if "update tm" in desc:
        return "Update TM", "Trados"
    if any(x in desc for x in ["ocr", "editable file"]):
        return "OCR", "Abby fine reader"
    if any(x in desc for x in ["ltb", "compare report", "comparison report"]):
        return "LTB change-comparison reports", "LTB"
    if any(x in desc for x in ["rp", "RP", "return package"]):
        return "Convert back", "Trados"
    if any(x in desc for x in ["x-bench", "x bench", "xbench", "qa"]):
        return "QA report", "Xbench"
    if "verifika" in desc or "verifica" in desc:
        return "QA report", "Verifika"
    if "convert back" in desc and "qa" in desc:
        return "Convert back", "MemoQ"
    if "convert back" in desc:
        return "Convert back", "MemoQ"
    if "trados" in desc:
        return ("Convert back", "Trados") if "convert back" in desc else ("File preparation", "Trados")
    if "apply tm" in desc:
        return ("Apply TM", "TWS") if "token" in desc else ("Apply TM", "Trados")
    if "log analysis" in desc or ("log" in desc and all(x not in desc for x in ["rtf", "trados", "apply tm"])):
        return "Analysis Report", "Trados"
    if "deliver" in desc:
        return "Finalization", "MemoQ"
    if "rtf" in desc:
        return "File preparation", "MemoQ"
    if "mt" in desc:
        return "Machine translation", "Memsource"
    return "", ""

def day_name_from_cell(cell):
    if pd.isna(cell):
        return ""
    if isinstance(cell, (pd.Timestamp, datetime)):
        return pd.to_datetime(cell).day_name()
    date_part = str(cell).split()[0]
    try:
        dt = pd.to_datetime(date_part, format="%d/%m/%Y")
        return dt.day_name()
    except ValueError:
        return ""

def brand_from_url(url):
    try:
        return url.split("//")[1].split(".")[0]
    except:
        return "Unknown"

# ──────── Global Driver ────────
driver = None

def open_browser():
    global driver
    try:
        driver = webdriver.Chrome()
        driver.maximize_window()
        wait = WebDriverWait(driver, 15)
        urls = [
            "https://bayantech.rasberryapp.com/LERequest/LERequestListIndex",
            #"https://sawatech.rasberryapp.com/LERequest/LERequestListIndex",
            "https://asialocalize.rasberryapp.com/LERequest/LERequestListIndex",
            "https://laoret.rasberryapp.com/LERequest/LERequestListIndex",
            "https://transpalm.rasberryapp.com/LERequest/LERequestListIndex"
        ]
        email = "hussam-sherif@teqneyat.com"
        password = "Huss@m2024"
        for u in urls:
            driver.get(u)
            wait.until(EC.presence_of_element_located((By.ID, "UserName"))).send_keys(email)
            driver.find_element(By.ID, "Password").send_keys(password)
            driver.find_element(By.ID, "btn_login").click()
            time.sleep(2)
        messagebox.showinfo("Browser ready", "Login completed. Now choose the .txt file of URLs to process.")
        process_button.config(state=tk.NORMAL)
    except Exception as exc:
        messagebox.showerror("Error opening browser", str(exc))

def process_urls():
    if not driver:
        messagebox.showerror("Error", "Browser not opened.")
        return
    path = filedialog.askopenfilename(title="Select URL list (TXT)", filetypes=[("Text files", "*.txt")])
    if not path:
        return
    with open(path, encoding="utf-8") as f:
        urls = [line.strip() for line in f if line.strip()]
    if not urls:
        messagebox.showwarning("No URLs", "The file is empty.")
        return

    allowed_users = {
        "hussam-sherif@teqneyat.com",
        "sara-hassan@teqneyat.com",
        "abdullah.nasr@teqneyat.com",
        "hana-tarek@teqneyat.com",
        "abdelrahman.bakr@teqneyat.com"
    }
    LEs = {"hussam", "sara hassan", "hana tarek","Abdullah Nasr","Abdulrahman Bakr"}
    rows_out = []
    wait = WebDriverWait(driver, 15)

    for url in urls:
        try:
            item_id_match = re.search(r"itemId=(\d+)", url)
            item_id = item_id_match.group(1) if item_id_match else None
            if not item_id:
                raise Exception("Could not extract itemId from URL.")

            driver.get(url)
            driver.execute_script("document.body.style.zoom='75%'")
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "Detailstab"))).click()
            brand = brand_from_url(url)
            date_txt = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="tabcontent"]/div[1]/div/div/div[2]/p'))).text.strip()
            description = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//div[@role="textbox" and @aria-label="Rich Text Editor, main"]'))).text.strip()

            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "Historytab"))).click()
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, '#LERequestHistoryListTabulator .tabulator-row')))

            rows = driver.find_elements(By.CSS_SELECTOR, '#LERequestHistoryListTabulator .tabulator-row')
            allowed_stamps, other_stamps = [], []
            allowed_users_lower = {u.lower() for u in allowed_users}

            for r in rows:
                try:
                    act = r.find_element(By.CSS_SELECTOR, 'div[tabulator-field="ActivityName"]').text.strip()
                    user = r.find_element(By.CSS_SELECTOR, 'div[tabulator-field="CreatedBy"]').text.strip().lower()
                    if act == "Changed Status":
                        ts = r.find_element(By.CSS_SELECTOR, 'div[tabulator-field="CreatedDate"]').text.strip()
                        (allowed_stamps if user in allowed_users_lower else other_stamps).append(ts)
                except:
                    continue

            allowed_stamps = sorted(allowed_stamps, key=parse_datetime)
            other_stamps = sorted(other_stamps, key=parse_datetime)

            if allowed_stamps:
                start_time, end_time = allowed_stamps[0], allowed_stamps[-1]
            elif other_stamps:
                start_time, end_time = other_stamps[0], other_stamps[-1]
            else:
                start_time = end_time = None

            if start_time and end_time:
                start_hr, end_hr = extract_hour(start_time), extract_hour(end_time)
                duration = calc_duration(start_time, end_time)
            else:
                start_hr = end_hr = duration = "Not found"

            time.sleep(5)
            le_request_text = f"LE Request #{item_id}"
            try:
                le_request_link = wait.until(EC.element_to_be_clickable((By.XPATH, f'//a[normalize-space(text())="{le_request_text}"]')))
                le_request_link.click()
            except: continue

            time.sleep(3)
            wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div[1]/div/div[4]/div[2]/div[3]/a[2]'))).click()
            wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="fileManagerViewList"]/div[2]')))
            wait.until(EC.element_to_be_clickable((By.XPATH, f'//*[@id="fileManager-tbl-content"]/div[2]/div//a[@class="btn btn-link tree-link" and normalize-space(text())="{le_request_text}"]'))).click()

            time.sleep(3)
            created_by_cells = driver.find_elements(By.XPATH, '//*[@id="fileManager-tbl-content"]/div[2]/div//div[@tabulator-field="CreatedByName"]')
            created_by_names = [cell.text.strip() for cell in created_by_cells if cell.text.strip()]
            name_counts = Counter(created_by_names)

            non_allowed_count = sum(count for name, count in name_counts.items() if name.lower() not in LEs)
            allowed_count = sum(count for name, count in name_counts.items() if name.lower() in LEs)
            amount = non_allowed_count if non_allowed_count > 0 else allowed_count
            task_type, tool = determine_task_type_and_tool(description)
            weekday = day_name_from_cell(date_txt)

            rows_out.append([weekday, date_txt, start_hr, end_hr, duration, item_id, task_type, amount, "File", "Easy", tool, brand])

        except Exception as exc:
            messagebox.showerror("Extraction error", f"{url}\n{str(exc)}")

    try:
        output_path = os.path.join(os.path.dirname(path), "project_data.xlsx")
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Project Data"
        ws.append(["Day", "Date", "Start Time", "End Time", "Total hours", "Task ID", "Task Type", "Amount", "Unit", "Complexity Level", "Tool", "Brand"])
        for row in rows_out:
            ws.append(row)
        wb.save(output_path)
        messagebox.showinfo("Done", f"Saved to: {output_path}")
    except Exception as exc:
        messagebox.showerror("Excel error", str(exc))

def on_close():
    if driver:
        driver.quit()
    root.destroy()

# ──────── GUI ────────
root = ttkb.Window(title="\U0001F9E0 Rasberry Project Extractor", themename="cyborg", size=(600, 450))
root.resizable(False, False)

frame = ttkb.Frame(root, padding=20)
frame.pack(fill=tk.BOTH, expand=True)

ttkb.Label(frame, text="\U0001F512 Step 1: Log in to Rasberry Portals", font=("Segoe UI", 12, "bold")).pack(pady=(10, 6))
ttkb.Button(frame, text="\U0001F310 Start Browser & Login", bootstyle=SUCCESS, width=40, command=open_browser).pack(pady=4)

ttkb.Label(frame, text="\U0001F4C4 Step 2: Load URLs & Extract Data", font=("Segoe UI", 12, "bold")).pack(pady=(20, 6))
process_button = ttkb.Button(frame, text="\u2699\ufe0f Process URLs from TXT", bootstyle=INFO, width=40, state=tk.DISABLED, command=process_urls)
process_button.pack(pady=4)

ttkb.Button(frame, text="\u274C Exit", bootstyle=DANGER, width=40, command=on_close).pack(pady=(30, 10))

root.protocol("WM_DELETE_WINDOW", on_close)
root.mainloop()
