import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from PIL import Image, ImageTk
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os
import re
import shutil
import json
from pathlib import Path

# keyring stores the password in Windows Credential Manager (encrypted by Windows).
# Install it once with:  pip install keyring
try:
    import keyring
    KEYRING_OK = True
except ImportError:
    KEYRING_OK = False

# ---------------- SETTINGS YOU CAN TUNE ----------------
REPORT_WAIT = 15          # seconds to let the report generate after clicking Execute
DOWNLOAD_TIMEOUT = 120    # max seconds to wait for one PDF to finish downloading
MAX_RETRIES = 2           # how many times to retry a row if it fails
# --------------------------------------------------------

FINAL_FOLDER = Path.home() / "Downloads"          # where renamed PDFs end up
TEMP_FOLDER = FINAL_FOLDER / "_wialon_tmp"        # Chrome downloads here first


CONFIG_FILE = Path.home() / ".wialon_scraper_config.json"   # stores username only
KEYRING_SERVICE = "AutoWebScraperPro_Wialon"                # password goes to keyring


def load_saved_credentials():
    """Return (username, password) if saved, else ('', '')."""
    username, password = "", ""
    try:
        if CONFIG_FILE.exists():
            data = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
            username = data.get("username", "")
        if username and KEYRING_OK:
            password = keyring.get_password(KEYRING_SERVICE, username) or ""
    except Exception as e:
        print(f"Could not load saved login: {e}")
    return username, password


def save_credentials(username, password):
    try:
        CONFIG_FILE.write_text(json.dumps({"username": username}), encoding="utf-8")
        if KEYRING_OK:
            keyring.set_password(KEYRING_SERVICE, username, password)
    except Exception as e:
        print(f"Could not save login: {e}")


def clear_credentials():
    try:
        if CONFIG_FILE.exists():
            data = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
            old_user = data.get("username", "")
            if old_user and KEYRING_OK:
                try:
                    keyring.delete_password(KEYRING_SERVICE, old_user)
                except Exception:
                    pass
            CONFIG_FILE.unlink()
    except Exception as e:
        print(f"Could not clear saved login: {e}")


def clean_value(value):
    """Return a clean string, or '' for blank/NaN cells."""
    if value is None or pd.isnull(value):
        return ""
    return str(value).strip()


def safe_filename(name):
    """Remove characters Windows doesn't allow in file names."""
    return re.sub(r'[\\/:*?"<>|]+', "_", name).strip()


def clear_temp_folder():
    TEMP_FOLDER.mkdir(parents=True, exist_ok=True)
    for f in TEMP_FOLDER.iterdir():
        try:
            f.unlink()
        except Exception:
            pass


def wait_for_download(timeout=DOWNLOAD_TIMEOUT):
    """
    Wait until exactly one finished PDF exists in the temp folder
    (no .crdownload / .tmp files, and file size has stopped changing).
    Returns the file path.
    """
    end_time = time.time() + timeout
    last_size = -1
    stable_count = 0

    while time.time() < end_time:
        files = list(TEMP_FOLDER.iterdir())
        in_progress = [f for f in files if f.suffix.lower() in (".crdownload", ".tmp")]
        pdfs = [f for f in files if f.suffix.lower() == ".pdf"]

        if pdfs and not in_progress:
            pdf = max(pdfs, key=lambda p: p.stat().st_mtime)
            size = pdf.stat().st_size
            if size > 0 and size == last_size:
                stable_count += 1
                if stable_count >= 2:      # size unchanged for ~2 checks
                    return pdf
            else:
                stable_count = 0
            last_size = size
        else:
            stable_count = 0
            last_size = -1

        time.sleep(1)

    raise TimeoutError("PDF download did not finish in time")


def move_with_retry(src, dst, attempts=10):
    """Move/rename the PDF, retrying if Chrome or antivirus still has it locked."""
    last_error = None
    for _ in range(attempts):
        try:
            os.replace(src, dst)   # overwrites existing file safely
            return
        except PermissionError as e:
            last_error = e
            time.sleep(1)
        except OSError:
            # os.replace can fail across drives; fall back to shutil
            try:
                shutil.move(str(src), str(dst))
                return
            except Exception as e:
                last_error = e
                time.sleep(1)
    raise last_error


def type_into(wait, element_id, text, pause=0.5):
    box = wait.until(EC.element_to_be_clickable((By.ID, element_id)))
    box.click()
    box.send_keys(Keys.CONTROL + "a")
    box.send_keys(Keys.BACKSPACE)
    time.sleep(pause)
    box.send_keys(text)
    time.sleep(pause)
    box.send_keys(Keys.ENTER)


def run_one_row(driver, wait, vehicle_number, from_date, to_date, job_number):
    # 1. Vehicle
    type_into(wait, "report_templates_filter_units", vehicle_number, pause=1)

    # 2. From / To dates
    if from_date:
        type_into(wait, "time_from_report_templates_filter_time", from_date)
    if to_date:
        type_into(wait, "time_to_report_templates_filter_time", to_date)

    # 3. Make sure the temp folder is empty so we only catch THIS download
    clear_temp_folder()

    # 4. Execute
    execute_btn = wait.until(EC.element_to_be_clickable((By.ID, "report_templates_filter_params_execute")))
    execute_btn.click()
    time.sleep(REPORT_WAIT)

    # 5. Export -> PDF
    export_btn = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#report_result_export .dropdown-btn")))
    driver.execute_script("arguments[0].click();", export_btn)
    time.sleep(2)

    pdf_btn = wait.until(EC.presence_of_element_located(
        (By.XPATH, "//span[contains(@class, 'option-label') and text()='PDF']")))
    driver.execute_script("arguments[0].click();", pdf_btn)

    # 6. Wait until the file is REALLY finished, then rename
    downloaded = wait_for_download()
    final_path = FINAL_FOLDER / f"{safe_filename(job_number)}.pdf"
    move_with_retry(downloaded, final_path)
    return final_path.name


def process_data(file_path, username, password, template_name, time_frame):
    driver = None
    try:
        # Read Job/Vehicle as text so numbers like 1001 don't become 1001.0
        df = pd.read_excel(file_path, dtype={"Job Number": str, "Vehicle Number": str})

        TEMP_FOLDER.mkdir(parents=True, exist_ok=True)

        options = webdriver.ChromeOptions()
        prefs = {
            "download.default_directory": str(TEMP_FOLDER),
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "plugins.always_open_pdf_externally": True,   # download PDF instead of opening it
            "profile.default_content_setting_values.automatic_downloads": 1,  # allow multiple downloads
        }
        options.add_experimental_option("prefs", prefs)

        driver = webdriver.Chrome(options=options)
        wait = WebDriverWait(driver, 20)

        # --- LOGIN ---
        driver.get("https://hosting.wialon.com/")

        user_box = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[data-testid='LoginMainEmailInput']")))
        user_box.send_keys(username)

        pass_box = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[data-testid='LoginMainPassInput']")))
        pass_box.send_keys(password)

        login_btn = wait.until(EC.element_to_be_clickable((By.ID, "monitoringLoginMainSubmitButton")))
        login_btn.click()

        WebDriverWait(driver, 45).until(EC.presence_of_element_located((By.ID, "report_templates_filter_units")))
        print("Login successful! Letting the dashboard fully load...")
        time.sleep(15)

        # --- TEMPLATE ---
        template_box = wait.until(EC.element_to_be_clickable((By.ID, "report_templates_filter_reports")))
        template_box.click()
        template_box.send_keys(Keys.CONTROL + "a")
        template_box.send_keys(Keys.BACKSPACE)
        time.sleep(1)
        template_box.send_keys(template_name)
        time.sleep(1)
        template_box.send_keys(Keys.ENTER)
        time.sleep(2)

        # --- TIME FRAME ---
        time_id_map = {
            "Today": "report_templates_filter_time-today-label",
            "Yesterday": "report_templates_filter_time-yesterday-label",
            "Week": "report_templates_filter_time-week-label",
            "Month": "report_templates_filter_time-month-label",
        }
        time_btn_id = time_id_map.get(time_frame, "report_templates_filter_time-today-label")
        time_btn = wait.until(EC.element_to_be_clickable((By.ID, time_btn_id)))
        time_btn.click()

        print("Settings applied! Starting the loop in 5 seconds...")
        time.sleep(5)

        # --- EXCEL LOOP ---
        results = []

        for index, row in df.iterrows():
            job_number = clean_value(row.get("Job Number", ""))
            vehicle_number = clean_value(row.get("Vehicle Number", ""))

            if not job_number or not vehicle_number:
                results.append("Skipped: Missing data")
                continue

            # Automatic date conversion: '10/21/2026 0:00' -> '2026 October 21 00:00'
            raw_from = row.get("From", "")
            raw_to = row.get("To", "")
            try:
                from_date = pd.to_datetime(raw_from).strftime("%Y %B %d %H:%M") if pd.notnull(raw_from) else ""
                to_date = pd.to_datetime(raw_to).strftime("%Y %B %d %H:%M") if pd.notnull(raw_to) else ""
            except Exception:
                from_date = clean_value(raw_from)
                to_date = clean_value(raw_to)

            status = None
            for attempt in range(1, MAX_RETRIES + 1):
                try:
                    name = run_one_row(driver, wait, vehicle_number, from_date, to_date, job_number)
                    status = f"Success: {name}"
                    print(status)
                    break
                except Exception as loop_error:
                    status = f"Failed on {job_number}: {str(loop_error)}"
                    print(f"Attempt {attempt} failed for {job_number}: {loop_error}")
                    time.sleep(3)

            results.append(status)

        # Make sure nothing is still downloading before closing the browser
        print("All rows processed! Checking no downloads are still running...")
        end_time = time.time() + 30
        while time.time() < end_time and any(
            f.suffix.lower() in (".crdownload", ".tmp") for f in TEMP_FOLDER.iterdir()
        ):
            time.sleep(1)
        time.sleep(2)

        driver.quit()
        driver = None

        # Clean up temp folder if empty
        try:
            TEMP_FOLDER.rmdir()
        except OSError:
            pass

        df["Status"] = results
        output_filename = os.path.join(os.path.dirname(file_path), "Scraped_Results.xlsx")
        df.to_excel(output_filename, index=False)
        messagebox.showinfo("Success", f"Process complete! Status saved to {output_filename}")

    except Exception as e:
        messagebox.showerror("Error", f"A fatal error occurred: {str(e)}")
    finally:
        if driver is not None:
            try:
                driver.quit()
            except Exception:
                pass


def start_process():
    username = user_entry.get()
    password = pass_entry.get()
    template_name = template_combo.get()
    time_frame = time_combo.get()

    if not username or not password:
        messagebox.showwarning("Warning", "Please enter both Username and Password.")
        return

    if remember_var.get():
        if not KEYRING_OK:
            messagebox.showinfo(
                "Password not saved",
                "Only the username can be remembered because the 'keyring' package "
                "is not installed.\n\nInstall it with:  pip install keyring"
            )
        save_credentials(username, password)
    else:
        clear_credentials()

    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if file_path:
        root.update()
        process_data(file_path, username, password, template_name, time_frame)


root = tk.Tk()
root.title("Auto Web Scraper Pro")
root.geometry("450x520")
root.configure(bg="#f3f4f6")

style = ttk.Style()
style.theme_use("clam")

style.configure("TLabel", background="#f3f4f6", font=("Segoe UI", 10), foreground="#374151")
style.configure("TButton", font=("Segoe UI", 10, "bold"), padding=8, background="#3b82f6", foreground="white")
style.map("TButton", background=[("active", "#2563eb")])
style.configure("TCombobox", padding=5, font=("Segoe UI", 10))
style.configure("TEntry", padding=5, font=("Segoe UI", 10))

# --- Logo ---
try:
    logo_image = Image.open("logo.png").convert("RGBA")
    datas = logo_image.getdata()
    newData = []
    for item in datas:
        if item[0] < 30 and item[1] < 30 and item[2] < 30:
            newData.append((255, 255, 255, 0))
        else:
            newData.append(item)
    logo_image.putdata(newData)
    logo_image = logo_image.resize((160, 48), Image.Resampling.LANCZOS)
    logo_photo = ImageTk.PhotoImage(logo_image)

    logo_label = tk.Label(root, image=logo_photo, bg="#f3f4f6", borderwidth=0)
    logo_label.image = logo_photo
    logo_label.pack(pady=(25, 15))
except Exception as e:
    print(f"Error loading logo: {str(e)}")

frame = tk.Frame(root, bg="#f3f4f6")
frame.pack(padx=40, fill="both", expand=True)

ttk.Label(frame, text="Wialon Username").pack(anchor="w", pady=(0, 4))
user_entry = ttk.Entry(frame)
user_entry.pack(fill="x", pady=(0, 15))

ttk.Label(frame, text="Wialon Password").pack(anchor="w", pady=(0, 4))
pass_entry = ttk.Entry(frame, show="•")
pass_entry.pack(fill="x", pady=(0, 8))

# Remember me checkbox
style.configure("TCheckbutton", background="#f3f4f6", font=("Segoe UI", 9), foreground="#374151")
remember_var = tk.BooleanVar(value=False)
remember_check = ttk.Checkbutton(frame, text="Remember my login", variable=remember_var)
remember_check.pack(anchor="w", pady=(0, 15))

# Auto-fill saved login (if any)
_saved_user, _saved_pass = load_saved_credentials()
if _saved_user:
    user_entry.insert(0, _saved_user)
    pass_entry.insert(0, _saved_pass)
    remember_var.set(True)

ttk.Label(frame, text="Select Report Template").pack(anchor="w", pady=(0, 4))
template_combo = ttk.Combobox(frame, state="readonly")
template_combo["values"] = [
    "OkiDoki MC - Customer Location IN and OUT- Unit Time/ Mileage",
    "External Power Report",
    "Accident Speed Monitoring"
]
template_combo.current(0)
template_combo.pack(fill="x", pady=(0, 15))

ttk.Label(frame, text="Select Time Frame").pack(anchor="w", pady=(0, 4))
time_combo = ttk.Combobox(frame, state="readonly")
time_combo["values"] = ["Today", "Yesterday", "Week", "Month"]
time_combo.current(0)
time_combo.pack(fill="x", pady=(0, 25))

upload_btn = ttk.Button(frame, text="Upload Excel & Run", command=start_process)
upload_btn.pack(fill="x", pady=(10, 20))

root.mainloop()