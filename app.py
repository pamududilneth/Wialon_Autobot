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
import glob
from pathlib import Path

def process_data(file_path, username, password, template_name, time_frame):
    try:
        df = pd.read_excel(file_path)
        
        options = webdriver.ChromeOptions()
        driver = webdriver.Chrome(options=options)
        wait = WebDriverWait(driver, 15)
        
        # 1. Navigate to the login page
        driver.get("https://hosting.wialon.com/") 
        
        # --- AUTOMATED LOGIN SEQUENCE ---
        
        # A. Wait for the Username box to appear and type the username
        user_box = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[data-testid='LoginMainEmailInput']")))
        user_box.send_keys(username)
        
        # B. Find the Password box and type the password
        pass_box = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[data-testid='LoginMainPassInput']")))
        pass_box.send_keys(password)
        
        # C. Click the Login Button
        login_btn = wait.until(EC.element_to_be_clickable((By.ID, "monitoringLoginMainSubmitButton")))
        login_btn.click()
        
       # Wait until the main dashboard loads before continuing
        WebDriverWait(driver, 45).until(EC.presence_of_element_located((By.ID, "report_templates_filter_units")))
        
        print("Login successful! Taking a 20-second break to let the dashboard fully load...")
        time.sleep(20) 
        
        # --- SET TEMPLATE AND TIME FRAME ---
        # ... (keep your existing template selection code here) ...
        
        print("Settings applied! Starting the loop in 5 seconds...")
        time.sleep(5) # <-- INCREASED TO 5 SECONDS

        # --- SET TEMPLATE AND TIME FRAME ---
        
        # Set Template
        template_box = wait.until(EC.element_to_be_clickable((By.ID, "report_templates_filter_reports")))
        template_box.click()
        template_box.send_keys(Keys.CONTROL + "a")
        template_box.send_keys(Keys.BACKSPACE)
        time.sleep(1)
        template_box.send_keys(template_name)
        time.sleep(1)
        template_box.send_keys(Keys.ENTER)
        
        # Set Time Frame
        time_id_map = {
            "Today": "report_templates_filter_time-today-label",
            "Yesterday": "report_templates_filter_time-yesterday-label",
            "Week": "report_templates_filter_time-week-label",
            "Month": "report_templates_filter_time-month-label"
        }
        time_btn_id = time_id_map.get(time_frame, "report_templates_filter_time-today-label")
        time_btn = wait.until(EC.element_to_be_clickable((By.ID, time_btn_id)))
        time_btn.click()
        
        print("Settings applied! Starting the loop in 3 seconds...")
        time.sleep(3) # <-- NEW PAUSE: Let the template and time filters apply

        # --- START EXCEL LOOP ---
        results = []
        downloads_folder = str(Path.home() / "Downloads")

        for index, row in df.iterrows():
            job_number = str(row.get('Job Number', '')) 
            vehicle_number = str(row.get('Vehicle Number', ''))
            
            # --- AUTOMATIC DATE CONVERSION ---
            raw_from = row.get('From', '')
            raw_to = row.get('To', '')
            
            try:
                # Converts '10/21/2026 0:00' into '2026 October 21 00:00'
                from_date = pd.to_datetime(raw_from).strftime('%Y %B %d %H:%M') if pd.notnull(raw_from) else ''
                to_date = pd.to_datetime(raw_to).strftime('%Y %B %d %H:%M') if pd.notnull(raw_to) else ''
            except Exception:
                # Fallback just in case the cell is blank or weirdly formatted
                from_date = str(raw_from)
                to_date = str(raw_to)
            # ---------------------------------
            
            if not job_number or not vehicle_number:
                results.append("Skipped: Missing data")
                continue

            try:
                # 1. Type Vehicle Number
                object_input = wait.until(EC.element_to_be_clickable((By.ID, "report_templates_filter_units")))
                object_input.click()
                object_input.send_keys(Keys.CONTROL + "a")
                object_input.send_keys(Keys.BACKSPACE)
                time.sleep(1) 
                object_input.send_keys(vehicle_number)
                time.sleep(1) 
                object_input.send_keys(Keys.ENTER)
                
                # 2. Type 'From' Date
                if from_date and from_date.lower() != 'nan':
                    from_input = wait.until(EC.element_to_be_clickable((By.ID, "time_from_report_templates_filter_time")))
                    from_input.click()
                    from_input.send_keys(Keys.CONTROL + "a")
                    from_input.send_keys(Keys.BACKSPACE)
                    time.sleep(0.5)
                    from_input.send_keys(from_date)
                    from_input.send_keys(Keys.ENTER)

                # 3. Type 'To' Date
                if to_date and to_date.lower() != 'nan':
                    to_input = wait.until(EC.element_to_be_clickable((By.ID, "time_to_report_templates_filter_time")))
                    to_input.click()
                    to_input.send_keys(Keys.CONTROL + "a")
                    to_input.send_keys(Keys.BACKSPACE)
                    time.sleep(0.5)
                    to_input.send_keys(to_date)
                    to_input.send_keys(Keys.ENTER)
                
                # 4. Click Execute
                execute_btn = wait.until(EC.element_to_be_clickable((By.ID, "report_templates_filter_params_execute")))
                execute_btn.click()
                
                # Wait for report
                time.sleep(15) 
                
                # Click Export Menu
                export_btn = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#report_result_export .dropdown-btn")))
                driver.execute_script("arguments[0].click();", export_btn)
                time.sleep(2) 
                
                # Click PDF
                pdf_btn = wait.until(EC.presence_of_element_located((By.XPATH, "//span[contains(@class, 'option-label') and text()='PDF']")))
                driver.execute_script("arguments[0].click();", pdf_btn)
                
                # Wait for download
                time.sleep(5) 
                
                # Rename PDF
                list_of_pdfs = glob.glob(os.path.join(downloads_folder, '*.pdf')) 
                
                if list_of_pdfs:
                    latest_file = max(list_of_pdfs, key=os.path.getctime) 
                    new_file_path = os.path.join(downloads_folder, f"{job_number}.pdf")
                    
                    if os.path.exists(new_file_path):
                        os.remove(new_file_path)
                    os.rename(latest_file, new_file_path)
                    results.append(f"Success: {job_number}.pdf")
                else:
                    results.append(f"Error: No PDF downloaded for {job_number}")

            except Exception as loop_error:
                results.append(f"Failed on {job_number}: {str(loop_error)}")
        
        # --- NEW FINAL BUFFER ---
        print("All rows processed! Waiting 5 seconds to ensure the final download is completely saved...")
        time.sleep(5)
        # ------------------------

        driver.quit()

        df['Status'] = results
        output_filename = "Scraped_Results.xlsx"
        df.to_excel(output_filename, index=False)
        messagebox.showinfo("Success", f"Process complete! Status saved to {output_filename}")

    except Exception as e:
        messagebox.showerror("Error", f"A fatal error occurred: {str(e)}")

def start_process():
    username = user_entry.get()
    password = pass_entry.get()
    template_name = template_combo.get()
    time_frame = time_combo.get()
    
    if not username or not password:
        messagebox.showwarning("Warning", "Please enter both Username and Password.")
        return

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
root.configure(bg="#f3f4f6") # Modern light gray background

# Apply a modern clean theme
style = ttk.Style()
style.theme_use("clam") 

# Configure specific widget styles
style.configure("TLabel", background="#f3f4f6", font=("Segoe UI", 10), foreground="#374151")
style.configure("TButton", font=("Segoe UI", 10, "bold"), padding=8, background="#3b82f6", foreground="white")
style.map("TButton", background=[("active", "#2563eb")]) # Darker blue on hover
style.configure("TCombobox", padding=5, font=("Segoe UI", 10))
style.configure("TEntry", padding=5, font=("Segoe UI", 10))

# --- Add the Logo ---
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

# UI Container Frame for padding
frame = tk.Frame(root, bg="#f3f4f6")
frame.pack(padx=40, fill="both", expand=True)

# Username Field
ttk.Label(frame, text="Wialon Username").pack(anchor="w", pady=(0, 4))
user_entry = ttk.Entry(frame)
user_entry.pack(fill="x", pady=(0, 15))

# Password Field
ttk.Label(frame, text="Wialon Password").pack(anchor="w", pady=(0, 4))
pass_entry = ttk.Entry(frame, show="•")
pass_entry.pack(fill="x", pady=(0, 15))

# Template Dropdown
ttk.Label(frame, text="Select Report Template").pack(anchor="w", pady=(0, 4))
template_combo = ttk.Combobox(frame, state="readonly")
template_combo['values'] = [
    "OkiDoki MC - Customer Location IN and OUT- Unit Time/ Mileage",
    "External Power Report",
    "Accident Speed Monitoring"
]
template_combo.current(0)
template_combo.pack(fill="x", pady=(0, 15))

# Time Frame Dropdown
ttk.Label(frame, text="Select Time Frame").pack(anchor="w", pady=(0, 4))
time_combo = ttk.Combobox(frame, state="readonly")
time_combo['values'] = ["Today", "Yesterday", "Week", "Month"]
time_combo.current(0)
time_combo.pack(fill="x", pady=(0, 25))

# Upload Button
upload_btn = ttk.Button(frame, text="Upload Excel & Run", command=start_process)
upload_btn.pack(fill="x", pady=(10, 20))

root.mainloop()