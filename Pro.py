import threading
import keyboard  # ✅ Emergency stop
import time
import pandas as pd  # ✅ Excel handling
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.edge.service import Service

# ✅ Emergency Stop Variable
stop_flag = False

import pandas as pd  # ✅ Excel handling

import warnings

# ✅ Suppress FutureWarnings
warnings.simplefilter(action='ignore', category=FutureWarning)

# =========================================
# STEP 1: Load and Process Excel Data Correctly
# =========================================

# Define file path and sheet name
file_path = "Dummy Details.xlsx"
sheet_name = "MAXIMO_READY"  # Ensure this matches your sheet name

# ✅ Load the data without headers
data = pd.read_excel(file_path, sheet_name=sheet_name, header=None)

# ✅ Identify all Employee Name rows
employee_indices = data[data.iloc[:, 0].astype(str).str.contains("Employee Name", na=False, case=False)].index.tolist()

# ✅ Extract and Process Each Employee's Data
employees = []
for i, emp_idx in enumerate(employee_indices):
    # ✅ Extract Employee Name & KNPC ID
    employee_name = data.iloc[emp_idx, 1]  # Name is in Column 2
    knpc_id = data.iloc[emp_idx + 1, 1]  # ID is in Column 2 (next row)

    # ✅ Find the header row (contains "Date (DD/MM/YYYY)")
    header_idx = emp_idx + 3  # Header row is 3 rows below the Employee Name row
    headers = data.iloc[header_idx].astype(str).str.strip()

    # ✅ Extract the employee's table data
    next_emp_idx = employee_indices[i + 1] if i + 1 < len(employee_indices) else len(data)
    emp_data = data.iloc[header_idx + 1: next_emp_idx].reset_index(drop=True)

    # ✅ Assign headers & remove "Unnamed" columns
    emp_data.columns = headers
    emp_data = emp_data.loc[:, ~emp_data.columns.str.contains("Unnamed", na=False)]

    # ✅ Convert Date column to explicit DD/MM/YYYY format
    if "Date (DD/MM/YYYY)" in emp_data.columns:
        emp_data["Date (DD/MM/YYYY)"] = pd.to_datetime(emp_data["Date (DD/MM/YYYY)"], errors="coerce")

        # ✅ Drop invalid dates before formatting
        emp_data = emp_data.dropna(subset=["Date (DD/MM/YYYY)"])

        # ✅ Explicitly format the date as string in DD/MM/YYYY to prevent misinterpretation
        emp_data["Date (DD/MM/YYYY)"] = emp_data["Date (DD/MM/YYYY)"].dt.strftime("%d/%m/%Y")

    # ✅ Remove any "Total" rows
    emp_data = emp_data[~emp_data["Date (DD/MM/YYYY)"].astype(str).str.contains("Total", na=False, case=False)]

    # ✅ Replace NaN with empty strings
    emp_data = emp_data.fillna("").infer_objects(copy=False)

    # ✅ Store Employee Data
    employees.append({"knpc_id": knpc_id, "name": employee_name, "data": emp_data})

# ✅ Show only a simple confirmation message
print("\n✅ **Step 1 Completed: Employee Data Successfully Extracted!** 🚀")

# =========================================
# STEP 2: Connect to Existing Edge Session
# =========================================

edge_driver_path = r"D:\FLARE\Automated Script\Programm\edgedriver_win64\msedgedriver.exe"
options = webdriver.EdgeOptions()
options.debugger_address = "localhost:9222"

try:
    service = Service(edge_driver_path)
    driver = webdriver.Edge(service=service, options=options)
    print("✅ Successfully attached to existing Edge session.")
except Exception as e:
    print(f"❌ ERROR: Unable to attach to Edge session: {e}")
    exit()

# Identify tabs
maximo_tab = None
kconnect_tab = None
for handle in driver.window_handles:
    driver.switch_to.window(handle)
    if "Labor Reporting" in driver.title:
        maximo_tab = handle
    elif driver.current_url.startswith("https://webportal.knpc.com/"):
        kconnect_tab = handle

if not maximo_tab or not kconnect_tab:
    print("❌ ERROR: Could not find both Maximo and KConnect tabs!")
    exit()

driver.switch_to.window(maximo_tab)
print("✅ Starting in Maximo tab.")

# **Initialize the refresh timer**
last_refresh_time = time.time()

## =========================================
# ✅ STEP 3: HELPER FUNCTIONS FOR MAXIMO AUTOMATION
# =========================================

import time
import threading
import keyboard
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ✅ Emergency Stop Listener
def listen_for_stop():
    """Listens for the 'Q' key to stop the script immediately."""
    global stop_flag
    while True:
        if keyboard.is_pressed("q"):
            print("\n🚨 Emergency Stop Activated! Exiting safely... 🚨")
            stop_flag = True  
            driver.quit()
            exit()

threading.Thread(target=listen_for_stop, daemon=True).start()

# ✅ Function to Ensure Maximo is Fully Loaded
def wait_for_maximo_load(step_name="Unknown Step"):
    """Waits for Maximo to fully load before proceeding."""
    retries = 3  
    for attempt in range(retries):
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//input[contains(@id,'search')]"))
            )
            print(f"✅ Maximo fully loaded and ready for input at step: {step_name}.")
            return True  
        except:
            print(f"⚠️ Attempt {attempt + 1} failed: Maximo not fully loaded. Retrying in 20 seconds...")
            time.sleep(20)  

    print(f"❌ Maximo is not responding after 3 attempts at step: {step_name}.")
    return False  

# ✅ Detect Whether Maximo is in ID Filling Page or Hour Filling Page
def get_maximo_page():
    """Determines if Maximo is on the ID Filling Page or Hour Filling Page."""
    try:
        if WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.XPATH, "//button[contains(text(),'New Row')]"))):
            return "Hour Filling Page"
    except:
        pass  

    try:
        if WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.XPATH, "//input[contains(@id,'search')]"))):
            return "ID Filling Page"
    except:
        return "Unknown"

    return "Unknown"

# ✅ Function to Enter Employee ID in Maximo
def enter_knpc_id(knpc_id):
    """Enters the employee KNPC ID and ensures we reach the Hour Filling Page."""
    if get_maximo_page() == "Hour Filling Page":
        print(f"✅ KNPC ID {knpc_id} is already processed. Skipping ID entry.")
        return True  

    if not wait_for_maximo_load("ID Filling Page"):
        return False  

    retries = 3
    for attempt in range(retries):
        if get_maximo_page() == "Hour Filling Page":
            print(f"✅ KNPC ID {knpc_id} is already processed. Skipping ID entry.")
            return True  

        try:
            print(f"🔍 Attempt {attempt + 1}: Entering KNPC ID: {knpc_id}")
            search_box = WebDriverWait(driver, 7).until(
                EC.presence_of_element_located((By.XPATH, "//input[contains(@id,'search')]"))
            )
            search_box.clear()
            search_box.send_keys(knpc_id)
            search_box.send_keys(Keys.RETURN)
            time.sleep(3)  

            if get_maximo_page() == "Hour Filling Page":
                print(f"✅ KNPC ID {knpc_id} entered successfully. Now on Hour Filling Page.")
                return True  
            
        except Exception as e:
            print(f"⚠️ KNPC ID entry failed: {e}. Retrying...")
            time.sleep(5)

    print(f"❌ ERROR: Failed to enter KNPC ID {knpc_id} after 3 attempts. Skipping employee.")
    return False
    time.sleep(3)

# ✅ Function to Click "New Row" in Maximo
def click_new_row():
    """Clicks 'New Row' button in Maximo's Hour Filling Page."""
    retries = 3
    for attempt in range(retries):
        try:
            print("🔄 Clicking 'New Row' button...")
            new_row_button = WebDriverWait(driver, 7).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'New Row')]"))
            )
            new_row_button.click()
            time.sleep(5)  

            print("✅ 'New Row' button clicked successfully.")
            return True  

        except Exception as e:
            print(f"⚠️ Attempt {attempt + 1} failed to click 'New Row'. Retrying in 5 seconds...")
            time.sleep(5)

    print("❌ ERROR: Failed to click 'New Row' after 3 attempts.")
    return False

# ✅ Function to Refresh KConnect & Maximo Every 5 Minutes
def refresh_kconnect_maximo():
    """Refreshes KConnect & Maximo every 5 minutes to keep sessions active."""
    global last_refresh_time
    elapsed_time = time.time() - last_refresh_time

    if elapsed_time >= 5 * 60:
        print(f"\n⏳ {round(elapsed_time / 60, 2)} minutes reached, refreshing KConnect & Maximo...")
        for tab, name in [(kconnect_tab, "KConnect"), (maximo_tab, "Maximo")]:
            driver.switch_to.window(tab)
            driver.refresh()
            time.sleep(5)  
            print(f"✅ {name} refreshed successfully.")

        print("⏳ Waiting 5 seconds before starting next employee...")
        time.sleep(5)
        last_refresh_time = time.time()
    else:
        print(f"⏳ Elapsed time: {round(elapsed_time / 60, 2)} minutes. Still within limit, skipping refresh.")

# ✅ Function to Enter Data Fields (Handles Focus & Enter Key)
def enter_data(xpath, value, label, press_enter=False, last_row=False):
    """Enters data into a specified field in Maximo and handles focus issues."""
    if stop_flag or value in [None, "", "nan"]:
        return  

    retries = 3
    for attempt in range(retries):
        try:
            field = WebDriverWait(driver, 7).until(EC.presence_of_element_located((By.XPATH, xpath)))
            field.click()
            time.sleep(2)
            field.clear()
            field.send_keys(value)
            time.sleep(1.5)  

            if press_enter:
                field.send_keys(Keys.RETURN)
                print(f"✅ {label} entered successfully: {value} (Enter pressed)")
                time.sleep(3)

            return  

        except Exception as e:
            print(f"⚠️ Attempt {attempt + 1} failed to enter {label}. Retrying in 6 seconds...")
            time.sleep(6)  

    print(f"❌ ERROR: Failed to enter {label} after {retries} attempts.")


# =========================================
# ✅ STEP 4: PROCESS EMPLOYEES & ENTER DATA
# =========================================

# ✅ Function to Format Hours Correctly
def format_hours(value):
    """ Ensure hours are properly formatted as '0:00' """
    try:
        return f"{int(float(value))}:00" if str(value).strip() else "0:00"
    except ValueError:
        return "0:00"  # Default if conversion fails

# ✅ Process Each Employee
for employee in employees:
    if stop_flag:
        break  

    knpc_id = employee["knpc_id"]
    employee_name = employee["name"]
    emp_data = employee["data"]

    print(f"\n🔄 Processing Employee: {employee_name} | KNPC ID: {knpc_id}")

    # ✅ Ensure Maximo is on the ID Filling Page
    if not enter_knpc_id(knpc_id):
        continue  

    first_entry = True  # Track if it's the first entry for this employee

    # ✅ Iterate Over Each Entry in the Employee's Work Data
    for index, test_entry in emp_data.iterrows():
        if stop_flag:
            break  

        # ✅ Check for Empty Row (Stop Processing if Found)
        if test_entry.isnull().any() or test_entry.astype(str).str.strip().eq("").all():
            print("⚠️ Detected empty row, stopping employee entry.")
            break  

        # ✅ Ensure "New Row" is clicked only ONCE per employee
        if first_entry:
            print("🔄 Clicking 'New Row' button for first entry.")
            if not click_new_row():
                print(f"❌ Skipping Entry: WO = {test_entry.get('Work Order', 'N/A')} | Date = {test_entry.get('Date (DD/MM/YYYY)', 'N/A')} | Hours = {test_entry.get('Regular Hours', 'N/A')}")
                break  # Skip the entire employee entry if "New Row" fails
            first_entry = False  

        # ✅ Ensure Maximo is on the Hour Filling Page Before Entering Data
        wait_for_maximo_load(step_name="Hour Filling Page")

        # ✅ Check if This is the Last Row
        last_row = index == len(emp_data) - 1  

        # ✅ Check if There is a Next Row with Valid Data
        next_row_has_data = (
            index + 1 < len(emp_data)
            and "Date (DD/MM/YYYY)" in emp_data.columns
            and isinstance(emp_data.iloc[index + 1]["Date (DD/MM/YYYY)"], str)
            and emp_data.iloc[index + 1]["Date (DD/MM/YYYY)"].strip() != ""
        )

        # ✅ Extract Work Order, Date, and Overtime Fields (Ensure Proper Formatting)
        work_order = test_entry.get("Work Order", "N/A")
        date = test_entry.get("Date (DD/MM/YYYY)", "N/A")
        regular_hours = format_hours(test_entry.get("Regular Hours", ""))
        normal_ot = format_hours(test_entry.get("Normal OT", ""))
        friday_ot = format_hours(test_entry.get("Friday OT", ""))
        holiday_ot = format_hours(test_entry.get("Holiday OT", ""))

        print(f"📝 Entering Data: WO = {work_order} | Date = {date}")

        # ✅ Enter Work Order & Date
        enter_data("//*[@id='m867d5646-tb']", work_order, "Work Order")
        enter_data("//*[@id='mc4a7c56c-tb']", date, "Date")

        # ✅ Enter Hours (Only if the Value is Not "0:00")
        if regular_hours != "0:00":
            enter_data("//*[@id='m4d450696-tb']", regular_hours, "Regular Hours", press_enter=next_row_has_data, last_row=last_row)

        if normal_ot != "0:00":
            enter_data("//*[@id='m29ceedaf-tb']", normal_ot, "Normal OT", press_enter=next_row_has_data, last_row=last_row)

        if friday_ot != "0:00":
            enter_data("//*[@id='m1695cf5f-tb']", friday_ot, "Friday OT", press_enter=next_row_has_data, last_row=last_row)

        if holiday_ot != "0:00":
            enter_data("//*[@id='m5ec9dd39-tb']", holiday_ot, "Holiday OT", press_enter=next_row_has_data, last_row=last_row)

        # ✅ Ensure Proper Focus Shift on Last Row
        if last_row:
            try:
                # ✅ If no valid date in the next row, click another field instead of pressing Enter
                print("✅ Shifted focus for last row.")
                focus_field = WebDriverWait(driver, 7).until(
                    EC.presence_of_element_located((By.XPATH, "//*[@id='m1695cf5f-tb']"))  # Click Friday OT for focus shift
                )
                focus_field.click()
            except Exception as e:
                print(f"⚠️ WARNING: Could not shift focus on the last row. Error: {e}")

    # ✅ Click 'List' Button to Return to Main Page
    print("🔄 Clicking 'List' button...")
    try:
        list_button = WebDriverWait(driver, 7).until(
            EC.presence_of_element_located((By.XPATH, "//*[@id='m9fa3e414-tab_anchor']"))
        )
        list_button.click()
        time.sleep(5)  
    except Exception as e:
        print(f"⚠️ WARNING: Could not click 'List' button. {e}")

    # ✅ Click 'Yes' Button to Confirm
    print("🔄 Clicking 'Yes' button...")
    try:
        yes_button = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, "//*[@id='me1720906-pb']"))
        )
        yes_button.click()
        time.sleep(5)  
        print("✅ 'Yes' button clicked.")
    except Exception as e:
        print(f"⚠️ WARNING: Could not click 'Yes' button. {e}")

    # ✅ Refresh KConnect & Maximo Every 5 Minutes
    refresh_kconnect_maximo()

print("\n✅ **Test Completed for ALL Employees! Entries are now SAVED in Maximo! 🚀**")

