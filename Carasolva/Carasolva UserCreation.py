# Carasolva_UserCreation.py
# Usage:
#   python Carasolva_UserCreation.py <username> <password> --file "C:\path\to\users.xlsx" --role "Non Med Cert Staff" --driver "C:\path\to\msedgedriver.exe"
#
# Notes:
# - Supports .xlsx/.xls/.csv
# - Accepts flexible headers for first/last/email/employee id
# - Keeps your original log format ([OK], [ERROR], [WARN], [FAIL])

import argparse
import os
import sys
import time
import pandas as pd

from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    NoSuchElementException,
    TimeoutException,
    ElementNotInteractableException,
    StaleElementReferenceException,
)

# ----------------------------
# Args
# ----------------------------
def parse_args():
    p = argparse.ArgumentParser(
        description="Create Carasolva users from a spreadsheet (.xlsx/.xls/.csv)."
    )
    p.add_argument("username", help="Carasolva username")
    p.add_argument("password", help="Carasolva password")
    p.add_argument(
        "--file",
        required=True,
        help="Path to spreadsheet (.xlsx, .xls, or .csv) containing users",
    )
    p.add_argument(
        "--driver",
        default=r"C:\Users\kjcalkins\Downloads\edgedriver_win64\msedgedriver.exe",
        help="Path to msedgedriver.exe",
    )
    p.add_argument(
        "--role",
        default="Non Med Cert Staff",
        help="Role to assign (visible text in the dropdown)",
    )
    p.add_argument(
        "--pause-on-exit",
        action="store_true",
        help="Pause for Enter before quitting the browser",
    )
    return p.parse_args()

# ----------------------------
# Flexible file reader
# ----------------------------
def read_users_table(path):
    if not os.path.isfile(path):
        print(f"[ERROR] File not found: {path}")
        sys.exit(1)

    ext = os.path.splitext(path)[1].lower()
    try:
        if ext in [".xlsx", ".xls"]:
            df = pd.read_excel(path)
        elif ext == ".csv":
            df = pd.read_csv(path)
        else:
            print(f"[ERROR] Unsupported file extension: {ext}. Use .xlsx, .xls, or .csv")
            sys.exit(1)
    except Exception as e:
        print(f"[ERROR] Failed to read file: {e}")
        sys.exit(1)

    # Normalize column names
    norm_map = {}
    for c in df.columns:
        key = (
            str(c)
            .strip()
            .lower()
            .replace("_", " ")
            .replace("-", " ")
            .replace(".", " ")
        )
        key = " ".join(key.split())  # collapse spaces
        norm_map[c] = key
    df.rename(columns=norm_map, inplace=True)

    # Candidate sets for each required field
    FIRST_CANDS = {"first name", "firstname", "first", "f name", "given name"}
    LAST_CANDS = {"last name", "lastname", "last", "l name", "surname", "family name"}
    EMAIL_CANDS = {"email", "e-mail", "email address", "mail"}
    EMP_CANDS = {"employee id", "employee number", "emp id", "emp number", "employeeid"}

    def choose(colset):
        for c in df.columns:
            if c in colset:
                return c
        return None

    col_first = choose(FIRST_CANDS)
    col_last = choose(LAST_CANDS)
    col_email = choose(EMAIL_CANDS)
    col_emp = choose(EMP_CANDS)  # optional

    missing = []
    if not col_first: missing.append("First Name")
    if not col_last:  missing.append("Last Name")
    if not col_email: missing.append("Email")

    if missing:
        print(
            "[ERROR] Missing required columns. Need at least: First Name, Last Name, Email.\n"
            f"Detected headers: {list(df.columns)}"
        )
        sys.exit(1)

    # Build normalized view (don’t alter original data types more than needed)
    norm_rows = []
    for _, r in df.iterrows():
        first = str(r.get(col_first, "")).strip()
        last = str(r.get(col_last, "")).strip()
        email = str(r.get(col_email, "")).strip()
        emp = r.get(col_emp, "")
        if pd.isna(emp):
            emp = ""
        emp = str(emp).strip()

        # Skip completely empty rows
        if not (first or last or email):
            continue

        norm_rows.append(
            {"First Name": first, "Last Name": last, "Email": email, "Employee ID": emp}
        )

    out_df = pd.DataFrame(norm_rows, columns=["First Name", "Last Name", "Email", "Employee ID"])
    print(f"[OK] Loaded {len(out_df)} users from '{os.path.basename(path)}'.")
    return out_df

# ----------------------------
# Selenium helpers
# ----------------------------
def fill_field(driver, field_id, value, label="Field"):
    for attempt in range(3):
        try:
            element = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, field_id))
            )
            driver.execute_script("arguments[0].scrollIntoView(true);", element)
            element.clear()
            element.send_keys(value)
            print(f"[OK] Filled {label} with: {value}")
            return True
        except (ElementNotInteractableException, StaleElementReferenceException, TimeoutException) as e:
            print(f"[WARN] Attempt {attempt + 1} failed for {label}: {e}")
            time.sleep(1)
            continue
    print(f"[FAIL] Failed to fill {label} after retries.")
    return False

def main():
    args = parse_args()

    username = args.username
    password = args.password
    driver_path = args.driver
    role_text = args.role
    file_path = args.file

    # Read spreadsheet (flex headers + csv/xlsx)
    df = read_users_table(file_path)
    if df.empty:
        print("[ERROR] No valid rows found in the spreadsheet.")
        sys.exit(1)

    # Setup Selenium
    service = Service(driver_path)
    driver = webdriver.Edge(service=service)

    # Screenshot dir next to data file for convenience
    screenshot_dir = os.path.join(os.path.dirname(file_path) or ".", "carasolva_errors")
    os.makedirs(screenshot_dir, exist_ok=True)

    try:
        # Login
        driver.get("https://carasolva-medsupport-prod.azurewebsites.net/Masters/UserLogin.aspx")
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.ID, "txtUserName")))
        driver.find_element(By.ID, "txtUserName").send_keys(username)
        driver.find_element(By.ID, "txtPassword").send_keys(password + Keys.RETURN)
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        driver.maximize_window()

        # Expand company
        try:
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//img[@alt='Expand Company ']"))
            ).click()
            print("[OK] Clicked '+' to expand Company.")
        except Exception as e:
            print(f"[ERROR] Error expanding company: {e}")

        # Go to Users
        try:
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.LINK_TEXT, "Users"))
            ).click()
            print("[OK] Clicked on 'Users' link.")
        except Exception as e:
            print(f"[ERROR] Error clicking Users link: {e}")

        # Main loop
        for index, row in df.iterrows():
            first = row["First Name"]
            last = row["Last Name"]
            email = row["Email"]
            emp_id = row["Employee ID"]

            print(f"\nCreating user {index + 1}: {first} {last}")

            try:
                # New user
                WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "ctl00_DefaultContent_btnActionButtons_btnAdd"))
                ).click()

                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.ID, "ctl00_DefaultContent_uscPersonControl_txtSearchFName"))
                )

                # Search to skip duplicates
                fill_field(driver, "ctl00_DefaultContent_uscPersonControl_txtSearchFName", first, "First Name (Search)")
                fill_field(driver, "ctl00_DefaultContent_uscPersonControl_txtSearchLName", last, "Last Name (Search)")
                fill_field(driver, "ctl00_DefaultContent_uscPersonControl_txtSearchEmail", email, "Email (Search)")

                driver.find_element(By.ID, "ctl00_DefaultContent_uscPersonControl_imgbtnSearch").click()
                time.sleep(3)

                if "No records found" not in driver.page_source:
                    print("User already exists. Skipping.")
                    try:
                        WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.ID, "ctl00_DefaultContent_btnReturn"))
                        ).click()
                        WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.ID, "ctl00_DefaultContent_btnActionButtons_btnAdd"))
                        )
                        print("Returned to User List page.")
                    except Exception as e:
                        print(f"[ERROR] Error returning to User List after skip: {e}")
                    continue

                # Begin creation
                WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "ctl00_DefaultContent_uscPersonControl_btnSave"))
                ).click()

                WebDriverWait(driver, 15).until(
                    EC.presence_of_element_located((By.ID, "ctl00_DefaultContent_uscPersonControl_txtFirstName"))
                )

                # Fill details
                fill_field(driver, "ctl00_DefaultContent_uscPersonControl_txtFirstName", first, "First Name")
                fill_field(driver, "ctl00_DefaultContent_uscPersonControl_txtLastName", last, "Last Name")

                initials = (first[:1] if first else "") + (last[:1] if last else "")
                fill_field(driver, "ctl00_DefaultContent_uscPersonControl_txtInitials", initials.upper(), "Initials")

                fill_field(driver, "ctl00_DefaultContent_txtLogin", email, "User Name")
                if emp_id:
                    fill_field(driver, "ctl00_DefaultContent_txtSSOLogin", str(emp_id), "SSO Login")
                    fill_field(driver, "ctl00_DefaultContent_txtEmployeeNumber", str(emp_id), "Employee Number")
                else:
                    print("[WARN] Employee ID missing/blank; skipping SSO/Employee Number.")

                fill_field(driver, "ctl00_DefaultContent_txtPassword", "A12345", "Password")
                fill_field(driver, "ctl00_DefaultContent_txtConfirmPassword", "A12345", "Confirm Password")
                fill_field(driver, "ctl00_DefaultContent_uscPersonControl_txtTitle", "DSP", "Title")

                # Role
                try:
                    role_dropdown = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.ID, "ctl00_DefaultContent_cmbRole"))
                    )
                    Select(role_dropdown).select_by_visible_text(role_text)
                    print(f"[OK] Selected Role: {role_text}")
                except Exception as e:
                    print(f"[FAIL] Failed to select Role '{role_text}': {e}")

                # Save
                WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "ctl00_DefaultContent_btnActionButtons_btnSave"))
                ).click()
                print("[OK] Clicked Save button to finalize user creation.")

                time.sleep(2)

                # Back to list
                WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "ctl00_DefaultContent_btnReturn"))
                ).click()
                WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "ctl00_DefaultContent_btnActionButtons_btnAdd"))
                )
                print("Returned to User List page, ready for next user.")

            except Exception as e:
                print(f"[FAIL] Error during user creation: {e}")
                screenshot_path = os.path.join(screenshot_dir, f"user_{index + 1}_error.png")
                try:
                    driver.save_screenshot(screenshot_path)
                    print(f"[INFO] Screenshot saved to {screenshot_path}")
                except Exception as se:
                    print(f"[WARN] Failed to save screenshot: {se}")
                # Try to recover back to list
                try:
                    WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.ID, "ctl00_DefaultContent_btnReturn"))
                    ).click()
                    WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.ID, "ctl00_DefaultContent_btnActionButtons_btnAdd"))
                    )
                    print("Returned to User List page after error.")
                except Exception:
                    pass
                continue

        print("\nAll users processed.")

    finally:
        if args.pause_on_exit:
            try:
                input("\nScript complete. Press Enter to close browser...")
            except EOFError:
                pass
        driver.quit()


if __name__ == "__main__":
    main()
