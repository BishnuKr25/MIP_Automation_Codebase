from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os
import sys
import requests
import json
from datetime import datetime

# Credentials (update as needed)
USERNAME = "username"
PASSWORD = "password"

# Set up Chrome options
chrome_options = webdriver.ChromeOptions()
chrome_options.add_experimental_option("detach", True)
chrome_options.add_argument("--remote-debugging-port=9222")

# Initialize the driver and start the login process
driver = webdriver.Chrome(options=chrome_options)
driver.get("https://portal.mip.com/signin")

# --- Login Process ---
print("🔄 Starting login process...")

# Enter username and password
WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, "/html/body/div/form/div[2]/div[3]/div/input"))
).send_keys(USERNAME)

WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, "/html/body/div/form/div[2]/div[4]/div/input"))
).send_keys(PASSWORD + Keys.RETURN)

# Click "Send me a passcode" - new XPath
print("🔄 Requesting passcode...")
WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "/html/body/div/form/div[2]/div[5]/div/div/a[2]"))
).click()

# --- Auto-Enter MFA Code with File Reset ---
print("\n🚨 Preparing to fetch new MFA code... Clearing previous codes.")
# Clear the mfa_code.txt file to ensure only the new code is used
mfa_file = "mfa_code.txt"
if os.path.exists(mfa_file):
    os.remove(mfa_file)
    print("🗑️ Previous MFA codes deleted.")
else:
    print("ℹ️ No previous MFA file found, starting fresh.")

print("⏳ Waiting for new MFA code from Flask server...")
mfa_code = None
timeout = 300  # 5-minute timeout for MFA code
poll_interval = 2  # Seconds between poll attempts
start_time = time.time()

while time.time() - start_time < timeout:
    try:
        response = requests.get("http://localhost:8502/latest_mfa", timeout=5)
        if response.status_code == 200:
            data = response.json()
            # Check if a non-empty mfa_code is received
            if "mfa_code" in data and data["mfa_code"].strip():
                mfa_code = data["mfa_code"]
                print(f"✅ Received new MFA code: {mfa_code}")
                break
            else:
                print("⏳ No MFA code received yet, polling again...")
        else:
            print(f"⚠️ Server response: {response.status_code} - {response.text}")
    except requests.exceptions.RequestException as e:
        print(f"⚠️ Connection error: {str(e)}")
    time.sleep(poll_interval)

if not mfa_code:
    print("❌ No new MFA code received within timeout period")
    driver.quit()
    sys.exit(1)

# --- Enter MFA Code and Click Verify with new XPaths ---
# Enter passcode
mfa_field_xpath = "/html/body/div/form/div[2]/div[3]/div[1]/input"
mfa_field = WebDriverWait(driver, 10).until(
    EC.visibility_of_element_located((By.XPATH, mfa_field_xpath))
)
mfa_field.clear()
mfa_field.send_keys(mfa_code)

# Click Verify button 
verify_button_xpath = "/html/body/div/form/div[2]/div[7]/button"
verify_button = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, verify_button_xpath))
)
driver.execute_script("arguments[0].scrollIntoView(true);", verify_button)
time.sleep(1)  # Brief pause to ensure scroll completes
driver.execute_script("arguments[0].click();", verify_button)
print("🔄 Verify button clicked.")

# Click on "Launch Now" button
launch_now_xpath = "/html/body/div[2]/div/div[1]/div[3]/div/div[2]/div[3]/a"
launch_button = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, launch_now_xpath)))
# driver.execute_script("arguments[0].scrollIntoView(true);", launch_button)
launch_button.click()
print("🔄 Launch Now button clicked.")

# Switch to new window/tab
WebDriverWait(driver, 20).until(EC.number_of_windows_to_be(2))
driver.switch_to.window(driver.window_handles[1])
print("🔄 Switched to new window.")

# Wait for the MIP login page to load
WebDriverWait(driver, 30).until(
    EC.url_contains("https://login.mip.com")
)
print("✅ MIP login page loaded.")

# Enter username
username_field_xpath = "/html/body/div/ui-view/gateway-mfa-component/div/div[1]/ui-view/identity-component/form/div[2]/div[1]/input[1]"
WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, username_field_xpath))
).send_keys(USERNAME)

# Click Continue
continue_button_xpath = "/html/body/div/ui-view/gateway-mfa-component/div/div[1]/ui-view/identity-component/form/div[3]/button"
WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, continue_button_xpath))
).click()
print("🔄 Username entered and Continue clicked.")

# Enter password
password_field_xpath = "/html/body/div/ui-view/gateway-mfa-component/div/div[1]/ui-view/signin-component/form/div[2]/div/input[2]"
WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, password_field_xpath))
).send_keys(PASSWORD)

# Click Sign-in button
signin_button_xpath = "/html/body/div/ui-view/gateway-mfa-component/div/div[1]/ui-view/signin-component/form/div[4]/button"
WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, signin_button_xpath))
).click()
print("🔄 Password entered and Sign-in clicked.")

# Wait for dashboard to load
dashboard_timeout = 60  # seconds
dash_start = time.time()
while time.time() - dash_start < dashboard_timeout:
    if "https://adv.mip.com/#/dashboard/1" in driver.current_url:
        print("\n✅ Dashboard confirmed. Resuming automation...\n")
        break
    time.sleep(2)
else:
    print("❌ Dashboard not reached within timeout.")
    driver.quit()
    sys.exit(1)

# --- Navigate to Reports ---
WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "/html/body/mad-main/div/div/div[1]/mad-menu/div/div[6]/div"))
).click()

# Open 'General Ledger Analysis'
WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "/html/body/mad-main/div/div/div[1]/div[2]/div[1]/landing-page/mad-single-container/div/div[2]/div/div/div[2]/div/ng-transclude/mad-navigation-card[8]/div/ng-include/div/div[1]"))
).click()

# Expand 'General Ledger'
WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "/html/body/mad-main/div/div/div[1]/div[2]/div[1]/report-types/mad-single-container/div/div[2]/div/div/div[2]/div[2]/ng-transclude/mad-navigation-card[2]/div/ng-include/div/div[1]"))
).click()

# Select the specific report
WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "/html/body/mad-main/div/div/div[1]/div[2]/div[1]/saved-reports/mad-single-container/div/div[2]/div/div/div[2]/div/div/ng-transclude/mad-report-card[2]/div/div[3]/div/button[2]"))
).click()
time.sleep(7)

print("\n✅ Successfully navigated to 'Expanded General Ledger - This Year'!")

# --- Updated set_date Helper Function ---
def set_date(xpath, date_value, wait_time=20):
    date_field = WebDriverWait(driver, wait_time).until(
        EC.element_to_be_clickable((By.XPATH, xpath))
    )
    driver.execute_script("arguments[0].scrollIntoView(true);", date_field)
    time.sleep(0.5)
    date_field.click()
    date_field.clear()
    date_field.send_keys(date_value)
    driver.execute_script("arguments[0].dispatchEvent(new Event('change', { bubbles: true }));", date_field)
    time.sleep(1)

# --- Read Dates from extracted_dates.json ---
try:
    with open("extracted_dates.json", "r") as f:
        dates = json.load(f)
    from_date = dates["from"]
    to_date = dates["to"]
except Exception as e:
    print(f"❌ Error reading dates from JSON: {e}. Using defaults.")
    from_date = "2020-01-10"  # Fallback
    to_date = "2021-09-30"    # Fallback

# --- Set Report Date Fields ---
set_date("/html/body/mad-main/div/div/div[1]/div[2]/div[1]/report-builder/mad-single-container/div/div[2]/div/div/mad-form/div/div[2]/div/div/div[2]/div[2]/div/mad-report-dates/div/mad-paper/div/div/div[1]/div[1]/div/div[1]/mad-input/div/div/div/span/span/input", from_date)
set_date("/html/body/mad-main/div/div/div[1]/div[2]/div[1]/report-builder/mad-single-container/div/div[2]/div/div/mad-form/div/div[2]/div/div/div[2]/div[2]/div/mad-report-dates/div/mad-paper/div/div/div[1]/div[1]/div/div[2]/mad-input/div/div/div/span/span/input", from_date)
set_date("/html/body/mad-main/div/div/div[1]/div[2]/div[1]/report-builder/mad-single-container/div/div[2]/div/div/mad-form/div/div[2]/div/div/div[2]/div[2]/div/mad-report-dates/div/mad-paper/div/div/div/div[2]/div/div[2]/mad-input/div/div/div/span/span/input", from_date)
set_date("/html/body/mad-main/div/div/div[1]/div[2]/div[1]/report-builder/mad-single-container/div/div[2]/div/div/mad-form/div/div[2]/div/div/div[2]/div[2]/div/mad-report-dates/div/mad-paper/div/div/div/div[2]/div/div[3]/mad-input/div/div/div/span/span/input", to_date)

print("\n✅ Dates have been successfully updated!")

# --- Remove and Add Selections ---
WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "/html/body/mad-main/div/div/div[1]/div[2]/div[1]/report-builder/mad-single-container/div/div[2]/div/div/mad-form/div/div[2]/div/div/div[2]/div[2]/div/div[1]/mad-paper/div/div/div[2]/div[1]/div/div[1]/div[2]/div/button[2]"))
).click()

WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "/html/body/mad-main/div/div/div[1]/div[2]/div[1]/report-builder/mad-single-container/div/div[2]/div/div/mad-form/div/div[2]/div/div/div[2]/div[2]/div/div[1]/mad-paper/div/div/div[2]/div[1]/div/div[1]/div[2]/div/button[1]"))
).click()

for _ in range(24):
    WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "/html/body/mad-main/div/div/div[1]/div[2]/div[1]/report-builder/mad-single-container/div/div[2]/div/div/mad-form/div/div[2]/div/div/div[2]/div[2]/div/div[1]/mad-paper/div/div/div[2]/div[1]/div/div[2]/div/div[2]/div[2]/div[13]/div/div/ng-include/div[1]/div/span[2]/i"))
    ).click()
WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "/html/body/mad-main/div/div/div[1]/div[2]/div[1]/report-builder/mad-single-container/div/div[2]/div/div/mad-form/div/div[2]/div/div/div[1]/div/div[2]/span"))
).click()
WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "/html/body/mad-main/div/div/div[1]/div[2]/div[1]/report-builder/mad-single-container/div/div[2]/div/div/mad-form/div/div[2]/div/div/div[1]/div/div[3]/span"))
).click()
checkbox = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, "/html/body/mad-main/div/div/div[1]/div[2]/div[1]/report-builder/mad-single-container/div/div[2]/div/div/mad-form/div/div[2]/div/div/div[2]/div[6]/div/div/div[1]/mad-paper/div/div/div/div[3]/mad-input/div/div/fieldset/input"))
)
if not checkbox.is_selected():
    checkbox.click()
WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "/html/body/mad-main/div/div/div[1]/div[2]/div[1]/report-builder/mad-single-container/div/div[3]/div/div/div[2]/div[1]/button"))
).click()
WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "/html/body/mad-main/div/div/div[1]/div[2]/div[1]/report-builder/mad-single-container/div/div[3]/div/div/div[2]/div[1]/ul/li[2]"))
).click()
WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div/default-modal/div/div/div/div[3]/button[3]"))
).click()
print("\n⏳ Waiting for file download (approx 3 minutes)...")
time.sleep(180)
print("\n✅ Report is downloaded!")
print("\n🚀 Automation completed successfully! Closing browser...")
driver.quit()
