"""
Telephone Number Finder
=========================
What this program does:
1. Receives the path to an Excel file (command-line argument) - the same file
   that app3.py creates.
2. Opens gulesider.no via Selenium and, for each shareholder on the
   "Private 2024" and "Selskap 2024" sheets, looks up a phone number using
   name + zip code.
3. Writes the found numbers into the "TELEPHONE" column and saves the file
   with a " (T)" suffix.

telephone2.py version notes:
- MAIN BUG FIX: removed the cv2, numpy, PIL.Image, BytesIO imports - they
  were never used anywhere in the code, and the opencv-python (cv2) library
  simply wasn't installed, which made the script crash right at startup
  with ModuleNotFoundError.
- Added headless mode (needed to run on a server, where there is no display).
- Added column checks ("SHAREHOLDER", "ZIP CODE / CITY", "TELEPHONE")
  before processing, so the script doesn't crash with a KeyError if a
  column is missing.
- driver.quit() now always runs (even if something fails) - previously
  Chrome could stay open/hang in memory after an error.
- Renaming the final file no longer crashes if a file with the " (T)"
  suffix already exists.
- Code split into sections with comments.
- FIX: force stdout/stderr to UTF-8 so the script doesn't crash on Windows
  terminals using a legacy codepage (e.g. cp1251) when printing emoji
  characters like the ones used in the log messages below.
"""

# ============================================================
# 0. FORCE UTF-8 OUTPUT (fixes UnicodeEncodeError on Windows)
# ============================================================
import sys
import io

# On Windows, the default console codepage (e.g. cp1251) often can't encode
# emoji characters. Reconfigure stdout/stderr to UTF-8 so print() never
# crashes on these log messages, regardless of the terminal's codepage.
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
else:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")


# ============================================================
# 1. IMPORTS
# ============================================================
import time
import re
import os
import random

import pandas as pd
from openpyxl import load_workbook

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from webdriver_manager.chrome import ChromeDriverManager

# Note: cv2 (opencv-python), numpy and PIL.Image were in the original file
# but never used anywhere - they were removed so the script doesn't require
# extra dependencies that aren't even in requirements.


# ============================================================
# 2. READ COMMAND-LINE ARGUMENTS
# ============================================================
if len(sys.argv) < 2:
    print("❌ Error: No file path provided!")
    sys.exit(1)

file_path = sys.argv[1]  # Excel file path received from app3.py
sheets = ["Private 2024", "Selskap 2024"]  # sheets that need to be processed

print(f"ℹ️ Received file path: {file_path}")
print(f"ℹ️ File exists: {os.path.exists(file_path)}")
if os.path.exists(file_path):
    print(f"ℹ️ File size: {os.path.getsize(file_path)} bytes")
else:
    print("❌ ERROR: File does not exist!")
    sys.exit(1)


# ============================================================
# 3. CHROME SETUP
# ============================================================
def build_chrome_driver(headless=True):
    """
    Creates a Chrome WebDriver configured to block ads and run reliably on
    a server (headless).

    headless=True is required on a server (no graphical display there).
    It's also recommended to test locally with headless=True, so behavior
    matches what happens on the server.
    """
    options = Options()

    if headless:
        options.add_argument("--headless=new")

    # Ad blocking / stability
    options.add_argument("--disable-extensions-except")
    options.add_argument("--disable-plugins-discovery")
    options.add_argument("--disable-web-security")
    options.add_argument("--disable-features=VizDisplayCompositor")
    options.add_argument("--disable-ipc-flooding-protection")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")

    options.add_experimental_option("prefs", {
        "profile.default_content_setting_values": {
            "notifications": 2,   # block notifications
            "popups": 2,          # block popups
            "media_stream": 2,    # block camera/microphone access
            "plugins": 2,
            "images": 1,          # images are needed for the site to work properly
            "javascript": 1,
        },
        "profile.managed_default_content_settings": {"images": 1}
    })

    options.add_argument(
        "--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )

    service = Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=service, options=options)


driver = build_chrome_driver(headless=True)
print("✅ ChromeDriver started successfully")


# ============================================================
# 4. FUNCTIONS FOR HANDLING ADS / MODAL WINDOWS
# ============================================================
def detect_and_close_ad(driver):
    """Tries to find and close ads/popups using several strategies."""
    print(f"ℹ️ Searching for ads at {time.strftime('%H:%M:%S', time.localtime())}")

    google_ad_selectors = [
        "iframe[src*='googleads']",
        "iframe[src*='googlesyndication']",
        "iframe[src*='doubleclick']",
        "div[id*='google_ads']",
        "div[class*='google-ad']",
        "div[data-google-av-cxn]",
        "div[data-google-av-cid]",
        "div[style*='position: fixed'][style*='z-index']",
        "div[style*='position: absolute'][style*='z-index']",
        "[style*='z-index: 9999']",
        "[style*='z-index: 999999']",
        ".modal-overlay",
        ".modal-backdrop",
        ".ad-overlay"
    ]

    close_button_selectors = [
        "button[aria-label*='close' i]",
        "button[aria-label*='lukk' i]",  # "close" in Norwegian
        "button[title*='close' i]",
        "button[title*='lukk' i]",
        "button.close",
        "button[class*='close']",
        "span.close",
        "div.close",
        "[data-dismiss]",
        "[data-close]",
        ".ad-close",
        ".close-btn",
        ".close-button"
    ]

    try:
        time.sleep(2)  # give ads time to load

        # Strategy 1: hide/remove ad blocks
        for selector in google_ad_selectors:
            try:
                ad_elements = driver.find_elements(By.CSS_SELECTOR, selector)
                for ad_element in ad_elements:
                    try:
                        driver.execute_script("arguments[0].style.display = 'none';", ad_element)
                    except Exception:
                        pass
                    if ad_element.tag_name == 'iframe':
                        try:
                            driver.execute_script("arguments[0].remove();", ad_element)
                        except Exception:
                            pass
            except Exception:
                continue

        # Strategy 2: click the close button, if found
        for selector in close_button_selectors:
            try:
                close_buttons = driver.find_elements(By.CSS_SELECTOR, selector)
                for button in close_buttons:
                    try:
                        if button.is_displayed() and button.is_enabled():
                            driver.execute_script("arguments[0].click();", button)
                            time.sleep(1)
                            return True
                    except Exception:
                        continue
            except Exception:
                continue

        # Strategy 3: find close button via XPath by text
        xpath_selectors = [
            "//button[contains(text(), '×')]",
            "//button[contains(text(), '✕')]",
            "//span[contains(text(), '×')]",
            "//div[contains(text(), '×')]",
            "//button[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'close')]",
            "//button[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'lukk')]",
            "//button[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'skip')]",
            "//button[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'hopp over')]"
        ]
        for xpath in xpath_selectors:
            try:
                close_buttons = driver.find_elements(By.XPATH, xpath)
                for button in close_buttons:
                    try:
                        if button.is_displayed() and button.is_enabled():
                            driver.execute_script("arguments[0].click();", button)
                            time.sleep(1)
                            return True
                    except Exception:
                        continue
            except Exception:
                continue

        # Strategy 4: Escape key
        try:
            driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
            time.sleep(1)
        except Exception as e:
            print(f"ℹ️ Could not press Escape: {e}")

        # Strategy 5: click in "safe" screen areas to close the overlay
        safe_click_areas = [(10, 10), (10, 100), (100, 10), (50, 50)]
        for x, y in safe_click_areas:
            try:
                action = ActionChains(driver)
                action.move_by_offset(x, y).click().perform()
                time.sleep(0.5)
                action.move_by_offset(-x, -y).perform()
            except Exception as e:
                print(f"ℹ️ Could not click at ({x}, {y}): {e}")
                continue

        print("ℹ️ Ad-closing attempts finished")
        return False

    except Exception as e:
        print(f"⚠️ Error in detect_and_close_ad: {e}")
        return False


def check_for_blocking_elements(driver):
    """Checks whether the page has large elements that might be blocking content (ads/modals)."""
    blocking_selectors = [
        "[style*='z-index: 9999']",
        "[style*='z-index: 999999']",
        "iframe[src*='google']",
        ".modal-overlay",
        ".modal-backdrop",
        "div[style*='position: fixed']",
        "div[style*='position: absolute'][style*='top: 0'][style*='left: 0']"
    ]
    for selector in blocking_selectors:
        try:
            elements = driver.find_elements(By.CSS_SELECTOR, selector)
            for element in elements:
                if element.is_displayed():
                    size = element.size
                    if size['height'] > 200 and size['width'] > 200:
                        print(f"⚠️ Found a potentially blocking element: {selector}")
                        return True
        except Exception:
            continue
    return False


# ============================================================
# 5. INITIAL SITE LOAD
# ============================================================
driver.get("https://www.gulesider.no/")

try:
    WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, "button.css-1b11xpz"))
    )
    button = driver.find_element(By.CSS_SELECTOR, "button.css-1b11xpz")
    driver.execute_script("arguments[0].click();", button)
    print("✅ Button clicked (e.g. cookie consent)")
    time.sleep(2)
except Exception as e:
    print(f"ℹ️ Button not found or could not be clicked (it may not have appeared): {e}")

try:
    detect_and_close_ad(driver)
except Exception as e:
    print(f"⚠️ Initial ad handling failed: {e}")


# ============================================================
# 6. LOOK UP THE PHONE NUMBER FOR ONE SHAREHOLDER
# ============================================================
def get_phone_number(name, zip_code):
    """Searches gulesider.no for a phone number by name + zip code. Returns "" if not found."""
    try:
        if check_for_blocking_elements(driver):
            print("⚠️ Blocking elements found, closing ads...")
            detect_and_close_ad(driver)
            time.sleep(1)

        search_bar = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.ID, "searchbox-input-app-input"))
        )
        search_bar.clear()
        search_bar.send_keys(f"{name} {zip_code}")
        search_bar.send_keys(Keys.RETURN)
        time.sleep(2)

        if check_for_blocking_elements(driver):
            print("⚠️ Blocking elements found after search, closing ads...")
            detect_and_close_ad(driver)
            time.sleep(1)

        phone_elements = None
        max_attempts = 3

        for attempt in range(max_attempts):
            try:
                phone_elements = WebDriverWait(driver, 4).until(
                    EC.presence_of_all_elements_located((By.CLASS_NAME, "text-nowrap"))
                )
                if phone_elements:
                    break
            except TimeoutException:
                if attempt < max_attempts - 1:
                    print(f"ℹ️ Attempt {attempt + 1} failed, checking for ads...")
                    if check_for_blocking_elements(driver):
                        detect_and_close_ad(driver)
                        time.sleep(1)
                    continue
                else:
                    return "-"

        if not phone_elements:
            return "-"

        phone_number = phone_elements[0].text.strip()

        # If the number is truncated behind a "Vis" (show) link - open the full card
        if "Vis" in phone_number:
            try:
                vis_button = WebDriverWait(driver, 4).until(
                    EC.element_to_be_clickable((By.CLASS_NAME, "text-nowrap"))
                )
                vis_button.click()
                time.sleep(2)

                if check_for_blocking_elements(driver):
                    print("⚠️ Blocking elements found on the details page, closing ads...")
                    detect_and_close_ad(driver)
                    time.sleep(1)

                full_phone_element = WebDriverWait(driver, 4).until(
                    EC.presence_of_element_located((By.XPATH, "//p[contains(@class, 'text-2xl')]"))
                )
                phone_number = full_phone_element.text.strip()

            except TimeoutException:
                return "Error retrieving full number"

        if re.search(r"Endre bedriftsinformasjon", phone_number):
            return ""

        return phone_number

    except Exception as e:
        print(f"⚠️ Error in get_phone_number for '{name}': {e}")
        try:
            detect_and_close_ad(driver)
        except Exception:
            pass
        return ""


# ============================================================
# 7. LOAD THE EXCEL FILE
# ============================================================
try:
    print(f"ℹ️ Loading file: {file_path}")
    wb = load_workbook(file_path)
    print(f"✅ File loaded. Sheets: {wb.sheetnames}")
except Exception as e:
    print(f"❌ Could not load the Excel file: {e}")
    driver.quit()
    sys.exit(1)


# ============================================================
# 8. PROCESS SHEETS: LOOK UP AND WRITE PHONE NUMBERS
# ============================================================
try:
    for sheet_name in sheets:
        if sheet_name not in wb.sheetnames:
            print(f"⚠️ Sheet '{sheet_name}' not found in the file, skipping")
            continue

        print(f"ℹ️ Processing sheet: {sheet_name}")
        df = pd.read_excel(file_path, sheet_name=sheet_name)

        # Check that the required columns exist, to avoid a KeyError
        required_columns = ["SHAREHOLDER", "ZIP CODE / CITY"]
        missing = [c for c in required_columns if c not in df.columns]
        if missing:
            print(f"⚠️ Sheet '{sheet_name}' is missing columns {missing} - skipping sheet")
            continue

        if "TELEPHONE" not in df.columns:
            print(f"⚠️ Sheet '{sheet_name}' has no 'TELEPHONE' column - adding it")
            ws_check = wb[sheet_name]
            new_col_idx = ws_check.max_column + 1
            ws_check.cell(row=1, column=new_col_idx, value="TELEPHONE")
            df["TELEPHONE"] = ""

        phone_numbers = []

        for index, row in df.iterrows():
            # Guard against empty/NaN name or zip code values, so the search doesn't break
            name = "" if pd.isna(row["SHAREHOLDER"]) else str(row["SHAREHOLDER"]).strip()
            postcode = "" if pd.isna(row["ZIP CODE / CITY"]) else str(row["ZIP CODE / CITY"])[:4]

            if not name:
                print(f"ℹ️ Row {index + 1}: empty name, skipping")
                phone_numbers.append("")
                continue

            print(f"ℹ️ Processing {index + 1}/{len(df)}: {name}")

            # Periodic ad check (every 5 searches)
            if index % 5 == 0 and index > 0:
                print("ℹ️ Periodic ad check...")
                if check_for_blocking_elements(driver):
                    detect_and_close_ad(driver)
                    time.sleep(1)

            phone_number = get_phone_number(name, postcode)
            phone_numbers.append(phone_number)
            print(f"   Found: {phone_number if phone_number else 'nothing'}")

            time.sleep(random.uniform(2, 4))  # random delay to avoid looking like a bot

        # Write the results back into the file
        ws = wb[sheet_name]
        if len(phone_numbers) == len(df):
            telephone_col_idx = df.columns.get_loc("TELEPHONE") + 1
            for i, value in enumerate(phone_numbers, start=2):  # row 1 is headers
                ws.cell(row=i, column=telephone_col_idx, value=value)
        else:
            print(
                f"⚠️ Number of phone numbers found ({len(phone_numbers)}) does not match "
                f"the number of rows ({len(df)}) on sheet '{sheet_name}' - data NOT written"
            )

    # ============================================================
    # 9. SAVE THE FILE
    # ============================================================
    wb.save(file_path)
    wb.close()
    print("✅ File saved with phone numbers")

    # Rename the file, adding " (T)" to the end of the name
    file_dir, file_name = os.path.split(file_path)
    file_base, file_ext = os.path.splitext(file_name)
    new_file_name = f"{file_base} (T){file_ext}"
    new_file_path = os.path.join(file_dir, new_file_name)

    if os.path.exists(new_file_path):
        # If a file with this name already exists (script run again) - don't
        # crash, instead add a timestamp so nothing gets overwritten or fails
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        new_file_name = f"{file_base} (T) {timestamp}{file_ext}"
        new_file_path = os.path.join(file_dir, new_file_name)
        print(f"ℹ️ A file with this name already exists, saving as: {new_file_name}")

    os.rename(file_path, new_file_path)
    print(f"✅ Finished file: {new_file_path}")

except Exception as e:
    print(f"❌ Error while processing phone numbers: {e}")
    raise

finally:
    # Regardless of success or failure - always close the browser,
    # otherwise the Chrome process could stay hanging in memory.
    driver.quit()
    print("ℹ️ ChromeDriver closed")
