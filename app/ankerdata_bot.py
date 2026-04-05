import os
import time
import shutil
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

def fetch_reports():
    print("Launching Chrome for Ankerdata...")
    options = webdriver.ChromeOptions()
    
    # Setup persistent profile so login state is saved between runs
    profile_dir = os.path.join(os.getcwd(), 'data', 'profiles', 'AnkerdataProfile')
    options.add_argument(f"user-data-dir={profile_dir}")
    
    # Set download directory so we know exactly where files land
    download_dir = os.path.join(os.getcwd(), 'tmp')
    os.makedirs(download_dir, exist_ok=True)
    prefs = {"download.default_directory" : download_dir}
    options.add_experimental_option("prefs", prefs)
    
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    
    try:
        driver.get("https://enterprise.web.za/")
        
        try:
            # Attempt to login
            WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "input#Email")))
            print("Logging in...")
            driver.find_element(By.CSS_SELECTOR, "input#Email").send_keys("Dylan.lloyd25@gmail.com")
            driver.find_element(By.CSS_SELECTOR, "input#Password").send_keys("aiNO5Bpo")
            driver.find_element(By.CSS_SELECTOR, "button.btn-signin").click()
        except:
            print("Already logged in or login fields not found, proceeding...")
            
        import traceback

        # Navigate to Cashiers/Servers
        print("Navigating to Cashiers/Servers...")
        time.sleep(3)
        cashiers_elem = WebDriverWait(driver, 25).until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Cashiers/Servers')]")))
        driver.execute_script("arguments[0].scrollIntoView();", cashiers_elem)
        time.sleep(1)
        try:
            cashiers_elem.click()
        except Exception:
            driver.execute_script("arguments[0].click();", cashiers_elem)
        
        # Navigate to Clock In/Out
        time.sleep(2)
        print("Navigating to Clock In/Out...")
        clockin_elem = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, "//a[contains(@href, 'ClockInOut')]")))
        driver.execute_script("arguments[0].scrollIntoView();", clockin_elem)
        time.sleep(1)
        try:
            clockin_elem.click()
        except:
            driver.execute_script("arguments[0].click();", clockin_elem)
        
        print("Waiting for Report to load...")
        time.sleep(10)
        
        # Open Export Menu Toggle
        print("Opening export menu...")
        try:
            # Try a more resilient XPath for the sidebar toggle
            menu_toggle = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//a[contains(@class, 'sidebar-toggle') and contains(@class, 'right')]")))
            driver.execute_script("arguments[0].scrollIntoView();", menu_toggle)
            time.sleep(1)
            menu_toggle.click()
        except Exception as e:
            print(f"Standard toggle failed, trying JS click: {e}")
            driver.execute_script("document.querySelector('.right.sidebar-toggle').click();")
        
        time.sleep(2)
        print("Downloading to Excel...")
        try:
            export_btn = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a[title='Export to Excel']")))
            export_btn.click()
        except:
            driver.execute_script("document.querySelector('a[title=\"Export to Excel\"]').click();")
        
        # Wait for download to actually start/finish
        print("Waiting for download...")
        time.sleep(15)
        
        # Move the newest downloaded file to 'latest_clockin.xlsx'
        # Check both .xlsx and .csv
        download_dir = os.path.join(os.getcwd(), 'tmp')
        files = [f for f in os.listdir(download_dir) if (f.endswith('.xlsx') or f.endswith('.csv')) and 'Template' not in f and 'EFT' not in f and 'latest' not in f]
        files.sort(key=lambda x: os.path.getmtime(os.path.join(download_dir, x)), reverse=True)
        
        if files:
            newest_file = os.path.join(download_dir, files[0])
            dest_file = os.path.join(os.getcwd(), 'data', 'input', 'latest_clockin.xlsx')
            print(f"Found recently downloaded file: {newest_file}")
            
            if newest_file.endswith('.csv'):
                print("Converting CSV to XLSX for processor...")
                df_temp = pd.read_csv(newest_file)
                df_temp.to_excel(dest_file, index=False)
            else:
                shutil.copy(newest_file, dest_file)
                
            print(f"Verified: {newest_file} processed as {dest_file}")
        else:
            print("Could not find any recently downloaded files in the directory.")

    except Exception as e:
        print(f"Automation failed: {e}")
        # Save a screenshot for debugging if it fails
        os.makedirs("output", exist_ok=True)
        driver.save_screenshot("output/ankerdata_fail.png")
        print("Screenshot saved to output/ankerdata_fail.png")
    finally:
        driver.quit()
    print("Session over.")

if __name__ == '__main__':
    import pandas as pd
    fetch_reports()
