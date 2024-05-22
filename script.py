import os
import subprocess
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException
import time
import pyautogui
import pyperclip
import smtplib
from email.message import EmailMessage
import datetime
from datetime import timezone
import traceback
from pywinauto import Application

# Replace this with the actual path to your Chromedriver executable
chrome_driver_path = r"C:\Users\zulema\Documents\chromedriver-win64\chromedriver.exe"

def open_chrome():
    try:
        # Close any existing Chrome instances
        os.system("taskkill /f /im chrome.exe")
        
        # Wait for a short delay to ensure Chrome is closed
        time.sleep(2)
        
        # Open a new Chrome instance with remote debugging enabled
        subprocess.Popen([r"C:\Program Files\Google\Chrome\Application\chrome.exe", "--remote-debugging-port=9222"])
        
        # Wait for a short delay to ensure Chrome is opened
        time.sleep(5)
        
        print("Chrome opened successfully.")
    except Exception as e:
        print(f"Error opening Chrome: {e}")

# Function to initialize the Chrome WebDriver instance
def init_driver():
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("webdriver.chrome.driver=" + chrome_driver_path)
    chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")  # Use the correct port
    driver = webdriver.Chrome(options=chrome_options)
    driver.get("https://admin.servicefusion.com/jobs")

    # Wait for 7 seconds before proceeding
    time.sleep(7)

    return driver

# Path to the Notepad file
notepad_path = r"C:\Users\zulema\Documents\invoices.txt"

# Paths to the "done.txt" and "ghost.txt" files
done_file_path = r"C:\Users\zulema\Documents\done.txt"
ghost_file_path = r"C:\Users\zulema\Documents\ghost.txt"

# Function to activate the file explorer pop-up
def activate_file_explorer_popup():
    pyautogui.press('tab')
    time.sleep(1)

def close_utility_windows_task():
    try:
        # Open the Chrome task manager
        pyautogui.hotkey('shift', 'esc')
        time.sleep(2)

        # Connect to the Chrome task manager window
        app = Application(backend='uia').connect(title_re='Task Manager')
        task_manager = app.window(title_re='Task Manager')

        # Find the "Utility: Windows Utilities" task
        utility_task = task_manager.child_window(title_re='Utility: Windows Utilities')
        if utility_task.exists():
            utility_task.click_input()
            pyautogui.press('end')
            pyautogui.press('enter')
            print("Closed the 'Utility: Windows Utilities' task.")
            return True
        else:
            print("No 'Utility: Windows Utilities' task found.")
            return False
    except Exception as e:
        print(f"Error closing the 'Utility: Windows Utilities' task: {e}")
        return False

# Function to paste the folder name and perform additional actions using pyautogui
def paste_folder_and_additional_actions(driver, folder_name):
    try:
        folder_path = fr"C:\unprocessed\{folder_name}"
        pyperclip.copy(folder_path)
        activate_file_explorer_popup()
        pyautogui.press('backspace')
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(1)
        pyautogui.press('enter')

        # Add a delay to give time for a potential error message to appear
        time.sleep(1)

        # Rest of your existing code to select all items and perform the upload
        for _ in range(11):
            pyautogui.press('tab')
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('enter')
        pyautogui.press('tab')

        # Initialize a flag for tracking errors
        error_occurred = False

        while True:
            try:
                error_message = WebDriverWait(driver, 2).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, ".jquery-msgbox-wrapper.jquery-msgbox-error"))
                )
                accept_button = error_message.find_element(By.CSS_SELECTOR, ".jquery-msgbox-button-submit")
                accept_button.click()
                # Wait for the page to process the click and load the next popup
                time.sleep(2)
                # Set the error flag to True since an error has occurred
                error_occurred = True
            except TimeoutException:
                # If no error message appears within the timeout, break out of the loop
                break

        # Error handling finished, write the folder name to the ghost file only if an error occurred
        if error_occurred:
            with open(ghost_file_path, "a") as ghost_file:
                ghost_file.write(folder_name + ": File too large\n")

        start_button = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".btn.plupload_start"))
        )
        start_button.click()

        WebDriverWait(driver, 25).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "li.plupload_droptext"))
        )

        # Open the Chrome task manager
        pyautogui.hotkey('shift', 'esc')
        time.sleep(2)

        # Close the 'Utility: Windows Utilities' task if it exists
        utility_task_closed = close_utility_windows_task()

        # If the utility task was closed, restart Chrome and the script
        if utility_task_closed:
            print("Restarting Chrome and the script...")
            driver.quit()
            time.sleep(5)
            open_chrome()
            driver = init_driver()
            return False

        return True

    except Exception as e:
        print(f"Failed to upload folder {folder_name}. Error: {e}")
        with open(ghost_file_path, "a") as ghost_file:
            ghost_file.write(folder_name + "\n")
        # Send an email with the error details
        error_message = f"An error occurred while processing folder {folder_name}:\n{traceback.format_exc()}"
        send_email(error_message)

        # Open the Chrome task manager
        pyautogui.hotkey('shift', 'esc')
        time.sleep(2)

        # Close the 'Utility: Windows Utilities' task if an exception occurs
        close_utility_windows_task()

        # Add a short delay to allow the task to close completely
        time.sleep(2)

        return False

# Function to send email using Gmail SMTP server
def send_email(message):
    sender_email = "wasserstiefell@gmail.com"
    receiver_email = "chavezromeroivan@gmail.com"
    password = "dxwi ofmn raxf dvfq"  # Replace with the app password you generated

    msg = EmailMessage()
    msg['Subject'] = "Script Error" if "An error occurred" in message else "Script Finished"
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg.set_content(message)

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(sender_email, password)
            smtp.send_message(msg)
        print("Email sent successfully")
    except Exception as e:
        print(f"Error sending email: {e}")

def process_invoice(driver, folder_name, line_number):
    try:
        search_bar = driver.find_element("id", "global-search-box")
        search_bar.clear()
        search_bar.send_keys(folder_name)

        wait = WebDriverWait(driver, 25)

        try:
            drop_down_result = wait.until(EC.presence_of_element_located((By.ID, "ui-id-2")))
        except TimeoutException:
            print(f"Failed to find dropdown element for folder {folder_name}")
            with open(ghost_file_path, "a") as ghost_file:
                ghost_file.write(folder_name + "\n")
            # Send an email with the error details
            error_message = f"Failed to find dropdown element for folder {folder_name}:\n{traceback.format_exc()}"
            send_email(error_message)
            return False
        except WebDriverException as e:
            # If Chrome crashes or becomes unresponsive, restart the driver instance
            if "out of memory" in str(e) or "unknown error" in str(e):
                print("Chrome has crashed or become unresponsive. Restarting the driver instance...")
                driver.quit()
                driver = init_driver()
                # Send an email notification
                error_message = f"Chrome crashed or became unresponsive. Restarting the driver instance.\n{traceback.format_exc()}"
                send_email(error_message)
                return False
            else:
                raise e

        # Check for the specific element with the class and text
        is_customers_element_present = False
        try:
            customers_element = driver.find_element(By.XPATH, "//li[contains(@class, 'ui-autocomplete-category') and contains(text(), 'Customers')]")
            if customers_element.is_displayed():
                is_customers_element_present = True
        except:
            pass

        if is_customers_element_present:
            drop_down_result = wait.until(EC.presence_of_element_located((By.ID, "ui-id-3")))

        drop_down_result.click()

        pictures_tab = driver.find_element("id", "pictures-title")
        pictures_tab.click()

        upload_btn = driver.find_element("id", "upload-btn")
        upload_btn.click()

        browse_button = driver.find_element("id", "plupload-upload-image_browse")
        browse_button.click()

        upload_successful = paste_folder_and_additional_actions(driver, folder_name)
        
        if upload_successful:
            with open(done_file_path, "a") as done_file:
                done_file.write(folder_name + "\n")
        else:
            return False
        
        # Add a short delay to allow the task to close completely
        time.sleep(2)
        
        return True
    except Exception as e:
        print(f"Failed to process invoice {folder_name}. Error: {e}")
        # Handle the exception (e.g., logging, sending an email)
        return False

# Main execution starts
with open(notepad_path, "r") as notepad_file:
    total_invoices = sum(1 for _ in notepad_file)

invoices_processed = 0
last_processed_line = 0

open_chrome()
driver = init_driver()

while True:
    with open(notepad_path, "r") as notepad_file:
        for line_number, line in enumerate(notepad_file, start=1):
            if line_number <= last_processed_line:
                continue  # Skip already processed invoices
            
            folder_name = line.strip()
            success = process_invoice(driver, folder_name, line_number)
            
            if success:
                invoices_processed += 1
                last_processed_line = line_number  # Update the last processed line number
                percentage_done = round((invoices_processed / total_invoices) * 100, 2)
                print(f"Processed {invoices_processed}/{total_invoices} invoices ({percentage_done}%)")
            else:
                print(f"Skipping invoice {folder_name} due to an error.")
                break  # Break out of the inner loop to restart the script from the last processed line

    if invoices_processed == total_invoices:
        break  # Break out of the outer loop if all invoices have been processed

# Send email when script finishes
utc_time = datetime.datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")
message = f"The script finished at {utc_time} (UTC)"
send_email(message)

# Close the browser
driver.quit()