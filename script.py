from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import time
import pyautogui
import pyperclip

# Set the path to the Chrome WebDriver executable on your system
chrome_driver_path = r"C:\Users\wasse\Documents\mattucci\chromedr\chromedriver.exe"

# Configure the Chrome WebDriver options
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("webdriver.chrome.driver=" + chrome_driver_path)
chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")  # Connect to an existing Chrome session

# Initialize the Chrome WebDriver instance
driver = webdriver.Chrome(options=chrome_options)

# Navigate to the target website
driver.get("https://admin.servicefusion.com/jobs")

# Specify the paths for the files involved in the process
notepad_path = r"C:\Users\wasse\Documents\mattucci\invoices.txt"  # File containing folder names
done_file_path = r"C:\Users\wasse\Documents\mattucci\done.txt"  # File to track successfully uploaded folders
ghost_file_path = r"C:\Users\wasse\Documents\mattucci\ghost.txt"  # File to track failed uploads

# Helper function to activate the file explorer pop-up
def activate_file_explorer_popup():
    pyautogui.press('tab')  # Press the 'tab' key to switch focus to the file explorer
    time.sleep(1)  # Wait for a moment to ensure the focus has shifted

# Helper function to paste the folder path and perform additional actions
def paste_folder_and_additional_actions(folder_name):
    try:
        folder_path = fr"C:\Users\wasse\Documents\mattucci\pictures\{folder_name}"  # Construct the folder path
        pyperclip.copy(folder_path)  # Copy the folder path to the clipboard
        activate_file_explorer_popup()  # Activate the file explorer pop-up
        pyautogui.press('backspace')  # Clear any pre-existing text in the file path field
        pyautogui.hotkey('ctrl', 'v')  # Paste the folder path from the clipboard
        time.sleep(1)  # Wait for a moment to ensure the folder path is pasted
        pyautogui.press('enter')  # Press 'Enter' to confirm the folder selection

        # Add a delay to give time for a potential error message to appear
        time.sleep(1)

        # Select all files in the folder and initiate the upload process
        for _ in range(11):
            pyautogui.press('tab')  # Press 'tab' multiple times to navigate through the elements
        pyautogui.hotkey('ctrl', 'a')  # Select all files in the folder
        pyautogui.press('enter')  # Confirm the file selection
        pyautogui.press('tab')  # Switch focus to the next element

        # Check for potential error messages during the upload process
        error_occurred = False
        while True:
            try:
                # Wait for an error message element to appear (if any)
                error_message = WebDriverWait(driver, 2).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, ".jquery-msgbox-wrapper.jquery-msgbox-error"))
                )
                accept_button = error_message.find_element(By.CSS_SELECTOR, ".jquery-msgbox-button-submit")
                accept_button.click()  # Click the 'Accept' button to dismiss the error message
                time.sleep(2)  # Wait for the page to process the click and load the next popup (if any)
                error_occurred = True  # Set the flag to indicate that an error occurred
            except TimeoutException:
                # If no error message appears within the timeout, break out of the loop
                break

        # If an error occurred during the upload, log the folder name to the 'ghost.txt' file
        if error_occurred:
            with open(ghost_file_path, "a") as ghost_file:
                ghost_file.write(folder_name + ": File too large\n")

        # Start the actual upload process
        start_button = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".btn.plupload_start"))
        )
        start_button.click()

        # Wait for the upload progress indicator to appear
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "li.plupload_droptext"))
        )

        # Log the successfully uploaded folder name to the 'done.txt' file
        with open(done_file_path, "a") as done_file:
            done_file.write(folder_name + "\n")

    except Exception as e:
        print(f"Failed to upload folder {folder_name}. Error: {e}")
        # If an exception occurs, log the folder name to the 'ghost.txt' file
        with open(ghost_file_path, "a") as ghost_file:
            ghost_file.write(folder_name + "\n")

# Main execution starts here
with open(notepad_path, "r") as notepad_file:
    for line_number, line in enumerate(notepad_file, start=1):
        folder_name = line.strip()  # Get the folder name from the current line
        search_bar = driver.find_element("id", "global-search-box")
        search_bar.clear()  # Clear the search bar
        search_bar.send_keys(folder_name)  # Enter the folder name in the search bar

        wait = WebDriverWait(driver, 10)
        drop_down_result = wait.until(EC.presence_of_element_located((By.ID, "ui-id-2")))

        # Check if the 'Customers' element is present in the drop-down
        is_customers_element_present = False
        try:
            customers_element = driver.find_element(By.XPATH, "//li[contains(@class, 'ui-autocomplete-category') and contains(text(), 'Customers')]")
            if customers_element.is_displayed():
                is_customers_element_present = True
        except:
            pass

        # If the 'Customers' element is present, select a different drop-down option
        if is_customers_element_present:
            drop_down_result = wait.until(EC.presence_of_element_located((By.ID, "ui-id-3")))

        drop_down_result.click()  # Click the selected drop-down result

        pictures_tab = driver.find_element("id", "pictures-title")
        pictures_tab.click()  # Navigate to the 'Pictures' tab

        upload_btn = driver.find_element("id", "upload-btn")
        upload_btn.click()  # Click the 'Upload' button

        browse_button = driver.find_element("id", "plupload-upload-image_browse")
        browse_button.click()  # Click the 'Browse' button to open the file explorer

        paste_folder_and_additional_actions(folder_name)  # Paste the folder path and initiate the upload process

# Do not close the browser automatically
# driver.quit()