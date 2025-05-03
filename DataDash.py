# Version ONE, no need for cgm csv

# CMD (-m) pip install selenium
# Needs Chromedriver.exe for your version of Chrome, 32 or 62 windows
# CMD (-m) pip install pywin32
# CMD (-m) pip install auto-py-to-exe

import win32com.client
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException, ElementClickInterceptedException
import glob
import os


def get_user_input(prompt):
    return input(prompt + ": ")

def find_chromedriver():
    # List of common directories to search for chromedriver
    search_dirs = [
        os.path.expanduser("~"),  # User's home directory
        os.path.join(os.path.expanduser("~"), "Downloads"),
        os.path.join(os.path.expanduser("~"), "Desktop"),
        "C:\\Program Files",
        "C:\\Program Files (x86)",
        "D:\\chromedriver-win64",  # Example directory on USB
        "/usr/local/bin",
        "/usr/bin",
    ]

    # Loop through each directory and search for chromedriver.exe
    for search_dir in search_dirs:
        for root, dirs, files in os.walk(search_dir):
            for file in files:
                if file.lower() == "chromedriver.exe":
                    return os.path.join(root, file)
    
    return None

def login(username, password):
    driver_path = find_chromedriver()
    if driver_path:
        print(f"Chromedriver found at: {driver_path}")
    else:
        print("Chromedriver not found. Please ensure it is installed and accessible.")
        
    service = Service(executable_path=driver_path)
    driver = webdriver.Chrome(service=service)

    driver.get('https://clarity.dexcom.com/professional/')

    username_field = driver.find_element(By.NAME, 'username')
    password_field = driver.find_element(By.NAME, 'password')

    username_field.send_keys(username)
    password_field.send_keys(password)
    password_field.send_keys(Keys.RETURN)

    WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.ID, 'ember21')))  # Wait for search field after login

    return driver

def add_patient(driver, study_id, last_name, month_of_birth, day_of_birth, year_of_birth):
    try:
        # Find and click the "Add new patient" button
        new_patient_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, 'ember23')))
        new_patient_button.click()
        print("Clicked on 'add new patient' button")

        # Locate and fill the patient information fields
        first_name_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'patient-form-firstname')))
        last_name_field = driver.find_element(By.ID, 'patient-form-lastname')
        month_of_birth_dropdown = driver.find_element(By.CLASS_NAME, 'localized-inline-date-picker-form--month')
        day_of_birth_field = driver.find_element(By.ID, 'ember75')
        year_of_birth_field = driver.find_element(By.ID, 'ember76')
        profile_id_field = driver.find_element(By.ID, 'patient-form-patient-id')

        first_name_field.send_keys(study_id)
        last_name_field.send_keys(last_name)

        select_month = Select(month_of_birth_dropdown)
        select_month.select_by_visible_text(month_of_birth)

        day_of_birth_field.send_keys(day_of_birth)
        year_of_birth_field.send_keys(year_of_birth)
        profile_id_field.send_keys(study_id)

        # Wait for the "Save" button to be clickable and click it using JavaScript
        save_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, '.btn[type="submit"]')))
        driver.execute_script("arguments[0].click();", save_button)
        print("Clicked on 'Save' button")

        # Return to main dashboard
        driver.get('https://clarity.dexcom.com/professional/patients')

        # Search for the patient again
        search_patient_id_again = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'ember8')))
        search_patient_id_again.send_keys(study_id)
        search_patient_id_again.send_keys(Keys.RETURN)
        print("Searched for patient ID again")

        # Click on the patient name containing 'AI-READI, {study_id}'
        patient_xpath = f"//td[contains(@class, 'patient-list__patient-name') and contains(., 'AI-READI, {study_id}')]"
        patient_name_cell = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, patient_xpath)))
        patient_name_cell.click()
        print(f"Clicked on patient 'AI-READI, {study_id}' in search results")

        # Click upload (first upload button)
        upload_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, '.btn.btn-primary.patient-upload-data__button.ember-view')))
        driver.execute_script("arguments[0].click();", upload_button)
        print("Clicked on 'Upload' button")

        # Upload from reader (second upload button)
        reader_upload_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.mdc-button.mdc-button--raised.uploadDevice.mdc-ripple-upgraded')))
        driver.execute_script("arguments[0].click();", reader_upload_button)
        print("Clicked on 'Upload' button for reader")

        # Continue button
        continue_button = WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.mdc-button.mdc-button--raised.exitUploadAfterSuccess.mdc-ripple-upgraded')))
        driver.execute_script("arguments[0].click();", continue_button)
        print("Clicked on 'Continue' button")

        # Return to main dashboard again
        driver.get('https://clarity.dexcom.com/professional/patients')

        # Search for the patient again
        search_patient_id_again = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'ember8')))
        search_patient_id_again.send_keys(study_id)
        search_patient_id_again.send_keys(Keys.RETURN)
        print("Searched for patient ID again")

        # Click on the patient name containing 'AI-READI, {study_id}'
        patient_xpath = f"//td[contains(@class, 'patient-list__patient-name') and contains(., 'AI-READI, {study_id}')]"
        patient_name_cell = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, patient_xpath)))
        patient_name_cell.click()
        print(f"Clicked on patient 'AI-READI, {study_id}' in search results")

        # Click "Save or print report"
        to_pdf_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, '.btn btn-primary patient-generate-report__button')))
        driver.execute_script("arguments[0].click();", to_pdf_button)
        print("Clicked on 'Upload' button")

        # Click "Save as PDF"
        the_pdf_button = WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.CSS_SELECTOR, '.btn submit-button')))
        driver.execute_script("arguments[0].click();", the_pdf_button)
        print("Clicked on 'Upload' button")

    except Exception as e:
        print(f"An error occurred: {e}")

def download_report(driver, study_id):
    try:
        # Return to main dashboard
        driver.get('https://clarity.dexcom.com/professional/patients')

        # Search for the patient
        search_patient_id = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'ember8')))
        search_patient_id.clear()  # Clear the search field
        search_patient_id.send_keys(study_id)
        search_patient_id.send_keys(Keys.RETURN)
        print("Searched for patient ID")

        # Click on the patient name containing 'AI-READI, {study_id}'
        patient_xpath = f"//td[contains(@class, 'patient-list__patient-name') and contains(., 'AI-READI, {study_id}')]"
        patient_name_cell = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, patient_xpath)))
        print("Found the patient element")

        # Scroll the element into view
        driver.execute_script("arguments[0].scrollIntoView(true);", patient_name_cell)

        # Attempt to click the element
        try:
            patient_name_cell.click()
            print(f"Clicked on patient 'AI-READI, {study_id}' in search results")
        except ElementClickInterceptedException:
            print(f"ElementClickIntercepted: Another element is blocking the click for XPath: {patient_xpath}")
            driver.execute_script("arguments[0].click();", patient_name_cell)
            print(f"Clicked on patient 'AI-READI, {study_id}' in search results using JavaScript")

        # Click "Save or print report"
        to_pdf_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, '.btn.btn-primary.patient-generate-report__button')))
        to_pdf_button.click()  # Click the button directly

        # Click "Save as PDF"
        the_pdf_button = WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.CSS_SELECTOR, '.btn.submit-button')))
        the_pdf_button.click()  # Click the button directly

        # Click "Close" button
        close_button = WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "//button[@class='btn' and text()='Close']")))
        close_button.click()  # Click the button directly

        # Search for the patient again
        search_patient_id_again = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, 'ember8')))
        search_patient_id_again.clear()  # Clear the search field
        search_patient_id_again.send_keys(study_id)
        search_patient_id_again.send_keys(Keys.RETURN)
        print("Searched for patient ID again")

        # Click on the patient name containing 'AI-READI, {study_id}'
        patient_name_cell = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, patient_xpath)))
        print("Found the patient element")

        # Scroll the element into view
        driver.execute_script("arguments[0].scrollIntoView(true);", patient_name_cell)

        # Attempt to click the element
        try:
            patient_name_cell.click()
            print(f"Clicked on patient 'AI-READI, {study_id}' in search results")
        except ElementClickInterceptedException:
            print(f"ElementClickIntercepted: Another element is blocking the click for XPath: {patient_xpath}")
            driver.execute_script("arguments[0].click();", patient_name_cell)
            print(f"Clicked on patient 'AI-READI, {study_id}' in search results using JavaScript")
        
        # Print the page source for debugging
        page_source = driver.page_source
        print(page_source)

        # Find the "Export" button and click it
        export_button = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, "//a[@class='ember-view']/span[text()='Export']")))
        export_button.click()
        print("Clicked on 'Export' button")

        # Click "Export" button
        yes_export_button = WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "//button[@class='btn btn-primary' and text()='Export']")))
        driver.execute_script("arguments[0].click();", yes_export_button)
        print("Clicked on 'Close' button")

    except Exception as e:
        print(f"An error occurred: {e}")

import shutil
import os

def garmin_copy_folder(source_folder_one, destination_folder_one):
    # Check if the source folder exists
    if not os.path.exists(source_folder_one):
        print(f"The source folder '{source_folder_one}' does not exist.")
        return

    # Check if the destination folder exists
    if not os.path.exists(destination_folder_one):
        # Create the destination folder if it does not exist
        os.makedirs(destination_folder_one)

    # Copy the entire folder
    try:
        shutil.copytree(source_folder_one, os.path.join(destination_folder_one, os.path.basename(source_folder_one)))
        print(f"Folder copied successfully from '{source_folder_one}' to '{destination_folder_one}'.")
    except Exception as e:
        print(f"Error: {e}")

def env_copy_files(source_folder_two, destination_folder_two):
    # Check if the source folder exists
    if not os.path.exists(source_folder_two):
        print(f"The source folder '{source_folder_two}' does not exist.")
        return

    # Check if the destination folder exists
    if not os.path.exists(destination_folder_two):
        # Create the destination folder if it does not exist
        os.makedirs(destination_folder_two)

    # Copy files from the source folder to the destination folder
    try:
        for item in os.listdir(source_folder_two):
            source_item_path = os.path.join(source_folder_two, item)
            if os.path.isfile(source_item_path):  # Only copy files
                shutil.copy2(source_item_path, destination_folder_two)
        print(f"Files copied successfully from '{source_folder_two}' to '{destination_folder_two}'.")
    except Exception as e:
        print(f"Error: {e}")

def delete_folder_contents(folder):
    # Check if the folder exists
    if not os.path.exists(folder):
        print(f"The folder '{folder}' does not exist.")
        return

    # Iterate over all the files and folders in the folder and delete them
    for filename in os.listdir(folder):
        file_path = os.path.join(folder, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        except Exception as e:
            print(f"Failed to delete {file_path}. Reason: {e}")

def send_msg():
    outlook = win32com.client.Dispatch("Outlook.Application")
    msg = outlook.CreateItem(0)
    msg.Subject = "AI-READI CGM & Compensation" 
    msg.HTMLBody = f"""
    <p>Hello [NAME],</p>
    <p>Thank you for your participation in our study! </p>
    <p>As a token of our appreciation and as written in the consent, below is a link to a $200 gift card</p>
    <p>$200: 
    <p>Any complications/questions regarding the e-gift card should be directed to the Tengo vendor customer support line @ 877-558-2646. Their hours of operation are Monday to Friday, from 6:00 AM to 5:00PM Pacific Time.</p>
    <p>We are pleased to enclose the results of your 10-day continuous glucose monitoring by Dexcom.  We recommend that you review these results with your primary healthcare provider. We are unable to provide any direct medical interpretation or guidance regarding the results. If you have any difficulty downloading and viewing these results or have additional questions, please contact an AI-READI team member using the information provided below.</p>
    <p>The file is password-protected. Your password is 0101 followed by your four-digit birth year.  For example, if someone's birth year was 2000, then their password would be 01012000.</p> 
    <p>Thank you again for participating in this important study!</p>
    <p>
        \t All the best,<br>
        \t [YOUR NAME]<br>
    </p>
    """
    # Opens a preview of your message! 
    msg.Display()


def get_desktop_path():
    home = os.path.expanduser("~")
    desktop_path = os.path.join(home, "Desktop")

    # Check for OneDrive Desktop path (common on Windows)
    onedrive_path = os.path.join(home, "OneDrive", "Desktop")
    if os.path.exists(onedrive_path):
        return onedrive_path

    # Default to the regular Desktop path
    return desktop_path

def main():
    username = get_user_input("What is your username")
    password = get_user_input("What is your password")
    study_id = get_user_input("What is the study ID")
    env_num = get_user_input("What is the environmental sensor number")
    last_name = "AI-READI"
    month_of_birth = "January"
    day_of_birth = "01"
    year_of_birth = get_user_input("Patient DOB (year)")
    # Prompt user for root directories, not case sensitive
    root_dir_one = input("Enter the root directory (D, G, E, F, etc...) to search for GARMIN: ")
    root_dir_two = input("Enter the root directory (D, G, E, F, etc...) to search for Environmental Sensor: ")

    desktop_path = get_desktop_path()

    study_folder_path = os.path.join(desktop_path, study_id)
    os.makedirs(study_folder_path, exist_ok=True)
    
    subfolder_paths = [
        os.path.join(study_folder_path, f"FIT-{study_id}"),
        os.path.join(study_folder_path, f"ENV-{study_id}-{env_num}")
    ]
    
    for folder_path in subfolder_paths:
        os.makedirs(folder_path, exist_ok=True)
        print(f"Folder created at: {folder_path}")

    driver = login(username, password)

    add_patient(driver, study_id, last_name, month_of_birth, day_of_birth, year_of_birth)
    
    download_report(driver, study_id)

    # DYNAMICALLY SEARCH FOR GARMIN
    source_folder_one = rf"{root_dir_one}:\\GARMIN"  # Replace with the path to "Garmin" folder 
    destination_folder_one = os.path.join(desktop_path, f"{study_id}", f"FIT-{study_id}")  # Replace with your desired destination folder on the desktop

    # DYNAMICALLY SEARCH FOR ENV S
    source_folder_two = rf"{root_dir_two}:\\" # Replace with the path to Environmental Sensor
    destination_folder_two = os.path.join(desktop_path, f"{study_id}", f"ENV-{study_id}-{env_num}")

    garmin_copy_folder(source_folder_one, destination_folder_one)
    env_copy_files(source_folder_two, destination_folder_two)

    # Delete the contents of the specific inner folders, DYNAMICALLY 
    delete_folder_contents(rf"{root_dir_one}:\\GARMIN\\Monitor") # Replace with respective folder paths
    delete_folder_contents(rf"{root_dir_one}:\\GARMIN\\Sleep") # ^^
    delete_folder_contents(rf"{root_dir_two}:\\") # ^^

    send_msg()

    input("Press Enter to close the browser")

if __name__ == "__main__":
    main()
 