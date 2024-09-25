import os
import time

import requests
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

# Figma > settings > API keys/access_token
FIGMA_TOKEN = os.environ.get('FIGMA_TOKEN')

# project needs to be saved to a team space so we can find the File key
TEAM_ID = '1420750927076874147'

FILE_NAME = 'JLL-test'

GOOGLE_SHEET = 'https://docs.google.com/spreadsheets/d/1dQq2mxBSAAc6K2i_8zbNsR9IZgsZmJTFlZ4rk4dgCGM/edit?gid=0#gid=0'

# Also, "Google Sheets Plugin" and "Pitchdeck presentation studio" plugins need to be "Saved plugin" within this workspace on Figma


def login(driver):
    driver.get('https://www.figma.com/login')

    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, 'email')))
    driver.find_element(By.NAME, 'email').send_keys('julesterrien+figma@gmail.com')

    driver.find_element(By.NAME, 'password').send_keys('@Iron123man')

    driver.find_element(By.CSS_SELECTOR, 'button[type="submit"]').click()

    WebDriverWait(driver, 10).until(EC.url_contains('/files'))
    

def get_figma_files():
    response = requests.get(
        f'https://api.figma.com/v1/teams/{TEAM_ID}/projects',
        headers={'X-Figma-Token': FIGMA_TOKEN}
    )
    response.raise_for_status()

    projects = response.json().get('projects', [])
    
    for project in projects:
        project_id = project['id']
        file_response = requests.get(
            f'https://api.figma.com/v1/projects/{project_id}/files',
            headers={'X-Figma-Token': FIGMA_TOKEN}
        )
        file_response.raise_for_status()
        files = file_response.json().get('files', [])

        for file in files:
            if file['name'] == FILE_NAME:
                return file['key']
    
    return None

def open_saved_plugins(driver):
    FIGMA_LOGO = 'toggle-menu-button'
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, FIGMA_LOGO)))
    driver.find_element(By.ID, FIGMA_LOGO).click()

    PLUGIN_MENU = 'mainMenu-plugins-menu-12'
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, PLUGIN_MENU)))
    driver.find_element(By.ID, PLUGIN_MENU).click()

    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//*[text()='Saved plugins']"))
    )
    plugins_element = driver.find_element(By.XPATH, "//*[text()='Saved plugins']")
    plugins_element.click()

def fetch_and_sync_google_sheet(driver):
    open_saved_plugins(driver)

    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//*[text()='Google Sheets Sync']"))
    )
    google_sheets_plugin = driver.find_element(By.XPATH, "//*[text()='Google Sheets Sync']")
    google_sheets_plugin.click()
    
    WebDriverWait(driver, 10).until(
        EC.frame_to_be_available_and_switch_to_it((By.ID, "plugin-iframe-in-modal"))
    )

    WebDriverWait(driver, 10).until(
        EC.frame_to_be_available_and_switch_to_it((By.NAME, "Network Plugin Iframe"))
    )
    
    WebDriverWait(driver, 10).until(
        EC.frame_to_be_available_and_switch_to_it((By.ID, "plugin-iframe"))
    )
    
    input_element = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Enter Sheets shareable link here...']"))
    )
    driver.execute_script("arguments[0].scrollIntoView(true);", input_element)
    input_element.send_keys(GOOGLE_SHEET)
    
    button = driver.find_element(By.XPATH, "//button[text()='Fetch & Sync']")
    button.click()

def export_file(driver):
    open_saved_plugins(driver) 
    
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//*[text()='Pitchdeck Presentation Studio']"))
    )
    google_sheets_plugin = driver.find_element(By.XPATH, "//*[text()='Pitchdeck Presentation Studio']")
    google_sheets_plugin.click()


def open_file_and_run_plugin(driver, file_key):
    file_url = f'https://www.figma.com/design/{file_key}'
    driver.get(file_url)
    
    fetch_and_sync_google_sheet(driver)

    print('syncing from google')
    # wait for google sycn to happen
    time.sleep(10)
    
    print('done sleeping')
    
    export_file(driver)
    
    time.sleep(120)
    

# Step 4: Main Script Execution
def main():
    # Initialize WebDriver (use Chrome or any other browser you prefer)
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service)

    try:
        login(driver)
            
        print('Logged in!')

        # Get the file key from Figma API
        file_key = get_figma_files()
        if file_key:
            open_file_and_run_plugin(driver, file_key)
        else:
            print("File not found!")
    
    finally:
        # Close the driver
        # driver.quit()
        print('Closing')

if __name__ == '__main__':
    main()

