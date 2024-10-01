import os
import re
import time
import urllib.parse

import pandas as pd
import requests
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

# ENV VARS
FIGMA_TOKEN = os.environ.get('FIGMA_TOKEN') # Figma > settings > API keys/access_token
USERNAME = 'julesterrien+figma@gmail.com' # user needs username + pw to loginto FIgma 
PASSWORD = '@Iron123man'

# For this script to work:
# - "Google Sheets Plugin" and "Pitchdeck presentation studio" plugins
#   need added to "Saved plugin" for the given project
# - The figma project can't be a 'draft'.
#   Make sure it's saved to a team space so the script can find the file key and access the key
#   The script will read the google sheet provided to find
#   the team ID and file name that are requested to be automated
GOOGLE_SHEET = 'https://docs.google.com/spreadsheets/d/1dQq2mxBSAAc6K2i_8zbNsR9IZgsZmJTFlZ4rk4dgCGM/edit?gid=0#gid=0'
SHEET_NAME = "project metadata"

# TODO: fetch team id  and file name from second sheet of google sheet
# TODO: toggle off compress and downsize to get speed down
# TODO: on export, click top left to toggle to "Sort from Top to Bottom" in Sort Visually rows

def get_csv_url(sheet_url, sheet_name):
    # Extract the document ID using regex
    match = re.search(r'/d/([a-zA-Z0-9-_]+)', sheet_url)
    if not match:
        raise ValueError("Invalid Google Sheet URL provided.")
    
    doc_id = match.group(1)
    
    encoded_sheet_name = urllib.parse.quote(sheet_name)

    csv_url = f"https://docs.google.com/spreadsheets/d/{doc_id}/gviz/tq?tqx=out:csv&sheet={encoded_sheet_name}"
    
    return csv_url

def fetch_file_key():
    print('fetching client config')
    csv_url = get_csv_url(GOOGLE_SHEET, SHEET_NAME)

    # Read the CSV data into a pandas DataFrame
    df = pd.read_csv(csv_url, header=None)

    # Get the first 3 rows from the DataFrame
    first_three_rows = df.head(3)
    
    # Transpose the DataFrame so that the first column becomes keys and the second column becomes values
    client_configs = first_three_rows.set_index(first_three_rows.columns[0])[first_three_rows.columns[1]].to_dict()

    team_id = client_configs.get('FigmaTeamId')
    file_name = client_configs.get('FileName')
    client_email = client_configs.get('ClientEmail')
    
    if not team_id:
        raise ValueError(f"Team ID was not provided in the 'project metadata' sheet of {csv_url}")

    if not file_name:
        raise ValueError(f"File name was not provided in the 'project metadata' sheet of {csv_url}")

    if not client_email:
        raise ValueError(f"Client email was not provided in the 'project metadata' sheet of {csv_url}")
    
    return team_id, file_name, client_email

def login(driver):
    driver.get('https://www.figma.com/login')

    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, 'email')))
    driver.find_element(By.NAME, 'email').send_keys(USERNAME)

    driver.find_element(By.NAME, 'password').send_keys(PASSWORD)

    driver.find_element(By.CSS_SELECTOR, 'button[type="submit"]').click()

    WebDriverWait(driver, 10).until(EC.url_contains('/files'))
    
    print('Logged in!')

def get_figma_files(team_id, file_name):
    response = requests.get(
        f'https://api.figma.com/v1/teams/{team_id}/projects',
        headers={'X-Figma-Token': FIGMA_TOKEN}
    )
    response.raise_for_status()

    projects = response.json().get('projects', [])
    
    for project in projects:
        project_id = project['id']
        files_in_project = requests.get(
            f'https://api.figma.com/v1/projects/{project_id}/files',
            headers={'X-Figma-Token': FIGMA_TOKEN}
        )
        files_in_project.raise_for_status()
        files = files_in_project.json().get('files', [])

        for file in files:
            if file['name'] == file_name:
                return file['key']
    
    return None

def count_frames(data):
    frame_count = 0
    
    def count_frame_recursively(node):
        nonlocal frame_count
        if node.get("type") == "FRAME":
            frame_count += 1
        if "children" in node:
            for child in node["children"]:
                count_frame_recursively(child)
    
    count_frame_recursively(data)
    
    return frame_count

def count_file_frames(file_key):
    response = requests.get(
        f'https://api.figma.com/v1/files/{file_key}',
        headers={'X-Figma-Token': FIGMA_TOKEN}
    )
    response.raise_for_status()

    response = response.json()
    
    document = response['document']
    
    # Call the function and print the result
    number_of_frames = count_frames(document)
    return number_of_frames

def go_to_plugin_iframe(driver):
    WebDriverWait(driver, 10).until(
        EC.frame_to_be_available_and_switch_to_it((By.ID, "plugin-iframe-in-modal"))
    )

    WebDriverWait(driver, 10).until(
        EC.frame_to_be_available_and_switch_to_it((By.NAME, "Network Plugin Iframe"))
    )
    
    WebDriverWait(driver, 10).until(
        EC.frame_to_be_available_and_switch_to_it((By.ID, "plugin-iframe"))
    )

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
    
    go_to_plugin_iframe(driver)
    
    input_element = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Enter Sheets shareable link here...']"))
    )
    driver.execute_script("arguments[0].scrollIntoView(true);", input_element)
    input_element.send_keys(GOOGLE_SHEET)
    
    button = driver.find_element(By.XPATH, "//button[text()='Fetch & Sync']")
    button.click()
    
    print('syncing file from google')
    time.sleep(10)
    
    # leave iframe context and return to default document context
    driver.switch_to.default_content()

def export_file(driver, frame_count):
    open_saved_plugins(driver) 
    
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//*[text()='Pitchdeck Presentation Studio']"))
    )
    google_sheets_plugin = driver.find_element(By.XPATH, "//*[text()='Pitchdeck Presentation Studio']")
    google_sheets_plugin.click()
    
    # plugin modal can take some time to appear
    print('loading export plugin')
    time.sleep(10)
    
    go_to_plugin_iframe(driver)
    
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//div[@class='slide__orders']//button[@class='select-menu__button']"))
    )
    sort_button = driver.find_element(By.XPATH, "//div[@class='slide__orders']//button[@class='select-menu__button']")
    sort_button.click()
    
    sort_item = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Sort from Top to Bottom')]"))
    )
    sort_item.click()
    
    # export_button = driver.find_element(By.XPATH, "//button[text()='Export']")
    export_button = driver.find_element(By.CSS_SELECTOR, "button.button--primary")
    export_button.click()
    
    button_element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'settings__content')]//button[contains(@class, 'select-menu__button')]"))
    )
    button_text_content = driver.execute_script("return arguments[0].textContent;", button_element)
    export_option_button = driver.find_element(By.XPATH, "//div[contains(@class, 'settings__content')]//button[contains(@class, 'select-menu__button')]")
    
    # it sometimes takes a few ms for this button to become interactable
    time.sleep(2)
    
    if "Google Slides (.pptx file)" not in button_text_content:
        print('selected option is not correct...changing to: Google Slides (.pptx file)')
        export_option_button.click()

        google_slides_item = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Google Slides (.pptx file)')]"))
        )
        google_slides_item.click()

    # disable compress/XL image options for faster export
    compress_images_checkbox = driver.find_element(By.XPATH, "//label[contains(text(), 'Compress Images')]/preceding-sibling::input[@type='checkbox']")
    if compress_images_checkbox.is_selected():
        print("Compress Images is enabled, disabling it...")
        compress_images_toggle = driver.find_element(By.XPATH, "//label[contains(text(), 'Compress Images')]/parent::div")
        compress_images_toggle.click()
    
    # 2x Retina Images setting
    retina_images_checkbox = driver.find_element(By.XPATH, "//label[contains(text(), '2x Retina Images')]/preceding-sibling::input[@type='checkbox']")
    if retina_images_checkbox.is_selected():
        print("2x Retina Images is enabled, disabling it...")
        retina_images_toggle = driver.find_element(By.XPATH, "//label[contains(text(), '2x Retina Images')]/parent::div")
        retina_images_toggle.click()
    
    # Downsize Large Figma Image Fills setting
    downsize_figma_images_checkbox = driver.find_element(By.XPATH, "//label[contains(text(), 'Downsize Large Figma Image Fills')]/preceding-sibling::input[@type='checkbox']")
    if downsize_figma_images_checkbox.is_selected():
        print("Downsize Large Figma Image Fills is enabled, disabling it...")
        downsize_figma_images_toggle = driver.find_element(By.XPATH, "//label[contains(text(), 'Downsize Large Figma Image Fills')]/parent::div")
        downsize_figma_images_toggle.click()

    # export_for_google_slides_button = driver.find_element(By.XPATH, "//button[contains(@class, 'button--primary') and .//span[contains(text(), 'Export for Google Slides')]]")
    # export_for_google_slides_button.click()
    
    # # it takes a moment for the export to be generated
    # print('waiting for export to be generated')
    
    # # this took 3.50s on my macbook for a ~50-70 slide deck so we should beef this up to ~5mins or calculate based on number of sheets/nodes on this workspace
    # # so avg. ~4s per frame
    # # TODO: shorten this via toggle settling
    # wait_time = frame_count * 4
    # print('number of frames to process', frame_count)
    # print('estimated wait time to prepare download', wait_time)
    # time.sleep(wait_time)
    # print('export should be generated by now')
    
    # confirm_export_button = driver.find_element(By.XPATH, "//button[contains(@class, 'button--primary') and .//span[contains(text(), 'Download your .pptx file')]]")
    # confirm_export_button.click()
    # print('file downloaded')

def open_file_and_run_plugin(driver, file_key, frame_count):
    file_url = f'https://www.figma.com/design/{file_key}'
    driver.get(file_url)
    
    fetch_and_sync_google_sheet(driver)

    print('syncing from google')
    
    export_file(driver, frame_count)
    
    time.sleep(120)

def main():
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service)

    try:
        print("""      
        ⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀
        ⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀
        ⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀
        ⠀⠀⠀⠀⣀⣤⣴⣶⣶⣦⣄⡀⠀⠀⠀⠀⠀⠀⢀⣤⣶⣶⣶⣦⣤⡀⠀⠀⠀⠀
        ⠀⠀⢀⣾⣿⣿⣿⣿⣿⣿⣿⣷⣄⠀⠀⠀⢀⣾⣿⣿⣿⣿⣿⣿⣿⣿⣦⡀⠀⠀
        ⠀⠀⣾⣿⠟⠋⠉⠀⠀⠉⠙⠻⣿⣷⡀⣰⣿⣿⣿⠟⠉⠀⠀⠀⠈⠙⣿⣷⠀⠀
        ⠀⢸⣿⠏⠀⠀⠀⠀⠀⠀⠀⠀⠈⢻⣿⣿⣿⡿⠃⠀⠀⠀⠀⠀⠀⠀⠸⣿⡇⠀
        ⠀⢸⣿⠀⠀⠀⠀⠀⠀⠀⠀⠀⢀⣾⣿⣿⡿⠁⠀⠀⠀⠀⠀⠀⠀⠀⠀⣿⡇⠀
        ⠀⢸⣿⡆⠀⠀⠀⠀⠀⠀⠀⢀⣾⣿⣿⣿⣧⡀⠀⠀⠀⠀⠀⠀⠀⠀⢰⣿⡇⠀
        ⠀⠀⢿⣿⣄⡀⠀⠀⠀⢀⣴⣿⣿⣿⠟⠘⢿⣿⣦⣀⡀⠀⠀⢀⣀⣴⣿⡿⠀⠀
        ⠀⠀⠈⠻⣿⣿⣿⣿⣿⣿⣿⣿⡿⠁⠀⠀⠀⠙⣿⣿⣿⣿⣿⣿⣿⣿⡿⠁⠀⠀
        ⠀⠀⠀⠀⠈⠛⠻⠿⠿⠿⠛⠁⠀⠀⠀⠀⠀⠀⠈⠙⠻⠿⠿⠿⠛⠉⠀⠀⠀⠀
        ⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀
        ⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀
        ⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀
              """)
        
        team_id, file_name, client_email = fetch_file_key()
        
        login(driver)

        # Get the file key from Figma API
        file_key = get_figma_files(team_id=team_id, file_name=file_name)
        
        if not file_key:
            raise ValueError(f"Figma file is missing: {file_name}")
        
        frame_count = count_file_frames(file_key)

        open_file_and_run_plugin(driver, file_key, frame_count)
    
    finally:
        print('Closing driver')
        # driver.quit()

if __name__ == '__main__':
    main()
