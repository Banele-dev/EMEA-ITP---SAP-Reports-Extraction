from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.edge.options import Options as EdgeOptions
import time 
import pyperclip
from datetime import datetime, timedelta
import os
import logging
import shutil
import traceback
import re
import mss
import pandas as pd
import glob

attempts = 0
user_name = os.getlogin()
start_time = time.time()

################################ LOG PREPARATION ##################################

# Get the path of the directory where the script is located
script_directory = f"C:/Users/{user_name}/Anglo American/GSS Automation Team - General/03_Documentation - Automation Initiatives/02_EMEA/EMEA_ITP_SAP Reports Extraction"

# Create the path for the LogControl folder 
log_control_path = os.path.join(script_directory, 'LogControl')
# If the LogControl folder doesn't exist, create it
if not os.path.exists(log_control_path):
    os.makedirs(log_control_path)

# Create the full path to the log file within the LogControl folder
log_file_name = f"EMEA_Validation_Log_{datetime.now().strftime('%d%m%Y%H%M')}"+".txt"
log_file_path = os.path.join(log_control_path, log_file_name)

logging.basicConfig(
    level=logging.INFO,  
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file_path), 
        logging.StreamHandler()  
    ])

while attempts < 3:
    try:
        # Setting up the webdriver
        edge_path = f"C:/Users/{user_name}/Anglo American/GSS Automation Team - General/03_Documentation - Automation Initiatives/02_EMEA/EMEA_ITP_SAP Reports Extraction/msedgedriver.exe"
        edge_options = webdriver.EdgeOptions()
        edge_options.add_argument(f"executable_path={edge_path}")
        edge_options.add_argument("--start-maximized")
        edge_options.add_argument("--lang=en-US")
        edge_options.add_argument("--disable-notifications")
        edge_options.add_argument('--disable-features')
        edge_options.add_argument('--disable-features=ClipboardEvent')
        edge_options.add_argument("--disable-clipboard-read-protection")
        edge_options.add_argument("--disable-clipboard-write-protection")

        driver = webdriver.Edge(options=edge_options)
        time.sleep(2)

        #####################################   Turning Clipboard Notification to OFF  #####################################
        driver.get("edge://settings/content/clipboard")
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="permission-toggle-row"]'))).click()

        #####################################   EMEA_Validation Transaction  #####################################
        ################################## ---> Open VIM Analytics  Transaction 
        driver.get("https://aopfiori.angloamerican.com/sap/bc/gui/sap/its/webgui#")

        search_bar = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="ToolbarOkCode"]')))
        if not search_bar:
            driver.get("https://aopfiori.angloamerican.com/sap/bc/gui/sap/its/webgui#")

        time.sleep(8)
        search_bar.send_keys('/n/OPT/VIM_VA2')
        search_bar.send_keys(Keys.RETURN)
        logging.info(f"VIM Analytics Transaction opened successfully.\n")

        ################################## ---> Importing data from PF03_WP report (just generated) and consolidate it
        sharepoint_path = f"C:/Users/{user_name}/Anglo American/SSBI - GSS Americas - Finance Services - PTP Backlog Management EMEA"
        file_pattern = "PF03_WP*"
        matching_files = glob.glob(os.path.join(sharepoint_path, file_pattern))

        if matching_files:
            # Read the first matching file found
            df_pf03_wp_report = pd.read_excel(matching_files[0])

        df_pf03_wp_report['Target Process Key (PF / VIM)'] = df_pf03_wp_report['Target Process Key (PF / VIM)'].astype(str)
        df_pf03_wp_report = df_pf03_wp_report.drop_duplicates(subset='Target Process Key (PF / VIM)')
        df_pf03_wp_report = df_pf03_wp_report[['Target Process Key (PF / VIM)']]
        #df_pf03_wp_report = df_pf03_wp_report[['Target Process Key (PF / VIM)']].head(20)

        df_list = df_pf03_wp_report['Target Process Key (PF / VIM)'].tolist()
        df_string = "\n".join(map(str, df_list))
        pyperclip.copy(df_string)

        logging.info(f"PF03_WP's Report data imported successfully.\n")

        ################################## ---> Insert Document Processing Number
        WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="M0:46:::16:78"]'))).click()
        WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="M1:48::btn[24]"]'))).click()
        time.sleep(5)
        control_v = WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.ID, 'urPopupWindowBlockLayer'))).send_keys(Keys.CONTROL, 'v')

        ###-> Button to go forward
        time.sleep(5)
        WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="M1:48::btn[8]"]'))).click()
        logging.info(f"Document Processing Number inserted successfully.")

        ################################## ---> Change the field “Maximum Number of Records” 
        time.sleep(5)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="M0:46:::114:34"]'))).clear()
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="M0:46:::114:34"]'))).send_keys('99999999')

        logging.info(f"Number of records increased successfully.")

        ################################## ---> Choose the right layout 
        time.sleep(5)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="M0:46:::113:34"]'))).clear()
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="M0:46:::113:34"]'))).send_keys('DASHBOARD')

        logging.info(f"Layout inserted successfully.\n")

        ################################## ---> Generate the report

        ###-> Click on the execution button
        time.sleep(5)
        WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="M0:50::btn[8]"]'))).click()

        ###-> Loop until "loading" pop up goes off
        print( "Loading... Please Wait..")
        while True:
            try:
                # Try to find the element (may throw an exception if it is no longer present)
                loading_icon = WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="ur-loading-box"]')))
            except:
                # If the exception is thrown, it means the element is no longer present
                break

        ###-> Check if the page loaded successfully
        results_label = WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="M0:46"]')))
        results_label = results_label.text

        if not 'Results' in results_label:
            attempts += 1
            continue

        ###-> Choose the right layout
        toolbar = WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.XPATH, '//*[contains(@id, "toolbar")]'))).get_attribute("id")
        id_number = re.findall(r'\d{3}', toolbar)[0]
        xpath_three_dots = f'//*[@id="C{id_number}_toolbar-hiddenOpener"]'
        xpath_select_layout_btm = f'//*[@id="_MB_VARIANT{id_number}-BtnChoiceMenu-img"]'

        ###-> Click in the "Export" icon
        time.sleep(2)
        xpath_select_export_btm = f'//*[@id="_MB_EXPORT{id_number}-BtnChoiceMenu-img"]'

        WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, xpath_three_dots))).click() #command needed so automation can run in VDI with no user connected
        WebDriverWait(driver, 100).until(EC.visibility_of_element_located((By.XPATH, xpath_select_export_btm))).click()

        ###-> Select in "Spreadsheet" option
        time.sleep(2)
        WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, '//td[@class="urMnuTxt"]/span[text()="Spreadsheet"]'))).click()
        time.sleep(5)

        ###-> Click Continue and wait for the next step
        WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="M1:50::btn[0]"]'))).click()
        print( "Loading... Please Wait..")
        while True:
            try:
                # Try to find the element (may throw an exception if it is no longer present)
                element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="ur-loading-box"]')))
            except:
                # If the exception is thrown, it means the element is no longer present
                break

        ###-> Renaming file
        report_name = f'EMEA_Validation_Report_Oficial_{datetime.now().strftime("%d%m%Y%H%M")}'
        time.sleep(3)
        WebDriverWait(driver, 40).until(EC.presence_of_element_located((By.XPATH, '//*[@id="popupDialogInputField"]'))).clear()
        time.sleep(1)
        WebDriverWait(driver, 40).until(EC.presence_of_element_located((By.XPATH, '//*[@id="popupDialogInputField"]'))).send_keys(report_name)

        ###-> Click in OK button and wait to load
        WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="UpDownDialogChoose"]'))).click()
        print( "Loading... Please Wait..")
        while True:
            try:
                # Try to find the element (may throw an exception if it is no longer present)
                element = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="ur-loading-box"]')))
            except:
                # If the exception is thrown, it means the element is no longer present
                break

        time.sleep(15)
        logging.info(f"EMEA_Validation Report generated successfully.\n")

        ################################## ---> Salve file in sharepoint folder
        report_name = report_name + ".xlsx"
        reports_local_path = f"C:/Users/{user_name}/Downloads/{report_name}"
        sharepoint_path = f"C:/Users/{user_name}/Anglo American/SSBI - GSS Americas - Finance Services - PTP Backlog Management EMEA"
        reports_final_path = f"C:/Users/{user_name}//Anglo American/SSBI - GSS Americas - Finance Services - PTP Backlog Management EMEA/{report_name}"
        history_folder_final_path = f"C:/Users/{user_name}//Anglo American/SSBI - GSS Americas - Finance Services - PTP Backlog Management EMEA/Historic"
        file_to_find = "EMEA_Validation"


        if not os.path.exists(reports_local_path): #it will only execute the tasks below if the file has indeed been generated and saved in Downloads
            logging.error(f"EMEA_Validation Report not found in 'Downloads' folder.\n")
            result = "Failed"
            attempts += 1
            continue

        try:
            df = pd.read_excel(reports_local_path)
        except:
            logging.error(f"Cannot read EMEA_Validation Report in 'Downloads' folder.\n")
            result = "Failed"
            attempts += 1
            continue

        if df.empty:
            logging.error(f"EMEA_Validation Report(s) is empty in 'Downloads' folder.\n")
            result = "Failed"
            attempts += 1
            shutil.move(reports_local_path, history_folder_final_path) #move empty report to "Histórico"
            continue

        else:
        ###-> Move old file from root folder to historic folder 
            for filename in os.listdir(sharepoint_path):
                if file_to_find in filename:
                    file_path = os.path.join(sharepoint_path, filename)
                    shutil.move(file_path, os.path.join(history_folder_final_path, filename))

        ###-> Move new file from local folder to sharepoint
            shutil.move(reports_local_path, reports_final_path)
            logging.info(f"EMEA_Validation Report saved in Sharepoint folder successfully.\n")
            result = "Success"
            break 

    except Exception as e:
        with mss.mss() as sct:
            screenshot = sct.shot(output=f"EMEA_Validation_Error_Screenshot_from_attempt_{attempts}.png")
        attempts += 1
        trace = traceback.format_exc()
        logging.error(f"Error during automation's execution, automation is designed to try 3 attempts. Here it follows the error:\n{trace}\n\n ---------------------------------- TRYING AGAIN --------------------------------------")
        result = "Failed"



end_time = time.time()
execution_duration = round(end_time - start_time, 2)
logging.info("Automation runtime: " + str((execution_duration)) + ' seconds\n\n')

for handler in logging.getLogger().handlers:
    if isinstance(handler, logging.FileHandler):
        handler.close()
        

print(result)