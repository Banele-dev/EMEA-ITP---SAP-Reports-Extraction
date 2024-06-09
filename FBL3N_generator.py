from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.common.action_chains import ActionChains
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
log_file_name = f"FBL3N_Log_{datetime.now().strftime('%d%m%Y%H%M')}"+".txt"
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

        #####################################   FBL3N Transaction  #####################################
        ### ---> Open FBL3N Transaction 
        driver.get("https://aopfiori.angloamerican.com/sap/bc/gui/sap/its/webgui#")

        search_bar = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="ToolbarOkCode"]')))
        if not search_bar:
            driver.get("https://aopfiori.angloamerican.com/sap/bc/gui/sap/its/webgui#")

        time.sleep(8)
        search_bar.send_keys('/nFBL3N')
        search_bar.send_keys(Keys.RETURN)
        logging.info(f"FBL3N Transaction open successfully.\n")

        ### ---> Insert G/L Account 
        time.sleep(2) 
        element = WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="M0:46:::1:34"]')))
        element.clear()
        element.send_keys('71000080')
        element.send_keys(Keys.TAB)

        ### ---> Select Company Code 
        time.sleep(5)
        WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="M0:46:::2:78"]'))).click()
        WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="M1:48::btn[16]"]'))).click()
        time.sleep(5)

        #company_code = ['K200']
        company_code = ['X003','X713','X201', 'X200','X005', 'X750', 'X023' , 'X046' , 'X035', 'X604', 'X401', 'X020', 'X608', 'X852', 'X895', 'X892','X894', 'X890', 'X891', 'X899', 'X898', 'X870', 'X889', 'K400', 'X909', 'X893', 'x950', 'X859', 'X851', 'X948', 'X853' , 'X857', 'X951', 'K300', 'R204', 'R202', 'X874', 'R000', 'R355', 'R700','K200']
        company_code_string = "\n".join(company_code)
        pyperclip.copy(company_code_string) 

        try:
            actions = ActionChains(driver)
            actions.key_down(Keys.SHIFT).send_keys(Keys.F12).key_up(Keys.SHIFT).perform()
        except:
            WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="M1:48::btn[24]"]'))).click()

        time.sleep(5)
        control_v = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'urPopupWindowBlockLayer'))).send_keys(Keys.CONTROL, 'v')

        time.sleep(2)
        WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="M1:48::btn[8]"]'))).click()
        logging.info(f"Company Code filtered successfully.")

        ### ---> Select Open Itens
        time.sleep(2)
        WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="M0:46:::11:4"]'))).click()
        logging.info(f"Open Itens selected successfully.")

        ### ---> Select Normal Itens
        time.sleep(2)
        check_box = WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="M0:46:::22:7"]')))
        check_box_stauts = check_box.get_attribute('aria-checked')
        if not check_box_stauts:
            WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="M0:46:::22:7"]'))).click()
        
        logging.info(f"Normal Itens selected successfully.")

        ###---> Choose the right layout 
        time.sleep(5)
        element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="M0:46:::27:34"]')))
        element.clear()
        #element.send_keys('/GRACE')
        element.send_keys('/AUTOMATIONV')
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

        ###-> Check if the page loaded successfully G/L Account
        results_label = WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="userarealist0-content"]')))
        results_label = results_label.text

        if not 'G/L Account' in results_label:
            attempts += 1
            continue

        ###-> Open Spreadsheet 
        actions = ActionChains(driver)
        actions.key_down(Keys.SHIFT).send_keys(Keys.F4).key_up(Keys.SHIFT).perform()

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
        report_name = f'FBL3N_Report_Oficial_{datetime.now().strftime("%d%m%Y%H%M")}'
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
        logging.info(f"FBL3N Report generated successfully.\n")

        ################################## ---> Salve file in sharepoint folder
        report_name = report_name + ".xlsx"
        reports_local_path = f"C:/Users/{user_name}/Downloads/{report_name}"
        sharepoint_path = f"C:/Users/{user_name}/Anglo American/SSBI - GSS Americas - Finance Services - PTP Backlog Management EMEA"
        reports_final_path = f"C:/Users/{user_name}//Anglo American/SSBI - GSS Americas - Finance Services - PTP Backlog Management EMEA/{report_name}"
        history_folder_final_path = f"C:/Users/{user_name}//Anglo American/SSBI - GSS Americas - Finance Services - PTP Backlog Management EMEA/Historic"
        file_to_find = "FBL3N"

        if not os.path.exists(reports_local_path): #it will only execute the tasks below if the file has indeed been generated and saved in Downloads
            logging.error(f"FBL3N Report not found in 'Downloads' folder.\n")
            result = "Failed"
            attempts += 1
            continue

        try:
            df = pd.read_excel(reports_local_path)
        except:
            logging.error(f"Cannot read FBL3N Report in 'Downloads' folder.\n")
            result = "Failed"
            attempts += 1
            continue

        if df.empty:
            logging.error(f"FBL3N Report(s) is empty in 'Downloads' folder.\n")
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
            logging.info(f"FBL3N Report saved in Sharepoint folder successfully.\n")
            result = "Success"

            break 


    except Exception as e:
        with mss.mss() as sct:
            screenshot = sct.shot(output=f"FBL3N_Error_Screenshot_from_attempt_{attempts}.png")
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
        
today_date = datetime.now().strftime("%Y-%m-%d")
today_date_folder = f'{today_date} - Anglo American'
print(result)
