import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.edge.options import Options as EdgeOptions
import time 
import pyperclip
import shutil
from datetime import datetime, timedelta
import os
import logging
import traceback
import glob
import re
import mss

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
log_file_name = f"VIM_ILC_Log_{datetime.now().strftime('%d%m%Y%H%M')}"+".txt"
log_file_path = os.path.join(log_control_path, log_file_name)

logging.basicConfig(
    level=logging.INFO,  
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file_path), 
        logging.StreamHandler()  
    ])


###--> Importing data from VIM_VA2 report (just generated)
sharepoint_path = f"C:/Users/{user_name}/Anglo American/SSBI - GSS Americas - Finance Services - PTP Backlog Management EMEA"
file_pattern = "VIM_VA2*"
matching_files = glob.glob(os.path.join(sharepoint_path, file_pattern))

if matching_files:
    # Read the first matching file found
    df_va2_report = pd.read_excel(matching_files[0])

df_va2_report['Document Id'] = df_va2_report['Document Id'].astype(str)
df_va2_report = df_va2_report[['Document Id']]
#df_va2_report = df_va2_report[['Document Id']].head(20)
logging.info(f"VIM_VA2 data imported Successfully.\n")

###--> Spliting data into batches of 5000 rows (max)
number_of_lines = 5000
number_batches = len(df_va2_report) // number_of_lines + (1 if len(df_va2_report) % number_of_lines != 0 else 0)
splitted_dataframes = [df_va2_report.iloc[i * number_of_lines:(i + 1) * number_of_lines] for i in range(number_batches)]
logging.info(f"Number of lines in VA2:" + str((len(df_va2_report) ))+ "")
logging.info(f"Number of batches (reports) that will be generated:" + str((number_batches)) + " \n")

### ---> Generating reports in batches 
batches = 0
attempts = 0
while batches < number_batches:
    try:

        ### ---> Setting up the webdriver
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

        #####################################   Turning Clipboard Notification OFF  #####################################
        time.sleep(2)
        driver.get("edge://settings/content/clipboard")
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="permission-toggle-row"]'))).click()

        ##########################################   VIM_ILC Transaction  #####################################
        driver.get("https://aopfiori.angloamerican.com/sap/bc/gui/sap/its/webgui#")
        search_bar = WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.XPATH, '//*[@id="ToolbarOkCode"]')))
        if not search_bar:
            driver.get("https://aopfiori.angloamerican.com/sap/bc/gui/sap/its/webgui#")

        # Automation will try to execute 3 times the same "attempt"
        if attempts >= 4:
            logging.error(f"Number of attemps exceed. Automation tried " + str((attempts)) + " attempts but could not generated all the " + str((number_batches)) + " batches.\n")
            result = "Failed"
            break

        current_dataframe = splitted_dataframes[batches]
        column_to_copy = current_dataframe['Document Id']
        df_list = column_to_copy.tolist()
        df_string = "\n".join(map(str, df_list))
        pyperclip.copy(df_string)

        ### ---> Open VIM_ILC2 Transaction 
        time.sleep(5)
        search_bar = WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.XPATH, '//*[@id="ToolbarOkCode"]')))
        search_bar.send_keys('/n/OPT/VIM_ILC')
        search_bar.send_keys(Keys.RETURN)

        ### ---> Open multiple selection section -  ID do document field 
        WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="M0:46:::1:78"]'))).click()

        ### ---> Saving the IDs in clipboard
        upload_from_clipboard_button = WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="M1:48::btn[24]"]'))).click()
        time.sleep(3)
        control_v = WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.ID, 'urPopupWindowBlockLayer'))).send_keys(Keys.CONTROL, 'v')

        ###-> Button to go forward
        time.sleep(5)
        WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="M1:48::btn[8]"]'))).click()

        ###---> Choose the right layout 
        time.sleep(5)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="M0:46:::15:34"]'))).clear()
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="M0:46:::15:34"]'))).send_keys('DASHBOARD')

        logging.info(f"Layout inserted successfully.\n")


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

        ###-> Check if the page loaded Successfully 
        label = WebDriverWait(driver, 40).until(EC.visibility_of_element_located((By.CSS_SELECTOR, '.urST3Tit.urST4LbTitBg'))).get_attribute('id')
        id_number = re.findall(r'\d{3}', label)[0]
        xpath_label = f'//*[@id="C{id_number}-title"]'

        invoice_lifecycle_report_label = WebDriverWait(driver, 40).until(EC.visibility_of_element_located((By.XPATH, xpath_label))).text
        if not 'Invoice Lifecycle Report' in invoice_lifecycle_report_label:
            batches += 1
            result = "Failed"
            continue


        ###-> Select in "Spreadsheet" option
        time.sleep(2)
        xpath_export_buttom = f'//*[@id="_MB_EXPORT{id_number}"]'
        try:
            WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, xpath_export_buttom))).click()
        except:
            WebDriverWait(driver, 100).until(EC.visibility_of_element_located((By.XPATH, '//*[contains(@id, "_MB_EXPORT")]'))).click()
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
        deployment_date = datetime.now().strftime("%d%m%Y%H%M")
        report_name = "VIM_ILC_Report_" + "batch_" + str(batches) + "_" + str(deployment_date)
        time.sleep(3)
        WebDriverWait(driver, 40).until(EC.presence_of_element_located((By.XPATH, '//*[@id="popupDialogInputField"]'))).clear()
        time.sleep(1)
        WebDriverWait(driver, 40).until(EC.presence_of_element_located((By.XPATH, '//*[@id="popupDialogInputField"]'))).send_keys(report_name)

        ###-> Click in OK button and wait to load (command is duplicated due to SAP condition "it pops up 2 loading windows sometimes")
        WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="UpDownDialogChoose"]'))).click()
        print( "Loading... Please Wait..")
        while True:
            try:
                # Try to find the element (may throw an exception if it is no longer present)
                element = WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="ur-loading-box"]')))
            except:
                # If the exception is thrown, it means the element is no longer present
                break

        print( "Loading... Please Wait..")
        while True:
            try:
                # Try to find the element (may throw an exception if it is no longer present)
                element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="ur-loading-box"]')))
            except:
                # If the exception is thrown, it means the element is no longer present
                break

        time.sleep(60)
        result = "Success"
        logging.info(f"Batch number " + str((batches+1)) + " generated successfully.\n")
        batches += 1
        driver.quit()

    except Exception as e:
        with mss.mss() as sct:
            screenshot = sct.shot(output=f"ILC_Error_Screenshot_from_attempt_{attempts}.png")
        attempts += 1
        trace = traceback.format_exc()
        logging.error(f"Error during automation's execution, automation is designed to try 4 attempts. Here it follows the error:\n{trace}\n\n ---------------------------------- TRYING AGAIN --------------------------------------")
        result = "Failed"
        driver.quit()




################################## Salve file in sharepoint folder ##################################
reports_local_path = f"C:/Users/{user_name}/Downloads"
sharepoint_path = f"C:/Users/{user_name}/Anglo American/SSBI - GSS Americas - Finance Services - PTP Backlog Management EMEA"
history_folder_final_path = f"C:/Users/{user_name}/Anglo American/SSBI - GSS Americas - Finance Services - PTP Backlog Management EMEA/Historic"
file_to_find = "VIM_ILC"

if result == "Success":
    ###-> Move old files from root folder to historic folder 
    for filename in os.listdir(sharepoint_path):
        if file_to_find in filename:
            file_path = os.path.join(sharepoint_path, filename)
            shutil.move(file_path, os.path.join(history_folder_final_path, filename))

    ###-> Move new file from local folder to sharepoint
    files_in_folder = os.listdir(reports_local_path)
    while not any(file_to_find in filename for filename in files_in_folder):
        logging.info("Waiting for the report to download...")
        time.sleep(60) 
        files_in_folder = os.listdir(reports_local_path)

    for filename in os.listdir(reports_local_path):
        if file_to_find in filename:
            file_path = os.path.join(reports_local_path, filename)
            shutil.move(file_path, os.path.join(sharepoint_path, filename))
    
    logging.info(f"VIM_ILC Report(s) saved in Sharepoint folder successfully\n")

else:
    ###-> Delete new files from local folder to avoid problems in the next execution
    for filename in os.listdir(reports_local_path):
        if file_to_find in filename:
            file_path = os.path.join(reports_local_path, filename)
            os.remove(file_path)
    logging.error(f"VIM_ILC Report(s) that were generated by this execution were deleted from the local folder.\n")


end_time = time.time()
execution_duration = round(end_time - start_time, 2)
logging.info("Automation runtime: " + str((execution_duration)) + ' seconds\n')

for handler in logging.getLogger().handlers:
    if isinstance(handler, logging.FileHandler):
        handler.close()

print(result)