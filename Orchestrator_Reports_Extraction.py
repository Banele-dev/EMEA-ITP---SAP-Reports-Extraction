import subprocess
import logging
import os
from datetime import datetime, timedelta
import time
import win32com.client as win32
import pandas as pd

user_name = os.getlogin()
start_time = time.time()

# Get the path of the directory where the script is located
script_directory = f"C:/Users/{user_name}/Anglo American/GSS Automation Team - General/03_Documentation - Automation Initiatives/02_EMEA/EMEA_ITP_SAP Reports Extraction"

# Create the path for the LogControl folder 
log_control_path = os.path.join(script_directory, 'LogControl')
if not os.path.exists(log_control_path):
    os.makedirs(log_control_path) # If the LogControl folder doesn't exist, create it

# Create the full path to the log file within the LogControl folder
log_file_name = f"OrchestratorLog_{datetime.now().strftime('%d%m%Y%H%M')}"+".txt"
log_file_path = os.path.join(log_control_path, log_file_name)

logging.basicConfig(
    level=logging.INFO,  
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file_path), 
        logging.StreamHandler()  
    ])


## Setting variables to check is this version matches with the GSS Automation Team's control
# app = "EMEA ITP_SAP Reports Extraction"
# version = "v01"
# user_name = os.getlogin()
# path = r'C:\\Users\\' + user_name +  r"\\Box\Automation Script Versions\\versions.xlsx"
# df = pd.read_excel(path)
# filter_criteria = (df['app'] == app) & (df['versão'] == version)

# if not filter_criteria.any():
#     logging.info('Outdated script, talk to the automation team. ')
#     quit()


def execute_script(script_path, script_name):
    try:
        logging.info(f'Starting script execution: {script_name}')
        result = subprocess.run(['python', script_path], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)

        # Print the subprocess output
        output = result.stdout

        # Check for success or failure in the subprocess output
        if "Success" in output:
            logging.info(f'Script executed successfully: {script_name}\n')
            return "Success"
        elif "Failed" in output:
            logging.error(f"Error during {script_name} execution. See specific execution log for details.\n")
            return "Failed"
        else:
            logging.error(f"Unexpected output during {script_name} execution:\n{output}\n")
            return "Failed"

    except Exception as e:
        logging.error(f"Unexpected error during {script_name} execution: {e}\n")
        return "Failed"


# Subscripts path
vim_va2_path = f'{script_directory}\\VIM_VA2_generator.py'
vim_ilc_path = f'{script_directory}\\VIM_ILC_generator.py'
me2l_path = f'{script_directory}\\ME2L_generator.py'
fbl1n_path = f'{script_directory}\\FBL1N_generator.py'
fbl3n_path = f'{script_directory}\\FBL3N_generator.py'
pf03_wp_path = f'{script_directory}\\PF03_WP_generator.py'
emea_validation_path = f'{script_directory}\\EMEA_Validation_generator.py'

# Execution Criteria
run_vim_va2 = execute_script(vim_va2_path, 'VIM_VA2_generator.py')
if run_vim_va2 == "Success": #it will only execute both scripts if VA2 has been sucessfully generated
    run_vim_ilc = execute_script(vim_ilc_path, 'VIM_ILC_generator.py')
    run_me2l = execute_script(me2l_path, 'ME2L_generator.py')
    quantity_of_reports = 3


run_fbl1n = execute_script(fbl1n_path, 'FBL1N_generator.py')
if run_fbl1n == "Success":
    quantity_of_reports = 4
    
run_fbl3n = execute_script(fbl3n_path, 'FBL3N_generator.py')
if run_fbl3n == "Success":
    quantity_of_reports = 5

run_pf03_wp = execute_script(pf03_wp_path, 'PF03_WP_generator.py')
run_pf03_wp = "Success"
if run_pf03_wp == "Success":
    quantity_of_reports = 6
    run_emea_validation = execute_script(emea_validation_path, 'EMEA_Validation_generator.py')
    if run_emea_validation == "Success":
        quantity_of_reports = 7

if run_vim_va2 == "Failed" or run_vim_ilc == "Failed" or run_me2l == "Failed" or run_fbl1n == "Failed" or run_fbl3n == "Failed" or run_pf03_wp == "Failed" or run_emea_validation=="Failed" :
    status_automation = "Failed"
else:
    status_automation = "Success"


#Send e-mail to users for generating the report manually
failure_time = datetime.now().strftime('%d/%m/%Y')

if status_automation == "Failed":
    outlook = win32.Dispatch('Outlook.Application')
    email = outlook.CreateItem(0)
    #email.SentOnBehalfOfName = "marcelo.s.araujo@angloamerican.com" #adjust to Bot Agent's adress
    email.To = "marcelo.s.araujo@angloamerican.com"
    #email.CC = "marcelo.s.araujo@angloamerican.com"
    email.Subject = "Automação de Extração dos Relatorios no SAP (EMEA) - Rotina Falhou"
    email.HTMLBody = """
                        <html>
                            <body>
                                <p>Dear All,</p>
                                <p>The automatic extraction routine for the SAP reports failed to execute on <strong>{}</strong></p>
                                <p>To avoid issues in other processes that depend on these reports, we request that you check which report failed during the extraction and then generate it <strong>manually</strong>.</p>
                                <p>Please note that the automatic extraction routine runs twice a day (at 05:00 AM - BRT and 10:00 PM - BRT).</p> 
                                <p>Regarding the automation, we advise that you <strong>wait until the next execution</strong> to assess if it stabilizes. If there is a failure in tomorrow's execution, please contact the Automation Team for evaluation.</p>
                                <p>Sincerely,</p>
                                <p>GSS Automation Team</p>
                            </body>
                        </html>
                    """.format(failure_time)

    email.Send()
    logging.info(f"Notification sent to the business so users can generate the reports manually.")


end_time = time.time()
execution_duration = round(end_time - start_time, 2)
logging.info("Numberes of reports genereted:" + str((quantity_of_reports)) + '\n')
logging.info("Automation runtime: " + str((execution_duration)) + ' seconds\n\n')

try:
    outlook = win32.Dispatch('Outlook.Application')
    email = outlook.CreateItem(0)
    #email.SentOnBehalfOfName = "marcelo.s.araujo@angloamerican.com" #adjust to Bot Agente's adress
    email.To = "breno.andrade@angloamerican.com"
    #email.To = "marcelo.s.araujo@angloamerican.com"
    email.Subject = "Automation Team - Automation Log"
    # automation name_ date of execution_status of execution_duration of execution_process data_type of process data
    email.HTMLBody = "EMEA ITP - SAP Reports Extraction" + "_" + str(datetime.today()) + "_" + str(status_automation) + "_" + str(execution_duration) + "_" + str(quantity_of_reports) + "_" + "reports generated"
    attach_log_file= email.Attachments.Add(log_file_path)
    email.Send()

    logging.info(f"Notification sent to the GSS Automation Team (Sharepoint Data Base).")
except:
    logging.error(f"Error trying to send notification to the GSS Automation Team (Sharepoint Databse).")

for handler in logging.getLogger().handlers:
    if isinstance(handler, logging.FileHandler):
        handler.close()

print("FINISHED")