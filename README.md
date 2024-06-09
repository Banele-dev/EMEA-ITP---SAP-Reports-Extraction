
# Project Title:
EMEA_ITP_SAP Reports Extraction

# Description:
The current process of data extraction from SAP for the Backlog Management Dashboard in EMEA ITP is impaired by inefficiencies, consuming a significant amount of time ranging from 45 to 60 minutes per day. This prolonged duration not only hampers operational efficiency but also exposes the system to the risk of human errors, undermining the accuracy of the extracted data.

To address these challenges, we proposed the implementation of an automated solution aimed at streamlining the extraction of SAP reports essential for the Backlog Management Dashboard. The initial focus was on automating the extraction process for these specific reports, with the potential for replication across various reports currently managed manually by different teams.

By implementing this automated solution, we aimed to significantly reduce the time spent on data extraction, enhance operational efficiency, and improve the accuracy of the extracted data. This initiative not only mitigated the risk of human errors but also freed up valuable resources, allowing teams to focus on more strategic tasks. Ultimately, the automation of SAP report extraction contributed to a more efficient, accurate, and reliable reporting process within EMEA ITP.

# Prerequisites and Dependencies:
To proceed with the task, I needed need access to all SAP and to the coreect transactional codes and company codes

# Code Explanation: 
I have developed eight scripts for this project, seven of which are dedicated to extracting various reports based on different business and process rules. The remaining script orchestrates the execution of these scripts according to specific criteria. Below is the list of scripts I have developed:
1. VIM_VA2_generator
2. VIM_ILC_generator
3. PF03_WP_generator
4. ME2L_generator
5. FBL3N_generator
6. FBL1N_generator
7. EMEA_Validation_generator
8. Orchestrator_Reports_Extraction

# Outcome Achieved:
1. Time Efficiency: Reduced the data extraction time from 45-60 minutes to a fraction of that, enabling quicker access to essential reports.
2. Operational Efficiency: Streamlined the extraction process, allowing for smoother and more efficient operations.
3. Accuracy Improvement: Minimized the risk of human errors, leading to more accurate and reliable data extraction.
4. Resource Optimization: Freed up valuable resources, allowing teams to focus on strategic tasks rather than manual data extraction.
5. Consistency: Ensured consistent and reliable extraction of SAP reports, contributing to better overall data management.
6. Scalability: Created a scalable solution that can be replicated across various reports managed manually by different teams, enhancing the overall efficiency of the organization.
