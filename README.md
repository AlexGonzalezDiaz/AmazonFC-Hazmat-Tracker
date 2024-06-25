# AmazonFC-Hazmat-Tracker
A Script created in python to automatically check the attendance and send the results via slack webhook. 

# Code Summary:
Main Loop:
- **theHarvester** | Python: Pulls only the relevant data from the excell sheet that the team needs to recertify an associate and paste it into specific columns in the excel
- **HazmatCocTracker** | VBA: Module that is our runner to fire all the python and vba scripts in order.
- **DoIt** | VBA: A module that checks the current attendance onsite and will pull the information from that database. 
- **theChef_2** | Python: Parses the data on the spreadsheet such as the associates name, login, location, and Hazmat/Code of Conduct Compliance level and seperates them in an excel sheet tab by dept name. 
The information then gets organized on a markdown table and the payload gets sent through slack.

Automation:
- **ScheduleRun** | Python: Schedules the time intervals when to fire the code again. 
- **RunHazmat** | Python: Fires the main loop and calls the appropraite classes and VBA. 

The solution also has a VBA component which currently authenticates via midway key. That wasn't included since that will contain sensitive information. 
# Workflow Summary:
1. Pulling the .csv containing the information of certificate expiration. 
2. Send the data to the Hazmat_Tracker excel and let the script parse and send the webhook. 
3. Run the ScheduleRun script, which will automatically continue pulling attendance for the entire day at scheduled intervals. 
4. Repeat for the next oncoming shift. 

# My Contribution:
1. Reverse engineered the VBA code for our provided attendance tracker and created VBA macros to facilitate running auxiliary Python code.
2. Created theChef_2 which parses the data into seperate sheets and sends the webhook to slack.
3. Created ScheduleRun and RunHazmat Scripts that fires the code on a scheduled time.

# Workflow Detailed:
The excel sheet has button to run the code manually: 
![image](https://github.com/AlexGonzalezDiaz/AmazonFC-Hazmat-Tracker/assets/138159711/f5097a7f-d161-41c1-9b64-b25a787c1f49)

This will fill in the bolded columns in the sheet wiht all sensitive information that will be parsed out and sent via webhook to slack.  

The script will send a table through slack webhook with logins, name,  locations, when the certificate will expire, and cohort of the associate.

![image](https://github.com/AlexGonzalezDiaz/AmazonFC-Hazmat-Tracker/assets/138159711/57991a2f-c111-4aee-b233-dcfa27c776e5)

After the information is saved onto the spreadsheet, the user will then execute "ScheduleRun.exe" which will then check attendance again through programmed time intervals.  

(Login and Name not shown)

# Outcomes:

1. This lead to the team having consistent 98%+ on compliance for our building. 
2. Always making sure FLEX AAs are compliant, the script is automatically checking attendance. 
3. Easier workflow for the team since they don't need to run the script manually throughout the day.  

# Future Updates:

1. We are attempting to pull the data directly from the database and have the script read the data realtime. This will take out the need to pull the .csv from the website and instead just run the script automatically.  
