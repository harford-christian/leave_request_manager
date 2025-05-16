Description of the program:

This program is designed to take information from the “Leave & Sub Requests” Workflow on Ubiquiti’s UID Workspace (https://harfordchristian.ui.com/cloud/workflow/approvals) and use it to create an event on Google Calendar. To accomplish this task, a script will open Ubiquiti, sign in as a specified user, and extract the information as an Excel file. This file is automatically downloaded into the downloads folder. Next, the program will read specified fields from the Excel file and use them to create an event in a specified user’s Google Calendar using a service account. The batch file will run this program.

Description of each file in the project directory:

chromedriver-win64 – windows 64-bit installation of chromedriver, the driver that runs the script

venv – a virtual environment; creates a separate space for the program to run in

.env – a file used to store sensitive information

deleted_events_list – a collection of IDs of events that were deleted; prevents unwanted events

leave_requests – the program itself

requirements – a list of all dependencies and their versions; used to aid in dependency installation

run_leave_requests.bat – batch file that runs the program

credentials.json – this file must be downloaded from the Google Cloud Console after a service account is created

How to run the program manually:

 1. open a terminal in the project directory (or cd into the project directory)
 2. start the virtual environment (in command prompt, venv\scripts\activate)
 3. run the program (python leave_requests.py)

A more detailed description of the program itself can be found in the comments within the program
