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

How cancelled leave is handled:

 - Revoking or rejecting the request in Ubiquiti removes the calendar event and the Google Sheet row on the next run. This is the authoritative way to cancel leave and works for past and future dates.
 - Deleting the event from Google Calendar by hand also works for leave that has not started yet. On the next run the program notices the event is gone, drops the row from the Google Sheet (so dashboards built on it update), and records the Approval ID in the deleted-events file so the event is not re-created.
 - Deleting a calendar event for leave whose date has already passed does not remove the record; the event is re-created on the next run. Revoke it in Ubiquiti instead.
 - To bring a request back after it was removed by hand, delete its Approval ID from the deleted-events file; the event and sheet row are re-created on the next run.
