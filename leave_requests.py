"""This program is designed to take information from the "Leave & Sub Requests" Workflow on Ubiquiti's UID Workspace 
(https://{WORKSPACE_DOMAIN}.ui.com/cloud/workflow/approvals) and use it to create an event on Google Calendar. To 
accomplish this task, a script will open Ubiquiti, sign in as a specified user, and extract the information as an 
Excel file. This file is automatically downloaded into the downloads folder. Next, the program will read specified 
fields from the Excel file and use them to create an event in a specified user's Google Calendar using a service 
account. The batch file will run this program at regular intervals determined by Windows Task Scheduler."""

"""All sensitive information (passwords,usernames,IDs,filepaths) must be set in the .env file."""

"""All comments are either green or red. Green comments are used to explain blocks of code, while red comments 
are used to explain functions."""

# Installed libraries
import os
import time
import logging
import pandas as pd
import google.oauth2.service_account as service_account
from datetime import datetime
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager 
from google.oauth2 import service_account
from googleapiclient.discovery import build

# Set up logging
log_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'logs')
os.makedirs(log_dir, exist_ok=True)
log_file = os.path.join(log_dir, f'leave_requests_{datetime.now().strftime("%Y%m%d")}.log')
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file),
        logging.StreamHandler()
    ]
)

# Loads the variables from the .env file
load_dotenv() 

# Path to a file for deleted event IDs
deleted_events_file = os.getenv("DELETED_EVENTS")

# Get website for login
website = os.getenv("LOGIN_URL")

# Get impersonated user email
impersonated_user = os.getenv("IMPERSONATED_USER_EMAIL")

# Main part of the function
class LeaveRequestCalendar:
    def __init__(self, service_account_file, calendar_id, sheets_id=None):
        self.calendar_id = calendar_id
        self.sheets_id = sheets_id
        self.SCOPES = [
            'https://www.googleapis.com/auth/calendar',
            'https://www.googleapis.com/auth/spreadsheets'
        ]
        self.calendar_service = self.setup_google_calendar(service_account_file)
        self.sheets_service = self.setup_google_sheets(service_account_file)
        self.deleted_set = self.load_and_save_deleted_events()
        
    def setup_chrome_driver(self):
        """Set up Chrome WebDriver with download settings"""
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument("--headless") # Runs the script without opening a browser window
        chrome_options.add_argument("--disable-gpu") # Disable GPU rendering
        chrome_options.add_argument("--disable-extensions") # Disables extensions that may interfere with the script
        chrome_options.add_argument("--disable-software-rasterizer")  # Add this to avoid GPU rendering issues
        chrome_options.add_argument("--disable-dev-shm-usage")  # Helps avoid shared memory issues
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
        return driver 
    
    def download_excel(self):
        """Download Excel file from the UI website"""
        
        driver = self.setup_chrome_driver()
        try:
            # Navigate to the website
            driver.get(website)
                      
            # Wait for page to load and elements to be present
            wait = WebDriverWait(driver, 20)
            
            # Load username from .env file
            uid_username = os.getenv("UID_USERNAME")
            
            # Enter username
            username_field = wait.until(EC.presence_of_element_located((By.ID, "Email"))) 
            username_field.send_keys(uid_username) 
            
            # Click Next
            next_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[.//span[text()='Next']]")))
            next_button.click()
            
            # Load password from .env file
            uid_password = os.getenv("PASSWORD")
            
            # Enter password 
            password_field = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='password']")))  
            password_field.send_keys(uid_password)
            
            # Click login button
            signin_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'button__Ad6q97AA') and contains(@class, 'button_ui-login-button__V7mNB') and .//span[text()='Sign In']]")))  # Adjust selector as needed for the login button
            signin_button.click()
            
            # Wait for page to load and elements to be present
            wait = WebDriverWait(driver, 20)
            
            # Click on "Manager Portal"
            manager_portal = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "svg path[d^='M0 20C0 12.9993 0 9.49902 1.36242']")))            
            manager_portal.click()
            
            # Wait for new window and switch to it
            wait.until(lambda driver: len(driver.window_handles) > 1)
            driver.switch_to.window(driver.window_handles[-1])  # Switch to the latest window
            
            # After switching to new window, maximize it
            wait.until(lambda driver: len(driver.window_handles) > 1)
            driver.switch_to.window(driver.window_handles[-1])
            driver.maximize_window()
            
            # Click on "Workflows and Approvals"
            workflows_approvals = wait.until(EC.element_to_be_clickable((By.XPATH, "//li[@data-testid='Workflows and Approvals' and contains(@class, 'navItem__TNKCRCxr')]")))
            workflows_approvals.click()
            
            # Click on "Requests"
            requests = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@type='button' and @data-testid='approvals' and span[text()='Requests']]")))
            requests.click()
            
            # Click on download button
            download_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "div > svg > path[d*='M10.367 13.887']")))
            download_button.click()
            
            # Click on the approval dropdown
            approval_dropdown = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Select']")))
            approval_dropdown.click()
            approval_dropdown.clear()
            approval_dropdown.send_keys('Leave & Sub Request')
            approval_dropdown.send_keys(Keys.ENTER)
            approval_dropdown.send_keys('Leave: HCS-Related')
            approval_dropdown.send_keys(Keys.ENTER)
            approval_dropdown.send_keys('Leave: Late Arrival/Early Departure')
            approval_dropdown.send_keys(Keys.ENTER)
            approval_dropdown.send_keys('Leave: Other Paid Leave')
            approval_dropdown.send_keys(Keys.ENTER)
            approval_dropdown.send_keys('Leave: Personal Leave')
            approval_dropdown.send_keys(Keys.ENTER)
            approval_dropdown.send_keys('Leave: Sick Leave')
            approval_dropdown.send_keys(Keys.ENTER)
            
            # Select the time period to be the last month (can be commented out if needed)
            # button = driver.find_element(By.XPATH, "//div[normalize-space(text())='1 Month']")
            # driver.execute_script("arguments[0].click();", button)
            
            # Select the "Status" dropdown
            status_dropdown = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@data-testid='inputReadOnly']//span[text()='Select']")))
            status_dropdown.click()
            
            '''# Uncomment this section if you want to select "pending" from the dropdown
            # Select "Pending" from the dropdown
            select_pending = wait.until(EC.element_to_be_clickable((By.ID, 'dropdownOptions_pending')))
            select_pending.click()'''
            
            # Select "Approved" from the dropdown
            select_approved = wait.until(EC.element_to_be_clickable((By.ID, 'dropdownOptions_pass')))
            select_approved.click()
            
            # Select "Rejected" from the dropdown
            select_denied = wait.until(EC.element_to_be_clickable((By.ID, 'dropdownOptions_reject')))
            select_denied.click()
            
            # Select "Revoked" from the dropdown
            select_revoked = wait.until(EC.element_to_be_clickable((By.ID, 'dropdownOptions_cancel')))
            select_revoked.click()
            
            # Click Export button
            export_button = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//div[contains(@class, 'wrap__VCR3r9bC')]//button[.//span[text()='Export']]")))
            export_button.click()
            
            # Wait for download to complete
            time.sleep(2)  # Adjust if needed
        
        # Close the browser    
        finally:
            driver.quit()
    
    def setup_google_calendar(self, service_account_file):
        """Set up Google Calendar API with service account"""
        
        # Load credentials from service account file
        credentials = service_account.Credentials.from_service_account_file(
            service_account_file, scopes=self.SCOPES, subject=impersonated_user)
        
        # Return the API client
        return build('calendar', 'v3', credentials=credentials)
    
    def setup_google_sheets(self, service_account_file):
        """Set up Google Sheets API with service account"""
        
        # Load credentials from service account file
        credentials = service_account.Credentials.from_service_account_file(
            service_account_file, scopes=self.SCOPES, subject=impersonated_user)
        
        # Return the API client
        return build('sheets', 'v4', credentials=credentials)
    
    def get_latest_excel_file(self, download_dir):
        """Get the most recently downloaded Excel file"""
        files = [f for f in os.listdir(download_dir) if f.endswith('.xlsx')]
        if not files:
            raise FileNotFoundError("No Excel files found in download directory")
        return max(
            [os.path.join(download_dir, f) for f in files],
            key=os.path.getmtime
        )
        
    def get_existing_events(self):
        """Get all existing leave request events from the calendar"""
        events = []
        page_token = None
        
        while True:
            # Query events with a prefix match on summary to get leave requests
            events_result = self.calendar_service.events().list(
                calendarId=self.calendar_id,
                pageToken=page_token,
            ).execute()
            
            for event in events_result.get('items', []):
                """Extract Approval ID from event description if it exists"""
                description = event.get('description', '')
                if 'Approval ID:' in description:
                    approval_id = description.split('Approval ID:')[1].split('\n')[0].strip()
                    
                    # Extract substitute name from existing calendar event
                    current_summary = event.get('summary', '')
                    sub_name_in_calendar = None
                    
                    # Check for the format: "{sub_name}/{full_name}/{type}"
                    if '/' in current_summary:
                        parts = current_summary.split('/')
                        if len(parts) >= 1:
                            sub_name_in_calendar = parts[0].strip()
                    
                    # Get event creation/update time for duplicate detection
                    event_created = event.get('created', '')
                    event_updated = event.get('updated', '')
                    
                    events.append({
                        'event_id': event['id'],
                        'approval_id': approval_id,
                        'event': event,
                        'calendar_sub_name': sub_name_in_calendar,
                        'created': event_created,
                        'updated': event_updated
                    })
            page_token = events_result.get('nextPageToken')
            if not page_token:
                break
        
        # Handle duplicates: keep the most recent event, delete older ones
        approval_id_map = {}
        duplicates_deleted = 0
        
        for event_info in events:
            approval_id = event_info['approval_id']
            
            if approval_id in approval_id_map:
                # Duplicate found - compare timestamps to keep the newest
                existing_event = approval_id_map[approval_id]
                
                # Compare by 'updated' timestamp (most recent modification)
                existing_updated = existing_event['updated']
                current_updated = event_info['updated']
                
                if current_updated > existing_updated:
                    # Current event is newer - delete the old one
                    try:
                        self.calendar_service.events().delete(
                            calendarId=self.calendar_id,
                            eventId=existing_event['event_id'],
                            sendUpdates='none'  # Don't send emails for duplicate cleanup
                        ).execute()
                        logging.info(f"Deleted duplicate event (older) for Approval ID: {approval_id}")
                        print(f"Deleted duplicate event (older) for Approval ID: {approval_id}")
                        duplicates_deleted += 1
                        # Keep the newer event
                        approval_id_map[approval_id] = event_info
                    except Exception as e:
                        logging.error(f"Error deleting duplicate event: {str(e)}")
                else:
                    # Existing event is newer - delete the current one
                    try:
                        self.calendar_service.events().delete(
                            calendarId=self.calendar_id,
                            eventId=event_info['event_id'],
                            sendUpdates='none'  # Don't send emails for duplicate cleanup
                        ).execute()
                        logging.info(f"Deleted duplicate event (older) for Approval ID: {approval_id}")
                        print(f"Deleted duplicate event (older) for Approval ID: {approval_id}")
                        duplicates_deleted += 1
                        # Keep the existing (newer) event
                    except Exception as e:
                        logging.error(f"Error deleting duplicate event: {str(e)}")
            else:
                # First occurrence of this approval_id
                approval_id_map[approval_id] = event_info
        
        if duplicates_deleted > 0:
            logging.info(f"Removed {duplicates_deleted} duplicate calendar events")
            print(f"Removed {duplicates_deleted} duplicate calendar events")
        
        return approval_id_map
    
    def update_calendar_event(self, event_id, event_body):
        """Update an existing calendar event"""
        try:
            self.calendar_service.events().update(
                calendarId=self.calendar_id,
                eventId=event_id,
                body=event_body,
                sendUpdates='all'
            ).execute()
            return True
        except Exception as e:
            print(f"Error updating calendar event: {str(e)}")
            return False
        
    def load_and_save_deleted_events(self, approval_id=None):
        """Load the deleted events from a file, optionally adding a new approval_id and saving back"""
        deleted_set = set() # Initialize empty set
        try:
            with open(deleted_events_file, 'r') as f:
                deleted_set = set(f.read().splitlines())  # Convert lines into a set
        except FileNotFoundError:
            deleted_set = set()  # Initialize empty set if the file doesn't exist

        # If an approval_id is provided, add it to the set
        if approval_id and approval_id not in deleted_set:
            deleted_set.add(approval_id)
            # Save the updated set back to the file
            with open(deleted_events_file, 'w') as f:
                f.write('\n'.join(deleted_set))  # Write each approval_id on a new line

        return deleted_set  # Return the current set of deleted events
    
    def delete_sheet_row(self, row_index):
        """Delete a specific row from Google Sheets"""
        try:
            # Create delete request
            delete_request = {
                'deleteDimension': {
                    'range': {
                        'sheetId': 0,  # Assuming first sheet, adjust if needed
                        'dimension': 'ROWS',
                        'startIndex': row_index - 1,  # Convert to 0-based index
                        'endIndex': row_index  # End index is exclusive
                    }
                }
            }
            
            # Execute the delete request
            self.sheets_service.spreadsheets().batchUpdate(
                spreadsheetId=self.sheets_id,
                body={'requests': [delete_request]}
            ).execute()
            
            print(f"Deleted row {row_index} from Google Sheets")
            logging.info(f"Deleted row {row_index} from Google Sheets")
            return True
            
        except Exception as e:
            print(f"Error deleting row {row_index} from Google Sheets: {str(e)}")
            logging.error(f"Error deleting row {row_index} from Google Sheets: {str(e)}")
            return False
    
    def setup_sheets_headers(self):
        """Set up the headers in the Google Sheet if they don't exist"""
        if not self.sheets_id:
            return
            
        try:
            # Check if sheet exists and has headers
            result = self.sheets_service.spreadsheets().values().get(
                spreadsheetId=self.sheets_id,
                range='A1:Z1'
            ).execute()
            
            values = result.get('values', [])
            
            # Define the headers we want
            headers = [
                'Approval ID', 'First Name', 'Last Name', 'Time Off Type', 
                'Status', 'Start Time', 'End Time', 'Duration', 'Substitute', 'Sub Required?',
                'Reason', 'Additional comments', 'Last Updated', 'Calendar Event Status'
            ]
            
            # If no headers exist or headers are different, set them
            if not values or values[0] != headers:
                result = self.sheets_service.spreadsheets().values().update(
                    spreadsheetId=self.sheets_id,
                    range='A1:N1',
                    valueInputOption='RAW',
                    body={'values': [headers]}
                ).execute()
                logging.info("Headers set up in Google Sheet")
        
        except Exception as e:
            logging.error(f"Error setting up sheet headers: {str(e)}")
            import traceback
            logging.error(f"Full traceback: {traceback.format_exc()}")
    
    def get_existing_sheet_data(self):
        """Get existing data from Google Sheets to avoid duplicates"""
        if not self.sheets_id:
            return {}
            
        try:
            result = self.sheets_service.spreadsheets().values().get(
                spreadsheetId=self.sheets_id,
                range='A:N'  # Include all columns including Calendar Event Status
            ).execute()
            
            values = result.get('values', [])
            
            if len(values) <= 1:  # Only headers or empty
                return {}
            
            # Create a dictionary with Approval ID as key
            existing_data = {}
            
            for i, row in enumerate(values[1:], 2):  # Start from row 2 (skip headers)
                if row and len(row) > 0:  # Make sure row has data
                    approval_id = row[0] if len(row) > 0 else ''
                    if approval_id:
                        existing_data[approval_id] = {
                            'row_index': i,
                            'data': row
                        }
            return existing_data
            
        except Exception as e:
            logging.error(f"Error getting existing sheet data: {str(e)}")
            import traceback
            logging.error(f"Full traceback: {traceback.format_exc()}")
            return {}
    
    def sync_calendar_to_sheets(self):
        """Sync manually edited substitute names from calendar back to Google Sheets"""
        if not self.sheets_id:
            logging.info("No Google Sheets ID provided, skipping calendar sync")
            return
        
        try:
            logging.info("Syncing calendar edits to Google Sheets...")
            
            # Get all existing calendar events
            existing_events = self.get_existing_events()
            
            # Get existing sheet data
            existing_sheet_data = self.get_existing_sheet_data()
            
            updates = []
            
            for approval_id, event_info in existing_events.items():
                # Check if this event has a substitute in the calendar
                calendar_sub_name = event_info.get('calendar_sub_name')
                
                if approval_id in existing_sheet_data and calendar_sub_name:
                    sheet_row = existing_sheet_data[approval_id]
                    sheet_data = sheet_row['data']
                    
                    # Check if sheet has the substitute column (index 8)
                    if len(sheet_data) > 8:
                        sheet_sub_name = sheet_data[8] if sheet_data[8] else ''
                        
                        # If calendar substitute differs from sheet substitute, update the sheet
                        # Don't sync "NEEDS SUB" placeholders
                        if calendar_sub_name != sheet_sub_name and calendar_sub_name != 'NEEDS SUB':
                            # Update the substitute column (column I, index 8)
                            row_index = sheet_row['row_index']
                            updates.append({
                                'range': f'I{row_index}',
                                'values': [[calendar_sub_name]]
                            })
                            logging.info(f"Syncing substitute for approval {approval_id}: '{sheet_sub_name}' -> '{calendar_sub_name}'")
            
            # Perform batch updates
            if updates:
                batch_update_body = {
                    'valueInputOption': 'RAW',
                    'data': updates
                }
                self.sheets_service.spreadsheets().values().batchUpdate(
                    spreadsheetId=self.sheets_id,
                    body=batch_update_body
                ).execute()
                logging.info(f"Synced {len(updates)} manually edited substitutes from calendar to sheets")
                print(f"Synced {len(updates)} manually edited substitutes from calendar to sheets")
            else:
                logging.info("No calendar edits to sync to sheets")
                
        except Exception as e:
            logging.error(f"Error syncing calendar to sheets: {str(e)}")
            import traceback
            logging.error(f"Full traceback: {traceback.format_exc()}")
    
    def update_sheets_data(self, df, calendar_events_status, manually_edited_subs):
        """Update Google Sheets with leave request data"""
        if not self.sheets_id:
            logging.info("No Google Sheets ID provided, skipping sheets update")
            return
        
        try:
            # Set up headers first
            self.setup_sheets_headers()
            
            # Get existing data to determine what to update vs insert
            existing_data = self.get_existing_sheet_data()
            
            # Prepare data for batch update
            updates = []
            new_rows = []
            rows_to_delete = []
            
            # First, check existing sheet data for "Previously Deleted" entries that should be removed
            for approval_id, data_info in existing_data.items():
                if len(data_info['data']) >= 14:  # Make sure we have the calendar status column
                    existing_calendar_status = data_info['data'][13]  # Calendar Event Status is column N (index 13)
                    if existing_calendar_status == 'Previously Deleted':
                        rows_to_delete.append({
                            'approval_id': approval_id,
                            'row_index': data_info['row_index']
                        })
                        print(f"Marking previously deleted row for removal: {approval_id}")
            
            for _, row in df.iterrows():
                approval_id = str(row['Approval ID'])
                
                # Get calendar event status
                calendar_status = calendar_events_status.get(approval_id, 'Unknown')
                
                # If the calendar event was deleted, mark the row for deletion
                if calendar_status == 'Deleted':
                    if approval_id in existing_data:
                        rows_to_delete.append({
                            'approval_id': approval_id,
                            'row_index': existing_data[approval_id]['row_index']
                        })
                    continue  # Skip processing this row further
                
                # Skip if this approval_id is marked for deletion (either just deleted or previously deleted)
                if any(item['approval_id'] == approval_id for item in rows_to_delete):
                    continue
                
                # Get duration from Excel file (already calculated by Ubiquiti)
                duration = ""
                if 'Amount (Hours)' in row and pd.notna(row['Amount (Hours)']):
                    amount_str = str(row['Amount (Hours)'])
                    # Extract numeric value from strings like "8 Hours" or "8.5"
                    import re
                    hours_match = re.search(r'(\d+\.?\d*)', amount_str)
                    if hours_match:
                        hours = float(hours_match.group(1))
                        if hours < 24:
                            duration = f"{hours:.1f} hours"
                        else:
                            days = int(hours // 24)
                            remaining_hours = hours % 24
                            if remaining_hours == 0:
                                duration = f"{days} day{'s' if days != 1 else ''}"
                            else:
                                duration = f"{days} day{'s' if days != 1 else ''}, {remaining_hours:.1f} hours"
                
                # Determine which substitute name to use
                # Priority: manually edited sub from calendar > Ubiquiti data
                substitute_to_use = manually_edited_subs.get(approval_id, row['Substitute'] if pd.notna(row['Substitute']) else '')
                
                # Prepare row data for non-deleted events
                row_data = [
                    approval_id,
                    row['First Name'],
                    row['Last Name'], 
                    row['Time Off Type'],
                    row['Status'],
                    row['Start Time'].strftime('%Y-%m-%d %H:%M:%S') if pd.notna(row['Start Time']) else '',
                    row['End Time'].strftime('%Y-%m-%d %H:%M:%S') if pd.notna(row['End Time']) else '',
                    duration,
                    substitute_to_use,  # Use the calendar version if it was manually edited
                    row['Sub Required?'],
                    row['Reason'] if pd.notna(row['Reason']) else '',
                    row['Additional comments'] if pd.notna(row['Additional comments']) else '',
                    datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    calendar_status
                ]
                
                if approval_id in existing_data:
                    # Update existing row
                    row_index = existing_data[approval_id]['row_index']
                    updates.append({
                        'range': f'A{row_index}:N{row_index}',
                        'values': [row_data]
                    })
                else:
                    # New row to be added
                    new_rows.append(row_data)
            
            # Delete rows first (start from highest row index to avoid index shifting)
            if rows_to_delete:
                # Sort by row_index in descending order to delete from bottom up
                rows_to_delete.sort(key=lambda x: x['row_index'], reverse=True)
                deleted_count = 0
                
                for row_info in rows_to_delete:
                    if self.delete_sheet_row(row_info['row_index']):
                        deleted_count += 1
                        print(f"Deleted sheet row for approval ID: {row_info['approval_id']}")
                
                logging.info(f"Deleted {deleted_count} rows from Google Sheets")
            
            # Perform batch updates for existing rows
            if updates:
                batch_update_body = {
                    'valueInputOption': 'RAW',
                    'data': updates
                }
                self.sheets_service.spreadsheets().values().batchUpdate(
                    spreadsheetId=self.sheets_id,
                    body=batch_update_body
                ).execute()
                logging.info(f"Updated {len(updates)} existing rows in Google Sheets")
            
            # Append new rows
            if new_rows:
                self.sheets_service.spreadsheets().values().append(
                    spreadsheetId=self.sheets_id,
                    range='A:N',
                    valueInputOption='RAW',
                    insertDataOption='INSERT_ROWS',
                    body={'values': new_rows}
                ).execute()
                logging.info(f"Added {len(new_rows)} new rows to Google Sheets")
                
        except Exception as e:
            logging.error(f"Error updating Google Sheets: {str(e)}")
    
    def create_calendar_events(self, excel_path):
        """Read Excel (all sheets) and create calendar events for each leave request"""
        
        # Read all sheets into a dict {sheet_name: DataFrame}
        sheets = pd.read_excel(excel_path, sheet_name=None)
        
        # Get existing events from calendar (calendar entries take precedence)
        existing_events = self.get_existing_events()
        
        # Initialize counters and status map
        created_count = existing_count = deleted_count = ignored_count = 0
        calendar_events_status = {}
        manually_edited_subs = {}  # Track manually edited substitutes
        
        # Combine all sheets into a single DataFrame
        all_data = []
        for sheet_name, df in sheets.items():
            if df is None or df.empty:
                continue
            # Ensure datetimes are parsed per-sheet
            df['Start Time'] = pd.to_datetime(df['Start Time'], errors='coerce')
            df['End Time'] = pd.to_datetime(df['End Time'], errors='coerce')
            all_data.append(df)
        
        # If no data found, just sync calendar to sheets and return
        if not all_data:
            logging.info("No new data in Excel file - syncing existing calendar events to sheets")
            self.sync_calendar_to_sheets()
            return
        
        combined_df = pd.concat(all_data, ignore_index=True)
        logging.info(f"Processing {len(combined_df)} total rows from {len(all_data)} sheets")
        
        # Process the combined DataFrame
        for _, row in combined_df.iterrows():
            # Create full name
            full_name = f"{row['First Name']} {row['Last Name']}"
            
            # Initialize approval_id early since we need it for all-day check
            approval_id = str(row['Approval ID'])
            
            # Define default summary based on Ubiquiti data
            default_summary = (f"{row['Substitute']}/{full_name}/{row['Time Off Type']}" 
                            if pd.notna(row['Substitute']) and str(row['Substitute']).strip() != ''
                            else f"NEEDS SUB/{full_name}/{row['Time Off Type']}"
                            if row['Sub Required?'].lower() == 'yes'
                            else f"No Sub/{full_name}/{row['Time Off Type']}")
            
            # Determine if this should be an all-day event
            start_time = row['Start Time']
            end_time = row['End Time']
            is_all_day = False
            
            if pd.notna(start_time) and pd.notna(end_time):
                # Get duration from Excel file (already calculated by Ubiquiti)
                duration_hours = 0
                if 'Amount (Hours)' in row and pd.notna(row['Amount (Hours)']):
                    amount_str = str(row['Amount (Hours)'])
                    # Extract numeric value from strings like "8 Hours" or "8.5"
                    import re
                    hours_match = re.search(r'(\d+\.?\d*)', amount_str)
                    if hours_match:
                        duration_hours = float(hours_match.group(1))
                
                # Check if it covers the full school day
                start_hour = start_time.hour + start_time.minute / 60
                end_hour = end_time.hour + end_time.minute / 60
                
                # Criteria: starts at or before 9 AM, ends at or after 3 PM, and duration >= 6 hours
                if start_hour <= 9 and end_hour >= 15.0 and duration_hours >= 6.0:
                    is_all_day = True
            
            # Create event structure based on whether it's all-day
            if is_all_day:
                # All-day events use 'date' instead of 'dateTime'
                event = {
                    'summary': default_summary,
                    'description': (f"Approval ID: {row['Approval ID']}\n\n"
                                    f"Reason: {row['Reason']}\n\n"
                                    f"Additional Comments: {row['Additional comments']}"),
                    'start': {
                        'date': start_time.strftime('%Y-%m-%d'),
                        'timeZone': 'America/New_York',
                    },
                    'end': {
                        'date': end_time.strftime('%Y-%m-%d'),
                        'timeZone': 'America/New_York',
                    },
                    'reminders': {
                        'useDefault': True
                    }
                }
            else:
                # Regular timed event
                event = {
                    'summary': default_summary,
                    'description': (f"Approval ID: {row['Approval ID']}\n\n"
                                    f"Reason: {row['Reason']}\n\n"
                                    f"Additional Comments: {row['Additional comments']}"),
                    'start': {
                        'dateTime': start_time.isoformat(),
                        'timeZone': 'America/New_York',
                    },
                    'end': {
                        'dateTime': end_time.isoformat(),
                        'timeZone': 'America/New_York',
                    },
                    'reminders': {
                        'useDefault': True
                    }
                }
            
            if approval_id in existing_events and (row['Status'] == 'Approved'):
                
                # Check if there's a substitute in the calendar
                existing_event_info = existing_events[approval_id]
                existing_sub_name = existing_event_info.get('calendar_sub_name')
                
                # Get the Ubiquiti substitute data
                ubiquiti_sub = str(row['Substitute']).strip() if pd.notna(row['Substitute']) else ''
                
                # Check if calendar has a valid substitute name (not "NEEDS SUB" placeholder)
                if existing_sub_name and existing_sub_name != 'NEEDS SUB':
                    # Calendar has a real substitute name
                    # Always preserve what's in the calendar and use it
                    event['summary'] = f"{existing_sub_name}/{full_name}/{row['Time Off Type']}"
                    manually_edited_subs[approval_id] = existing_sub_name
                    
                    # Log only if it differs from Ubiquiti (indicating a manual edit)
                    if existing_sub_name != ubiquiti_sub:
                        logging.info(f"Preserving manually edited sub '{existing_sub_name}' for {full_name} (Ubiquiti has '{ubiquiti_sub}')")

                # Update the event
                try:    
                    success = self.update_calendar_event(
                        existing_events[approval_id]['event_id'],
                        event
                    )
                    if success:
                        existing_count += 1
                        calendar_events_status[approval_id] = 'Updated'
                        print(f"Updated calendar event for {full_name}")
                    else:
                        calendar_events_status[approval_id] = 'Update Failed'
                except Exception as e:
                    calendar_events_status[approval_id] = 'Update Error'
                    print(f"Error updating calendar event for {full_name}: {str(e)}")
                    
            elif approval_id in existing_events and (row['Status'] == 'Rejected' or row['Status'] == 'Revoked'):
                # Delete the event
                try:
                    self.calendar_service.events().delete(
                        calendarId=self.calendar_id,
                        eventId=existing_events[approval_id]['event_id'],
                        sendUpdates='all'
                    ).execute()
                    deleted_count += 1
                    calendar_events_status[approval_id] = 'Deleted'
                    print(f"Deleted calendar event for {full_name}")
                    try: 
                        self.load_and_save_deleted_events(approval_id)
                        print(f"Added event to deleted events: {approval_id}")
                    except Exception as e:
                        print(f"Error adding event to deleted events: {str(e)}")
                except Exception as e:
                    calendar_events_status[approval_id] = 'Delete Error'
                    print(f"Error deleting event: {str(e)}")
                    
            elif approval_id in self.deleted_set: 
                # Event was previously deleted - ignore it
                ignored_count += 1
                calendar_events_status[approval_id] = 'Previously Deleted'
                
            else:
                # Create new event
                try:
                    self.calendar_service.events().insert(
                        calendarId=self.calendar_id,
                        body=event,
                        sendUpdates='all'
                    ).execute()
                    created_count += 1
                    calendar_events_status[approval_id] = 'Created'
                    print(f"Created new calendar event for {full_name}")
                except Exception as e:
                    calendar_events_status[approval_id] = 'Create Error'
                    print(f"Error creating calendar event for {full_name}: {str(e)}")
        
        print(f"\nCalendar Summary:")
        print(f"Created {created_count} new events")
        print(f"Updated {existing_count} existing events")
        print(f"Deleted {deleted_count} existing events")
        print(f"Ignored {ignored_count} previously deleted events")
        if manually_edited_subs:
            print(f"Preserved {len(manually_edited_subs)} manually edited substitutes")
        
        # Pass the combined DataFrame to update_sheets_data with manually edited subs
        if self.sheets_id:
            logging.info("Updating Google Sheets...")
            self.update_sheets_data(combined_df, calendar_events_status, manually_edited_subs)
            print("Google Sheets updated successfully")

def main():
    try:
        logging.info("Starting leave request calendar update...")
        
        # Information for the Google Calendar API
        service_account_file = os.getenv("SERVICE_ACCOUNT_FILE")
        CALENDAR_ID = os.getenv("CALENDAR_ID")
        SHEETS_ID = os.getenv("SHEETS_ID")
        DOWNLOAD_DIR = os.path.expanduser("~/Downloads")
        
        # Create calendar manager
        calendar_manager = LeaveRequestCalendar(service_account_file, CALENDAR_ID, SHEETS_ID)
        
        
        # Download Excel file
        logging.info("Downloading Excel file...")
        calendar_manager.download_excel()
        
        # Get latest Excel file
        excel_file = calendar_manager.get_latest_excel_file(DOWNLOAD_DIR)
        logging.info(f"Found Excel file: {excel_file}")
        
        # Create calendar events and update sheets
        calendar_manager.create_calendar_events(excel_file)
        
        logging.info("Calendar and Sheets update completed successfully")
        
    except Exception as e:
        logging.error(f"Error in main execution: {str(e)}")
        import traceback
        logging.error(f"Full traceback: {traceback.format_exc()}")
        raise

if __name__ == "__main__":
    main()
