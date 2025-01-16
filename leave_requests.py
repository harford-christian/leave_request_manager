# Installed libraries
import os
import time
import logging
import pandas as pd
from datetime import datetime
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
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
deleted_events_file = r"Path\to\file"

# Main part of the function
class LeaveRequestCalendar:
    def __init__(self, service_account_file, calendar_id):
        self.calendar_id = calendar_id
        self.SCOPES = ['https://www.googleapis.com/auth/calendar']
        self.service = self.setup_google_calendar(service_account_file)
        self.deleted_set = self.load_and_save_deleted_events()
        
    def setup_chrome_driver(self):
        """Set up Chrome WebDriver with download settings"""
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument("--headless") # Runs the script without opening a browser window
        chrome_options.add_argument("--disable-gpu") # Disable GPU rendering
        chrome_options.add_argument("--disable-extensions") # Disables extensions that may interfere with the script
        chrome_options.add_argument("--disable-software-rasterizer")  # Add this to avoid GPU rendering issues
        chrome_options.add_argument("--disable-dev-shm-usage")  # Helps avoid shared memory issues
        service = Service(r"Path\to\file")
        return webdriver.Chrome(service=service, options=chrome_options)
    
    def download_excel(self):
        '''Download Excel file from the UI website'''
        
        driver = self.setup_chrome_driver()
        try:
            # Navigate to the website
            driver.get("https://{WORKSPACE_DOMAIN}.ui.com/login")
            
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
            manager_portal = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(@class, 'sc-fihHvN')][.//span[text()='Manager Portal']]")))
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
            approval_dropdown.send_keys('Leave & Sub Request')
            approval_dropdown.send_keys(Keys.ENTER)
            
            # Click somewhere neutral to close the previous dropdown
            body_element = driver.find_element(By.TAG_NAME, "body")
            body_element.click()
            
            # Select the "Status" dropdown
            status_dropdown = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@data-testid='inputReadOnly']//span[text()='Select']")))
            status_dropdown.click()
            
            # Select "Pending" from the dropdown
            select_pending = wait.until(EC.element_to_be_clickable((By.ID, 'dropdownOptions_pending')))
            select_pending.click()
            
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
                (By.XPATH, "//button[contains(., 'Export')]")))
            export_button.click()
            
            # Wait for download to complete
            time.sleep(2)  # Adjust if needed
        
        # Close the browser    
        finally:
            driver.quit()
    
    def setup_google_calendar(self, service_account_file):
        """Set up Google Calendar API with service account"""
        
        # Define the user email to impersonate
        IMPERSONATED_USER_EMAIL = 'email goes here'
        
        # Load credentials from service account file
        credentials = service_account.Credentials.from_service_account_file(
            service_account_file, scopes=self.SCOPES)
        
        # Create delegated credentials
        delegated_credentials = credentials.with_subject(IMPERSONATED_USER_EMAIL)
        return build('calendar', 'v3', credentials=delegated_credentials)

    
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
            events_result = self.service.events().list(
                calendarId=self.calendar_id,
                pageToken=page_token,
                q="Leave Request -"  # This will match our event naming pattern
            ).execute()
            
            for event in events_result.get('items', []):
                # Extract Approval ID from event description if it exists
                description = event.get('description', '')
                if 'Approval ID:' in description:
                    approval_id = description.split('Approval ID:')[1].split(',')[0].strip()
                    events.append({
                        'event_id': event['id'],
                        'approval_id': approval_id,
                        'event': event
                    })
            
            page_token = events_result.get('nextPageToken')
            if not page_token:
                break
                
        return {event['approval_id']: event for event in events}
    
    def update_calendar_event(self, event_id, event_body):
        """Update an existing calendar event"""
        try:
            self.service.events().update(
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
        
    def create_calendar_events(self, excel_path):
        """Read Excel and create calendar events for each leave request"""
        
        # Read Excel file
        df = pd.read_excel(excel_path)
        
        # Convert 'Start Time' and 'End Time' to datetime
        df['Start Time'] = pd.to_datetime(df['Start Time']) 
        df['End Time'] = pd.to_datetime(df['End Time'])
        
        # Get existing events
        existing_events = self.get_existing_events()
        
        # Initialize counters
        created_count = 0
        existing_count = 0
        deleted_count = 0
        ignored_count = 0
        
        for _, row in df.iterrows():
            # Create full name
            full_name = f"{row['First Name']} {row['Last Name']}"
            
            # Create event body
            event = {
                'summary': f"Leave Request - {full_name}",
                'description': (f"Approval ID: {row['Approval ID']}, "
                                f"Sub Required: {row['Sub Required?']}, "
                                f"Reason: {row['Reason']}, "
                                f"Additional Comments: {row['Additional comments']}"),
                'start': {
                    'dateTime': row['Start Time'].isoformat(),
                    'timeZone': 'America/New_York',
                },
                'end': {
                    'dateTime': row['End Time'].isoformat(),
                    'timeZone': 'America/New_York',
                },
                'reminders': {
                    'useDefault': True
                }
            }
            
            # Initialize approval_id
            approval_id = str(row['Approval ID'])
            
            if approval_id in existing_events and (row['Status'] == 'Approved' or row['Status'] == 'Pending'):
                # Path 1: Event exists - update it if needed
                try:    
                    success = self.update_calendar_event(
                        existing_events[approval_id]['event_id'],
                        event
                    )
                    if success:
                        existing_count += 1
                        print(f"Existing calendar event for {full_name}")
                except Exception as e:
                    print(f"Error updating calendar event for {full_name}: {str(e)}")
            elif approval_id in existing_events and (row['Status'] == 'Rejected' or row['Status'] == 'Revoked'):
                # Path 2: Event exists - delete it
                try:
                    self.service.events().delete(
                        calendarId=self.calendar_id,
                        eventId=existing_events[approval_id]['event_id'],
                        sendUpdates='all'
                    ).execute()
                    deleted_count += 1
                    print(f"Deleted calendar event for {full_name}")
                    try: 
                        self.load_and_save_deleted_events(approval_id)
                        print(f"Added event to deleted events: {approval_id}")
                    except Exception as e:
                        print(f"Error adding event to deleted events: {str(e)}")
                except Exception as e:
                    print(f"Error deleting event: {str(e)}")
            elif approval_id in self.deleted_set: 
                # Path 3: Event was previously deleted - ignore it
                ignored_count += 1  
            else:
                # Path 4: Completely new event - create a new event
                try:
                    self.service.events().insert(
                        calendarId=self.calendar_id,
                        body=event,
                        sendUpdates='all'
                    ).execute()
                    created_count += 1
                    print(f"Created new calendar event for {full_name}")
                except Exception as e:
                    print(f"Error creating calendar event for {full_name}: {str(e)}")
        
        print(f"\nSummary:")
        print(f"Created {created_count} new events")
        print(f"Updated {existing_count} existing events")
        print(f"Deleted {deleted_count} existing events")
        print(f"Ignored {ignored_count} previously deleted events")

def main():
    try:
        logging.info("Starting leave request calendar update...")
        
        # Information for the Google Calendar API
        SERVICE_ACCOUNT_FILE = r'Path\to\file'
        CALENDAR_ID = 'primary'
        DOWNLOAD_DIR = os.path.expanduser("~/Downloads")
        
        # Create calendar manager
        calendar_manager = LeaveRequestCalendar(SERVICE_ACCOUNT_FILE, CALENDAR_ID)
        
        # Download Excel file
        logging.info("Downloading Excel file...")
        calendar_manager.download_excel()
        
        # Get latest Excel file
        excel_file = calendar_manager.get_latest_excel_file(DOWNLOAD_DIR)
        logging.info(f"Found Excel file: {excel_file}")
        
        # Create calendar events
        calendar_manager.create_calendar_events(excel_file)
        
        logging.info("Calendar update completed successfully")
        
    except Exception as e:
        print(f"Error: {str(e)}")

if __name__ == "__main__":
    main()