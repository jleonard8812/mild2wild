import requests
import login
import re
import gspread
from google.oauth2.service_account import Credentials
from gspread.exceptions import APIError
from datetime import datetime

# API credentials
api_url = "https://mild2wild.arcticres.com/api/rest"
api_username = login.username
api_password = login.password

# Endpoint for reservations
reservations_url = f"{api_url}/reservation?query=activity.status+IN+%28%27unfinished%27%2C+%27finished%27%2C+%27over%27%29+AND+trip.canceled+%3D+false+AND+activity.start.datetimerelative+APPLY%28%27operator%27%2C%27on%27%2C%27count%27%2C%270%27%2C%27units%27%2C%27day%27%2C%27direction%27%2C%27future%27%29+AND+activity.businessgroupid.businessgroupcondition+APPLY%28%27operator%27%2C%27is-or-within%27%2C%27value%27%2C%2727%27%29+AND+allcomponents+LIKE+%27%25Photos%25%27"

# Fetch reservations data
response = requests.get(reservations_url, auth=(api_username, api_password))
response.raise_for_status()
reservations_data = response.json()

# Extract activity IDs
activity_ids = [entry['activityid'] for entry in reservations_data['entries']]

# Function to get customer information from activity ID
def get_customer_info(activity_id):
    activity_url = f"{api_url}/activity/{activity_id}"
    response = requests.get(activity_url, auth=(api_username, api_password))
    response.raise_for_status()
    activity_data = response.json()
    
    person_data = activity_data.get('person', {})
    customer_name = f"{person_data.get('namefirst', 'N/A')} {person_data.get('namelast', 'N/A')}"
    customer_email = person_data.get('emailaddresses', [{}])[0].get('emailaddress', 'N/A')
    customer_phone = person_data.get('phonenumbers', [{}])[0].get('phonenumber', 'N/A')
    
    # Extract Trip Type and Trip Time from description
    trip_type = trip_time = 'N/A'
    if 'invoice' in activity_data and 'groups' in activity_data['invoice']:
        for group in activity_data['invoice']['groups']:
            for item in group.get('items', []):
                description = item.get('description', '')
                trip_type_match = re.search(r"1/2 Day - Kayak|1/2 Day|1/2Day-Premium|3/4-Day|1/4 Day", description)
                trip_time_match = re.search(r"\d{1,2}:\d{2}\s*(AM|PM)", description)
                if trip_type_match:
                    trip_type = trip_type_match.group(0)
                if trip_time_match:
                    trip_time = trip_time_match.group(0)
    
    return customer_name, customer_email, customer_phone, trip_type, trip_time

# Google Sheets setup
# Path to the downloaded JSON file
creds_file = 'g_creds.json'  # Use the correct path

# Define the scopes
scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

# Load the credentials
creds = Credentials.from_service_account_file(creds_file, scopes=scopes)
client = gspread.authorize(creds)

# Open the existing spreadsheet by ID
spreadsheet_id = '1uBYSAal74K-tTWzcTp2QkdEQDSWO3Rf07ptAhI35OUM'  # Use your spreadsheet ID
try:
    spreadsheet = client.open_by_key(spreadsheet_id)
    print(f"Opened existing spreadsheet: {spreadsheet.url}")

    # Select the first sheet
    sheet = spreadsheet.sheet1

    # Clear existing data in the sheet
    sheet.clear()
    print("Cleared existing data.")

    # Set up the header row if not already set
    existing_headers = sheet.row_values(1)
    header = ['Name', 'Email', 'Phone', 'Trip Type', 'Trip Time', 'Guide']
    if existing_headers != header:
        sheet.clear()  # Optional: Clear existing content if headers are not correct
        sheet.append_row(header)

    # Function to convert time string to a sortable datetime object
    def convert_time(trip_time):
        if trip_time == 'N/A':
            return datetime.min
        if not trip_time[-2] == ' ':
            trip_time = trip_time[:-2] + ' ' + trip_time[-2:]
        return datetime.strptime(trip_time, '%I:%M %p')

    # Collect customer information
    customer_data = []
    for activity_id in activity_ids:
        name, email, phone, trip_type, trip_time = get_customer_info(activity_id)
        customer_data.append((name, email, phone, trip_type, trip_time))

    # Sort the data by Trip Time
    customer_data.sort(key=lambda x: convert_time(x[4]))  # Assuming trip_time is the 5th element in the tuple

    # Append sorted data to the sheet
    for row in customer_data:
        sheet.append_row(list(row))

    print("Data uploaded successfully.")

except APIError as e:
    print(f"APIError: {e}")
except Exception as e:
    print(f"An error occurred: {e}")
