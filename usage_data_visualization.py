import json
import requests
from datetime import datetime
import gspread
from google.oauth2 import service_account
from googleapiclient.discovery import build
from rich import print

# Setup values
client_name = "Client Name"
start_date = "20230601"
end_date = "20230831"
api_key = "Org API Key"
spreadsheet_title = f'{client_name} - Usage Data {start_date} - {end_date}'
header = {'Authorization': f'Bearer {api_key}'}

# 1. Make an API call to gather usage metrics
url = f"https://{api_key}@api-dev.headspin.io/v1/usage"
params_total_time = {
    "start_date": start_date,
    "end_date": end_date,
    "metric": "total_time",
    "timezone": "US/Pacific",
    "bin_size": "month"
}

params_device_count = {
    "start_date": start_date,
    "end_date": end_date,
    "metric": "device_count",
    "timezone": "US/Pacific",
    "bin_size": "month"
}

print("Making API call to gather usage metrics...")
total_time_response = requests.get(url, params=params_total_time, headers=header)
total_time_data = total_time_response.json()

device_count_response = requests.get(url, params=params_device_count, headers=header)
device_count_data = device_count_response.json()

print(total_time_data)
print(device_count_data)

print("Merging data...")
merged_data = {'report': []}
for num in range(len(total_time_data['report'])):
    merged_data['report'].append({
        'datetime': total_time_data['report'][num]['datetime'],
        'total_time': total_time_data['report'][num]['total_time'],
        'unit': total_time_data['report'][num]['unit'],
        'device_count': device_count_data['report'][num]['device_count']
    })

print(merged_data)


# 2. Write the data to Google Sheets
scope = [
    'https://spreadsheets.google.com/feeds',
    'https://www.googleapis.com/auth/drive.file',
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/spreadsheets'
]

creds = service_account.Credentials.from_service_account_file('usage-visualization-8960ba280706.json', scopes=scope)
client = gspread.authorize(creds)
service = build('sheets', 'v4', credentials=creds)  # Add this for the googleapiclient

print("Writing data to Google Sheets...")
try:
    # Try to open the existing spreadsheet
    spreadsheet = client.open(spreadsheet_title)
except gspread.SpreadsheetNotFound:
    # If not found, create a new one
    spreadsheet = client.create(spreadsheet_title)
    # And share it with your email if needed (otherwise, only the service account can access it)
    spreadsheet.share('jose@headspin.io', perm_type='user', role='writer')


# Create or fetch "Usage Data" sheet
number_of_rows = len(merged_data['report']) + 1
number_of_columns = len(merged_data['report'][0]) + 1

try:
    usage_data_sheet = spreadsheet.worksheet("Usage Data")
    usage_data_sheet.clear()
except gspread.exceptions.WorksheetNotFound:
    usage_data_sheet = spreadsheet.add_worksheet(title="Usage Data", rows=number_of_rows, cols=number_of_columns)

usage_data_sheet.append_row(['Month', 'Total Time', 'Unit', 'Device Count'])

for entry in merged_data['report']:
    month_name = datetime.strptime(entry['datetime'], '%Y-%m-%d').strftime('%B')
    usage_data_sheet.append_row([month_name, float(entry['total_time']), entry['unit'], int(entry['device_count'])])

# 3. Create sheet for combo graph
if "Device and Platform Usage" not in [ws.title for ws in spreadsheet.worksheets()]:
    spreadsheet.add_worksheet(title="Device and Platform Usage", rows="100", cols="20")

chart_sheet = spreadsheet.worksheet('Device and Platform Usage')

number_of_rows = len(merged_data['report']) + 1  # +1 accounts for the header row

# Define the chart request
chart_data = {
    "requests": [
        {
            "addChart": {
                "chart": {
                    "spec": {
                        "title": "Device Count and Platform Usage in Hours",
                        "basicChart": {
                            "chartType": "COMBO",
                            "legendPosition": "BOTTOM_LEGEND",
                            "axis": [
                                {"position": "BOTTOM_AXIS", "title": "Month"},
                                {"position": "LEFT_AXIS", "title": "Device Count"},
                                {"position": "RIGHT_AXIS", "title": "Total Time (hours)"}
                            ],
                            "domains": [
                            {
                                "domain": {
                                    "sourceRange": {
                                        "sources": [
                                            {
                                                "sheetId": usage_data_sheet._properties['sheetId'],
                                                "startRowIndex": 0,
                                                "endRowIndex": number_of_rows,
                                                "startColumnIndex": 0,
                                                "endColumnIndex": 1
                                            }
                                        ]
                                    }
                                }
                            }
                            ],
                            "series": [
                                {
                                    "series": {
                                        "sourceRange": {
                                            "sources": [
                                                {
                                                    "sheetId": usage_data_sheet._properties['sheetId'],
                                                    "startRowIndex": 0,
                                                    "endRowIndex": number_of_rows,
                                                    "startColumnIndex": 1,
                                                    "endColumnIndex": 2
                                                }
                                            ]
                                        }
                                    },
                                    "targetAxis": "RIGHT_AXIS",
                                    "type": "LINE"
                                },
                                {
                                    "series": {
                                        "sourceRange": {
                                            "sources": [
                                                {
                                                    "sheetId": usage_data_sheet._properties['sheetId'],
                                                    "startRowIndex": 0,
                                                    "endRowIndex": number_of_rows,
                                                    "startColumnIndex": 3,
                                                    "endColumnIndex": 4
                                                }
                                            ]
                                        }
                                    },
                                    "targetAxis": "LEFT_AXIS",
                                    "type": "COLUMN"
                                }
                            ],
                            "headerCount": 1,
                            "stackedType": "STACKED"
                        }
                    },
                    "position": {
                        "overlayPosition": {
                            "anchorCell": {"sheetId": chart_sheet.id, "rowIndex": 5, "columnIndex": 0},
                            "widthPixels": 600,
                            "heightPixels": 400
                        }
                    }
                }
            }
        }
    ]
}


chart_update = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet.id, body=chart_data).execute()
