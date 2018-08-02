# Install the smartsheet sdk with the command: pip install smartsheet-python-sdk
import pandas
import smartsheet
import logging
import os.path
import json
from datetime import datetime

# TODO: Set your API access token here, or leave as None and set as environment variable "SMARTSHEET_ACCESS_TOKEN"
access_token = "y65rugwjfgncy3h9an5vjtn9fv"

_dir = os.path.dirname(os.path.abspath(__file__))

#Store Master Client List and Pipeline sheet ID of Interest in a variable
CLIENT_LIST_ID = 8950161956202372
PIPELINE_LIST_ID = 8257272599078788
STATS_SHEET_ID = 166602185435012

DASHBOARD_ID = 5619544673806212

# Initialize client
ss = smartsheet.Smartsheet(access_token)
# Make sure we don't miss any error
ss.errors_as_exceptions(True)

# Helper function to find cell in a row
def get_cell_by_column_name(map_obj, row, column_name):
    column_id = map_obj[column_name]
    return row.get_column(column_id)

def get_cell_value_by_column_name(map_obj, row, column_name):
    column_id = map_obj[column_name]
    cell = row.get_column(column_id)
    return cell.value

def create_map_from_columns(smartsheet_obj):
    map_obj = {}
    for column in smartsheet_obj.columns:
        map_obj[column.title] = column.id
    return map_obj

# print("Starting ...")


# Log all calls
logging.basicConfig(filename='rwsheet.log', level=logging.INFO)

# # Load entire client sheet and pipeline sheet
client_sheet = ss.Sheets.get_sheet(CLIENT_LIST_ID)
stats_sheet = ss.Sheets.get_sheet(STATS_SHEET_ID)
# pipeline_sheet = ss.Sheets.get_sheet(PIPELINE_LIST_ID)

# print ("Loaded " + str(len(client_sheet.rows)) + " rows from sheet: " + client_sheet.name)

# The API identifies columns by Id, but it's more convenient to refer to column names. Store a map here
# Build column map for later reference - translates column names to column id
stats_column_map = create_map_from_columns(stats_sheet) #map for the salesforce excel sheet
clients_column_map = create_map_from_columns(client_sheet) #map for client list smartsheet
# pl_column_map = create_map_from_columns(pipeline_sheet) #map for pipeline list smartsheet


#GO THROUGH EACH STUDY
    #GO THROUGH EACH UPCOMING MONTH
        #IF CONTRACT STARTED BEFORE THAT MONTH
            # CALCULATE PAYMENTS FOR THAT MONTH FROM THIS CLIENT (# Sites * 1 if Monthly Frequency, * 2 if bi-monthly etc)
            # ADD THIS CALCULATION TO THE TOTAL FOR THE MONTH
signed_clients_row = None
pipeline_row = None

for stats_row in stats_sheet.rows:
    first_column_val = get_cell_value_by_column_name(stats_column_map, stats_row, "Source List")
    if (first_column_val == "Signed Clients"):
        signed_clients_row = stats_row
    elif (first_column_val == "Pipeline"):
        pipeline_row = stats_row


new_clients_row = ss.models.Row()
new_clients_row.id = signed_clients_row.id

for month_column in stats_sheet.columns:
    if (month_column.index == 0): #skip the first column, as this has no data
        continue
    monthly_payments = 0
    column_date = datetime.strptime(month_column.title, '%b %Y')
    print(type(column_date))
    
    for client_study_row in client_sheet.rows:
        num_sites = get_cell_value_by_column_name(clients_column_map, client_study_row, "# of Sites")
        start_date = get_cell_value_by_column_name(clients_column_map, client_study_row, "Start Date")
        print(start_date)
        ## THIS START DATE CONVERSION IS BROKEN, MAKING EVERYTHING 01-01-18
        start_date = datetime.strptime(start_date, '%Y-%m-%d')
        payment_frequency = get_cell_value_by_column_name(clients_column_map, client_study_row, "Payment Frequency")
        payments_per_month = 1 if payment_frequency == "Monthly" else 2
        month_study_payments = num_sites * payments_per_month
        
        if (start_date < column_date):
            print(start_date)
            monthly_payments += month_study_payments
    print("Total for Month....")
    print(monthly_payments)
    
    month_cell = ss.models.Cell()    
    month_cell.column_id = month_column.id
    month_cell.value = monthly_payments
    
    new_clients_row.cells.append(month_cell)

updated_row = ss.Sheets.update_rows(
  STATS_SHEET_ID,      # sheet_id
  [new_clients_row])
    
# sight = ss.Sights.get_sight(
#   DASHBOARD_ID)     # sightId

# sight_dict = sight.to_dict()

# sight_dict['widgets'][0]['title'] = 'New Title'

# response = ss.Sights.set_publish_status(
#   DASHBOARD_ID,       # sight_id
#   ss.models.SightPublish({
#     'read_only_full_enabled': True
#   })
# )

# print(response)

