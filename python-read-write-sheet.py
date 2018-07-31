# Install the smartsheet sdk with the command: pip install smartsheet-python-sdk
import pandas
import smartsheet
import logging
import os.path

# TODO: Set your API access token here, or leave as None and set as environment variable "SMARTSHEET_ACCESS_TOKEN"
access_token = "9ffzzdb2pammh8gqpp8w5w6ucy"

_dir = os.path.dirname(os.path.abspath(__file__))

# # The API identifies columns by Id, but it's more convenient to refer to column names. Store a map here
# xl_column_map = {} #map for the salesforce excel sheet
# ss_column_map = {} #map for client list smartsheet
# pl_column_map = {} #map for pipeline list smartsheet

#Store Master Client List and Pipeline sheet ID of Interest in a variable
CLIENT_LIST_ID = 8950161956202372
PIPELINE_LIST_ID = 8257272599078788

# Helper function to find cell in a row
def get_cell_by_column_name(map_obj, row, column_name):
    column_id = map_obj[column_name]
    return row.get_column(column_id)

def create_map_from_columns(smartsheet_obj):
    map_obj = {}
    for column in smartsheet_obj.columns:
        map_obj[column.title] = column.id
    return map_obj

print("Starting ...")

# Initialize client
ss = smartsheet.Smartsheet(access_token)
# Make sure we don't miss any error
ss.errors_as_exceptions(True)

# Log all calls
logging.basicConfig(filename='rwsheet.log', level=logging.INFO)

# # Import the excel sheet into a smartsheet object (this creates a smartsheet to be deleted later)
result = ss.Sheets.import_xlsx_sheet(_dir + '/Fake-Salesforce-Clients.xlsx', header_row_index=0)
salesforce_data = ss.Sheets.get_sheet(result.data.id)

# Load entire client sheet and pipeline sheet
client_sheet = ss.Sheets.get_sheet(CLIENT_LIST_ID)
pipeline_sheet = ss.Sheets.get_sheet(PIPELINE_LIST_ID)

print ("Loaded " + str(len(client_sheet.rows)) + " rows from sheet: " + client_sheet.name)

# The API identifies columns by Id, but it's more convenient to refer to column names. Store a map here
# Build column map for later reference - translates column names to column id
xl_column_map = create_map_from_columns(salesforce_data) #map for the salesforce excel sheet
ss_column_map = create_map_from_columns(client_sheet) #map for client list smartsheet
pl_column_map = create_map_from_columns(pipeline_sheet) #map for pipeline list smartsheet

def check_if_opp_ID_exists_in_sheet(opp_id, smartsheet_obj, map_obj):
    already_exists = False
    for ss_row in smartsheet_obj.rows:     
        ss_opp_ID_cell = get_cell_by_column_name(map_obj, ss_row, "OppID")
        ss_opp_ID_value = ss_opp_ID_cell.display_value
        if (ss_opp_ID_value == opp_id):
            already_exists = True
            break
    return already_exists

# Accumulate rows needing update here
AddedRowIDs = []

for xl_row in salesforce_data.rows:
    xl_status_cell = get_cell_by_column_name(xl_column_map, xl_row, "Status")
    xl_status_value = xl_status_cell.display_value
    
    xl_opp_ID_cell = get_cell_by_column_name(xl_column_map, xl_row, "OppID")
    xl_opp_ID_value = xl_opp_ID_cell.display_value
    
    is_in_client_list = check_if_opp_ID_exists_in_sheet(xl_opp_ID_value, client_sheet, ss_column_map) 
    #check if this oppID is already present in the Client List Smartsheet
#     already_exists = False
#     for ss_row in client_sheet.rows:     
#         ss_opp_ID_cell = get_cell_by_column_name(ss_column_map, ss_row, "OppID")
#         ss_opp_ID_value = ss_opp_ID_cell.display_value
#         if (ss_opp_ID_value == xl_opp_ID_value):
#             already_exists = True
#             break
            
    if (xl_status_value == "Closed" and not is_in_client_list):          
            AddedRowIDs.append(xl_row.id) 
    
    
# Finally, write updated cells back to Smartsheet
if AddedRowIDs:
    print("Writing " + str(len(AddedRowIDs)) + " rows back to sheet id " + str(client_sheet.id))
    response = ss.Sheets.copy_rows(
      salesforce_data.id,               # sheet_id of rows to be copied
      ss.models.CopyOrMoveRowDirective({
        'row_ids': AddedRowIDs,
        'to': ss.models.CopyOrMoveRowDestination({
          'sheet_id': CLIENT_LIST_ID
        })
      })
    )
    
    email = ss.models.MultiRowEmail({
        #hard-coded, but this should pull in the value in the email column
        "sendTo": [{
            "email": "tallen@mdsol.com" 
        }],
        "subject": "Action Required: Payments Data Needed",
        "message": "Hi Travis. New opportunities have appeared in the Payments List.  Please update the missing fields.  Payments Team",
        "ccMe": False,
        "includeAttachments": False,
        "includeDiscussions": False
    })
    email.row_ids = AddedRowIDs

    # Send rows via email
    email_response = ss.Sheets.send_rows(
      salesforce_data.id,       # sheet_id
      email)
    print(email_response)

else:
    print("No updates required")
        
## Delete the Salesforce sheet that you've created:
ss.Sheets.delete_sheet(
  salesforce_data.id) 
print ("Done")



