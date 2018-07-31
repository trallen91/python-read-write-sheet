# Install the smartsheet sdk with the command: pip install smartsheet-python-sdk
import pandas
import smartsheet
import logging
import os.path

# TODO: Set your API access token here, or leave as None and set as environment variable "SMARTSHEET_ACCESS_TOKEN"
access_token = "9ffzzdb2pammh8gqpp8w5w6ucy"

_dir = os.path.dirname(os.path.abspath(__file__))

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
AddToClientList = []
AddToPipelineList = []

for xl_row in salesforce_data.rows:
    xl_status_cell = get_cell_by_column_name(xl_column_map, xl_row, "Status")
    xl_status_value = xl_status_cell.display_value
    
    xl_opp_ID_cell = get_cell_by_column_name(xl_column_map, xl_row, "OppID")
    xl_opp_ID_value = xl_opp_ID_cell.display_value
    
    is_in_client_list = check_if_opp_ID_exists_in_sheet(xl_opp_ID_value, client_sheet, ss_column_map)     
    is_in_pipeline_list = check_if_opp_ID_exists_in_sheet(xl_opp_ID_value, pipeline_sheet, pl_column_map)
    # IF IT IS A CLOSED OPPORTUNITY
    # Check if it is in the pipeline list
        # If Yes: Copy row from Pipeline List to Client List, delete from Pipeline List 
        # If No: Check if it is in Client List, then add if not 
    if (xl_status_value == "Closed"):          
        if (is_in_client_list):
            continue
        elif (not is_in_pipeline_list and not is_in_client_list):
            AddToClientList.append(xl_row.id)
#         elif (is_in_pipeline_list and not is_in_client_list):
            #logic to delete from pipeline list and add that to client list
    else:
        if (is_in_pipeline_list):
            continue
        else:
            AddToPipelineList.append(xl_row.id)
    
    #IF IT IS AN UN-CLOSED OPPORTUNITY
    # Check if it is in the pipeline list
        # If Yes:  Do nothing
        # If No: Add to pipeline list


def copy_to_smartsheet_list(source_sheet, target_sheet_ID, rows_to_copy):
    print("Writing " + str(len(rows_to_copy)) + " rows back to sheet id " + str(target_sheet_ID))
    response = ss.Sheets.copy_rows(
        source_sheet.id,               # sheet_id of rows to be copied
        ss.models.CopyOrMoveRowDirective({
            'row_ids': rows_to_copy,
            'to': ss.models.CopyOrMoveRowDestination({
              'sheet_id': target_sheet_ID
            })
          })
    )
    
# Finally, write updated cells back to Smartsheet
if AddToClientList:
    copy_to_smartsheet_list(salesforce_data, CLIENT_LIST_ID, AddToClientList)
       
#     email = ss.models.MultiRowEmail({
#         #hard-coded, but this should pull in the value in the email column
#         "sendTo": [{
#             "email": "tallen@mdsol.com" 
#         }],
#         "subject": "Action Required: Payments Data Needed",
#         "message": "Hi Travis. New opportunities have appeared in the Payments List.  Please update the missing fields.  Payments Team",
#         "ccMe": False,
#         "includeAttachments": False,
#         "includeDiscussions": False
#     })
#     email.row_ids = AddToClientList

#     # Send rows via email
#     email_response = ss.Sheets.send_rows(
#       salesforce_data.id,       # sheet_id
#       email)
#     print(email_response)
if AddToPipelineList:
    copy_to_smartsheet_list(salesforce_data, PIPELINE_LIST_ID, AddToPipelineList)
else:
    print("No updates required")
        
## Delete the Salesforce sheet that you've created:
ss.Sheets.delete_sheet(
  salesforce_data.id) 
print ("Done")



