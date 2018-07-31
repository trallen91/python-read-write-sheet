# Install the smartsheet sdk with the command: pip install smartsheet-python-sdk
import pandas
import smartsheet
import logging
import os.path

# TODO: Set your API access token here, or leave as None and set as environment variable "SMARTSHEET_ACCESS_TOKEN"
access_token = "9ffzzdb2pammh8gqpp8w5w6ucy"

_dir = os.path.dirname(os.path.abspath(__file__))

# The API identifies columns by Id, but it's more convenient to refer to column names. Store a map here
xl_column_map = {}
ss_column_map = {}

#Store Master List sheet ID of Interest in a variable
SMARTSHEET_ID = 8950161956202372

# Helper function to find cell in a row
def get_cell_by_column_name(map_obj, row, column_name):
    column_id = map_obj[column_name]
    return row.get_column(column_id)

print("Starting ...")

# Initialize client
ss = smartsheet.Smartsheet(access_token)
# Make sure we don't miss any error
ss.errors_as_exceptions(True)

# Log all calls
logging.basicConfig(filename='rwsheet.log', level=logging.INFO)

# # Import the excel sheet
result = ss.Sheets.import_xlsx_sheet(_dir + '/Fake-Salesforce-Clients.xlsx', header_row_index=0)
salesforce_data = ss.Sheets.get_sheet(result.data.id)

# Load entire sheet
sheet = ss.Sheets.get_sheet(SMARTSHEET_ID)

print ("Loaded " + str(len(sheet.rows)) + " rows from sheet: " + sheet.name)

# Build column map for later reference - translates column names to column id
for ss_column in sheet.columns:
    ss_column_map[ss_column.title] = ss_column.id

for xl_column in salesforce_data.columns:
    xl_column_map[xl_column.title] = xl_column.id
            
# Accumulate rows needing update here
AddedRowIDs = []

for xl_row in salesforce_data.rows:
    xl_status_cell = get_cell_by_column_name(xl_column_map, xl_row, "Status")
    xl_status_value = xl_status_cell.display_value
    
    xl_opp_ID_cell = get_cell_by_column_name(xl_column_map, xl_row, "OppID")
    xl_opp_ID_value = xl_opp_ID_cell.display_value
    
    #check if this oppID is already present in the Smartsheet
    already_exists = False
    for ss_row in sheet.rows:     
        ss_opp_ID_cell = get_cell_by_column_name(ss_column_map, ss_row, "OppID")
        ss_opp_ID_value = ss_opp_ID_cell.display_value
        if (ss_opp_ID_value == xl_opp_ID_value):
            already_exists = True
            break
            
    if (xl_status_value == "Closed" and not already_exists):          
            AddedRowIDs.append(xl_row.id) 
    
    
# Finally, write updated cells back to Smartsheet
if AddedRowIDs:
    print("Writing " + str(len(AddedRowIDs)) + " rows back to sheet id " + str(sheet.id))
    response = ss.Sheets.copy_rows(
      salesforce_data.id,               # sheet_id of rows to be copied
      ss.models.CopyOrMoveRowDirective({
        'row_ids': AddedRowIDs,
        'to': ss.models.CopyOrMoveRowDestination({
          'sheet_id': SMARTSHEET_ID
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



