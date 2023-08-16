import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
import os
from OSGridConverter import grid2latlong
import io
import numpy as np
import pdfplumber
from office365.sharepoint.client_context import ClientContext;
from office365.runtime.auth.client_credential import ClientCredential
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.files.file import File

# Set up link to sharepoint
site_url = 'https://companyName.sharepoint.com/teams/sharepointName/'
read_folder_URL = '/teams/sharepointName/Shared Documents/Structural Report Examples/'
output_folder_URL ='/teams/sharepointName/Shared Documents/Structural Report Output/'

client_id = #### enter key from share point ####
client_secret =  #### enter secret from share point ####
client_credentials = ClientCredential(client_id, client_secret)

ctx = ClientContext(site_url).with_credentials(client_credentials)

read_folder = ctx.web.get_folder_by_server_relative_url(read_folder_URL)
output_folder = ctx.web.get_folder_by_server_relative_url(output_folder_URL)

# Retrieve the files within the folder
files = read_folder.files
ctx.load(files)
ctx.execute_query()

################## Information ##################
start_row = "Examination Type:"
end_row = "Section A:"
columns=['Examination Type: ', 'NR ID: ', 'Exam Date: ', 'Area: ', 'BRS: ', 'OS Ref: ', 'Structure Name: ', 'Type:', 'Exam ID: ', 'Route: ', 'Complete Exam: ']
information = pd.DataFrame(columns=columns)
id = 0
table=[]

for file in files:
    get_information = pd.DataFrame(columns=columns)
    response = file.open_binary(ctx, file.serverRelativeUrl)
    bytes_file_obj = io.BytesIO()
    bytes_file_obj.write(response.content)
    bytes_file_obj.seek(0)
    page = 0

    with pdfplumber.open(bytes_file_obj) as pdf:
        current_page = pdf.pages[page]
        for row in current_page.extract_table():
            cleaned_row = [item for item in row if item != '' and item is not None]
            table.append(cleaned_row)

        for clmn in columns:
            for row in table:
                for item in row:
                    if clmn in item:
                        updated_item = item.replace(clmn, '')
                        row.remove(item)
                        get_information[clmn] = pd.Series([updated_item])

    get_information.insert(0, "ID: ", id)
    information = pd.concat([information, get_information], ignore_index=True)

    id += 1

# Move the "ID" column to the first position
information.insert(0, "ID: ", information.pop("ID: "))      
information['Latitude: '] = ''
information['Longitude: '] = ''

##add lat/long
# Loop through each row
for index, row in information.iterrows():
    os_reference = row['OS Ref: ']
    l = grid2latlong(os_reference)
    information.at[index, 'Latitude: '] = l.latitude
    information.at[index, 'Longitude: '] = l.longitude


         
################## SECTION A ##################
start_row = "DESCRIPTION"
end_row = "History of Live Significant Defects"
sectiona = pd.DataFrame()
id = 0

for file in files:
    end_row_index = None
    response = file.open_binary(ctx, file.serverRelativeUrl)
    bytes_file_obj = io.BytesIO()
    bytes_file_obj.write(response.content)
    bytes_file_obj.seek(0)
    page = 0

    while end_row_index is None:

        with pdfplumber.open(bytes_file_obj) as pdf:
            current_page = pdf.pages[page]
            table = current_page.extract_table()

            table = pd.DataFrame(table[1:], columns=table[0])
            table.replace('\n', ' ', regex=True, inplace=True)


        for index, row in enumerate(table.iterrows()):
            if start_row in row[1].values:
                start_row_index = index
            if end_row in row[1].values:
                end_row_index = index
                break
            else:
                end_row_index = None

        if end_row_index is not None:
            target_table = table.iloc[start_row_index+1:end_row_index]
        else:
            target_table = table.iloc[start_row_index+1:]

        target_table = target_table.dropna(axis=0, how="all")
        target_table = target_table.dropna(axis=1, how="all")
        target_table.columns = ['Item', 'Description', 'Location', 'Est. Cost £ +/- 20%', 'Priority Within', 'Quantity', 'Severity', 'Probability', 'Risk Score', 'Works Category']
        target_table.insert(0, "ID", id)

        sectiona = pd.concat([sectiona, target_table], ignore_index=True)
        page += 1

        id += 1
        


################## History ##################
start_row =  "History of Live Significant Defects"
end_row = "Engineers Notes"
history_df = pd.DataFrame()
id = 0

for file in files:
    end_row_index = None
    response = file.open_binary(ctx, file.serverRelativeUrl)
    bytes_file_obj = io.BytesIO()
    bytes_file_obj.write(response.content)
    bytes_file_obj.seek(0)
    page = 0

    while end_row_index is None:

        with pdfplumber.open(bytes_file_obj) as pdf:
            current_page = pdf.pages[page]
            table = current_page.extract_table()

            table = pd.DataFrame(table[1:], columns=table[0])
            table.replace('\n', ' ', regex=True, inplace=True)


        for index, row in enumerate(table.iterrows()):
            if start_row in row[1].values:
                start_row_index = index
            if end_row in row[1].values:
                end_row_index = index
                break
            else:
                end_row_index = None

        if end_row_index is not None:
            target_table = table.iloc[start_row_index+2:end_row_index]
        else:
            target_table = table.iloc[start_row_index+2:]
        
        target_table = table.iloc[start_row_index+2:end_row_index]
        target_table.replace(r'^\s*$', pd.NA, regex=True, inplace=True)
        target_table = target_table.dropna(axis=0, how="all")
        target_table = target_table.dropna(axis=1, how="all")

        if not target_table.empty:
            target_table.columns = ['No', 'Description', 'Location', 'Exam Date', 'Access Gained', 'Exam Type', 'Rec Raised', 'Risk Score', 'Access Required', 'Deterioration', 'Repaired', 'Flagged for Closure', 'Engineer Comments']
            target_table.insert(0, "ID", id)
            history_df = pd.concat([history_df, target_table], ignore_index=True)
            page += 1
        else:

          id -= 1
    id += 1
 
################## MASTERLIST ##################

id = 0
masterlist = pd.DataFrame(columns=['ID', 'Document Name'])

for file in files:
  response = file.open_binary(ctx, file.serverRelativeUrl)
  bytes_file_obj = io.BytesIO()
  bytes_file_obj.write(response.content)
  bytes_file_obj.seek(0)
  masterlist = pd.concat([masterlist, pd.DataFrame([[id, file.properties["Name"]]], columns=['ID', 'Document Name'])], ignore_index=True)
  id += 1

#Combine everything

#Export example_df
buffer = io.BytesIO()      # Create a buffer object
#master_df.to_excel(buffer, index=False) # Write the dataframe to the buffer
writer = pd.ExcelWriter(buffer, engine='xlsxwriter')

# Write each DataFrame to a separate sheet in the Excel file
masterlist.to_excel(writer, sheet_name='Masterlist', index=False)
information.to_excel(writer, sheet_name='Information', index=False)
sectiona.to_excel(writer, sheet_name='Section A', index=False)
history_df.to_excel(writer, sheet_name='History', index=False)

# Save and close the Excel writer
writer._save()
#writer.close()

# Exporting
# Retrieve the file content
file_content = buffer.getvalue()

buffer.seek(0)
file_content = buffer.read()

#Create output path
path = "/Outputs/Bridge_Report_Data.xlsx"

target_folder = output_folder
name = os.path.basename(path)
#Here is where we actually upload the file - using the "execute_query()" command again from ctx
target_file = target_folder.upload_file(name, file_content).execute_query()
print("File has been uploaded to url: {0}".format(target_file.serverRelativeUrl))

