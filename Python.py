import openpyxl
import os
from openpyxl import Workbook
import pandas as pd
from datetime import datetime, timedelta
from openpyxl.styles import Font, Alignment 
from openpyxl.styles import PatternFill

#Get todays date and format to string 4-digit year, 2-digit month, 2-digit day
today = (datetime.now().strftime('%Y%m%d'))

#Define directory where to find files to read
file_death = r'C:\Users\Rudolph Van Rensburg\Desktop\KHUSA\Clarissa Summary Report\Data\Death.xlsx'
file_disablity = r'C:\Users\Rudolph Van Rensburg\Desktop\KHUSA\Clarissa Summary Report\Data\Disability.xlsx'
file_underwriting = r'C:\Users\Rudolph Van Rensburg\Desktop\KHUSA\Clarissa Summary Report\Data\Medical Underwriting.xlsx'

#Read File
df_death = pd.read_excel(file_death)
df_disability = pd.read_excel(file_disablity)
df_underwriting = pd.read_excel(file_underwriting)

# Clean column names
df_death.columns = df_death.columns.str.strip()
df_disability.columns = df_disability.columns.str.strip()
df_underwriting.columns = df_underwriting.columns.str.strip()

#Count cell contents
funeral_Claim = df_death['Claim type'].astype(str).str.contains('funeral', case=False, na=False).sum()
gla_Claim = df_death['Claim type'].astype(str).str.contains('GLA', case=False, na=False).sum()
finalised_Funeral_Claim = df_death[df_death['Status of the Claim'].astype(str).str.contains('Paid', case=False, na=False)]['Claim type'].astype(str).str.contains('funeral', case=False, na=False).sum()
finalised_Gla_Claim = df_death[df_death['Status of the Claim'].astype(str).str.contains('Paid', case=False, na=False)]['Claim type'].astype(str).str.contains('GLA', case=False, na=False).sum()
disability_Claim = df_disability['Member'].count()
gla_Underwriting_Requested  = df_underwriting[df_underwriting['Benefit'].astype(str).str.contains('GLA', case=False, na=False) & df_underwriting['Decision'].isna()].shape[0]
disability_Underwriting_Requested  = df_underwriting[~df_underwriting['Benefit'].astype(str).str.contains('GLA', case=False, na=False) & df_underwriting['Decision'].isna()].shape[0]
gla_Underwriting_Decisioned  = df_underwriting[df_underwriting['Benefit'].astype(str).str.contains('GLA', case=False, na=False) & df_underwriting['Decision'].notna()]['Decision'].count()
disability_Underwriting_Decisioned  = df_underwriting[~df_underwriting['Benefit'].astype(str).str.contains('GLA', case=False, na=False) & df_underwriting['Decision'].notna()]['Decision'].count()

# Sheet names
clients = ['Client 1', 'Client 2', 'Client 3']

#Define Directory
file_directory = r'C:\Users\Rudolph Van Rensburg\Desktop\KHUSA\Clarissa Summary Report'

#Define File Name
filename = f'{today} - Claims & Underwriting Summarized Reporting.xlsx'

#Create full path
file_path = os.path.join(file_directory, filename)

# Create workbook and remove default sheet
workbook = Workbook()
default_sheet = workbook.active
workbook.remove(default_sheet)

#Data template
def get_sheet_data():
    return [
        ['Claims','', '','','','','Underwriting','','','',''],
        ['', 'Funeral', 'Death', 'Disability', 'Notes', '', '', 'GLA', 'Disability', 'Notes'],
        ['Received', funeral_Claim, gla_Claim, disability_Claim, '', '', 'Requested', gla_Underwriting_Requested, disability_Underwriting_Requested, ''],
        ['Finalised', finalised_Funeral_Claim, finalised_Gla_Claim, 'TBC', '', '', 'Decisioned', gla_Underwriting_Decisioned, disability_Underwriting_Decisioned, ''],
        [''],
        ['ADDITIONAL FEEDBACK/NOTES']
    ]

# Loop through each client and create formatted sheets
for client in clients:
    sheet = workbook.create_sheet(title=client)

    # Add data
    for row in get_sheet_data():
        sheet.append(row)

    # Merge across columns
    sheet.merge_cells('A1:E1')
    sheet.merge_cells('G1:J1')
    sheet.merge_cells('A6:J6')

    # Format the "Claims" cell
    cell = sheet['A1']
    cell.font = Font(size=14, bold=True, color='FFFFFF')
    cell.alignment = Alignment(horizontal='center', vertical='center')
    cell.fill = PatternFill(start_color="006666", end_color="006666", fill_type="solid")

    # Format the "Underwriting" cell
    cell = sheet['G1']
    cell.font = Font(size=14, bold=True, color='FFFFFF')
    cell.alignment = Alignment(horizontal='center', vertical='center')
    cell.fill = PatternFill(start_color="006666", end_color="006666", fill_type="solid")

    # Format the "ADDITIONAL FEEDBACK/NOTES" cell
    cell = sheet['A6']
    cell.font = Font(size=14, bold=True, color='FFFFFF')
    cell.alignment = Alignment(horizontal='center', vertical='center')
    cell.fill = PatternFill(start_color="006666", end_color="006666", fill_type="solid")

#Save Workbook
workbook.save(file_path)
print("file Saved", today)