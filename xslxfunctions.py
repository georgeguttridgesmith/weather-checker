import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from pprint import pprint
from os.path import join
from datetime import date

date_string = str(date.today())



def read_xlsx_file(filepath):
    wb = openpyxl.load_workbook(filepath)
    ws = wb.active
    
    # Create a list to hold all the dictionaries
    data = []
    
    # Get the header row and use it as keys for each row
    keys = [cell.value for cell in ws[1]]
    
    # Loop over the rows starting from the second row
    for row in ws.iter_rows(min_row=2, values_only=True):
        # Create a dictionary using the keys and row values
        row_dict = dict(zip(keys, row))
        # Add the dictionary to the list
        data.append(row_dict)
    
    return data

pprint(read_xlsx_file())



def create_xlsx_workbook(filename_prefix='', sheet_names=[], xlsx_folder_path=''):
    # Create file path
    xlsx_wb_name = f"{date_string}{filename_prefix}.xlsx"
    xlsx_file_path_full = join(xlsx_folder_path, xlsx_wb_name)

    # Create workbook and sheets
    wb = Workbook()
    for sheet_name in sheet_names:
        wb.create_sheet(sheet_name)

    # Save workbook to file path
    wb.save(xlsx_file_path_full)

    return [xlsx_file_path_full, xlsx_wb_name]

def write_xlsx_worksheet(list_of_dict, workbook_name='', sheet_name='', xlsx_file_loc=''):
    # Open workbook and worksheet
    wb = load_workbook(xlsx_file_loc)
    ws = wb[sheet_name]

    # Write headings
    headings = list(list_of_dict[0].keys())
    ws.append(headings)

    # Write data rows
    for cus_dict in list_of_dict:
        ship_data = list(cus_dict.values())
        ws.append(ship_data)

    # Make headings bold
    for col in range(1, len(headings) + 1):
        ws.cell(row=1, column=col).font = Font(bold=True)

    # Remove empty sheets
    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        if sheet.max_row == 0 and sheet.max_column == 0:
            wb.remove(sheet)

    # Save workbook
    wb.save(xlsx_file_loc)
