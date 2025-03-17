import win32com.client as win32
import os

import tkinter as tk
from tkinter import filedialog

import shutil
import re

from datetime import datetime

# Function that opens file explorer and allows user to select file
def select_file():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename()
    return file_path

# Function copies all of the sheets from the report file into the same directory as .csv's
# Note: it is expected that the report_file has already been moved into a folder such as PC_{timestamp}
def sheets_to_csv(report_file, sens_name, timestamp):
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False
    workbook = excel.Workbooks.Open(report_file)

    directory = os.path.dirname(report_file)
    for sheet in workbook.Sheets:
        try:
            sheet_name = sheet.Name
            csv_filename = rf'{sens_name}_{timestamp}-{sheet_name}.csv'
            save_location = (rf'{directory}/{csv_filename}').replace("/", "\\")
            
            sheet.SaveAs(save_location, FileFormat=6)
            print(f"Duplicated sheet '{sheet_name}' successfully.") 
        except Exception as e:
            print(f"Failed to duplicate sheet '{sheet_name}': {e}") 
    
    workbook.Close(False)
    excel.Quit()

# Function copies sheets from the source path onto our new excel sheet for pivot cache creation
def copy_sheets(source_path, dest_path, sheet_names):
    # Start an instance of excel
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False

    # Open our source workbook and create our destination workbook
    source_wb = excel.Workbooks.Open(source_path)
    dest_wb = excel.Workbooks.Add()

    # Loop through each sheet name and copy it over 
    for sheet_name in sheet_names: 
        try: 
            src_sheet = source_wb.Sheets(sheet_name)

            # Copy the sheet to the end of the destination workbook 
            src_sheet.Copy(After=dest_wb.Sheets(dest_wb.Sheets.Count)) 
            dest_wb.Sheets(sheet_name).Visible = win32.constants.xlSheetHidden

            print(f"Copied sheet '{sheet_name}' successfully.") 
        except Exception as e: 
            print(f"Failed to copy sheet '{sheet_name}': {e}") 

    # Save and close the destination workbook, quit this instance of excel
    dest_wb.SaveAs(dest_path, FileFormat=51) 
    dest_wb.Close() 
    source_wb.Close(False)
    excel.Quit()
 
# Function gets the next available row in the spreadsheet, with padding (+4) for a new table location
def get_next_available_row(sheet, col=1):
    # If the excel file is empty return A1
    if sheet.Cells(1, 1).Value is None and sheet.UsedRange.Count == 1:
        return 1
    
    # Otherwise, find the last used row and give padding
    last_used_row = sheet.Cells(sheet.Rows.Count, col).End(win32.constants.xlUp).Row
    return last_used_row + 4

# Used to hide #!DIV0 errors and make them blank cells
# This is OK because we aren't expecting values in such cells anyways
def hide_errors(pivot_table):
    pivot_range = pivot_table.TableRange1
    formula = "=ISERROR(INDIRECT(ADDRESS(ROW(), COLUMN())))"

    cf = pivot_range.FormatConditions.Add(
        Type=win32.constants.xlExpression,
        Formula1=formula
    )

    cf.Font.Color = 16777215

# Entire class is essentially a structure and stores the information we need for any pivot table
# This is NOT fully dynamic -- only the required variabilities have been accounted for
# For example, only "sum" and "average" are valid, extra features must be implemented as needed
class Pivot_Def: 
    def __init__(self, source_sheet, pivot_rows, pivot_cols, pivot_data, agg_methods, pivot_filters, dest_sheet_pivot): 
        self.source_sheet = source_sheet # Name of the sheet that is the data source (copied sheet) 
        self.pivot_rows = pivot_rows # Row field(s) 
        self.pivot_cols = pivot_cols # Column field(s) 
        self.pivot_data = pivot_data # Data field(s)
        self.agg_methods = agg_methods # Aggregation method(s)
        self.pivot_filters = pivot_filters # Filter(s)
        self.dest_sheet_pivot = dest_sheet_pivot # Destination sheet for the pivot table 

class Pivot_Creator:
    def __init__(self, source_path, dest_path, sheet_names, wsColor):
        # Save our destination path and start an instance of excel
        self.dest_path = dest_path
        self.excel = win32.gencache.EnsureDispatch('Excel.Application')
        self.excel.Visible = False

        # Copy sheets that we need from our EnCompass report
        copy_sheet_names = [
            "Resource Annual",
            "Resource Annual Fuel",
            "Resource Annual Emissions",
            "Company Annual",
            "Company Annual Programs"
        ]

        copy_sheets(source_path, dest_path, copy_sheet_names)

        # Open the destination workbook (which now contains the copied sheets) 
        self.dest_wb = self.excel.Workbooks.Open(self.dest_path)
        
        # Create the empty sheets that we will populate in the future and get rid of the default sheet
        for index, sheet_name in enumerate(sheet_names, start=0):
            self.dest_wb.Sheets.Add(After=self.dest_wb.Sheets(self.dest_wb.Sheets.Count)).Name = sheet_name
            self.dest_wb.Sheets(sheet_name).Tab.Color = wsColor[index]
        
        self.dest_wb.Sheets("Sheet1").Delete()

    # This function does the bulk of the work
    # It is responsibel for determining the pivot table location, the pivot cache, the creation of the pivot table
    # And the population of the pivot table after it has been created
    # If you are looking for extra features, this is where you would likely develop them
    def create_pivot_tables(self, pivot_definitions):
        # A dictionary that maps aggregation methods to their respective win32 constants
        agg_dict = {"sum": win32.constants.xlSum, "average": win32.constants.xlAverage}

        # Here, I store columns that need to be computed
        # Key will be the column, should return the string of cols to compute
        compute_dict = {
            "Fuel Dispatch ($/MWh)" : "'Fuel Costs ($000)'/'Generation (GWh)'",
            "VOM Dispatch ($/MWh)" : "'Non-Fuel Variable Cost ($000)'/'Generation (GWh)'",
            "Emissions Dispatch ($/MWh)" : "'Program Costs ($000)'/'Generation (GWh)'"
        }

        # Iterate through all of our pivot deifnitions, and create each table
        for idx, pivot_def in enumerate(pivot_definitions, start=1): 
            print(f"Creating pivot table #{idx} from sheet '{pivot_def.source_sheet}'")
            PivotTableName = f'PivotTable{idx}'

            # Using the information from our pivot definition,
            # Get the destination sheet and determine where our pivot table should go
            dest_sheet = self.dest_wb.Sheets(pivot_def.dest_sheet_pivot)
            available_row = get_next_available_row(dest_sheet)

            pivot_table_location = f'A{available_row}'

            # Creating the pivot cache (where the data in our piovt table "comes from")
            pivot_cache = self.dest_wb.PivotCaches().Create(
                SourceType=win32.constants.xlDatabase,
                SourceData=self.dest_wb.Sheets(pivot_def.source_sheet).UsedRange
            )

            # Creating the actual pivot table
            pivot_table = pivot_cache.CreatePivotTable(
                TableDestination=dest_sheet.Range(pivot_table_location),
                TableName=PivotTableName
            )

            for row in pivot_def.pivot_rows:
                pivot_table.PivotFields(row).Orientation = win32.constants.xlRowField
            
            for col in pivot_def.pivot_cols:
                pivot_table.PivotFields(col).Orientation = win32.constants.xlColumnField
           
            # Stores all of the data fields
            data_field = []
            for index, pivot_data in enumerate(pivot_def.pivot_data, start=0):
                method = agg_dict[pivot_def.agg_methods[index]]
                method_string = pivot_def.agg_methods[index].capitalize()
                
                # Handling our computed columns differently than our basic aggregations
                if (pivot_data in compute_dict.keys()):
                    pivot_table.CalculatedFields().Add(pivot_data, compute_dict[pivot_data])
                    data_field.append(pivot_table.AddDataField(pivot_table.PivotFields(pivot_data),
                                                f'{method_string} of {pivot_data}', method))
                    data_field[index].NumberFormat = '###0.00'
                else:                    
                    data_field.append(pivot_table.AddDataField(pivot_table.PivotFields(pivot_data),
                                                f'{method_string} of {pivot_data}', method))
                    data_field[index].NumberFormat = '###0'

            if (len(data_field) > 1):
                pivot_table.DataPivotField.Orientation = win32.constants.xlRowField
                pivot_table.DataPivotField.Position = 1
            
            pivot_table.RowGrand = False

            # If there are filters desired, add them
            if (pivot_def.pivot_filters):
                pivot_field = pivot_table.PivotFields(pivot_def.pivot_filters)
                pivot_field.Orientation = win32.constants.xlPageField
                pivot_field.Position = 1

            # Gets rid of "(blank)" column
            for field in pivot_table.ColumnFields:
                for item in field.PivotItems():
                    if item.Name == "(blank)":
                        item.Visible = False

            # Table styling
            pivot_table.TableStyle2 = "PivotStyleMedium14"
            hide_errors(pivot_table)

    # Last function called in main
    # Saves our workbook, closes it, and quits our instance of excel
    def finish(self):
        self.dest_wb.Save()
        self.dest_wb.Close()
        self.excel.Quit()

# The function where we define what sheets and pivot tables we desire, utilizing the interface defined above
def main(source_path, dest_path):
    # All the sheets that we want to create and populate, with colors for the sheets
    sheet_names = ["Unit Detail", "Generation", "Unit Dispatch", "Availability", 
                   "Fuel Consumption", "Emissions", "Heat Rate", "Transactions", "Chemical Costs"]
    
    wsColor = ["6299648", "12611584", "15773696", "5296274", "255", "4626167", "65535", "16777215", "26347235"]
    
    # Defines all of the pivot tables we want to create, gets passed as pivot_definitions
    pivots = [
        # Pivot for the Unit Detail Tab
        Pivot_Def(
            source_sheet="Resource Annual",
            pivot_rows=[],
            pivot_cols=['Year'],
            pivot_data= [
                        'Capacity (MW)',
                        'Availability (%)',
                        'Generation (GWh)',
                        'Unit Hours',
                        'Avg Heat Rate (Btu/kWh)',
                        'Heat Required (mmBtu)',
                        'Fuel Costs ($000)',
                        'Program Costs ($000)',
                        'Avg Fuel Cost ($/mmBtu)',
                        'Commitment Costs ($000)',
                        'Non-Fuel Variable Cost ($000)',
                        'Total Energy Cost ($000)',
                        'Average Energy Cost ($/MWh)'
                    ],
            agg_methods=(['sum'] * 13),
            pivot_filters="Resource",
            dest_sheet_pivot="Unit Detail"
        ),

        # Pivot for the Generation Tab
        Pivot_Def(
            source_sheet="Resource Annual",
            pivot_rows=['Resource'],
            pivot_cols=['Year'],
            pivot_data=['Generation (GWh)', 'Capacity (MW)', 'Firm Capacity (MW)'],
            agg_methods=['sum', 'sum', 'sum'],
            pivot_filters="Type",
            dest_sheet_pivot="Generation"
        ),

        # Pivot for the Unit Dispatch Tab
        Pivot_Def(
            source_sheet="Resource Annual",
            pivot_rows=['Resource'],
            pivot_cols=['Year'],
            pivot_data=['Average Energy Cost ($/MWh)', 'Fuel Dispatch ($/MWh)', 'VOM Dispatch ($/MWh)', 'Emissions Dispatch ($/MWh)'],
            agg_methods=['average', 'sum', 'sum', 'sum'],
            pivot_filters="",
            dest_sheet_pivot="Unit Dispatch"
        ),

        # Pivot for the Availability Tab
        Pivot_Def(
            source_sheet="Resource Annual",
            pivot_rows=['Resource'],
            pivot_cols=['Year'],
            pivot_data=['Availability (%)'],
            agg_methods=['sum'],
            pivot_filters="",
            dest_sheet_pivot="Availability"
        ),

        # Pivot 1 for the Fuel Consumption Tab
        Pivot_Def(
            source_sheet="Resource Annual Fuel",
            pivot_rows=['Fuel', 'Resource'],
            pivot_cols=['Year'],
            pivot_data=['Consumption (FUnits)'],
            agg_methods=['sum'],
            pivot_filters="",
            dest_sheet_pivot="Fuel Consumption"
        ),

        # Pivot 2 for the Fuel Consumption Tab
        Pivot_Def(
            source_sheet="Resource Annual",
            pivot_rows=['Resource'],
            pivot_cols=['Year'],
            pivot_data=['Heat Required (mmBtu)'],
            agg_methods=['sum'],
            pivot_filters="Heat Required (mmBtu)",
            dest_sheet_pivot="Fuel Consumption"
        ),

        # Pivot 1 for the Emissions Tab
        Pivot_Def(
            source_sheet="Resource Annual Emissions",
            pivot_rows=['Emission', 'Resource'],
            pivot_cols=['Year'],
            pivot_data=['Released (tons)', 'Release Rate (lb/mmBtu)'],
            agg_methods=['sum', 'sum'],
            pivot_filters="",
            dest_sheet_pivot="Emissions"
        ),

        # Pivot 2 for the Emissions Tab
        Pivot_Def(
            source_sheet="Resource Annual Emissions",
            pivot_rows=['Emission'],
            pivot_cols=['Year'],
            pivot_data=['Released (tons)'],
            agg_methods=['sum'],
            pivot_filters="",
            dest_sheet_pivot="Emissions"
        ),

        # Pivot for the Heat Rate Tab
        Pivot_Def(
            source_sheet="Resource Annual",
            pivot_rows=['Resource'],
            pivot_cols=['Year'],
            pivot_data=['Avg Heat Rate (Btu/kWh)'],
            agg_methods=['average'],
            pivot_filters="",
            dest_sheet_pivot="Heat Rate"
        ),

        # Pivot for the Transactions Tab
        Pivot_Def(
            source_sheet="Company Annual",
            pivot_rows=[],
            pivot_cols=['Year'],
            pivot_data=['Sales (GWh)', 'Purchases (GWh)', 'Sales Revenue ($000)', 'Purchase Cost ($000)'],
            agg_methods=['sum', 'sum', 'sum', 'sum'],
            pivot_filters="",
            dest_sheet_pivot="Transactions"
        ),

        # Pivot 1 for the Chemical Costs Tab
        Pivot_Def(
            source_sheet="Company Annual Programs",
            pivot_rows=['Emission'],
            pivot_cols=['Year'],
            pivot_data=['Externality Costs ($000)'],
            agg_methods=['sum'],
            pivot_filters="Externality Costs ($000)",
            dest_sheet_pivot="Chemical Costs"
        ),

        # Pivot 2 for the Chemical Costs Tab
        Pivot_Def(
            source_sheet="Company Annual Programs",
            pivot_rows=['Emission'],
            pivot_cols=['Year'],
            pivot_data=['Required'],
            agg_methods=['sum'],
            pivot_filters="Externality Costs ($000)",
            dest_sheet_pivot="Chemical Costs"
        )
    ]

    creator = Pivot_Creator(source_path, dest_path, sheet_names, wsColor)
    creator.create_pivot_tables(pivots)
    creator.finish()
    
if __name__ == '__main__':
    # Get the timestamp for the time the script is being executed, and the report file
    timestamp = datetime.now().strftime(r"%Y-%m-%d_%H%M%S")

    # This will give us something like:
    # .../DTE_IRP_Studies/Simulations/2024_Studies/{sens_name}/Outputs/{report_name}.xlsx
    source_path = select_file()
    
    # Using RegEx to find our sens_name from the given directory
    pattern = r'([^\/]+)\/Outputs'
    sens_name = re.search(pattern, source_path).group(1)
    
    # Get the report name
    file_name = os.path.basename(source_path)
    report_name = os.path.splitext(file_name)[0]
 
    # Creating the folder PC_{timestamp}
    new_dir = rf'{os.path.dirname(source_path)}/PC_{timestamp}'
    os.mkdir(new_dir)
    
    # Move the report file into the new folder and update source_path to reflect this change
    report_path = rf'{new_dir}/{file_name}'
    shutil.move(source_path, report_path)

    # Calling our function to create all the csv's
    sheets_to_csv(report_path, sens_name, timestamp)

    # Creating the destination path
    dest_path = rf'{new_dir}/Annual_Generation_Report_{report_name}_{timestamp}.xlsx'    
    
    # Formatting that is windows compatible
    report_path = report_path.replace("/", "\\")
    dest_path = dest_path.replace("/", "\\")

    try:
        main(report_path, dest_path)
        print("Report successfully created!")
    except Exception as e:
        print(f"Error: {e}")
