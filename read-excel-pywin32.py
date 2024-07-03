import win32com.client
import random
import json
import os

# Constants
INPUT_CELLS = ['C2', 'C3', 'C4', 'C7', 'C8', 'C9']
CALCULATION_CELLS = ['C12', 'C13']
DEFAULT_FILENAME = 'data.xlsx'
DEFAULT_SHEET_NAME = 'SheetName2'
DEFAULT_ITERATIONS = 10
RANDOM_RANGE = (0.95, 1.05)
DECIMAL_PLACES = 4
OUTPUT_FILENAME = 'calculations-pywin32.json'

def open_excel():
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    return excel

def load_workbook(excel, filename):
    return excel.Workbooks.Open(os.path.abspath(filename))

def read_cell_value(sheet, cell):
    value = sheet.Range(cell).Value
    if isinstance(value, (int, float)):
        return round(value, DECIMAL_PLACES)
    return value

def update_cell_value(sheet, cell, value):
    sheet.Range(cell).Value = value

def update_input_randomly(sheet, input_cells):
    for cell in input_cells:
        current_value = read_cell_value(sheet, cell)
        if isinstance(current_value, (int, float)):
            new_value = current_value * random.uniform(*RANDOM_RANGE)
            update_cell_value(sheet, cell, new_value)

def get_all_values(sheet, input_cells, calculation_cells):
    values = {}
    for cell in input_cells + calculation_cells:
        label = sheet.Range(f'A{cell[1:]}').Value
        unit = sheet.Range(f'B{cell[1:]}').Value
        value = read_cell_value(sheet, cell)
        values[label] = {'value': value, 'unit': unit}
    return values

def main(filename=DEFAULT_FILENAME, sheet_name=DEFAULT_SHEET_NAME, n=DEFAULT_ITERATIONS):
    all_runs = {}
    
    excel = open_excel()
    wb = load_workbook(excel, filename)
    sheet = wb.Worksheets(sheet_name)
    
    try:
        for iteration in range(DEFAULT_ITERATIONS):
            update_input_randomly(sheet, INPUT_CELLS)
            
            # Force recalculation
            wb.Application.Calculate()
            
            iteration_values = get_all_values(sheet, INPUT_CELLS, CALCULATION_CELLS)
            all_runs[f"Iteration {iteration + 1}"] = iteration_values
            
        # Save all_runs to calculations.json
        with open(OUTPUT_FILENAME, 'w') as f:
            json.dump(all_runs, f, indent=2)
        
        print(f"Calculations saved to {OUTPUT_FILENAME}")
    
    finally:
        wb.Close(SaveChanges=False)
        excel.Quit()

if __name__ == "__main__":
    main()