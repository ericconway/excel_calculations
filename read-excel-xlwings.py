import xlwings as xw
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
OUTPUT_FILENAME = 'calculations-xlwings.json'

def read_cell_value(sheet, cell):
    value = sheet.range(cell).value
    if isinstance(value, (int, float)):
        return round(value, DECIMAL_PLACES)
    return value

def update_cell_value(sheet, cell, value):
    sheet.range(cell).value = value

def update_input_randomly(sheet, input_cells):
    for cell in input_cells:
        current_value = read_cell_value(sheet, cell)
        if isinstance(current_value, (int, float)):
            new_value = current_value * random.uniform(*RANDOM_RANGE)
            update_cell_value(sheet, cell, new_value)

def get_all_values(sheet, input_cells, calculation_cells):
    values = {}
    for cell in input_cells + calculation_cells:
        label = sheet.range(f'A{cell[1:]}').value
        unit = sheet.range(f'B{cell[1:]}').value
        value = read_cell_value(sheet, cell)
        values[label] = {'value': value, 'unit': unit}
    return values

def main(filename=DEFAULT_FILENAME, sheet_name=DEFAULT_SHEET_NAME, n=DEFAULT_ITERATIONS):
    all_runs = {}
    
    with xw.App(visible=False) as app:
        wb = app.books.open(filename)
        sheet = wb.sheets[sheet_name]
        
        for iteration in range(DEFAULT_ITERATIONS):
            update_input_randomly(sheet, INPUT_CELLS)
            
            # Force recalculation
            wb.app.calculate()
            
            iteration_values = get_all_values(sheet, INPUT_CELLS, CALCULATION_CELLS)
            all_runs[f"Iteration {iteration + 1}"] = iteration_values
        
        wb.save()
        wb.close()
    
    # Save all_runs to calculations.json
    with open(OUTPUT_FILENAME, 'w') as f:
        json.dump(all_runs, f, indent=2)
    
    print(f"Calculations saved to {OUTPUT_FILENAME}")

if __name__ == "__main__":
    main()