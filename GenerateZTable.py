import pandas as pd
import scipy.stats
from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.styles import NamedStyle, PatternFill, numbers

def generate_z_table():
    """Generate Z-table from -3.99 to 3.99 in 0.01 increments"""
    z_values = [round(z * 0.01, 2) for z in range(-399, 400)]
    probabilities = scipy.stats.norm.cdf(z_values)
    return pd.DataFrame({"Z-Value": z_values, "Probability": probabilities})
        
def create_named_ranges(wb, ws_title, data_length, named_cells):
    # 1. Create column ranges
    column_ranges = {
        'Z_Values': f'{ws_title}!$A$2:$A${data_length+1}',
        'Probabilities': f'{ws_title}!$B$2:$B${data_length+1}'
    }
    # 2. Create single cell references
    for name, cell_ref in named_cells.items():
        if not cell_ref or len(cell_ref) < 2:
            continue  # Skip invalid
            
        col = cell_ref[0].upper()
        row = cell_ref[1:]
        column_ranges[name] = f'{ws_title}!${col}${row}'
    # 3. Add to workbook
    for name, ref_text in column_ranges.items():
        wb.defined_names.add(DefinedName(name, attr_text=ref_text))
        
def style_cells(ws, cells, style):
    for cell in cells:
        ws[cell].style = style
        
def format_output_cells(ws):
    for cell in ws['B'][1:]:
        cell.number_format = '0.0000'
    two_dec = ['E1', 'E4', 'H9', 'K10']
    four_dec = ['E2', 'E5', 'E7', 'H7', 'K7', 'K8']
    for cell in two_dec:
        ws[cell].number_format = '0.00'
    for cell in four_dec:
        ws[cell].number_format = '0.0000'
    
def set_column_width(ws, columns, width):
    for col in columns:
        ws.column_dimension[col].width = width

def create_excel_file():
    # Create workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Z_Table"
    
    # Add Z-table data
    df = generate_z_table()
    
    # Write headers
    ws['A1'] = "Z-Value"
    ws['B1'] = "Probability"
    ws['G1'] = "Sample Mean"
    ws['J1'] = "Sample Proportion"
    heading_cells: list = ['A1', 'B1', 'G1', 'J1']
    
    # Write data
    for idx, row in df.iterrows():
        ws.cell(row=idx+2, column=1, value=row['Z-Value'])
        ws.cell(row=idx+2, column=2, value=row['Probability'])
        
    named_cells:dict = {
        'Z1_Input': 'E1',
        'Probability_Z1': 'E2',
        'Z2_Input': 'E4',
        'Probability_Z2': 'E5',
        'mu': 'H2',
        'sigma': 'H3',
        'sampleSize_mean': 'H4',
        'popSize_mean': 'H5',
        'x_mean': 'H6',
        'sigma_x': 'H7',
        'p': 'K2',
        'sampleSize_prop': 'K3',
        'popSize_prop': 'K4',
        'x_prop': 'K5',
        'pbar_man': 'K6',
        'pbar_calc': 'K7',
        'sigma_p': 'K8'
    }
    create_named_ranges(wb, ws.title, len(df), named_cells)
    
    # Add cell labels
    # Z-lookup
    ws['D1'] = "Find Z_1"
    ws['D2'] = "Probability"
    ws['D4'] = "Find Z_2"
    ws['D5'] = 'Probability'
    ws['D7'] = "P(z1 < z < z2)"
    # sample mean
    ws['G2'] = "μ"
    ws['G3'] = "σ"
    ws['G4'] = "n"
    ws['G5'] = "N"
    ws['G6'] = "x"
    ws['G7'] = "σₓ"
    ws['G9'] = "zₓ"
    #sample proportion
    ws['J2'] = "p"
    ws['J3'] = "n"
    ws['J4'] = "N"
    ws['J5'] = "x"
    ws['J6'] = "p̄ (manual)"
    ws['J7'] = "p̄ (calc)"
    ws['J8'] = "σₚ"
    ws['J10'] = "zₚ"
    
    input_cells: list[str] = ['E1', 'E4', 'H2', 'H3', 'H4', 'H5', 'H6', 'K2', 'K3', 'K4', 'K5', 'K6']
    output_cells: list[str] = ['E2', 'E5', 'E7', 'H9', 'K10']
    calculation_cells: list[str] = ['H7', 'K7', 'K8']
    
    # Style cells
    style_cells(ws, heading_cells, "Headline 3")
    style_cells(ws, input_cells, 'Input')
    style_cells(ws, output_cells, 'Output')
    style_cells(ws, calculation_cells, 'Calculation')
    
    format_output_cells(ws)
    set_column_width(['A'], 7)
    set_column_width(['B'], 9)
    set_column_width(['D'], 11)
    set_column_width(['J'], 10)
    set_column_width(['G'], 5)
    
    # Z1_Input
    ws['E1'] = 0
    # Probability_Z1
    ws['E2'] = '=_xlfn.XLOOKUP(Z1_Input, Z_Values, Probabilities)'
    # Z2_Input
    ws['E4'] = 0
    # Probability_Z2
    ws['E5'] = '=_xlfn.XLOOKUP(Z2_Input, Z_Values, Probabilities)'
    # P(z1 < z < z2)
    ws['E7'] = '=ABS(Probability_Z2 - Probability_Z1)'
    
    # Save
    wb.save("z_table.xlsx")
    print("Successfully created z_table.xlsx")
    print("\nTo add the 'keep selected' functionality:")
    print("1. Open z_table.xlsx and press Alt+F11")
    print("2. Paste this code into Sheet1:")
    print("""
Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Range("E1")) Is Nothing Then
        Application.EnableEvents = False
        Range("E1").Select
        Application.EnableEvents = True
    End If
End Sub
""")
    print("3. Save as z_table.xlsm (macro-enabled workbook)")

if __name__ == "__main__":
    create_excel_file()
    