import pandas as pd
import scipy.stats
from openpyxl import Workbook
from openpyxl.styles import NamedStyle, PatternFill

def generate_z_table():
    """Generate Z-table from -3.99 to 3.99 in 0.01 increments"""
    z_values = [round(z * 0.01, 2) for z in range(-399, 400)]
    probabilities = scipy.stats.norm.cdf(z_values)
    return pd.DataFrame({"Z-Value": z_values, "Probability": probabilities})

def create_excel_file():
    # Create workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Z-Table"
    
    # Create and register styles
    input_style = NamedStyle(name="input_style")
    input_style.fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")  # Yellow
    input_style.number_format = '0.0000'
    
    output_style = NamedStyle(name="output_style")
    output_style.fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")  # Gray
    output_style.number_format = '0.000000'
    
    # Register styles (important!)
    wb.add_named_style(input_style)
    wb.add_named_style(output_style)
    
    # Add Z-table data (simpler method without dataframe_to_rows)
    df = generate_z_table()
    
    # Write headers
    ws['A1'] = "Z-Value"
    ws['B1'] = "Probability"
    
    # Write data
    for idx, row in df.iterrows():
        ws.cell(row=idx+2, column=1, value=row['Z-Value'])
        ws.cell(row=idx+2, column=2, value=row['Probability'])
    
    # Add lookup cells
    ws['D1'] = "Find"
    ws['D2'] = "Probability"
    ws['E1'] = -1.1
    ws['E1'].style = "input_style"
    ws['E2'] = '=_xlfn.XLOOKUP(E1,A:A,B:B)' 
    # ws['E2'].data_type = 'f'
    ws['E2'].style = "output_style"
    
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