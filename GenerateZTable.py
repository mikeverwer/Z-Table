import pandas as pd
import scipy.stats
import xlwings as xw
import os
import sys

dependencies = {
    'pandas': '2.3.0',
    'scipy': '1.16.0',
    'xlwings': '0.33.15',
    'pywin32': '310'
}
missing = []
for package, version in dependencies.items():
    try:
        __import__(package)
    except ImportError:
        missing.append(package)
if missing:
    print(f"Missing packages: {', '.join(missing)}")
    print("Install with: pip install -r requirements.txt")
    sys.exit(1)

def generate_z_table():
    z_values = [round(z * 0.01, 2) for z in range(-399, 400)]
    probabilities = scipy.stats.norm.cdf(z_values)
    return pd.DataFrame({"Z-Value": z_values, "Probability": probabilities})

def create_excel_with_lookup():
    df = generate_z_table()
    
    with xw.App(visible=False) as app:
        wb = app.books.add()
        ws = wb.sheets[0]
        
        # Add lookup cells
        ws.range('D1').value = "Find"
        ws.range('D2').value = "Probability"
        ws.range('E1').value = 0.0
        ws.range('E2').formula = '=XLOOKUP(E1,A:A,B:B)'
        
        # Apply Excel's built-in styles
        ws.range('E1').api.Style = "Input"  # Light yellow background
        ws.range('E2').api.Style = "Output"  # Light gray background
        
        # Formatting
        ws.range('E1').number_format = '0.00'
        ws.range('E2').number_format = '0.0000'
        
        # Add Z-table data
        ws.range('A1').options(index=False).value = df
        
        # Save as regular .xlsx (no macro)
        output_path = os.path.abspath('z_table_with_lookup.xlsx')
        wb.save(output_path)
        wb.close()
    
    print(f"Created {output_path}")
    print("NOTE: To add the 'keep selected' functionality:")
    print("1. Open the file in Excel")
    print("2. Press Alt+F11 to open VBA Editor")
    print("3. Paste the macro code (provided below) into Sheet1")
    print("""
Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Range("E1")) Is Nothing Then
        Application.EnableEvents = False
        Range("E1").Select
        Application.EnableEvents = True
    End If
End Sub
""")

if __name__ == "__main__":
    create_excel_with_lookup()
    