# Import the openpyxl library for working with Excel files
import openpyxl

# Import the l93_to_wgs84 function from the coordinates_converter.py file
from coordinates_converter import l93_to_wgs84

# Define the path to the input Excel file
wbk_name = "C:/Users/liang.jingyi/Documents/CALCUL DE CHARGE/BLM/31600/Tableau CAP FT.xlsx"

# Define the name of the worksheet to work with
sheet_name = "Saisies terrain"

# Define the commune name and INSEE code
commune = "Dame-Marie"
insee = '61142'
adresse = "Lieu dit le Pommerais"

# Load the workbook and get the active worksheet
try:
    wb = openpyxl.load_workbook(wbk_name)
    ws = wb[sheet_name]
except Exception as e:
    print(f"Error loading workbook or worksheet: {e}")
    # You could add additional error handling here if needed, such as exiting the script

# Get the last row number in the worksheet
last_row = ws.max_row

# Convert Lambert 93 coordinates to WGS84 (latitude, longitude) and update the worksheet
for i in range(9, last_row+1):
    # Check if the first column of the current row is not empty and if the fourth column is a float
    if ws.cell(row=i, column=1).value is not None and type(ws.cell(row=i, column=4).value) is float:
        # Get the Lambert 93 coordinates from the fourth and fifth columns of the current row
        x = ws.cell(row=i, column=4).value
        y = ws.cell(row=i, column=5).value
        # Convert the Lambert 93 coordinates to WGS84 (latitude, longitude) using the l93_to_wgs84 function
        lat, lon = l93_to_wgs84(x, y)
        # Update the fourth and fifth columns of the current row with the converted WGS84 coordinates
        ws.cell(row=i, column=4).value = lat
        ws.cell(row=i, column=5).value = lon
    else:
        # If the first column of the current row is empty or the fourth column is not a float, skip this row
        continue

# Fill in specific values
for i in range(9, last_row+1):
    if ws.cell(row=i, column=1).value is not None:
        if ws.cell(row=i, column=6).value == 'FEN' or ws.cell(row=i, column=6).value == 'EPA' or ws.cell(row=i, column=6).value == 'CHO':
            ws.cell(row=i, column=14).value == 'Non'
            for j in range(7,14):
                ws.cell(row=i, column=j).value = 'Oui'
        elif ws.cell(row=i, column=6).value == 'AUT':
            ws.cell(row=i, column=14).value == 'Non'
            for j in range(7,14):
                ws.cell(row=i, column=j).value = 'Oui'
        elif ws.cell(row=i, column=13).value == 'Oui' or ws.cell(row=i, column=6).value is None:
            for j in range(6,15):
                ws.cell(row=i, column=j).value = 'Oui'
        ws.cell(row=i, column=15).value = 'TER'
        ws.cell(row=i, column=23).value = '0'
        ws.cell(row=i, column=24).value = '15'
        ws.cell(row=i, column=26).value = 'Oui'
        ws.cell(row=i, column=27).value = 'Non'
        ws.cell(row=i, column=35).value = 'Oui'
        ws.cell(row=i, column=36).value = 'Non'
        ws.cell(row=i, column=37).value = '0'
        ws.cell(row=i, column=38).value = 'Non'
        ws.cell(row=i, column=39).value = '0'
        ws.cell(row=i, column=40).value = 'Non'
        if ws.cell(row=i, column=3).value is None or ws.cell(row=i, column=3).value == 'Adresse non trouvée':
            ws.cell(row=i, column=3).value = adresse
        if ws.cell(row=i, column=2).value == 'BH6':
            ws.cell(row=i, column=2).value = 'BH6 S30'
        elif ws.cell(row=i, column=2).value == 'BH7':
            ws.cell(row=i, column=2).value = 'BH7 S30'
        elif ws.cell(row=i, column=2).value == 'MH7':
            ws.cell(row=i, column=2).value = 'MH7 S30'
        elif ws.cell(row=i, column=2).value == 'BH8':
            ws.cell(row=i, column=2).value = 'BH8 S30'
        elif ws.cell(row=i, column=2).value == 'MH8':
            ws.cell(row=i, column=2).value = 'MH8 S30'
        if ws.cell(row=i, column=2).value == 'FR7':
            ws.cell(row=i, column=2).value = 'BS7'
        if ws.cell(row=i, column=2).value == 'MR0':
            ws.cell(row=i, column=2).value = 'M20'
        elif 'Ab' in ws.cell(row=i, column=1).value:
            ws.cell(row=i, column=2).value = 'POT MIN'
        elif ws.cell(row=i, column=2).value == 'BETON':
            ws.cell(row=i, column=2).value = 'EDF'

    if ws.cell(row=i, column=19).value in ['L1092-11', 'L1092-13', 'L1092-14', 'L1092-15']:
        ws.cell(row=i, column=19).value = ws.cell(row=i, column=19).value + '-A'
        ws.cell(row=i, column=19).font = openpyxl.styles.Font(bold=True)
    elif ws.cell(row=i, column=19).value == 'L1092-12-P':
        ws.cell(row=i, column=19).value = 'L1092-12-A'
        ws.cell(row=i, column=19).font = openpyxl.styles.Font(bold=True)

wb.save(wbk_name)