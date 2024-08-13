import json
import odf
from odf.opendocument import load
from odf.table import Table, TableRow, TableCell
from odf.text import P

def extract_text(cell):
    text_content = []
    for paragraph in cell.getElementsByType(P):
        text_content.append(str(paragraph))
    return ' '.join(text_content).strip()

ods_file = "O4.ods"
spreadsheet = load(ods_file)
data = {}

for table in spreadsheet.getElementsByType(Table):
    tab_name = str(table.getAttribute("name"))
    print(f"Tab Name Var: {tab_name}")
    tab_data = {}
    
    for row in table.getElementsByType(TableRow):
        cells = row.getElementsByType(TableCell)
        print(f"Tab Name Var: {cells}")
        
        ion_name = extract_text(cells[0]) if len(cells) > 0 else None
        wavelength = extract_text(cells[1]) if len(cells) > 1 else None

        if ion_name:  # if there is an ion name
            tab_data[ion_name] = wavelength
        else:  # if no ion name, use the tab's file name as key
            tab_data[tab_name] = wavelength

    # Save the tab data under the tab's name in the main dictionary
    data[tab_name] = tab_data

# Save the extracted data to a JSON file
output_file = "spectra.json"

with open(output_file, "w") as json_file:
    json.dump(data, json_file, indent=4)

print(f"Data Dumped and (hopefully) saved to {output_file}")