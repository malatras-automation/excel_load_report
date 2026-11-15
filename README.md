
# Excel Load Report
Reads circuit loads from an Excel file and creates an Excel report showing which circuits are overloaded against a 16 A breaker.

# What it does
- Reads circuit names and currents from loads.xlsx 
- Skips empty rows 
- Marks OVER/OK against a 16 A breaker
- Writes report_writer.xlsx with totals

# Example output
[Report](report_screenshot.PNG)

# How to run
1. Install openpyxl:
```
pip install openpyxl
```
2. Put `loads.xlsx` in the same folder as the script, then run:
```
python report_writer.py
```
The report is saved as `report_writer.xlsx`
# Skills shown
openpyxl, loops, data cleaning, Excel automation

# Author
Michalis Malatras, Electrical Engineer, learning Python for industrial automation
