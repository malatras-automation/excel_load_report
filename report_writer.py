"""
Python script that turns an Excel list of circuit loads into an overload report.
"""

from openpyxl import load_workbook, Workbook
count = 0 
over_count = 0
normal_count = 0
total = 0

BREAKER = 16

wb_in = load_workbook("loads.xlsx")
ws_in = wb_in.active

report_wb_out = Workbook()
report_ws_out = report_wb_out.active
report_ws_out.append(["Circuit", "Current(A)", "Status"])

for name, current in ws_in.iter_rows(min_row=2, values_only=True):
    
    if current is None:
        continue
    total += float(current)
    count += 1
    
    if current > BREAKER:
        print(name, "OVER:", "->", current, "A")
        status = "OVER"
        over_count += 1
        
    else:
        print( name, "OK:", "->", current, "A")
        status = "OK"
        normal_count += 1
           
    report_ws_out.append([name, current, status])

print(f"circuits counted: {count}")    
print(f"Total current: {total} A")

report_ws_out.append([])
report_ws_out.append(["CIRCUITS:", count, f"  OVER: {over_count} |  OK: {normal_count}"])
report_ws_out.append(["TOTAL current(A):", total,])

report_wb_out.save("report_writer.xlsx")
print(f"📊 Excel Report saved to: report_writer.xlsx")
