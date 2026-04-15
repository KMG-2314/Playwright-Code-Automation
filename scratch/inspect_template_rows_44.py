import openpyxl

input_path = 'd:/Playwright/Data/Quincy-April-Week-2-Input.xlsx'
wb = openpyxl.load_workbook(input_path, data_only=False)
ws = wb.active

print("Detailed check of Rows 44-50:")
for r in range(44, 51):
    vals = [ws.cell(r, c).value for c in range(1, 15)]
    print(f"Row {r}: {vals}")
