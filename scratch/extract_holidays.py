import pdfplumber
import json

pdf_path = 'd:/Playwright/Data/Holiday List 2026.pdf'
with pdfplumber.open(pdf_path) as pdf:
    all_data = []
    for page in pdf.pages:
        table = page.extract_table()
        if table:
            all_data.extend(table)

with open('d:/Playwright/scratch/holiday_table.json', 'w', encoding='utf-8') as f:
    json.dump(all_data, f, ensure_ascii=False, indent=2)

print("Saved holiday table to d:/Playwright/scratch/holiday_table.json")
