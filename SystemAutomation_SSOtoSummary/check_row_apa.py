import openpyxl
from openpyxl.utils import column_index_from_string
wb = openpyxl.load_workbook(r'd:\magang\ITMG-Internship-All-Project-2025\SystemAutomation_SSOtoSummary\data\10 August_Margin.xlsx', data_only=False)
ws = wb['ITM Summary']
for r in range(1, 1500):
    val = ws.cell(row=r, column=column_index_from_string("APA")).value
    if val:
        print(f"Row {r}: {val}")
