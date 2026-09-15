import openpyxl
from openpyxl.utils import column_index_from_string
wb = openpyxl.load_workbook(r'd:\magang\ITMG-Internship-All-Project-2025\SystemAutomation_SSOtoSummary\data\10 August_Margin.xlsx', data_only=False)
ws = wb['ITM Summary']
r = 264
for col in ['AOV', 'AOW', 'AOX', 'AOY', 'AOZ', 'APA', 'APB', 'APC', 'APD', 'APE', 'APF', 'APG', 'APH']:
    print(f"{col}: {ws.cell(row=r, column=column_index_from_string(col)).value}")
