import openpyxl
wb = openpyxl.load_workbook(r'd:\magang\ITMG-Internship-All-Project-2025\SystemAutomation_SSOtoSummary\data\10 August_Margin.xlsx', data_only=True)
ws = wb['FC Quality Master']
months = {}
for r in range(1, 1500):
    c = ws.cell(row=r, column=2).value
    if c and hasattr(c, 'month') and hasattr(c, 'year'):
        if c.month not in months:
            months[c.month] = r
print(months)

