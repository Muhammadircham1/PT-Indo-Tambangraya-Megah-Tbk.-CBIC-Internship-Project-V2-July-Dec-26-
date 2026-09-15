import win32com.client, os; xl = win32com.client.Dispatch('Excel.Application'); xl.Visible=False; wb = xl.Workbooks.Open(os.getcwd()+'\\MasterData.xlsx'); sh = wb.Sheets('Fact_Shipments'); 
sh.Range('P3:P68').ClearContents()
sh.Range('P3:P68').Formula = '=IF(N3=1, "implements", "non-implements")'
print('Row 45 formula:', sh.Cells(45, 16).Formula)
wb.Save()
wb.Close()
xl.Quit()
