import win32com.client
excel = win32com.client.Dispatch('Excel.Application')
print("Excel version:", excel.Version)
excel.Visible = False
wb = excel.Workbooks.Open(r"Z:/00 Pastas de trabalho/Asantos/05 - X_TEMPLATE_PREENCHIDO/Fornecedor/Conversor/EXP.MIG.LAY.MM-BP01-BP_For_Preenchido_ARR_PF_LP.xlsx")
print(wb.Name)
wb.Close()
excel.Quit()
