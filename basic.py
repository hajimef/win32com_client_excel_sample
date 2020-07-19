import win32com.client, time, os

# Excelを起動する
xl = win32com.client.Dispatch("Excel.Application")
# Excelを表示する
xl.Visible = True
# ブックを新規作成する
wb = xl.Workbooks.Add()
# ブックの先頭のワークシートを得る
ws = wb.Worksheets(1)
# ワークシートの1行1列のセルに「Hello Excel」を入力する
ws.Cells(1, 1).Value = "Hello Excel"
# カレントディレクトリに「basic.xlsx」というファイル名で保存する
xl.DisplayAlerts = False
wb.SaveAs(os.getcwd() + "/basic.xlsx")
xl.DisplayAlerts = True
# Excelを終了する
xl.Quit()
