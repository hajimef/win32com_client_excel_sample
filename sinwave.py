import win32com.client, time, math, os

# Excelを起動する
xl = win32com.client.Dispatch("Excel.Application")
xl.Visible = True
# カレントディレクトリの「sinwave.xlsx」ファイルを開く
wb = xl.Workbooks.Open(os.getcwd() + "/sinwave.xlsx")
# 先頭のワークシートを得る
ws = wb.Worksheets(1)
# 画面を再描画をオンにする
xl.ScreenUpdating = True
p = 0
while True:
    # 画面の再描画を止める
    xl.ScreenUpdating = False
    # B2セル～C5セルの値を書き換える
    for i in range(0, 4):
        ws.Cells(i + 2, 2).Value = math.sin(math.pi / 180 * (i * 90 + p)) * 50 + 100
        ws.Cells(i + 2, 3).Value = math.cos(math.pi / 180 * (i * 90 + p)) * 50 + 100
    # 画面を再描画をオンにする
    xl.ScreenUpdating = True
    p += 5
    p %= 360
xl.Quit()
