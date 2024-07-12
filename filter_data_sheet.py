import openpyxl

# "Book1.xlsx"を開く
try:
    workbook = openpyxl.load_workbook('Book1.xlsx')
except FileNotFoundError:
    print("ファイル 'Book1.xlsx' が見つかりません。")
    exit()

# "Sheet1" を取得
sheet1 = workbook.get_sheet_by_name('Sheet1')

# "Sheet2" が既に存在するか確認
if 'Sheet2' in workbook.sheetnames:
    print("エラー: 'Sheet2' シートが既に存在します。")
    exit()

# "Sheet2" を新規作成
sheet2 = workbook.create_sheet('Sheet2')

# "Sheet1"のA列の値が「済」ではない行だけを"Sheet2"にコピー
for row in sheet1.iter_rows(min_row=2, values_only=True):
    if row[0] != '済':
        sheet2.append(row)

# ファイルを保存
workbook.save('Book1.xlsx')

print("'Sheet2'シートにデータをコピーしました。")