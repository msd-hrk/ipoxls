import openpyxl
from conf import config
from dbquery import access_db
from draw import sheet_creater

# 設定ファイル読み込み
conf = config.Config()

# 対象データ抽出
db = access_db.AccessDb()
datas = db.get_secrities_num_list()

# ワークブック作成
workbook = openpyxl.Workbook()

# 会社別シートの描画
for data in datas:
    print("write:" + data["securitiesNo"])
    ccs = sheet_creater.SheetCreater(workbook)
    ccs.add_individual_sheet(data["securitiesNo"])
    
# 一覧シート描画

# 終了
workbook.save("/tmp/test.xlsx")
