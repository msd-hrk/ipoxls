from dbquery.AccessDb import accessDb
from draw.CrtDrawArry import CrtDrawArry
from draw.CtrlXlsx import CtrlXlsx

print("process start")

# mongoDBからデータを取得
print("DB access")
dbInfo = accessDb()
# data = dbInfo.find_one_from_securities_no("5131")
datas = dbInfo.find_listed_data()
# DataSet(data)
d_arry = CrtDrawArry()
d_obj = []
for data in datas:
    d_obj = d_arry.addDrawColum(data, d_obj)
    

path = "test.xlsx"
c_xlsx = CtrlXlsx(path)
c_xlsx.writeXlsx("sample", d_obj)
# データをエクセルに描画
print("process end")


