from dbquery import accessDb

print("process start")

# mongoDBからデータを取得
print("DB access")
dbInfo = accessDb.accessDb()
data = dbInfo.find_all_data()
print(data)
# データをエクセルに描画
print("process end")


