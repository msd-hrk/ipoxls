from conf.DrawConf import DrawConf

class CrtDrawArry():
    def __init__(self) -> None:
        self.title = 0
        self.field_name = 1
        self.draw_conf = DrawConf()
        pass
    
    def addDrawColum(self, data, draw_obj):
        # タイトル行作成
        if len(draw_obj) == 0:
            title_colum = []
            for dt in self.draw_conf.draw_arry:
                title_colum.append(dt[self.title])
            draw_obj.append(title_colum)
        
        # 1行ずつ情報を描画
        colum = []
        for target in self.draw_conf.draw_arry:
            f_name = target[self.field_name]
            try:
                colum.append(data[f_name])
            except:
                colum.append("-")
        colum.extend(self.crtDiaryPriceArry(data))

        draw_obj.append(colum)
        print(draw_obj)
        return draw_obj
    
    def crtDiaryPriceArry(self, data):
        dairy_price_arry = []
        for day_data in data["priceDiary"]:
            dairy_price_arry.append(day_data["closePrice"])
        return dairy_price_arry