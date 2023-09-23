from calendar import month


class CtrlStr():
    def __init__(self) -> None:
        pass
    
    def remove(self, before_data, *args):
        target = before_data
        for item in args:
            target = str(target).replace(item, "")
        return target
    
    def cnv_ja_money(self, money: int) -> str:
        str_in_money = str(money)
        minus_str = ""
        # 負の数か？
        if str_in_money[0:1] == "-":
            minus_str = str_in_money[0:1]
            str_in_money = str_in_money[1:]
        ja_money = ""
        tens_of_thousands = 0
        if len(str_in_money) > 8:
            ja_money = str_in_money[:-8] + '億'
            tens_of_thousands = int(str_in_money[-8:-4])
        else:
            tens_of_thousands = int(str_in_money[:-4])
            
        if not tens_of_thousands == 0:
            ja_money = ja_money + str(tens_of_thousands) + '万'
        
        return minus_str + ja_money + '円'

    def cnv_ja_yyyy_mm_dd(self, date_num: int):
        year = str(date_num)[:-4]
        month = str(date_num)[-4:-2]
        day = str(date_num)[-2:]
        return year + "年" + month + "月" + day + "日"