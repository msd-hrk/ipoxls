import openpyxl.worksheet
from common import ctrl_str
from draw import data_converter

class DataWriter():
    
    def __init__(self) -> None:
        # create BindingDesigner instance
        self.ctrl_str = ctrl_str.CtrlStr()
        pass
    
    def write_individual(self, data, sheet: openpyxl.worksheet):
        self.company_summary(data, sheet)
        self.ipo_info(data, sheet)
        self.listing_forecast(data, sheet)
        self.financial_info(data, sheet)
        self.news_list(data, sheet)
        self.lockup_info(data, sheet)
        pass
    
    def company_summary(self, data, sheet: openpyxl.worksheet):
        dc = data_converter.DataConverter(data)
        sheet['A1'] = '証券コード'
        sheet['B1'] = dc.i_securities_num()
        sheet['C1'] = '会社名'
        sheet['D1'] = dc.i_company()
        sheet['D1'].hyperlink = dc.i_web()
        
        sheet['A3'] = '事業内容'
        sheet['B3'] = dc.i_business()
        sheet['A4'] = '詳細'
        sheet['B4'] = dc.i_business_detail()
        
        sheet['A17'] = '会社情報'
        sheet['A18'] = '業種'
        sheet['B18'] = dc.i_category()
        sheet['A19'] = '設立'
        sheet['B19'] = dc.i_build()
        sheet['A20'] = '資本金'
        sheet['B20'] = dc.i_capital()
        sheet['D18'] = '従業員'
        sheet['E18'] = '人数'
        sheet['F18'] = dc.i_emp_num()
        sheet['E19'] = '平均年齢'
        sheet['F19'] = dc.i_emp_age()
        sheet['E20'] = '平均年収'
        sheet['F20'] = dc.i_emp_salary()
        pass
    
    def ipo_info(self, data, sheet: openpyxl.worksheet):
        dc = data_converter.DataConverter(data)
        sheet['A22'] = "IPO情報"
        sheet['A23'] = "上場日"
        sheet['C23'] = dc.i_listing_date()
        sheet['E23'] = "市場"
        sheet['F23'] = dc.i_market()
        sheet['A24'] = "購入期間"
        sheet['C24'] = dc.i_purchase_st()
        sheet['E24'] = "〜"
        sheet['F24'] = dc.i_purchase_ed()
        sheet['A25'] = "ブックビル期間"
        sheet['C25'] = dc.i_bookbuilding_st()
        sheet['E25'] = "〜"
        sheet['F25'] = dc.i_bookbuilding_ed()
        sheet['A26'] = "当選口数"
        sheet['B26'] = dc.i_winning_num()
        sheet['D26'] = "吸収金額"
        sheet['D27'] = dc.i_absorption_amount()
        sheet['F26'] = "公募価格"
        sheet['G26'] = dc.i_public_offer_price()
        sheet['A27'] = "公募株数"
        sheet['B27'] = dc.i_public_offer_shares()
        sheet['C27'] = dc.i_sum_shares()
        sheet['F27'] = "株主数"
        sheet['G27'] = dc.i_holder_cnt()
        sheet['A28'] = "売出株数"
        sheet['B28'] = dc.i_sell_shares()
        sheet['F28'] = "VC数"
        sheet['G28'] = dc.i_vc_cnt()
        sheet['A29'] = "QA分"
        sheet['B29'] = dc.i_oa_shares()
        sheet['F29'] = "VC保有率"
        sheet['G29'] = dc.i_vc_rate()
        sheet['A30'] = "発行済み株式数"
        sheet['C30'] = dc.i_issued_shares()
        sheet['D30'] = "主幹事銀行"
        sheet['E30'] = dc.i_bank_name()
        sheet['G30'] = "評価"
        sheet['G31'] = dc.i_grade()
        sheet['A31'] = "QA率"
        sheet['C31'] = dc.i_oar()
        sheet['D31'] = "取扱割合"
        sheet['E31'] = dc.i_bank_rate()
    
    def listing_forecast(self, data, sheet: openpyxl.worksheet):
        dc = data_converter.DataConverter(data)
        sheet['A33'] = "上場予想"
        sheet['A34'] = "仮条件決定後"
        sheet['E34'] = "上場結果"
        sheet['A35'] = "想定価格"
        sheet['B35'] = dc.i_assumed_price()
        sheet['E35'] = "公募価格"
        sheet['F35'] = dc.i_public_offer_price()
        sheet['A36'] = "初値予想"
        sheet['B36'] = dc.i_initial_price_forecast()
        sheet['E36'] = "初値"
        sheet['F36'] = dc.i_init_price()
        sheet['A37'] = "予想利益"
        sheet['B37'] = dc.i_expected_profit()
        sheet['E37'] = "初値損益"
        sheet['F37'] = dc.i_init_price_sell_profit()
        sheet['A38'] = "吸収金額"
        sheet['B38'] = dc.i_abs_amount_aft_td()
        sheet['E38'] = "騰落率"
        sheet['F38'] = dc.i_rf_rate()
        sheet['A39'] = "時価総額"
        sheet['B39'] = dc.i_market_capitalization_aft_td()
    
    def financial_info(self, data, sheet: openpyxl.worksheet):
        dc = data_converter.DataConverter(data)
        sheet['A41'] = "単独財務情報"
        sheet['A42'] = "決算期"
        sheet['A43'] = "売上高"
        sheet['A44'] = "経常利益"
        sheet['A45'] = "当期利益"
        sheet['A46'] = "純資産"
        sheet['A47'] = "配当金"
        sheet['A48'] = "EPS"
        sheet['A49'] = "BPS"
        
        i = 0
        row = ["B", "D", "F"]
        dts = dc.i_indpndnt_fin_info()
        for dt in dts:
            try:
                sheet[ row[i] + '42'] = dt["fiscalYear"][:-2] + "年" + dt["fiscalYear"][-2:] + "月"
                sheet[ row[i] + '43'] = self.ctrl_str.cnv_ja_money(dt["amountOfSalls"])
                sheet[ row[i] + '44'] = self.ctrl_str.cnv_ja_money(dt["ordinaryIncome"])
                sheet[ row[i] + '45'] = self.ctrl_str.cnv_ja_money(dt["netIncome"])
                sheet[ row[i] + '46'] = self.ctrl_str.cnv_ja_money(dt["netWorth"])
                sheet[ row[i] + '47'] = dt["dividend"]
                sheet[ row[i] + '48'] = dt["bps"]
                sheet[ row[i] + '49'] = dt["eps"]
                i = i + 1
            except:
                break
        

        sheet['A51'] = "連結財務情報"
        sheet['A52'] = "決算期"
        sheet['A53'] = "売上高"
        sheet['A54'] = "経常利益"
        sheet['A55'] = "当期利益"
        sheet['A56'] = "純資産"
        sheet['A57'] = "EPS"
        sheet['A58'] = "BPS"

        i = 0
        row = ["B", "D", "F"]
        dts = dc.i_cnsldtd_fin_info()
        for dt in dts:
            try:
                sheet[ row[i] + '52'] = dt["fiscalYear"][:-2] + "年" + dt["fiscalYear"][-2:] + "月"
                sheet[ row[i] + '53'] = self.ctrl_str.cnv_ja_money(dt["amountOfSalls"])
                sheet[ row[i] + '54'] = self.ctrl_str.cnv_ja_money(dt["ordinaryIncome"])
                sheet[ row[i] + '55'] = self.ctrl_str.cnv_ja_money(dt["netIncome"])
                sheet[ row[i] + '56'] = self.ctrl_str.cnv_ja_money(dt["netWorth"])
                sheet[ row[i] + '57'] = dt["bps"]
                sheet[ row[i] + '58'] = dt["eps"]
                i = i + 1
            except:
                break

    def news_list(self, data, sheet: openpyxl.worksheet):
        sheet['A60'] = "ニュース"
        sheet['A61'] = "日付"
        sheet['C61'] = "タイトル"

    def lockup_info(self, data, sheet: openpyxl.worksheet):
        dc = data_converter.DataConverter(data)

        cell_row = 82

        sheet['A' + str(cell_row)] = "ロックアップ情報"
        
        cnt = 0
        for row in range(cell_row + 1, cell_row + 40, 4):
            cnt = cnt + 1
            sheet["A" + str(row)] = str(cnt)
            sheet["B" + str(row)] = "割合"
            sheet["C" + str(row)] = "所有者"
            sheet["C" + str(row + 1)] = "役職"
            sheet["C" + str(row + 2)] = "ロック"
            sheet["D" + str(row + 2)] = "日数"
            sheet["F" + str(row + 2)] = "レート"
            sheet["B" + str(row + 3)] = "保有株式"
            sheet["D" + str(row + 3)] = "売出株式"
            sheet["F" + str(row + 3)] = "潜在株式"
            
        i = cell_row + 1
        for dt in dc.i_share_holders():
            sheet["D" + str(i)] = dt["name"]
            sheet["B" + str(i + 1)] = dt["rate"]
            sheet["D" + str(i + 1)] = dt["position"]
            sheet["E" + str(i + 2)] = dt["lockup"]["day"]
            sheet["G" + str(i + 2)] = dt["lockup"]["rate"]
            sheet["C" + str(i + 3)] = dt["sheredNum"]
            sheet["E" + str(i + 3)] = dt["outShares"]
            sheet["G" + str(i + 3)] = dt["potentialShares"]
            i = i + 4
            