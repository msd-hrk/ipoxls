from common import ctrl_str


class DataConverter():
    
    def __init__(self, data) -> None:
        self.data = data
        self.ctrl_str = ctrl_str.CtrlStr()
        pass
    
    def i_securities_num(self):
        return self.data["securitiesNo"]
    
    def i_company(self):
        return self.data["company"]
    
    def i_web(self):
        return self.data["web"]
    def i_business(self):
        return self.data["business"]
    
    def i_business_detail(self):
        return self.ctrl_str.remove(self.data["businessDetail"], " ", "　", "\r\n", "\n")
    
    def i_category(self):
        return self.data["category"]
    
    def i_build(self):
        return str(self.data["build"]) + "年"
    
    def i_capital(self):
        return self.ctrl_str.cnv_ja_money(self.data["capital"])
    
    def i_emp_num(self):
        return str(self.data["employee"]["num"]) + '人'
    
    def i_emp_age(self):
        return str(self.data["employee"]["age"]) + '歳'
    
    def i_emp_salary(self):
        return str(self.data["employee"]["salary"]) + '万'
    
    def i_listing_date(self):
        return self.ctrl_str.cnv_ja_yyyy_mm_dd(self.data["listingDate"])
    
    def i_purchase_st(self):
        return self.ctrl_str.cnv_ja_yyyy_mm_dd(self.data["purchasePeriod"]["start"])
    
    def i_purchase_ed(self):
        return self.ctrl_str.cnv_ja_yyyy_mm_dd(self.data["purchasePeriod"]["end"])

    def i_bookbuilding_st(self):
        return self.ctrl_str.cnv_ja_yyyy_mm_dd(self.data["bookbuilding"]["start"])

    def i_bookbuilding_ed(self):
        return self.ctrl_str.cnv_ja_yyyy_mm_dd(self.data["bookbuilding"]["end"])
    
    def i_winning_num(self):
        return str(self.data["winningNum"]) + '口'
    
    def i_public_offer_price(self):
        return str(self.data["pubOfferPrice"]) + "円"
    
    def i_public_offer_shares(self):
        return str(self.data["pubOfferedShares"]) + "株"
    
    def i_sell_shares(self):
        return str(self.data["sellShares"]) + "株"
    
    def i_oa_shares(self):
        return str(self.data["oaShared"]) + "株"
    
    def i_sum_shares(self):
        return str(self.data["pubOfferedShares"] + self.data["oaShared"] + self.data["sellShares"]) + "株"

    def i_issued_shares(self):
        return str(self.data["issuedShares"]) + "株"
    
    def i_oar(self):
        return str(self.data["oar"] * 100) + "%"
    
    def i_absorption_amount(self):
        pub_offer_price = self.data["pubOfferPrice"]
        sum_shares = self.data["pubOfferedShares"] + self.data["oaShared"] + self.data["sellShares"]
        absorption_amount = sum_shares * pub_offer_price
        return self.ctrl_str.cnv_ja_money(absorption_amount)
    
    def i_holder_cnt(self):
        return str(self.data["holder_num"]) + "人"
    
    def i_vc_cnt(self):
        vc_cnt = 0
        for holder in self.data["shareholders"]:
            if holder["position"] in "vc" or holder["position"] in "VC":
                vc_cnt = vc_cnt + 1
        return str(vc_cnt) + "社"
    
    def i_vc_rate(self):
        vc_rate = 0
        for holder in self.data["shareholders"]:
            if holder["position"] in "vc" or holder["position"] in "VC":
                vc_rate = vc_rate + float(str(holder["rate"])[:-1])
        return str(round(vc_rate, 2)) + "%"

    def i_bank_name(self):
        return self.data["bank"]["name"]
    
    def i_bank_rate(self):
        return self.data["bank"]["rate"]
    
    def i_grade(self):
        return self.data["grade"]
    
    def i_market(self):
        return self.data["market"]
    
    def i_assumed_price(self):
        min = self.data["expectedProfitAfterTD"]["tdPrice"]["min"]
        max = self.data["expectedProfitAfterTD"]["tdPrice"]["max"]
        return str(min) + "円　〜　" + str(max) + "円"
    
    def i_initial_price_forecast(self):
        min = self.data["expectedProfitAfterTD"]["initPriceEx"]["min"]
        max = self.data["expectedProfitAfterTD"]["initPriceEx"]["max"]
        return str(min) + "円　〜　" + str(max) + "円"

    def i_expected_profit(self):
        min = self.data["expectedProfitAfterTD"]["exProfit"]["min"]
        max = self.data["expectedProfitAfterTD"]["exProfit"]["max"]
        return str(min) + "円　〜　" + str(max) + "円"

    def i_abs_amount_aft_td(self):
        new_shares = self.data["pubOfferedShares"] + self.data["oaShared"] + self.data["sellShares"]
        min = self.data["expectedProfitAfterTD"]["tdPrice"]["min"]
        max = self.data["expectedProfitAfterTD"]["tdPrice"]["max"]
        return self.ctrl_str.cnv_ja_money(min * new_shares) + "　〜　" + self.ctrl_str.cnv_ja_money(max * new_shares)

    def i_market_capitalization_aft_td(self):
        new_shares = self.data["pubOfferedShares"] + self.data["oaShared"] + self.data["sellShares"]
        min = self.data["expectedProfitAfterTD"]["tdPrice"]["min"]
        max = self.data["expectedProfitAfterTD"]["tdPrice"]["max"]
        return self.ctrl_str.cnv_ja_money(min * new_shares) + "　〜　" + self.ctrl_str.cnv_ja_money(max * new_shares)

    def i_init_price(self):
        return str(self.data["initPrice"]) + "円"
    
    def i_init_price_sell_profit(self):
        return str(self.data["InitPriceSellProfit"]) + "円"
    
    def i_rf_rate(self):
        return str(self.data["rfRate"]) + "%"
    
    def i_indpndnt_fin_info(self):
        return self.data["indpndntFinInfo"]
    
    def i_cnsldtd_fin_info(self):
        return self.data["cnsldtdFinInfo"]

    def i_share_holders(self):
        return self.data["shareholders"]