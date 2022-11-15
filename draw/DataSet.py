
class DataSet():
    """
    mongoDB登録データを保持するクラス
    """
    def __init__(self, data):
        # 会社名
        self.company = data["company"]
        # 証券コード
        self.securitiesNo = data["securitiesNo"]
        # ブックビル期間
        self.bookbuilding = data["bookbuilding"]
        # 上場日
        self.listingDate = data["listingDate"]
        # 評価
        self.grade = data["grade"]
        # 仮条件決定日
        self.tdd = data["tdd"]
        # 吸収金額合計
        self.absorbentAmount = data["absorbentAmount"]
        # 銀行情報（主幹事データ）
        self.bank = data["bank"]
        # 設立
        self.build = data["build"]
        # 事業内容
        self.business = data["business"]
        # 事業内容詳細
        self.businessDetail = data["businessDetail"]
        # 資本金
        self.capital = data["capital"]
        # 業種
        self.category = data["category"]
        # 連結財務情報
        self.cnsldtdFinInfo = data["cnsldtdFinInfo"]
        # 従業員情報
        self.employee = data["employee"]
        # 予想利益（仮条件決定後）
        self.expectedProfitAfterTD = data["expectedProfitAfterTD"]
        # 予想利益（仮条件決定前）
        self.expectedProfitBeforeTD = data["expectedProfitBeforeTD"]
        # 株主数
        self.holder_num = data["holder_num"]
        # 単独財務情報
        self.indpndntFinInfo = data["indpndntFinInfo"]
        # 発行済株式数
        self.issuedShares = data["issuedShares"]
        # 市場
        self.market = data["market"]
        # 時価総額
        self.mrktcptlz = data["mrktcptlz"]
        # OA分
        self.oaShared = data["oaShared"]
        # オファリングレシオ
        self.oar = data["oar"]
        # 公募価格決定日
        self.pdd = data["pdd"]
        # 公募株数
        self.pubOfferedShares = data["pubOfferedShares"]
        # 購入期間
        self.purchasePeriod = data["purchasePeriod"]
        # 売出株数
        self.sellShares = data["sellShares"]
        # 上位10株主情報
        self.shareholders = data["shareholders"]
        # 単元株
        self.unitShare = data["unitShare"]
        # サイト
        self.web = data["web"]
        # 当選口数
        self.winningNum = data["winningNum"]
        # 公募価格
        self.pubOfferPrice = data["pubOfferPrice"]
        # 初値売り損益
        self.InitPriceSellProfit = data["InitPriceSellProfit"]
        # 初値
        self.initPrice = data["initPrice"]
        # 騰落率
        self.rfRate = data["rfRate"]
        # 日毎データ
        self.priceDiary = data["priceDiary"]
        
        pass