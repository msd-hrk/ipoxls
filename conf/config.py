import json

class Config():
    """
    設定値を保持するクラス
    """
    def __init__(self):
        """
        config.jsonから値を取得する
        """
        config_path = "config/ipocc_config.json"
        config = json.load(open(config_path , "r"))
        self.mongo_url = config["MONGODB_URL"]  
