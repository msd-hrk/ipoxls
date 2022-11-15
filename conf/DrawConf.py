from conf.Config import Config
from common.ctrlStr import CtrlStr
class DrawConf():
    
    def __init__(self) -> None:
        conf = Config()
        self.draw_conf_path = conf.draw_conf_path
        
        self.draw_arry = []
        f = open(self.draw_conf_path)
        lines = f.readlines()
        for line in lines:
            if line[0:1] == '#':
                continue
            
            conf_arry = CtrlStr.remove(line[line.index('=') + 1:], '\n', '\r\n').split(",")
            self.draw_arry.append(conf_arry)
            
        print(self.draw_arry)
        pass