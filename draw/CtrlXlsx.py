from draw import DataSet
import openpyxl

class CtrlXlsx():
    
    def __init__(self, path) -> None:
        self.out_path = path
        self.wb = openpyxl.Workbook()
        self.wb.save(self.out_path)
        pass
    
    def writeXlsx(self, sheet_name, draw_obj):
        ws = self.wb.active
        ws.title = sheet_name

        for row in draw_obj:
            ws.append(row)
        
        self.wb.save(self.out_path)
    