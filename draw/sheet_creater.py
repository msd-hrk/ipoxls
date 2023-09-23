import openpyxl
from openpyxl.styles.borders import Border, Side
from conf import config
from dbquery import access_db
from draw import frame_designer, binding_designer, paint_designer, data_writer
from common import ctrl_str

class SheetCreater():
    
    def __init__(self, wb: openpyxl.Workbook) -> None:
        # load config
        self.conf = config.Config()
        # create FrameDesigner instance
        self.frame_d = frame_designer.FrameDesigner()
        # create BindingDesigner instance
        self.binding_d = binding_designer.BindingDesigner()
        # create PaintDesigner instance
        self.paint_d = paint_designer.PaintDesigner()
        # create BindingDesigner instance
        self.data_w = data_writer.DataWriter()
        
        # keep Workbook obj
        self.workbook = wb
    
    def add_individual_sheet(self, sec_num: str) -> None:
        # get data from security num
        data = access_db.AccessDb().find_one_from_securities_no(sec_num)
        # sheet name
        sheet_name = data["securitiesNo"] + "_" + data["company"]
        # creat a new sheet
        sheet = self.workbook.create_sheet(sheet_name, len(self.workbook.sheetnames))

        self.binding_d.create_individual_binding(sheet)
        self.frame_d.create_individual_frame(sheet)
        self.paint_d.paint_individual(sheet)
        self.data_w.write_individual(data, sheet)
        pass
    
