import openpyxl.worksheet
from openpyxl.styles.borders import Border, Side

class FrameDesigner():
    
    def __init__(self) -> None:
        # Border
        side = Side(style="thin", color="000000")
        self.black_thin = Border(left=side, right=side, top=side, bottom=side)
        self.black_thin_rtb = Border(right=side, top=side, bottom=side)
        self.black_thin_ltb = Border(left=side, top=side, bottom=side)
        self.black_thin_lrb = Border(left=side, right=side, bottom=side)
        self.black_thin_lrt = Border(left=side, right=side, top=side)
        self.black_thin_lt = Border(left=side, top=side)
        self.black_thin_rt = Border(right=side, top=side)
        self.black_thin_lb = Border(left=side, bottom=side)
        self.black_thin_rb = Border(right=side, bottom=side)
        self.black_thin_b = Border(bottom=side)
        self.black_thin_t = Border(top=side)
        self.sheet = openpyxl.worksheet

    def create_individual_frame(self, sheet: openpyxl.worksheet):
        self.sheet = sheet
        self.company_summary(sheet)
        self.ipo_info(sheet)
        self.listing_forecast(sheet)
        self.financial_info(sheet)
        self.news_list(sheet)
        self.lockup_info(sheet)
        pass
    
    def company_summary(self, sheet: openpyxl.worksheet):
        self.draw_merge_frame("A1:O1", self.black_thin)
        self.draw_merge_frame("A3:O15", self.black_thin)
        self.draw_merge_frame("A17:G20", self.black_thin)
        pass
    
    def ipo_info(self, sheet: openpyxl.worksheet):
        self.draw_merge_frame("A22:G31", self.black_thin)
        sheet['C24'].border = self.black_thin_lt
        sheet['C25'].border = self.black_thin_lb
        sheet['G24'].border = self.black_thin_rt
        sheet['G25'].border = self.black_thin_rb
        self.draw_merge_frame("D24:F24", self.black_thin_t)
        self.draw_merge_frame("D25:F25", self.black_thin_b)
        pass
    
    def listing_forecast(self, sheet: openpyxl.worksheet):
        self.draw_merge_frame("A33:G38", self.black_thin)
        self.draw_merge_frame("A39:D39", self.black_thin)
    
    def financial_info(self, sheet: openpyxl.worksheet):
        self.draw_merge_frame("A41:G49", self.black_thin)
        self.draw_merge_frame("A51:G58", self.black_thin)

    def news_list(self, sheet: openpyxl.worksheet):
        self.draw_merge_frame("A60:G80", self.black_thin)
        
    def lockup_info(self, sheet: openpyxl.worksheet):
        cell_row = 82
        self.draw_merge_frame("A" + str(cell_row) + ":G" + str(cell_row + 40), self.black_thin)
        
    def draw_merge_frame(self, range: str, border_type: Border):
        cells = self.sheet[range]
        for row in cells:
            for cell in row:
                cell.border = border_type
    