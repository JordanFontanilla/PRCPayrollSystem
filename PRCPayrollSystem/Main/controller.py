import customtkinter as ctk
from excelImportPage import ExcelImportPage
from importEmployee import ImportEmployeePage
from reportsPage import ReportsPage
from generatePayslip import GeneratePayslipPage
from historyPage import HistoryPage

class AppController(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Payroll System")
        self.geometry("950x600")
        self.frames = {}
        from Main import MainMenu  
        for F in (MainMenu, ExcelImportPage, ImportEmployeePage, ReportsPage, GeneratePayslipPage, HistoryPage):
            frame = F(self, self)
            self.frames[F.__name__] = frame
            frame.place(relwidth=1, relheight=1)
        self.show_frame("MainMenu")

    def show_frame(self, page_name):
        frame = self.frames[page_name]
        frame.tkraise()

    def get_page(self, page_name):
        return self.frames.get(page_name)
