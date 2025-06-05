import customtkinter as ctk
from PRCPayrollSystem.Main.excelImportPage import ExcelImportPage
from PRCPayrollSystem.Main.importEmployee import ImportEmployeePage
from PRCPayrollSystem.Main.reportsPage import ReportsPage
from PRCPayrollSystem.Main.generatePayslip import GeneratePayslipPage
from PRCPayrollSystem.Main.historyPage import HistoryPage

class AppController(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Payroll System")
        self.geometry("950x600")
        self.frames = {}
        from PRCPayrollSystem.Main.Main import MainMenu  # Import here to avoid circular import
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
