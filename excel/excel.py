from tkinter.filedialog import *
from openpyxl import Workbook


class Excel:
    def __init__(self):
        self.path = None
        self.filename = None
        self.wb = None
        self.ws = []
        self.sheet = None
        self.active_sheet = None


class Load(Excel):
    def __init__(self):
        super().__init__()
        self.image = None
        self.label = None
        self.state = False
        self.button = None
        self.options = None
        self.var = None
        self.cell = None
        self.month = []

    def load(self):
        self.path = askopenfilename(initialdir='C:/Users/mohin/Downloads/',
                                    title='Open File',
                                    filetypes=(("Excel File", "*xlsx"),
                                               ("All Files", "*.*")))
        self.filename = self.path.split('/')[-1]
        if self.path == '':
            return


class Create(Excel):
    def __init__(self):
        super().__init__()
        self.pc_sheet = None
        self.odoo_sheet = None

    def create_excel(self, pc, odoo):
        self.wb = Workbook()
        self.sheet = self.wb.sheetnames
        self.pc_sheet = self.wb.create_sheet(pc.sheet.get())
        self.odoo_sheet = self.wb.create_sheet(odoo.sheet.get())

