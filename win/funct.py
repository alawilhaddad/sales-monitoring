import sys
from excel.var import new, pc, odoo
from excel.copy_sheet_data import copy_sheet
from excel.adjustment.eni import eni
from excel.adjustment.phkt import phkt
from excel.adjustment.phm_ade import phm_ade
from datetime import *
from tkinter.messagebox import showinfo
from tkinter.filedialog import asksaveasfilename
from os import startfile


def _quit():
    quit()
    sys.exit()


def start():
    new.create_excel(pc, odoo)
    pc.active_sheet = pc.wb[pc.sheet.get()]
    copy_sheet(pc.active_sheet, new.pc_sheet)
    odoo.active_sheet = odoo.wb[odoo.sheet.get()]
    copy_sheet(odoo.active_sheet, new.odoo_sheet)
    if 'Sheet' in new.sheet:
        new.wb.remove(new.wb['Sheet'])
    if new.wb.sheetnames[1] == '3250':
        eni()
        project = 'ENI'
    elif new.wb.sheetnames[1] == '3235':
        phm_ade()
        project = 'PHM'
    elif new.wb.sheetnames[1] == '3247':
        phkt()
        project = 'PHKT'

    dates = datetime.now()

    path = asksaveasfilename(
        initialdir='C:/Users/mohin/Downloads/',
        title='Save File',
        initialfile=f'{project}-{new.wb.sheetnames[1]}-{dates.strftime("%Y%m%d")}-0{int(dates.strftime("%d"))//7}\
        -Monitoring-v01.xlsx',
        filetypes=(("Excel File", "*xlsx"), ("All Files", "*.*")))
    if path == '':
        return
    else:
        new.wb.save(path)
        showinfo(title='Done', message='Done!')
        startfile(path)
    quit()


if __name__ == "__main__":
    start()
