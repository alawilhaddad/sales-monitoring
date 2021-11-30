from win32com.client import Dispatch


def remove_password_xlsx(filename, pw_str):
    xcl = Dispatch("Excel.Application")
    wb = xcl.Workbooks.Open(filename, False, False, None, pw_str)
    xcl.DisplayAlerts = False
    filename_split = filename.split('.')
    filename_split[-2] += '_unlocked'
    filename = '.'.join(filename_split)
    wb.SaveAs(filename, None, '', '')
    xcl.Quit()
    return filename


def convert_slash(path):
    path_list = list(path)
    for char in path_list:
        if char == '/':
            index = path_list.index(char)
            path_list[index] = '\\'
    path = ''.join(path_list)
    return path
