from win32com.client import Dispatch


def unlock_excel(path, pw_str):
    path = convert_slash(path)
    xcl = Dispatch("Excel.Application")

    wb = xcl.Workbooks.Open(path, False, False, None, pw_str)
    xcl.DisplayAlerts = False
    filename_split = path.split('.')
    filename_split[-2] += '_unlocked'
    path = '.'.join(filename_split)
    wb.SaveAs(path, None, '', '')

    xcl.Quit()
    return path


def convert_slash(path):
    path_list = list(path)
    for char in path_list:
        if char == '/':
            index = path_list.index(char)
            path_list[index] = '\\'
    path = ''.join(path_list)
    return path


if __name__ == "__main__":
    unlock_excel("C:/Users/mohin/Downloads/Progress Control OF PHKT - Week 46 OF 2021.xlsx", input("Password: "))
