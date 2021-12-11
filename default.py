import os
from excel import var


def default():
    if os.path.isfile("config.txt"):
        try:
            with open("config.txt", "r+") as file:
                lines = file.readlines()
                open_dir = lines[0].split("=")[1].strip("\n")
                save_dir = lines[1].split("=")[1]
                if os.path.isdir(open_dir) and os.path.isdir(save_dir):
                    var.default_open_dir = open_dir
                    var.default_save_dir = save_dir
                if not os.path.isdir(open_dir):
                    var.default_save_dir = open_dir
                    var.default_open_dir = ""
                    lines[0] = 'open= \n'
                if not os.path.isdir(save_dir):
                    var.default_open_dir = save_dir
                    var.default_save_dir = ""
                    lines[1] = 'save= '
                if not os.path.isdir(open_dir) and not os.path.isdir(save_dir):
                    var.default_open_dir = ""
                    var.default_save_dir = ""
                    lines[0] = "open= \n"
                    lines[1] = "save= "
            with open("config.txt", "w") as file:
                file.writelines(lines)
        except IndexError:
            with open("config.txt", "w") as file:
                file.write(f"open= \n")
                file.write(f"save= ")
            var.default_save_dir = ""
            var.default_open_dir = ""

    else:
        with open("config.txt", "w") as file:
            file.write(f"open= \n")
            file.write(f"save= ")
        var.default_save_dir = ""
        var.default_open_dir = ""


