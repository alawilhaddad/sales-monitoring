import os
from tkinter.filedialog import askdirectory


if os.path.isfile("setting.txt"):
    with open("setting.txt", "r") as file:
        lines = file.readlines()
        default_load_dir = lines[0].split("=")[1]
        default_save_dir = lines[1].split("=")[1]
        print(default_load_dir)
        print(default_save_dir)
else:
    print("no")
    default_open_directory = askdirectory()
    default_save_directory = askdirectory()
    with open("setting.txt", "w") as file:
        file.write(f"open={default_open_directory}\n")
        file.write(f"save={default_save_directory}")
