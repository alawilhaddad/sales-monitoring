from win.var import *


def main():
    ws = window.winfo_screenwidth()
    hs = window.winfo_screenheight()
    x = (ws / 2) - (window_width / 2)
    y = (hs / 2) - (window_height / 2)
    window.geometry(f"{window_width}x{window_height}+{int(x)}+{int(y)}")
    canvas.place(x=0, y=0)
    # home()
    window.mainloop()


if __name__ == "__main__":
    main()
