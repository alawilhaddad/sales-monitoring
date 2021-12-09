from win.config import Win
from win.window import Main
from win.controller import home_show


if __name__ == "__main__":
    root = Win()
    main = Main(root)
    home_show(main)
    root.mainloop()
    