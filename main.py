from win.config import Win
from win.window import Main
from win.controller import home_show
from default import default
from excel.var import pc, odoo


if __name__ == "__main__":
    root = Win()
    default()
    main = Main(root)
    home_show(main, pc, odoo)
    root.mainloop()
    