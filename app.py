from win.config import Win
from tkinter import *
from tkinter import ttk
from tkinter.messagebox import *
from tkinter.filedialog import *
from excel.helper import convert_slash, remove_password_xlsx
from openpyxl import *
from zipfile import BadZipfile
from excel.var import pc, odoo
from win.funct import start, _quit


def _open(self, cvs):
    self.path = askopenfilename(initialdir='C:/Users/mohin/Downloads/',
                                title='Open File',
                                filetypes=(("Excel File", "*xlsx"), ("All Files", "*.*")))
    self.filename = self.path.split('/')[-1]
    self.path = convert_slash(self.path)
    if self.path is None:
        return
    else:
        try:
            self.wb = load_workbook(self.path, data_only=True)
            self.sheet_list = self.wb.sheetnames
            self.active_sheet = self.sheet_list[0]
            cvs.itemconfigure(self.label, text=self.filename)
            self.options['values'] = self.sheet_list
            self.options['state'] = 'readonly'
            self.sheet.set(self.sheet_list[0])
            self.state = True
        except BadZipfile:
            showinfo(title='Password Protected', message='Your file is password protected.')
            self.path = remove_password_xlsx(self.path, input("password: "))
            self.filename = self.path.split('\\')[-1]
            self.wb = load_workbook(self.path, data_only=True)
            self.sheet_list = self.wb.sheetnames
            self.active_sheet = self.sheet_list[0]
            canvas.itemconfigure(self.label, text=self.filename)
            self.options['values'] = self.sheet_list
            self.options['state'] = 'readonly'
            self.sheet.set(self.sheet_list[0])
            self.state = True
        if pc.state and odoo.state:
            start_button.config(state=NORMAL)


def quit_help():
    help_window.destroy()
    window.deiconify()


def _help():
    # noinspection PyGlobalUndefined
    global help_window
    window.withdraw()
    help_window = Toplevel()
    help_width = help_window.winfo_screenwidth()
    help_height = help_window.winfo_screenheight()
    help_position_x = (help_width / 2) - 200
    help_position_y = (help_height / 2) - 150
    help_window.geometry(f"400x300+{int(help_position_x)}+{int(help_position_y)}")
    help_window.configure(bg="#ffffff")
    help_window.overrideredirect(True)

    help_canvas = Canvas(
        help_window,
        bg="#ffffff",
        height=300,
        width=400,
        bd=0,
        highlightthickness=0,
        relief="ridge")
    help_canvas.place(x=0, y=0)

    help_background_img = PhotoImage(file=f"win/img/help_background.png")
    help_canvas.create_image(
        200.0, 150.0,
        image=help_background_img)

    img0 = PhotoImage(file=f"win/img/help_close.png")
    b0 = Button(
        help_window,
        image=img0,
        borderwidth=0,
        highlightthickness=0,
        command=quit_help,
        relief="flat")
    b0.place(
        x=375, y=7,
        width=19,
        height=19)

    window.resizable(False, False)
    help_window.mainloop()


if __name__ == "__main__":
    window = Win()
    ws = window.winfo_screenwidth()
    hs = window.winfo_screenheight()
    x = (ws / 2) - 275
    y = (hs / 2) - 225
    window.geometry(f"550x450+{int(x)}+{int(y)}")
    window.configure(bg="#ffffff")

    canvas = Canvas(
        window,
        bg="#ffffff",
        height=450,
        width=550,
        bd=0,
        highlightthickness=0,
        relief="ridge")
    canvas.place(x=0, y=0)

    background_img = PhotoImage(file=f"win/img/background.png")
    canvas.create_image(
        275.0, 225.0,
        image=background_img)

    help_icon = PhotoImage(file=f"win/img/img2.png")
    help_button = Button(
        image=help_icon,
        borderwidth=0,
        highlightthickness=0,
        command=_help,
        relief="flat"
    )
    help_button.place(
        x=3, y=420,
        width=24,
        height=24)

    close_image = PhotoImage(file=f"win/img/img4.png")
    close_button = Button(
        image=close_image,
        borderwidth=0,
        highlightthickness=0,
        command=_quit,
        relief="flat")
    close_button.place(
        x=524, y=6,
        width=19,
        height=19)

    start_image = PhotoImage(file=f"win/img/img3.png")
    start_button = Button(
        image=start_image,
        borderwidth=0,
        highlightthickness=0,
        command=start,
        relief="flat",
        state=DISABLED)
    start_button.place(
        x=418, y=349,
        width=88,
        height=28)

    pc.image = PhotoImage(file=f"win/img/img0.png")
    pc.button = Button(
        image=pc.image,
        borderwidth=0,
        highlightthickness=0,
        relief="flat")
    pc.button['command'] = lambda source=pc, cv=canvas: _open(source, cv)
    pc.button.place(
        x=74, y=219,
        width=284,
        height=28)
    pc.label = canvas.create_text(500, 254, text='', font=('calibri light', 8), fill='white', anchor=E)
    pc.sheet = StringVar(window)
    pc.options = ttk.Combobox(
        window,
        textvariable=pc.sheet,
        width=12,
        value=pc.ws,
        font=('karla tamil upright', 9),
        justify='right',
        state="readonly")
    pc.options.place(x=352, y=219, width=155, height=28)
    pc.options['state'] = 'disabled'

    odoo.image = PhotoImage(file=f"win/img/img1.png")
    odoo.button = Button(
        image=odoo.image,
        borderwidth=0,
        highlightthickness=0,
        relief="flat")
    odoo.button['command'] = lambda source=odoo, cv=canvas: _open(source, cv)
    odoo.button.place(
        x=74, y=275,
        width=284,
        height=28)
    odoo.label = canvas.create_text(500, 309, text='', font=('calibri light', 8), fill='white', anchor=E)
    odoo.sheet = StringVar(window)
    odoo.options = ttk.Combobox(
        window,
        textvariable=odoo.sheet,
        width=12,
        value=odoo.ws,
        font=('karla tamil upright', 9),
        justify='right',
        state="readonly")
    odoo.options.place(x=352, y=275, width=155, height=28)
    odoo.options['state'] = 'disabled'
    window.mainloop()
