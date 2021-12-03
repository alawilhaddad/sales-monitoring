from win.config import Win
from tkinter import *
from tkinter import ttk
from tkinter.messagebox import *
from tkinter.filedialog import *
from excel.helper import convert_slash, remove_password_xlsx
from openpyxl import *
from zipfile import BadZipfile
from excel.var import pc, odoo
from win.funct import *
from win.var import *


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
    x = (ws / 2) - (window_width/2)
    y = (hs / 2) - (window_height/2)
    window.geometry(f"{window_width}x{window_height}+{int(x)}+{int(y)}")

    canvas = Canvas(
        window,
        bg="#ffffff",
        height=600,
        width=800,
        bd=0,
        highlightthickness=0,
        relief="ridge")
    canvas.place(x=0, y=0)

    background_img = PhotoImage(file=f"win/img/background.png")
    canvas.create_image(
        400, 300,
        image=background_img)

    main_title_img = PhotoImage(file=f"win/img/main_title.png")
    canvas.create_image(
        140, 92,
        anchor="nw",
        image=main_title_img)

    label_canvas = PhotoImage(file=f"win/img/label_canvas.png")
    canvas.create_image(
        140, 310,
        anchor="nw",
        image=label_canvas)

    canvas.create_image(
        140, 390,
        anchor="nw",
        image=label_canvas)

    canvas.create_rectangle(
        370, 350, 370 + 120, 350 + 40,
        fill="#81c0f2",
        outline="")

    canvas.create_rectangle(
        370, 270, 370 + 120, 270 + 40,
        fill="#81c0f2",
        outline="")

    canvas.create_text(
        405.0, 290.0,
        text="Pass: ",
        fill="#ffffff",
        font=("KarlaTamilUpright-Bold", int(11.0)))

    canvas.create_text(
        155.0, 562.0,
        text="Haddaddegusti 2021 | v. 2.3.0",
        fill="#ffffff",
        anchor="w",
        font=("Taprom", int(10.0)))

    home_icon = PhotoImage(file=f"win/img/home_icon.png")
    home_button = Button(
        image=home_icon,
        borderwidth=0,
        highlightthickness=0,
        command=home,
        relief="flat")
    home_button.place(
        x=30, y=370,
        width=60,
        height=60)

    setting_icon = PhotoImage(file=f"win/img/setting_icon.png")
    setting_button = Button(
        image=setting_icon,
        borderwidth=0,
        highlightthickness=0,
        command=setting,
        relief="flat")
    setting_button.place(
        x=30, y=440,
        width=60,
        height=60)

    help_icon = PhotoImage(file=f"win/img/help_icon.png")
    help_button = Button(
        image=help_icon,
        borderwidth=0,
        highlightthickness=0,
        command=guide,
        relief="flat")
    help_button.place(
        x=30, y=510,
        width=60,
        height=60)

    close_image = PhotoImage(file=f"win/img/close_icon.png")
    close_button = Button(
        image=close_image,
        borderwidth=0,
        highlightthickness=0,
        command=exit_app,
        relief="flat")
    close_button.place(
        x=750, y=30,
        width=20,
        height=20)

    pc.image = PhotoImage(file=f"win/img/pc_icon.png")
    pc.button = Button(
        image=pc.image,
        borderwidth=0,
        highlightthickness=0,
        relief="flat")
    pc.button['command'] = lambda source=pc, cv=canvas: _open(source, cv)
    pc.button.place(
        x=140, y=270,
        width=230,
        height=40)
    pc.label = canvas.create_text(
        150.0, 322.0,
        fill="#ffffff",
        anchor="w",
        font=("KarlaTamilUpright-Regular", int(9.0)))
    pc.sheet = StringVar(window)
    pc.options = ttk.Combobox(
        window,
        textvariable=pc.sheet,
        width=12,
        value=pc.ws,
        font=('karla tamil upright', 9),
        justify='right',
        state="readonly")
    pc.options.place(x=490, y=270, width=120, height=40)
    pc.options['state'] = 'disabled'

    entry0_img = PhotoImage(file=f"win/img/img_textBox0.png")
    entry0_bg = canvas.create_image(
        448.0, 290.0,
        image=entry0_img)
    entry0 = Entry(
        bd=0,
        fg="#7e91d6",
        bg="#ffffff",
        highlightthickness=0,
        font=("KarlaTamilUpright-Regular", int(11.0)))
    entry0.place(
        x=430.0, y=280,
        width=36.0,
        height=22)

    odoo.image = PhotoImage(file=f"win/img/odoo_icon.png")
    odoo.button = Button(
        image=odoo.image,
        borderwidth=0,
        highlightthickness=0,
        relief="flat")
    odoo.button['command'] = lambda source=odoo, cv=canvas: _open(source, cv)
    odoo.button.place(
        x=140, y=350,
        width=230,
        height=40)
    odoo.label = canvas.create_text(
        150.0, 401.0,
        fill="#ffffff",
        anchor="w",
        font=("KarlaTamilUpright-Regular", int(9.0)))
    odoo.sheet = StringVar(window)
    odoo.options = ttk.Combobox(
        window,
        textvariable=odoo.sheet,
        width=12,
        value=odoo.ws,
        font=('karla tamil upright', 9),
        justify='right',
        state="readonly")
    odoo.options.place(x=490, y=350, width=120, height=40)
    odoo.options['state'] = 'disabled'

    start_image = PhotoImage(file=f"win/img/start_icon.png")
    start_button = Button(
        image=start_image,
        borderwidth=0,
        highlightthickness=0,
        command=start,
        relief="flat",
        state=DISABLED)
    start_button.place(
        x=660, y=460,
        width=100,
        height=40)
    window.mainloop()
