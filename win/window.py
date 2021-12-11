from tkinter import *
from tkinter import ttk
from win.controller import *


class Main:
# Initiate all widget =================================================================================================
    def __init__(self, window, window_height=600, window_width=800):
        # Determine window properties
        self.window = window
        self.screen_w = window.winfo_screenwidth()
        self.screen_h = window.winfo_screenheight()
        self.pos_x = (self.screen_w / 2) - (window_width / 2)
        self.pos_y = (self.screen_h / 2) - (window_height / 2)
        self.window.geometry(f"{window_width}x{window_height}+{int(self.pos_x)}+{int(self.pos_y)}")

# Main Window ---------------------------------------------------------------------------------------------------------

        # Initiate base canvas
        self.canvas = Canvas(
            window,
            bg="#ffffff",
            height=600,
            width=800,
            bd=0,
            highlightthickness=0,
            relief="ridge")
        self.canvas.place(x=0, y=0)

        # Initiate background
        self.background_img = PhotoImage(file=f"win/img/background.png")
        self.canvas.create_image(400, 300, image=self.background_img)

        # Initiate close button
        self.close_image = PhotoImage(file=f"win/img/close_icon.png")
        self.close_button = Button(
            image=self.close_image,
            borderwidth=0,
            highlightthickness=0,
            command=exit_app,
            relief="flat")
        self.close_button.place(
            x=750, y=30,
            width=20,
            height=20)

        # Initiate copyright and versioning
        self.canvas.create_text(
            155.0, 562.0,
            text="Haddaddegusti 2021 | v. 2.3.0",
            fill="#ffffff",
            anchor="w",
            font=("Taprom", int(10.0)))

# Menu Bar ------------------------------------------------------------------------------------------------------------

        # Initiate home button
        self.home_icon = PhotoImage(file=f"win/img/home_icon.png")
        self.home_button = Button(
            image=self.home_icon,
            borderwidth=0,
            highlightthickness=0,
            command=lambda menu=self : home_show(menu),
            relief="flat")
        self.home_button.place(
            x=30, y=370,
            width=60,
            height=60)

        # Initiate setting button
        self.setting_icon = PhotoImage(file=f"win/img/setting_icon.png")
        self.setting_button = Button(
            image=self.setting_icon,
            borderwidth=0,
            highlightthickness=0,
            command=lambda menu=self : setting_show(menu),
            relief="flat")
        self.setting_button.place(
            x=30, y=440,
            width=60,
            height=60)

        # Initiate help button
        self.help_icon = PhotoImage(file=f"win/img/help_icon.png")
        self.help_button = Button(
            image=self.help_icon,
            borderwidth=0,
            highlightthickness=0,
            command=lambda menu=self : guide_show(menu),
            relief="flat")
        self.help_button.place(
            x=30, y=510,
            width=60,
            height=60)

# Home_Section --------------------------------------------------------------------------------------------------------

        # Initiate main title
        self.main_title_img = PhotoImage(file=f"win/img/main_title.png")

        # Initiate label background
        self.label_canvas = PhotoImage(file=f"win/img/label_canvas.png")

        self.pc_state = False # State for start condition
        self.pc_label = self.canvas.create_text(
            150.0, 322.0,
            fill="#ffffff",
            anchor="w",
            font=("KarlaTamilUpright-Regular", 9),
            tags="home")

        # Initiate "load pc monitoring" button
        self.pc_image = PhotoImage(file=f"win/img/pc_icon.png")
        self.pc_button = Button(
            image=self.pc_image,
            borderwidth=0,
            highlightthickness=0,
            command=lambda app=self, source=pc : open_excel(self, pc, odoo, self.pc_options, self.pc_label),
            relief="flat")

        # Initiate combobox for pc
        self.pc_selected_sheet = StringVar(window)
        self.pc_options = ttk.Combobox(
            window,
            textvariable=pc.sheet_list,
            width=12,
            value=self.pc_selected_sheet,
            font=('karla tamil upright', 9),
            justify='right',
            state="readonly")
        self.pc_options['state'] = 'disabled'

        # Initiate reload image
        self.reload_image = PhotoImage(file=f"win/img/reload_icon.png")

        # Initiate reload button for PC
        self.reload_pc_button = Button(
            image=self.reload_image,
            borderwidth=0,
            highlightthickness=0,
            command=lambda main=self: reload(main, pc, self.pc_options, self.pc_label),
            relief="flat")

        # Initiate textbox for PC monitoring password
        self.password = StringVar(window)
        self.entry0_img = PhotoImage(file=f"win/img/img_textBox0.png")
        self.entry0 = Entry(
            bd=0,
            fg="#7e91d6",
            bg="#ffffff",
            textvariable=self.password,
            highlightthickness=0,
            font=("KarlaTamilUpright-Regular", 11),
            justify="center",
            show="*")

        # State for start condition
        self.odoo_state = False

        self.odoo_label = self.canvas.create_text(
            150.0, 401.0,
            fill="#ffffff",
            anchor="w",
            font=("KarlaTamilUpright-Regular", 9),
            tags="home")

        # Initiate "load odoo report" button
        self.odoo_image = PhotoImage(file=f"win/img/odoo_icon.png")
        self.odoo_button = Button(
            image=self.odoo_image,
            borderwidth=0,
            highlightthickness=0,
            command=lambda app=self: open_excel(app, odoo, pc, self.odoo_options, self.odoo_label),
            relief="flat")

        # Initiate reload button for Odoo
        self.reload_odoo_button = Button(
            image=self.reload_image,
            borderwidth=0,
            highlightthickness=0,
            command=lambda main=self: reload(main, odoo, self.odoo_options, self.odoo_label),
            relief="flat")

        # Initiate "load odoo report" button
        self.odoo_selected_sheet = StringVar(window)
        self.odoo_options = ttk.Combobox(
            window,
            textvariable=odoo.sheet_list,
            width=12,
            value=self.odoo_selected_sheet,
            font=('karla tamil upright', 9),
            justify='right',
            state="readonly")
        self.odoo_options['state'] = 'disabled'

        # Initiate start button
        self.start_image = PhotoImage(file=f"win/img/start_icon.png")
        self.start_button = Button(
            image=self.start_image,
            borderwidth=0,
            highlightthickness=0,
            command=lambda app=self: start(app, new, pc, odoo),
            state=DISABLED,
            relief="flat")

# Help_Section --------------------------------------------------------------------------------------------------------

        # Initiate help title
        self.help_title_img = PhotoImage(file=f"win/img/help_title.png")

# Setting_Section -----------------------------------------------------------------------------------------------------

        # Initiate help setting
        self.setting_title_img = PhotoImage(file=f"win/img/setting_title.png")

        # Initiate open default open directory
        self.open_d_image = PhotoImage(file=f"win/img/open_directory.png")
        self.open_d_button = Button(
            image=self.open_d_image,
            borderwidth=0,
            highlightthickness=0,
            command=lambda: print("click"),
            relief="flat")

        # Initiate open default close directory
        self.save_d_image = PhotoImage(file=f"win/img/save_directory.png")
        self.save_d_button = Button(
            image=self.save_d_image,
            borderwidth=0,
            highlightthickness=0,
            command=lambda: print("click"),
            relief="flat")

        # Initiate round background
        self.label_round = PhotoImage(file=f"win/img/label_round.png")
