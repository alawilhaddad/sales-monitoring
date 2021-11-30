from tkinter import Tk


class Win(Tk):
    def __init__(self):
        super().__init__()
        super().overrideredirect(True)
        self.offset_x = 0
        self.offset_y = 0
        super().bind("<Button-1>", self.click_win)
        super().bind("<B1-Motion>", self.drag_win)

    def drag_win(self, event):
        psx = super().winfo_pointerx() - self.offset_x
        psy = super().winfo_pointery() - self.offset_y
        super().geometry(f"+{psx}+{psy}")
        return event

    def click_win(self, event):
        self.offset_x = super().winfo_pointerx() - super().winfo_rootx()
        self.offset_y = super().winfo_pointery() - super().winfo_rooty()
        return event
