import tkinter as tk

__all__ = ["HoverButton"]


class HoverButton(tk.Button):
    def __init__(self, master=None, cnf=None, **kwargs):
        self.kwargs = kwargs
        if cnf is None:
            cnf = {}
        cnf = tk._cnfmerge((cnf, kwargs))
        super(HoverButton, self).__init__(master=master, cnf=cnf, **kwargs)

        self.DefaultBG = self.kwargs['bg']
        self.HoverBG = self.kwargs['activeback']
        self.bind_class(self, "<Enter>", self.Enter)
        self.bind_class(self, "<Leave>", self.Leave)

    def change_color_bind(self, DefaultBG, HoverBG, ActiveBG):
        self.configure(bg=DefaultBG, activebackground=ActiveBG)
        self.kwargs['bg'] = self.DefaultBG = DefaultBG
        self.kwargs['activeback'] = self.HoverBG = HoverBG

    def Enter(self, event):
        self['bg'] = self.HoverBG

    def Leave(self, event):
        self['bg'] = self.DefaultBG
