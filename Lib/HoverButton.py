import tkinter as tk

__all__ = ["HoverButton"]


class HoverButton(tk.Button):
    def __init__(self, master=None, cnf=None, **kwargs):
        if cnf is None:
            cnf = {}
        cnf = tk._cnfmerge((cnf, kwargs))
        super(HoverButton, self).__init__(master=master, cnf=cnf, **kwargs)
        self.DBG = kwargs['background']
        self.ABG = kwargs['activeback']
        self.bind_class(self, "<Enter>", self.Enter)
        self.bind_class(self, "<Leave>", self.Leave)

    def Enter(self, event):
        self['bg'] = self.ABG

    def Leave(self, event):
        self['bg'] = self.DBG
