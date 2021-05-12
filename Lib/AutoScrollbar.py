import tkinter as tk

__all__ = ["AutoScrollbar"]


class AutoScrollbar(tk.Scrollbar):
    """Scrollbar that automatically hides when not needed."""

    def __init__(self, master=None, **kwargs):
        """
        Create a scrollbar.

        :param master: master widget
        :type master: widget
        :param kwargs: options to be passed on to the :class:`ttk.Scrollbar` initializer
        """
        tk.Scrollbar.__init__(self, master=master, **kwargs)
        self._pack_kw = {}
        self._place_kw = {}
        self._layout = 'place'

    def set(self, lo, hi):
        """
        Set the fractional values of the slider position.

        :param lo: lower end of the scrollbar (between 0 and 1)
        :type lo: float
        :param hi: upper end of the scrollbar (between 0 and 1)
        :type hi: float
        """
        if float(lo) <= 0.0 and float(hi) >= 1.0:
            if self._layout == 'place':
                self.place_forget()
            elif self._layout == 'pack':
                self.pack_forget()
            else:
                self.grid_remove()
        else:
            if self._layout == 'place':
                self.place(**self._place_kw)
            elif self._layout == 'pack':
                self.pack(**self._pack_kw)
            else:
                self.grid()
        tk.Scrollbar.set(self, lo, hi)

    def _get_info(self, layout):
        """Alternative to pack_info and place_info in case of bug."""
        info = str(self.tk.call(layout, 'info', self._w)).split("-")
        dic = {}
        for i in info:
            if i:
                key, val = i.strip().split()
                dic[key] = val
        return dic

    def place(self, **kw):
        tk.Scrollbar.place(self, **kw)
        try:
            self._place_kw = self.place_info()
        except TypeError:
            # bug in some tkinter versions
            self._place_kw = self._get_info("place")
        self._layout = 'place'

    def pack(self, **kw):
        tk.Scrollbar.pack(self, **kw)
        try:
            self._pack_kw = self.pack_info()
        except TypeError:
            # bug in some tkinter versions
            self._pack_kw = self._get_info("pack")
        self._layout = 'pack'

    def grid(self, **kw):
        tk.Scrollbar.grid(self, **kw)
        self._layout = 'grid'
