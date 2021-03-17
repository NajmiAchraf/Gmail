from email.message import EmailMessage
import tkinter as tk
from tkinter import *
from tkinter import filedialog, messagebox
import smtplib
import pandas as pd
import time
import ntpath
from concurrent import futures
import webbrowser

"""
#### version 0.0.0.4 RC

- add button to delete PDF file from listbox
"""


def SMTPGmail():
    return smtplib.SMTP_SSL("smtp.gmail.com", 465)


def FileDirectionPDF():
    return filedialog.askopenfilenames(
        initialdir='/',
        title='Select PDF files',
        filetypes=(("Portable Document Format (.pdf)", "*.pdf"), ("All Files (*.*)", "*.*"))
    )


def FileDirectionCSV():
    return filedialog.askopenfilenames(
        initialdir='/',
        title='Select CSV file',
        filetypes=(("CSV (.csv)", "*.csv"), ("All Files (*.*)", "*.*"))
    )


def path_leaf(path):
    head, tail = ntpath.split(path)
    return tail or ntpath.basename(head)


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


class ScrolledListbox(Listbox):
    def __init__(self, master, *args, **kwargs):
        self.canvas = Canvas(master)
        self.canvas.rowconfigure(0, weight=1)
        self.canvas.columnconfigure(0, weight=1)

        self.frame = Frame(self.canvas)
        self.frame.grid(row=0, column=0, sticky=NSEW)
        self.frame.rowconfigure(0, weight=1)
        self.frame.columnconfigure(0, weight=1)

        Listbox.__init__(self, self.frame, *args, **kwargs)
        self.grid(row=0, column=0, sticky=NSEW)

        self.vbar = AutoScrollbar(self.canvas, orient=VERTICAL)
        self.hbar = AutoScrollbar(self.canvas, orient=HORIZONTAL)

        self.configure(xscrollcommand=self.hbar.set, yscrollcommand=self.vbar.set)

        self.vbar.grid(row=0, column=1, sticky=NS)
        self.vbar.configure(command=self.yview)
        self.hbar.grid(row=1, column=0, sticky=EW)
        self.hbar.configure(command=self.xview)

        # Copy geometry methods of self.canvas without overriding Listbox
        # methods -- hack!
        listbox_meths = vars(Listbox).keys()
        methods = vars(Pack).keys() | vars(Grid).keys() | vars(Place).keys()
        methods = methods.difference(listbox_meths)

        for m in methods:
            if m[0] != '_' and m != 'config' and m != 'configure':
                setattr(self, m, getattr(self.canvas, m))

    def __str__(self):
        return str(self.canvas)


class ScrolledTextbox(Text):
    def __init__(self, master, *args, **kwargs):
        self.canvas = Canvas(master)
        self.canvas.rowconfigure(0, weight=1)
        self.canvas.columnconfigure(0, weight=1)

        self.frame = Frame(self.canvas)
        self.frame.grid(row=0, column=0, sticky=NSEW)
        self.frame.rowconfigure(0, weight=1)
        self.frame.columnconfigure(0, weight=1)

        Text.__init__(self, self.frame, *args, **kwargs)
        self.grid(row=0, column=0, sticky=NSEW)

        self.vbar = AutoScrollbar(self.canvas, orient=VERTICAL)
        self.hbar = AutoScrollbar(self.canvas, orient=HORIZONTAL)

        self.configure(xscrollcommand=self.hbar.set, yscrollcommand=self.vbar.set)

        self.vbar.grid(row=0, column=1, sticky=NS)
        self.vbar.configure(command=self.yview)
        self.hbar.grid(row=1, column=0, sticky=EW)
        self.hbar.configure(command=self.xview)

        # Copy geometry methods of self.canvas without overriding Text
        # methods -- hack!
        textbox_meths = vars(Listbox).keys()
        methods = vars(Pack).keys() | vars(Grid).keys() | vars(Place).keys()
        methods = methods.difference(textbox_meths)

        for m in methods:
            if m[0] != '_' and m != 'config' and m != 'configure':
                setattr(self, m, getattr(self.canvas, m))

    def __str__(self):
        return str(self.canvas)


class Gmail:
    __author__ = 'DeepEastWind'
    __version__ = '0.0.0.4 RC'
    __name__ = 'Gmail'

    btn_prm = {'padx': 18,
               'pady': 1,
               'bd': 1,
               'background': '#4d4d4d',
               'fg': 'white',
               'bg': '#4d4d4d',
               'font': ('DejaVu Sans', 18),
               'width': 4,
               'height': 1,
               'relief': 'raised',
               'activeback': '#3d3d3d',
               'activebackground': '#2d2d2d',
               'activeforeground': "white"}
    lbl_prm = {'fg': 'black',
               'bg': '#F0F0F0',
               'font': ('DejaVu Sans', 20),
               'relief': 'flat'}
    ent_prm = {'fg': 'black',
               'bg': 'white',
               'font': ('DejaVu Sans', 20)}

    TXT = {'cnting': 'Connecting ...',
           'cnted': 'Connected',
           'cntf': 'Field to connect', 'check': 'check the internet connexion',
           'login': 'Login ...',
           'loged': 'Login success',
           'logf': 'Field to login', 'incrt': 'the gmail or password are incorrect',
           'halfmin': '30 seconds to reconnect',
           'msgto': 'Message sent to :',
           'discnt': 'Disconnect',
           'msgtof': 'Field to send the message to',
           'success': 'Sending Complete Successfully !!',
           'stopped': 'The Sending is Stopped'
           }

    def __init__(self):
        self.win = Tk()

        self.thread_pool_executor = futures.ThreadPoolExecutor(max_workers=1)
        self.win.title(f'{self.__name__} v{self.__version__}')

        self.win.rowconfigure(0, weight=1)
        self.win.columnconfigure(0, weight=1)

        self.canvas = Canvas(self.win)

        self.connect = False
        self.login = False
        self.first_root = False
        self.Sending = False

        self.frame = Frame

        self.button = HoverButton
        self.buttonS = HoverButton
        self.listpdf = ScrolledListbox
        self.listsend = ScrolledListbox
        self.entry0 = Entry
        self.label0 = Label
        self.label1 = Label
        self.text = ScrolledTextbox
        self.lblvar = StringVar()

        self.data = pd.read_csv

        self.CSVFile = []
        self.PDFiles = []
        self.user_gmail = ''
        self.pass_gmail = ''
        self.Message = ''
        self.staff = int
        self.count = 0
        self.total = int

        try:
            self.gmail = SMTPGmail()
        except Exception:
            pass

        self.First_Mainloop()

        self.win.mainloop()

    def First_Mainloop(self):
        self.canvas.destroy()
        self.canvas = Canvas(self.win)

        self.canvas.grid(row=0, column=0, sticky=NSEW)

        self.canvas.rowconfigure(0, weight=4)
        self.canvas.rowconfigure(1, weight=4)
        self.canvas.rowconfigure(2, weight=4)
        self.canvas.rowconfigure(3, weight=4)
        self.canvas.columnconfigure(0, weight=1)
        self.canvas.columnconfigure(1, weight=4)

        self.lblvar = StringVar()

        self.CSVFile = []
        self.entry0 = []
        self.label0 = []
        text_lbl = ['Gmail :', 'Password :']
        for ent in range(2):
            self.entry0.append(Entry(self.canvas, **self.ent_prm))
            self.entry0[ent].grid(row=ent, column=1)
            self.label0.append(Label(self.canvas, **self.lbl_prm, text=text_lbl[ent]))
            self.label0[ent].grid(row=ent, column=0)

        self.button = HoverButton(self.canvas, **self.btn_prm, text="Connect",
                                  command=lambda: self.thread_pool_executor.submit(self.Enter))
        self.button.grid(row=2, column=0, columnspan=2)

        self.label1 = Label(self.canvas, **self.lbl_prm, textvariable=self.lblvar)
        self.label1.grid(row=3, column=0, columnspan=2)
        self.first_root = True

    def Second_Mainloop(self):
        self.canvas.destroy()
        self.canvas = Canvas(self.win)

        self.canvas.grid(row=0, column=0, sticky=NSEW)

        self.canvas.columnconfigure(0, weight=1)
        self.canvas.columnconfigure(1, weight=1)
        self.canvas.rowconfigure(1, weight=1)

        self.frame = []
        for ol in range(4):
            self.frame.append(Frame(self.canvas))
            self.frame[ol].grid(row=ol, column=0, sticky=NSEW)

        self.frame[1].rowconfigure(1, weight=1)
        self.frame[1].columnconfigure(1, weight=1)
        self.frame[2].columnconfigure(2, weight=1)

        self.frame[3].grid(row=0, column=1, rowspan=3, sticky=NSEW)
        self.frame[3].rowconfigure(0, weight=1)
        self.frame[3].columnconfigure(1, weight=1)

        # CSV & PDF label
        self.label0 = []
        text_lbl = ['CSV file', 'Add PDF files', 'Delete PDF file']
        # buttons add files (.pdf) & (.csv)
        self.button = []
        func_but = [self.AddCSVFiles, self.AddPDFiles, self.DeletePDFiles]
        ent = 0
        for ku in range(0, 3, 2):
            self.label0.append(Label(self.frame[ku], **self.lbl_prm, text=text_lbl[ent]))
            self.label0[ent].grid(row=0, column=0, sticky=NSEW)
            self.button.append(HoverButton(self.frame[ku], **self.btn_prm, text="Click", command=func_but[ent]))
            self.button[ent].grid(row=0, column=1)
            ent += 1

        # button delete file (.pdf)
        self.label0.append(Label(self.frame[2], **self.lbl_prm, text=text_lbl[2]))
        self.label0[2].grid(row=1, column=0, sticky=NSEW)
        self.button.append(HoverButton(self.frame[2], **self.btn_prm, text="Click", command=func_but[2]))
        self.button[2].grid(row=1, column=1)

        # label for imported CSV's file
        self.labelcsv = Label(self.frame[0], **self.lbl_prm)
        self.labelcsv.grid(row=0, column=2)
        # listbox for imported PDF's files
        self.listpdf = ScrolledListbox(self.frame[2], **self.ent_prm, width=5, height=5)
        self.listpdf.grid(row=0, column=2, rowspan=2, sticky=NSEW)

        # object label & entry, message label & text
        self.label1 = []
        text_lbl = ['Object :', 'Message :']
        for ko in range(2):
            self.label1.append(Label(self.frame[1], **self.lbl_prm, text=text_lbl[ko]))
            self.label1[ko].grid(row=ko, column=0, sticky=NSEW)

        self.entry0 = []
        self.text = []
        for ok in range(1):
            self.entry0.append(Entry(self.frame[1], **self.ent_prm))
            self.entry0[ok].grid(row=0, column=1, sticky=NSEW)

            self.text.append(ScrolledTextbox(self.frame[1], **self.ent_prm, width=5, height=5))
            self.text[ok].grid(row=1, column=1, sticky=NSEW)

        # send : button & listbox
        self.buttonS = HoverButton(self.frame[3], **self.btn_prm, text='Send', command=lambda: self.RunSending())
        self.buttonS.grid(row=0, column=0, sticky=NSEW)

        self.listsend = ScrolledListbox(self.frame[3], **self.ent_prm, width=5, height=5)
        self.listsend.grid(row=0, column=1, sticky=NSEW)

        self.first_root = False

    def AddCSVFiles(self):
        self.CSVFile = FileDirectionCSV()
        try:
            self.data = pd.read_csv(self.CSVFile[0])
            self.labelcsv.configure(text=path_leaf(self.CSVFile[0]))
            self.total = self.data.Email.count()
            print(self.total)
        except IndexError:
            return print('IndexError: string index out of range')
        print(self.CSVFile)

    def AddPDFiles(self):
        DirePDFiles = FileDirectionPDF()
        print(DirePDFiles)
        for dr in range(len(DirePDFiles)):
            if DirePDFiles[dr] not in self.PDFiles:
                self.PDFiles.append(DirePDFiles[dr])
                self.listpdf.insert(END, path_leaf(DirePDFiles[dr]))
        self.listpdf.see(END)
        print(self.PDFiles)

    def DeletePDFiles(self):
        try:
            # Delete from Listbox
            selection = self.listpdf.curselection()
            self.listpdf.delete(selection[0])
            # Delete from list that provided it
            self.PDFiles.pop(selection[0])
            print(self.PDFiles)
        except Exception:
            print("error")

    def TextLabel(self, var):
        self.lblvar.set(var)

    def ListSend(self, *elements):
        self.listsend.insert(END, *elements)
        self.listsend.see(END)

    def Connect(self):
        try:
            if self.first_root:
                print(self.TXT['cnting'])
                self.canvas.after_idle(self.TextLabel, self.TXT['cnting'])

            else:
                print(self.TXT['cnting'])
                self.canvas.after_idle(self.ListSend, self.TXT['cnting'])

            self.gmail = SMTPGmail()

            if self.first_root:
                print(self.TXT['cnted'])
                self.canvas.after_idle(self.TextLabel, self.TXT['cnted'])

            else:
                print(self.TXT['cnted'])
                self.canvas.after_idle(self.ListSend, self.TXT['cnted'])

            self.connect = True

        except Exception:
            if self.first_root:
                self.canvas.after_idle(self.TextLabel, self.TXT['cntf'])
                print(f"{self.TXT['cntf']}, {self.TXT['check']}")
                messagebox.showwarning(title='Warning!', message=f"{self.TXT['cntf']}, {self.TXT['check']}")

            else:
                print(f"{self.TXT['cntf']}, {self.TXT['check']}")
                self.canvas.after_idle(self.ListSend, f"{self.TXT['cntf']}, {self.TXT['check']}")

            self.connect = False

    def Login(self):
        try:
            if self.first_root:
                print(self.TXT['login'])
                self.canvas.after_idle(self.TextLabel, self.TXT['login'])

            else:
                print(self.TXT['login'])
                self.canvas.after_idle(self.listsend.insert, END, self.TXT['login'])

            self.gmail.login(self.user_gmail, self.pass_gmail)

            if self.first_root:
                print(self.TXT['loged'])
                self.canvas.after_idle(self.TextLabel, self.TXT['loged'])

            else:
                print(self.TXT['loged'])
                self.canvas.after_idle(self.ListSend, self.TXT['loged'])
            self.login = True

        except Exception:
            if self.first_root:
                print(self.TXT['logf'])
                self.canvas.after_idle(self.TextLabel, self.TXT['logf'])
                messagebox.showerror(title='Error !', message=f"{self.TXT['logf']}, {self.TXT['incrt']}")

                print('visit : https://myaccount.google.com/lesssecureapps')
                MsgBox = messagebox.askquestion('Open Browser', 'visit theose web sites to get access to your email :'
                                                                '\nGmail : https://myaccount.google.com/lesssecureapps')
                # '\nOutlook : https://outlook.live.com/mail/0/options/mail/accounts/popImap')
                if MsgBox == 'yes':
                    webbrowser.open_new("https://myaccount.google.com/lesssecureapps")
                else:
                    pass

            else:
                print(self.TXT['cntf'])
                print(self.TXT['halfmin'])
                self.canvas.after_idle(self.ListSend, self.TXT['cntf'], self.TXT['halfmin'])
                self.canvas.after_idle(self.ListSend, "0s")
                for tm in range(30):
                    if not self.Sending:
                        break
                    label = f"{tm}s"
                    idx = self.listsend.get(0, tk.END)
                    idx = idx.index(label)
                    self.canvas.after_idle(self.listsend.delete, idx)
                    print(f"{tm + 1}s", end=' ')
                    self.canvas.after_idle(self.ListSend, f"{tm + 1}s")
                    time.sleep(1)
            self.login = False

    def Enter(self):
        self.user_gmail = str(self.entry0[0].get())
        self.pass_gmail = str(self.entry0[1].get())
        self.Connect()
        time.sleep(2)

        if self.connect:
            self.Login()

        if self.login:
            self.Second_Mainloop()

    def Reconnect(self):
        self.Connect()
        # sleep to establish the connection
        time.sleep(2)
        self.Login()
        if not self.Sending:
            return self.ShowStopInfo()
        if not self.login:
            return self.Reconnect()
        # continue sending
        return self.Send()

    def RunSending(self):
        self.count = 0
        self.listsend.delete(0, END)
        self.Subject = str(self.entry0[0].get())
        self.Message = self.text[0].get(1.0, END)
        self.entry0[0].configure(state='disabled')
        self.text[0].configure(state='disabled')
        self.listpdf.configure(state='disabled')
        for ml in range(3):
            self.button[ml].configure(state='disabled')
        self.buttonS.configure(text='Stop', command=lambda: self.EndSending())
        self.Sending = True
        self.thread_pool_executor.submit(self.Send)

    def EndSending(self):
        self.entry0[0].configure(state='normal')
        self.text[0].configure(state='normal')
        self.listpdf.configure(state='normal')
        for ml in range(3):
            self.button[ml].configure(state='normal')
        self.buttonS.configure(text='Send', command=lambda: self.RunSending())
        self.Sending = False

    def ShowStopInfo(self):
        print(self.TXT['stopped'])
        self.canvas.after_idle(self.ListSend, self.TXT['stopped'])
        messagebox.showinfo(title='Stop !', message=self.TXT['stopped'])

    def Send(self):
        sending_error = False

        for self.staff in range(self.count, self.total):
            if not self.Sending:
                break

            receiver_email = self.data.iloc[self.staff, 0]

            try:
                msg = EmailMessage()
                msg['Subject'] = self.Subject
                msg['From'] = self.user_gmail
                msg['To'] = receiver_email
                msg.set_content(self.Message)

                for file in self.PDFiles:
                    file_name = path_leaf(file)
                    with open(file, 'rb') as f:
                        file_data = f.read()

                    msg.add_attachment(file_data,
                                       maintype='application',
                                       subtype='octect-stream',
                                       filename=file_name)

                self.gmail.send_message(msg)

                print(f"{self.staff + 1} {self.TXT['msgto']} {receiver_email}")
                self.canvas.after_idle(self.ListSend, f"{self.staff + 1} {self.TXT['msgto']} {receiver_email}")

            except Exception:
                self.count = self.staff

                print(self.TXT['discnt'])
                print(self.TXT['msgtof'], self.count, receiver_email)
                self.canvas.after_idle(self.ListSend, self.TXT['discnt'],
                                       f"{self.TXT['msgtof']} {self.count + 1} {receiver_email}")
                sending_error = True
                break
            if self.total == (self.staff + 1) or not self.Sending:
                break

        if sending_error and self.Sending:
            return self.Reconnect()
        elif not sending_error and not self.Sending:
            return  self.ShowStopInfo()
        else:
            print(self.TXT['success'])
            self.canvas.after_idle(self.ListSend, self.TXT['success'])
            messagebox.showinfo(title='Congratulation !', message=self.TXT['success'])
            return  self.EndSending()



if __name__ == '__main__':
    Gmail()
