from email.message import EmailMessage
import tkinter as tk
from tkinter import *
from tkinter import filedialog, messagebox
import smtplib
import pandas as pd
import time
import ntpath

"""
#### version 0.0.0.2 Beta

1. setting scheme's color & organizing widgets
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

        Listbox.__init__(self, self.canvas, *args, **kwargs)
        self.grid(row=0, column=0, sticky=NSEW)

        self.vbar = Scrollbar(self.canvas, orient=VERTICAL)
        self.hbar = Scrollbar(self.canvas, orient=HORIZONTAL)

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

        Text.__init__(self, self.canvas, *args, **kwargs)
        self.grid(row=0, column=0, sticky=NSEW)

        self.vbar = Scrollbar(self.canvas, orient=VERTICAL)
        self.hbar = Scrollbar(self.canvas, orient=HORIZONTAL)

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


class Gmail:
    __author__ = 'Achraf Najmi'
    __version__ = '0.0.0.2 Beta'
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

    def __init__(self):
        self.win = Tk()

        self.win.title(f'{self.__name__} v{self.__version__}')

        self.win.rowconfigure(0, weight=1)
        self.win.columnconfigure(0, weight=1)

        self.canvas = Canvas(self.win)

        self.connect = False
        self.login = False
        self.first_root = False

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

        self.button = HoverButton(self.canvas, **self.btn_prm, text="Connect", command=lambda: self.Enter())
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
        text_lbl = ['CSV file', 'PDF files']
        # buttons add files (.pdf) & (.csv)
        self.button = []
        func_but = [self.AddCSVFiles, self.AddPDFiles]
        ent = 0
        for ku in range(0, 3, 2):
            self.label0.append(Label(self.frame[ku], **self.lbl_prm, text=text_lbl[ent]))
            self.label0[ent].grid(row=0, column=0, sticky=NSEW)
            self.button.append(HoverButton(self.frame[ku], **self.btn_prm, text="Click", command=func_but[ent]))
            self.button[ent].grid(row=0, column=1)
            ent += 1

        # label for imported CSV's file
        self.labelcsv = Label(self.frame[0], **self.lbl_prm)
        self.labelcsv.grid(row=0, column=2)
        # listbox for imported PDF's files
        self.listpdf = ScrolledListbox(self.frame[2], **self.ent_prm, width=5, height=5)
        self.listpdf.grid(row=0, column=2, sticky=NSEW)
        # TODO delete files (.pdf) button
        #######
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
        self.buttonS = HoverButton(self.frame[3], **self.btn_prm, text='Send', command=self.RunSending)
        self.buttonS.grid(row=0, column=0, sticky=NSEW)

        self.listsend = ScrolledListbox(self.frame[3], **self.ent_prm, width=5, height=5)
        self.listsend.grid(row=0, column=1, sticky=NSEW)

        self.first_root = False

    def AddCSVFiles(self):
        self.CSVFile = FileDirectionCSV()
        try:
            self.data = pd.read_csv(self.CSVFile[0])
            self.labelcsv.configure(text=path_leaf(self.CSVFile[0]))
            self.total = self.data.Email.count() - 1
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

    def TextLabel(self, var):
        self.lblvar.set(var)

    def Connect(self):
        try:
            if self.first_root:
                print('Connecting ...')
                self.TextLabel('Connecting ...')
            else:
                print('Connecting ...')
                self.listsend.insert(END, 'Connecting ...')

            self.gmail = SMTPGmail()

            if self.first_root:
                print('Connected')
                self.TextLabel('Connected')
            else:
                print('Connected')
                self.listsend.insert(END, 'Connected')
                self.listpdf.see(END)

            self.connect = True
        except Exception:
            if self.first_root:
                self.TextLabel('Field to connect')
                print('Field to connect, check the internet connexion')
                messagebox.showwarning(title='Warning!', message='Field to connect, check the internet connexion')
            else:
                print('Field to connect, check the internet connexion')
                self.listsend.insert(END, 'Field to connect, check the internet connexion')
                self.listpdf.see(END)
            self.connect = False

    def Login(self):
        try:
            if self.first_root:
                print('Login ...')
                self.TextLabel('Login ...')
            else:
                print('Login ...')
                self.listsend.insert(END, 'Login ...')

            self.gmail.login(self.user_gmail, self.pass_gmail)

            if self.first_root:
                print('Login success')
                self.TextLabel('Login success')
            else:
                print('Login success')
                self.listsend.insert(END, 'Login success')
                self.listsend.see(END)
            self.login = True
        except Exception:
            if self.first_root:
                print('Field to login')
                self.TextLabel('Field to login')
                messagebox.showerror(title='Error !', message='Field to login, the gmail or password are incorrect')
                print('visit ')
                messagebox.showinfo(title='Notice !', message='visit')  # TODO : lien to less secure apps
            else:
                print('Field to login, Error !')
                self.listsend.insert(END, 'Field to login, Error !')
                self.listsend.see(END)
            self.login = False

    def Enter(self):
        self.user_gmail = str(self.entry0[0].get())
        self.pass_gmail = str(self.entry0[1].get())
        self.Connect()
        # time.sleep(2)
        if self.connect:
            self.Login()
        if self.login:
            self.Second_Mainloop()

    def Reconnect(self):
        try:
            self.Connect()
            # sleep to establish the connection
            time.sleep(2)
            self.Login()
            # continue sending
            self.Send()
        except Exception:
            print('Field to connect')
            print('1 minute to reconnect')
            self.listsend.insert(END, 'Field to connect',
                                 '1 minute to reconnect')
            self.listsend.see(END)
            time.sleep(60)
            self.Reconnect()

    def RunSending(self):
        self.listsend.delete(0, END)
        self.Subject = str(self.entry0[0].get())
        self.Message = self.text[0].get(1.0, END)
        self.entry0[0].configure(state='disabled')
        self.text[0].configure(state='disabled')
        self.listpdf.configure(state='disabled')
        for ml in range(2):
            self.button[ml].configure(state='disabled')
        self.buttonS.configure(text='Stop', command=self.Stopending)
        self.Send()

    def Stopending(self):
        self.Subject = str(self.entry0[0].get())
        self.Message = self.text[0].get(1.0, END)
        self.entry0[0].configure(state='normal')
        self.text[0].configure(state='normal')
        self.listpdf.configure(state='normal')
        for ml in range(2):
            self.button[ml].configure(state='normal')
        self.buttonS.configure(text='Send', command=self.RunSending)

    def Send(self):
        for self.staff in range(self.count, self.total):
            receiver_email = self.data.iloc[self.staff, 0]
            try:
                msg = EmailMessage()
                msg['Subject'] = self.Subject
                msg['From'] = self.user_gmail
                msg['To'] = receiver_email
                msg.set_content(self.Message)

                for file in self.PDFiles:
                    with open(file, 'rb') as f:
                        file_data = f.read()
                        file_name = f.name

                    msg.add_attachment(file_data,
                                       maintype='application',
                                       subtype='octect-stream',
                                       filename=file_name)

                self.gmail.send_message(msg)
                print(f"Message sent to : {self.staff} {receiver_email}")
                self.listsend.insert(END, f"Message sent to : {self.staff} {receiver_email}")
                self.listsend.see(END)
            except Exception:
                self.count = self.staff
                print('Disconnect')
                print('Field to send the message to', self.count, receiver_email)
                self.listsend.insert(END, 'Disconnect',
                                     'Field to send the message to')
                self.listsend.see(END)
                self.Reconnect()
        print('Sending Complete Successfully !!')
        self.listsend.insert(END, 'Sending Complete Successfully !!')
        self.listsend.see(END)
        return messagebox.showinfo(title='Congratulation !', message='Sending Complete Successfully !!'), \
               self.Second_Mainloop()


if __name__ == '__main__':
    Gmail()
