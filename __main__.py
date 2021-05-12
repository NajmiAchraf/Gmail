from Lib import *

from email.message import EmailMessage
import smtplib
import pandas as pd

import tkinter as tk
from tkinter import *
from tkinter import filedialog, messagebox

import time

import ntpath
import webbrowser
import configparser
from os import path

import concurrent.futures

"""
#### version 0.1.0.1 beta

- fix adding files empty

- add auto fill up the account gmail and password after success sign up

- fix button delete : activate when file is selected
"""

__author__ = 'Najmi Achraf'
__version__ = '0.1.0.1 beta'
__title__ = 'Gmail'

if not path.exists('Account.ini'):
    Create_Settings_File()

parser = configparser.ConfigParser()
parser.read('Account.ini')


def SMTPGmail():
    return smtplib.SMTP_SSL("smtp.gmail.com", 465)


def CheckTheFormat(file, extension):
    try:
        form = file[0][-3:].lower()
        if form != extension:
            file = None
    except IndexError:
        file = None
    return file


def FileDirectionPDF():
    file = filedialog.askopenfilenames(
        initialdir='/',
        title='Select PDF files',
        filetypes=(("Portable Document Format (.pdf)", "*.pdf"), ("All Files (*.*)", "*.*"))
    )
    return CheckTheFormat(file=file, extension='pdf')


def FileDirectionCSV():
    file = filedialog.askopenfilenames(
        initialdir='/',
        title='Select CSV file',
        filetypes=(("Comma-Separated Values (.csv)", "*.csv"), ("All Files (*.*)", "*.*"))
    )
    return CheckTheFormat(file=file, extension='csv')


def path_leaf(path):
    head, tail = ntpath.split(path)
    return tail or ntpath.basename(head)


class Gmail:
    btn_prm = {'padx': 18,
               'pady': 1,
               'bd': 1,
               'background': '#4d4d4d',
               'fg': 'white',
               'bg': '#4d4d4d',
               'font': ('DejaVu Sans', 12),
               'width': 4,
               'height': 1,
               'relief': 'raised',
               'activeback': '#3d3d3d',
               'activebackground': '#2d2d2d',
               'activeforeground': "white"}
    lbl_prm = {'fg': 'black',
               'bg': '#F0F0F0',
               'font': ('DejaVu Sans', 12),
               'relief': 'flat'}
    ent_prm = {'fg': 'black',
               'bg': 'white',
               'font': ('DejaVu Sans', 12)}

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
           'stopped': 'The Sending is Stopped',
           'visit': 'visit theose web sites to get access to your email :',
           'gmail': 'Gmail : https://myaccount.google.com/lesssecureapps',
           }

    def __init__(self):
        self.win = Tk()

        self.thread_pool_executor = concurrent.futures.ThreadPoolExecutor(max_workers=1)
        self.win.title(f'{__title__} v{__version__}')

        self.win.rowconfigure(0, weight=1)
        self.win.columnconfigure(0, weight=1)

        self.win.bind_all('<Button-1>', self.SetDeleteButtonState)

        self.connect = False
        self.login = False
        self.first_root = False
        self.Sending = False

        self.frame = Frame

        self.buttons = HoverButton
        self.buttonS = HoverButton
        self.listpdf = ScrolledListbox
        self.listsend = ScrolledListbox
        self.entry1 = Entry
        self.label1 = Label
        self.label2 = Label
        self.lblvar = StringVar()

        self.data = pd.read_csv

        self.CSVFile = []
        self.PDFiles = []
        self.user_gmail = ''
        self.pass_gmail = ''
        self.Message = ''
        self.staff = int(0)
        self.count = int(0)
        self.total = int(0)

        # self.First_Mainloop()-----------------------------------------------------------------------------------------
        self.canvas1 = Canvas(self.win)

        self.canvas1.grid(row=0, column=0, sticky=NSEW)

        self.canvas1.rowconfigure(0, weight=4)
        self.canvas1.rowconfigure(1, weight=4)
        self.canvas1.rowconfigure(2, weight=4)
        self.canvas1.rowconfigure(3, weight=4)
        self.canvas1.columnconfigure(0, weight=1)
        self.canvas1.columnconfigure(1, weight=4)

        self.win.resizable(width=False, height=False)

        self.CSVFile = []
        self.entry1 = []
        self.label1 = []
        text_lbl = ['Gmail :', 'Password :']
        for ent in range(2):
            self.entry1.append(Entry(self.canvas1, **self.ent_prm, width=25))
            self.entry1[ent].grid(row=ent, column=1)
            self.label1.append(Label(self.canvas1, **self.lbl_prm, text=text_lbl[ent]))
            self.label1[ent].grid(row=ent, column=0)
        self.entry1[1].config(show="*")
        try:
            self.entry1[0].insert(0, parser.get('settings', 'gmail'))
            self.entry1[1].insert(0, parser.get('settings', 'password'))
        except configparser.NoOptionError:
            pass

        self.button = HoverButton(self.canvas1, **self.btn_prm, text="Connect",
                                  command=lambda: self.thread_pool_executor.submit(self.Enter))
        # command=lambda: self.Enter())
        self.button.grid(row=2, column=0, columnspan=2)

        self.label1 = Label(self.canvas1, **self.lbl_prm, textvariable=self.lblvar)
        self.label1.grid(row=3, column=0, columnspan=2)
        self.first_root = True

        # self.Second_Mainloop()----------------------------------------------------------------------------------------
        self.canvas2 = Canvas(self.win)

        self.canvas2.grid(row=0, column=0, sticky=NSEW)

        self.canvas2.columnconfigure(0, weight=1)
        self.canvas2.columnconfigure(1, weight=1)
        self.canvas2.rowconfigure(1, weight=1)

        self.canvas2.grid_forget()

        self.frame = []
        for ol in range(4):
            self.frame.append(Frame(self.canvas2))
            self.frame[ol].grid(row=ol, column=0, sticky=NSEW)

        self.frame[1].rowconfigure(1, weight=1)
        self.frame[1].columnconfigure(1, weight=1)

        self.frame[2].columnconfigure(2, weight=1)

        self.frame[3].grid(row=0, column=1, rowspan=3, sticky=NSEW)
        self.frame[3].rowconfigure(0, weight=1)
        self.frame[3].columnconfigure(1, weight=1)

        # CSV & PDF label
        self.label2 = []
        text_lbl = ['CSV file (Email Column)', 'Add PDF files', 'Delete PDF file']
        # buttons add files (.pdf) & (.csv)
        self.buttons = []
        func_but = [self.AddCSVFiles, self.AddPDFiles]
        ent = 0
        for ku in range(0, 3, 2):
            self.label2.append(Label(self.frame[ku], **self.lbl_prm, text=text_lbl[ent]))
            self.label2[ent].grid(row=0, column=0, sticky=NSEW)
            self.buttons.append(HoverButton(self.frame[ku], **self.btn_prm, text="Add", command=func_but[ent]))
            self.buttons[ent].grid(row=0, column=1)
            ent += 1

        # button delete file (.pdf)
        self.label2.append(Label(self.frame[2], **self.lbl_prm, text=text_lbl[2]))
        self.label2[2].grid(row=1, column=0, sticky=NSEW)
        self.buttons.append(HoverButton(self.frame[2], **self.btn_prm, text="Del", command=self.DeletePDFiles))
        self.buttons[2].grid(row=1, column=1)

        # label for imported CSV's file
        self.labelcsv = Label(self.frame[0], **self.lbl_prm)
        self.labelcsv.grid(row=0, column=2)
        # listbox for imported PDF's files
        self.listpdf = ScrolledListbox(self.frame[2], **self.ent_prm, width=5, height=5)
        self.listpdf.grid(row=0, column=2, rowspan=2, sticky=NSEW)

        # object label & entry, message label & text
        self.label2 = []
        text_lbl = ['Object :', 'Message :']
        for ko in range(2):
            self.label2.append(Label(self.frame[1], **self.lbl_prm, text=text_lbl[ko]))
            self.label2[ko].grid(row=ko, column=0, sticky=NSEW)

        self.entry2 = Entry(self.frame[1], **self.ent_prm)
        self.entry2.grid(row=0, column=1, sticky=NSEW)

        self.text = ScrolledTextbox(self.frame[1], **self.ent_prm, width=5, height=5)
        self.text.grid(row=1, column=1, sticky=NSEW)

        # send : button & listbox
        self.buttonS = HoverButton(self.frame[3], **self.btn_prm, text='Send', command=lambda: self.RunSending())
        self.buttonS.grid(row=0, column=0, sticky=NSEW)

        self.listsend = ScrolledListbox(self.frame[3], **self.ent_prm, width=5, height=5)
        self.listsend.grid(row=0, column=1, sticky=NSEW)

        try:
            self.gmail = SMTPGmail()
        except Exception:
            pass

        self.win.mainloop()

    def Second_Mainloop(self):
        self.canvas1.grid_forget()

        self.canvas2.grid(row=0, column=0, sticky=NSEW)
        self.canvas2.columnconfigure(0, weight=1)
        self.canvas2.columnconfigure(1, weight=1)
        self.canvas2.rowconfigure(1, weight=1)

        self.win.resizable(width=True, height=True)

        self.first_root = False

    def SetDeleteButtonState(self, event=None):
        selection = len(self.listpdf.curselection())
        if selection == 0:
            self.buttons[2]['state'] = DISABLED
        else:
            self.buttons[2]['state'] = NORMAL

    def AddCSVFiles(self):
        self.CSVFile = FileDirectionCSV()
        if self.CSVFile is None:
            pass
        else:
            try:
                self.data = pd.read_csv(self.CSVFile[0])
                self.total = self.data.Email.count()
                print(self.total)
                self.labelcsv.configure(text=f"{path_leaf(self.CSVFile[0])} \ Emails : {self.total}")
            except IndexError:
                return print('IndexError: string index out of range')
            print(self.CSVFile)

    def AddPDFiles(self):
        DirePDFiles = FileDirectionPDF()
        if DirePDFiles is None:
            pass
        else:
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
            self.SetDeleteButtonState()
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
                self.canvas1.after_idle(self.TextLabel, self.TXT['cnting'])

            else:
                print(self.TXT['cnting'])
                self.canvas2.after_idle(self.ListSend, self.TXT['cnting'])

            self.gmail = SMTPGmail()

            if self.first_root:
                print(self.TXT['cnted'])
                self.canvas1.after_idle(self.TextLabel, self.TXT['cnted'])

            else:
                print(self.TXT['cnted'])
                self.canvas2.after_idle(self.ListSend, self.TXT['cnted'])

            self.connect = True

        except Exception:
            if self.first_root:
                self.canvas1.after_idle(self.TextLabel, self.TXT['cntf'])
                print(f"{self.TXT['cntf']}, {self.TXT['check']}")
                messagebox.showwarning(title='Warning!', message=f"{self.TXT['cntf']}, {self.TXT['check']}")

            else:
                print(f"{self.TXT['cntf']}, {self.TXT['check']}")
                self.canvas2.after_idle(self.ListSend, f"{self.TXT['cntf']}, {self.TXT['check']}")

            self.connect = False

    def Login(self):
        try:
            if self.first_root:
                print(self.TXT['login'])
                self.canvas1.after_idle(self.TextLabel, self.TXT['login'])

            else:
                print(self.TXT['login'])
                self.canvas2.after_idle(self.listsend.insert, END, self.TXT['login'])

            self.gmail.login(self.user_gmail, self.pass_gmail)

            if self.first_root:
                print(self.TXT['loged'])
                self.canvas1.after_idle(self.TextLabel, self.TXT['loged'])

            else:
                print(self.TXT['loged'])
                self.canvas2.after_idle(self.ListSend, self.TXT['loged'])
            self.login = True

        except Exception:
            if self.first_root:
                print(self.TXT['logf'])
                self.canvas1.after_idle(self.TextLabel, self.TXT['logf'])
                messagebox.showerror(title='Error !', message=f"{self.TXT['logf']}, {self.TXT['incrt']}")

                print('visit : https://myaccount.google.com/lesssecureapps')
                MsgBox = messagebox.askquestion('Open Browser', f"{self.TXT['visit']}\n{self.TXT['gmail']}")
                # '\nOutlook : https://outlook.live.com/mail/0/options/mail/accounts/popImap')
                if MsgBox == 'yes':
                    webbrowser.open_new("https://myaccount.google.com/lesssecureapps")
                else:
                    pass

            else:
                print(self.TXT['cntf'])
                print(self.TXT['halfmin'])
                self.canvas2.after_idle(self.ListSend, self.TXT['cntf'], self.TXT['halfmin'])
                self.canvas2.after_idle(self.ListSend, "0s")
                for tm in range(30):
                    if not self.Sending:
                        break
                    label = f"{tm}s"
                    idx = self.listsend.get(0, tk.END)
                    idx = idx.index(label)
                    self.canvas2.after_idle(self.listsend.delete, idx)
                    print(f"{tm + 1}s", end=' ')
                    self.canvas2.after_idle(self.ListSend, f"{tm + 1}s")
                    time.sleep(1)
            self.login = False

    def Enter(self):
        self.user_gmail = str(self.entry1[0].get())
        self.pass_gmail = str(self.entry1[1].get())
        self.Connect()
        time.sleep(0.5)

        if self.connect:
            self.Login()

        if self.login:
            self.Second_Mainloop()
            Create_Settings_File(gmail=self.user_gmail, password=self.pass_gmail)

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
        self.Subject = str(self.entry2.get())
        self.Message = self.text.get(1.0, END)
        self.entry2.configure(state='disabled')
        self.text.configure(state='disabled')
        self.listpdf.configure(state='disabled')
        for ml in range(3):
            self.buttons[ml].configure(state='disabled')
        self.buttonS.configure(text='Stop', command=lambda: self.EndSending())
        self.Sending = True
        self.thread_pool_executor.submit(self.Send)

    def EndSending(self):
        self.entry2.configure(state='normal')
        self.text.configure(state='normal')
        self.listpdf.configure(state='normal')
        for ml in range(2):
            self.buttons[ml].configure(state='normal')
        self.buttonS.configure(text='Send', command=lambda: self.RunSending())
        self.Sending = False

    def ShowStopInfo(self):
        print(self.TXT['stopped'])
        self.canvas2.after_idle(self.ListSend, self.TXT['stopped'])
        messagebox.showinfo(title='Stop !', message=self.TXT['stopped'])

    def Send(self):
        sending_error = False

        for self.staff in range(self.count, self.total):
            if not self.Sending:
                break

            receiver_email = self.data.iloc[self.staff, 0]

            try:
                msg = self.CreateEmailMessage(receiver_email=receiver_email)
                self.gmail.send_message(msg)
                del msg

                print(f"{self.staff + 1} {self.TXT['msgto']} {receiver_email}")
                self.canvas2.after_idle(self.ListSend, f"{self.staff + 1} {self.TXT['msgto']} {receiver_email}")

            except Exception:
                self.count = self.staff

                print(self.TXT['discnt'])
                print(self.TXT['msgtof'], self.count + 1, receiver_email)
                self.canvas2.after_idle(self.ListSend, self.TXT['discnt'],
                                        f"{self.TXT['msgtof']} {self.count + 1} {receiver_email}")
                sending_error = True
                break
            if self.total == (self.staff + 1) or not self.Sending:
                break

        if sending_error and self.Sending:
            return self.Reconnect()
        elif not sending_error and not self.Sending:
            return self.ShowStopInfo()
        else:
            print(self.TXT['success'])
            self.canvas2.after_idle(self.ListSend, self.TXT['success'])
            messagebox.showinfo(title='Congratulation !', message=self.TXT['success'])
            return self.EndSending()

    def CreateEmailMessage(self, receiver_email):
        msg = EmailMessage()
        msg['From'] = self.user_gmail
        msg['To'] = receiver_email
        msg['Subject'] = self.Subject
        msg.set_content(self.Message)

        for file_path in self.PDFiles:
            file_name = path_leaf(file_path)
            with open(file_path, 'rb') as file_pdf:
                file_data = file_pdf.read()

            msg.add_attachment(file_data,
                               maintype='application',
                               subtype='octect-stream',
                               filename=file_name)
        return msg


if __name__ == '__main__':
    Gmail()
