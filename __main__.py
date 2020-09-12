from email.message import EmailMessage
from tkinter import *
from tkinter import filedialog, messagebox
import smtplib
import pandas as pd
import time
import ntpath


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
    def __init__(self):
        self.win = Tk()
        self.connect = False
        self.login = False
        self.first_root = False

        self.canvas0 = Canvas(self.win)

        self.button = Button
        self.buttonS = Button
        self.listpdf = ScrolledListbox
        self.listsend = ScrolledListbox
        self.entry0 = Entry
        self.label0 = Label
        self.label1 = Label
        self.text = ScrolledTextbox

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
            self.mail = smtplib.SMTP_SSL("smtp.gmail.com", 465)
        except Exception:
            pass

        self.First_Mainloop()
        self.win.columnconfigure(0, weight=1)
        self.win.rowconfigure(0, weight=1)
        self.win.mainloop()

    def First_Mainloop(self):
        self.canvas0.destroy()
        self.canvas0 = Canvas(self.win)
        self.canvas0.grid(row=0, column=0, sticky=NSEW)

        self.CSVFile = []
        self.entry0 = []
        self.label0 = []
        text_lbl = ['Gmail :', 'Password :']
        for ent in range(2):
            self.entry0.append(Entry(self.canvas0))
            self.entry0[ent].grid(row=ent, column=1, sticky=NSEW)
            self.label0.append(Label(self.canvas0, text=text_lbl[ent]))
            self.label0[ent].grid(row=ent, column=0, sticky=NSEW)

        self.button = Button(self.canvas0, text="Connect", command=lambda: self.Enter())
        self.button.grid(row=2, column=0, columnspan=2, sticky=NSEW)

        self.label1 = Label(self.canvas0, text="")
        self.label1.grid(row=3, column=0, columnspan=2, sticky=NSEW)
        self.first_root = True

    def Second_Mainloop(self):
        self.canvas0.destroy()
        self.canvas0 = Canvas(self.win)
        self.canvas0.grid(row=0, column=0, sticky=NSEW)
        self.canvas0.columnconfigure(0, weight=1)
        self.canvas0.columnconfigure(1, weight=1)
        self.canvas0.rowconfigure(1, weight=1)

        self.frame = []
        for ol in range(4):
            self.frame.append(Frame(self.canvas0))
            self.frame[ol].grid(row=ol, column=0, sticky=NSEW)


        self.frame[1].rowconfigure(1, weight=1)
        self.frame[1].columnconfigure(1, weight=1)
        self.frame[2].columnconfigure(2, weight=1)

        self.frame[3].grid(row=0, column=1, rowspan=3, sticky=NSEW)
        self.frame[3].rowconfigure(0, weight=1)
        self.frame[3].columnconfigure(1, weight=1)

        # CSV & PDF label
        self.label0 = []
        text_lbl = ['Select CSV file', 'Select PDF files']
        # buttons add files (.pdf) & (.csv)
        self.button = []
        func_but = [self.AddCSVFiles,  self.AddPDFiles]
        ent = 0
        for ku in range(0, 3, 2):
            self.label0.append(Label(self.frame[ku], text=text_lbl[ent]))
            self.label0[ent].grid(row=0, column=0, sticky=NSEW)
            self.button.append(Button(self.frame[ku], text="Click", command=func_but[ent]))
            self.button[ent].grid(row=0, column=1)
            ent += 1

        # label for imported CSV's file
        self.labelcsv = Label(self.frame[0])
        self.labelcsv.grid(row=0, column=2, sticky=NSEW)
        # listbox for imported PDF's files
        self.listpdf = ScrolledListbox(self.frame[2])
        self.listpdf.grid(row=0, column=2, sticky=NSEW)
        # TODO delete files (.pdf) button
        #######
        # object label & entry, message label & text
        self.label1 = []
        text_lbl = ['Object :', 'Message :']
        for ko in range(2):
            self.label1.append(Label(self.frame[1], text=text_lbl[ko]))
            self.label1[ko].grid(row=ko, column=0, sticky=NSEW)

        self.entry0 = []
        self.text = []
        for ok in range(1):
            self.entry0.append(Entry(self.frame[1], width=50))
            self.entry0[ok].grid(row=0, column=1, sticky=NSEW)

            self.text.append(ScrolledTextbox(self.frame[1], width=5, height=5))
            self.text[ok].grid(row=1, column=1, sticky=NSEW)

        # send : button & listbox
        self.buttonS = Button(self.frame[3], text='Send', command=self.RunSending)
        self.buttonS.grid(row=0, column=0, sticky=NSEW)

        self.listsend = ScrolledListbox(self.frame[3])
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

    def Connect(self):
        try:
            if self.first_root:
                print('Connecting ...')
                self.label1.configure(text='Connecting ...')
            else:
                print('Connecting ...')
                self.listsend.insert(END, 'Connecting ...')

            self.mail = smtplib.SMTP_SSL("smtp.gmail.com", 465)

            if self.first_root:
                print('Connected')
                self.label1.configure(text='Connected')
            else:
                print('Connected')
                self.listsend.insert(END, 'Connected')
                self.listpdf.see(END)
            self.connect = True
        except Exception:
            if self.first_root:
                self.label1.configure(text='Field to connect')
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
                self.label1.configure(text='Login ...')
            else:
                print('Login ...')
                self.listsend.insert(END, 'Login ...')

            self.mail.login(self.user_gmail, self.pass_gmail)

            if self.first_root:
                print('Login success')
                self.label1.configure(text='Login success')
            else:
                print('Login success')
                self.listsend.insert(END, 'Login success')
                self.listsend.see(END)
            self.login = True
        except Exception:
            if self.first_root:
                print('Field to login')
                self.label1.configure(text='Field to login')
                messagebox.showerror(title='Error !', message='Field to login, the gmail or password incorrect')
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
        time.sleep(2)
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

                self.mail.send_message(msg)
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
