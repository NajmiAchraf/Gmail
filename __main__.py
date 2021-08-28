from Lib import *

from email.message import EmailMessage
import smtplib
import pandas as pd

from tkinter import *
from tkinter import filedialog, messagebox

import time

import ntpath
import webbrowser
import configparser
import os

import concurrent.futures

"""
#### version 0.1.0.3 FV

- add new icons symbolize each button in the app

- add user Gmail login
"""

__author__ = 'Najmi Achraf'
__version__ = '0.1.0.3 FV'
__title__ = 'Gmail'

if not os.path.exists('Account.ini'):
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


def ModifyEmail(CSVFile):
	a_file = open(str(path_leaf(CSVFile[0])), "r")
	list_of_lines = a_file.readlines()
	list_of_lines[0] = "Email\n"
	a_file = open(str(path_leaf(CSVFile[0])), "w")
	a_file.writelines(list_of_lines)
	a_file.close()


class Gmail:
	btn_prm = {'padx': 18,
	           'pady': 1,
	           'bd': 1,
	           'fg': 'white',
	           'bg': 'firebrick2',
	           'font': ('DejaVu Sans', 12),
	           'width': 5,
	           'height': 1,
	           'relief': 'raised',
	           'activeback': 'firebrick3',
	           'activebackground': 'firebrick4',
	           'activeforeground': "white"}
	lbl_prm = {'fg': 'black',
	           'bg': '#F0F0F0',
	           'font': ('DejaVu Sans', 12),
	           'relief': 'flat'}
	ent_prm = {'fg': 'black',
	           'bg': 'white',
	           'font': ('DejaVu Sans', 12)}
	frm_prm = {'font': ('DejaVu Sans', 12)}

	TXT = {'cnting': 'Connecting ...',
	       'cnted': 'Connected',
	       'cntf': 'Field to connect', 'check': 'check the internet connexion',
	       'login': 'Log in ...',
	       'loged': 'Log in success',
	       'logf': 'Field to log in', 'incrt': 'the gmail or password are incorrect',
	       'halfmin': '30 seconds to reconnect',
	       'msgto': 'Message sent to :',
	       'discnt': 'Disconnect',
	       'msgtof': 'Field to send the message to',
	       'success': 'Sending Complete Successfully !!',
	       'stopped': 'The Sending is Stopped',
	       'error': 'Cannot begin sending until you add CSV and PDF Files',
	       'visit': 'visit theose web sites to get access to your email :',
	       'gmail': 'Gmail : https://myaccount.google.com/lesssecureapps',
	       }

	def __init__(self):
		self.win = Tk()

		self.thread_pool_executor = concurrent.futures.ThreadPoolExecutor(max_workers=1)
		self.win.title(f'{__title__} v{__version__}')
		self.win.protocol("WM_DELETE_WINDOW", self.ExitApplication)
		self.win.bind_all('<Button-1>', self.SetDeleteButtonState)

		self.win.resizable(width=False, height=False)
		self.win.rowconfigure(0, weight=1)
		self.win.columnconfigure(0, weight=1)

		self.data = pd.read_csv

		self.csv_file_existing = False
		self.connect = False
		self.login = False
		self.first_root = False
		self.Sending = False

		self.buttons = HoverButton
		self.buttonS = HoverButton
		self.listpdf = ScrolledListbox
		self.listsend = ScrolledListbox
		self.entry1 = Entry
		self.label0 = Label
		self.label1 = Label
		self.lblvar = StringVar()
		self.csvar = StringVar()
		self.actvar = StringVar()

		self.CSVFile = []
		self.PDFiles = []
		self.user_gmail = ''
		self.pass_gmail = ''
		self.Subject = ''
		self.Message = ''
		self.staff = int(0)
		self.count = int(0)
		self.total = int(0)

		# self.First_Mainloop()-----------------------------------------------------------------------------------------
		self.labelframe = LabelFrame(self.win, text="Login with Gmail", font=('DejaVu Sans', 24))
		self.labelframe.grid(row=0, column=0, padx=10, pady=10)
		self.labelframe.columnconfigure(1, weight=1)

		frames0 = []
		for ji in range(2):
			frames0.append(Frame(self.labelframe))
			frames0[ji].grid(row=ji, column=0, columnspan=2, padx=10, pady=10, sticky=NSEW)

		self.CSVFile = []
		self.entry1 = []
		self.label0 = []
		text_lbl = ['Gmail :', 'Password :']
		gt = 0
		for ent in range(2):
			self.label0.append(Label(frames0[ent], **self.lbl_prm, text=text_lbl[ent], anchor=W))
			self.label0[ent].grid(row=ent + gt, column=0, sticky=EW)
			self.entry1.append(Entry(frames0[ent], **self.ent_prm, width=25))
			self.entry1[ent].grid(row=ent + gt + 1, column=0)
			gt += 1
		self.entry1[1].config(show="*")

		try:
			self.entry1[0].insert(0, parser.get('settings', 'gmail'))
			self.entry1[1].insert(0, parser.get('settings', 'password'))
		except configparser.NoOptionError:
			pass

		log_img = PhotoImage(file='icons/login.png')
		Label(self.labelframe, image=log_img).grid(row=4, column=0)

		self.button = HoverButton(self.labelframe, **self.btn_prm, text="Connect",
		                          command=lambda: self.thread_pool_executor.submit(self.Enter))
		# command=lambda: self.Enter())
		self.button.grid(row=5, column=0, padx=10, pady=10)
		self.button.change_color_bind(DefaultBG='#4d4d4d', HoverBG='#3d3d3d', ActiveBG='#2d2d2d')

		self.label1 = Label(self.labelframe, **self.lbl_prm, textvariable=self.lblvar)
		self.label1.grid(row=4, column=1, rowspan=2, padx=10, pady=10, sticky=EW)

		self.first_root = True

		# self.Second_Mainloop()----------------------------------------------------------------------------------------
		self.canvas1 = Canvas(self.win)

		frame1 = Frame(self.canvas1)
		frame1.grid(row=0, column=0, padx=10, pady=10, sticky=NSEW)
		frame1.columnconfigure(1, weight=1)

		labels = ["To", 'Subject', 'Message', "PDF files"]
		labelframes = []
		for ol in range(4):
			labelframes.append(LabelFrame(self.canvas1, text=labels[ol], **self.frm_prm))
			labelframes[ol].grid(row=ol+1, column=0, padx=10, pady=10, sticky=NSEW)

		labelframes[1].columnconfigure(0, weight=1)

		labelframes[2].rowconfigure(0, weight=1)
		labelframes[2].columnconfigure(0, weight=1)

		labelframes[3].columnconfigure(1, weight=1)

		frame2 = Frame(self.canvas1)
		frame2.grid(row=5, column=0, padx=10, pady=10, sticky=NSEW)
		frame2.columnconfigure(0, weight=1)

		# logged label account
		box_img = PhotoImage(file='icons/inbox.png')
		Label(frame1, image=box_img).grid(row=0, column=0, rowspan=2)

		label_account = Label(frame1, **self.lbl_prm, textvariable=self.actvar, anchor=W)
		label_account.grid(row=0, column=1, rowspan=2, padx=10, pady=10, sticky=EW)

		# disconnect : button
		home_img = PhotoImage(file='icons/home.png')
		Label(frame1, image=home_img).grid(row=0, column=2)

		self.disconnect_button = HoverButton(frame1, **self.btn_prm, text='Disconnect',
		                                     command=lambda: self.First_Mainloop())
		self.disconnect_button.grid(row=1, column=2, padx=10, pady=10)
		self.disconnect_button.change_color_bind(DefaultBG='#4d4d4d', HoverBG='#3d3d3d', ActiveBG='#2d2d2d')

		# logout : button
		out_img = PhotoImage(file='icons/logout.png')
		Label(frame1, image=out_img).grid(row=0, column=3)

		self.logout_button = HoverButton(frame1, **self.btn_prm, text='Log out',
		                                 command=lambda: self.First_Mainloop(logout=True))
		self.logout_button.grid(row=1, column=3, padx=10, pady=10)

		self.buttons = []

		# button add file (.csv)
		self.buttons.append(HoverButton(labelframes[0], **self.btn_prm, text="Add (➕)", command=self.AddCSVFiles))
		self.buttons[0].grid(row=0, column=0, padx=10, pady=10)
		self.buttons[0].change_color_bind(DefaultBG='Royalblue2', HoverBG='Royalblue3', ActiveBG='Royalblue4')

		# label for imported CSV's file
		self.labelcsv = Label(labelframes[0], textvariable=self.csvar, **self.lbl_prm)
		self.labelcsv.grid(row=0, column=1, padx=10, pady=10)
		self.csvar.set("CSV file (Email Column)")

		# object : entry, message : text
		self.entry2 = Entry(labelframes[1], **self.ent_prm)
		self.entry2.grid(row=0, column=0, padx=10, pady=10, sticky=NSEW)

		self.text = ScrolledTextbox(labelframes[2], **self.ent_prm, width=5, height=5)
		self.text.grid(row=0, column=0, padx=10, pady=10, sticky=NSEW)

		# buttons add & delete files (.pdf)
		btn_lbl = ["Add (➕)", "Del (➖)"]
		func_but = [self.AddPDFiles, self.DeletePDFiles]
		for ent in range(0, 2):
			self.buttons.append(HoverButton(labelframes[3], **self.btn_prm, text=btn_lbl[ent], command=func_but[ent]))
			self.buttons[ent + 1].grid(row=ent, column=0, padx=10, pady=10)
		self.buttons[1].change_color_bind(DefaultBG='Royalblue2', HoverBG='Royalblue3', ActiveBG='Royalblue4')

		# listbox for imported PDF's files
		self.listpdf = ScrolledListbox(labelframes[3], **self.ent_prm, width=5, height=5)
		self.listpdf.grid(row=0, column=1, rowspan=2, padx=10, pady=10, sticky=NSEW)

		# send : button
		send_img = PhotoImage(file='icons/send.png')
		Label(frame2, image=send_img).grid(row=0, column=1)

		self.send_button = HoverButton(frame2, **self.btn_prm, text='Send', command=lambda: self.RunSending())
		self.send_button.grid(row=1, column=1, padx=10, pady=10)
		self.send_button.change_color_bind(DefaultBG='#20B645', HoverBG='#009C27', ActiveBG='#00751E')

		# self.Third_Mainloop()----------------------------------------------------------------------------------------
		self.canvas2 = Canvas(self.win)

		labelframe1 = LabelFrame(self.canvas2, text="Sending Treatment", **self.frm_prm)
		labelframe1.grid(row=0, column=0, columnspan=2, padx=10, pady=10, sticky=NSEW)
		labelframe1.rowconfigure(0, weight=1)
		labelframe1.columnconfigure(0, weight=1)

		frame3 = Frame(self.canvas2)
		frame3.grid(row=1, column=0, padx=10, pady=10, sticky=NSEW)
		frame3.columnconfigure(0, weight=1)

		# stop button & listbox : Sending Treatment
		self.listsend = ScrolledListbox(labelframe1, **self.ent_prm, width=5, height=5)
		self.listsend.grid(row=0, column=0, padx=10, pady=10, sticky=NSEW)

		# stop : button
		stop_img = PhotoImage(file='icons/stop.png')
		Label(frame3, image=stop_img).grid(row=0, column=1)

		self.stop_button = HoverButton(frame3, **self.btn_prm, text='Stop', command=lambda: self.EndSending())
		self.stop_button.grid(row=1, column=1, padx=10, pady=10)

		try:
			self.gmail = SMTPGmail()
		except Exception:
			pass

		self.win.mainloop()

	def ExitApplication(self):
		self.gmail.quit()
		print('quit done')
		sys.exit()

	def First_Mainloop(self, logout=None):
		if logout:
			Create_Settings_File()
		os.startfile(sys.argv[0])
		self.ExitApplication()

	def Second_Mainloop(self):
		self.labelframe.grid_forget()
		self.canvas2.grid_forget()

		self.canvas1.grid(row=0, column=0, sticky=NSEW)
		self.canvas1.columnconfigure(0, weight=1)
		self.canvas1.rowconfigure(3, weight=1)

		if self.first_root:
			self.win.geometry("800x700")
			self.win.resizable(width=True, height=True)

		self.first_root = False

	def Third_Mainloop(self):
		self.canvas1.grid_forget()

		self.canvas2.grid(row=0, column=0, sticky=NSEW)
		self.canvas2.rowconfigure(0, weight=1)
		self.canvas2.columnconfigure(0, weight=1)

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
				ModifyEmail(self.CSVFile)

				self.data = pd.read_csv(self.CSVFile[0])
				self.total = self.data.Email.count()
				print(self.total)

				self.csvar.set(f"{path_leaf(self.CSVFile[0])} \ Emails : {self.total}")
				self.csv_file_existing = True
			except IndexError:
				return print('IndexError: string index out of range')
			print(self.CSVFile)

	def AddPDFiles(self):   # TODO switch to adding all types of files
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
				self.labelframe.after_idle(self.TextLabel, self.TXT['cnting'])

			else:
				print(self.TXT['cnting'])
				self.canvas1.after_idle(self.ListSend, self.TXT['cnting'])

			self.gmail = SMTPGmail()

			if self.first_root:
				print(self.TXT['cnted'])
				self.labelframe.after_idle(self.TextLabel, self.TXT['cnted'])

			else:
				print(self.TXT['cnted'])
				self.canvas1.after_idle(self.ListSend, self.TXT['cnted'])

			self.connect = True

		except Exception:
			if self.first_root:
				self.labelframe.after_idle(self.TextLabel, self.TXT['cntf'])
				print(f"{self.TXT['cntf']}, {self.TXT['check']}")
				messagebox.showwarning(title='Warning!', message=f"{self.TXT['cntf']}, {self.TXT['check']}")

			else:
				print(f"{self.TXT['cntf']}, {self.TXT['check']}")
				self.canvas1.after_idle(self.ListSend, f"{self.TXT['cntf']}, {self.TXT['check']}")

			self.connect = False

	def Login(self):
		try:
			if self.first_root:
				print(self.TXT['login'])
				self.labelframe.after_idle(self.TextLabel, self.TXT['login'])

			else:
				print(self.TXT['login'])
				self.canvas1.after_idle(self.listsend.insert, END, self.TXT['login'])

			self.gmail.login(self.user_gmail, self.pass_gmail)

			if self.first_root:
				print(self.TXT['loged'])
				self.labelframe.after_idle(self.TextLabel, self.TXT['loged'])

			else:
				print(self.TXT['loged'])
				self.canvas1.after_idle(self.ListSend, self.TXT['loged'])
			self.login = True

		except Exception:
			if self.first_root:
				print(self.TXT['logf'])
				self.labelframe.after_idle(self.TextLabel, self.TXT['logf'])
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
				self.canvas1.after_idle(self.ListSend, self.TXT['cntf'], self.TXT['halfmin'])
				self.canvas1.after_idle(self.ListSend, "0s")
				for tm in range(30):
					if not self.Sending:
						break
					label = f"{tm}s"
					idx = self.listsend.get(0, END)
					idx = idx.index(label)
					self.canvas1.after_idle(self.listsend.delete, idx)
					print(f"{tm + 1}s", end=' ')
					self.canvas1.after_idle(self.ListSend, f"{tm + 1}s")
					time.sleep(1)
			self.login = False

	def Enter(self):
		self.button.configure(state='disabled')

		self.user_gmail = str(self.entry1[0].get())
		self.pass_gmail = str(self.entry1[1].get())
		self.Connect()
		time.sleep(0.5)

		if self.connect:
			self.Login()

		if self.login:
			self.Second_Mainloop()
			Create_Settings_File(gmail=self.user_gmail, password=self.pass_gmail)
			# reset label string variable
			self.labelframe.after_idle(self.TextLabel, '')
			self.actvar.set(f"Logged in as : {self.user_gmail}")

		self.button.configure(state='normal')

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
		if not self.csv_file_existing or len(self.PDFiles) == 0:
			return self.ShowErrorInfo()
		self.count = 0
		self.listsend.delete(0, END)
		self.Subject = str(self.entry2.get())
		self.Message = self.text.get(1.0, END)
		self.entry2.configure(state='disabled')
		self.text.configure(state='disabled')
		self.listpdf.configure(state='disabled')
		for ml in range(3):
			self.buttons[ml].configure(state='disabled')
		self.logout_button.configure(state='disabled')
		self.disconnect_button.configure(state='disabled')
		self.send_button.configure(state='disabled')
		self.stop_button.configure(state='normal')
		self.Sending = True
		self.Third_Mainloop()
		self.thread_pool_executor.submit(self.Send)

	def EndSending(self):
		self.entry2.configure(state='normal')
		self.text.configure(state='normal')
		self.listpdf.configure(state='normal')
		for ml in range(2):
			self.buttons[ml].configure(state='normal')
		self.logout_button.configure(state='normal')
		self.disconnect_button.configure(state='normal')
		self.send_button.configure(state='normal')
		self.stop_button.configure(state='disabled')
		self.Sending = False

	def ShowErrorInfo(self):
		print(self.TXT['error'])
		messagebox.showerror(title='Error !', message=self.TXT['error'])

	def ShowStopInfo(self):
		print(self.TXT['stopped'])
		self.canvas1.after_idle(self.ListSend, self.TXT['stopped'])
		messagebox.showinfo(title='Stop !', message=self.TXT['stopped'])
		self.Second_Mainloop()

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
				self.canvas1.after_idle(self.ListSend, f"{self.staff + 1} {self.TXT['msgto']} {receiver_email}")

			except Exception:
				self.count = self.staff

				print(self.TXT['discnt'])
				print(self.TXT['msgtof'], self.count + 1, receiver_email)
				self.canvas1.after_idle(self.ListSend, self.TXT['discnt'],
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
			self.canvas1.after_idle(self.ListSend, self.TXT['success'])
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
			                   subtype='octet-stream',
			                   filename=file_name)
		return msg


if __name__ == '__main__':
	Gmail()
