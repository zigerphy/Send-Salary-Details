# -*- coding: utf-8 -*-

import tkinter as tk
import random
from tkinter.filedialog import askopenfilename
from tkinter import messagebox
import os,time
import json
import Analysis_excel
import passwordUtils
import mail_handler
import base64

class Application(tk.Frame):
	def __init__(self, master=None):

		super().__init__(master)
		self.master = master
		self.configFile = f'{os.path.abspath(os.path.curdir)}/config.json'
		self.mail_file,self.sal_file,self.lastChosenDir,self.userConfig = self.readConfig()
		self.salt = 'IWASTOLDTOKILLYOU'
		self.topic = self.topicInit()
		# self.cipher = AES.new(self.salt,AES.MODE_ECB)
		self.master.protocol("WM_DELETE_WINDOW",self.on_closing)
		self.create_widgets()
		self.pack()
		
	def topicInit(self):
		day_now = time.localtime()
		topic_year = day_now.tm_year
		topic_mon = day_now.tm_mon - 1
		if day_now.tm_mon == 1:
			topic_year -= 1
			topic_mon = 12
		return f'{topic_year}年{topic_mon}月份工资'

	def emailLogin(self,email,passwd,autoLogin):
		email = email.strip()
		passwd = passwd.strip()
		print(email,passwd)
		result,message = mail_handler.init_smtp('smtp.mxhichina.com',email,passwd)
		if result:
			messagebox.showinfo(title="提示信息",detail = message)
			self.userConfig['autoLogin'] = autoLogin
			self.winNew.destroy()
			if autoLogin:
				identifyPatten = bytes(email+'$'+passwd+self.salt,encoding='utf-8')
				
				self.userConfig['identify'] = str(base64.encodebytes(identifyPatten),encoding='utf-8')
			else:
				self.userConfig['identify'] = ''
		else:
			messagebox.showerror(title="错误",detail=message)
		self.above = 0

	def loginWindow(self):
		self.above = 1
		size = '320x240+%d+%d' % ((self.winfo_screenwidth() - 320)/2, (self.winfo_screenheight() - 240)/2)
		self.winNew = tk.Toplevel(self)
		self.winNew.geometry(size)
		self.winNew.resizable(width=False,height=False)
		self.winNew.title('登录邮箱账号')
		self.winNew.grab_set()
		self.winNew.protocol("WM_DELETE_WINDOW",self.on_closing)
		title = tk.Label(self.winNew, text='登录邮箱')
		title.grid(row = 1, column=2, padx=8,pady=5)
		title.config(font=('黑体',20,'normal'))
		tk.Label(self.winNew, text='电子邮箱').grid(row = 2, column=1,sticky='E', padx=8,pady=5)
		emailText = tk.Text(self.winNew, wrap=tk.NONE, height = 1, width= 30)
		emailText.grid(row = 2,column = 2, padx=8,pady=5, sticky='E')

		password = tk.StringVar()
		autoLogin = tk.BooleanVar()
		tk.Label(self.winNew, text='密码').grid(row = 4, column=1,sticky='E', padx=8,pady=5)
		passwdText = tk.Entry(self.winNew,show='*',textvariable=password)
		passwdText.grid(row = 4,column = 2, padx=8,pady=5,sticky='W')
		tk.Checkbutton(self.winNew, text='记住密码', variable=autoLogin).grid(row=6,column=2,sticky='E')
		loginButton = tk.Button(self.winNew, text = "登录", height=1, width=10, command = lambda : self.emailLogin(emailText.get(1.0,tk.END).strip(),password.get(),autoLogin.get())).grid(row = 7,column=2,sticky='E')
		notNowButton = tk.Button(self.winNew, text = "暂不", height=1, width=10, command = self.winNew.destroy).grid(row = 8,column=2,sticky='E')
		
		if self.userConfig['autoLogin']:
			if self.userConfig['identify']:
				email,passwd = str(base64.decodebytes(bytes(self.userConfig['identify'],encoding='utf-8')),encoding='utf-8').split(self.salt)[0].split('$')
				self.emailLogin(email,passwd,self.userConfig['autoLogin'])
		return True

	def cancelAutoLogin(self):
		self.userConfig['autoLogin'] = False
		self.userConfig['identify'] = ''
		mail_handler.smtp_logout()
		self.loginWindow()
		self.winNew.lift(aboveThis=self.master)

	def on_closing(self):
		if messagebox.askokcancel("退出", "你确定要退出吗？"):
			self.writeConfig()
			try:
				mail_handler.smtp_logout()
			except Exception as e:
				print(e)
			self.master.destroy()

	def readConfig(self):
		try:
			if os.path.exists(self.configFile) and os.path.isfile(self.configFile):
				with open(self.configFile,'r') as config:
					configs = json.loads(config.read())
					return configs['mail_file'],configs['sal_file'],configs['lastChosenDir'],configs['userConfig']
			else:
				return '','','~',{'identify':'','autoLogin':'False'}
		except json.decoder.JSONDecodeError as jsonerror:
			print(jsonerror)
			os.unlink(self.configFile)
			return '','','~',{'identify':'','autoLogin':'False'}

	def writeConfig(self):
		config = {
				'mail_file' :self.mail_file,
				'sal_file' :self.sal_file,
				'lastChosenDir' :self.lastChosenDir,
				'userConfig':self.userConfig
				}
		configs = json.dumps(config,indent=4, sort_keys=True)
		with open(self.configFile,'w') as fileObj:
			print(configs)
			fileObj.write(configs)
	def textInsert(self, widget, texts):
		widget.config(state='normal')
		widget.delete(1.0,tk.END)
		widget.insert(tk.INSERT, texts)
		widget.config(state='disable')
		
	def create_widgets(self):
		self.topic_label = tk.Label(self, text='邮件主题：').grid(row = 1, column=1,sticky='W', padx=8,pady=5)
		self.topic_text = tk.Text(self, wrap='none', height = 1, width=45)
		self.topic_text.grid(row = 1,column = 2,columnspan = 3, padx=8,pady=5)
		self.topic_text.insert(tk.INSERT, self.topic)

		self.mail_label = tk.Label(self, text='邮箱表：').grid(row = 2, column=1,sticky='W', padx=8,pady=5)
		self.mail_text = tk.Text(self, wrap='none', height = 1, width=45)
		self.mail_text.grid(row = 2,column = 2,columnspan = 3, padx=8,pady=5)
		self.textInsert(self.mail_text, self.mail_file)

		self.mail_sel = tk.Button(self, text = "选择", width=8, command = lambda : self.fileSelect(self.mail_text,'选择邮箱列表'))
		self.mail_sel.grid(row = 2, column = 5, padx=8,pady=5)

		self.sal_label = tk.Label(self, text='工资表：').grid(row = 3, column=1,sticky='W', padx=8,pady=5)
		self.sal_text = tk.Text(self,wrap=tk.NONE, height = 1, width=45)
		self.sal_text.grid(row = 3,column = 2,columnspan = 3, padx=8, pady=5 )
		self.textInsert(self.sal_text, self.sal_file)

		self.sal_sel = tk.Button(self, text = "选择", width=8, command = lambda : self.fileSelect(self.sal_text,'选择工资列表')).grid(row = 3, column = 5, padx=8,pady=5)

		tk.Button(self, text="分析", width=8, command=self.annalysis).grid(row = 4, column=2)

		self.sendButton = tk.Button(self, text="发送", width=8, state=tk.DISABLED, command=self.sendEmails)
		self.sendButton.grid(row = 4, column=5)
		self.selectAllButton = tk.Button(self, text="全选", width=8, state=tk.DISABLED, command= lambda : self.selectAllBox(True))
		self.selectAllButton.grid(row = 4,column = 3)
		self.selectNoneButton = tk.Button(self, text="全不选", width=8, state=tk.DISABLED, command= lambda : self.selectAllBox(False))
		self.selectNoneButton.grid(row = 4,column = 4)

		# canvas=tk.Canvas(self,width=200,height=180,scrollregion=(0,0,520,520)) #创建canvas
		# canvas.place(x = 175, y = 265)
		# frame=tk.Frame(canvas) #把frame放在canvas里
		# frame.place(width=180, height=180) #frame的长宽，和canvas差不多的
		# vbar=tk.Scrollbar(canvas,orient=tk.VERTICAL) #竖直滚动条
		# vbar.place(x = 180,width=20,height=180)
		# vbar.configure(command=canvas.yview)
		# hbar=tk.Scrollbar(canvas,orient=tk.HORIZONTAL)#水平滚动条
		# hbar.place(x =0,y=165,width=180,height=20)
		# hbar.configure(command=canvas.xview)
		# canvas.config(xscrollcommand=hbar.set,yscrollcommand=vbar.set) #设置  
		# canvas.create_window((90,240), window=frame)  #create_window


	def selectAllBox(self,selectOrNot):
		if selectOrNot:
			for checkbutton,value in self.checkbuttons:
				value.set(1)
		else:
			for checkbutton,value in self.checkbuttons:
				value.set(0)

	def sendEmails(self):
		result = messagebox.askokcancel('提示','确定开始发送？')
		# print(result)
		failed = []
		if not result:
			return
		for checkbutton,value in self.checkbuttons:
			if value.get() == 1:
				name = checkbutton.cget('text')
				res,fallback = mail_handler.send_mail(self.topic,self.nameSalaries[name],self.nameEmails[name])
				if res:
					checkbutton.deselect()
				else:
					failed.append(name+fallback)
					checkbutton.configure(selectcolor='red')
				time.sleep(1)
		if len(failed) == 0:
			messagebox.showinfo('成功','全部完成！')
		else:
			messagebox.showerror('错误',f'有一些人没有发送成功请检查。\n可以尝试再次点击发送按钮。{failed}')

	def fileSelect(self, widget, title):
		user_dir = os.path.expanduser(self.lastChosenDir)
		file_name = askopenfilename(initialdir=user_dir,
			filetypes =(("Excel File", "*.xlsx"),),
			title = title)
		if file_name:
			self.lastChosenDir = os.path.dirname(file_name)
			self.textInsert(widget, file_name)

	def checkButtons(self):
		self.checkbuttons = []
		row = 6
		col = 2
		for name,email in self.nameSalaries.items():
			selectcolor='white'
			if name == '茹凯':
				selectcolor='gold'
			value = tk.IntVar()
			value.set(1)
			button = tk.Checkbutton(self,text=name,width=8,variable=value,selectcolor=selectcolor)
			button.grid(row=row,column = col,sticky=tk.W, padx=8, pady=5)
			if not name in self.aviliableNames:
				button.configure(state=tk.DISABLED)
			self.checkbuttons.append([button,value])
			if col < 4:
				col+=1
			else:
				col = 2
				row+=1
		return True

	def annalysis(self):
		self.mail_file = self.mail_text.get(1.0,tk.END).strip()
		self.sal_file = self.sal_text.get(1.0,tk.END).strip()
		if self.mail_file == '' or self.sal_file == '':
			messagebox.showerror('错误',f'请先选择工资表以及邮箱联系人表')
			return
		if not os.path.isfile(self.mail_file) or not os.path.isfile(self.sal_file):
			messagebox.showerror('错误',f'文件不存在:\n{self.mail_file}\n{self.sal_file}')
			return
		self.nameEmails = Analysis_excel.read_emails(self.mail_file)
		self.nameSalaries = Analysis_excel.salaries_handler(self.sal_file)

		self.aviliableNames = []
		for key,value in self.nameSalaries.items():
			if key in self.nameEmails.keys():
				self.aviliableNames.append(key)

		print(self.nameEmails)
		self.checkButtons()
		self.sendButton.configure(state='normal')
		self.selectAllButton.configure(state='normal')
		self.selectNoneButton.configure(state='normal')

# class Login(tk.Frame):
# 	"""docstring for Login"""
# 	def __init__(self, master=None):
# 		super().__init__(master)
# 		self.master = master
# 		self.result = True
# 		self.pack()

# 	def create_widgets():
# 		pass



def main():
	try:
		root = tk.Tk()
		root.title('发工资条')
		root.resizable(width=False,height=True)
		root.attributes("-topmost", True)
		screenwidth = root.winfo_screenwidth()
		screenheight = root.winfo_screenheight()
		size = '%dx%d+%d+%d' % (510, 600, (screenwidth - 510)/2, (screenheight - 600)/2)

		root.geometry(size)

		app = Application(master=root)
		mainmenu = tk.Menu(root)
		menuFile = tk.Menu(mainmenu,tearoff=0)
		mainmenu.add_cascade(label='菜单',menu=menuFile)
		menuFile.add_command(label='取消自动登录',command=app.cancelAutoLogin)
		root.config(menu=mainmenu)
		app.loginWindow()
		if app.above == 1:
			app.winNew.lift(aboveThis=app.master)
		app.mainloop()
		
	except Exception as e:
		messagebox.showerror('Error',e)
		raise e


if __name__ == '__main__':
	main()