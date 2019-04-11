# mail_handler

import smtplib
import time
from email.mime.text import MIMEText
from email.header import Header


def init_smtp(sender_host,sender_addr,sender_pwd):
	try:
		global smtp
		globals()['sender_addr'] = sender_addr
		smtp = smtplib.SMTP_SSL(sender_host)
		smtp.login(sender_addr,sender_pwd)
		return True,'登录成功'
	except smtplib.SMTPAuthenticationError as e:
		print('账号或者密码错误。')
		return False,'账号或者密码错误。'


def send_mail(topic,content,receiver):
	msg = MIMEText(content,'html','utf-8')
	msg['Subject'] = Header(topic,'utf-8')
	msg['From'] = sender_addr
	msg['To'] = receiver
	try:
		send = smtp.sendmail(sender_addr,receiver,msg.as_string())
		print(send)
		return True,str(send)
	except Exception as e:
		print(e)
		return False,repr(e)

def smtp_logout():
	try:
		smtp.quit()
	except Exception as e:
		print(e)
		exit()

def main():
	init_smtp('smtp.mxhichina.com','kai.ru@eltbio.com','Perhaps0')
	send_mail('Hello','This is a test Message from python.','kaidanalenko@outlook.com')
	smtp_logout()

if __name__ == '__main__':
	main()