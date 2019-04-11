# salary_detail_sender

import msvcrt, sys, os, time
import mail_handler as mail
import Analysis_excel as xls

def getpass(prompt):
    count = 0
    chars = []
    for x in prompt:
        msvcrt.putch(bytes(x,'utf8'))
    while True:
        new_char = msvcrt.getch()
        if new_char in b'\r\n':
            break
        elif new_char == b'\0x3': #ctrl + c
            raise KeyboardInterrupt
        elif new_char == b'\b':
            if chars and count >= 0:
                count -= 1
                chars =  chars[:-1]
                msvcrt.putch(b'\b')
                msvcrt.putch(b'\x20')#space
                msvcrt.putch(b'\b')
        else:
            if count < 0:
                count = 0
            count += 1
            chars.append(new_char.decode('utf8'))
            msvcrt.putch(b'*')
    print()
    return ''.join(chars)

def are_u_sure(yes_or_no_answer):
	print(yes_or_no_answer)
	answer = input()
	if answer.lower() == 'yes':
		return True
	elif answer.lower() == 'no':
		return False
	else:
		print('请只输入 yes 或者 no !')
		are_u_sure()

def confirm_names(name_email_dic,name_salary_dic):
	time.sleep(1)
	os.system('cls')
	names = [ str(name).strip().split('.')[0] for name in name_salary_dic.keys() if str(name).strip().split('.')[0] in name_email_dic.keys() ]
	names_non = [ str(name).strip().split('.')[0] for name in name_salary_dic.keys() if not str(name).strip().split('.')[0] in name_email_dic.keys() ]

	# return names
	print('下面的名字是同时存在于邮箱表和工资表的，可以发送邮件：')
	print('===========================')
	count = 0
	for i in range(len(names)):
		count += 1
		print(f'{i+1}） {names[i]}',end='\t\t')
		if count == 3:
			count = 0
			print()
	if count in [1,2]:
		print()
	print('===========================')
	if not len(names_non) == 0:
		print(f'\n下面的名字只在工资表中找到，没有找到邮箱，无法发送邮件：')
		print('---------------------------')
		print(f'{",".join(names_non)}')
		print('---------------------------')
		print(f'你可以忽略并按下回车，只发送给有邮箱的人，或者ctrl+c结束程序然后检查一下')
		input()
	print('\n选择发送：输入名字前的序号, 用空格隔开 \n全部发送：输入‘0’ , 然后按回车')
	select = input()
	res = []
	selects = ','.join(select.split(' ')).split(',')
	if '0' in selects:
		print('输入中有“0”, 将给所有人发送信息。')
		if are_u_sure('确定? yes/no'):
			return names
		else:
			return confirm_names(name_salary_dic,name_email_dic)
	for x in selects:

		if not x.isdigit():
			print('输入错误，只能输入数字。')
			return confirm_names(name_salary_dic,name_email_dic)
		else:
			x = int(x) - 1
			if not x in range(len(names)):
				print(f'输入错误，上面没有数字{x+1}')
				return confirm_names(name_salary_dic,name_email_dic)
			else:
				res.append(names[x])
	print(f'将发送给这些人:{",".join(res)}')
	if are_u_sure(f'确认? yes/no'):
		return res
	else:
		return confirm_names(name_salary_dic,name_email_dic)

def send(names,mails,salaries):
	topic = input('请输入邮件主题，例如：“七月份工资明细”\n')
	failed = []
	print('=========================')
	for name in names:
		try:
			print(f'发送 {topic} 给 {name}',end='...\t')
			mail.send_mail(topic,salaries[name],mails[name])
			print(f'成功!')
		except Exception as e:
			print(e,end='失败\n')
			failed.append(name)
			continue
	print('=========================')
	return failed

def table_analysis():
	os.system('cls')

	print("""把邮箱表拖进这个窗口, 然后按回车:""")
	receiver_table = input()
	name_email = xls.read_emails(receiver_table)
	os.system('cls')
	print("""把工资表拖进这个窗口, 然后按回车:""")
	salaries_table = input()
	name_salary = xls.salaries_handler(salaries_table)
	return name_email,name_salary

def main():
	try:
		email_dic,salary_dic = table_analysis()
		targets = confirm_names(email_dic,salary_dic)
		os.system('cls')
		while True:
			try:
				time.sleep(1)
				os.system('cls')
				sender_addr = input("发件人邮箱地址（只支持@eltbio.com）:")
				if not sender_addr.endswith('@eltbio.com'):
					print('邮箱输入错误。')
					continue
				sender_pwd = getpass('Password: ')
				res = mail.init_smtp('smtp.mxhichina.com',sender_addr,sender_pwd)
				if res:
					print('登录成功')
					time.sleep(1)
					break
				else:
					time.sleep(1)
			except Exception as e:
				print(e)
				time.sleep(3)
				continue
		
		os.system('cls')
		failed_send = send(targets,email_dic,salary_dic)
		if failed_send:
			print('这些人发送失败：')
			print(f'{",".join(failed_send)}')
		else:
			print('全部完成。。。')
		mail.smtp_logout()
		time.sleep(2)
	except KeyboardInterrupt as e:
		os.system('cls')
		print('再见.')
	except Exception as e2:
		print(e2)
		input()
		exit()

if __name__ == '__main__':
	main()