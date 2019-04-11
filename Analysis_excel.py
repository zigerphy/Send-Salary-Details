# Analysis_excel
import xlrd
import table_utils

def salaries_handler(salary_table):
	name_salary = {}
	table = xlrd.open_workbook(salary_table).sheet_by_index(0)
	# print(dir(table))
	mcells = table.merged_cells

	mailBody = '<!DOCTYPE html> \
				<html><body>'
	mailBody = '<div style="overflow-x:auto;"><p>您好，感谢您一个月来对公司的付出，您的本月工资明细如下，请查收，如对工资有疑问，请反馈。</p> \
			<table style="border-collapse: collapse;">'
	for row_num in range(2):
		mailBody += '<tr style="background-color: grey">'
		for col_num in range(table.ncols):
			mailBody += table_utils.get_cellspan(table,row_num,col_num)
		mailBody += '</tr>'

	for row_num in range(2,table.nrows):
		single_line = mailBody
		single_line += '<tr>'
		name = table.cell_value(row_num,0)
		for col_num in range(table.ncols):
			cv = table.cell_value(row_num,col_num)
			# print(cv)
			if isinstance(cv,float):
				cv = '%.2f' % cv
			# print(cv)
			single_line += f'<td style="border: 1px solid black;white-space: nowrap;">{cv}</td>'

		single_line += '</tr></table></div></body></html>'
		name_salary[name] = single_line

	return name_salary

def read_emails(email_table):
	table = xlrd.open_workbook(email_table).sheet_by_index(0)
	name_email = {}
	for row_num in range(table.nrows):
		name = table.cell_value(row_num,0).strip()
		email = table.cell_value(row_num,1).strip()
		if not '@' in email:
			print(f"{name} 的邮箱似乎有问题")
			continue
		name_email[name] = email
	return name_email

def main():
	print(salaries_handler('E:\\Users\\kaidan\\Desktop\\工资(1).xlsx').keys())
	# print(read_emails('E:\\Users\\kaidan\\Desktop\\名单.xlsx').keys())

if __name__ == '__main__':
	main()
