import time, datetime, os
import tkinter.filedialog
import tkinter as tk

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options
from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC

import pyautogui, schedule, openpyxl
COUNTRY_CODE = '91'


def get_whatsapp_logged_in_driver():
	chrome_options = Options()
	chrome_options.add_argument('--user-data-dir=User_Data')
	chrome_options.add_experimental_option("detach", True)
	chrome_path = './chromedriver.exe'
	driver = webdriver.Chrome(executable_path=chrome_path, options=chrome_options)
	wait = WebDriverWait(driver, 300)
	driver.get('https://web.whatsapp.com')
	driver.maximize_window()
	WebDriverWait(driver, 1000).until(
		EC.presence_of_element_located((By.XPATH, '//*[@id="side"]/header/div[2]/div/span/div[2]/div/span')))
	return driver


def open_contact(driver, contact_row, contacts_sheet = None):
	if contacts_sheet:
		phone_no = COUNTRY_CODE + str(contact_row[0].value).replace('-', '').replace(' ', '')[-10:]
	else:
		phone_no = COUNTRY_CODE+str(contact_row)
	print(phone_no)
	try:
		driver.get('https://web.whatsapp.com/send?phone=' + phone_no)
	except:
		alert = driver.switch_to.alert()
		alert.dismiss()
		open_contact(driver, contact_row)

	def wait_for_message_box():
		try:
			WebDriverWait(driver, 15).until(
				EC.presence_of_element_located((By.XPATH, '//*[@id="main"]/footer/div[1]/div[2]')))
			return True
		except:
			contact_not_found = driver.find_element_by_xpath(
				'//*[@id="app"]/div/span[2]/div/span/div/div/div/div/div/div[2]/div')
			if contact_not_found:
				contact_not_found.click()
				if contacts_sheet:
					contacts_sheet.cell(row=contact_row[0].row, column=contact_row[0].column+1).value = 'Contact Not Found'
					return False
			else:
				time.sleep(3)
				wait_for_message_box()

	return wait_for_message_box()


def send_message(driver, message, attachments_paths=None):
	try:
		WebDriverWait(driver, 1000).until(
			EC.presence_of_element_located((By.XPATH, '//*[@id="main"]/footer/div[1]/div[2]')))
		input_box = driver.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[2]/div/div[2]')

		if attachments_paths:
			driver.find_element_by_xpath('//*[@id="main"]/header/div[3]/div/div[2]/div').click()
			time.sleep(1)
			driver.find_element_by_xpath(
				'//*[@id="main"]/header/div[3]/div/div[2]/span/div/div/ul/li[1]/button').click()
			time.sleep(2)
			pyautogui.typewrite(attachments_paths)
			# time.sleep(2)
			pyautogui.press('enter')
			WebDriverWait(driver, 2000).until(EC.presence_of_element_located(
				(By.XPATH, '//*[@id="app"]/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/span/div')))
			time.sleep(1)
			attachment_send_button = driver.find_element_by_xpath(
				'//*[@id="app"]/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/span/div')
			attachment_send_button.click()
			time.sleep(1)
		for line in message.split('\n'):
			input_box.send_keys(line)
			ActionChains(driver).key_down(Keys.SHIFT).key_down(Keys.ENTER).key_up(Keys.ENTER).key_up(
				Keys.SHIFT).perform()
		input_box.send_keys(Keys.ENTER)


	# time.sleep(2)
	except Exception as e:
		print('Message sending failed because: ' + str(e))


multi_contacts_frame = None
single_contact_frame = None
# attachment_frame = None
variable_attachment_frame = None
attachment_paths = None
contacts_file_field = None
single_contact = None
count = None


def main():
	root = tk.Tk()
	root.iconbitmap('logo_black.ico')
	root.title('AutoWhatsapp')
	root.resizable(False, False)
	main_frame = tk.Frame(root)
	main_frame.pack(fill='both', padx=30, pady=30)
	top_frame = tk.Frame(main_frame)
	top_frame.grid(sticky='n', row=0)

	variable_grid_frame = tk.Frame(main_frame)
	variable_grid_frame.grid(sticky='n', row=1)

	def show_multiple_contacts_input_field():
		global multi_contacts_frame
		global single_contact_frame
		global contacts_file_field

		if single_contact_frame:
			single_contact_frame.destroy()
		multi_contacts_frame = tk.Frame(main_frame)
		multi_contacts_frame.grid(sticky='n', row=3)
		contacts_file_field = tk.Entry(multi_contacts_frame, width=30)

		def open_file():
			filename = tk.filedialog.askopenfilename(initialdir='/', title='Select Contact Excel File',
			                                         filetypes=[("Excel files", ".xlsx .xls")])
			contacts_file_field.insert(tk.END, filename)

		contacts_file_field.grid(column=0, row=1)
		contacts_button = tk.Button(multi_contacts_frame, text='Browse for excel file', command=open_file)
		contacts_button.grid(column=1, row=1)

	def show_single_contact_input_field():
		global single_contact_frame
		global multi_contacts_frame
		global single_contact
		global count
		if multi_contacts_frame:
			multi_contacts_frame.destroy()
		single_contact_frame = tk.Frame(main_frame)
		single_contact_frame.grid(sticky='n', row=3)
		single_contact_label = tk.Label(single_contact_frame, text='Enter the contact \nnumber (with code)')
		single_contact = tk.Entry(single_contact_frame)
		count = tk.Entry(single_contact_frame)
		count_label = tk.Label(single_contact_frame, text='Enter the number of times \nyou want to send the message')
		single_contact_label.grid(column=0, row=0)
		single_contact.grid(column=1, row=0)
		count_label.grid(column=2, row=0)
		count.grid(column=3, row=0)

	show_multiple_contacts_input_field()
	# TODO add radiobuttons in row 2
	operation_type_frame = tk.Frame(main_frame, pady=20)
	operation_type_frame.grid(row=2, columnspan=2)
	operation_choice = tk.IntVar()
	operation_choice.set(1)
	multi_contacts = tk.Radiobutton(operation_type_frame, command=show_multiple_contacts_input_field,
	                                text='Send message to Multiple contacts in excel sheet', indicatoron=False, value=1,
	                                variable=operation_choice)
	multi_contacts.grid(column=0, row=0, sticky='w')
	spam_contact = tk.Radiobutton(operation_type_frame, command=show_single_contact_input_field,
	                              text='Send message to Single contact multiple times', indicatoron=False, value=2,
	                              variable=operation_choice)
	spam_contact.grid(column=1, row=0, sticky='e')

	permanent_grid_frame = tk.Frame(main_frame)
	permanent_grid_frame.grid(sticky='n', row=4)

	message_frame = tk.Frame(permanent_grid_frame, pady=10)
	message_label = tk.Label(message_frame, text='Enter your Message')
	message_input = tk.Text(message_frame, height=8)
	message_frame.grid(row=0)
	message_label.grid(column=0, row=0)
	message_input.grid(column=0, row=1)

	# global attachment_frame
	# attachment_frame = tk.Frame(main_frame)
	def show_attachment_fields():
		# global attachment_frame
		global variable_attachment_frame
		global attachment_paths
		variable_attachment_frame = tk.Frame(attachment_frame, pady=10)
		variable_attachment_frame.grid(sticky='n', row=3, column=0, columnspan=2)

		def open_file():
			filename = tk.filedialog.askopenfilenames(initialdir='/', title='all attachments',
			                                          filetypes=[("All files", "*.*")])
			# if len(filename)>1:
			file_paths = ''
			for name in filename:
				file_paths += '; ' + str(name)
			file_paths = file_paths[2:]
			# else:
			# 	file_paths = filename[0]
			attachment_paths.insert(tk.END, file_paths)

		# attachment_frame.grid(sticky='n', row=5)
		attachment_paths = tk.Entry(variable_attachment_frame, width=30)
		attachment_button = tk.Button(variable_attachment_frame, text='Browse for attachments', command=open_file)
		attachment_paths.grid(row=0, column=0)
		attachment_button.grid(row=0, column=1)

	def hide_attachments_fields():
		global variable_attachment_frame
		if variable_attachment_frame:
			variable_attachment_frame.destroy()

	attachment_frame = tk.Frame(permanent_grid_frame, pady=10)
	attachment_label = tk.Label(attachment_frame, text='Do you want to add attachments?')
	attachment_choice = tk.IntVar()
	attachment_choice.set(1)
	attachment_yes = tk.Radiobutton(attachment_frame, text='Yes', value=2, variable=attachment_choice,
	                                command=show_attachment_fields)
	attachment_no = tk.Radiobutton(attachment_frame, text='No', value=1, variable=attachment_choice,
	                               command=hide_attachments_fields)
	attachment_frame.grid(row=1)
	attachment_label.grid(row=1, column=0, columnspan=2)
	attachment_yes.grid(row=2, column=0, sticky='en')
	attachment_no.grid(row=2, column=1, sticky='wn')

	def wait_till_message_sent(driver):
		message_wait = driver.find_elements_by_css_selector('span[data-icon=msg-time]')
		if len(message_wait):
			time.sleep(1)
			wait_till_message_sent(driver)
		else:
			print('Message sent')

	def validate_and_send():
		print('checking')
		message = message_input.get('1.0', 'end')

		def get_attachment_paths():
			global attachment_paths
			try:
				attachment_paths_string = attachment_paths.get()
				return ' '.join(attachment_paths_string.split(';')).replace("/", "\\")
			except:
				return None

		if operation_choice.get() == 1:
			global contacts_file_field
			global file_address
			file_address = contacts_file_field.get()
			wb = openpyxl.load_workbook(file_address)
			contacts_sheet = wb.active
			driver = get_whatsapp_logged_in_driver()

			for current_row in contacts_sheet.iter_rows():
				if contacts_sheet.cell(row=current_row[0].row, column=current_row[0].column+1).value != 'Message Sent':
					contact_found = open_contact(driver, current_row, contacts_sheet)
					if contact_found:
						send_message(driver, message, get_attachment_paths())
						wait_till_message_sent(driver)
						contacts_sheet.cell(row=current_row[0].row, column=current_row[0].column + 1).value = 'Message Sent'
			wb.save(filename='result.xlsx')
			driver.quit()
		else:
			global single_contact
			global count
			contact_no = str(single_contact.get())
			spam_count = int(count.get())
			driver = get_whatsapp_logged_in_driver()
			open_contact(driver, contact_no)
			for _ in range(spam_count):
				send_message(driver, message, get_attachment_paths())
			wait_till_message_sent(driver)
			driver.quit()

	submit_button = tk.Button(main_frame, text='Start Sending', command=validate_and_send)
	submit_button.grid()
	my_name = tk.Label(root, text='Reach out the developer on sid.kasat@gmail.com :)')
	my_name.pack(side='right', padx=10, pady=5)

	root.mainloop()

# driver.quit()

if __name__ == "__main__":
	main()
