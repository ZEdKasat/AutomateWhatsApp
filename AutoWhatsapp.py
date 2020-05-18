from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options
from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
import time, datetime, os, schedule, openpyxl
import os.path
import pyautogui


def get_message():
	message = []
	print('Enter the message you want to send: Add ~~ at the end to mark message as complete \n')
	while True:
		temp = str(input())
		if temp[-2:] == '~~':
			message.append(temp[:-2])
			break
		else:
			message.append(temp)
	return message


def get_whatsapp_logged_in_driver():
	chrome_options = Options()
	chrome_options.add_argument('--user-data-dir=User_Data')
	chrome_path = './chromedriver.exe'
	driver = webdriver.Chrome(executable_path=chrome_path, options=chrome_options)
	wait = WebDriverWait(driver, 300)
	driver.get('https://web.whatsapp.com')
	driver.maximize_window()
	WebDriverWait(driver, 1000).until(EC.presence_of_element_located((By.XPATH, '//*[@id="side"]/header/div[2]/div/span/div[2]/div/span')))
	return driver

def open_contact(driver, contact_row):
	phone_no = '91' + str(contact_row[0])[-10:]
	driver.get('https://web.whatsapp.com/send?phone=' + phone_no)
	WebDriverWait(driver, 1000).until(EC.presence_of_element_located((By.XPATH, '//*[@id="main"]/footer/div[1]/div[2]')))

def send_message(driver, message, attachments_paths=None):
	try:
		# phone_no = '91'+str(contact_row[0])[-10:]
		# driver.get('https://web.whatsapp.com/send?phone=' + phone_no)
		WebDriverWait(driver, 1000).until(EC.presence_of_element_located((By.XPATH, '//*[@id="main"]/footer/div[1]/div[2]')))
		input_box = driver.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[2]/div/div[2]')

		if attachments_paths:
			driver.find_element_by_xpath('//*[@id="main"]/header/div[3]/div/div[2]/div').click()
			time.sleep(1)
			driver.find_element_by_xpath('//*[@id="main"]/header/div[3]/div/div[2]/span/div/div/ul/li[1]/button').click()
			time.sleep(2)
			pyautogui.typewrite(attachments_paths)
			time.sleep(2)
			pyautogui.press('enter')
			WebDriverWait(driver, 2000).until(EC.presence_of_element_located((By.XPATH, '//*[@id="app"]/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/span/div')))
			time.sleep(1)
			attachment_send_button = driver.find_element_by_xpath('//*[@id="app"]/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/span/div')
			attachment_send_button.click()
			# time.sleep(1)
			# pyautogui.press('enter')
			# time.sleep(2)
			# input_box.send_keys(Keys.ENTER)

			time.sleep(1)
		for line in message:
			input_box.send_keys(line)
			ActionChains(driver).key_down(Keys.SHIFT).key_down(Keys.ENTER).key_up(Keys.ENTER).key_up(Keys.SHIFT).perform()
		input_box.send_keys(Keys.ENTER)

		def wait_till_message_sent():
			message_wait = driver.find_elements_by_class_name('jdhpF')
			message_wait = driver.find_elements_by_css_selector('span[data-icon=msg-time]')
			# message_wait = driver.find_elements_by_('jdhpF')
			if len(message_wait):
				time.sleep(1)
				wait_till_message_sent()
			else:
				print('Message sent')
		# wait_till_message_sent()
		# time.sleep(2)
	except Exception as e:
		print('Message sending failed because: '+str(e))


def get_attachments():
	no_of_attachments = int(input('Enter the Number of Attachments: '))
	paths = ''
	for number_of_file in range(no_of_attachments):
		file_path = input('Enter the address for file number '+ str(number_of_file))
		if os.path.isfile(file_path):
			paths = paths + ' "' + file_path + '"'
		else:
			print('File Not Found!\n')
	return paths

def main():
	message = get_message()
	print('')
	attachments_choice = True if str(input('\n\nDo you want to add attachments? |nType "yes" if you want to add attachments: ')).lower() == 'yes' else False
	if attachments_choice:
		attachments_paths = get_attachments()
	else:
		attachments_paths = None

	operation_choice = int(input('Press 1 to send bulk message to multiple contacts\nPress 2 to spam message to single contact\n'))
	if operation_choice == 1:
		file_address = str(input('Enter address to xls file that contains contacts: '))
	# file_address = 'E:\Development\WhatsappMessenger\resultcopy.xlsx'
		wb = openpyxl.load_workbook(file_address)
		contacts_sheet = wb.active
		driver = get_whatsapp_logged_in_driver()
		for row in contacts_sheet.iter_rows(values_only=True):
			open_contact(driver, row)
			send_message(driver, message, attachments_paths)
		wb.save(filename='result.xlsx')
		driver.quit()
	elif operation_choice == 2:
		contact_no = input('Enter Contact Number (10 digits only): ')
		spam_count = int(input('Enter the number of times you want to spam the contact: '))
		driver = get_whatsapp_logged_in_driver()
		open_contact(driver, [contact_no])
		for _ in range(spam_count):
			send_message(driver, message, attachments_paths)
		driver.quit()

if __name__ == "__main__":
	main()