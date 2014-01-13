#!/usr/bin/env python2.7
# -*- coding: utf-8 -*-
"""
Модуль для работы с почтой по протоколу IMAP.
"""
import sys, os, time, email, imaplib, smtplib
print sys.path
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate
from email import Encoders
from email.header import decode_header
import re
import datetime
import maillib
#import email.charset
#from win32com.server.exception import COMException
#import pywintypes

internal_date_re = re.compile(r'.*INTERNALDATE "'
		r'(?P<day>[ 0123][0-9])-(?P<mon>[A-Z][a-z][a-z])-(?P<year>[0-9][0-9][0-9][0-9])'
		r' (?P<hour>[0-9][0-9]):(?P<min>[0-9][0-9]):(?P<sec>[0-9][0-9])'
		r' (?P<zonen>[-+])(?P<zoneh>[0-9][0-9])(?P<zonem>[0-9][0-9])'
		r'"')

mon2num = {'Jan': 1, 'Feb': 2, 'Mar': 3, 'Apr': 4, 'May': 5, 'Jun': 6,
		'Jul': 7, 'Aug': 8, 'Sep': 9, 'Oct': 10, 'Nov': 11, 'Dec': 12}

def internalDate_to_datetime(idate):
	"""
	convert IMAP4 INTERNALDATE to python's type DateTime.
	"""
	mo = internal_date_re.match(idate)
	if not mo:
		return None
	mon = mon2num[mo.group('mon').decode()]
	zonen = mo.group('zonen')
	day = int(mo.group('day'))
	year = int(mo.group('year'))
	hour = int(mo.group('hour'))
	min = int(mo.group('min'))
	sec = int(mo.group('sec'))
	return datetime.datetime(year, mon, day, hour, min, sec) # Преобразуем лучше в DateTime

def str_to_internaldate(datestring="", formatstring='%d.%m.%Y'):
	"""
	convert string ('10.12.2011') and python/c format string ('%d.%m.%Y') into email InternalDate string
	"""
	return formatdate(timeval=time.mktime(time.strptime(datestring,formatstring)), localtime=True)

def decode_text(encoded_text):
	"""
	decode '=?=utf-8=?=' -like text, if it encoded.
	"""
	text, coding = decode_header(encoded_text)[0]
	#print 'HEADER CODING: %s' % (header_coding,)
	if coding:
		return text.decode(coding, errors='replace')
	else:
		return text.decode('ascii', errors='replace')

def decode_from_header(from_header):
	real_name = u''
	address = u''
	real_name, address = email.utils.parseaddr(from_header)
	encoded_text, coding = decode_header(real_name)[0]
	if coding:
		name = unicode(encoded_text, coding, errors='replace')
	else:
		name = unicode(encoded_text, errors='replace')
	return name, address

def decode_email_header(email_object, email_field):
	header_text = email_object[email_field]
	if email_field.upper() == "FROM":
		return decode_from_header(header_text)
	else:
		header_text, header_coding = decode_header(header_text)[0]
		if header_coding:
			return header_text.decode(header_coding, errors='replace')
		else:
			return header_text.decode('ascii', errors='replace')

class SmartMail(object):
	"""
	Класс для работы с почтой по протоколам SMTP и IMAP.
	"""
	_public_methods_ = [ 'connect', 'set_filter', 'set_fields', 'get_messages', 'messages_count', 'get_message',
						 'files_count', 'save_file', 'get_file_name', 'set_dir', 'set_folder',
						 'add_recipient','send','add_file']
	_public_attrs_ = ['date', 'sender', 'recipient', 'subject', 'body', 'type']
	_readonly_attrs_ = ['detach_dir']
	_reg_progid_ = "lkl.SmartMail"
	# NEVER copy the following ID
	# Use "print pythoncom.CreateGuid()" to make a new one.
	_reg_clsid_ = "{CC16108F-63D9-497F-8C26-D7861000192B}"

	def __init__(self):
		self.folder = "INBOX"

	def clear_current_data(self):
		pass #TODO: Вынести сюда очистку данных по текущему письму

	def connect(self, host, port, username, userpass, connection_type, use_ssl=0):
		"""
		Синтаксис:
			connect(host, port, username, userpass, connection_type, use_ssl=0): int, 1 or 0
		Параметры:
			host - str, Строка - Имя или IP почтового сервера;
			port - str or int, Строка или Число - Порт почтового сервера;
			username - str, Строка - Имя пользователя, e-mail адрес;
			userpass - str, Строка - Пароль пользователя;
			connection_type - str, Строка, ('imap' or 'pop3' or 'smtp') - Тип подключения к серверу;
			use_ssl - int, Число, (1 or 0) - Использовать SSL (1) или нет (0);
		Описание:
			Инициализация подключения, установка первоначальных параметров.
		Возвращаемое значение:
			Результат, число, 1 - удачно, 0 - неудачно.
		"""
		self.host = str(host).strip()
		self.port = int(port)
		self.username = str(username).strip()
		self.userpass = str(userpass).strip()
		self.connection_type = str(connection_type).strip().lower()
		if int(use_ssl):
			self.use_ssl = True
		else:
			self.use_ssl = False
		self.detach_dir = os.getcwd()
		self._sender 	  = None
		self._recipient   = None
		self._date 		  = None
		self.type         = 1 # integer, "text/plain" (0) or "text/html" (1)
		self.coding       = 'us-ascii'
		self._body        = None
		self._attachments = []
		self.current_mail = None
		## ----------------IMAP-----------------
		self.filter_string = ''
		self.fields = []
		self.items = []
		self.folder = "INBOX"
		if self.connection_type == 'imap':
			if self.use_ssl:
				self.connection = imaplib.IMAP4_SSL(self.host, self.port)
			else:
				self.connection = imaplib.IMAP4(self.host, self.port)
			self.connection.login(self.username,self.userpass)
			self.connection.select(self.folder)
		## ----------------SMTP-----------------
		elif self.connection_type == 'smtp':
			self.connection = smtplib.SMTP(self.host, self.port)
			if self.use_ssl:
				self.connection.ehlo()
				self.connection.starttls()
				self.connection.ehlo()
			self.connection.login(self.username, self.userpass)
			self.current_mail = MIMEMultipart()
		return 1

	def set_dir(self, path):
		"""
		Синтаксис:
			set_dir(path): int, 1 or 0
		Параметры:
			path - str, Строка - Каталог для сохранения вложений;
		Описание:
			Устанавливает путь к каталогу вложений.
		Возвращаемое значение:
			Результат, число, 1 - удачно, 0 - неудачно.
		"""
		if not path:
			return 0
		norm_path = os.path.normpath(path)
		if os.path.isdir(norm_path):
			self.detach_dir = norm_path
			return 1
		else:
			return 0

	def set_filter(self, filter_string=""):
		"""
		Синтаксис:
			set_filter(filter_string): int, 1 or 0
		Параметры:
			filter_string - str, Строка - Строка фильтра;
		Описание:
			Устанавливает строку фильтра IMAP перед выборкой,
			к примеру "(SINCE 01-Mar-2011 BEFORE 05-Mar-2011)"
		Возвращаемое значение:
			Результат, число, 1 - удачно, 0 - неудачно.
		"""
		if filter_string:
			self.filter_string = str(filter_string)
		return 1

	def set_folder(self, folder_name="Inbox"):
		"""
		Синтаксис:
			set_folder(self, folder_name="Inbox"): int, 1 or 0
		Параметры:
			folder_name - str, Строка - Название почтового каталога;
		Описание:
			Устанавливает текущий почтовый каталог,
			по-умолчанию - "Inbox"
		Возвращаемое значение:
			Результат, число, 1 - удачно, 0 - неудачно.
		"""
		if folder_name:
			self.folder = str(folder_name).strip()
			return 1
		return 0

	def set_fields(self, fields_string=""):
		""" (Пока не используется) """
		self.fields = str(fields_string).split(',')
		return 1

	def get_messages(self):
		"""
		Синтаксис:
			get_messages(): int, 1 or 0
		Параметры:
			нет.
		Описание:
			Инициализирует выборку писем.
		Возвращаемое значение:
			Результат, число, 1 - удачно, 0 - неудачно.
		"""
		response, items = self.connection.search(None, self.filter_string)
		print response, items
		self.items = items[0].split(' ')
		return response

	def messages_count(self):
		"""
		Синтаксис:
			messages_count(): int
		Параметры:
			нет.
		Описание:
			Возвращает кол-во писем в выборке.
		Возвращаемое значение:
			Количество писем, число.
		"""
		return len(self.items)

	def get_message(self, message_number=1):
		"""
		Синтаксис:
			get_message(message_number): int, 1 or 0
		Параметры:
			message_number - int, Число - Порядковый номер письма.
		Описание:
			Получает письмо из выборки по порядковому номеру.
			Нумерация идет с 1.
		Возвращаемое значение:
			Результат, число, 1 - удачно, 0 - неудачно.
		"""
		emailid = self.items[int(message_number)-1]
		try:
			resp, (data, internaldate) = self.connection.fetch(emailid, "(RFC822 INTERNALDATE)")
		except Exception as e:
			print e
			return -1
		self._date = internalDate_to_datetime(internaldate) # сразу получаем дату сообщения
		#return "OK!"
		email_body = data[1]
		mail 	= email.message_from_string(email_body)
		mail2 	= maillib.Message.from_message(mail)
		# if mail.get_content_maintype() != 'multipart':
			# return 0
		# else:
		self.current_mail 	= mail
		self.current_mail2 	= mail2
		return 1

	def files_count(self):
		"""
		Синтаксис:
			files_count(): int
		Параметры:
			нет.
		Описание:
			Возвращает кол-во вложений в текущем письме.
		Возвращаемое значение:
			Количество вложеных файлов, число.
		"""
		if not self.current_mail:
			return 0
		files_count = 0
		for part in self.current_mail.walk():
			#### multipart/* are just containers
			if part.get_content_maintype() == 'multipart':
				continue
			if part.get('Content-Disposition') is None:
				continue
			files_count += 1
		return files_count

	def save_file(self, file_number=1):
		"""
		Синтаксис:
			save_file(file_number): str
		Параметры:
			file_number - int, Число - Порядковый номер вложения.
		Описание:
			Сохраняет файл-вложение с порядковым номером <file_number>
			в каталог вложений. Возвращает имя файла. Нумерация идет с 1.
		Возвращаемое значение:
			Имя файла, строка.
		"""
		if not file_number:
			return 0
		if not self.detach_dir:
			self.detach_dir = os.getcwd()
		num = 0
		for part in self.current_mail.walk():
			#### multipart/* are just containers
			if part.get_content_maintype() == 'multipart':
				continue
			if part.get('Content-Disposition') is None:
				continue
			num += 1
			if num <> int(file_number):
				continue # получаем только нужный файл
			filename = decode_text(part.get_filename())
			counter = 1
			if not filename:
				filename = 'part-%03d%s' % (counter, 'bin')
				counter += 1
			att_path = os.path.abspath(os.path.join(self.detach_dir, filename))
			try:
				fp = open(att_path, 'wb+') # перезаписываем, если такой файл уже есть
				fp.write(part.get_payload(decode=True))
				fp.close()
				return att_path
			except:
				return 0

	def get_file_name(self, file_number=1):
		"""
		Синтаксис:
			get_file_name(file_number): str
		Параметры:
			file_number - int, Число - Порядковый номер вложения.
		Описание:
			Возвращает имя файла-вложения с порядковым номером <file_number>.
			Нумерация идет с 1.
		Возвращаемое значение:
			Имя файла, строка.
		"""
		if not file_number:
			return 0
		if not self.detach_dir:
			self.detach_dir = os.getcwd()
		num = 0
		for part in self.current_mail.walk():
			#### multipart/* are just containers
			if part.get_content_maintype() == 'multipart':
				continue
			if part.get('Content-Disposition') is None:
				continue
			num += 1
			if num <> int(file_number):
				continue # получаем только нужный файл
			filename = decode_text(part.get_filename())
			counter = 1
			if not filename:
				filename = 'part-%03d%s' % (counter, 'bin')
				counter += 1
			return filename

	def get_body(self):
		"""
		Синтаксис:
			get_body(): str
		Параметры:
			нет.
		Описание:
			Возвращает тело текущего письма, приоритет -
			если есть HTML возвращается он, в противном случае
			возвращается PLAIN.
		Возвращаемое значение:
			Тело письма, строка.
		"""
		if not self.current_mail:
			return 0
		# maintype = self.current_mail.get_content_maintype()
		# if maintype == 'multipart':
			# for part in self.current_mail.get_payload():
				# if part.get_content_maintype() == 'text':
					# return part.get_payload(decode=True)
		# elif maintype == 'text':
			# return self.current_mail.get_payload(decode=True)
		text_html = self.current_mail2.html
		if not text_html:
			text_html = self.current_mail2.body
		self._body = text_html
		return self._body

	def add_recipient(self, recipient):
		"""
		Синтаксис:
			add_recipient(recipient): None
		Параметры:
			recipient - str, Строка, список получателей письма 
		Описание:
			Возвращает тело текущего письма, приоритет -
			если есть HTML возвращается он, в противном случае
			возвращается PLAIN.
		Возвращаемое значение:
			Тело письма, строка.
		"""
		if (self.connection_type == 'smtp') and (self.current_mail):
			#self._recipient.append(recipient.strip())
			self._recipient.extend(recipient.strip())


	def send(self):
		"""
		Отправляет письмо по протоколу SMTP
		"""
		if self.connection_type != 'smtp':
			return -1
		if not len(self._recipient):
			return -2
		if not self._sender:
			self._sender = self.username
		if not self._date:
			self._date = formatdate(localtime=True)
		self.current_mail['From']    = self._sender
		self.current_mail['To']      = COMMASPACE.join(self._recipient)
		self.current_mail['Date']    = self._date
		self.current_mail['Subject'] = self._subject
		body_type = 'plain'
		if self.type == 1:
			body_type = 'html'
		self.current_mail.attach(MIMEText(self._body, body_type, self.coding))
		for f in self._attachments:
			part = MIMEBase('application', "octet-stream")
			part.set_payload(open(f,"rb").read())
			email.Encoders.encode_base64(part)
			part.add_header('Content-Disposition', 'attachment; filename="%s"'
											   % os.path.basename(f))
			self.current_mail.attach(part)
		self.connection.sendmail(self._sender, self._recipient, self.current_mail.as_string())
		#self.connection.rset() #???
		self.connection.close()

	def add_file(self, filename):
		if (self.connection_type != 'smtp') and (not self.current_mail):
			return -1
		f = os.path.abspath(filename)
		if not (os.path.isfile(f)):
			return -2
		self._attachments.append(f)

#--------date attribute---------
	@property
	def date(self):
		return self._date

	@date.getter
	def date(self):
		if not self.current_mail:
			return 0
		else:
			return self._date

	@date.setter
	def date(self, value):
		if self.connection_type == 'smtp':
			self._date = str_to_internaldate(value.strip()) # date in format dd.mm.yyyy (%d.%m.%Y)

#--------sender attribute------
	@property
	def sender(self):
		return self._sender

	@sender.getter
	def sender(self):
		if self.connection_type == 'imap':
			if not self.current_mail:
				return ""
			else:
				return decode_email_header(self.current_mail, "From")[1]

	@sender.setter
	def sender(self, value):
		if self.connection_type == 'smtp':
			self._sender = value.strip()

#--------recipient attribute----
	@property
	def recipient(self):
		return self._recipient

	@recipient.getter
	def recipient(self):
		if not self.current_mail:
			return 0
		else:
			tos = self.current_mail.get_all('to', [])
			tos = email.utils.getaddresses(tos)
			self._recipient = [i[1] for i in tos]
			return ';'.join(self._recipient)

	@recipient.setter
	def recipient(self, value):
		self._recipient = value.strip()

#--------subject attribute------
	@property
	def subject(self):
		# return self._subject
		pass

	@subject.getter
	def subject(self):
		if not self.current_mail:
			return 0
		else:
			return decode_email_header(self.current_mail, "Subject")

	@subject.setter
	def subject(self, value):
		if self.connection_type == 'smtp':
			self._subject = value.strip()

#--------body attribute---------
	@property
	def body(self):
		pass

	@body.getter
	def body(self):
		if self.connection_type == 'imap':
			return self.get_body()
		elif self.connection_type == 'smtp':
			return self._body
		else:
			return None

	@body.setter
	def body(self, value):
		if self.connection_type == 'smtp':
			self._body = value


#---------------TEST------------

def test_send():
	obj = SmartMail()
	obj.connect("smtp.gmail.com", "587", "botkakbot@gmail.com", "*********", "smtp", 1)
	obj.sender = 'botkakbot@gmail.com'
	obj.recipient = 'botkakbot@yandex.ru'
	obj.subject = 'smartmail smtp test #4'
	obj.body = '<b>test</b> привет! <i>test</i>'
	obj.send()


# Add code so that when this script is run by
# Python.exe, it self-registers.
if __name__=='__main__':
	print "Registering COM server..."
	import win32com.server.register
	win32com.server.register.UseCommandLine(SmartMail)

