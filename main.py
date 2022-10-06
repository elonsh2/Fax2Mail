import smtplib, ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import os
import shutil
from time import sleep
from tqdm import tqdm

from mailmerge import MailMerge
import openpyxl
from docx2pdf import convert
import os
from PyPDF2 import PdfFileMerger

root_folder = "C:\\Users\\Elon\\Desktop\\Exceltoword\\work\\combined\\"
to_folder = "C:\\Users\\Elon\\Desktop\\Exceltoword\\work\\sent\\"

class Mail:

	def __init__(self):
		self.port = 587
		self.smtp_server_domain_name = "smtp.gmail.com"
		self.sender_mail = "office.hadar.shamir@gmail.com"
		self.password = "apgqxpmrocszhnqm"

	def send(self, email, subject, path):
		message = MIMEMultipart('mixed')
		message['From'] = 'Contact <{sender}>'.format(sender=self.sender_mail)
		message['To'] = email
		message['CC'] = email
		message['Subject'] = subject

		attachmentPath = f"{root_folder}{path}.pdf"
		with open(attachmentPath, "rb") as attachment:
			p = MIMEApplication(attachment.read(), _subtype="pdf")
			p.add_header('Content-Disposition', "attachment", filename=path)
			message.attach(p)
		msg_full = message.as_string()
		context = ssl.create_default_context()
		with smtplib.SMTP(self.smtp_server_domain_name, self.port) as server:
			server.ehlo()
			server.starttls(context=context)
			server.ehlo()
			server.login(self.sender_mail, self.password)
			server.sendmail(self.sender_mail, email, msg_full)
			server.quit()


if __name__ == '__main__':
	for filename in tqdm(os.listdir(root_folder)):
		path = filename[0:-4]
		mails = "03-6926484@myfax.co.il"
		subject = path
		mail = Mail()
		mail.send(mails, subject, path)
