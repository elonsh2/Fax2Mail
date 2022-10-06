import smtplib, ssl
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

from tqdm import tqdm
import os

# root folder: folder containing the PDFs to fax
root_folder = "C:\\Users\\Elon\\Desktop\\Exceltoword\\work\\combined\\"
# to_folder: folder that will contain all the PDFs that have been sent
to_folder = "C:\\Users\\Elon\\Desktop\\Exceltoword\\work\\sent\\"

class Mail:

	def __init__(self):
		self.port = 587
		self.smtp_server_domain_name = "smtp.gmail.com"
		# add Email address
		self.sender_mail = ""
		# add one time password
		self.password = ""

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
		# add the fax number/mail
		mails = ""
		subject = path
		mail = Mail()
		mail.send(mails, subject, path)
