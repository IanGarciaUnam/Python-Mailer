import smtplib, ssl,os,sys
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.encoders import encode_base64
from tkinter import messagebox
port= 465
smtp_server = "smtp.gmail.com"

class Sender:
	"""
	This class is the model to send emails massively and automatically
	"""
	def __init__(self, sender_email:str, password:str, receivers_emails):
		"""
		Builder for Sender
		Params: sender_email, just the email, configured for smtp.gmail
				receivers_emails: list of emails to receive the message
		"""
		self.sender_email=sender_email
		self.receivers_emails=receivers_emails
		self.context = ssl.create_default_context()
		self.password=password

	def send_message(self, message:str):
		"""
		Send message function is designed to send the same message for every user

		"""
		with smtplib.SMTP_SSL(smtp_server, port, context=self.context) as server:
			counter=0
			try:
				server.login(self.sender_email, self.password)
			
			except :
				messagebox.showerror(title="Credenciales incorrectas", message="Tu correo o contraseña son incorrectos, a continuación se terminará el programa")
				sys.exit(1)

			for receiver in self.receivers_emails:
				server.sendmail(self.sender_email, receiver, message)
				counter+=1
			print("Mensaje enviado con éxito a "+ str(counter)+" contactos")
"""
lista =[]
for i in range(10):
	lista.append("gonzalezlin24@gmail.com")

sender= Sender("canitogarciavazquez@gmail.com", "Maquinadeguerra2", lista)
message=str(i)
sender.send_message(message)
"""


class Sender_Improved:
	"""
	Class to build a sender that could sent images and the message 
	is in html
	"""
	def __init__(self,email:str, password:str, contactos):
		self.email=email
		self.password=password
		self.contactos=contactos
		self.server=smtplib.SMTP("smtp.gmail.com", 587)
		self.server.starttls()#Protocolo de cifrado de datos
		try:
			self.server.login(self.email, self.password)
		except:
			messagebox.showerror(title="Credenciales incorrectas", message="Tu correo o contraseña son incorrectos")
			self.finalize()

	def send_message(self, asunto, message, contacto, file=None,):
		"""
		Params:
			asunto : Subject of the email 
			message : Ought to be in HTML Format
			contacto: Email contact
		"""
		header=MIMEMultipart()
		header['Subject']= asunto
		header['From']=self.email
		header['To']=contacto
		mensaje = MIMEText(message, "html")#Content-type /html
		header.attach(mensaje)
		if file!=None and os.path.isfile(file):
			attached=MIMEBase('application', 'octet-stream')
			attached.set_payload(open(file, "rb").read())
			encode_base64(attached)
			attached.add_header('Content-Disposition', 'attachment; filename="%s"' % os.path.basename(file))
			header.attach(attached)
		self.server.sendmail(self.email,contacto, header.as_string())

	def finalize(self):
			"""
			To end the process, put this function to close the connection with the
			server
			"""
			self.server.quit()
"""
file=sys.argv[1]
contactos=["archivosian508@gmail.com", "gonzalezlin24@gmail.com", "canitogarciavazquez@gmail.com"]

s= Sender_Improved("canitogarciavazquez@gmail.com","Maquinadeguerra2",contactos)
mensaje= \"""<html><h1>  Espero que se vea cool</h1>
	<body>Sí le sé </body>

 </html>\"""

for contacto in contactos:
	s.send_message("ENVÍO DE PRUEBA CON IMAGEN", mensaje, file, contacto)

s.finalize()
"""