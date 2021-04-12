from Mailer import Sender_Improved as Sender

class Receiver:
	"""
	Class to get the data from user
	"""
	def __init__(self, main_mail, password, file):
		self.file=file
		self.main_mail=main_mail
		self.password=password
		self.contactos=self.get_contacts(None)

	def send_message(self, asunto, message):
		sender= Sender(self.main_mail, self.password, self.contactos)
		for contacto in self.contactos:
			sender.send_message(asunto, message, contacto, file=self.file)
		sender.finalize()

	def get_contacts(self, idk):
		return ["archivosian508@gmail.com", "canitogarciavazquez@gmail.com"]

if __name__=="__main__":
	r=Receiver("canitogarciavazquez@gmail.com", "Maquinadeguerra2", None)
	mensaje= """<html><h1>  Espero que se vea cool</h1>
	<body>Sí le sé </body>
	</html>"""
	r.send_message("HOLA",mensaje)
