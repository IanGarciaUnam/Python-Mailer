from Mailer import Sender_Improved as Sender
import xlrd
import sys
from tkinter import StringVar, filedialog
import tkinter as tk

class Receiver:
	"""
	Class to get the data from user
	"""
	def __init__(self, main_mail, password, file, file_xlxs_path):
		self.file=file
		self.main_mail=main_mail
		self.password=password
		self.contactos:list=self.get_data_from_xlxs(file_xlxs_path, 1)
		self.nombres:list=self.get_data_from_xlxs(file_xlxs_path,2)

	def send_message(self, asunto, message):
		sender=Sender(self.main_mail, self.password, self.contactos)
		for contacto in self.contactos:
			sender.send_message(asunto, message, contacto, file=self.file)
		sender.finalize()
		print("Succesfully sent")

	def get_data_from_xlxs(self, file_path, x):
		"""
		Returns a list with email of each contact, it
		uses xlrd library
			Param:
				file_path: XLXS file path
				x: 
		"""
		agenda=[]
		loc=str(file_path)
		wb=xlrd.open_workbook(loc)
		sheet = wb.sheet_by_index(0)#Sheet file
		for i in range(sheet.nrows):
			agenda.append(str(sheet.cell_value(i,x)))#We have to acces to the column (1), and the row (i)
		del agenda[0]#The first element will always be 'correo electrónico'
		return agenda

class Ventana:

	def __init__(self):
		self.ventana=tk.Tk()
		self.ventana.geometry("800x400")
		self.mail_label=tk.Label(self.ventana, text="Email")
		self.mail_label.place(x=20, y=20)
		self.email=tk.StringVar()
		self.caja_email=tk.Entry(self.ventana,textvariable=self.email)
		self.caja_email.place(x=100, y=20)
		self.password_label=tk.Label(self.ventana, text="Password")
		self.password_label.place(x=20, y=40)
		self.password=tk.StringVar()
		self.caja_password=tk.Entry(self.ventana,textvariable=self.password)
		self.caja_password.place(x=100, y=40)
		self.caja_password.config(show="*")
		self.asunto_label=tk.Label(self.ventana, text="Asunto")
		self.asunto_label.place(x=20, y=60)
		self.asunto=tk.StringVar()
		self.caja_asunto=tk.Entry(self.ventana, textvariable=self.asunto)
		self.caja_asunto.place(x=100,y=60)
		self.mensaje_label=tk.Label(self.ventana, text="Mensaje")
		self.mensaje_label.place(x=20,y=80)
		self.Mensaje=tk.StringVar()
		self.caja_texto=tk.Text(self.ventana,height=5, width=50)
		self.scroll=tk.Scrollbar(self.ventana)
		self.caja_texto.configure(yscrollcommand=self.scroll.set)
		#self.caja_texto.pack(side=tk.BOTTOM)
		self.caja_texto.place(x=100, y=80)
		self.scroll.config(command=self.caja_texto.yview)
		#self.caja_mensaje=Entry(self.ventana, textvariable=self.Mensaje)
		#self.caja_mensaje.place(x=100, y=80, height=100, width=100)
		self.boton_archivos=tk.Button(self.ventana, text="Abrir archivo", command=self.guardar_archivos)
		self.boton_archivos.pack(side=tk.BOTTOM)
		self.boton_enviar=tk.Button(self.ventana, text="Enviar", command=self.enviar)
		self.boton_enviar.pack(side=tk.BOTTOM)
		
		self.ventana.mainloop()


	def guardar_archivos(self):
		self.files=tk.filedialog.askopenfilenames()
		self.advertisement_label=tk.Label(self.ventana, text="File selected: "+str(self.files))
		self.advertisement_label.place(x=100, y=200)
	def enviar(self):
		correo=self.caja_email.get()
		print(correo)




if __name__=="__main__":
	#file=sys.argv[2]
	"""
	xlxs=sys.argv[1]
	r=Receiver("canitogarciavazquez@gmail.com", "Maquinadeguerra2", None,xlxs)
	mensaje="hola me agradas"
	r.send_message("Mensaje de Bienvenida",mensaje)
	"""
	#root=Tk()
	Window=Ventana()
	#root.mainloop()