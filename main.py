from Mailer import Sender_Improved as Sender
import xlrd
import sys
import re
from tkinter import StringVar, filedialog,messagebox
import tkinter as tk
import socket

class Receiver:
	"""
	Class to get the data from user, and connect with Sende_Improve from
	Mailer
	Params:
		main_mail:str:Correo electrónico del remitente
		password:str: Contraseña del remitente
		file: Archivo adjunto, solo se permite uno por correo
		file_xlxs_path:str: Ruta absoluta de la base de datos en formato xlsx
	"""
	def __init__(self, main_mail, password, file, file_xlxs_path):
		self.file=file
		self.main_mail=main_mail
		self.password=password
		self.contactos:list=self.get_data_from_xlxs(file_xlxs_path, 1)
		self.nombres:list=self.get_data_from_xlxs(file_xlxs_path,2)
		if not self.es_correo_valido(self.main_mail):
			messagebox.showerror("Error","Upssss!, Al parecer tu correo no es válido")
			return

	def es_correo_valido(self, correo):
		expresion_regular = r"(?:[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*|\"(?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21\x23-\x5b\x5d-\x7f]|\\[\x01-\x09\x0b\x0c\x0e-\x7f])*\")@(?:(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?|\[(?:(?:(2(5[0-5]|[0-4][0-9])|1[0-9][0-9]|[1-9]?[0-9]))\.){3}(?:(2(5[0-5]|[0-4][0-9])|1[0-9][0-9]|[1-9]?[0-9])|[a-z0-9-]*[a-z0-9]:(?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21-\x5a\x53-\x7f]|\\[\x01-\x09\x0b\x0c\x0e-\x7f])+)\])"
		return re.match(expresion_regular, correo) is not None

	def send_message(self, asunto, message, ventana=None):
		"""
		ENvía los mensajes a cada correo electrónico, en caso de error, notifica la posiblidad de
		que alguno de los contactos sea incorrecto
		Param:
			message:str : Puede estar escrito en Html
			asunto:str : Asunto del correo
		"""
		flag=False
		if ventana!=None:
			flag=True
			etiqueta_contactos=tk.Label(ventana, text="Por enviar --")
			etiqueta_contactos.place(x="100",y="300")
		sender=Sender(self.main_mail, self.password, self.contactos)
		for contacto in self.contactos:
			if not self.es_correo_valido(contacto):
				messagebox.showinfo(message="Usuario erronéo: "+ contacto+ " ,será omitido", title="Error en ejecución")
				continue
			try:
				sender.send_message(asunto, message, contacto, file=self.file)
			except:
				messagebox.showinfo(message="Archivo erroneo ó Posible usuario erronéo: "+ contacto, title="Error en ejecución")
		sender.finalize()
		etiqueta_contactos=tk.Label(ventana, text="Proceso finalizado")
		etiqueta_contactos.place(x="100",y="300")
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
		"""
		Inicializa y defina la ventana por defecto, así como sus atributos
		"""
		if not self.net_connected():
			messagebox.showwarning("Internet te extraña","Parece que no estás conectado a internet, verifica tu conexión :(, Mailer no podrá notificarte si es que los correo no logran ser entregados\n ")

		self.ventana=tk.Tk(className="Mailer")
		self.ventana.geometry("800x500")
		self.mail_label=tk.Label(self.ventana, text="Email")
		self.mail_label.place(x=20, y=20)
		self.email=None#tk.StringVar()
		self.caja_email=tk.Entry(self.ventana,textvariable=self.email)
		self.caja_email.place(x=100, y=20)
		self.password_label=tk.Label(self.ventana, text="Password")
		self.password_label.place(x=20, y=40)
		self.password=None#tk.StringVar()
		self.caja_password=tk.Entry(self.ventana,textvariable=self.password)
		self.caja_password.place(x=100, y=40)
		self.caja_password.config(show="*")
		self.asunto_label=tk.Label(self.ventana, text="Asunto")
		self.asunto_label.place(x=20, y=60)
		self.asunto=None
		self.caja_asunto=tk.Entry(self.ventana, textvariable=self.asunto)
		self.caja_asunto.place(x=100,y=60)
		self.mensaje_label=tk.Label(self.ventana, text="Mensaje")
		self.mensaje_label.place(x=20,y=80)
		self.Mensaje=None#tk.StringVar()
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

		self.boton_guardar_xlxs=tk.Button(self.ventana, text="Abrir Base de Datos", command=self.guardar_archivo_xlxs)
		self.boton_guardar_xlxs.pack(side=tk.BOTTOM)

		self.boton_enviar=tk.Button(self.ventana, text="Enviar", command=self.enviar)
		self.boton_enviar.pack(side=tk.BOTTOM)
		self.xlsx=None
		self.files=None
		self.advertisement_xlsx=None
		self.advertisement_label=None
		self.ventana.mainloop()

	def guardar_archivo_xlxs(self):
		"""
		Guarda la ruta absoluta de nuestro archivo bbdd, solo acepta archivos xlxs
		"""
		self.xlsx=tk.filedialog.askopenfilename(filetypes = (("xlsx files","*.xlsx"),("xlsx files","*.xlxs*")))
		self.advertisement_xlsx=tk.Label(self.ventana, text="BBDD seleccionada: "+str(self.xlsx))
		self.advertisement_xlsx.place(x=100, y=250)


	def guardar_archivos(self):
		"""
		Guarda la ruta absoluta de un archivo y notifica creando una etiqueta en la ventana
		"""
		self.files=tk.filedialog.askopenfilename()
		self.advertisement_label=tk.Label(self.ventana, text="File seleccionada: "+str(self.files))
		self.advertisement_label.place(x=100, y=200)

	def enviar(self):
		"""
		Comportamiento designado tras presionar el boton enviar
		forma un objeto receiver, verifica que los datos sean consistentes
		,envía avisos en caso de no marcar las bbdd de referencia.
		AL finalizar muestra un cuadro para notificar que todo fue correcto
		"""
		self.email=self.caja_email.get()
		messagebox.showinfo(message="Comenzando envíos", title="Proceso iniciado")
		#print("email: ", self.email)
		self.password=self.caja_password.get()
		self.asunto=self.caja_asunto.get()
		self.Mensaje=self.caja_texto.get("1.0","end-1c")
		if self.email=="" or self.password=="" or self.asunto=="" or self.Mensaje=="":
			messagebox.showerror(title="Campos vacíos", message="Algunos de los campos principales están vacíos, verifica \n IMPOSIBLE ENVIAR")
		if self.xlsx == None:
			messagebox.showerror(title="Campos vacíos", message="Base de datos no seleccionada, verifique que sea en formato .xlxs")
			return
		if self.files==None:
			if not messagebox.askyesno(message="¿Desea continuar?", title="No contiene archivo adjunto"):
				return

		r=Receiver(self.email, self.password, self.files, self.xlsx)
		r.send_message(self.asunto, self.Mensaje, ventana=self.ventana)
		messagebox.showinfo(message="Mensaje enviado a cada contacto con éxito", title="Proceso finalizado")

	def net_connected(self):
		s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
		s.settimeout(5)
		try:
			s.connect(("www.google.com", 80))
		except (socket.gaierror, socket.timeout):
			return False
		return True










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
