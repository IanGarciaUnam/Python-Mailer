from Mailer import Sender_Improved as Sender
import xlrd
import sys
import re
from tkinter import StringVar, filedialog,messagebox,ttk
import tkinter as tk
import threading
import socket
import os
import time
from io import open
from validate_email import validate_email
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
		#self.nombres:list=self.get_data_from_xlxs(file_xlxs_path,2)
		if not self.es_correo_valido(self.main_mail):
			messagebox.showerror("Error","Upssss!, Al parecer tu correo no es válido")
			return

	def es_correo_valido(self, correo):
		return validate_email(correo) and "@" in correo

	def send_message(self, asunto, message, ventana, progressbar:ttk.Progressbar):
		"""
		ENvía los mensajes a cada correo electrónico, en caso de error, notifica la posiblidad de
		que alguno de los contactos sea incorrecto
		Param:
			message:str : Puede estar escrito en Html
			asunto:str : Asunto del correo
		"""
		MAXIMALE=len(self.contactos)
		sender=Sender(self.main_mail, self.password, self.contactos)

		if self.contactos==None:
			messagebox.showerror("Credenciales inválidas", "Correo o contraseñas incorrectas, tus correos no pueden ser enviados")
			return
		self.cantidad_de_perdidos=0
		self.perdidos=""	
		contador:int=0
		for i in range(len(self.contactos)):
			if not self.es_correo_valido(self.contactos[i]):
				self.perdidos+=self.contactos[i]+"\n"
				self.cantidad_de_perdidos+=1
				continue
			if i%50==0 and i>0:
				sender.finalize()
				time.sleep(300)
				sender=Sender(self.main_mail, self.password, self.contactos)
			
			if i%10==0:
				progressbar['value']=((i*100)/MAXIMALE)# 200==100
				ventana.update_idletasks()
				time.sleep(1)
			contador+=1
			try:
				#print(self.contactos[i],i)
				sender.send_message(asunto, message, self.contactos[i], file=self.file)
			except:
				if self.contactos[i] == " ":
					messagebox.showinfo(message="Usuario vacío en casilla "+str(i+2) , title="Error en adquisición")
					self.perdidos+="Casilla vacía: "+str(i+2)
					self.cantidad_de_perdidos+=1
				else:
					messagebox.showinfo(message="Archivo erroneo ó Posible usuario erronéo: "+ self.contactos[i], title="adquisición")
					self.perdidos+=str(self.contactos[i])
					self.cantidad_de_perdidos+=1

		progressbar['value']=100
		ventana.update_idletasks()
		time.sleep(1)
		sender.finalize()
		etiqueta_contactos=tk.Label(ventana, text="Proceso finalizado, mensajes enviados: "+str(contador))
		etiqueta_contactos.place(x="100",y="300")
		messagebox.showinfo(message="Mensajes enviados exitosamente", title="Proceso finalizado")

		if self.cantidad_de_perdidos>0:
			messagebox.showinfo(message="Un total de "+str(self.cantidad_de_perdidos)+" contactos no localizables, el reporte correspondiente será escrito", title="Sobre contactos perdidos")
			self.write_report(self.perdidos)
		print("Succesfully sent")

	def write_report(self, content):
		"""
		Write a report in txt
			Param: 
				content=Lecture gotten
		"""
		textfile=open("Report.txt", "w")
		textfile.write(content)
		textfile.close()


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
			try:
				agenda.append(str(sheet.cell_value(i,x)))#We have to acces to the column (1), and the row (i)
			except:
				messagebox.showerror(title="Campos vacíos (ERROR)", message="Al parecer la base de datos contiene espacios vacíos, imposible procesar")
		if len(agenda)<1:
			messagebox.showerror(title="Base de datos vacía", message="Verifica que tu base de datos no este vacía, y que los correos eléctrónicos se ubiquen en la columna 1 y despues de la fila 1")
			return
		else:
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
		self.caja_texto.place(x=100, y=80)
		self.scroll.config(command=self.caja_texto.yview)
		self.boton_archivos=tk.Button(self.ventana, text="Adjuntar archivo", command=self.guardar_archivos)
		self.boton_archivos.pack(side=tk.BOTTOM)
		self.boton_guardar_xlxs=tk.Button(self.ventana, text="Abrir Base de Datos", command=self.guardar_archivo_xlxs)
		self.boton_guardar_xlxs.pack(side=tk.BOTTOM)
		self.boton_enviar=tk.Button(self.ventana, text="Enviar", command=self.enviar)
		self.boton_enviar.pack(side=tk.BOTTOM)
		self.xlsx=None
		self.files=None
		self.advertisement_xlsx=None
		self.advertisement_label=None
		self.progreso_label=tk.Label(self.ventana, text="Progreso")
		self.progreso_label.place(x=20, y=180)
		self.progressbar=ttk.Progressbar(self.ventana,length=100, mode="determinate")
		self.progressbar.place(x=100, y=180, width=300)
		self.ventana.mainloop()#Cualquier elemento debe ser agregado previo a esta instrucción


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
		self.advertisement_label=tk.Label(self.ventana, text="Archivo adjunto: "+str(self.files))
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
			return
		if self.xlsx == None or not os.path.isfile(self.xlsx):
			messagebox.showerror(title="Campos vacíos", message="Base de datos no seleccionada, verifique que sea en formato .xlxs")
			return
		if self.files==None:
			if not messagebox.askyesno(message="No contiene archivo adjunto¿Desea continuar?", title="Aviso"):
				return
		self.contador_etiqueta=tk.Label(self.ventana, text="Comienzan envíos")
		self.contador_etiqueta.place(x="100",y="300")
		r=Receiver(self.email, self.password, self.files, self.xlsx)
		r.send_message(self.asunto, self.Mensaje, self.ventana, self.progressbar)


		return
			#t3=threading.Thread(target=self.progressbar.stop)
			#t3.start()
		#self.progressbar.stop()

	def net_connected(self):
		s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
		s.settimeout(5)
		try:
			s.connect(("www.google.com", 80))
		except (socket.gaierror, socket.timeout):
			return False
		return True










if __name__=="__main__":
	Window=Ventana()
