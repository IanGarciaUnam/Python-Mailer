## Python mailer

## Settings
It's necessary to able the smtp port of your current mail, the Mailer at this point its designed to work with smtp.google
It requires  python>=3.0

## Usage for Linux's Users
 ´´´
 $python main.py
 ´´´
 if you want to create a beautiful script in Linux

´´´
pyinstaller --windowed --onefile --icon=./mailme.ico main.py
´´´
##### --windowed
Quit the terminal to just use the GUI Interphace
#### --onefile
Just create An executable script
#### --icon
Define the icon for the script
#### main.py
The main class to execute in Python

## Usage for Window's Users
Just go to the directory dist, and apply double click in
´´´
main.exe
´´´
The option --Adjuntar archivo--, will let you to chose a file to attach at the mail, just one for each email.
The option --abrir base de datos-- will only leave you to select an xlsx format file, be sure that in the 2nd column, commonly known as B are located the emails. Mailer will only read the 1st sheet of the file.

##Requirements
*Your email ought to be type Gmail.

*Be sure to configure your email with:
 https://www.google.com/settings/security/lesssecureapps
 ==========================================================
                          Español

 Para Windows
 **Dirigete a la carpeta dist/
 **Da doble click en main.exe
 Tras llenar todos los campos requeridos (Es necesario llenar todos los campos)

 **Para adjuntar un archivo da click en el botón, abrir archivo, seleccionalo, y en la ventana principal quedará anexada la ruta absoluta de dicho archivo indicando que ha sido agregada, este paso es opcional

 **Para adjuntar la base de datos, es necesario que se encuentre en formato xlsx, asimismo en la 2da columna del archivo deberan encontrarse escritos los correos correspondientes; no es necesario preocuparse si alguno es incorrecto, será omitido tras un aviso. Este paso es indispensable

 **Para enviar, Presione el boton enviar, este deberá ser el último paso, tras presionar, recibirá un aviso del comienzo del proceso, en caso de no haber adjuntado ningún archivo será notificado si desea continuar o cancelar, tras ello el proceso de envío tomará un momento.Si durante el proceso se presentará alguna problemática con algún correo, recibirá una advertencia esperando por aprobación y podrá continuar.
 Al finalizar el proceso de envío será notificado del termino exitoso.

 *Si es necesario puede repetir el proceso.

 ### Nota importante
 El mensaje puede ser escrito en formato plano o HTML.
