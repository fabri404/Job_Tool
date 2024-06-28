from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import tkinter as tk
from tkinter import messagebox, filedialog, ttk
import time
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from pytube import YouTube
from tkinter import *
import os
import webbrowser
import instaloader
from urllib.parse import urlparse, parse_qs
from pathlib import Path
from PIL import Image, ImageTk
from rembg import remove
import io
import threading

# Contraseña predefinida
PASSWORD = "3517141146404"

# Función para la ventana de autenticación
def authenticate():
    def check_password():
        if entry_password.get() == PASSWORD:
            auth_window.destroy()
            show_main_window()
        else:
            messagebox.showerror("Error", "Contraseña incorrecta")

    auth_window = tk.Tk()
    auth_window.title("Autenticación")
    auth_window.geometry("400x350")

    tk.Label(auth_window, text="Ingrese la contraseña:").pack(pady=10)
    entry_password = tk.Entry(auth_window, show="*")
    entry_password.pack(pady=5)
    tk.Button(auth_window, text="Ingresar", command=check_password).pack(pady=10)

    auth_window.mainloop()

# Función para la ventana principal con botones para cada script
def show_main_window():
    main_window = tk.Tk()
    main_window.title("Seleccionar Programa")
    main_window.geometry("400x350")
    
    def run_script1():
        
        def remove_background(input_path: str, output_path: str, progress_callback=None, callback=None):
            try:
                with open(input_path, "rb") as image_file:
                    input_image = image_file.read()

                if progress_callback:
                    progress_callback(10)  # Progreso inicial

                output_image = remove(input_image)

                if progress_callback:
                    progress_callback(80)  # Progreso intermedio

                img = Image.open(io.BytesIO(output_image)).convert("RGBA")
                img.save(output_path)

                if progress_callback:
                    progress_callback(100)  # Progreso completo

                if callback:
                    callback(None)  # Indica éxito
            except Exception as e:
                if callback:
                    callback(e)  # Indica error

        def select_image():
            file_path = filedialog.askopenfilename(
                filetypes=[
                    ("All image files", "*.png;*.jpg;*.jpeg;*.bmp;*.tiff;*.gif;*.ico;*.webp;*.jfif;*.pjpeg;*.pjp;*.svg")
                ]
            )
            if file_path:
                input_path.set(file_path)
                load_image(file_path)
                lbl_output.grid()
                btn_save.grid()
                progress_bar.grid()

        def save_image():
            file_path = filedialog.asksaveasfilename(
                defaultextension=".png",
                filetypes=[("PNG files", "*.png"), ("All files", "*.*")]
            )
            if file_path:
                output_path.set(file_path)
                btn_save.config(state=tk.DISABLED)
                progress_bar.start()
                threading.Thread(target=remove_background, args=(input_path.get(), output_path.get(), update_progress, on_save_complete)).start()

        def load_image(path):
            img = Image.open(path)
            img.thumbnail((300, 300))  # Ajusta la imagen para que no sea demasiado grande
            img_tk = ImageTk.PhotoImage(img)
            lbl_image.config(image=img_tk)
            lbl_image.image = img_tk
            lbl_image.config(width=300, height=300)

        def on_save_complete(error):
            progress_bar.stop()
            progress_bar.grid_remove()
            btn_save.config(state=tk.NORMAL)
            if error:
                messagebox.showerror("Error", str(error))
            else:
                exito_label.config(text="Hecho!")

        def update_progress(value):
            progress_bar['value'] = value
            root.update_idletasks()

        root = tk.Tk()
        root.title("Eliminador de Fondo de Imágenes")

        root.geometry("400x550")  
        root.resizable(False, False)  

        input_path = tk.StringVar()
        output_path = tk.StringVar()

        frm = tk.Frame(root, padx=10, pady=10)
        frm.pack(padx=10, pady=10)

        lbl_select = tk.Label(frm, text="Selecciona una imagen:")
        lbl_select.grid(row=0, column=0, padx=5, pady=5)

        btn_select = tk.Button(frm, text="Buscar", command=select_image)
        btn_select.grid(row=0, column=1, padx=5, pady=5)

        lbl_image = tk.Label(frm, text="Previsualización", width=40, height=20, bg="lightgray")
        lbl_image.grid(row=1, column=0, columnspan=2, pady=10)

        lbl_output = tk.Label(frm, text="Guardar como:")
        lbl_output.grid(row=2, column=0, padx=5, pady=5)
        lbl_output.grid_remove()

        btn_save = tk.Button(frm, text="Guardar", command=save_image)
        btn_save.grid(row=2, column=1, padx=5, pady=5)
        btn_save.grid_remove()

        progress_bar = ttk.Progressbar(frm, orient="horizontal", length=300, mode="determinate")
        progress_bar.grid(row=3, column=0, columnspan=2, pady=10)
        progress_bar.grid_remove()

        exito_label = tk.Label(frm, text="", fg="green")
        exito_label.grid(row=4, column=0, columnspan=2, pady=10)

        root.mainloop()   
          
    def run_script2():
        
        def seleccionar_imagenes():
            rutas_nuevas = filedialog.askopenfilenames(
                title="Seleccionar imágenes",
                filetypes=[
                    ("Image files", "*.jpg *.jpeg *.png *.bmp *.gif *.tiff *.tif *.ico *.webp *.jfif *.pjpeg *.pjp *.svg")
                ]
            )
            if rutas_nuevas:
                for ruta in rutas_nuevas:
                    nombre_archivo = os.path.basename(ruta)
                    if nombre_archivo not in lista_imagenes.get(0, tk.END):
                        lista_imagenes.insert(tk.END, nombre_archivo)
                        rutas_imagenes[nombre_archivo] = ruta
                previsualizar_imagen(rutas_nuevas[0])

        def previsualizar_imagen(ruta):
            imagen = Image.open(ruta)
            imagen.thumbnail((200, 200))
            imagen_tk = ImageTk.PhotoImage(imagen)
            preview_label.config(image=imagen_tk)
            preview_label.image = imagen_tk

        def convertir_imagenes():
            nombres_imagenes = lista_imagenes.get(0, tk.END)
            if not nombres_imagenes:
                messagebox.showerror("Error", "Por favor, selecciona al menos una imagen.")
                return

            formato_seleccionado = formato_var.get().lower()
            carpeta_descargas = r"C:\Users\Peluca\Downloads"

            exito_label.config(text="Convirtiendo...")

            errores = []
            for nombre in nombres_imagenes:
                ruta = rutas_imagenes[nombre]
                try:
                    imagen = Image.open(ruta)
                    nombre_archivo = os.path.splitext(nombre)[0] + f".{formato_seleccionado}"
                    ruta_salida = os.path.join(carpeta_descargas, nombre_archivo)
                    if formato_seleccionado == 'ico':
                        # Convertir y guardar como ICO
                        imagen.save(ruta_salida, format='ICO', sizes=[(256, 256)])
                    else:
                        # Convertir y guardar en el formato seleccionado
                        imagen.save(ruta_salida, format=formato_seleccionado.upper())
                except Exception as e:
                    errores.append(f"{ruta}: {str(e)}")

            if errores:
                messagebox.showerror("Errores", f"Algunas imágenes no pudieron ser convertidas:\n" + "\n".join(errores))
                exito_label.config(text="")
            else:
                exito_label.config(text="Hecho!")

        # Crear la ventana principal
        ventana = tk.Tk()
        ventana.title("Convertidor de Imágenes")

        # Botón para seleccionar imágenes
        seleccionar_btn = tk.Button(ventana, text="Seleccionar Imágenes", command=seleccionar_imagenes)
        seleccionar_btn.grid(row=0, column=0, padx=10, pady=10)

        # Lista de imágenes seleccionadas
        lista_imagenes = tk.Listbox(ventana, width=50, height=10)
        lista_imagenes.grid(row=1, column=0, padx=10, pady=10)
        lista_imagenes.bind('<<ListboxSelect>>', lambda e: previsualizar_imagen(rutas_imagenes[lista_imagenes.get(tk.ACTIVE)]))

        # Diccionario para almacenar las rutas completas de las imágenes
        rutas_imagenes = {}

        # Etiqueta de previsualización
        preview_label = tk.Label(ventana)
        preview_label.grid(row=1, column=1, padx=10, pady=10)

        # Menú desplegable para seleccionar el formato de salida
        formatos = ["ICO", "PNG", "JPG", "JPEG", "BMP", "GIF", "TIFF", "WEBP"]
        formato_var = tk.StringVar(value="ICO")
        formato_menu = tk.OptionMenu(ventana, formato_var, *formatos)
        formato_menu.config(indicatoron=0, compound='right')
        formato_menu.grid(row=2, column=0, padx=10, pady=10)

        # Botón para convertir las imágenes
        convertir_btn = tk.Button(ventana, text="Convertir Imágenes", command=convertir_imagenes)
        convertir_btn.grid(row=2, column=1, pady=10)

        # Etiqueta para mostrar el estado de la conversión
        exito_label = tk.Label(ventana, text="")
        exito_label.grid(row=3, column=0, columnspan=2, pady=10)

        # Iniciar el bucle principal de la interfaz gráfica
        ventana.mainloop()      
    
    def run_script3():
        
        def descargar_mp3(link, download_path):
            try:
                video = YouTube(link)
                mp3_filename = video.title + ".mp3"
                audio_stream = video.streams.filter(only_audio=True).first()
                audio_stream.download(output_path=download_path, filename=mp3_filename)
                return mp3_filename
            except Exception as e:
                print(e)
                return None


        def descargar_video_youtube(link, download_path, formato):
            try:
                video = YouTube(link)
                if formato == "MP3":
                    mp3_filename = descargar_mp3(link, download_path)
                    if mp3_filename:
                        return f"El audio en formato MP3 ha sido descargado como '{mp3_filename}' en la carpeta seleccionada."
                else:
                    descarga = video.streams.get_highest_resolution()
                    descarga.download(download_path)
                    return "El video en formato MP4 ha sido descargado en la carpeta seleccionada."
            except Exception as e:
                return f"No se pudo procesar el enlace: {e}"


        # Funciones para la descarga de video de Instagram
        def descargar_video_instagram(link, download_path):
            descarga = instaloader.Instaloader()

            try:
                parsed_url = urlparse(link)
                query_params = parse_qs(parsed_url.query)
                shortcode = query_params.get('igshid', [''])[0]

                if not shortcode:
                    path_components = parsed_url.path.strip('/').split('/')
                    shortcode_index = path_components.index('p') + 1
                    shortcode = path_components[shortcode_index]

                post = instaloader.Post.from_shortcode(descarga.context, shortcode)
                descarga.download_post(post, target=download_path)
                return "El contenido ha sido descargado en la carpeta seleccionada."
            except ValueError:
                return "El enlace no parece ser un enlace válido de Instagram."
            except Exception as e:
                return f"No se pudo procesar el enlace: {e}"

        # Funciones para el botón de Descargar
        def descargar_click():
            link = videos.get()

            if not link:
                messagebox.showwarning("ERROR", "Por favor ingresa un enlace válido.")
                return

            destino = destino_var.get()
            base_path = os.path.expanduser("~")
            download_path = os.path.join(base_path, "Downloads" if destino == "Descargas" else "Desktop")

            if "youtube.com" in link:
                formato = formato_var.get()
                mensaje = descargar_video_youtube(link, download_path, formato)

            elif "instagram.com" in link:
                try:
                    # Obtiene el directorio de descargas del sistema
                    download_path = Path.home() / "Downloads"
                    mensaje = descargar_video_instagram(link, download_path)
                except Exception as e:
                    return f"No se pudo procesar el enlace: {e}"
            else:
                mensaje = "Enlace no reconocido."

            messagebox.showinfo("Descarga completada", mensaje)


        # Configuración de la ventana principal
        root = Tk()
        root.config(bd=15)
        root.title("Descargas")

        # Agregar una opción para seleccionar el formato de descarga (solo para YouTube)
        formato_var = StringVar()
        formato_var.set("MP4")
        formato_label = Label(root, text="Seleccione un formato (solo para YouTube):")
        formato_label.grid(row=6, column=0, columnspan=2)

        mp4_radio = Radiobutton(root, text="MP4", variable=formato_var, value="MP4")
        mp3_radio = Radiobutton(root, text="MP3", variable=formato_var, value="MP3")
        mp4_radio.grid(row=7, column=0, columnspan=2)
        mp3_radio.grid(row=8, column=0, columnspan=2)

        # Agregar una opción para seleccionar la ubicación de descarga
        destino_var = StringVar()
        destino_var.set("Descargas")  # Valor predeterminado
        destino_label = Label(root, text="Seleccione una ubicación (solo para YouTube):")
        destino_label.grid(row=3, column=0, columnspan=2)

        descargas_radio = Radiobutton(root, text="Descargas", variable=destino_var, value="Descargas")
        escritorio_radio = Radiobutton(root, text="Escritorio", variable=destino_var, value="Escritorio")
        descargas_radio.grid(row=4, column=0, columnspan=2)
        escritorio_radio.grid(row=5, column=0, columnspan=2)

        # Crear el botón de Descargar y asociarlo con la función descargar_click
        boton = Button(root, text="Descargar", command=descargar_click)
        boton.grid(row=2, column=1)

        # Agregar la entrada para la URL
        texto = Label(root, text="Ingrese la URL e indique las opciones :\n"
                                "\n"
                                "´´Las URLS de instagram se guardaran en descargas\n"
                                "unicamente en formato MP4´´")
        texto.grid(row=0, column=1)
        videos = Entry(root)
        videos.grid(row=1, column=1)

        # Iniciar el bucle de la interfaz gráfica
        root.mainloop()
                
    def run_script4():
        
        def leer_direcciones():
            scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
            json_keyfile_path = r"D:\Fabri\cursado2023prgramacion\proyectos\Proyectos concretados\envios de mails automaticos\psyched-era-426120-i3-53869cc8b98a.json"
            creds = ServiceAccountCredentials.from_json_keyfile_name(json_keyfile_path, scope)

            client = gspread.authorize(creds)
            
            sheet = client.open('Proyecto de Automatizacion').sheet1
            data = sheet.col_values(1)[1:]  # Leer todas las direcciones de correo, omitiendo el título

            return data

        # Configuración del servidor SMTP y envío de correos electrónicos
        smtp_server = 'smtp.gmail.com'
        smtp_port = 587
        smtp_user = 'Fabri404lopez@gmail.com'
        smtp_password = 'yckt ittq szmc eiiq'

        def enviar_correo(electronico, asunto, mensaje, adjuntos=[]):
            msg = MIMEMultipart()
            msg['From'] = smtp_user
            msg['To'] = electronico
            msg['Subject'] = asunto

            msg.attach(MIMEText(mensaje, 'plain'))

            for adjunto in adjuntos:
                attachment = open(adjunto, 'rb')
                part = MIMEBase('application', 'octet-stream')
                part.set_payload((attachment).read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f"attachment; filename= {adjunto.split('/')[-1]}")
                msg.attach(part)

            server = smtplib.SMTP(smtp_server, smtp_port)
            server.starttls()
            server.login(smtp_user, smtp_password)
            server.send_message(msg)
            server.quit()

        # Interfaz gráfica
        def crear_interfaz():
            def seleccionar_adjuntos():
                archivos = filedialog.askopenfilenames(title="Seleccionar archivos")
                for archivo in archivos:
                    lista_adjuntos.insert(tk.END, archivo)

            def enviar_mensaje():
                asunto = entrada_asunto.get()
                mensaje = entrada_mensaje.get("1.0", tk.END)
                adjuntos = lista_adjuntos.get(0, tk.END)
                correos = leer_direcciones()

                progreso["maximum"] = len(correos)
                for i, correo in enumerate(correos):
                    enviar_correo(correo.strip(), asunto, mensaje, adjuntos)
                    progreso["value"] = i + 1
                    ventana.update_idletasks()

            ventana = tk.Tk()
            ventana.title("Envío de Correos")

            tk.Label(ventana, text="Asunto:").pack()
            entrada_asunto = tk.Entry(ventana, width=50)
            entrada_asunto.pack()

            tk.Label(ventana, text="Mensaje:").pack()
            entrada_mensaje = tk.Text(ventana, height=10, width=50)
            entrada_mensaje.pack()

            tk.Button(ventana, text="Seleccionar adjuntos", command=seleccionar_adjuntos).pack()
            lista_adjuntos = tk.Listbox(ventana, selectmode=tk.MULTIPLE)
            lista_adjuntos.pack()

            tk.Button(ventana, text="Enviar mensaje", command=enviar_mensaje).pack()

            progreso = ttk.Progressbar(ventana, orient="horizontal", length=300, mode="determinate")
            progreso.pack()

            ventana.mainloop()

        # Llamar a la función para crear la interfaz gráfica
        crear_interfaz()
     
    def run_script5():
        
        RUTA_CHROME = r"C:\Program Files\Google\Chrome\Application\chrome-win64\chrome.exe"
        RUTA_CHROMEDRIVER = r"D:\Fabri\cursado2023prgramacion\proyectos\Proyectos concretados\envios automatizados instagram\Version2024\chromedriver-win64\chromedriver.exe"

        # Configurar y ejecutar la automatización
        def iniciar_automatizacion():
            cuentas_destinatarias = entrada_cuentas.get("1.0", tk.END).strip().split('\n')  # Obtener las cuentas destinatarias
            mensajes = obtener_mensajes()

            # Configurar las opciones del navegador Chrome
            options = webdriver.ChromeOptions()
            options.add_argument("--incognito")  # Opcional: Modo incógnito u otras opciones
            options.binary_location = RUTA_CHROME  # Especificar la ruta del binario de Chrome
            options.add_argument(f"webdriver.chrome.driver={RUTA_CHROMEDRIVER}")  # Especificar la ruta del ChromeDriver

            driver = None  # Inicializar driver fuera del try para manejar el caso de excepción

            try:
                # Iniciar el navegador Chrome utilizando ChromeDriverManager
                driver = webdriver.Chrome(service=Service(RUTA_CHROMEDRIVER), options=options)

                # Iniciar sesión en Instagram
                login_ig(driver)

                # Enviar mensajes a las cuentas destinatarias
                enviar_mensajes(driver, cuentas_destinatarias, mensajes)

                messagebox.showinfo("Información", "Los mensajes han sido enviados correctamente.")

            except Exception as e:
                messagebox.showerror("Error", f"Ocurrió un error: {str(e)}")

            finally:
                if driver:
                    driver.quit()

        # Función para obtener mensajes desde la interfaz
        def obtener_mensajes():
            mensajes = []
            for entry in mensajes_entries:
                mensaje = entry.get().strip()
                if mensaje:
                    mensajes.append(mensaje)
            return mensajes

        # Función para iniciar sesión en Instagram
        def login_ig(driver):
            driver.get("https://www.instagram.com/")
            time.sleep(5)

            entrada_usuario = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.NAME, "username")))
            entrada_contra = driver.find_element(By.NAME, "password")

            entrada_usuario.send_keys(USUARIO_IG)
            entrada_contra.send_keys(CONTRASENA_IG)
            entrada_contra.send_keys(Keys.ENTER)
            time.sleep(5)

            try:
                # Manejar el diálogo de guardar información de inicio de sesión
                dialogo_guardar_info = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//div[contains(text(), 'Ahora no') and @role='button']")))
                dialogo_guardar_info.click()

                # Esperar hasta que aparezca el botón "Ahora no"
                siguiente_boton_ahora_no = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//button[contains(@class, '_a9--') and "
                                                            "contains(@class, '_a9_1') and text()='Ahora no']")))
                siguiente_boton_ahora_no.click()
            except Exception as e:
                print(f"No se encontró el diálogo de guardar información de inicio de sesión: {str(e)}")

        # Función para enviar mensajes a cuentas destinatarias
        def enviar_mensajes(driver, cuentas_destinatarias, mensajes):
            for cuenta in cuentas_destinatarias:
                try:
                    # Navegar hasta el perfil del destinatario
                    driver.get(f"https://www.instagram.com/{cuenta}/")
                    time.sleep(4)  # Espera 4 segundos para cargar el perfil completamente

                    # Hacer clic en el botón de enviar mensaje
                    boton_mensaje = WebDriverWait(driver, 15).until(
                        EC.element_to_be_clickable(
                            (By.XPATH,
                            "//div[contains(@class, 'x9f619')]/div[contains(@class, 'x1i10hfl') and contains(text(),"
                            " 'Enviar mensaje')]")))
                    driver.execute_script("arguments[0].click();", boton_mensaje) 

                    # Escribir y enviar mensajes
                    for mensaje in mensajes:
                        entrada_mensaje = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'x9f619')]//div[@aria-label='Mensaje']")))
                        entrada_mensaje.send_keys(mensaje)
                        entrada_mensaje.send_keys(Keys.ENTER)
                        time.sleep(1)

                    time.sleep(2)

                except Exception as e:
                    print(f"No se pudo enviar mensaje a {cuenta}: {str(e)}")

        # Función para añadir un nuevo campo de entrada de mensajes
        def agregar_mensaje_entry(event=None):
            new_entry = tk.Entry(root, width=50)
            new_entry.grid(row=len(mensajes_entries) + 1, column=1)
            mensajes_entries.append(new_entry)
            new_entry.focus_set()
            # Desvincular el evento anterior
            new_entry.unbind('<Return>')
            # Volver a vincular el evento Enter al nuevo campo de entrada
            new_entry.bind('<Return>', agregar_mensaje_entry)
            # Reubicar el botón "Iniciar Automatización" debajo de la última entrada
            boton_iniciar.grid(row=len(mensajes_entries) + 2, column=0, columnspan=2)

        # Función principal para iniciar la interfaz
        def iniciar_proyecto():
            global root
            root = tk.Tk()
            root.title("Automatización de Envío de Mensajes en Instagram")

            tk.Label(root, text="Cuentas destinatarias (una por línea):").grid(row=0, column=0, sticky='e')
            global entrada_cuentas
            entrada_cuentas = tk.Text(root, width=50, height=4)
            entrada_cuentas.grid(row=0, column=1)

            tk.Label(root, text="Mensajes:").grid(row=1, column=0, sticky='e')
            global mensajes_entries
            mensajes_entries = []
            new_entry = tk.Entry(root, width=50)
            new_entry.grid(row=1, column=1)
            mensajes_entries.append(new_entry)

            # Asociar la función agregar_mensaje_entry al evento Enter
            new_entry.bind('<Return>', agregar_mensaje_entry)

            # Botón "Iniciar Automatización"
            global boton_iniciar
            boton_iniciar = tk.Button(root, text="Iniciar Automatización", command=iniciar_automatizacion)
            boton_iniciar.grid(row=2, column=0, columnspan=2)

            root.mainloop()

        # Configurar las credenciales de Instagram
        USUARIO_IG = "Fabri404.lopez"  # Cambiar por tu usuario de Instagram
        CONTRASENA_IG = "ELEFANTES"   # Cambiar por tu contraseña de Instagram

        if __name__ == "__main__":
            iniciar_proyecto()
            
    def show_about_me():
        webbrowser.open_new("https://www.linkedin.com/in/qa-adriel")

    tk.Button(main_window, text="Borrar Fondos", command=run_script1).pack(pady=5)
    tk.Button(main_window, text="Convertir Imagenes", command=run_script2).pack(pady=5)
    tk.Button(main_window, text="Descargar Videos", command=run_script3).pack(pady=5)
    tk.Button(main_window, text="Enviar MAILS Automaticos", command=run_script4).pack(pady=5)
    tk.Button(main_window, text="Enviar Mensajes Instagram Automatico", command=run_script5).pack(pady=5)
    tk.Button(main_window, text="Mas Sobre el Creador", command=show_about_me).pack(pady=5)
    main_window.mainloop()

# Iniciar la aplicación mostrando primero la ventana de autenticación
if __name__ == "__main__":
    authenticate()