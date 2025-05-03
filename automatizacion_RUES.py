# %% [markdown]
# ## Extracción de información de la página del Rues con autenticación

# %% [markdown]
# <p><b>TIP:</b> Para descomentar una línea de código en Visual Studio Code, puede señalar la línea que se quiera descomentar y utilizar el comando <b>'Ctrl + k + u'</b>. Por el contrario, para comentar una línea puede usar el comando <b>'Ctrl + k + c'</b>. En caso de ya tenerlos instalados no se debe hacer nada.</p><p><b>NOTA:</b> Se debe descargar <b>Chrome Driver</b> que es el navegador de prueba que permitirá ejecutar la automatización, se recomienda instalarlo en el escritorio ya que el código está para que lo ejecute ahí. Puede descargar Chrome Driver en el siguiente link: <a href='URL'>https://developer.chrome.com/docs/chromedriver/downloads?hl=es-419</a>
# </p>

# %% [markdown]
# <p>Librerías a cargar, debemos cargar <b>Selenium</b> (automatización de procesos), <b>Beautifulsoup</b> (web scrapping), <b>pandas</b>, <b>time</b> y <b>pathlib</b> (generalizar rutas).</p> El bloque de abajo detecta automáticamente si ya se tienen las librerías necesarias. En caso de no tenerlas instaladas este mismo las instalará automáticamente

# %%
import subprocess
import sys

print('Fase 1: Instalación de Paquetes.')
print('Verificando instalación de librerías')

# Lista de librerías requeridas
required_packages = [
    ("pandas", "pandas"),
    ("beautifulsoup4", "bs4"),
    ("selenium", "selenium"),
    ("time", "time"),
    ("pathlib", "pathlib"),
    ("customtkinter", "customtkinter")
]

def install_package(package_name):
    """Instala un paquete usando pip."""
    subprocess.check_call([sys.executable, "-m", "pip", "install", package_name])

def check_and_install_packages(packages):
    """Comprueba si los paquetes están instalados, y los instala si no lo están."""
    for package_name, module_name in packages:
        try:
            __import__(module_name)
            print(f"{package_name} ya está instalado.")
        except ImportError:
            print(f"{package_name} no está instalado. Instalando...")
            install_package(package_name)

# Comprobar e instalar los paquetes necesarios
check_and_install_packages(required_packages)

# Ejemplo de uso de las librerías (para verificar que se instalaron correctamente)
import sys
import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException, ElementNotInteractableException
from bs4 import BeautifulSoup
import pandas as pd
import time
from pathlib import Path
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox as mb
from tkinter import PhotoImage
from customtkinter import *

print("Todas las librerías están instaladas y listas para usar.")

print('Fase 1 completada.')

# %% [markdown]
# <p><b>Ingresar los NITs a extraer información</b>. Si los tiene en un archivo de excel se pueden cargar con lel siguiente bloque de código abajo, solo debe especificar la ruta donde está el archivo y luego especificar el nombre de la columna donde están los nits en el archivo.</p>

# %%
print('Fase 2: Extracción de información desde la web.')

from customtkinter import *
from tkinter import filedialog

file_path = None

# Función para manejar la selección de archivos Excel
def open_file_dialog():
    global file_path
    file_path = filedialog.askopenfilename(title="Seleccionar archivo Excel", 
                                           filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        label_file.configure(text=f"Archivo seleccionado: {file_path}")
        # Leer el archivo Excel y guardar la columna NIT como una lista
        try:
            df = pd.read_excel(file_path)  # Cargar el archivo Excel
            globals()["nits"] = df['NIT']  # Guardar la columna NIT como una lista
            print(f"NITs cargados: {len(nits)}")  # Para verificar que los NITs se han cargado
            nits.info()
        except Exception as e:
            label_file.configure(text=f"Error al leer el archivo: {e}")
            print(f"Error al leer el archivo: {e}")

app = CTk()
app.geometry("500x400")

set_appearance_mode("dark")
set_default_color_theme("dark-blue")

# Función para cerrar el evento cuando se dé click en x:
def on_closing():
    if not entry_username.get() or not entry_password.get():
        result = mb.askyesno("Advertencia", "Los campos estan vacíos. ¿Desea cerrar la aplicación de todos modos?")
        if result:
            app.destroy()
            sys.exit()
    else:
        app.destroy()
        sys.exit()

# Labels
label_titulo = CTkLabel(master=app, text="Consulta Información RUES", font=("Arial", 30),
                        text_color="#4aaea7").place(relx=0.5, rely=0.2, anchor="center")

label_username = CTkLabel(master=app, text="Usuario RUES", font=("Arial", 20),
                          text_color="#e90574").place(relx=0.3, rely=0.4, anchor="center")

label_password = CTkLabel(master=app, text="Contraseña", font=("Arial", 20),
                          text_color="#e90574").place(relx=0.7, rely=0.4, anchor="center")

# Input info
entry_username = CTkEntry(master=app, placeholder_text="Username RUES", width=150,
                          text_color="#e90574")
entry_username.place(relx=0.3, rely=0.5, anchor="center")

entry_password = CTkEntry(master=app, placeholder_text="Contraseña RUES", width=150,
                          text_color="#e90574", show="*")
entry_password.place(relx=0.7, rely=0.5, anchor="center")

### Botón seleccionador de año

# opciones de años 
def actualizar_ano(opcion):
    global ano
    ano = opcion

# Lista de años
anos = [str(year) for year in range(2010, 2025)]

# Menú de años
menu_anos = CTkOptionMenu(app, values=anos, command=actualizar_ano, button_color="#243e94", fg_color="#243e94")
menu_anos.place(relx=0.7, rely=0.7, anchor="center")
menu_anos.set("2024")

# Función para manejar el click del botón
def click_handler():
    if not entry_username.get() or not entry_password.get():
        mb.showwarning(title='ADVERTENCIA', message="Llene todos los campos antes de cerrar")
    elif file_path is None:
        mb.showwarning("ADVERTENCIA", message="Por favor, selecciona un archivo antes de buscar.")
    else:
        print("Button clicked")
        globals()["username"] = entry_username.get()
        globals()["password"] = entry_password.get()
        print(username)
        print(password)
        app.destroy()

# Botón para realizar la acción de buscar
btn = CTkButton(master=app, text="Buscar", corner_radius=32, fg_color="#73378c",
                hover_color="#73378c", command=click_handler)
btn.place(relx=0.5, rely=0.8, anchor="center")

# Botón para abrir el explorador de archivos
btn_file = CTkButton(master=app, text="Seleccionar Archivo Excel", corner_radius=32, 
                     fg_color="#e90574", hover_color="#4158D0", command=open_file_dialog)
btn_file.place(relx=0.3, rely=0.7, anchor="center")

# Etiqueta para mostrar el archivo seleccionado
label_file = CTkLabel(master=app, text="No se ha seleccionado archivo", font=("Arial", 14),
                      text_color="#4aaea7")
label_file.place(relx=0.5, rely=0.6, anchor="center")

# Interceptar el cierre de la ventana
app.protocol("WM_DELETE_WINDOW", on_closing)

# Iniciar la aplicación
app.mainloop()

# Imprimir la lista
print(f'Hay {len(nits)} nits a los que se les buscará la información')

# %% [markdown]
# **Iniciar sesión**. Se debe ingresar el usuario y contraseña

# %% [markdown]
# Abrir el chromeDriver. Aquí se debe actualizar la ruta donde se haya instalado el ChromeDriver. Por recomendación, lo mejor es instalarlo en la aplicación de escritorio (el usuario se llama practeec, en caso de que ya no sea así debe cambiarlo). Además iniciaremos sesión automaticamente, aunque también se puede hacer de forma manual (si el usuario y/o contraseña cambia, deben cambairse estos valores en el código)

# %%
# Ruta al ejecutable de ChromeDriver (ajusta esta ruta según donde hayas descargado el driver)
chrome_driver_path = Path.home() / 'Downloads' / 'Chromedriver.exe'
# Crear una instancia del servicio de ChromeDriver
service = Service(chrome_driver_path)
# Inicializar el navegador usando el servicio
driver = webdriver.Chrome(service=service)
# URL de la página de login
login_url = "https://www.rues.org.co/?old=true"
# Ruta carpeta para guardar el archivo
rues_descarga_path = Path.home() / 'Downloads' / f'info_rues_{ano}.xlsx'

# Navegar a la página de login
driver.get(login_url)

# Darle tiempo (5 segundos)
time.sleep(5)

# Botón cerrar emergente
xpath_boton_cerrar_emergente = "//button[@class='btn btn-success btn-sm']"
boton_cerrar_emergente = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, xpath_boton_cerrar_emergente)))
driver.execute_script("arguments[0].scrollIntoView();", boton_cerrar_emergente)
boton_cerrar_emergente.click()

#Hacer click en el botón de acceso
WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.CLASS_NAME, 'login-call'))).click()

# Encontrar los campos de login y enviar las credenciales
try:
    email_element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, 'Email')))
    password_element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, 'Password')))
    
    # Desplazarse a los elementos para asegurarse de que están visibles
    driver.execute_script("arguments[0].scrollIntoView();", email_element)
    driver.execute_script("arguments[0].scrollIntoView();", password_element)
    
    email_element.send_keys(username)
    password_element.send_keys(password)
except (TimeoutException, ElementNotInteractableException) as e:
    print("No se pudo interactuar con los campos de login:", e)
    driver.quit()
    raise

# Encontrar el botón por su texto y hacer clic
driver.find_element(By.CLASS_NAME, 'btn-blue').click() # Ajusta según el botón de submit

############## CICLO FOR PARA LOS N NITS ####################

# Matriz a rellenar más adelante
data = {
    'NIT': None,
    'Activo_Total': None,
    'Ingresos_Actividad_Ordinaria': None,
}

def dar_click_boton_regresar():
    boton_regresar_xpath = "//a[@class='btn-gt']"
    ver_boton_regresar = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, boton_regresar_xpath)))
    WebDriverWait(driver, 5).until(EC.visibility_of(ver_boton_regresar))
    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, boton_regresar_xpath)))
    driver.execute_script("arguments[0].scrollIntoView(true);", ver_boton_regresar)
    driver.execute_script("arguments[0].click();", ver_boton_regresar)

# Ciclo for para los n NITs enlistados
results = []

for nit in nits:
   
    try:
        # Esperar hasta que el campo de "Número de identificación" esté presente y realizar la consulta
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'txtNIT')))
        driver.find_element(By.ID, 'txtNIT').send_keys(nit)
        driver.find_element(By.ID, "btnConsultaNIT").click()

        # Esperar hasta que el botón esté presente y hacer clic
        
        # Encontrar la celda que contiene el texto "ACTIVA"
        try:
            activa_td_xpath = "//table[@id='rmTable2']//tbody//tr//td[text()='ACTIVA']"
            activa_td = WebDriverWait(driver, 3).until(
                EC.presence_of_element_located((By.XPATH, activa_td_xpath))
            )
        except (TimeoutException, NoSuchElementException, NameError):
            data = {'NIT': nit,
                'Activo_Total': 'No se encuentra la caja de activa',
                'Ingresos_Actividad_Ordinaria': 'No se encuentra la caja de activa'
                }
            results.append(data)
            driver.find_element(By.ID, 'txtNIT').clear()
            continue

        else:    
            # Encontrar la fila padre de la celda encontrada
            parent_row = activa_td.find_element(By.XPATH, "./ancestor::tr")
            # Encontrar la celda con tabindex="0" en la misma fila
            target_td = parent_row.find_element(By.XPATH, ".//td[@tabindex='0']")
            # Esperar a que la celda con tabindex="0" sea clicable
            WebDriverWait(driver, 3).until(
                EC.element_to_be_clickable((By.XPATH, ".//td[@tabindex='0']"))
            )
            # Hacer clic en la celda con tabindex="0"
            target_td.click()

            # Dar click en el link de info
            encontrar_link = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.LINK_TEXT, "nfo")))
            encontrar_link.click()

            # Dar click en el botón de info financiera
            try:
                xpath_ver_info_financiera = "(//span[contains(text(), 'Ver información')])[1]"
                ver_info_link = WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.XPATH, xpath_ver_info_financiera)))
                WebDriverWait(driver, 3).until(EC.visibility_of(ver_info_link))
                WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, xpath_ver_info_financiera)))
                driver.execute_script("arguments[0].scrollIntoView(true);", ver_info_link)
                driver.execute_script("arguments[0].click();", ver_info_link)
            except (TimeoutException, NoSuchElementException, NameError):
                data = {'NIT': nit,
                    'Activo_Total': 'No hay información financiera',
                    'Ingresos_Actividad_Ordinaria': 'No hay información financiera'
                    }
                results.append(data)
                dar_click_boton_regresar()
                continue
            else: 
                # Dar click en el botón del año a buscar
                try:
                    xpath_ano = f"//div[@id='{ano}']//button"
                    ver_ano_link = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, xpath_ano)))
                    WebDriverWait(driver, 3).until(EC.visibility_of(ver_ano_link))
                    WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, xpath_ano)))
                    driver.execute_script("arguments[0].scrollIntoView(true);", ver_ano_link)
                    driver.execute_script("arguments[0].click();", ver_ano_link)
                except (TimeoutException, NoSuchElementException, NameError):
                    data = {'NIT': nit,
                        'Activo_Total': 'No hay información financiera',
                        'Ingresos_Actividad_Ordinaria': 'No hay información financiera',
                        }
                    results.append(data)
                    dar_click_boton_regresar()
                    continue
                else:

                    # Esperar hasta que la tabla esté presente
                    xpath_tabla_ano = f"//div[@id='collapse-{ano}']//table"
                    WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.XPATH, xpath_tabla_ano)))

                    # Extraer el contenido de la página
                    page_source = driver.page_source
                    # Analizar el contenido con BeautifulSoup
                    soup = BeautifulSoup(page_source, 'html.parser')

                    # Buscar la tabla con datos
                    table = soup.select_one(f'div#collapse-{ano} table')
                    data['NIT'] = nit
                    if table:
                        # Extraer todas las filas de la tabla
                        rows = table.find_all('tr')
                        for row in rows:
                            # Extraer todas las celdas de la fila
                            cells = row.find_all('td')
                            cell_data = [cell.text.strip() for cell in cells]
                            if 'Activo Total' in cell_data[0]:
                                data['Activo Total'] = cell_data[1]
                            elif 'Ingresos Actividad Ordinaria' in cell_data[0]:
                                data['Ingresos Actividad Ordinaria'] = cell_data[1]
                    else:
                        print("No se encontró la tabla de datos")
                        data = {
                            'Activo_Total': 'No se encontró la tabla de datos',
                            'Ingresos_Actividad_Ordinaria': 'No se encontró la tabla de datos'
                        }
                        results.append(data)
                        dar_click_boton_regresar()
                        continue

                # Agregar los datos del NIT actual a la lista de resultados
                results.append(data)
                # Dar click en el botón de regresar para que se pueda iniciar la otra consulta
                dar_click_boton_regresar()

    except (TimeoutException, NoSuchElementException, NameError):
        data = {
            'NIT': nit,
            'Activo_Total': 'Problema al interactuar con el NIT',
            'Ingresos_Actividad_Ordinaria': 'Problema al interactuar con el NIT',
        }
        results.append(data)
        driver.find_element(By.ID, 'txtNIT').clear()
        continue

print(f'Fase 2 completada. Se han extraído exitosamente la información financiera de {len(nits)} nits')
print('Fase 3. Descarga del archivo a excel')
# Limpiar y convertir los valores a numérico
# Convertir los resultados a un DataFrame de pandas para facilitar la manipulación
if results:
    df = pd.DataFrame(results)
    # Crear el archivo excel 
    df.to_excel(rues_descarga_path, index=False)

# Cerrar el navegador
driver.quit()

print(df.info())

print(f'Se ha descargado el archivo en la ruta: {rues_descarga_path}')
print('Fase 3 completada.')
# %%
