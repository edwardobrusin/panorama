from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import time
import os
import glob
import pandas as pd

tiempo_inicio_script = time.time() - 2

# ==========================================
# 1. CONFIGURACIÓN DEL NAVEGADOR Y RUTAS
# ==========================================
opciones = webdriver.ChromeOptions()

# Definimos las rutas de trabajo dinámicas
try:
    PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
    if os.path.basename(PROJECT_ROOT) == 'scripts':
        PROJECT_ROOT = os.path.dirname(PROJECT_ROOT)
except NameError:
    PROJECT_ROOT = os.getcwd()
    if os.path.basename(PROJECT_ROOT) == 'scripts':
        PROJECT_ROOT = os.path.dirname(PROJECT_ROOT)

raw_dir = os.path.join(PROJECT_ROOT, "data", "raw")
intermediate_dir = os.path.join(PROJECT_ROOT, "data", "intermediate")

# Aseguramos que la carpeta intermediate exista
os.makedirs(intermediate_dir, exist_ok=True)

prefs = {"download.default_directory" : raw_dir}
opciones.add_experimental_option("prefs", prefs)

# Forzamos el idioma a español
opciones.add_argument("--lang=es-MX") 

driver = webdriver.Chrome(options=opciones)
# Entramos al mismo enlace (nos dejará en la pestaña por defecto)
driver.get("https://public.tableau.com/app/profile/imss.cpe/viz/Histrico_4/Empleo_h?publish=yes")

wait = WebDriverWait(driver, 20)

# ==========================================
# 2. PREPARACIÓN DEL DASHBOARD
# ==========================================
# A. Manejar el banner de cookies
try:
    print("Buscando banner de cookies...")
    btn_cookies = wait.until(EC.element_to_be_clickable((By.ID, "onetrust-accept-btn-handler")))
    btn_cookies.click()
    print("✅ Cookies aceptadas.")
    time.sleep(1) 
except TimeoutException:
    print("ℹ️ No apareció el aviso de cookies.")

# B. Entrar al iFrame
print("Buscando el iFrame del dashboard...")
iframe = wait.until(EC.presence_of_element_located((By.TAG_NAME, "iframe")))
driver.switch_to.frame(iframe)

print("Esperando a que carguen los datos iniciales...")
time.sleep(2) # Pausa para asegurar que la tabla por defecto renderice bien

# NUEVO PASO: Seleccionar la pestaña "Nacional - Puestos de Trabajo"
print("Cambiando a la pestaña 'Nacional - Puestos de Trabajo'...")
xpath_pestana = "//div[contains(@class, 'tabStoryPointContent') and contains(normalize-space(), 'Nacional - Puestos de Trabajo')]"
tab_puestos = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_pestana)))
tab_puestos.click()
time.sleep(5) # Pausa un poco más larga para dejar que los datos de la pestaña rendericen

# ==========================================
# 3. EXTRACCIÓN DIRECTA
# ==========================================
try:
    print("Iniciando secuencia de descarga...")
    xpath_btn_descarga = "//button[@data-tb-test-id='viz-viewer-toolbar-button-download']"
    btn_descarga = wait.until(EC.presence_of_element_located((By.XPATH, xpath_btn_descarga)))
    driver.execute_script("arguments[0].click();", btn_descarga)
    time.sleep(1) 

    xpath_crosstab = "//div[@data-tb-test-id='download-flyout-download-crosstab-MenuItem']"
    btn_crosstab = wait.until(EC.presence_of_element_located((By.XPATH, xpath_crosstab)))
    driver.execute_script("arguments[0].click();", btn_crosstab)
    time.sleep(2)
    
    # Nota: Si Tableau pide seleccionar qué hoja descargar antes de elegir CSV, 
    # tendremos que agregar un clic extra aquí. Por ahora, asumimos flujo directo.

    print("Seleccionando formato CSV...")
    xpath_csv_label = "//label[@data-tb-test-id='crosstab-options-dialog-radio-csv-Label']"
    btn_csv = wait.until(EC.presence_of_element_located((By.XPATH, xpath_csv_label)))
    driver.execute_script("arguments[0].click();", btn_csv)
    time.sleep(1) 

    xpath_descarga_final = "//button[@data-tb-test-id='export-crosstab-export-Button']"
    btn_descarga_final = wait.until(EC.presence_of_element_located((By.XPATH, xpath_descarga_final)))
    driver.execute_script("arguments[0].click();", btn_descarga_final)

    print("📥 Descargando archivo maestro...")
    time.sleep(5) 
    print("✅ Descarga completada.")

except Exception as e:
    print(f"❌ Error durante la descarga: {e}")

# ==========================================
# 4. PROCESAMIENTO Y LIMPIEZA CON PANDAS
# ==========================================
print("🔄 Procesando y limpiando el archivo descargado...")
# 1. Búsqueda quirúrgica del archivo descargado
prefijo_descarga = "Nacional" 

patron_busqueda = os.path.join(raw_dir, f"{prefijo_descarga}*.csv")
archivos_candidatos = glob.glob(patron_busqueda)

# Filtrar SOLO los creados/modificados DESPUÉS de que inició el script
archivos_validos = [f for f in archivos_candidatos if os.path.getmtime(f) >= tiempo_inicio_script]

if archivos_validos:
    archivo_reciente = max(archivos_validos, key=os.path.getmtime)
    
    # Leer el archivo (Tableau usa utf-16 y tabulaciones), saltando la fila 1
    df = pd.read_csv(
        archivo_reciente, 
        skiprows=1, 
        encoding='utf-16', 
        sep='\t'
    )
    
    # Renombrar las dos primeras columnas dinámicamente (índices 0 y 1)
    nuevos_nombres = {
        df.columns[0]: 'Año',
        df.columns[1]: 'Mes'
    }
    df.rename(columns=nuevos_nombres, inplace=True)
    
    # Identificar columnas de estados (todas a partir de la tercera)
    columnas_estados = df.columns[2:]
    
    # Limpiar el formato de miles y asegurar tipo numérico para sumar correctamente
    for col in columnas_estados:
        df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
        
    # Crear la columna Total y conservar solo las columnas necesarias
    df['Total'] = df[columnas_estados].sum(axis=1).astype(int)
    df = df[['Año', 'Mes', 'Total']]
    
    # Eliminar el archivo temporal raw
    try:
        os.remove(archivo_reciente)
        print("🗑️ Archivo temporal eliminado de la carpeta raw.")
    except Exception as e:
        print(f"⚠️ No se pudo eliminar el archivo temporal: {e}")
        
    # Guardar en intermediate
    print("\n💾 Guardando el consolidado en la carpeta intermediate...")
    # Le cambié el nombre a imss_empleo_nacional.xlsx para diferenciarlo del de salarios
    ruta_salida = os.path.join(intermediate_dir, "puestos_imss.xlsx") 
    
    df.to_excel(ruta_salida, index=False)
    print(f"✅ ¡Éxito! Archivo maestro guardado en:\n{ruta_salida}")

else:
    print("⚠️ No se detectó ningún archivo CSV descargado. Revisa el flujo de Selenium.")

# Cerrar el navegador
print("\nCerrando el navegador...")
driver.quit()