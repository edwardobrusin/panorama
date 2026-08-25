import os
import time
import requests
import pandas as pd
from tqdm.auto import tqdm

# ==========================================
# CONFIGURACIÓN ESPECÍFICA (MODIFICAR AQUÍ)
# ==========================================
NOMBRE_PROCESO = "Ingresos Remesas (Desest.)"
ARCHIVO_SALIDA = "remesas.xlsx"

INDICADORES = {
    "SE44962": "Ingresos"
}

# ==========================================
# CONFIGURACIÓN GLOBAL Y RUTAS
# ==========================================
TOKEN_BANXICO = "3cf05ba180ebc8fa6bac83d6473f5c287fd4a5d28c8d0411ec9c2e896b844b3e"

try:
    SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
    if os.path.basename(SCRIPT_DIR) == 'scripts':
        PROJECT_ROOT = os.path.dirname(SCRIPT_DIR)
    else:
        PROJECT_ROOT = SCRIPT_DIR
except NameError:
    PROJECT_ROOT = os.getcwd()
    if os.path.basename(PROJECT_ROOT) == 'scripts':
        PROJECT_ROOT = os.path.dirname(PROJECT_ROOT)

INTERMEDIATE_DIR = os.path.join(PROJECT_ROOT, "data", "intermediate")
os.makedirs(INTERMEDIATE_DIR, exist_ok=True)

# ==========================================
# LÓGICA DE EXTRACCIÓN
# ==========================================
def obtener_serie_banxico(token, serie_id):
    url = f"https://www.banxico.org.mx/SieAPIRest/service/v1/series/{serie_id}/datos"
    headers = {"Bmx-Token": token}
    
    for intento in range(3):
        try:
            response = requests.get(url, headers=headers, timeout=10)
            if response.status_code == 200:
                data = response.json()
                serie_info = data['bmx']['series'][0]
                titulo = serie_info['titulo']
                datos = serie_info['datos']
                
                df = pd.DataFrame(datos)
                df.columns = ['fecha', titulo]
                df['fecha'] = pd.to_datetime(df['fecha'], format='%d/%m/%Y')
                df[titulo] = pd.to_numeric(df[titulo].str.replace(',', ''), errors='coerce')
                return df
            else:
                time.sleep(1)
        except requests.exceptions.RequestException:
            time.sleep(1)
            
    return pd.DataFrame(columns=['fecha', 'valor_nulo'])

def procesar_indicador():
    print(f"⏳ [{NOMBRE_PROCESO}] Iniciando extracción...")
    
    df_master = pd.DataFrame()
    errores = 0
    
    barra_progreso = tqdm(INDICADORES.items(), desc=f"📊 {NOMBRE_PROCESO[:15]:<15}", unit="serie")
    
    for serie_id, nombre_columna in barra_progreso:
        df_temp = obtener_serie_banxico(TOKEN_BANXICO, serie_id)
        
        if not df_temp.empty and len(df_temp.columns) == 2:
            df_temp.columns = ['fecha', nombre_columna]
            if df_master.empty:
                df_master = df_temp
            else:
                # Merge 'outer' en caso de que se agreguen más indicadores con distintas fechas
                df_master = df_master.merge(df_temp, on="fecha", how="outer")
        else:
            errores += 1
            barra_progreso.set_postfix({"Errores": errores})
            
    if not df_master.empty:
        df_master = df_master.sort_values(by='fecha').reset_index(drop=True)
        df_master['fecha'] = df_master['fecha'].dt.strftime('%Y-%m-%d')
        
        ruta_salida = os.path.join(INTERMEDIATE_DIR, ARCHIVO_SALIDA)
        df_master.to_excel(ruta_salida, index=False)
        
        print(f"✅ [{NOMBRE_PROCESO}] Completado ({len(df_master)} periodos). Guardado en {ruta_salida}")
    else:
        print(f"⚠️ [{NOMBRE_PROCESO}] No se pudieron extraer los datos.")

if __name__ == "__main__":
    procesar_indicador()
