import random
import time
import requests
import pandas as pd
import concurrent.futures
from tqdm.auto import tqdm
import os

# ==========================================
# CONFIGURACIÓN ESPECÍFICA (MODIFICAR AQUÍ)
# ==========================================
NOMBRE_PROCESO = "INPC Mensual"
ARCHIVO_SALIDA = "inpc_mensual.xlsx"

INDICADORES = {
    "910399": "Var. Mensual",
    "910406": "Var. Anual"
}

# ==========================================
# CONFIGURACIÓN GENERAL Y CREDENCIALES
# ==========================================
TOKENS_INEGI = [
    "129ac2e3-e8a6-72c7-58c1-acced5a601bd",
    "8ff1ca1f-4ba0-4abc-b98e-c8a5dffc9467",
    "9d9b582f-0cc1-6e57-9d97-2064cebd95d9",
    "8505df05-4276-f1b8-3c6a-437fe9d77c7a",
]

def obtener_token():
    return random.choice(TOKENS_INEGI)

# ==========================================
# LÓGICA DE EXTRACCIÓN
# ==========================================
def hacer_peticion(tarea):
    ind_clave, ind_nombre = tarea
    clave_estado = "00"  # Nacional por defecto
    token_actual = obtener_token()
    url = f"https://www.inegi.org.mx/app/api/indicadores/desarrolladores/jsonxml/INDICATOR/{ind_clave}/es/{clave_estado}/false/BIE-BISE/2.0/{token_actual}?type=json"
    
    errores_locales = 0
    for intento in range(3):
        try:
            r = requests.get(url, timeout=10)
            if r.status_code == 200:
                try:
                    data = r.json()
                    res_locales = []
                    if 'Series' in data and data['Series']:
                        serie = data['Series'][0].get('OBSERVATIONS', [])
                        serie_sorted = sorted(serie, key=lambda x: x.get('TIME_PERIOD', ''))
                        for obs in serie_sorted:
                            raw_val = obs.get('OBS_VALUE')
                            val_limpio = 0.0
                            
                            if raw_val is not None and str(raw_val).strip():
                                try:
                                    val_limpio = float(str(raw_val).strip())
                                except ValueError:
                                    val_limpio = 0.0
                                    
                            raw_periodo = obs.get('TIME_PERIOD', '0')
                            
                            # Intentamos castear el periodo si es estrictamente numérico
                            try:
                                periodo_limpio = int(raw_periodo)
                            except ValueError:
                                periodo_limpio = raw_periodo

                            res_locales.append({
                                'Indicador': ind_nombre, 
                                'Clave_Indicador': ind_clave,
                                'Periodo': periodo_limpio, 
                                'Valor': val_limpio
                            })
                    return res_locales, errores_locales, "EXITO"
                except Exception:
                    return [], errores_locales, "NO_DATA"
            elif r.status_code == 429 or r.status_code >= 500:
                errores_locales += 1
                time.sleep(1 + random.uniform(0.1, 1.5))
            else:
                return [], errores_locales, "NO_DATA"
        except requests.exceptions.RequestException:
            errores_locales += 1
            time.sleep(1 + random.uniform(0.1, 1.5))
    return [], errores_locales, "RETRY"

def procesar_indicadores():
    print(f"⏳ [{NOMBRE_PROCESO}] Iniciando extracción (Nacional)...")
    
    resultados = []
    tareas = [(clave, nombre) for clave, nombre in INDICADORES.items()]
    tareas_pendientes = tareas.copy()
    errores_totales = 0

    with concurrent.futures.ThreadPoolExecutor(max_workers=len(TOKENS_INEGI) * 2) as executor:
        barra_progreso = tqdm(
            total=len(tareas_pendientes), 
            desc=f"📊 {NOMBRE_PROCESO} ", 
            unit="req",
            position=0,
            leave=True
        )
        
        while tareas_pendientes:
            futuros = {executor.submit(hacer_peticion, t): t for t in tareas_pendientes}
            tareas_pendientes = []
            
            for futuro in concurrent.futures.as_completed(futuros):
                t = futuros[futuro]
                res, errs, estado = futuro.result()
                
                if res: resultados.extend(res)
                errores_totales += errs
                
                if estado == "RETRY":
                    tareas_pendientes.append(t)
                else:
                    barra_progreso.update(1)
                    
                if errores_totales > 0:
                    barra_progreso.set_postfix({"Errores de Red": errores_totales})
                    
        barra_progreso.close()

    if resultados:
        df = pd.DataFrame(resultados)
        
        ruta_directorio = os.path.join("data", "intermediate")
        os.makedirs(ruta_directorio, exist_ok=True)
        ruta_salida = os.path.join(ruta_directorio, ARCHIVO_SALIDA)
        
        df.to_excel(ruta_salida, index=False)
        print(f"✅ [{NOMBRE_PROCESO}] Completado ({len(df)} registros guardados en {ruta_salida}).")
    else:
        print(f"⚠️ [{NOMBRE_PROCESO}] No se obtuvieron datos.")

if __name__ == "__main__":
    procesar_indicadores()
