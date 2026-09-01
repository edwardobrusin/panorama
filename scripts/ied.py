import os
import pandas as pd

# ==========================================
# 1. CONFIGURACIÓN DEL ARCHIVO
# ==========================================
# Modifica el nombre del archivo aquí cuando cambie el trimestre/año
NOMBRE_ARCHIVO_IED = "2026_2T_Flujos_TI_OR_2.xlsx"

# ==========================================
# 2. CONFIGURACIÓN DE RUTAS DINÁMICAS
# ==========================================
try:
    PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
    if os.path.basename(PROJECT_ROOT) == 'scripts':
        PROJECT_ROOT = os.path.dirname(PROJECT_ROOT)
except NameError:
    PROJECT_ROOT = os.getcwd()
    if os.path.basename(PROJECT_ROOT) == 'scripts':
        PROJECT_ROOT = os.path.dirname(PROJECT_ROOT)

RAW_DIR = os.path.join(PROJECT_ROOT, "data", "raw")
INTERMEDIATE_DIR = os.path.join(PROJECT_ROOT, "data", "intermediate")

# Asegurar que el directorio de salida exista
os.makedirs(INTERMEDIATE_DIR, exist_ok=True)

# Rutas completas
archivo_entrada = os.path.join(RAW_DIR, NOMBRE_ARCHIVO_IED)
archivo_salida = os.path.join(INTERMEDIATE_DIR, "ied.xlsx")

# ==========================================
# 3. LÓGICA DE PROCESAMIENTO
# ==========================================
def procesar_ied_simplificado():
    print(f"⏳ Iniciando procesamiento de: {NOMBRE_ARCHIVO_IED}")
    
    if not os.path.exists(archivo_entrada):
        print(f"❌ Error: No se encontró el archivo en la ruta:\n{archivo_entrada}")
        return

    try:
        # Leer la primera hoja del libro, sin asumir que la fila 0 son los encabezados
        df = pd.read_excel(archivo_entrada, sheet_name=0, header=None)
        
        fila_anios = None
        fila_trimestres = None
        fila_total = None
        
        # 3.1 Identificar dinámicamente en qué fila está cada cosa
        for idx, row in df.iterrows():
            texto_celda = str(row[0]).strip().lower()
            
            if "tipo de inversión" in texto_celda:
                fila_anios = idx
                fila_trimestres = idx + 1 # Los trimestres siempre están justo abajo de los años
            elif "total general" in texto_celda:
                fila_total = idx
                break # Una vez encontramos el total general, no necesitamos buscar más abajo
                
        if fila_total is None or fila_anios is None:
            print("❌ Error: No se pudo identificar la estructura (años, trimestres o total general).")
            return

        datos_limpios = []
        anio_actual = None
        
        # 3.2 Extraer los datos iterando por las columnas (de izquierda a derecha)
        for col in range(1, len(df.columns)):
            val_anio = str(df.iloc[fila_anios, col]).strip()
            val_trimestre = str(df.iloc[fila_trimestres, col]).strip()
            val_total = df.iloc[fila_total, col]
            
            # Si la celda del año tiene un número de 4 dígitos, actualizamos nuestro "anio_actual"
            # (Limpiamos los ".0" en caso de que pandas lo lea como float)
            val_anio_limpio = val_anio.replace('.0', '')
            if val_anio_limpio.isdigit() and len(val_anio_limpio) == 4:
                anio_actual = int(val_anio_limpio)
                
            # Si ya tenemos un año registrado y estamos en una columna de trimestre válido (1, 2, 3 o 4)
            val_trimestre_limpio = val_trimestre.replace('.0', '')
            if anio_actual and val_trimestre_limpio in ['1', '2', '3', '4']:
                trimestre = int(val_trimestre_limpio)
                
                # Limpiar el valor total (quitar comas y convertir a float)
                try:
                    total_numerico = float(str(val_total).replace(',', '').strip())
                except ValueError:
                    total_numerico = 0.0
                    
                # Guardar el registro
                datos_limpios.append({
                    'Año': anio_actual,
                    'Trimestre': trimestre,
                    'Total': total_numerico
                })
                
        # 3.3 Crear el DataFrame final y exportar
        df_final = pd.DataFrame(datos_limpios)
        
        if df_final.empty:
            print("⚠️ Advertencia: El proceso finalizó pero no se extrajeron datos.")
        else:
            df_final.to_excel(archivo_salida, index=False)
            print(f"✅ ¡Éxito! Archivo procesado y exportado con {len(df_final)} registros.")
            print(f"📂 Ubicación: {archivo_salida}")

    except Exception as e:
        print(f"❌ Ocurrió un error inesperado: {e}")

if __name__ == "__main__":
    procesar_ied_simplificado()