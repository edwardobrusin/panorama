import tkinter as tk
from tkinter import messagebox
import os
import sys
import time
import subprocess

# Detección robusta de la raíz y carpeta de scripts
try:
    PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
except NameError:
    PROJECT_ROOT = os.getcwd()

SCRIPTS_DIR = os.path.join(PROJECT_ROOT, "scripts")

# ==========================================
# POP-UP SELECCIÓN
# ==========================================
def mostrar_interfaz_seleccion(modulos_disponibles):
    seleccion = []
    
    root = tk.Tk()
    root.title("Orquestador ETL")
    root.geometry("380x650")
    root.attributes('-topmost', True) # Mantiene la ventana siempre visible
    
    tk.Label(root, text="Selecciona los módulos a actualizar:", font=("Arial", 11, "bold")).pack(pady=10)
    
    variables_checkbox = {}
    
    # Crear un checkbox por cada módulo disponible
    for modulo in modulos_disponibles:
        var = tk.BooleanVar(value=False) # Por defecto desmarcados
        variables_checkbox[modulo] = var
        chk = tk.Checkbutton(root, text=modulo, variable=var, font=("Arial", 10))
        chk.pack(anchor='w', padx=40)
        
    def ejecutar_seleccion():
        for mod, var in variables_checkbox.items():
            if var.get():
                seleccion.append(mod)
        
        if not seleccion:
            messagebox.showwarning("Advertencia", "Debes seleccionar al menos un módulo.")
            return
            
        root.destroy()
        
    btn_iniciar = tk.Button(root, text="Iniciar ETL", command=ejecutar_seleccion, bg="#0078D7", fg="white", font=("Arial", 10, "bold"), width=15)
    btn_iniciar.pack(pady=20)
    
    root.mainloop()
    return seleccion

# ==========================================
# ORQUESTADOR
# ==========================================
def main():
    print("\n🚀 INICIANDO ORQUESTADOR ETL 🚀\n")
    print(f"📂 Directorio de scripts: {SCRIPTS_DIR}\n")
    
    # Diccionario relacionando el nombre visual con su archivo .py en /scripts
    # (Ajusta el nombre de los valores si tus archivos difieren ligeramente)
    mapa_tareas = {
        "PIB": "pib.py",
        "Exportaciones": "exportaciones.py",
        "Importaciones": "importaciones.py",
        "Consumo Privado": "consumo_privado.py",
        "IGAE": "igae.py",
        "Desocupación": "desocupacion.py",
        "IED": "ied.py",
        "Puestos IMSS": "puestos_imss.py",
        "Tasa Objetivo": "tasa_objetivo.py",
        "Tipo de Cambio FIX": "fix.py",
        "Actividad Industrial": "actividad_industrial.py",
        "FBKF": "fbkf.py",
        "INPC Mensual": "inpc_mensual.py",
        "INPC Quincenal": "inpc_quincenal.py",
        "Remesas": "remesas.py"
    }
    
    nombres_seleccionados = mostrar_interfaz_seleccion(list(mapa_tareas.keys()))
    
    if not nombres_seleccionados:
        print("⚠️ Operación cancelada. No se ejecutarán módulos.")
        return
        
    inicio_total = time.time()
    print(f"📋 Ejecutando {len(nombres_seleccionados)} módulos de forma secuencial...\n")
    
    # Ejecución secuencial para no superponer ventanas de Selenium ni mezclar las barras tqdm
    for nombre in nombres_seleccionados:
        archivo_script = mapa_tareas[nombre]
        ruta_script = os.path.join(SCRIPTS_DIR, archivo_script)
        
        print("\n" + "=" * 60)
        print(f"▶️ EJECUTANDO: {nombre}")
        print(f"📄 Archivo: {archivo_script}")
        print("=" * 60)
        
        if not os.path.exists(ruta_script):
            print(f"❌ Error: No se encontró '{archivo_script}' en la carpeta 'scripts/'.")
            continue
            
        try:
            # Usamos sys.executable para usar el mismo intérprete de Python activo
            # Sin capture_output para que los prints y barras tqdm del script hijo se muestren en vivo
            resultado = subprocess.run([sys.executable, ruta_script])
            
            if resultado.returncode != 0:
                print(f"\n⚠️ Hubo un error o interrupción al ejecutar {nombre}.")
        except Exception as e:
            print(f"\n❌ Falló la ejecución de {nombre}: {e}")
            
    print("\n" + "=" * 60)
    print(f"✨ PROCESO TOTAL TERMINADO EN {time.time()-inicio_total:.2f} SEGUNDOS ✨")
    print("=" * 60)

if __name__ == "__main__":
    main()