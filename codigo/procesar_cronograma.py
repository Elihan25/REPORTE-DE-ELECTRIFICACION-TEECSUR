import openpyxl
import json
import os

# Configuración de ruta y celdas
ruta_excel = r"C:\Users\EPOMAP\Downloads\CURVA S - LDS - SANTA CLARA.xlsx"
celdas_semanas = [
    "O164", "X164", "AG164", "AP164", "AY164", 
    "BH164", "BQ164", "BZ164", "CL164", "CR164", "CX164"
]

def extraer_datos():
    if not os.path.exists(ruta_excel):
        print(f"Error: No se encontró el archivo en {ruta_excel}")
        return

    # Cargar el libro de trabajo (data_only=True para obtener el valor de la fórmula)
    wb = openpyxl.load_workbook(ruta_excel, data_only=True)
    
    hoja_proy = wb["AVANCE PROYECTADO"]
    hoja_real = wb["AVANCE REAL"]

    datos_finales = []
    acum_proy_anterior = 0
    acum_real_anterior = 0

    print("Procesando semanas...")

    for i, celda in enumerate(celdas_semanas):
        # Obtener valores acumulados de las celdas
        # Multiplicamos por 100 para tener formato porcentaje (0.10 -> 10.0)
        v_proy_acu = (hoja_proy[celda].value or 0) * 100
        v_real_acu = (hoja_real[celda].value or 0) * 100

        # Calcular el valor semanal (la diferencia con la semana anterior)
        v_proy_sem = v_proy_acu - acum_proy_anterior
        v_real_sem = v_real_acu - acum_real_anterior
        
        # Calcular cumplimiento semanal
        cumplimiento = (v_real_sem / v_proy_sem * 100) if v_proy_sem > 0 else 0

        datos_finales.append({
            "semana": i + 1,
            "pSem": round(v_proy_sem, 2),
            "pAcu": round(v_proy_acu, 2),
            "rSem": round(v_real_sem, 2),
            "rAcu": round(v_real_acu, 2),
            "cump": round(cumplimiento, 2)
        })

        # Actualizar acumulados para el siguiente cálculo semanal
        acum_proy_anterior = v_proy_acu
        acum_real_anterior = v_real_acu

    # Guardar en archivo JS para el HTML
    with open("datos_proyecto.js", "w") as f:
        f.write(f"const datosProyecto = {json.dumps(datos_finales, indent=4)};")
    
    print("✅ ¡Éxito! Archivo 'datos_proyecto.js' generado.")

if __name__ == "__main__":
    extraer_datos()