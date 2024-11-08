import os
import pandas as pd
from tkinter import filedialog, messagebox

def procesar_archivos(archivos, fecha_inicio, fecha_fin, progress, root):
    columnas_requeridas = [
        "Cuenta", "Región", "Zona", "SIREC_SEGMENTO", "FECHA_OBSERVACION_NO_TRABAJADA",
        "OBSERVACIONES_NO_TRABAJADAS", "CONS_EAT_M3", "TOTAL", "Q_PESO_OBSERVACIONES", 
        "Q_BIMESTRAL_PESO", "Q_BAJO_PESO", "Q_PROMEDIO_1_PESO", "Q_PROMEDIO_2_PESO", 
        "Q_ANUAL_PESO", "Q_PESO_SIREC", "Q_PESO_%_DESVIO", "Q_PESO_KWH_DESVIO", 
        "CONTRATISTA_DIME", "FECHA_PEOR_SCORING_HISTORICO_SIN_TRABAJAR", 
        "SCORING_PEOR_HISTORICO_SIN_TRABAJAR", "DESVIO_BIMESTRAL", "DESVIO_ANUAL", 
        "SEGMENTACION", "FEC_TRABAJO", "RESULTADO_AC"
    ]

    df_estandarizado = pd.DataFrame(columns=columnas_requeridas)

    for i, ruta_archivo in enumerate(archivos):
        try:
            hoja = pd.read_excel(ruta_archivo, sheet_name="ANALISIS")
            hoja.columns = hoja.columns.str.strip().str.upper()  # Normalizar nombres de columnas
            columnas_existentes = [col for col in columnas_requeridas if col in hoja.columns]
            df_filtrado = hoja[columnas_existentes]
            
            df_filtrado["Nombre_Archivo"] = os.path.basename(ruta_archivo)
            df_estandarizado = pd.concat([df_estandarizado, df_filtrado], ignore_index=True)
        except Exception as e:
            messagebox.showerror("Error", f"Error al leer el archivo {ruta_archivo}: {e}")
        
        progress["value"] = (i + 1) / len(archivos) * 100
        root.update_idletasks()

    df_estandarizado.drop_duplicates(subset="Cuenta", keep="first", inplace=True)

    # Solicitar archivo REP 10 y realizar el merge
    ruta_rep10 = filedialog.askopenfilename(title="Seleccionar archivo REP 10", filetypes=[("Archivos de Excel", "*.xlsx")])
    if ruta_rep10:
        try:
            df_rep10 = pd.read_excel(ruta_rep10, sheet_name="R123")
            df_rep10.columns = df_rep10.columns.str.strip().str.upper()  # Normalizar nombres de columnas
            
            # Verificar si las columnas necesarias están presentes
            print("Columnas en REP 10 después de normalizar:", df_rep10.columns.tolist())  # <-- Verificación de columnas normalizadas
            
            # Asegurarse de que la columna 'FEC_TRABAJO' esté en el DataFrame
            if 'FEC_TRABAJO' not in df_rep10.columns:
                raise KeyError("La columna 'FEC_TRABAJO' no está en REP 10 después de la normalización")
            
            # Forzar la conversión a datetime
            df_rep10['FEC_TRABAJO'] = pd.to_datetime(df_rep10['FEC_TRABAJO'], errors='coerce')
            
            # Filtrar y procesar el archivo final
            columnas_necesarias_rep10 = ["CUENTA", "RESULTADO_AC", "FEC_TRABAJO", "INSTRUCCIONES", "MOTIVO_AC"]
            df_rep10 = df_rep10[columnas_necesarias_rep10]
            df_final = pd.merge(df_estandarizado, df_rep10, left_on="Cuenta", right_on="CUENTA", how="left")
            
            # Eliminar columnas duplicadas y renombrar para unificar nombres
            df_final.drop(columns=['FEC_TRABAJO_x', 'RESULTADO_AC_x'], inplace=True)
            df_final.rename(columns={'FEC_TRABAJO_y': 'FEC_TRABAJO', 'RESULTADO_AC_y': 'RESULTADO_AC'}, inplace=True)

            # Filtrar por fechas y condiciones especificadas
            fecha_inicio = pd.to_datetime(fecha_inicio)
            fecha_fin = pd.to_datetime(fecha_fin)
            df_final = df_final[(df_final['FEC_TRABAJO'] >= fecha_inicio) & (df_final['FEC_TRABAJO'] <= fecha_fin)]
            
            df_final = df_final[df_final['MOTIVO_AC'] == 'DIRECCIONAMIENTO']
            df_final = df_final[df_final['RESULTADO_AC'] != '#N/D']
            df_final = df_final.dropna(subset=['RESULTADO_AC'])
            
            output_path = filedialog.asksaveasfilename(title="Guardar archivo final", defaultextension=".xlsx", filetypes=[("Archivos de Excel", "*.xlsx")])
            if output_path:
                df_final.to_excel(output_path, index=False)
                messagebox.showinfo("Éxito", f"Procesamiento completo. Archivo guardado en {output_path}")
        except KeyError as e:
            messagebox.showerror("Error", f"Error al procesar REP 10: {e}")
        except Exception as e:
            messagebox.showerror("Error", f"Error al procesar REP 10: {e}")
