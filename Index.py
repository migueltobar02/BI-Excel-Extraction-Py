#Este codigo es con Google Colab

# Commented out IPython magic to ensure Python compatibility.
# %reset -f

import pandas as pd
from tabulate import tabulate
from pathlib import Path
from datetime import datetime
import re
import os

#obtener los archivos
from google.colab import drive
drive.mount('/content/drive')
CARPETA_ORIGEN = "/content/drive/MyDrive/PDF_laboratorios/"
ARCHIVO_CONTROL = os.path.join(CARPETA_ORIGEN, "control_procesados.xlsx")

# Función para limpiar valores
def limpiar_valores(valor):
    if pd.isna(valor): return ""
    valor_str = str(valor).strip().replace('↑','').replace('↓','')
    if re.match(r'^\d*\.\d+$', valor_str): return valor_str.replace('.', ',')
    if re.match(r'^\d+\.\d+\s*-\s*\d+\.\d+$', valor_str): return valor_str.replace('.', ',')
    if re.match(r'^\d+\.\d+\s*-\s*\d+\.\d+\s+[a-zA-Z]', valor_str):
        partes = valor_str.split()
        return f"{partes[0].replace('.',',')} - {partes[2].replace('.',',')} {partes[3]}"
    return valor_str

def obtener_tipo_examen(df):
    """Extrae el tipo de examen que aparece después de 'Fecha elaboración informe: DD MMº AAAA'"""
    try:
        # Patrón regex para encontrar la fecha de elaboración y lo que sigue
        patron = r"Fecha elaboración informe:\s*[A-Za-z]+\s*\d+º\s*\d{4}\s*(.*)"

        for i in range(len(df)):
            celda = str(df.iloc[i, 0])
            coincidencia = re.search(patron, celda, re.IGNORECASE)

            if coincidencia:
                examen = coincidencia.group(1).strip()

                # Limpieza adicional para eliminar posibles saltos de línea o espacios múltiples
                examen = ' '.join(examen.split())

                # Si hay contenido después de la fecha, devolver la primera línea no vacía
                if examen:
                    # Tomar solo hasta el primer salto de línea si hay múltiples líneas
                    return examen.split('\n')[0].strip()

        # Plan B: Buscar en la celda inferior si no se encontró en la misma celda
        for i in range(len(df)):
            if "Fecha elaboración informe" in str(df.iloc[i, 0]) and i+1 < len(df):
                examen = str(df.iloc[i+1, 0]).strip()
                if examen and not any(x in examen for x in ["ANÁLISIS", "TÉCNICA"]):
                    return examen

        return "Tipo de examen no identificado"

    except Exception as e:
        print(f"Error al obtener tipo de examen: {str(e)}")
        return "Tipo de examen no identificado"

def procesar_archivo(filepath):
    try:
        df = pd.read_excel(filepath, sheet_name='page 1')
        nombre_archivo = Path(filepath).name
        tipo_examen = obtener_tipo_examen(df)
        datos_paciente = {
            "Tipo Examen": tipo_examen,
            "Nombre": df.iloc[4, 7],
            "Especie": df.iloc[5, 7],
            "Raza": df.iloc[6, 7],
            "Sexo": df.iloc[4, 13],
            "Edad": df.iloc[5, 13],
            "Propietario": df.iloc[6, 13].replace("io:", "").strip(),
            "Veterinario": df.iloc[4, 0].replace("Médico veterinario:", "").strip(),
            "Fecha Muestreo": df.iloc[6, 2],
            "Código Informe": Path(filepath).stem

        }

        def buscar_seccion(texto):
            matches = df[df.iloc[:, 0].astype(str).str.contains(texto, case=False, na=False)]
            return matches.index[0] if not matches.empty else len(df)

        secciones = {
            "Analisis": buscar_seccion("ANÁLISIS"),
            "FIN": buscar_seccion("ah = Análisis habilitado")
        }

        resultados = []
        for seccion, inicio in secciones.items():
            if seccion == "FIN": continue
            fin = secciones.get(list(secciones.keys())[list(secciones.keys()).index(seccion)+1], secciones["FIN"])

            for i in range(inicio + 1, fin):
                fila = df.iloc[i]
                prueba = str(fila.iloc[0]).strip()
                if prueba and not any(s in prueba for s in ["SERIE", "RECUENTO", "PLAQUETAS", "ANÁLISIS"]):
                    resultados.append({
                        **datos_paciente,
                        "Prueba": prueba,
                        "Resultado": limpiar_valores(fila.iloc[8]),
                        "Unidades": limpiar_valores(fila.iloc[10]),
                        "Valor Referencia": limpiar_valores(fila.iloc[12])
                    })

        return resultados, nombre_archivo
    except Exception as e:
        print(f"ERROR procesando {Path(filepath).name}: {str(e)}")
        return [], None

def procesar_carpeta_control(carpeta_origen, archivo_control):
    ARCHIVOS_EXCLUIDOS = ['control_procesados.xlsx', 'resultados_consolidados.xlsx']

    # Verificar y crear carpeta si no existe
    os.makedirs(os.path.dirname(archivo_control), exist_ok=True)

    # Cargar o crear archivo de control
    try:
        if os.path.exists(archivo_control):
            df_control = pd.read_excel(archivo_control)
            archivos_registrados = set(df_control['Archivo'].tolist())
        else:
            archivos_registrados = set()
            df_control = pd.DataFrame(columns=['Archivo', 'Fecha_Registro'])
    except Exception as e:
        print(f"Error al leer archivo de control: {str(e)}")
        archivos_registrados = set()
        df_control = pd.DataFrame(columns=['Archivo', 'Fecha_Registro'])

    nuevos_archivos = []
    todos_resultados = []  # Cambiamos a lista para mantener el formato original

    # Buscar archivos Excel en la carpeta
    for archivo in os.listdir(carpeta_origen):
        # Verificar extensión y que no esté en excluidos
        if (archivo.lower().endswith(('.xlsx', '.xls'))
            and archivo not in archivos_registrados
            and archivo not in ARCHIVOS_EXCLUIDOS):

            filepath = os.path.join(carpeta_origen, archivo)

            print(f"\nProcesando archivo: {archivo}")

            try:
                # Procesar el archivo (igual que en el test individual)
                resultados, _ = procesar_archivo(filepath)

                if resultados:
                    todos_resultados.extend(resultados)  # Usamos extend en lugar de concat
                    nuevos_archivos.append({
                        'Archivo': archivo,
                        'Fecha_Registro': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    })
                    print(f"Procesado correctamente: {archivo} ({len(resultados)} registros)")

                    # Mostrar los datos del archivo actual (igual que en el test)
                    print("\nDatos extraídos de este archivo:")
                    display(pd.DataFrame(resultados))

                else:
                    print(f"Archivo {archivo} no produjo resultados")

            except Exception as e:
                print(f"Error al procesar {archivo}: {str(e)}")

    # Actualizar archivo de control si hay nuevos archivos
    if nuevos_archivos:
        df_nuevos = pd.DataFrame(nuevos_archivos)
        df_control = pd.concat([df_control, df_nuevos], ignore_index=True)

        try:
            with pd.ExcelWriter(archivo_control, engine='openpyxl') as writer:
                df_control.to_excel(writer, index=False)
            print(f"\nArchivo de control actualizado: {archivo_control}")
        except Exception as e:
            print(f"Error al guardar archivo de control: {str(e)}")

    # Convertimos a DataFrame al final para mantener compatibilidad
    return pd.DataFrame(todos_resultados) if todos_resultados else pd.DataFrame()

# ----------------------------
# POST-PROCESAMIENTO (NUEVAS FUNCIONALIDADES)
# ----------------------------

def limpiar_espacios(texto):
    """Elimina espacios innecesarios dejando solo un espacio entre palabras"""
    if pd.isna(texto):
        return ""
    return ' '.join(str(texto).strip().split())

def procesar_edad(df_resultados):
    """Procesa la columna Edad para separar en Años y Meses"""
    # Primero limpiamos la columna Edad
    df_resultados['Edad'] = df_resultados['Edad'].apply(limpiar_espacios)

    # Inicializamos las nuevas columnas
    df_resultados['Años'] = 0
    df_resultados['Meses'] = 0

    # Procesamos cada registro
    for idx, row in df_resultados.iterrows():
        edad = row['Edad'].lower()

        if 'año' in edad:
            años = re.search(r'\d+', edad)
            if años:
                df_resultados.at[idx, 'Años'] = int(años.group())
                # Convertir años a meses (1 año = 12 meses)
                df_resultados.at[idx, 'Meses'] = int(años.group()) * 12

        elif 'mes' in edad:
            meses = re.search(r'\d+', edad)
            if meses:
                df_resultados.at[idx, 'Meses'] = int(meses.group())

    return df_resultados

def limpiar_todos_los_textos(df_resultados):
    """Aplica limpieza de espacios a todas las columnas de texto"""
    for col in df_resultados.columns:
        if df_resultados[col].dtype == 'object':  # Solo para columnas de texto
            df_resultados[col] = df_resultados[col].apply(limpiar_espacios)
    return df_resultados

# 1. Procesar toda la carpeta (código original)
resultados_finales = procesar_carpeta_control(CARPETA_ORIGEN, ARCHIVO_CONTROL)

# 2. Aplicar post-procesamiento si hay resultados
if not resultados_finales.empty:
    # Aplicamos limpieza de textos
    resultados_finales = limpiar_todos_los_textos(resultados_finales)

    # Procesamos la edad
    resultados_finales = procesar_edad(resultados_finales)

    # Reordenar columnas para mejor visualización
    columnas = resultados_finales.columns.tolist()
    columnas.remove('Años')
    columnas.remove('Meses')
    idx_edad = columnas.index('Edad')
    columnas.insert(idx_edad + 1, 'Años')
    columnas.insert(idx_edad + 2, 'Meses')
    resultados_finales = resultados_finales[columnas]

    # Guardar resultados mejorados
    ruta_resultados = os.path.join(CARPETA_ORIGEN, 'resultados_consolidados.xlsx')
    resultados_finales.to_excel(ruta_resultados, index=False)
    print(f"\nResultados mejorados guardados en: {ruta_resultados}")

    # Mostrar resultados finales
    print("\nRESULTADOS CONSOLIDADOS MEJORADOS:")
    display(resultados_finales)
else:
    print("\nNo se encontraron archivos nuevos para procesar")
