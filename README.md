# BI-Excel-Extraction-Py

## 📊 BI_Excel_Extraction_Py:

El repositorio contiene el código desarrollado en Python para la extracción, limpieza y estructuración inicial de datos de laboratorio clínico veterinario almacenados en archivos Excel. Fue diseñado con el propósito de facilitar el procesamiento preliminar y la extracción de los datos para su posterior análisis, formando parte de una solución de Inteligencia de Negocios (BI). Este código se integra en la etapa de Extracción, Transformación y Carga (ETL), constituyendo el primer paso del proceso de extracción, cuyo resultado es utilizado por la herramienta Pentaho para, posteriormente, permitir la visualización de la información en Power BI.

🛠️ Tecnologías utilizadas:

* 💻 Google Colab: entorno para ejecución y desarrollo del código

* 🐍 Python 3.x: lenguaje de programación utilizado

* 🐼 pandas: librería para manipulación y análisis de datos

* 📦 openpyxl: lectura y escritura de archivos Excel

* 🔍 re: uso de expresiones regulares para búsqueda y limpieza de texto

* 📊 tabulate: formato y presentación tabular en consola

* 📁 pathlib, os, datetime: gestión eficiente de archivos y manejo de fechas

## 📝 Descripción del proceso:

El código automatiza la lectura y procesamiento de múltiples archivos Excel que contienen resultados de laboratorios clínicos veterinarios, los cuales han sido previamente convertidos desde archivos PDF. Los datos extraídos comprenden información de la mascota (nombre, especie, sexo, edad), propietario, veterinario, tipo de examen, resultados de pruebas, unidades y rangos de referencia. El proceso incluye la limpieza de datos textuales, la corrección de formatos numéricos y de fechas, así como el cálculo de variables derivadas, tales como la edad expresada en meses. Finalmente, toda la información se consolida en un único archivo Excel que sirve como insumo para la etapa posterior de ETL en Pentaho.

## 🚀 Instrucciones para la ejecución
1. Clonar el repositorio desde:
  https://github.com/migueltobar02/BI-Excel-Extraction-Py.git

2. Instalar las dependencias necesarias.

3. Crear el archivo de control "control_procesados.xlsx" para gestionar los archivos ya procesados.

4. Configurar la carpeta donde se almacenarán los archivos Excel a procesar.

5. Ejecutar el script principal para iniciar el procesamiento.

6. Revisar los resultados consolidados, los cuales se almacenarán en el archivo "resultados_consolidados.xlsx."
