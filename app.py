import streamlit as st
import pandas as pd
import zipfile
import os
import tempfile

# Función para cargar el archivo de referencia
def cargar_df_referencia():
    uploaded_file = st.file_uploader("Cargar archivo de referencia (XLSX)", type=["xlsx", "xls"])
    if uploaded_file:
        df_referencia = pd.read_excel(uploaded_file)
        st.write("Archivo de referencia cargado exitosamente:")
        st.dataframe(df_referencia.head())  # Mostrar las primeras filas
        return df_referencia
    return None

# Función para cargar el archivo del cliente
def cargar_df_cliente():
    uploaded_file = st.file_uploader("Cargar archivo del cliente (XLSX)", type=["xlsx", "xls"])
    if uploaded_file:
        df_cliente = pd.read_excel(uploaded_file)
        st.write("Archivo del cliente cargado exitosamente:")
        st.dataframe(df_cliente.head())  # Mostrar las primeras filas
        return df_cliente
    return None

# Función para realizar el cruce de datos
def cruce_datos(df_referencia, df_cliente):
    df_cruce = pd.merge(df_cliente, df_referencia, on="NIF", how="left")
    df_cruce['conteo de matriculaciones'] = df_cruce.groupby('NIF')['CURSO1'].transform('count')
    st.write("Cruce de datos realizado exitosamente. Primeras filas del resultado:")
    st.dataframe(df_cruce.head())
    return df_cruce

# Filtrar alumnos no matriculados
def filtrar_alumnos_no_matriculados(df_recuento_matriculaciones):
    motivos_exclusion = []

    filtros = [
        df_recuento_matriculaciones['conteo de matriculaciones'] > 3,
        df_recuento_matriculaciones['CURSO'] == 1,
        ~df_recuento_matriculaciones['E-MAIL'].str.contains('@', na=False),
        ~df_recuento_matriculaciones['NIF'].str.match(r'^[XYZ\d]\d{7}[A-Z]$', na=False),
        ~df_recuento_matriculaciones['CIF'].isin([
            "B62504105", "B96740659", "F20032553", "B48419378", "B01277268",
            "A78538774", "B43642222", "B55531495", "B20627196", "B09065236",
            "B81958134"
        ]),
        df_recuento_matriculaciones['APELLIDO 1º'].isna(),
        df_recuento_matriculaciones['TELÉFONO'].isna()
    ]

    motivos = [
        "alumno matriculado más de 3 veces",
        "alumno apto",
        "correo incorrecto",
        "NIF incorrecto",
        "CIF incorrecto",
        "usuario debe tener por lo menos un apellido",
        "usuario debe tener por lo menos un teléfono"
    ]

    df_alumnos_no_matriculados = pd.DataFrame()
    for filtro, motivo in zip(filtros, motivos):
        excluidos = df_recuento_matriculaciones[filtro].copy()
        excluidos['razon_exclusion'] = motivo
        df_alumnos_no_matriculados = pd.concat([df_alumnos_no_matriculados, excluidos])

    df_alumnos_no_matriculados = df_alumnos_no_matriculados.drop_duplicates()
    st.write("Alumnos no matriculados generados exitosamente. Primeras filas del resultado:")
    st.dataframe(df_alumnos_no_matriculados.head())
    return df_alumnos_no_matriculados

# Generar el DataFrame limpio
def generar_df_limpio(df_recuento_matriculaciones, df_alumnos_no_matriculados):
    df_limpio = df_recuento_matriculaciones[~df_recuento_matriculaciones['NIF'].isin(df_alumnos_no_matriculados['NIF'])].copy()
    
    # Mantener las columnas TELÉFONO y E-MAIL duplicadas
    df_limpio['TELÉFONO 1'] = df_limpio['TELÉFONO']
    df_limpio['E-MAIL 1'] = df_limpio['E-MAIL']
    
    # Seleccionar las columnas en el orden requerido
    df_limpio = df_limpio[['NIF', 'NOMBRE_x', 'APELLIDO 1º', 'APELLIDO 2º', 'TELÉFONO', 'TELÉFONO 1', 'E-MAIL', 'E-MAIL 1', 'NISS', 'F. NACIMIENTO', 'SEXO', 'DISCAPACITADO', 'NIVEL DE ESTUDIOS', 'CATEGORÍA PROFESIONAL', 'GRUPO DE COTIZACIÓN', 'CIF']]
    
    # Cambiar los nombres de columnas a los nombres finales que se requieren
    df_limpio.columns = ['NIF', 'NOMBRE', 'APELLIDO 1º', 'APELLIDO 2º', 'TELÉFONO', 'TELÉFONO', 'E-MAIL', 'E-MAIL', 'NISS', 'F. NACIMIENTO', 'SEXO', 'DISCAPACITADO', 'NIVEL DE ESTUDIOS', 'CATEGORÍA PROFESIONAL', 'GRUPO DE COTIZACIÓN', 'CIF']
    
    df_limpio = df_limpio.drop_duplicates()
    st.write("DataFrame limpio generado exitosamente. Primeras filas del resultado:")
    st.dataframe(df_limpio.head())
    return df_limpio


# Función para generar el Excel
def generar_excel(df_limpio, df_alumnos_no_matriculados, es_bonificada):
    tamaño_bloque = 80 if es_bonificada == 'Bonificada' else 300
    num_filas = df_limpio.shape[0]
    num_partes = (num_filas // tamaño_bloque) + (1 if num_filas % tamaño_bloque != 0 else 0)

    # Crear archivos temporales
    with tempfile.TemporaryDirectory() as tmpdirname:
        # Dividir y guardar cada parte en un archivo Excel separado
        for i in range(num_partes):
            inicio = i * tamaño_bloque
            fin = inicio + tamaño_bloque
            df_parte = df_limpio.iloc[inicio:fin]
            df_parte.to_excel(os.path.join(tmpdirname, f'df_limpio_parte_{i+1}.xlsx'), index=False)

        # Guardar el archivo de alumnos no matriculados
        df_alumnos_no_matriculados.to_excel(os.path.join(tmpdirname, 'df_alumnos_no_matriculados.xlsx'), index=False)

        # Crear un archivo ZIP
        zip_path = os.path.join(tmpdirname, 'archivos_alumnos.zip')
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for i in range(num_partes):
                zipf.write(os.path.join(tmpdirname, f'df_limpio_parte_{i+1}.xlsx'), arcname=f'df_limpio_parte_{i+1}.xlsx')
            zipf.write(os.path.join(tmpdirname, 'df_alumnos_no_matriculados.xlsx'), arcname='df_alumnos_no_matriculados.xlsx')

        # Descargar el archivo ZIP
        with open(zip_path, "rb") as f:
            st.download_button("Descargar archivos", f.read(), "archivos_alumnos.zip", "application/zip")

# Interfaz principal
def main():
    st.title("Aplicación de Procesamiento de Datos")

    df_referencia = cargar_df_referencia()
    df_cliente = cargar_df_cliente()

    if df_referencia is not None and df_cliente is not None:
        df_recuento_matriculaciones = cruce_datos(df_referencia, df_cliente)
        df_alumnos_no_matriculados = filtrar_alumnos_no_matriculados(df_recuento_matriculaciones)

        es_bonificada = st.radio("¿La formación es bonificada o privada?", ['Bonificada', 'Privada'])

        if st.button("Generar archivos"):
            df_limpio = generar_df_limpio(df_recuento_matriculaciones, df_alumnos_no_matriculados)
            generar_excel(df_limpio, df_alumnos_no_matriculados, es_bonificada)

if __name__ == "__main__":
    main()
