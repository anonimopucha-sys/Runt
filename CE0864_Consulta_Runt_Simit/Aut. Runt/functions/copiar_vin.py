import pandas as pd

# Cargar el archivo de Excel
def copiar_vin():
    try:
        df = pd.read_excel('C:\\CE0864_Consulta_RUNT_SIMIT\\Consulta_RUNT_VIM.xlsx')
        # Extraer la columna A (ignorando la primera fila)
        columna_a = df.iloc[0:, 0]
        # Recorrer la columna e imprimir cada celda hasta encontrar una vacía
        vins = []
        for i, celda in enumerate(columna_a):
            if pd.isna(celda):
                break
            vins.append(celda)
            print(f"{i+1}. {celda}")
        return vins
    except Exception as e:
        print("Error al ejecutar la funcion copiar vin",e)
        return []