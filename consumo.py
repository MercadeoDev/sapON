import pandas as pd

path = r'C:\Users\felipe.moncada\Downloads\121515.xlsb'

# Leemos la hoja completa
# Supongamos que tu tabla tabl_prof:
# - Tiene cabecera en la fila 10 del Excel (índice 9 en pandas)
# - Ocupa las columnas A, B, C y D (es decir, usecols="A:D")
#se consume el pedacito porque es muy lento todo a la vez
df_conSQL_pag = pd.read_excel(path,
    sheet_name='Consultas SQL',
    engine='pyxlsb',
    header=1, #primera fila es header
    usecols="A:F", nrows=40) #indicativo columnas 

print(df_conSQL_pag)
