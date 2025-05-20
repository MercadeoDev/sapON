#Interfaz gráfica
import tkinter as tk
from tkinter import messagebox, filedialog
from tkinter.scrolledtext import ScrolledText
import customtkinter as ctk
from PIL import Image, ImageTk 

import os
import json
import sys

user_pc = os.environ["USERPROFILE"]
user_name = os.path.basename(user_pc) 
name = user_name.split(".")
name = str(name[0]).capitalize()

base_path = fr"{user_pc}\35159_147728_DUPREE_VENTA_DIRECTA_S_A\MERCADEO PAISES - MERCADEO PAISES DOCUMENTOS\Mercadeo BI\1. Arquitectura\Bases\15. Base Revisión LLL" 
params_path = fr"{base_path}\2. Imágenes\params\params.json"
version_path = fr"{base_path}\2. Imágenes\version.json"


with open(params_path, "r", encoding="utf-8") as f:
    params = json.load(f)

with open(version_path, "r", encoding="utf-8") as f:
    version = json.load(f)

version_actual = "1.0"
version_sp = version.get("version")

files_loaded = False

version_path = fr"{base_path}/2. Imágenes/version.json"

if os.path.exists(version_path):
    with open(version_path, "r") as f:
        version_data = json.load(f)
        file_version = version_data.get("version", None)

#Variable global para insertar texto en la consola
consola = None

#Archivo seleccionado
selected_file = None

#Variables de estilos
font_h1 = "Gilroy Black", 16
font_h2 = "Gilroy Black", 12 
font_normal = "Gilroy Medium", 14
font_consola = "Consolas", 12
background_color = "#FFF"
items_color = "#4BCD5E"
text_color = "#292731"
disabled_color = "#7C8483"

""" items_color = "#FF4155"
text_color = "#0A100D" """

def main():
    ui()

def cargar_xlsx():
    global consola, btn_evaluar_lll, selected_file, files_loaded

    print("Cargando Leader List Lite...")

    # Abre explorador de archivos para un solo archivo
    path = filedialog.askopenfilename(
        title="Seleccione el archivo LLL",
        filetypes=[("Archivos Excel", ("*.xlsb", "*.xlsx"))]
    )

    if path:
        selected_file = path
        files_loaded = True

        archivo_msg = f">>{name}, el archivo elegido\nfue:→ {os.path.basename(path)}\n\n"
        consola.configure(state="normal")         
        consola.insert("end", archivo_msg)     
        consola.see("end")
        consola.configure(state="disabled")  

        activar_analisis()
    else:
        print("No se cargó ningún archivo.")

def ui():
    import datetime
    global consola, btn_evaluar_lll, files_loaded, pais, campana

    ventana = tk.Tk()
    ventana.title("Fallometro - Verificador de LLL")
    ventana.geometry("550x450")
    ventana.configure(bg=background_color)
    ventana.attributes('-alpha', 0.98)
    ruta_ico = fr"{base_path}/2. Imágenes/favicon.ico"

    #Estilos
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")          

    try:
        ventana.iconbitmap(ruta_ico) 
    except Exception as e:
        messagebox.showerror("Error de sincronización ♾️", "Marketing Países / Mercadeo BI se encuentra sincronizada de manera erronea.")
    
    ventana.resizable(False, False)    

    titulo = tk.Label(ventana, text="Fallometro — Verificador de Leader List Lite", font=font_h1, bg=background_color, fg=text_color)
    titulo.pack(pady=(15,0))

    #Establecer dos columnas
    frame_contenido = tk.Frame(ventana, bg="#FFF")
    frame_contenido.pack(fill="both", expand=True, padx=20, pady=5)

    #----------------------Columna izquierda
    columna_izquierda = tk.Frame(frame_contenido, bg="#FFF")
    columna_izquierda.grid(row=0, column=0, sticky="nsew")

    # Columna derecha
    columna_derecha = tk.Frame(frame_contenido, bg=background_color)
    columna_derecha.grid(row=0, column=1, sticky="nsew")

    # Configurar proporción de columnas
    frame_contenido.columnconfigure(0, weight=1)
    frame_contenido.columnconfigure(1, weight=1)    

    #both 
    frame_both = tk.Frame(ventana, bg="#FFF")
    frame_both.pack(fill="both", expand=True, padx=5, pady=5)  

    #Contenido en izquierda
    #--------------------------DropDowns
    dropdown_frame = ctk.CTkFrame(
        columna_izquierda,
        fg_color="transparent",
        corner_radius=0,
        width=200            
    )
    # evitar cambio de tamaño
    dropdown_frame.pack(pady=(16, 0), anchor="center")
    dropdown_frame.pack_propagate(False)

    # grid
    dropdown_frame.grid_columnconfigure(0, weight=3)   # 30% del espacio
    dropdown_frame.grid_columnconfigure(1, weight=7)   # 70% del espacio
    dropdown_frame.grid_rowconfigure(0, weight=1)    

    #pais
    pais = tk.StringVar(value="País")    
    pais.trace_add("write", lambda *args: activar_analisis())
    paises = params.get("paises")

    pais_dropdown = ctk.CTkOptionMenu(
        dropdown_frame,
        variable=pais,
        values=paises,
        fg_color=text_color,
        button_color=text_color,
        text_color=items_color,
        dropdown_fg_color=text_color,
        dropdown_text_color=items_color,
        corner_radius=6,
        height=30,
        width=80,
        dynamic_resizing=False,
        font=font_normal
    )
    pais_dropdown.grid(row=0, column=0, sticky="ew", padx=(0,5))    

    #campaña
    campana = tk.StringVar(value="Campaña")
    campana.trace_add("write", lambda *args: activar_analisis())
    campana_inicial = params.get("camp_inicial")
    fecha = datetime.date.today()
    mes = fecha.month
    anio = fecha.year
    #mes = 12

    camp  = mes + 7
    if camp > 18:
        camp -= 18
        anio += 1
    campana_final   = anio * 100 + camp

    #indices
    anio_inicial = campana_inicial // 100
    periodo_inicial = campana_inicial % 100
    idx_inicial = anio_inicial*18 + (periodo_inicial - 1)
    anio_final = campana_final // 100
    periodo_final = campana_final % 100
    idx_final = anio_final*18 + (periodo_final - 1)

    #range
    idxs = range(idx_inicial, idx_final + 1)
    campanas = [
        (i // 18)*100 + (i % 18 + 1)
        for i in idxs
    ]
    campanas.sort(reverse=True)  # para dropdown de mayor a menor

    campana_dropdown = ctk.CTkOptionMenu(
        dropdown_frame,
        variable=campana,
        values=[str(c) for c in campanas],
        fg_color=text_color,
        button_color=text_color,
        text_color=items_color,
        dropdown_fg_color=text_color,
        dropdown_text_color=items_color,
        corner_radius=6,
        height=30,
        width=110,
        dynamic_resizing=False,
        font=font_normal
    )
    campana_dropdown.grid(row=0, column=1, sticky="ew", padx=(5,0))

    #botones
    btn_cargar_lll = ctk.CTkButton(
        columna_izquierda, 
        text="Cargar Leader List Lite",
        command=cargar_xlsx,
        fg_color=text_color,
        text_color=items_color,
        font=font_normal,
        corner_radius=12,
        height=60, 
        width=200        
    )
    btn_cargar_lll.pack(padx=0, pady=(20, 0), anchor="center")    

    #botones
    btn_evaluar_lll = ctk.CTkButton(
        columna_izquierda, 
        text="Comenzar evaluación de\nLLL cargados",
        command=validar_lll,
        #fg_color=items_color,
        fg_color = disabled_color,
        text_color=text_color,
        font=font_normal,
        corner_radius=12,
        height=60, 
        width=200,
        state="disabled"        
    )
    btn_evaluar_lll.pack(padx=0, pady=(20,20), anchor="center")

    btn_descargar = ctk.CTkButton(
        columna_izquierda, 
        text="Descargar resultados",
        #command=cargar_xlsx,
        #fg_color="#005E9E",
        fg_color= disabled_color,
        text_color="#E6E6E6",
        font=font_normal,
        corner_radius=12,
        height=60, 
        width=200,
        state="disabled"        
    )
    btn_descargar.pack(padx=0, pady=0, anchor="center")

    #----------------------Columna izquierda
    
    consola = ctk.CTkTextbox(
    columna_derecha,
    corner_radius=12,
    fg_color=text_color,
    text_color=items_color,
    font=font_consola,
    height=345
    )
    consola.pack(pady=(16,0), padx=(0,20), fill="both", expand=True)  

    bienvenida = f">>¡Hola {name}!\nTe doy la bienvenida Fallometro\nPor favor, comienza cargando\nlos LLL a revisar...\n\n"
    consola.insert("0.0", bienvenida)

    consola.configure(state="disabled")

    #*----------------------Logo IN sobrepuesta
    # Carga y redimensiona la imagen a 10x10 píxeles
    try:
        imagen_original = Image.open(f"{base_path}/2. Imágenes/logo.png")
    except Exception as e:
        messagebox.showwarning("Error", "Los iconos no pueden ser encontrados, por favor, sincronice sus carpetas.")

    imagen_redimensionada = imagen_original.resize((130, 62), resample=Image.Resampling.LANCZOS)
    imagen_tk = ImageTk.PhotoImage(imagen_redimensionada)

    # Crea el Label con la imagen
    lbl_imagen = tk.Label(ventana, image=imagen_tk, bd=0, bg="#FFF")
    lbl_imagen.image = imagen_tk 

    # Delante de todo y coordenadas
    lbl_imagen.place(relx=1.0, y=350, x=-348, anchor="ne")    
    lbl_imagen.lift()

    ventana.mainloop()

def activar_analisis():
    global btn_evaluar_lll, files_loaded, pais_seleccionado, campana_seleccionada
    #get valores
    pais_seleccionado = pais.get()
    campana_seleccionada = campana.get()
    
    #condiciones
    if files_loaded and pais_seleccionado != "País" and campana_seleccionada != "Campaña":
        btn_evaluar_lll.configure(state="normal", fg_color=items_color)
    else:
        btn_evaluar_lll.configure(state="disabled", fg_color=disabled_color)

def validar_lll():
    import pandas as pd
    global selected_file
    
    print("País seleccionado: ", pais_seleccionado)
    print("Campaña: ", campana_seleccionada)

    #almacen de errores
    errores_lll = []
    fila_inicial = 6

    #encontrar el final del LLL
    def encontrar_fin_datos(df, fila_inicial):
        #buscar la primera fila completamente vacía desde fila_inicial
        for idx in range(fila_inicial, len(df)):
            if df.iloc[idx].isna().all():
                return idx  #devolver el indice de la ultima fila
        return len(df) #si no lo encuentra devolver el len del df    

    #validar igual a cero
    def validar_cero(columna:int):
        for idx in range(fila_inicial, fila_final):
            valor = df_CAT.iat[idx, columna]
            if valor == 0:
                errores_lll.append({'fila': idx, 'col': columna})

    #validar vacios o errores
    def validar_vacios(columna:int):
        for idx in range(fila_inicial, fila_final):
            valor = df_CAT.iat[idx, columna]
            # detección de vacío o error simple (#REF!, NaN, cadena vacía)
            if pd.isna(valor) or (isinstance(valor, str) and (valor.strip() == "" or valor.startswith("#"))):
                errores_lll.append({'fila': idx, 'col': columna})

    #validar que sea numero
    def validar_numerico(columna: int, es_entero: bool = False):
        for idx in range(fila_inicial, fila_final):
            valor = df_CAT.iat[idx, columna]

            #saltar celdas vacías
            if pd.isna(valor) or es_error_excel(valor):
                continue

            #intentar conversión
            try:
                if isinstance(valor, str):
                    valor = valor.replace(',', '.')
                    num = float(valor.strip())
                else:
                    num = float(valor)

                #validar tipo numérico
                if es_entero:
                    if not num.is_integer():
                        raise ValueError("No es entero")

            except (ValueError, TypeError, AttributeError):
                errores_lll.append({'fila': idx,'col': columna})

    #valores repetidos en la misma columna
    def validar_duplicados(columna:int):
        valores_vistos = set() #almacen

        for idx in range(fila_inicial, fila_final):
            valor = df_CAT.iat[idx, columna]

            # Verificar duplicados
            if valor in valores_vistos:
                errores_lll.append({'fila': idx, 'col': columna})
            else:
                valores_vistos.add(valor)

    #validar columnas con arrays de params, validar que este in, que no sean vacío u error
    def validacion_basica(columna:int, array_evaluativo):
        
        for idx in range(fila_inicial, fila_final):
            valor = df_CAT.iat[idx, columna]
            #celda vacía o NaN
            if pd.isna(valor) or (isinstance(valor, str) and valor.strip() == ""):
                errores_lll.append({'fila': idx, 'col': columna})

            #error de Excel (#REF!, #N/A…)
            if es_error_excel(valor):
                errores_lll.append({'fila': idx, 'col': columna})
                continue

            #validación contra validos
            v_str = str(valor).strip()
            if v_str not in array_evaluativo:
                errores_lll.append({'fila': idx, 'col': columna})

    def validar_rango_paginas(pagina_evaluada: int):
        col_vehiculo = 1 
        for idx in range(fila_inicial, fila_final):
            vehiculo = df_CAT.iat[idx, col_vehiculo]
            pagina = df_CAT.iat[idx, pagina_evaluada]

            try:
                #convertir int
                num_pagina = int(pagina)

                #filtrar rango del vehículo
                rango_vehiculo = rango_paginas[rango_paginas['codi_vehi'] == vehiculo]

                if rango_vehiculo.empty:
                    errores_lll.append({'fila': idx, 'col': col_vehiculo})
                else:
                    #obtener límites
                    pag_inicial = rango_vehiculo['PAG INICIAL'].values[0]
                    pag_final = rango_vehiculo['PAG FINAL'].values[0]

                    #validar rango
                    if not (pag_inicial <= num_pagina <= pag_final):
                        errores_lll.append({'fila': idx, 'col': pagina_evaluada})

            except Exception as e:
                errores_lll.append({'fila': idx, 'col': pagina_evaluada})

    #detector de errores nativos de excel
    def es_error_excel(v):
        return isinstance(v, str) and v.startswith("#") and v.endswith("!")

    def leer_lll(selected_file, sheets):
        ext = os.path.splitext(selected_file)[-1].lower()

        if ext == ".xlsb":
            df = pd.read_excel(selected_file, sheet_name=sheets, header=None, engine="pyxlsb")
        elif ext in [".xlsx", ".xls"]:
            df = pd.read_excel(selected_file, sheet_name=sheets, header=None, engine="openpyxl")  
        else:
            raise ValueError("Formato de archivo no soportado") 
        return df

    def consumir_conSQL(columna_inicial, columna_final, filas:int):
        #toca asi porque no es capaz con la hoja completa
        df_conSQL = pd.read_excel(selected_file,
            sheet_name='Consultas SQL',
            engine='pyxlsb',
            header=1, #primera fila es header
            usecols=f"{columna_inicial}:{columna_final}", nrows=int(filas)) #indicativo columnas
        
        df_conSQL = df_conSQL.dropna(how='all')
        #print(df_conSQL)
        return df_conSQL
    

    consola.configure(state="normal")         
    consola.insert("end", ">>Comenzando análisis del LLL cargado...\n\n")     
    consola.see("end")
    consola.configure(state="disabled")      

    if not selected_file:
        messagebox.showerror("Sin archivo", "No se ha seleccionado un archivo, por favor, seleccione alguno antes de continuar")
        return
                
    #carga  
    try:
        df_CAT = leer_lll(selected_file, "LL CAT")
    except Exception as e:
        messagebox.showerror("Error al leer archivo", f"El archivo cargado no corresponde a un LLL")
        consola.configure(state="normal")         
        consola.insert("end", "\n>>ERROR: El archivo cargado no corresponde a un LLL...")     
        consola.see("end")
        consola.configure(state="disabled")
        return

    #-----------------Validar que sea un LLL a través de las primeras 71 columnas
    columnas_lll = params.get("cols_LLL")
    fila_final = encontrar_fin_datos(df_CAT, fila_inicial)

    #print(columnas_lll)
    
    #Definición headers
    #print("Archivo cargado")
    headers = df_CAT.iloc[5].fillna("").tolist()

    #Si coinciden los headers significa que es un LLL
    if headers[:71] != columnas_lll:
        consola.configure(state="normal")         
        consola.insert("end", "\n>>ERROR: El archivo cargado no corresponde a un LLL...")     
        consola.see("end")
        consola.configure(state="disabled")
        messagebox.showerror("Error en la lectura", "El LLL seleccionado no tiene la estructura adecuada. Por favor, no modificar el formato original entregado por el Área de Precios y Optimización. Específicamente por Don Rodrigo")

    #----------------Validar que la campaña y el país sea igual
    campana_B5 = df_CAT.iloc[4,1]  
    if str(campana_B5) != str(campana_seleccionada):
        messagebox.showerror("Disparidad en campañas", "La campaña seleccionada no corresponde a la campaña del LLL cargado.")        
    
    pais_L5 = df_CAT.iloc[4,11]
    if str(pais_L5) != str(pais_seleccionado):
        messagebox.showerror("Disparidad en país", "El país seleccionado no corresponde al país del LLL cargado.")        
    
    consola.configure(state="normal")         
    consola.insert("end", f">>Analizando todos los datos\ningresados...\n\n")     
    consola.see("end")
    consola.configure(state="disabled") 

    #--------------------Validar Tipo de Venta
    tipos_venta = set(params.get("tipo_venta"))
    validacion_basica(0, tipos_venta)

    #--------------------Validar Vehiculo
    vehiculo = set(params.get("vehiculo"))
    validacion_basica(1, vehiculo)

    #--------------------Validar Estrategía
    codi_estrategia = set(params.get("codigos_estrategia"))
    validacion_basica(2, codi_estrategia)

    #--------------------Validar Tipo Programación
    tipo_prog = set(params.get("tipo_prog"))
    validacion_basica(12, tipo_prog)

    #--------------------Validar COLECCIÓN CAPSULA
    validar_vacios(13)

    #--------------------Validar PAG NAL
    rango_paginas = consumir_conSQL("A", "F", 40)
    #print(rango_paginas)
    validar_vacios(14)
    validar_rango_paginas(14)

    #--------------------Validar PAG NAL
    validar_vacios(15)
    validar_rango_paginas(15)

    #--------------------Validar COD VENTA
    validar_vacios(18)
    validar_duplicados(18)
    
    #--------------------Validar COD PROD
    validar_vacios(19)
    validar_duplicados(19)

    #--------------------Validar DESCRIPCIÓN COMERCIAL PAÍS
    validar_vacios(20)
    validar_duplicados(20)

    #--------------------Validar OBSERVACION PDF
    valores_pdf = consumir_conSQL("CJ", "CJ", 10)
    valores_pdf = valores_pdf.iloc[:, 0].tolist()
    validacion_basica(26, valores_pdf)

    #--------------------Validar COSTO INICIAL
    validar_vacios(29)

    #--------------------Validar CANASTA FINAL
    validar_vacios(30)

    #--------------------Validar FLETE
    validar_vacios(33)

    #--------------------Validar PP
    validar_vacios(34)

    #--------------------Validar PO
    validar_vacios(35)

    #--------------------Validar DCTO ASESORA
    validar_vacios(37)

    #--------------------Validar V NETO
    validar_vacios(40)

    #--------------------Validar FACTOR INICIAL/BASE
    validar_vacios(41)

    #--------------------Validar AJUSTE FACTOR
    for idx in range(fila_inicial, fila_final):
        valor = float(df_CAT.iat[idx, 42])
        if valor < 0 or valor > 2:
            errores_lll.append({'fila': idx, 'col': 42})

    #--------------------Validar FACTOR
    validar_vacios(43)
    validar_cero(43)

    #--------------------Validar UNIDAD NAL
    validar_vacios(44)

    #--------------------Validar UNIDAD ADV
    validar_vacios(45)

    #--------------------Validar UNIDADES TOTALES
    validar_vacios(46)

    #--------------------Validar INV SOBRANTE A C+3 
    validar_vacios(47)
    validar_cero(47)

    #--------------------Validar % CMV
    validar_vacios(51)
    #validar el procentaje
    for idx in range(fila_inicial, fila_final):
        valor = df_CAT.iat[idx, 51]
    
        try:
            # Convertir diferentes formatos a float
            if isinstance(valor, str):
                # Quitar caracteres no numéricos excepto . y %
                valor_limpio = valor.strip().replace(',', '.').replace('%', '')
                if not valor_limpio:
                    raise ValueError("Vacío")

                porcentaje = float(valor_limpio)

                # Verificar si tenía % para ajustar escala
                if '%' in valor:
                    porcentaje = porcentaje  # 50% = 50.0
                else:
                    porcentaje = porcentaje * 100  # 0.5 se considera 50%
            else:
                porcentaje = float(valor) * 100  # Si es decimal (0.37 -> 37%)

            # Validar rango
            if porcentaje <= 0 or porcentaje > 100:
                errores_lll.append({'fila': idx, 'col': 51})

        except (ValueError, TypeError, AttributeError):
            # Registrar errores de conversión
            print("Fallo en la evaluación del CMV")
    
    consola.configure(state="normal")         
    consola.insert("end", f">>Realizando cientos de cálculos\ncomplejos...\n\n")     
    consola.see("end")
    consola.configure(state="disabled")
    
    #--------------------Validar COD ESTIMACIÓN
    codis_estimacion = set(params.get("cod_estimacion"))
    validacion_basica(53, codis_estimacion)

    #--------------------Validar AGOTAR EXISTENCIA
    agotar_existencia = set(params.get("agotar_existencia"))
    validacion_basica(54, agotar_existencia)

    #--------------------Validar DIGITABLE
    digitables = set(params.get("digitables"))
    validacion_basica(55, digitables)

    #--------------------Validar ORIGEN
    validar_vacios(56)

    #--------------------Validar PEDIDOS 
    pedidos = consumir_conSQL("AK", "AN", 30)
    pedidos['codi_camp'] = pedidos["codi_camp"].astype(int)
    pedidos['tota_pedi'] = pedidos["tota_pedi"].astype(int)
    pedidos = pedidos[pedidos['codi_camp'] == int(campana_seleccionada)]['tota_pedi'].values[0]

    pedidos = pedidos

    for idx in range(fila_inicial, fila_final):
        valor = float(df_CAT.iat[idx, 57])
        if valor != pedidos:
            errores_lll.append({'fila': idx, 'col': 57})

    #--------------------Validar INDICADOR
    indicador = set(params.get("indicador"))
    validacion_basica(58, indicador)
    
    #--------------------Validar GRAMAJE
    #--------------------Validar UNIDAD DE MEDIDA
    #--------------------Validar PUM

    #--------------------Validar PUNTOS
    puntos = set(params.get("puntos"))
    validacion_basica(68, puntos)

    #--------------------Validar CAMPAÑA
    for idx in range(fila_inicial, fila_final):
        valor = float(df_CAT.iat[idx, 69])
        if valor != int(campana_seleccionada):
            errores_lll.append({'fila': idx, 'col': 69})

    #--------------------Validar PUNTOS/ MAXIPUNTAJE
    validar_vacios(70)

    for idx in range(fila_inicial, fila_final):
        valor = float(df_CAT.iat[idx, 70])
        if valor != int:
            errores_lll.append({'fila': idx, 'col': 70})




    print(errores_lll)

    num_errores = len(errores_lll)

    if num_errores == 0:
        consola.configure(state="normal")         
        consola.insert("end", f">>¡¡Felicitaciones!!, se han encontrado {num_errores} hallazgos :)\n")     
        consola.see("end")
        consola.configure(state="disabled")  
    else:
        consola.configure(state="normal")         
        consola.insert("end", f">>Análisis terminado. Se han encontrado {num_errores} hallazgos en el LLL cargado.\n")     
        consola.see("end")
        consola.configure(state="disabled")  

    print("perrito")



    


if version_actual == version_sp:
    main()
else:
    messagebox.showerror("Error de versión en Fallometro", f"Por favor, actualice a la versión más reciente: v{version_sp}")
















