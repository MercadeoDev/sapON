#Interfaz gráfica
import tkinter as tk
from tkinter import messagebox, filedialog
from tkinter.scrolledtext import ScrolledText
import customtkinter as ctk
from PIL import Image, ImageTk 

import os
import json
import sys


print(tk.TkVersion)
print(tk.TclVersion)


user_pc = os.environ["USERPROFILE"]
user_name = os.path.basename(user_pc) 
name = user_name.split(".")
name = str(name[0]).capitalize()

base_path = fr"{user_pc}\35159_147728_DUPREE_VENTA_DIRECTA_S_A\MERCADEO PAISES - MERCADEO PAISES DOCUMENTOS\Mercadeo BI\1. Arquitectura\Bases\15. Base Revisión LLL" 

version_actual = "1.0"
version_sp = "1.0"

version_path = fr"{base_path}/2. Imágenes/version.json"

if os.path.exists(version_path):
    with open(version_path, "r") as f:
        version_data = json.load(f)
        file_version = version_data.get("version", None)

consola = None
font_h1 = "Gilroy Black", 16
font_h2 = "Gilroy Black", 12 
font_normal = "Gilroy Medium", 14
font_consola = "Consolas", 12
background_color = "#FFF"
items_color = "#4BCD5E"
text_color = "#292731"
disabled_color = "#7C8483"

def main():
    ui()

def cargar_xlsx():
    global consola, btn_evaluar_lll
    print("Cargando Leader List Lite...")
    # Abre el explorador de archivos y permite múltiples selecciones
    paths_xlsb = filedialog.askopenfilenames(
        title="Seleccione los archivos LLL",
        filetypes=[("Archivos Excel", ("*.xlsb", "*.xlsx"))]
    )

    if paths_xlsb:
        print(f"{len(paths_xlsb)} archivo(s) cargado(s).")
        count_archivos = len(paths_xlsb)
        if count_archivos == 1:
            count_archivos = "un"
        archivos_msg = f">>{name}, cargaste " + str(count_archivos) + " LLL. \nSi es correcto oprime el botón para comenzar a evaluar...\nde lo contario, carga los\narchivos correctos...\n\n"
        #print(count_archivos)
        consola.configure(state="normal")         
        consola.insert("end", archivos_msg)     
        consola.see("end") #scroll automático
        consola.configure(state="disabled")  

        btn_evaluar_lll = ctk.CTkButton(
            #command=cargar_xlsx,
            fg_color=items_color,
            state="active"        
        )
        btn_evaluar_lll.pack(padx=0, pady=(20,20), anchor="center")

        for path in paths_xlsb:
            print(f"→ {os.path.basename(path)}")
    else:
        print("No se cargó ningún archivo.")

def ui():
    import time
    global consola, btn_evaluar_lll   

    ventana = tk.Tk()
    ventana.title("SapON - Verificador de Leader List Lite")
    ventana.geometry("550x390")
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

    titulo = tk.Label(ventana, text="SapON — Verificador de errores en LLL", font=font_h1, bg=background_color, fg=text_color)
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
    btn_cargar_lll.pack(padx=0, pady=(15, 0), anchor="center")    

    #botones
    btn_evaluar_lll = ctk.CTkButton(
        columna_izquierda, 
        text="Comenzar evaluación de\nLLL cargados",
        #command=cargar_xlsx,
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
    height=305
    )
    consola.pack(pady=(16,0), padx=(0,20), fill="both", expand=True)  

    bienvenida = f">>¡Hola {name}!\nTe doy la bienvenida SapON.\nPor favor, comienza cargando\nlos LLL a revisar...\n\n"
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
    lbl_imagen.place(relx=1.0, y=314, x=-348, anchor="ne")    
    lbl_imagen.lift()

    ventana.mainloop()

if version_actual == version_sp:
    main()
else:
    messagebox.showerror("Error de versión en Sap ON", f"Por favor, actualice a la versión más reciente: v{version_sp}")
















