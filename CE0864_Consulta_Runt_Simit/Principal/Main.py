import os
import subprocess
import tkinter as tk
import sys
from PIL import Image, ImageTk

def cerrar_ventana():
        sys.exit()
        
def consulta_vin():
    # Ruta del archivo .bat
    ruta_bat_consulta_vim = "C:\\RPA\\CE0864_Consulta_Runt_Simit\\Aut. Runt\\Data\\ConsultaVIM.bat"
    
    # Verificar si el archivo existe
    if os.path.exists(ruta_bat_consulta_vim):
        # Abrir una nueva instancia de cmd y ejecutar el archivo .bat dentro de ella
        subprocess.Popen(["cmd", "/c", "start", "cmd", "/k", "call", ruta_bat_consulta_vim], shell=True)
    else:
        # Si el archivo no existe, mostrar un mensaje de error
        print("El archivo .bat no existe en la ruta especificada.")
        
def notificar_vim():
    # Ruta del archivo .bat
    ruta_bat_notificar_vim = "C:\\RPA\\CE0864_Consulta_Runt_Simit\\Aut. Runt\\Data\\NotificarVIM.bat"
    
    # Verificar si el archivo existe
    if os.path.exists(ruta_bat_notificar_vim):
        # Abrir una nueva instancia de cmd y ejecutar el archivo .bat dentro de ella
        subprocess.Popen(ruta_bat_notificar_vim)
    else:
        # Si el archivo no existe, mostrar un mensaje de error
        print("El archivo .bat no existe en la ruta especificada.")
def consulta_placa_cedula():
    # Ruta del archivo .bat
    ruta_bat_consulta_placa_cedula = "C:\\RPA\\CE0864_Consulta_Runt_Simit\\Aut. Runt x Cedula y Placa\\Data\\Consulta_Placa_Cedula.bat"
    
    # Verificar si el archivo existe
    if os.path.exists(ruta_bat_consulta_placa_cedula):
        # Abrir una nueva instancia de cmd y ejecutar el archivo .bat dentro de ella
        subprocess.Popen(["cmd", "/c", "start", "cmd", "/k", "call", ruta_bat_consulta_placa_cedula], shell=True)
    else:
        # Si el archivo no existe, mostrar un mensaje de error
        print("El archivo .bat no existe en la ruta especificada.")
        
def notificar_SIMIT():
    # Ruta del archivo .bat
    ruta_bat_notificar_SIMIT = r"C:\RPA\CE0864_Consulta_Runt_Simit\Aut. Simit x Placa y Cedula\Data\Notificar.bat"                                
    
    # Verificar si el archivo existe
    if os.path.exists(ruta_bat_notificar_SIMIT):
        # Abrir una nueva instancia de cmd y ejecutar el archivo .bat dentro de ella
        subprocess.Popen(ruta_bat_notificar_SIMIT)
    else:
        # Si el archivo no existe, mostrar un mensaje de error
        print("El archivo .bat no existe en la ruta especificada.")

def notificar_cedula_placa():
    # Ruta del archivo .bat
    ruta_bat_notificar_cedula_placa = "C:\\RPA\\CE0864_Consulta_Runt_Simit\\Aut. Runt x Cedula y Placa\\Data\\NotificarPLacaCedula.bat"
    
    # Verificar si el archivo existe
    if os.path.exists(ruta_bat_notificar_cedula_placa):
        # Abrir una nueva instancia de cmd y ejecutar el archivo .bat dentro de ella
        subprocess.Popen(ruta_bat_notificar_cedula_placa)
    else:
        # Si el archivo no existe, mostrar un mensaje de error
        print("El archivo .bat no existe en la ruta especificada.")


def consulta_licencia():
    # Ruta del archivo .bat
    ruta_bat_consulta_licencia = "C:\\RPA\\CE0864_Consulta_Runt_Simit\\Aut. Runt Para Licencia\\Data\\ConsulteLicencia.bat"
    
    # Verificar si el archivo existe
    if os.path.exists(ruta_bat_consulta_licencia):
        # Abrir una nueva instancia de cmd y ejecutar el archivo .bat dentro de ella
        subprocess.Popen(["cmd", "/c", "start", "cmd", "/k", "call", ruta_bat_consulta_licencia], shell=True)
    else:
        # Si el archivo no existe, mostrar un mensaje de error
        print("El archivo .bat no existe en la ruta especificada.")
        
def notificar_licencia():
     #Ruta del archivo .bat
    ruta_bat_notificar_licencia = "C:\\RPA\\CE0864_Consulta_Runt_Simit\\Aut. Runt Para Licencia\\Data\\NotificarLicencia.bat"
    
    #Verificar si el archivo existe
    if os.path.exists(ruta_bat_notificar_licencia):
        # Abrir una nueva instancia de cmd y ejecutar el archivo .bat dentro de ella
        subprocess.Popen(ruta_bat_notificar_licencia)
    else:
        #Si el archivo no existe, mostrar un mensaje de error
        print("El archivo .bat no existe en la ruta especificada.")
def consulta_simit():
    # Ruta del archivo .bat
    ruta_bat_consulta_simit = "C:\\RPA\\CE0864_Consulta_Runt_Simit\\Aut. Simit x Placa y Cedula\\Data\\ConsultaSimit.bat"
    
    # Verificar si el archivo existe
    if os.path.exists(ruta_bat_consulta_simit):
        # Abrir una nueva instancia de cmd y ejecutar el archivo .bat dentro de ella
        subprocess.Popen(["cmd", "/c", "start", "cmd", "/k", "call", ruta_bat_consulta_simit], shell=True)
    else:
        # Si el archivo no existe, mostrar un mensaje de error
        print("El archivo .bat no existe en la ruta especificada.")


def main():
    ventana = tk.Tk()
    ventana.title('Automatización Validación de Requisitos Convenio de Rodamiento')
    ventana.geometry('500x310')
    ventana.configure(bg='white')
    ventana.iconbitmap('C:\\RPA\\CE0864_Consulta_Runt_Simit\\Principal\\img\\censLogo.ico')
    ventana.resizable(False,False)
    imagen = Image.open(r"C:\\RPA\\CE0864_Consulta_Runt_Simit\\Principal\\img\\CensLogo.png")
    imagen_resized = imagen.resize((550, 330), Image.Resampling.LANCZOS)
    fondo_imagen = ImageTk.PhotoImage(imagen_resized)
    etiqueta = tk.Label(ventana,image=fondo_imagen)
    etiqueta.pack()
    #ventana.configure(background='white')
    #Boton Cancelar
    boton_cancelar = tk.Button(ventana, text='Cancelar', command=cerrar_ventana, 
    bg='red', 
    fg='white',
    activebackground="lightblue",
    activeforeground="black",
    cursor="hand2",
    height="1", 
    width="8",
    bd="5",
    borderwidth=5,
    relief='raised',
    font=('Arial', 10, 'bold'))
    boton_cancelar.pack()
    boton_cancelar.place(relx=1.0, rely=1.0, x=-10, y=-15, anchor='se')
    #boton VIM
    icono_vim = tk.PhotoImage(file="C:\\RPA\\CE0864_Consulta_Runt_Simit\\Principal\\img\\VIMicono.png")
    boton_VIM = tk.Button(ventana, image=icono_vim, command=consulta_vin,
    compound=tk.CENTER, 
    bd="5",
    highlightthickness=0, 
    cursor="hand2", 
    activebackground="lightblue",
    activeforeground="black",
    background= "white",
    )
    boton_VIM.pack(padx=20, pady=20)
    boton_VIM.place(relx=0.0, rely=1.0, x=10, y=-205, anchor='sw')
    #boton notificar VIM
    icon_notificar_vim = tk.PhotoImage(file="C:\\RPA\\CE0864_Consulta_Runt_Simit\\Principal\\img\\icone-outlook-bleu.png")
    boton_notificar_vim = tk.Button(ventana, image=icon_notificar_vim, command=notificar_vim,
    bd="5",
    highlightthickness=0, 
    cursor="hand2", 
    activebackground="lightblue",
    activeforeground="black",
    background= "white",
    )
    boton_notificar_vim.pack()
    boton_notificar_vim.place(relx=0.0, rely=1.0, x=27, y=-130, anchor='sw')
    #boton Placa y Cedula
    icon_placa_cedula = tk.PhotoImage(file="C:\\RPA\\CE0864_Consulta_Runt_Simit\\Principal\\img\\PlacaCedulaicon.png")
    boton_placa_cedula = tk.Button(ventana, image=icon_placa_cedula, command=consulta_placa_cedula,
    bd="5",
    highlightthickness=0, 
    cursor="hand2", 
    activebackground="lightblue",
    activeforeground="black",
    background= "white",
    )
    boton_placa_cedula.pack()
    boton_placa_cedula.place(relx=0.0, rely=1.0, x=140, y=-206, anchor='sw')
    #boton notificar Placa y Cedula
    icon_notificar_placa_cedula = tk.PhotoImage(file="C:\\RPA\\CE0864_Consulta_Runt_Simit\\Principal\\img\\icone-outlook-bleu.png")
    boton_notificar_placa_cedula = tk.Button(ventana, image=icon_notificar_placa_cedula, command=notificar_cedula_placa,
    bd="5",
    highlightthickness=0, 
    cursor="hand2", 
    activebackground="lightblue",
    activeforeground="black",
    background= "white",                                      
    )
    boton_notificar_placa_cedula.pack()
    boton_notificar_placa_cedula.place(relx=0.0, rely=1.0, x=160, y=-131, anchor='sw')
    #boton licencia
    icon_licencia = tk.PhotoImage(file="C:\\RPA\\CE0864_Consulta_Runt_Simit\\Principal\\img\\Licenciaicon.png")
    boton_consulta_licencia = tk.Button(ventana, image=icon_licencia, command=consulta_licencia,
    bd="5",
    highlightthickness=0, 
    cursor="hand2",
    activebackground="lightblue",
    activeforeground="black",
    background="white" 
    #state="disabled"  # Deshabilitado

    )
    boton_consulta_licencia.pack()
    boton_consulta_licencia.place(relx=0.0, rely=1.0, x=270, y=-207, anchor='sw')
    #boton notificar licencia
    icon_notificar_licencia = tk.PhotoImage(file="C:\\RPA\\CE0864_Consulta_Runt_Simit\\Principal\\img\\icone-outlook-bleu.png")
    boton_notificar_licencia = tk.Button(ventana, image=icon_notificar_licencia, command=notificar_licencia,
    bd="5",
    highlightthickness=0, 
    activebackground="lightblue",
    activeforeground="black",
    cursor="hand2",
    background="white"
    #state="disabled"  # Deshabilitado
    )
    boton_notificar_licencia.pack()
    boton_notificar_licencia.place(relx=0.0, rely=1.0, x=290, y=-133, anchor='sw')

    #boton simit
    icon_simit = tk.PhotoImage(file="C:\\RPA\\CE0864_Consulta_Runt_Simit\\Principal\\img\\Simiticopng.png")
    boton_simit = tk.Button(ventana, image=icon_simit, command=consulta_simit,
    bd="5",
    highlightthickness=0, 
    cursor="hand2", 
    activebackground="lightblue",
    activeforeground="black",
    background= "white",
    )
    boton_simit.pack()
    boton_simit.place(relx=0.0, rely=1.0, x=400, y=-205, anchor='sw')
    
#boton notificar SIMIT
    icon_notificar_SIMIT = tk.PhotoImage(file="C:\\RPA\\CE0864_Consulta_Runt_Simit\\Principal\\img\\icone-outlook-bleu.png")
    boton_notificar_SIMIT = tk.Button(ventana, image=icon_notificar_SIMIT, command=notificar_SIMIT,
    bd="5",
    highlightthickness=0, 
    activebackground="lightblue",
    activeforeground="black",
    cursor="hand2",
    background= "white",                                      
    )
    boton_notificar_SIMIT.pack()
    boton_notificar_SIMIT.place(relx=0.0, rely=1.0, x=420, y=-131, anchor='sw')
    
    ventana.mainloop()
if __name__ == "__main__":
    main()