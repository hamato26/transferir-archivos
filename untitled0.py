import os
import shutil
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk  # Importa ttk para el ComboBox

# Función para seleccionar el archivo Excel
def seleccionar_archivo_excel():
    archivo_excel = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx")])
    archivo_excel_entry.delete(0, tk.END)
    archivo_excel_entry.insert(0, archivo_excel)

# Función para seleccionar la carpeta de origen
def seleccionar_carpeta_origen():
    carpeta_origen = filedialog.askdirectory()
    carpeta_origen_entry.delete(0, tk.END)
    carpeta_origen_entry.insert(0, carpeta_origen)

# Función para seleccionar la carpeta de destino
def seleccionar_carpeta_destino():
    carpeta_destino = filedialog.askdirectory()
    carpeta_destino_entry.delete(0, tk.END)
    carpeta_destino_entry.insert(0, carpeta_destino)

# Función para mover archivos PDF
def mover_archivos():
    archivo_excel = archivo_excel_entry.get()
    carpeta_origen = carpeta_origen_entry.get()
    carpeta_destino = carpeta_destino_entry.get()
    hoja_excel = hoja_excel_combo.get()

    try:
        if not os.path.isfile(archivo_excel):
            raise Exception("Selecciona un archivo Excel válido.")

        if not os.path.isdir(carpeta_origen):
            raise Exception("Selecciona una carpeta de origen válida.")

        if not os.path.isdir(carpeta_destino):
            raise Exception("Selecciona una carpeta de destino válida.")

        df = pd.read_excel(archivo_excel, sheet_name=hoja_excel)
        nombres_archivos = df["NOMBRE_DEL_ARCHIVO"].tolist()
        total_archivos = len(nombres_archivos)
        archivos_faltantes = []
        archivos_repetidos = []

        progress = ttk.Progressbar(ventana, orient="horizontal", length=400, mode="determinate")
        progress.grid(row=6, column=0, columnspan=3, padx=10, pady=10, sticky="w")
        progress["maximum"] = total_archivos

        for index, nombre_archivo in enumerate(nombres_archivos, start=1):
            progress["value"] = index
            ventana.update_idletasks()

            origen = os.path.join(carpeta_origen, nombre_archivo)
            destino = os.path.join(carpeta_destino, nombre_archivo)

            if os.path.exists(origen) and nombre_archivo.endswith(".pdf"):
                if not os.path.exists(destino):
                    shutil.move(origen, destino)
                else:
                    archivos_repetidos.append(nombre_archivo)
            else:
                archivos_faltantes.append(nombre_archivo)

        progress.grid_forget()  # Ocultar la barra de progreso después de completar la operación

        mensaje = "Operación completada."

        if archivos_faltantes:
            mensaje += f"\nArchivos faltantes: {len(archivos_faltantes)}. Archivos no encontrados en la carpeta de origen."

        if archivos_repetidos:
            mensaje += f"\nArchivos repetidos: {len(archivos_repetidos)}. Archivos con el mismo nombre ya existen en la carpeta de destino."

        resultado_label.config(text=mensaje)
        if archivos_faltantes or archivos_repetidos:
            messagebox.showwarning("Advertencia", mensaje)

    except FileNotFoundError as e:
        messagebox.showerror("Error", f"Error de archivo: {str(e)}")
    except PermissionError as e:
        messagebox.showerror("Error", f"Error de permisos: {str(e)}")
    except Exception as e:
        messagebox.showerror("Error", str(e))

# Función para mostrar la ayuda
def mostrar_ayuda():
    mensaje_ayuda = """\
    **Ayuda para Mover Archivos PDF**
    
    1. Selecciona un archivo Excel válido que contenga una hoja de datos con una columna llamada 'NOMBRE_DEL_ARCHIVO'. El programa moverá los archivos PDF correspondientes a esta lista.

    2. Selecciona una carpeta de origen donde se encuentren los archivos PDF que deseas mover.

    3. Selecciona una carpeta de destino donde se moverán los archivos PDF.

    4. Selecciona la hoja de Excel que contiene la lista de archivos a mover.

    5. Haz clic en el botón 'Mover Archivos' para iniciar la operación.

    Nota: El programa moverá los archivos PDF que coincidan con los nombres de la lista desde la carpeta de origen a la carpeta de destino. Los archivos repetidos no se sobrescribirán.
    """
    messagebox.showinfo("Ayuda", mensaje_ayuda)

# Crear la ventana de la GUI
ventana = tk.Tk()
ventana.title("Mover Archivos PDF")

# Configurar el tamaño de fuente
fuente_grande = ('Arial', 12)

# Colores y estilos
color_fondo = "#FFFFFF"  # Fondo blanco
color_boton = "#009688"  # Botón de color verde oscuro
color_texto = "#000000"  # Texto negro

# Configurar colores de fondo
ventana.configure(bg=color_fondo)

# Estilo para los botones
style = ttk.Style()
style.configure("TButton", padding=6, relief="flat", background=color_boton, foreground=color_texto)
style.map("TButton",
    background=[("active", "#007D68")],
    foreground=[("active", "#FFFFFF")]
)

# Etiqueta y entrada para el archivo Excel
archivo_excel_label = tk.Label(ventana, text="Archivo Excel:", font=fuente_grande, bg=color_fondo)
archivo_excel_label.grid(row=0, column=0, padx=10, pady=5, sticky="w")
archivo_excel_entry = tk.Entry(ventana, font=fuente_grande)
archivo_excel_entry.grid(row=0, column=1, padx=10, pady=5, sticky="ew")
seleccionar_archivo_excel_button = ttk.Button(ventana, text="Seleccionar Archivo", style="TButton", command=seleccionar_archivo_excel)
seleccionar_archivo_excel_button.grid(row=0, column=2, padx=10, pady=5)

# Etiqueta y entrada para la carpeta de origen
carpeta_origen_label = tk.Label(ventana, text="Carpeta de origen:", font=fuente_grande, bg=color_fondo)
carpeta_origen_label.grid(row=1, column=0, padx=10, pady=5, sticky="w")
carpeta_origen_entry = tk.Entry(ventana, font=fuente_grande)
carpeta_origen_entry.grid(row=1, column=1, padx=10, pady=5, sticky="ew")
seleccionar_origen_button = ttk.Button(ventana, text="Seleccionar Carpeta", style="TButton", command=seleccionar_carpeta_origen)
seleccionar_origen_button.grid(row=1, column=2, padx=10, pady=5)

# Etiqueta y entrada para la carpeta de destino
carpeta_destino_label = tk.Label(ventana, text="Carpeta de destino:", font=fuente_grande, bg=color_fondo)
carpeta_destino_label.grid(row=2, column=0, padx=10, pady=5, sticky="w")
carpeta_destino_entry = tk.Entry(ventana, font=fuente_grande)
carpeta_destino_entry.grid(row=2, column=1, padx=10, pady=5, sticky="ew")
seleccionar_destino_button = ttk.Button(ventana, text="Seleccionar Carpeta", style="TButton", command=seleccionar_carpeta_destino)
seleccionar_destino_button.grid(row=2, column=2, padx=10, pady=5)

# ComboBox para seleccionar la hoja de Excel
hoja_excel_label = tk.Label(ventana, text="Hoja de Excel:", font=fuente_grande, bg=color_fondo)
hoja_excel_label.grid(row=3, column=0, padx=10, pady=5, sticky="w")

hojas_excel = ["Hoja1", "Hoja2", "Hoja3", "Hoja4", "Hoja5", "Hoja6"]
hoja_excel_combo = ttk.Combobox(ventana, values=hojas_excel, font=fuente_grande)
hoja_excel_combo.grid(row=3, column=1, padx=10, pady=5, sticky="ew")
hoja_excel_combo.set(hojas_excel[0])  # Establece el valor predeterminado

# Botón para mover archivos
mover_button = ttk.Button(ventana, text="Mover Archivos", style="TButton", command=mover_archivos)
mover_button.grid(row=4, column=1, padx=10, pady=10)

# Botón de ayuda con un ícono de interrogación
ayuda_image = tk.PhotoImage(file="interrogation.png")  # Reemplaza "interrogation.png" con la ruta de tu imagen de interrogación
ayuda_button = tk.Button(ventana, image=ayuda_image, bg=color_fondo, command=mostrar_ayuda)
ayuda_button.grid(row=4, column=2, padx=10, pady=10)

# Etiqueta para mostrar el resultado
resultado_label = tk.Label(ventana, text="", wraplength=500, font=fuente_grande, bg=color_fondo)
resultado_label.grid(row=5, column=0, columnspan=3, padx=10, pady=10, sticky="w")

# Configurar opciones de expansión
ventana.columnconfigure(1, weight=1)
archivo_excel_entry.grid(sticky="ew")
carpeta_origen_entry.grid(sticky="ew")
carpeta_destino_entry.grid(sticky="ew")
hoja_excel_combo.grid(sticky="ew")
resultado_label.grid(sticky="ew")

ventana.mainloop()