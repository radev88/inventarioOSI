import tkinter as tk
import pandas as pd
from tkinter import messagebox
import os
import platform
import sys
from datetime import datetime

def obtener_ruta_archivo(nombre_archivo):
    if getattr(sys, 'frozen', False):
        ruta_base = os.path.dirname(sys.executable)
    else:
        ruta_base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(ruta_base, nombre_archivo)

archivo_excel = obtener_ruta_archivo("inventario_sistemas.xlsx")

# Función para cargar los datos desde el archivo Excel
def cargar_datos():
    if os.path.exists(archivo_excel):
        df = pd.read_excel(archivo_excel, engine='openpyxl', dtype=str)
        for col in ["Número de propiedad", "Número de serie", "Locación", "Decomisada", "Fecha de compra", "Descripción"]:
            if col not in df.columns:
                df[col] = ""
    else:
        df = pd.DataFrame(columns=["Número de propiedad", "Número de serie", "Locación", "Decomisada", "Fecha de compra", "Descripción"])
    return df

# Función para guardar los datos en el archivo Excel
def guardar_datos(df):
    df.to_excel(archivo_excel, index=False)

# Función para interpretar el estado de decomisado
def interpretar_decomisada(valor):
    return "Decomisada" if str(valor).strip().upper() == "X" else "No Decomisada"

# Función para verificar el estado de un equipo
def verificar_estado():
    num_propiedad = entry_num_propiedad.get().strip().upper()
    num_serie = entry_num_serie.get().strip().upper()

    if not num_propiedad and not num_serie:
        messagebox.showwarning("Advertencia", "Por favor, ingrese el número de propiedad o número de serie.")
        return

    df = cargar_datos()
    df["Número de propiedad"] = df["Número de propiedad"].astype(str).str.strip().str.upper()
    df["Número de serie"] = df["Número de serie"].astype(str).str.strip().str.upper()

    encontrado = None

    if num_propiedad:
        encontrado = df[df["Número de propiedad"] == num_propiedad]
    elif num_serie:
        encontrado = df[df["Número de serie"] == num_serie]

    if not encontrado.empty:
        fila = encontrado.iloc[0]

        decomisada_raw = fila["Decomisada"]
        decomisada_interpretada = interpretar_decomisada(decomisada_raw)
        locacion = fila.get("Locación", "No especificado") or "No especificado"
        fecha_compra = fila.get("Fecha de compra", "No especificado")
        descripcion = fila.get("Descripción", "No especificado") or "No especificado"

        messagebox.showinfo(
            "Estado",
            f"Computadora encontrada:\n"
            f"Número de Propiedad: {fila['Número de propiedad']}\n"
            f"Número de Serie: {fila['Número de serie']}\n"
            f"Estado: {decomisada_interpretada}\n"
            f"Locación: {locacion}\n"
            f"Fecha de compra: {fecha_compra}\n"
            f"Descripción: {descripcion}."
        )

        if decomisada_interpretada == "Decomisada":
            var_decomisada.set(1)
            var_no_decomisada.set(0)
        else:
            var_decomisada.set(0)
            var_no_decomisada.set(1)
    else:
        messagebox.showinfo("Estado", "No se encontró la computadora.")
        var_decomisada.set(0)
        var_no_decomisada.set(0)

# Función para cambiar el estado de decomisado
def on_checkbox_change(estado):
    num_propiedad = entry_num_propiedad.get().strip().upper()
    num_serie = entry_num_serie.get().strip().upper()

    df = cargar_datos()
    df["Número de propiedad"] = df["Número de propiedad"].astype(str).str.strip().str.upper()
    df["Número de serie"] = df["Número de serie"].astype(str).str.strip().str.upper()

    encontrado = None
    if num_propiedad:
        encontrado = df[df["Número de propiedad"] == num_propiedad]
    elif num_serie:
        encontrado = df[df["Número de serie"] == num_serie]

    if not encontrado.empty:
        index = encontrado.index[0]
        if estado == "Decomisada":
            df.at[index, "Decomisada"] = "X"
            var_decomisada.set(1)
            var_no_decomisada.set(0)
        else:
            df.at[index, "Decomisada"] = ""
            var_decomisada.set(0)
            var_no_decomisada.set(1)
        guardar_datos(df)
    else:
        var_decomisada.set(0)
        var_no_decomisada.set(0)

# Función para agregar un nuevo registro
def agregar_registro():
    num_propiedad = entry_num_propiedad.get().strip().upper()
    num_serie = entry_num_serie.get().strip().upper()
    locacion = entry_locacion.get().strip()
    decomisada = "X" if var_decomisada.get() == 1 else ""
    fecha_compra = entry_fecha_compra.get().strip()
    descripcion = entry_descripcion.get().strip()

    if not num_propiedad and not num_serie:
        messagebox.showwarning("Advertencia", "Debe ingresar el número de propiedad o número de serie.")
        return
    if not fecha_compra:
        messagebox.showwarning("Advertencia", "Debe ingresar la fecha de compra.")
        return

    try:
        datetime.strptime(fecha_compra, "%Y-%m-%d")
    except ValueError:
        messagebox.showwarning("Advertencia", "La fecha de compra debe estar en formato YYYY-MM-DD.")
        return

    df = cargar_datos()

    if ((num_propiedad and (num_propiedad in df["Número de propiedad"].values)) or 
        (num_serie and (num_serie in df["Número de serie"].values))):
        messagebox.showerror("Error", "El número de propiedad o número de serie ya existe en el inventario.")
    else:
        nuevo_registro = {
            "Número de propiedad": num_propiedad,
            "Número de serie": num_serie,
            "Locación": locacion,
            "Decomisada": decomisada,
            "Fecha de compra": fecha_compra,
            "Descripción": descripcion
        }
        df = pd.concat([df, pd.DataFrame([nuevo_registro])], ignore_index=True)
        guardar_datos(df)
        messagebox.showinfo("Éxito", "Computadora agregada exitosamente.")
        limpiar_campos()

# Función para limpiar los campos después de agregar
def limpiar_campos():
    entry_locacion.delete(0, tk.END)
    entry_fecha_compra.delete(0, tk.END)
    entry_descripcion.delete(0, tk.END)
    var_decomisada.set(0)
    var_no_decomisada.set(0)

# Función para editar el estado de un registro existente
def editar_estado():
    num_propiedad = entry_num_propiedad.get().strip().upper()
    num_serie = entry_num_serie.get().strip().upper()
    nueva_locacion = entry_locacion.get().strip()
    nueva_descripcion = entry_descripcion.get().strip()
    decomisada = "X" if var_decomisada.get() == 1 else ""

    if not num_propiedad and not num_serie:
        messagebox.showwarning("Advertencia", "Debe ingresar el número de propiedad o número de serie para editar.")
        return

    df = cargar_datos()
    df["Número de propiedad"] = df["Número de propiedad"].astype(str).str.strip().str.upper()
    df["Número de serie"] = df["Número de serie"].astype(str).str.strip().str.upper()

    encontrado = None
    if num_propiedad:
        encontrado = df[df["Número de propiedad"] == num_propiedad]
    elif num_serie:
        encontrado = df[df["Número de serie"] == num_serie]

    if not encontrado.empty:
        index = encontrado.index[0]
        if nueva_locacion:
            df.at[index, "Locación"] = nueva_locacion
        if nueva_descripcion:
            df.at[index, "Descripción"] = nueva_descripcion
        df.at[index, "Decomisada"] = decomisada

        guardar_datos(df)
        messagebox.showinfo("Éxito", "Registro actualizado exitosamente.")
        limpiar_campos()
    else:
        messagebox.showerror("Error", "No se encontró el registro.")

# Función para abrir el archivo Excel directamente
def abrir_excel():
    sistema = platform.system()
    if sistema == "Darwin":
        os.system(f"open '{archivo_excel}'")
    elif sistema == "Windows":
        os.startfile(archivo_excel)
    else:
        os.system(f"xdg-open '{archivo_excel}'")

# ------------------------------------------------------------
# Creación de la ventana principal con interfaz gráfica (GUI)
# ------------------------------------------------------------

root = tk.Tk()
root.title("Sistema de Inventario (OSI)")
root.configure(padx=20, pady=20)

# Inicializar las variables de los checkboxes
var_decomisada = tk.IntVar()
var_no_decomisada = tk.IntVar()

# Etiqueta y campo de número de propiedad
label_num_propiedad = tk.Label(root, text="Número de Propiedad:")
label_num_propiedad.grid(row=0, column=0, sticky="w")
entry_num_propiedad = tk.Entry(root, width=35)
entry_num_propiedad.grid(row=1, column=0, pady=(0, 10))

# Etiqueta y campo de número de serie
label_num_serie = tk.Label(root, text="Número de Serie:")
label_num_serie.grid(row=0, column=1, sticky="w")
entry_num_serie = tk.Entry(root, width=35)
entry_num_serie.grid(row=1, column=1, pady=(0, 10))

# Etiqueta y campo de locación
label_locacion = tk.Label(root, text="Locación:")
label_locacion.grid(row=2, column=0, sticky="w")
entry_locacion = tk.Entry(root, width=35)
entry_locacion.grid(row=3, column=0, columnspan=2, pady=(0, 10))

# Etiqueta y campo de fecha de compra
label_fecha_compra = tk.Label(root, text="Fecha de Compra (YYYY-MM-DD):")
label_fecha_compra.grid(row=4, column=0, sticky="w")
entry_fecha_compra = tk.Entry(root, width=35)
entry_fecha_compra.grid(row=5, column=0, columnspan=2, pady=(0, 10))

# Etiqueta y campo de descripción
label_descripcion = tk.Label(root, text="Descripción:")
label_descripcion.grid(row=6, column=0, sticky="w")
entry_descripcion = tk.Entry(root, width=35)
entry_descripcion.grid(row=7, column=0, columnspan=2, pady=(0, 10))

# Frame para los checkboxes
checkbox_frame = tk.Frame(root)
checkbox_frame.grid(row=8, column=0, columnspan=2, pady=10)

checkbox_decomisada = tk.Checkbutton(checkbox_frame, text="Decomisada", variable=var_decomisada,
                                      command=lambda: on_checkbox_change("Decomisada"))
checkbox_decomisada.pack(side="left", padx=10)

checkbox_no_decomisada = tk.Checkbutton(checkbox_frame, text="No Decomisada", variable=var_no_decomisada,
                                        command=lambda: on_checkbox_change("No Decomisada"))
checkbox_no_decomisada.pack(side="left", padx=10)

# Botón para verificar estado
boton_verificar = tk.Button(root, text="Verificar Estado", command=verificar_estado, width=25)
boton_verificar.grid(row=9, column=0, pady=10)

# Botón para agregar nuevo registro
boton_agregar = tk.Button(root, text="Agregar Registro", command=agregar_registro, width=25)
boton_agregar.grid(row=9, column=1, pady=10)

# Botón para editar registro existente
boton_editar = tk.Button(root, text="Editar Registro", command=editar_estado, width=25)
boton_editar.grid(row=10, column=0, columnspan=2, pady=10)

# Botón para abrir el archivo Excel
boton_abrir_excel = tk.Button(root, text="Abrir Archivo Excel", command=abrir_excel, width=25)
boton_abrir_excel.grid(row=11, column=0, columnspan=2, pady=20)

# Ejecutar la aplicación
root.mainloop()