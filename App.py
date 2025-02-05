
# Interfaz de un Excel básico

import tkinter as tk
from tkinter import ttk, font, messagebox, filedialog
import csv

ventana = tk.Tk()
ventana.title("Basic Excel")
ventana.geometry("1100x600")
ventana.resizable(0, 0)
ventana.iconbitmap("Logo.ico")

# Título
ttk.Label(ventana, text = "Welcome To Basic Excel", font = ("Arial", 18)).pack(pady = 30)

def cerrar_ventana():
    if messagebox.askokcancel("Exit", "Are you sure you want to leave?"):
        ventana.destroy()

# Lista global para almacenar todas las entradas
todas_las_entradas = []

def crear_celdas(texto, posicion_x, y):
    """
    Crear una fila de celdas con entradas.
    """
    tk.Label(ventana, text = texto, background = "#74c69d", 
        foreground = "#1b4332", padx = 20).place(x = 0, y = y)

    fila_entradas = []
    for x in posicion_x:
        entrada = ttk.Entry(ventana)
        entrada.place(x = x, y = y)
        fila_entradas.append(entrada)

    tk.Label(ventana, text = texto, background = "#74c69d", 
        foreground = "#1b4332", padx = 20).place(x = 1050, y = y)
    todas_las_entradas.append(fila_entradas)

def guardar_como_csv():
    """
    Guardar el contenido de las entradas en un archivo CSV.
    """
    archivo = filedialog.asksaveasfilename(defaultextension = ".csv", filetypes = [("Files CSV", "*.csv")])
    if not archivo:
        return # Si no se selecciona un archivo, salir de la función

    # Recoger los datos de todas las entradas
    datos = [[entrada.get() for entrada in fila] for fila in todas_las_entradas]

    # Guardar los datos en el archivo CSV
    with open(archivo, mode = "w", newline = "", encoding = "utf-8") as archivo_csv:
        escritor = csv.writer(archivo_csv)
        escritor.writerow(["Column " + c for c in columnas]) # Escribir encabezados
        escritor.writerows(datos) # Escribir datos

    messagebox.showinfo("Saved", "Data saved successfully")

def aplicar_operaciones():
    operacion = operacion_seleccionada.get()
    if operacion == "":
        messagebox.showerror("Error", "Please select an operation")
        return

    try:
        # Obtener índices globales desde los OptionMenu
        idx1, idx2 = int(opcion1.get()), int(opcion2.get())

        # Calcular fila y columna para cada índice
        fila1, col1 = idx1 // 8, idx1 % 8
        fila2, col2 = idx2 // 8, idx2 % 8

        # Verificar si los índices son válidos dentro de todas_las_entradas
        if fila1 >= len(todas_las_entradas) or fila2 >= len(todas_las_entradas):
            messagebox.showerror("Error", "Invalid cell selection")
            return

        valores = [[float(entrada.get()) if entrada.get() else 0 for entrada in fila] for fila in todas_las_entradas]

        # Obtener los valores correspondientes
        valor1, valor2 = valores[fila1][col1], valores[fila2][col2]

        # Realizar operación seleccionada
        if operacion == "Addition":
            resultado = valor1 + valor2
        elif operacion == "Subtraction":
            resultado = valor1 - valor2
        elif operacion == "Multiplication":
            resultado = valor1 * valor2
        elif operacion == "Division":
            if valor2 == 0:
                messagebox.showerror("Error", "Cannot divide by zero")
                return
            resultado = valor1 / valor2
        else:
            resultado = 0

        messagebox.showinfo("Result", f"The result is {resultado}")

    except ValueError:
        messagebox.showerror("Error", "Please enter valid numbers")

def aplicar_fuentes():
    messagebox.showinfo("Fonts", "Font style changed")
    for fila in todas_las_entradas:
        for entrada in fila:
            entrada.configure(font = (fuente_seleccionada.get(), 12))

# Crear la barra de menú
barra_menu = tk.Menu(ventana)

def agregar_menu(menu, etiqueta, funcion):
    menu.add_command(label = etiqueta, command = funcion)

# Agregar el menú Archivo
menu_archivo = tk.Menu(barra_menu, tearoff = 0)
agregar_menu(menu_archivo, "Save as CSV", guardar_como_csv)
barra_menu.add_cascade(label = "File", menu = menu_archivo)

# Agregar el menú Operaciones
menu_operaciones = tk.Menu(barra_menu, tearoff = 0)
operaciones = ["Addition", "Subtraction", "Multiplication", "Division"]
operacion_seleccionada = tk.StringVar()

for o in operaciones:
    menu_operaciones.add_radiobutton(label = o, variable = operacion_seleccionada, value = o, command = aplicar_operaciones)
barra_menu.add_cascade(label = "Operations", menu = menu_operaciones)

# Agregar el menú Fuentes
menu_fuentes = tk.Menu(barra_menu, tearoff = 0)
fuentes = font.families()[:10]
fuente_seleccionada = tk.StringVar()

for f in fuentes:
    menu_fuentes.add_radiobutton(label = f, variable = fuente_seleccionada, value = f, command = aplicar_fuentes)
barra_menu.add_cascade(label = "Fonts", menu = menu_fuentes)

# Crear dos menús desplegables para seleccionar las celdas para las operaciones
opcion1 = tk.IntVar(value = 0)
opcion2 = tk.IntVar(value = 1)

menu1 = tk.OptionMenu(ventana, opcion1, *map(str, range(175)))
menu1.place(x = 47, y = 52)

menu2 = tk.OptionMenu(ventana, opcion2, *map(str, range(175)))
menu2.place(x = 100, y = 52)

# Columnas
columnas = ["A", "B", "C", "D", "E", "F", "G", "H"]
posicion_x = [50, 175, 300, 425, 550, 674, 798, 924]

for c, x in zip(columnas, posicion_x):
    tk.Label(ventana, text = c, background = "#74c69d", foreground = "#1b4332", padx = 57).place(x = x, y = 80)

# Crear filas de celdas
for i, y in enumerate(range(100, 540, 20), start = 1):
    crear_celdas(str(i), posicion_x, y)

ventana.protocol("WM_DELETE_WINDOW", cerrar_ventana)
ventana.config(menu = barra_menu)
ventana.mainloop()