import tkinter as tk
from tkinter import ttk, messagebox
import threading


def cargar_modulos():
    global inicio_procesar_facturas
    try:
        from Libro_productos import inicio_procesar_facturas
    except Exception as e:
        print("Error al cargar módulos:", e)
        raise


#/////// CENTRAR VENTANA /////

def centrar_ventana(ventana, ancho, alto):
    ventana.update_idletasks()
    x = (ventana.winfo_screenwidth() // 2) - (ancho // 2)
    y = (ventana.winfo_screenheight() // 2) - (alto // 2)
    ventana.geometry(f"{ancho}x{alto}+{x}+{y}")


#//////INTERFAZ PRINCIPAL ////////

def crear_interfaz():
    ventana = tk.Tk()
    ventana.title("Procesador de Facturas")
    centrar_ventana(ventana, 700, 500)
    ventana.iconbitmap("logo.ico")

    # número con el cual se comenzara a registrar las facturas
    tk.Label(ventana, text="Ingrese el número de consecutivo inicial:").pack(pady=5)
    entrada = tk.Entry(ventana, width=20)
    entrada.pack()

    # Barra de progreso
    barra = ttk.Progressbar(ventana, orient="horizontal", length=500, mode="determinate")
    barra.pack(pady=10)

    texto = tk.Text(ventana, height=20, width=80)
    texto.pack(padx=10, pady=10)

    # Log
    def escribir_log(mensaje):
        texto.insert("end", mensaje + "\n")
        texto.see("end")
        ventana.update()

    # Barra
    def actualizar_barra(actual, total):
        barra["maximum"] = total
        barra["value"] = actual
        ventana.update()

    # Ejecutar proceso
    def ejecutar():
        try:
            consecutivo = int(entrada.get())
        except ValueError:
            messagebox.showerror("Error", "Ingrese un consecutivo válido.")
            return

        try:
            inicio_procesar_facturas(consecutivo, escribir_log, actualizar_barra)
            messagebox.showinfo("Finalizado", "Proceso terminado correctamente.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    tk.Button(ventana, text="Procesar facturas", command=ejecutar, width=25).pack(pady=10)

    ventana.mainloop()


# ///// SPLASH SCREEN /////

def abrir_interfaz():
    splash.destroy()
    crear_interfaz()


def iniciar_programa():
    try:
        cargar_modulos()
        splash.after(100, abrir_interfaz)
    except Exception as e:
        print("Error al iniciar programa:", e)
        splash.after(100, splash.destroy)


# Crear splash
splash = tk.Tk()
splash.overrideredirect(True)
centrar_ventana(splash, 300, 150)
tk.Label(splash, text="Cargando sistema...", font=("Arial", 12)).pack(expand=True)

threading.Thread(target=iniciar_programa, daemon=True).start()

splash.mainloop()
