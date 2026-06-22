import os
import sys
from datetime import datetime
from Obtener_informacion_PDF import extraer_datos_factura
from Facturas_compras import llenar_facturas_compras


def obtener_cliente_excel():
   import win32com.client as win32
   return win32.DispatchEx("Excel.Application")

# Ruta del archivo Excel
def obtener_ruta_base():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

BASE_DIR = obtener_ruta_base()

archivo_libro_productos = os.path.join(BASE_DIR, "Registro_total_productos.xlsx")
archivo_facturas_compras = os.path.join(BASE_DIR, "Registro_facturas_compras.xlsx")
carpeta_pdfs = os.path.join(BASE_DIR, "facturas")

#validar si existen las rutas dentro de la carpeta de origen
def validar_existencia(ruta, tipo):
    if not os.path.exists(ruta):
        raise Exception(f"No se encontró el {tipo}:\n{ruta}")
    
#validar si existe la hoja PLANTILLA en el archivo
def validar_plantilla(libro, nombre="PLANTILLA"):
    try:
        libro.Worksheets(nombre)
    except:
        raise Exception(f"No existe la hoja '{nombre}' en el archivo.")


####### FACTURAS A PROCESAR #########

def obtener_lista_pdfs(carpeta):
    return sorted([f for f in os.listdir(carpeta) if f.lower().endswith(".pdf")])


#### REGISRTO EN EL EXCEL ######

# Obtener el consecutivo automáticamente a partir del nombre de la última hoja del libro.”
def obtener_consecutivo(libro, consecutivo_inicial):
    consecutivos = [
        int(ws.Name)
        for ws in libro.Worksheets
        if ws.Name.isdigit() and ws.Name.upper() != "PLANTILLA"
    ]
    return consecutivo_inicial if not consecutivos else max(consecutivos) + 1


#Verificar si ya existe un registro previo o si este es el primero.”
def obtener_ultima_hoja(libro):
    consecutivos = [
        int(ws.Name)
        for ws in libro.Worksheets
        if ws.Name.isdigit() and ws.Name.upper() != "PLANTILLA"
    ]
    return max(consecutivos) if consecutivos else None


#Contar cuántas celdas están llenas dentro del rango permitido (15 productos por registro)
def contar_productos(hoja):
    contador = 0
    for fila in range(19, 34):
        valor = hoja.Cells(fila, 1).Value
        if valor is not None and str(valor).strip() != "":
            contador += 1
    return contador


# Determinar el espacio disponible del registro actual
def decision_sig_factura(hoja, cantidad):
    usados = contar_productos(hoja)
    return (usados + cantidad) <= 15


def obtener_primera_fila_vacia(hoja):
    for fila in range(19, 34):
        if hoja.Cells(fila, 1).Value is None:
            return fila
    return None


#registrar los productos en sus respectivas celdas
def registrar_productos(hoja, factura, fecha_obj):
    fila = obtener_primera_fila_vacia(hoja)
    item = contar_productos(hoja) +1


    for producto in factura["productos"]:

        hoja.Range(f"A{fila}").Value = item
        hoja.Range(f"B{fila}").Value = producto["cantidad"]
        hoja.Range(f"D{fila}").Value = producto["descripcion"]
        hoja.Range(f"J{fila}").Value = fecha_obj
        hoja.Range(f"M{fila}").Value = fecha_obj
        hoja.Range(f"N{fila}").Value = producto["cantidad"]
        hoja.Range(f"O{fila}").Value = "X"
        hoja.Range(f"Q{fila}").Value = "X"

        fila += 1
        item += 1

def registrar_facturas(libro, libro2, datos_factura, consecutivo_inicial):

    fecha_obj = datetime.strptime(datos_factura["fecha"], "%d/%m/%Y")
    cantidad_productos = len(datos_factura["productos"])

    ultima = obtener_ultima_hoja(libro)
    hoja_actual = libro.Worksheets(str(ultima)) if ultima else None

    # Determinar si usar hoja existente o crear nueva
    if hoja_actual and decision_sig_factura(hoja_actual, cantidad_productos):
        hoja = hoja_actual
        consecutivo = int(hoja.Name)
    else:
        consecutivo = obtener_consecutivo(libro, consecutivo_inicial)
        plantilla = libro.Worksheets("PLANTILLA")
        plantilla.Copy(Before=None, After=libro.Worksheets(libro.Worksheets.Count))

        hoja = libro.Worksheets(libro.Worksheets.Count)
        hoja.Name = str(consecutivo)

        hoja.Range("D8").Value = fecha_obj
        hoja.Range("Q8").Value = consecutivo

    registrar_productos(hoja, datos_factura, fecha_obj)
    llenar_facturas_compras(libro2, datos_factura, consecutivo)


def inicio_procesar_facturas(consecutivo, escribir_log, actualizar_barra):

    validar_existencia(archivo_libro_productos, "archivo de productos")
    validar_existencia(archivo_facturas_compras, "archivo de facturas")
    validar_existencia(carpeta_pdfs, "carpeta de facturas PDF")

    lista_pdfs = obtener_lista_pdfs(carpeta_pdfs)
    if not lista_pdfs:
        raise Exception("No se encontraron archivos PDF en la carpeta 'facturas'.")

    excel = obtener_cliente_excel()
    excel.Visible = False
    excel.DisplayAlerts = False

    libro1 = excel.Workbooks.Open(archivo_libro_productos, ReadOnly=False)
    libro2 = excel.Workbooks.Open(archivo_facturas_compras, ReadOnly=False)

    validar_plantilla(libro1)
    validar_plantilla(libro2)

    actualizar_barra(0, len(lista_pdfs))

    try:
        for i, nombre_pdf in enumerate(lista_pdfs):
            escribir_log(f"Procesando: {nombre_pdf}")

            ruta_pdf = os.path.join(carpeta_pdfs, nombre_pdf)
            datos_factura = extraer_datos_factura(ruta_pdf)

            registrar_facturas(libro1, libro2, datos_factura, consecutivo)

            actualizar_barra(i + 1, len(lista_pdfs))

        libro1.Save()
        libro2.Save()
        escribir_log("Proceso terminado correctamente.")

    except Exception as e:
        escribir_log(f"ERROR: {str(e)}")
        raise

    finally:
        libro1.Close(SaveChanges=True)
        libro2.Close(SaveChanges=True)
        excel.Quit()