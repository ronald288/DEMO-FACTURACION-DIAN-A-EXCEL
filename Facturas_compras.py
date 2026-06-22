def llenar_facturas_compras(libro, datos_factura, consecutivo):

    libro.Activate()

    # Copiar plantilla
    plantilla = libro.Worksheets("PLANTILLA")
    plantilla.Copy(
        Before=None,
        After=libro.Worksheets(libro.Worksheets.Count)
    )

    # Obtener la hoja recién creada
    hoja = libro.Worksheets(libro.Worksheets.Count)

    # Intentar asignar nombre único
    base_nombre = str(consecutivo)
    nombre_final = base_nombre
    contador = 2

    # Mientras exista una hoja con ese nombre, agregar sufijo (2), (3), ...
    while any(ws.Name == nombre_final for ws in libro.Worksheets):
        nombre_final = f"{base_nombre}({contador})"
        contador += 1

    # Renombrar la hoja con el nombre libre
    hoja.Name = nombre_final


    #///// LLENAR DATOS GENERALES /////
 
    hoja.Range("K4").Value = consecutivo
    hoja.Range("C1").Value = datos_factura["razon_social"]
    hoja.Range("C2").Value = datos_factura["nit"]
    hoja.Range("C3").Value = datos_factura["direccion"]
    hoja.Range("K1").Value = datos_factura["telefono"]
    hoja.Range("K2").Value = datos_factura["correo"]

    # Fecha
    dia, mes, anio = datos_factura["fecha"].split("/")

    hoja.Range("C4").Value = int(dia)
    hoja.Range("D4").Value = int(mes)
    hoja.Range("E4").Value = int(anio)

    hoja.Range("E27").Value = int(dia)
    hoja.Range("G27").Value = int(mes)
    hoja.Range("I27").Value = int(anio)

    # IVA
    hoja.Range("K23").Value = datos_factura["iva"]

    # Productos
    fila = 7

    for producto in datos_factura["productos"][:15]:
        hoja.Range(f"A{fila}").Value = producto["cantidad"]
        hoja.Range(f"C{fila}").Value = producto["descripcion"]
        hoja.Range(f"C{fila}").WrapText = False
        hoja.Range(f"J{fila}").Value = producto["precio_unitario"]

        fila += 1