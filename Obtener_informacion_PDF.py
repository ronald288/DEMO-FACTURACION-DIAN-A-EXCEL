import pdfplumber
import re

def limpiar_numero(valor):
    if not valor:
        return 0
    return float(valor.replace("$", "").replace(".", "").replace(",", "."))


def extraer_datos_factura(ruta_pdf):

    datos = {
        "fecha": None,
        "razon_social": None,
        "nit": None,
        "direccion": None,
        "telefono": None,
        "correo": None,
        "iva": 0,
        "productos": []
    }

    with pdfplumber.open(ruta_pdf) as pdf:
        texto = ""
        for pagina in pdf.pages:
            texto += pagina.extract_text() + "\n"


        # /////// FECHA /////
        m = re.search(r"Fecha de Emisión:\s*(\d{2}/\d{2}/\d{4})", texto)
        if m:
            datos["fecha"] = m.group(1)


        # ////// RAZÓN SOCIAL /////
        m = re.search(r"Razón Social:\s*(.+)", texto)
        if m:
            datos["razon_social"] = m.group(1).strip()


        # /////// NIT //////
        m = re.search(r"Nit del Emisor:\s*([0-9]+)", texto)
        if m:
            datos["nit"] = m.group(1)

  
        # ////// DIRECCIÓN ////
        m = re.search(r"Dirección:\s*:?\s*(.+)", texto)
        if m:
            datos["direccion"] = m.group(1).strip()


        # ///// TELÉFONO ////
        m = re.search(r"Teléfono\s*/\s*Móvil:\s*([0-9]+)", texto)
        if m:
            datos["telefono"] = m.group(1)

     
        # ///// CORREO ////
        m = re.search(r"[A-Za-z0-9\._%+-]+@[A-Za-z0-9\.-]+\.[A-Za-z]{2,}", texto)
        if m:
            datos["correo"] = m.group(0)


        # ///// PRODUCTOS (TABLA) /////
        # productos registrados en la factura 

        for pagina in pdf.pages:
            tablas = pagina.extract_tables()
            for tabla in tablas:
                for fila in tabla:
                    try:
                        nro = fila[0]
                        descripcion = fila[2]
                        cantidad = fila[4]
                        precio_unitario = fila[5]

                        # Solo filas válidas
                        if nro and nro.isdigit():
                            datos["productos"].append({
                                "descripcion": descripcion.strip(),
                                "cantidad": limpiar_numero(cantidad),
                                "precio_unitario": limpiar_numero(precio_unitario)
                            })
                    except:
                        continue

        # ///// IVA /////

        # Busca el bloque de texto "MONEDA COP" para seleccionar el valor del IVA correcto
        bloque_cop = re.search(r"MONEDA\s+COP(.*?)(?:RETENCIONES|Numero de Autorización)", texto, re.DOTALL)

        if bloque_cop:
        # Buscar el IVA dentro de ese bloque
         matches = re.findall(r"IVA[:\s]*\$?([\d.,]+)", bloque_cop.group(1))
         if matches:
          datos["iva"] = limpiar_numero(matches[-1])
         else:
          datos["iva"] = 0.0
        else:
         datos["iva"] = 0.0

    return datos
