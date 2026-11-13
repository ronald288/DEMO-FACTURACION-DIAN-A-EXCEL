Este proyecto, desarrollado en Python, tiene como objetivo automatizar la lectura y el procesamiento de facturas electrónicas en formato DIAN, permitiendo a la empresa registrar de manera eficiente sus compras.

La aplicación cuenta con una interfaz gráfica construida con Tkinter, diseñada para ofrecer una experiencia intuitiva y accesible al usuario. Para la extracción de información, el sistema utiliza PDFPlumber, una librería especializada en la lectura de tablas dentro de documentos PDF. Una vez obtenida la información, el proyecto emplea OpenPyXL para registrar los datos en plantillas o libros de Excel, generando el formato de compras requerido por la empresa.

COMO FUNCIONA:

  1.  en la carpeta "facturas" se agregan las facturas descargadas de la DIAN con formato (dia/mes/año y nombre de la empresa) en su nombre de archivo.
  2.  ejecutar la aplicacion y dar click en el boton "procesar facturas".
  3.  finalizar en "OK" y cerrar la pestaña.
  4.  las facturas quedaran guardadas en el formato excel, después de tener los formatos completos. es necesario retirar estos de aquí e ingresar la plantilla de nuevo, para proximos registros.
