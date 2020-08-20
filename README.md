# Ping_to_a_Batch_of_Hosts_Windows

El programa se ha creado para Windows en Español, si lo tienes en ingles hay que cambiar en el archivo de código
las palabra 'Respuesta' por 'Replay' y 'desde' por 'for'. Lo mismo con otro idioma.

Se ha creado un programa en VBS (visual basic script) que hace ping a una lista de hosts, los cuales los tenemos 
identificados por el nombre, no por IP (se podría modificar el programa fácilmente para poder hacerlo por el IP),
y se guarda en un Excel en el que hay tres columnas Hostname, IP, Status (El Excel se crea automáticamente y se
va rellenando cuando se abre).

El archivo con los HostNames se debe de llamar impLis.Txt (se puede cambiar el código), la estructura es simple,
debe de tener un hostname por línea (Un IP por línea en el caso de que quieras modificarlo para IP). Y debe de estar
en la misma ruta que el código.

Para ejecutar el programa abrir CMD, ir a la carpeta donde esta el archivo, una vez aquí ponemos:
```console
    cscript batchPing.vbs
```
El comando cscript sirve para ejecutar este tipo de archivos en CMD, y el segundo argumento es el nombre del archivo
con el código.

Si tienes alguna duda déjame un comentario en la pestaña issues o mándame un correo.
**No haría falta modificar nada para hacerlo por los IP's en vez de con los HosNames, solo ponerlos en el archivo impLis.Txt
