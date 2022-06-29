# Mover archivos adjunto de Outlook
Podemos mover los archivos adjuntos en Outlook a una carpeta en especifica

## 1. Debemos habilitar "Macros" en Outlook

- Ir a Archivo->Opciones
- En la ventana de "Opciones de outlook", ir "Centro de confianza" 
image.png

- Hacer click en el boton de "Configuracion de Centro de Confianza"
- En la ventana de "Centro de confianza", ir a "Configuración de macros"
image.png
- Seleccionar la opcion de "Habilitar todas las macros"
- Guardar los cambios

## Visual Basic de Outlook
- Presionamos las teclas de Alt+F11 en Outlook y nos vamos a la vista de desarrollo de Outlook
- Hacemos clic dentro de Proyecto1 (VbaOroject.OTM)->Microsoft Outlook Objetos->ThisOutlookSession
- Ahi copiar el codigo VBA del repositorio "save_Attchments.vba"