# Automatizar_Formatos
Auto-matizador y organizador de imagenes de archivos, plataforma para la automatizacion de los mismos
Aquí tienes un archivo `README.md` estructurado y redactado con un enfoque de Ingeniería de Software, ideal para presentarlo en tu portafolio, en GitHub, o como documentación oficial para tu proyecto en la ESIME.

---

#  Generador Multi-Reportes Operativos | SecureCorp

**Desarrollador:** Erick Daniel Azamar Silva

**Institución:** Instituto Politécnico Nacional – ESIME Culhuacán

**Versión:** 2.0 (Arquitectura Cliente-Servidor)

Un *middleware* de alta eficiencia diseñado para centros de monitoreo (CCTV y Rastreo GPS). Actúa como puente entre el analista humano y las plantillas operativas de Excel, automatizando el llenado de checklists, la recolección de evidencias fotográficas desde el portapapeles y el cálculo matemático de proporciones (Aspect Ratio) para inyectar imágenes sin distorsión.

##  Características Principales

* **Motor de Inyección Profunda (COM):** Utiliza la API nativa de Windows (`win32com`) para operar Microsoft Excel en segundo plano (modo fantasma), asegurando que los reportes generados mantengan el 100% de la fidelidad del archivo original.
* **Auto-Escalado Matricial Inteligente:** El backend de Python (con Pillow) analiza la resolución de cada imagen subida, detecta si es panorámica, cuadrada o "tira de texto", y crea un lienzo calculado matemáticamente para que encaje a la perfección en las celdas de destino sin dejar espacios en blanco.
* **Persistencia Local (Zero-Data-Loss):** El frontend está blindado con caché de `localStorage`. Si el navegador se cierra o recarga por accidente, toda la información de los formularios e imágenes se recupera instantáneamente.
* **Descargas Flexibles:** Permite exportar los resultados simultáneamente en formatos de alta resolución (.PDF) y documentos editables (.XLSX) mediante codificación Base64.
* **Interfaz Glassmorphism Dark UI:** Diseño estético y enfocado en la usabilidad nocturna, con componentes interactivos construidos 100% en vectores SVG (cero carga de imágenes externas).

##  Requisitos del Sistema

Debido al uso de librerías de interoperabilidad, este software requiere un entorno operativo específico:

* **Sistema Operativo:** Windows 10 o Windows 11 (Obligatorio para la librería COM).
* **Software Base:** Microsoft Excel instalado y activado localmente.
* **Entorno:** Python 3.8 o superior.

##  Instalación y Estructura

**1. Preparación del Directorio:**
Tu proyecto debe respetar rigurosamente la siguiente jerarquía de carpetas requerida por el framework Flask:

```text
📁 Proyecto_SecureCorp
 ├── 📄 app.py                             <-- Backend (Servidor API)
 ├── 📄 Formato Validación movil (1).xlsx  <-- Plantilla Excel Base 1
 ├── 📄 Formato validacion torres.xlsx     <-- Plantilla Excel Base 2
 ├── 📁 templates
 │    └── 📄 index.html                    <-- Frontend (UI / Vista Web)
 └── 📁 static
      └── 🖼️ SecureCorp.png                <-- Logo Corporativo

```

**2. Instalación de Dependencias:**
Abre una terminal (CMD o PowerShell) en la carpeta raíz del proyecto y ejecuta:

```bash
pip install flask flask-cors openpyxl pillow pywin32

```

##  Manual de Operación

1. **Arranque del Motor:** Ejecuta `python app.py` en la consola. La terminal no debe cerrarse durante el uso de la plataforma.
2. **Acceso a la Interfaz:** Ingresa a `[http://127.0.0.1:5000](http://127.0.0.1:5000)` en cualquier navegador web.
3. **Captura de Evidencias:** Utiliza el comando nativo de Windows (`Win + Shift + S`) para recortar pantallas de los sistemas de monitoreo y presiona el botón "Pegar foto" en los paneles de la plataforma web.
4. **Generación:** Selecciona mediante los *toggles* iluminados qué formato deseas descargar (PDF, Excel o ambos) y procesa el reporte. El servidor nombrará los archivos generados con el formato dinámico `MATRICULA_DD-MES-YYYY`.

## Resolución de Problemas Comunes (Troubleshooting)

| Código de Error | Causa Raíz | Solución Rápida |
| --- | --- | --- |
| **`Permission denied`** | Una de las plantillas originales de Excel (.xlsx) se encuentra abierta en el escritorio del usuario, por lo que Windows la bloquea. | Cierra la plantilla original en Excel y vuelve a procesar. |
| **`OLE error 0x800a01a8`** | Un proceso de Microsoft Excel se quedó "congelado" (Zombie) en la memoria RAM tras un cierre inesperado o crash del código. | Abre la terminal y ejecuta el comando de limpieza profunda: `taskkill /F /IM EXCEL.EXE` |
| **El servidor no arranca** | Los puertos locales (5000) pueden estar en uso por otro proceso web de Windows. | Modifica el puerto de arranque al final de `app.py` a `port=5001`. |
