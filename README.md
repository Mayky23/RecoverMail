# RecoverMail - Herramienta Forense de Recuperación de Correos Electrónicos

**Versión:** 3.0  
**Autor:** MARH  

---

## Descripción

RecoverMail es una herramienta forense diseñada para analizar y extraer información detallada de archivos MBOX, comúnmente utilizados para almacenar correos electrónicos. La herramienta permite generar informes en múltiples formatos (PDF, HTML y JSON) con el fin de facilitar la investigación forense y el análisis de datos.

Con RecoverMail, puedes:

- **Detectar automáticamente archivos MBOX**, incluso si no tienen extensión.
- **Extraer metadatos** como remitentes, destinatarios, asuntos, fechas y contenido de los correos.
- **Generar informes visuales y estructurados** que resumen el análisis forense.
- **Exportar resultados** en formatos compatibles con diferentes necesidades de análisis.

---

## Características principales

1. **Compatibilidad con MBOX:**
   - Soporte para archivos `.mbox` y archivos sin extensión detectados como MBOX.
   - Análisis robusto de metadatos y contenido de correos.

2. **Resúmenes interactivos:**
   - Informes en terminal con tablas claras y organizadas.
   - Exportación a PDF, HTML y JSON para compartir y documentar hallazgos.

3. **Extracción de detalles:**
   - Identificación de remitentes, destinatarios, fechas y asuntos.
   - Extracción del cuerpo del correo (texto plano) con manejo de errores.

4. **Interfaz amigable:**
   - Diseño visual con la biblioteca `rich` para mejorar la experiencia en terminal.
   - Informes generados con diseño limpio y profesional.

---

## Requisitos

Para ejecutar RecoverMail, necesitarás lo siguiente:

- **Python 3.8+**
- Bibliotecas requeridas (instala usando `pip install -r requirements.txt`):
  - `rich`
  - `mailbox`
  - `reportlab`
  - `argparse`

---

## Instalación

### 1. Clona este repositorio o descarga los archivos:
```bash
git clone https://github.com/Mayky23/RecoverMail.git
cd RecoverMail
```
### 2. Instala las dependencias:
```bash
pip install -r requirements.txt
```
### 3. Verifica la instalación ejecutando el comando de ayuda:
```bash
python recovermail.py --help
```

---

## Uso

### Sintaxis básica
```bash
python recovermail.py [archivos] [--output PREFIJO]
```
### Ejemplo de uso
Analiza un archivo MBOX y genera informes en PDF, HTML y JSON:

```bash
python recovermail.py correo.mbox --output resultados
```
Analiza múltiples archivos MBOX:

```bash
python recovermail.py correo1.mbox correo2.mbox --output analisis
```

---

## Opciones

| Opción      | Descripción                                      |
|-------------|--------------------------------------------------|
| `files`     | Archivos MBOX a analizar (uno o varios).         |
| `--output`  | Prefijo para los archivos de salida (opcional).  |

---

## Salida

RecoverMail genera los siguientes archivos:

| Tipo de archivo | Descripción                                                                 |
|-----------------|-----------------------------------------------------------------------------|
| **Informe PDF** | Resumen de archivos analizados y detalles de cada correo extraído.          |
| **Informe HTML**| Resumen interactivo con secciones desplegables para ver detalles de correos. |
| **Archivo JSON**| Datos completos en formato estructurado para análisis avanzado.             |