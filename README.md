# RecoverMail — MBOX Forensics Toolkit 📨

Suite forense para **analizar archivos MBOX** (aunque no tengan extensión), extraer metadatos, cuerpos y adjuntos (solo metadatos), y generar informes **PDF / HTML / JSON** sin modificar el original.

> Pensado para respuesta a incidentes, e-discovery, auditorías y análisis post-mortem.

---

## Características

- Detección de MBOX por firma (`From `) y extensiones comunes.
- Extracción robusta de:
  - `From / To / Cc / Bcc`
  - `Subject` (decodifica MIME headers)
  - `Date` (normaliza a **UTC ISO 8601** cuando se puede)
  - `Message-ID`
  - `Body` (`text/plain` o `text/html` convertido a texto)
  - Adjuntos (**solo metadatos**: nombre, tipo, tamaño)
- Métricas útiles:
  - Top remitentes/destinatarios/asuntos/dominios
  - Duplicados por `sha256` del cuerpo (si se incluye body)
  - Conteo de adjuntos
- Informes:
  - **HTML** con búsqueda y detalles desplegables
  - **JSON** estructurado (ideal para automatizar)
  - **PDF** con resumen y tablas

---

## Requisitos

- Python **3.9+** (recomendado 3.11+)
- Dependencias:
  - `rich`
  - `reportlab` (solo para exportar PDF)

Instalación rápida:

```bash
pip install rich reportlab
```

---

## Uso

### Analizar un MBOX y generar informes

```bash
python recovermail.py correo.mbox -o informe
```

Genera (por defecto) en el directorio actual:

- `informe.html`
- `informe.json`
- `informe.pdf`

### Analizar varios archivos

```bash
python recovermail.py correo1.mbox correo2.mbox -o caso_001 --outdir resultados
```

### Analizar una carpeta (y subcarpetas)

```bash
python recovermail.py ./evidencias_mail/ --recursive -o caso_002 --outdir resultados
```

---

## Opciones CLI

- `-o, --output`: prefijo de salida (sin extensión)
- `--outdir`: carpeta de salida (se crea si no existe)
- `--recursive`: buscar MBOX dentro de subcarpetas
- `--max-body-chars`: límite de caracteres del body en HTML/JSON (`0` = sin límite)
- `--top`: tamaño de listas “Top” (remitentes/asuntos/dominios)
- `--no-body`: **no** extraer body (más rápido y ligero)
- `--prefer-html`: prioriza `text/html` convertido a texto sobre `text/plain`
- `--no-html`, `--no-json`, `--no-pdf`: desactivar salidas

Ejemplo “solo JSON, sin body”:

```bash
python recovermail.py correo.mbox --no-body --no-html --no-pdf -o salida --outdir out
```

---

## Formato del JSON (resumen)

El JSON es una lista de “artifacts” (uno por MBOX). Campos principales:

- `file`, `count`
- `first_date_utc_iso`, `last_date_utc_iso`
- `top_senders`, `top_recipients`, `top_subjects`, `top_sender_domains`
- `attachments_total`, `duplicates_by_hash`
- `emails[]` con:
  - `from_`, `to`, `subject`, `date_utc_iso`, `message_id`
  - `body`, `body_sha256`
  - `attachments[]` (metadatos)
  - `parse_warnings[]`
