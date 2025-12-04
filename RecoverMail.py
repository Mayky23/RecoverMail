#!/usr/bin/env python3
"""
recovermail.py - Herramienta forense para análisis de archivos MBOX.

Sintaxis básica
---------------
    python recovermail.py [archivos] [--output PREFIJO]

Ejemplos
--------
Analizar un archivo MBOX y generar informes en PDF, HTML y JSON:

    python recovermail.py correo.mbox --output resultados

Analizar múltiples archivos MBOX:

    python recovermail.py correo1.mbox correo2.mbox --output analisis
"""

from __future__ import annotations

import argparse
import mailbox
import os
import json
import html
from pathlib import Path
from datetime import datetime, timezone
from collections import Counter
from typing import Any, Dict, List, Optional

from email.utils import parsedate_to_datetime

from rich.console import Console
from rich.table import Table

from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate,
    Table as PdfTable,
    TableStyle,
    Paragraph,
    PageBreak,
)
from reportlab.lib.styles import getSampleStyleSheet

console = Console()

MAX_BODY_CHARS = 2000  # Máximo de caracteres del cuerpo almacenados en JSON/HTML/PDF


def print_banner() -> None:
    banner = r"""
 ____                                 __  __       _ _     
|  _ \ ___  ___ _____   _____ _ __   |  \/  | __ _(_) |    
| |_) / _ \/ __/ _ \ \ / / _ \ '__|  | |\/| |/ _` | | |    
|  _ <  __/ (_| (_) \ V /  __/ |     | |  | | (_| | | |    
|_| \_\___|\___\___/ \_/ \___|_|     |_|  |_|\__,_|_|_|    
                                                           
[bold blue]Herramienta Forense de Recuperación de Correos Electrónicos[/bold blue]    
[bold cyan]v3.1[/bold cyan] • Soporte MBOX • By: MARH    
"""
    console.print(banner, justify="center")


def is_mbox(filepath: Path) -> bool:
    """
    Detecta si un archivo es MBOX aunque no tenga extensión.

    Estrategia:
      - El archivo debe existir y ser regular.
      - Intenta leer la primera línea y comprobar si comienza con 'From '.
      - Como respaldo, comprueba extensiones típicas de MBOX.
    """
    if not filepath.is_file():
        return False

    try:
        with filepath.open("rb") as f:
            first_line = f.readline().decode("utf-8", errors="ignore")
            if first_line.startswith("From "):
                return True
    except OSError:
        return False

    # Heurística por extensión
    return filepath.suffix.lower() in {".mbox", ".mbx", ".mbxrd"}


def _extract_body(msg: mailbox.mboxMessage) -> str:
    """
    Extrae el cuerpo en texto plano de un mensaje, ignorando adjuntos.

    Devuelve una cadena (posiblemente vacía).
    """
    try:
        if msg.is_multipart():
            parts: List[str] = []
            for part in msg.walk():
                content_type = part.get_content_type()
                disp = (part.get("Content-Disposition") or "").lower()
                if content_type == "text/plain" and "attachment" not in disp:
                    payload = part.get_payload(decode=True)
                    if isinstance(payload, bytes):
                        charset = part.get_content_charset() or "utf-8"
                        text = payload.decode(charset, errors="replace")
                    else:
                        text = payload or ""
                    parts.append(text)
            body = "\n".join(parts)
        else:
            payload = msg.get_payload(decode=True)
            if isinstance(payload, bytes):
                charset = msg.get_content_charset() or "utf-8"
                body = payload.decode(charset, errors="replace")
            else:
                body = payload or ""
    except Exception:
        body = ""

    body = body.replace("\r\n", "\n").replace("\r", "\n")
    if len(body) > MAX_BODY_CHARS:
        return body[:MAX_BODY_CHARS] + "\n[... contenido truncado ...]"
    return body


def _parse_date(raw_date: Optional[str]) -> tuple[str, Optional[datetime]]:
    """
    Intenta normalizar la fecha del correo.

    Devuelve:
        (cadena_para_mostrar, datetime_normalizado_o_None)
    """
    if not raw_date:
        return "N/D", None

    try:
        dt = parsedate_to_datetime(raw_date)
        if dt is None:
            return raw_date, None
        # Normalizamos a UTC sin tzinfo para facilitar comparaciones
        if dt.tzinfo is not None:
            dt = dt.astimezone(timezone.utc).replace(tzinfo=None)
        return dt.strftime("%Y-%m-%d %H:%M:%S"), dt
    except Exception:
        # Si no se puede parsear, devolvemos la cadena original
        return raw_date, None


def analyze_mbox(mbox_path: Path) -> Optional[Dict[str, Any]]:
    """
    Analiza un archivo MBOX y devuelve un diccionario con:
        - file: ruta del archivo
        - count: número de correos
        - first_date / last_date: fechas normalizadas
        - subjects: lista de asuntos más frecuentes
        - senders: lista de remitentes más frecuentes
        - emails: lista de correos con metadatos y cuerpo
    """
    try:
        mbox = mailbox.mbox(mbox_path)
    except Exception as e:
        console.print(f"[red]Error abriendo MBOX ({mbox_path}): {e}[/red]")
        return None

    emails: List[Dict[str, Any]] = []
    dates: List[datetime] = []

    for idx, msg in enumerate(mbox, start=1):
        sender = msg.get("from", "") or msg.get("From", "") or "Desconocido"
        to = msg.get("to", "") or msg.get("To", "") or "Desconocido"
        cc = msg.get("cc", "") or msg.get("Cc", "") or ""
        bcc = msg.get("bcc", "") or msg.get("Bcc", "") or ""
        subject = msg.get("subject", "") or msg.get("Subject", "") or "Sin asunto"

        raw_date = msg.get("date") or msg.get("Date")
        date_str, dt = _parse_date(raw_date)
        if dt is not None:
            dates.append(dt)

        body = _extract_body(msg)

        email_entry: Dict[str, Any] = {
            "id": idx,
            "from": sender,
            "to": to,
            "cc": cc,
            "bcc": bcc,
            "subject": subject,
            "date": date_str,
            "raw_date": raw_date or "",
            "body": body,
        }
        emails.append(email_entry)

    if not emails:
        return None

    # Estadísticos
    senders_counter = Counter(e["from"] for e in emails if e["from"] != "Desconocido")
    subjects_counter = Counter(e["subject"] for e in emails if e["subject"] != "Sin asunto")

    top_senders = [s for s, _ in senders_counter.most_common(5)] or ["N/D"]
    top_subjects = [s for s, _ in subjects_counter.most_common(5)] or ["N/D"]

    if dates:
        first_date = min(dates).strftime("%Y-%m-%d %H:%M:%S")
        last_date = max(dates).strftime("%Y-%m-%d %H:%M:%S")
    else:
        first_date = last_date = "N/D"

    summary: Dict[str, Any] = {
        "file": str(mbox_path),
        "count": len(emails),
        "first_date": first_date,
        "last_date": last_date,
        "subjects": top_subjects,
        "senders": top_senders,
        "emails": emails,
    }
    return summary


def generate_report(artifacts: List[Dict[str, Any]]) -> None:
    """
    Muestra en consola un resumen de los MBOX analizados.
    """
    table = Table(title="Resumen de Análisis Forense", show_lines=True)
    table.add_column("Archivo", style="magenta", no_wrap=True)
    table.add_column("Emails", style="green")
    table.add_column("Primera fecha", style="yellow")
    table.add_column("Última fecha", style="yellow")
    table.add_column("Asuntos destacados", style="blue")
    table.add_column("Remitentes destacados", style="red")

    for art in artifacts:
        table.add_row(
            Path(art["file"]).name,
            str(art["count"]),
            art["first_date"],
            art["last_date"],
            "\n".join(art.get("subjects", ["N/D"])),
            "\n".join(art.get("senders", ["N/D"])),
        )

    console.print(table)


def export_pdf(artifacts: List[Dict[str, Any]], filename: str) -> None:
    """
    Genera un informe PDF con un resumen general y tablas por archivo.
    """
    try:
        doc = SimpleDocTemplate(filename, pagesize=letter)
        styles = getSampleStyleSheet()
        elements: List[Any] = []

        # Título del informe
        elements.append(Paragraph("Informe RecoverMail", styles["Title"]))
        elements.append(
            Paragraph(
                f"Generado el {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
                styles["Normal"],
            )
        )
        elements.append(Paragraph(" ", styles["Normal"]))  # Espacio

        # Tabla resumen
        elements.append(Paragraph("Resumen de Archivos Analizados", styles["Heading2"]))
        summary_data = [["Archivo", "Emails", "Primera fecha", "Última fecha"]]
        for art in artifacts:
            summary_data.append(
                [
                    Path(art["file"]).name,
                    str(art["count"]),
                    art["first_date"],
                    art["last_date"],
                ]
            )

        summary_table = PdfTable(summary_data)
        summary_table.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                    ("FONTSIZE", (0, 0), (-1, -1), 9),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ]
            )
        )
        elements.append(summary_table)
        elements.append(PageBreak())

        # Detalle de cada archivo
        for art in artifacts:
            elements.append(
                Paragraph(
                    f"Detalle de archivo: {Path(art['file']).name}",
                    styles["Heading2"],
                )
            )

            emails = art.get("emails", [])
            if emails:
                data_emails = [["#", "De", "Para", "Asunto", "Fecha"]]
                for email in emails:
                    data_emails.append(
                        [
                            str(email["id"]),
                            (email["from"] or "")[:40],
                            (email["to"] or "")[:40],
                            (email["subject"] or "")[:60],
                            email["date"],
                        ]
                    )

                email_table = PdfTable(data_emails, repeatRows=1)
                email_table.setStyle(
                    TableStyle(
                        [
                            ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                            ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                            ("FONTSIZE", (0, 0), (-1, -1), 7),
                            ("VALIGN", (0, 0), (-1, -1), "TOP"),
                        ]
                    )
                )
                elements.append(email_table)
            else:
                elements.append(Paragraph("No se encontraron correos.", styles["Normal"]))

            elements.append(PageBreak())

        doc.build(elements)
        console.print(f"[green]PDF exportado: {filename}[/green]")
    except Exception as e:
        console.print(f"[red]Error generando PDF: {e}[/red]")


def export_html(artifacts: List[Dict[str, Any]], filename: str) -> None:
    """
    Genera un informe HTML interactivo con resumen y detalle de correos.
    """
    try:
        generated_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        parts: List[str] = [
            "<!DOCTYPE html>",
            "<html>",
            "<head>",
            "<meta charset='UTF-8' />",
            "<title>Informe RecoverMail</title>",
            "<style>",
            "body {font-family: Arial, sans-serif; margin: 0; padding: 0;}",
            ".container {width: 95%; margin: 0 auto; padding: 20px;}",
            "table {border-collapse: collapse; width: 100%; margin-top: 10px; margin-bottom: 25px;}",
            "th, td {border: 1px solid #ddd; padding: 6px 8px; text-align: left; font-size: 13px;}",
            "th {background-color: #f2f2f2; position: sticky; top: 0;}",
            "tr:nth-child(even) {background-color: #f9f9f9;}",
            ".header {background-color: #004080; color: white; padding: 15px; text-align: center; margin-bottom: 20px;}",
            "h2, h3 {color: #004080;}",
            ".summary {margin-bottom: 30px;}",
            ".email-body {white-space: pre-wrap; font-family: monospace; max-height: 250px; overflow-y: auto; border: 1px solid #ddd; padding: 8px;}",
            ".accordion {background-color: #eee; cursor: pointer; padding: 10px 14px; width: 100%; text-align: left; border: none; outline: none; transition: 0.3s; margin-bottom: 5px; font-size: 15px;}",
            ".accordion:hover {background-color: #ddd;}",
            ".panel {padding: 0 10px 10px 10px; display: none; overflow: hidden; border: 1px solid #ddd; margin-bottom: 10px;}",
            "</style>",
            "<script>",
            "function togglePanel(id) {",
            "  var el = document.getElementById(id);",
            "  if (!el) return;",
            "  if (el.style.display === 'block') {",
            "    el.style.display = 'none';",
            "  } else {",
            "    el.style.display = 'block';",
            "  }",
            "}",
            "</script>",
            "</head>",
            "<body>",
            "<div class='header'>",
            "<h1>Informe RecoverMail</h1>",
            f"<p>Generado el {html.escape(generated_at)}</p>",
            "</div>",
            "<div class='container'>",
            "<div class='summary'>",
            "<h2>Resumen de archivos analizados</h2>",
            "<table>",
            "<tr><th>Archivo</th><th>Emails</th><th>Primera fecha</th><th>Última fecha</th></tr>",
        ]

        # Resumen
        for art in artifacts:
            parts.append(
                "<tr>"
                f"<td>{html.escape(Path(art['file']).name)}</td>"
                f"<td>{art['count']}</td>"
                f"<td>{html.escape(art['first_date'])}</td>"
                f"<td>{html.escape(art['last_date'])}</td>"
                "</tr>"
            )

        parts.extend(["</table>", "</div>"])

        # Detalle de correos
        for idx, art in enumerate(artifacts):
            panel_id = f"panel-{idx}"
            parts.append(
                f"<button class='accordion' onclick=\"togglePanel('{panel_id}')\">"
                f"Ver detalles de {html.escape(Path(art['file']).name)} "
                f"({art['count']} correos)"
                "</button>"
            )
            parts.append(f"<div class='panel' id='{panel_id}'>")
            parts.append(
                f"<h3>Correos en archivo: {html.escape(Path(art['file']).name)}</h3>"
            )
            parts.append("<table>")
            parts.append(
                "<tr><th>#</th><th>De</th><th>Para</th><th>Asunto</th><th>Fecha</th><th>Cuerpo</th></tr>"
            )

            for email in art.get("emails", []):
                parts.append(
                    "<tr>"
                    f"<td>{email['id']}</td>"
                    f"<td>{html.escape(email.get('from', '')[:60])}</td>"
                    f"<td>{html.escape(email.get('to', '')[:60])}</td>"
                    f"<td>{html.escape(email.get('subject', '')[:60])}</td>"
                    f"<td>{html.escape(email.get('date', ''))}</td>"
                    "<td><div class='email-body'>"
                    f"{html.escape(email.get('body', '')[:MAX_BODY_CHARS])}"
                    "</div></td>"
                    "</tr>"
                )

            parts.append("</table>")
            parts.append("</div>")

        parts.extend(["</div>", "</body>", "</html>"])

        html_content = "\n".join(parts)

        with open(filename, "w", encoding="utf-8") as f:
            f.write(html_content)

        console.print(f"[green]HTML exportado: {filename}[/green]")
    except Exception as e:
        console.print(f"[red]Error generando HTML: {e}[/red]")


def export_json(artifacts: List[Dict[str, Any]], filename: str) -> None:
    """
    Exporta todos los datos en formato JSON estructurado.
    """
    try:
        with open(filename, "w", encoding="utf-8") as f:
            json.dump(artifacts, f, ensure_ascii=False, indent=2)
        console.print(f"[green]JSON exportado: {filename}[/green]")
    except Exception as e:
        console.print(f"[red]Error generando JSON: {e}[/red]")


def main(argv: Optional[List[str]] = None) -> int:
    print_banner()

    parser = argparse.ArgumentParser(
        description="Analizador forense de correos MBOX (sin subcomandos)."
    )
    parser.add_argument(
        "files",
        nargs="+",
        help="Archivos MBOX a analizar (pueden no tener extensión .mbox).",
    )
    parser.add_argument(
        "--output",
        "-o",
        default="informe_mbox",
        help="Prefijo para los archivos de salida (sin extensión).",
    )

    args = parser.parse_args(argv)

    artifacts: List[Dict[str, Any]] = []
    total_emails = 0

    with console.status("[bold green]Analizando archivos de correo..."):
        for file_path in args.files:
            path = Path(file_path)

            if not path.exists():
                console.print(f"[red]Archivo no encontrado: {file_path}[/red]")
                continue

            if not is_mbox(path):
                console.print(
                    f"[yellow]El archivo no parece ser un MBOX válido: {file_path}[/yellow]"
                )
                continue

            console.print(f"[cyan]→ Analizando:[/cyan] {path}")
            summary = analyze_mbox(path)
            if not summary:
                console.print(
                    f"[yellow]No se pudieron extraer correos de: {file_path}[/yellow]"
                )
                continue

            artifacts.append(summary)
            total_emails += summary["count"]

    if not artifacts:
        console.print(
            "[bold yellow]No se encontraron archivos MBOX válidos o sin correos para analizar.[/bold yellow]"
        )
        return 1

    console.print(
        f"\n[bold green]Análisis completado:[/bold green] {len(artifacts)} archivo(s), {total_emails} correo(s) procesados."
    )

    # Resumen en consola
    generate_report(artifacts)

    # Archivos de salida
    base = args.output
    output_dir = Path(os.getcwd())
    pdf_file = str(output_dir / f"{base}.pdf")
    html_file = str(output_dir / f"{base}.html")
    json_file = str(output_dir / f"{base}.json")

    with console.status("[bold yellow]Exportando resultados..."):
        export_pdf(artifacts, pdf_file)
        export_html(artifacts, html_file)
        export_json(artifacts, json_file)

    console.print("\n[bold green]Resultados exportados correctamente:[/bold green]")
    console.print(f"[cyan]► HTML:[/cyan] {html_file}")
    console.print(f"[cyan]► PDF: [/cyan] {pdf_file}")
    console.print(f"[cyan]► JSON:[/cyan] {json_file}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
