#!/usr/bin/env python3
"""
recovermail.py — Suite forense para análisis y extracción de correos MBOX.

Uso rápido
---------
  python recovermail.py correo.mbox -o informe
  python recovermail.py carpeta_con_mbox/ --recursive -o caso_001 --outdir resultados

Qué hace
--------
- Detecta y analiza archivos MBOX (aunque no tengan extensión).
- Extrae metadatos (From/To/Cc/Bcc/Subject/Date/Message-ID), cuerpo (texto) y adjuntos (solo metadatos).
- Normaliza fechas a UTC (ISO 8601) cuando es posible.
- Genera informes: PDF (resumen + tablas), HTML (buscable + desplegables) y JSON (estructurado).
- Incluye métricas útiles: top remitentes/asuntos/dominios, duplicados por hash, conteo de adjuntos, etc.

Dependencias
------------
- rich
- reportlab

Notas
-----
- No modifica el MBOX original.
- El cuerpo puede truncarse (configurable) para evitar salidas gigantes.
"""

from __future__ import annotations

import argparse
import hashlib
import html
import json
import os
import re
import sys
from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
from email.header import decode_header, make_header
from email.message import Message
from email.utils import getaddresses, parsedate_to_datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import mailbox

# Dependencias de salida / UI (cargadas con control de error)
try:
    from rich.console import Console
    from rich.table import Table
    from rich.progress import Progress, SpinnerColumn, TextColumn, BarColumn, TimeElapsedColumn
except Exception:  # pragma: no cover
    Console = None  # type: ignore
    Table = None  # type: ignore
    Progress = None  # type: ignore

try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Table as PdfTable, TableStyle
except Exception:  # pragma: no cover
    colors = None  # type: ignore
    letter = None  # type: ignore
    getSampleStyleSheet = None  # type: ignore
    PageBreak = None  # type: ignore
    Paragraph = None  # type: ignore
    SimpleDocTemplate = None  # type: ignore
    PdfTable = None  # type: ignore
    TableStyle = None  # type: ignore


APP_NAME = "RecoverMail"
APP_VERSION = "4.0"
DEFAULT_MAX_BODY_CHARS = 2000
DEFAULT_TOP_N = 8


# -----------------------------
# Modelos de datos
# -----------------------------
@dataclass
class AttachmentInfo:
    filename: str
    content_type: str
    size_bytes: int


@dataclass
class EmailRecord:
    id: int
    from_: str
    to: str
    cc: str
    bcc: str
    subject: str
    date_display: str
    date_utc_iso: str
    raw_date: str
    message_id: str
    body: str
    body_sha256: str
    attachments: List[AttachmentInfo] = field(default_factory=list)
    parse_warnings: List[str] = field(default_factory=list)


@dataclass
class MboxArtifact:
    file: str
    count: int
    first_date_utc_iso: str
    last_date_utc_iso: str
    top_senders: List[str]
    top_recipients: List[str]
    top_subjects: List[str]
    top_sender_domains: List[str]
    attachments_total: int
    duplicates_by_hash: int
    emails: List[EmailRecord]
    warnings: List[str] = field(default_factory=list)


# -----------------------------
# Utilidades generales
# -----------------------------
def _console() -> Any:
    if Console is None:
        return None
    return Console()


def print_banner(console: Any) -> None:
    if console is None:
        return
    banner = rf"""
 ____                                 __  __       _ _
|  _ \ ___  ___ _____   _____ _ __   |  \/  | __ _(_) |
| |_) / _ \/ __/ _ \ \ / / _ \ '__|  | |\/| |/ _` | | |
|  _ <  __/ (_| (_) \ V /  __/ |     | |  | | (_| | | |
|_| \_\___|\___\___/ \_/ \___|_|     |_|  |_|\__,_|_|_|

[bold blue]{APP_NAME}[/bold blue]  [bold cyan]v{APP_VERSION}[/bold cyan]
Herramienta forense de análisis MBOX • PDF/HTML/JSON • sin modificar evidencias
"""
    console.print(banner, justify="center")


def is_mbox(filepath: Path) -> bool:
    """
    Heurística de detección MBOX:
    - Debe ser archivo regular.
    - Primera línea suele empezar por 'From ' (mboxrd / mboxo).
    - Respaldo: extensiones típicas.
    """
    if not filepath.is_file():
        return False

    try:
        with filepath.open("rb") as f:
            first_line = f.readline(256).decode("utf-8", errors="ignore")
            if first_line.startswith("From "):
                return True
    except OSError:
        return False

    return filepath.suffix.lower() in {".mbox", ".mbx", ".mbxrd", ".mboxo", ".mboxrd"}


def safe_decode_header(value: str) -> str:
    """Decodifica headers MIME (=?utf-8?...?=) de forma robusta."""
    if not value:
        return ""
    try:
        # make_header + decode_header resuelve múltiples partes y encodings
        decoded = str(make_header(decode_header(value)))
        return decoded.strip()
    except Exception:
        return value.strip()


def normalize_whitespace(s: str) -> str:
    return re.sub(r"[ \t]+", " ", (s or "").strip())


def parse_addresses(header_value: str) -> str:
    """
    Convierte un header de direcciones en una cadena legible y estable.
    Ej: 'Nombre <a@b.com>, x@y.com' -> 'Nombre <a@b.com>, x@y.com'
    """
    if not header_value:
        return ""
    try:
        pairs = getaddresses([header_value])
        rendered: List[str] = []
        for name, addr in pairs:
            name = normalize_whitespace(safe_decode_header(name))
            addr = normalize_whitespace(addr)
            if name and addr:
                rendered.append(f"{name} <{addr}>")
            elif addr:
                rendered.append(addr)
            elif name:
                rendered.append(name)
        return ", ".join(rendered)
    except Exception:
        return normalize_whitespace(safe_decode_header(header_value))


def parse_date(raw_date: Optional[str]) -> Tuple[str, str]:
    """
    Devuelve:
      - date_display: 'YYYY-mm-dd HH:MM:SS (UTC)' o 'N/D'
      - date_utc_iso: ISO8601 en UTC con 'Z' (o cadena vacía si N/D)
    """
    if not raw_date:
        return "N/D", ""
    try:
        dt = parsedate_to_datetime(raw_date)
        if dt is None:
            return safe_decode_header(raw_date), ""
        if dt.tzinfo is None:
            # si no trae TZ, asumimos que ya está en UTC (mejor que inventar local)
            dt = dt.replace(tzinfo=timezone.utc)
        dt_utc = dt.astimezone(timezone.utc)
        display = dt_utc.strftime("%Y-%m-%d %H:%M:%S (UTC)")
        iso = dt_utc.replace(microsecond=0).isoformat().replace("+00:00", "Z")
        return display, iso
    except Exception:
        return safe_decode_header(raw_date), ""


def sha256_text(text: str) -> str:
    h = hashlib.sha256()
    h.update((text or "").encode("utf-8", errors="replace"))
    return h.hexdigest()


_TAG_RE = re.compile(r"<[^>]+>")
_WS_RE = re.compile(r"\s+")


def html_to_text_basic(s: str) -> str:
    # Conversión mínima sin dependencias externas.
    s = s.replace("\r\n", "\n").replace("\r", "\n")
    s = _TAG_RE.sub(" ", s)
    s = html.unescape(s)
    s = _WS_RE.sub(" ", s).strip()
    return s


def _payload_to_text(payload: Any, charset: Optional[str]) -> str:
    if payload is None:
        return ""
    if isinstance(payload, bytes):
        enc = charset or "utf-8"
        try:
            return payload.decode(enc, errors="replace")
        except Exception:
            return payload.decode("utf-8", errors="replace")
    if isinstance(payload, str):
        return payload
    return ""


def extract_body_text(msg: Message, prefer_plain: bool = True) -> Tuple[str, List[str]]:
    """
    Extrae cuerpo textual ignorando adjuntos.
    - prefer_plain=True: prioriza text/plain; si no, usa text/html convertido.
    Devuelve: (body_text, warnings)
    """
    warnings: List[str] = []
    texts_plain: List[str] = []
    texts_html: List[str] = []

    try:
        if msg.is_multipart():
            for part in msg.walk():
                ctype = (part.get_content_type() or "").lower()
                disp = (part.get("Content-Disposition") or "").lower()
                if "attachment" in disp:
                    continue
                if ctype not in {"text/plain", "text/html"}:
                    continue

                payload = part.get_payload(decode=True)
                charset = part.get_content_charset()
                txt = _payload_to_text(payload, charset)

                if ctype == "text/plain":
                    texts_plain.append(txt)
                else:
                    texts_html.append(txt)
        else:
            ctype = (msg.get_content_type() or "").lower()
            payload = msg.get_payload(decode=True)
            charset = msg.get_content_charset()
            txt = _payload_to_text(payload, charset)
            if ctype == "text/plain":
                texts_plain.append(txt)
            elif ctype == "text/html":
                texts_html.append(txt)
            else:
                # a veces correos raros sin ctype correcto
                texts_plain.append(txt)

    except Exception as e:
        warnings.append(f"Error extrayendo cuerpo: {e}")

    body = ""
    if prefer_plain and any(t.strip() for t in texts_plain):
        body = "\n\n".join(t.strip() for t in texts_plain if t is not None)
    elif any(t.strip() for t in texts_html):
        merged = "\n\n".join(t.strip() for t in texts_html if t is not None)
        body = html_to_text_basic(merged)
    else:
        body = "\n\n".join(t.strip() for t in texts_plain + texts_html if t is not None)

    body = body.replace("\r\n", "\n").replace("\r", "\n").strip()
    return body, warnings


def list_attachments(msg: Message) -> List[AttachmentInfo]:
    atts: List[AttachmentInfo] = []
    try:
        if not msg.is_multipart():
            return atts
        for part in msg.walk():
            disp = (part.get("Content-Disposition") or "").lower()
            if "attachment" not in disp:
                continue
            filename = safe_decode_header(part.get_filename() or "adjunto_sin_nombre")
            ctype = (part.get_content_type() or "application/octet-stream").lower()
            payload = part.get_payload(decode=True)
            size = len(payload) if isinstance(payload, (bytes, bytearray)) else 0
            atts.append(AttachmentInfo(filename=filename, content_type=ctype, size_bytes=size))
    except Exception:
        # si falla, devolvemos lo que haya (o vacío)
        pass
    return atts


def iter_input_paths(paths: Sequence[str], recursive: bool) -> Iterable[Path]:
    """
    Expande entradas:
    - Archivos directos
    - Directorios: lista archivos dentro (recursivo opcional)
    """
    for p in paths:
        path = Path(p)
        if path.is_file():
            yield path
        elif path.is_dir():
            if recursive:
                for f in path.rglob("*"):
                    if f.is_file():
                        yield f
            else:
                for f in path.glob("*"):
                    if f.is_file():
                        yield f
        else:
            # soporte básico de glob (p.ej. *.mbox)
            for f in Path(".").glob(p):
                if f.is_file():
                    yield f


# -----------------------------
# Análisis principal
# -----------------------------
def analyze_mbox(
    mbox_path: Path,
    *,
    max_body_chars: int,
    top_n: int,
    include_body: bool,
    prefer_plain: bool,
) -> Optional[MboxArtifact]:
    warnings: List[str] = []
    emails: List[EmailRecord] = []
    dates_utc: List[datetime] = []
    attachments_total = 0

    try:
        mbox = mailbox.mbox(mbox_path)
    except Exception as e:
        warnings.append(f"Error abriendo MBOX: {e}")
        return None

    # Nota: mailbox.mbox itera "mensajes" ya parseados; cada msg es email.message.Message
    for idx, msg in enumerate(mbox, start=1):
        parse_warnings: List[str] = []
        try:
            # Headers (decodificados)
            sender = parse_addresses(msg.get("From", "") or msg.get("from", "")) or "Desconocido"
            to = parse_addresses(msg.get("To", "") or msg.get("to", "")) or "Desconocido"
            cc = parse_addresses(msg.get("Cc", "") or msg.get("cc", "")) or ""
            bcc = parse_addresses(msg.get("Bcc", "") or msg.get("bcc", "")) or ""
            subject = safe_decode_header(msg.get("Subject", "") or msg.get("subject", "")) or "Sin asunto"
            raw_date = msg.get("Date", "") or msg.get("date", "") or ""
            date_display, date_utc_iso = parse_date(raw_date)
            message_id = normalize_whitespace(msg.get("Message-ID", "") or msg.get("message-id", ""))

            # Fechas para first/last
            if date_utc_iso:
                try:
                    dt = datetime.fromisoformat(date_utc_iso.replace("Z", "+00:00"))
                    dates_utc.append(dt.astimezone(timezone.utc))
                except Exception:
                    parse_warnings.append("Fecha ISO no parseable para estadísticas.")

            # Cuerpo
            body_text = ""
            body_hash = ""
            if include_body:
                body_text, w = extract_body_text(msg, prefer_plain=prefer_plain)
                parse_warnings.extend(w)
                if max_body_chars > 0 and len(body_text) > max_body_chars:
                    body_text = body_text[:max_body_chars] + "\n[... contenido truncado ...]"
                body_hash = sha256_text(body_text)
            else:
                body_hash = ""

            # Adjuntos
            atts = list_attachments(msg)
            attachments_total += len(atts)

            emails.append(
                EmailRecord(
                    id=idx,
                    from_=sender,
                    to=to,
                    cc=cc,
                    bcc=bcc,
                    subject=subject,
                    date_display=date_display,
                    date_utc_iso=date_utc_iso,
                    raw_date=raw_date,
                    message_id=message_id,
                    body=body_text,
                    body_sha256=body_hash,
                    attachments=atts,
                    parse_warnings=parse_warnings,
                )
            )
        except Exception as e:
            warnings.append(f"Error parseando mensaje #{idx}: {e}")

    if not emails:
        return None

    # Orden estable por fecha (si existe), si no por id
    def sort_key(e: EmailRecord) -> Tuple[int, str, int]:
        has_date = 1 if e.date_utc_iso else 0
        return (-has_date, e.date_utc_iso or "", e.id)

    emails.sort(key=sort_key)

    # Estadísticos
    from_list = [e.from_ for e in emails if e.from_ and e.from_ != "Desconocido"]
    to_list = [e.to for e in emails if e.to and e.to != "Desconocido"]
    subject_list = [e.subject for e in emails if e.subject and e.subject != "Sin asunto"]

    def top(items: List[str], n: int) -> List[str]:
        from collections import Counter

        if not items:
            return ["N/D"]
        return [k for k, _ in Counter(items).most_common(n)] or ["N/D"]

    def extract_domain(addr_field: str) -> List[str]:
        # extrae dominios de cualquier cosa que parezca correo
        domains: List[str] = []
        for _, addr in getaddresses([addr_field]):
            if "@" in addr:
                domains.append(addr.split("@", 1)[1].lower())
        return domains

    domains: List[str] = []
    for e in emails:
        domains.extend(extract_domain(e.from_))

    top_senders = top(from_list, top_n)
    top_recipients = top(to_list, top_n)
    top_subjects = top(subject_list, top_n)
    top_sender_domains = top(domains, top_n)

    # Duplicados (por hash de body; si no se incluye body, no aplica)
    duplicates_by_hash = 0
    if include_body:
        seen: Dict[str, int] = {}
        for e in emails:
            if not e.body_sha256:
                continue
            if e.body_sha256 in seen:
                duplicates_by_hash += 1
            else:
                seen[e.body_sha256] = 1

    # first/last
    if dates_utc:
        first_dt = min(dates_utc).replace(microsecond=0)
        last_dt = max(dates_utc).replace(microsecond=0)
        first_iso = first_dt.isoformat().replace("+00:00", "Z")
        last_iso = last_dt.isoformat().replace("+00:00", "Z")
    else:
        first_iso = last_iso = ""

    return MboxArtifact(
        file=str(mbox_path),
        count=len(emails),
        first_date_utc_iso=first_iso or "N/D",
        last_date_utc_iso=last_iso or "N/D",
        top_senders=top_senders,
        top_recipients=top_recipients,
        top_subjects=top_subjects,
        top_sender_domains=top_sender_domains,
        attachments_total=attachments_total,
        duplicates_by_hash=duplicates_by_hash,
        emails=emails,
        warnings=warnings,
    )


# -----------------------------
# Salidas / reportes
# -----------------------------
def print_summary_table(console: Any, artifacts: List[MboxArtifact]) -> None:
    if console is None or Table is None:
        # fallback simple
        for a in artifacts:
            print(f"- {Path(a.file).name}: {a.count} emails, {a.first_date_utc_iso} .. {a.last_date_utc_iso}")
        return

    table = Table(title="Resumen de análisis", show_lines=True)
    table.add_column("Archivo", style="magenta", no_wrap=True)
    table.add_column("Emails", style="green", justify="right")
    table.add_column("Primera fecha (UTC)", style="yellow")
    table.add_column("Última fecha (UTC)", style="yellow")
    table.add_column("Adjuntos", style="cyan", justify="right")
    table.add_column("Duplicados (hash)", style="red", justify="right")
    table.add_column("Top remitentes", style="blue")
    table.add_column("Top asuntos", style="blue")

    for a in artifacts:
        table.add_row(
            Path(a.file).name,
            str(a.count),
            a.first_date_utc_iso,
            a.last_date_utc_iso,
            str(a.attachments_total),
            str(a.duplicates_by_hash),
            "\n".join(a.top_senders[:3]) if a.top_senders else "N/D",
            "\n".join(a.top_subjects[:3]) if a.top_subjects else "N/D",
        )

    console.print(table)


def export_json(artifacts: List[MboxArtifact], filename: Path) -> None:
    payload = [asdict(a) for a in artifacts]
    filename.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def export_html(artifacts: List[MboxArtifact], filename: Path, max_body_chars: int) -> None:
    generated_at = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")

    # HTML con búsqueda (client-side) y detalles desplegables
    parts: List[str] = [
        "<!DOCTYPE html>",
        "<html><head>",
        "<meta charset='UTF-8' />",
        f"<title>{html.escape(APP_NAME)} — Informe</title>",
        "<meta name='viewport' content='width=device-width, initial-scale=1'/>",
        "<style>",
        "body{font-family:Arial, sans-serif; margin:0; background:#fafafa; color:#111;}",
        ".header{background:#0b3d91; color:#fff; padding:18px 14px;}",
        ".wrap{max-width:1200px; margin:0 auto; padding:18px 14px;}",
        ".card{background:#fff; border:1px solid #e6e6e6; border-radius:10px; padding:14px; margin:14px 0; box-shadow:0 1px 2px rgba(0,0,0,.04);}",
        "table{border-collapse:collapse; width:100%;}",
        "th,td{border:1px solid #e6e6e6; padding:8px; font-size:13px; vertical-align:top;}",
        "th{background:#f2f4f7; position:sticky; top:0; z-index:2;}",
        "tr:nth-child(even){background:#fcfcfc;}",
        ".muted{color:#555; font-size:13px;}",
        ".pill{display:inline-block; padding:2px 8px; border-radius:999px; background:#eef2ff; border:1px solid #dbe4ff; font-size:12px; margin-right:6px;}",
        ".search{width:100%; padding:10px 12px; border:1px solid #d0d5dd; border-radius:10px; font-size:14px;}",
        "details{border:1px solid #e6e6e6; border-radius:10px; background:#fff; padding:10px 12px; margin:10px 0;}",
        "summary{cursor:pointer; font-weight:600;}",
        "pre{white-space:pre-wrap; word-wrap:break-word; background:#0b1220; color:#e6edf3; padding:10px; border-radius:10px; overflow:auto;}",
        ".grid{display:grid; grid-template-columns:repeat(6, minmax(0,1fr)); gap:10px;}",
        ".grid > div{background:#fff; border:1px solid #e6e6e6; border-radius:10px; padding:10px;}",
        "@media(max-width:900px){.grid{grid-template-columns:repeat(2, minmax(0,1fr));}}",
        "</style>",
        "<script>",
        "function filterRows(inputId, tableId){",
        "  const q = document.getElementById(inputId).value.toLowerCase();",
        "  const rows = document.getElementById(tableId).getElementsByTagName('tr');",
        "  for (let i=1;i<rows.length;i++){",
        "    const row = rows[i];",
        "    const txt = row.innerText.toLowerCase();",
        "    row.style.display = txt.includes(q) ? '' : 'none';",
        "  }",
        "}",
        "</script>",
        "</head><body>",
        "<div class='header'>",
        f"<div class='wrap'><h1 style='margin:0'>{html.escape(APP_NAME)} — Informe forense</h1>",
        f"<div class='muted' style='color:#dbeafe'>Generado: {html.escape(generated_at)} • Versión: {html.escape(APP_VERSION)}</div></div>",
        "</div>",
        "<div class='wrap'>",
        "<div class='card'>",
        "<h2 style='margin:0 0 10px 0'>Resumen</h2>",
        "<div class='grid'>",
    ]

    total_files = len(artifacts)
    total_emails = sum(a.count for a in artifacts)
    total_atts = sum(a.attachments_total for a in artifacts)
    total_dups = sum(a.duplicates_by_hash for a in artifacts)

    parts += [
        f"<div><div class='muted'>Archivos</div><div style='font-size:22px; font-weight:700'>{total_files}</div></div>",
        f"<div><div class='muted'>Correos</div><div style='font-size:22px; font-weight:700'>{total_emails}</div></div>",
        f"<div><div class='muted'>Adjuntos</div><div style='font-size:22px; font-weight:700'>{total_atts}</div></div>",
        f"<div><div class='muted'>Duplicados (hash)</div><div style='font-size:22px; font-weight:700'>{total_dups}</div></div>",
        f"<div><div class='muted'>Truncado cuerpo</div><div style='font-size:22px; font-weight:700'>{max_body_chars if max_body_chars>0 else 'No'}</div></div>",
        "<div><div class='muted'>TZ</div><div style='font-size:22px; font-weight:700'>UTC</div></div>",
        "</div></div>",
    ]

    # Tabla resumen archivos
    parts += [
        "<div class='card'>",
        "<h2 style='margin:0 0 10px 0'>Archivos analizados</h2>",
        "<table>",
        "<tr><th>Archivo</th><th>Emails</th><th>Primera fecha (UTC)</th><th>Última fecha (UTC)</th><th>Adjuntos</th><th>Duplicados</th></tr>",
    ]
    for a in artifacts:
        parts.append(
            "<tr>"
            f"<td>{html.escape(Path(a.file).name)}</td>"
            f"<td>{a.count}</td>"
            f"<td>{html.escape(a.first_date_utc_iso)}</td>"
            f"<td>{html.escape(a.last_date_utc_iso)}</td>"
            f"<td>{a.attachments_total}</td>"
            f"<td>{a.duplicates_by_hash}</td>"
            "</tr>"
        )
    parts += ["</table></div>"]

    # Detalle por archivo
    for file_idx, a in enumerate(artifacts, start=1):
        table_id = f"t{file_idx}"
        input_id = f"s{file_idx}"

        parts += [
            "<div class='card'>",
            f"<h2 style='margin:0 0 8px 0'>Detalle: {html.escape(Path(a.file).name)}</h2>",
            "<div class='muted' style='margin-bottom:10px'>"
            f"<span class='pill'>Emails: {a.count}</span>"
            f"<span class='pill'>Adjuntos: {a.attachments_total}</span>"
            f"<span class='pill'>Duplicados: {a.duplicates_by_hash}</span>"
            f"<span class='pill'>Rango: {html.escape(a.first_date_utc_iso)} → {html.escape(a.last_date_utc_iso)}</span>"
            "</div>",
            "<div style='display:flex; gap:10px; flex-wrap:wrap; margin-bottom:10px'>",
            f"<input id='{input_id}' class='search' placeholder='Buscar en este archivo (from, to, asunto, fecha, body...)' "
            f"oninput=\"filterRows('{input_id}','{table_id}')\"/>",
            "</div>",
            f"<table id='{table_id}'>",
            "<tr><th>#</th><th>Fecha (UTC)</th><th>De</th><th>Para</th><th>Asunto</th><th>Adjuntos</th><th>Detalles</th></tr>",
        ]

        for e in a.emails:
            att_list = ""
            if e.attachments:
                att_list = "<br/>".join(
                    f"{html.escape(att.filename)} ({att.size_bytes} B, {html.escape(att.content_type)})"
                    for att in e.attachments[:6]
                )
                if len(e.attachments) > 6:
                    att_list += "<br/>…"
            else:
                att_list = "—"

            body_preview = e.body or ""
            if max_body_chars > 0 and len(body_preview) > max_body_chars:
                body_preview = body_preview[:max_body_chars] + "\n[... truncado ...]"

            # Detalles: body, message-id, raw-date, warnings
            details_parts = [
                "<details><summary>Ver</summary>",
                "<div class='muted' style='margin:8px 0'>",
                f"<div><b>Message-ID:</b> {html.escape(e.message_id or '—')}</div>",
                f"<div><b>Date raw:</b> {html.escape(e.raw_date or '—')}</div>",
                f"<div><b>Hash body (sha256):</b> {html.escape(e.body_sha256 or '—')}</div>",
                "</div>",
            ]
            if e.parse_warnings:
                details_parts.append("<div class='muted' style='margin:8px 0'><b>Warnings:</b><ul>")
                details_parts.extend([f"<li>{html.escape(w)}</li>" for w in e.parse_warnings[:10]])
                details_parts.append("</ul></div>")
            details_parts.append("<pre>" + html.escape(body_preview) + "</pre>")
            details_parts.append("</details>")

            parts.append(
                "<tr>"
                f"<td>{e.id}</td>"
                f"<td>{html.escape(e.date_utc_iso or e.date_display)}</td>"
                f"<td>{html.escape(e.from_[:90])}</td>"
                f"<td>{html.escape(e.to[:90])}</td>"
                f"<td>{html.escape(e.subject[:120])}</td>"
                f"<td>{att_list}</td>"
                f"<td>{''.join(details_parts)}</td>"
                "</tr>"
            )

        parts += ["</table>"]

        # Top lists
        parts += [
            "<div style='margin-top:12px' class='muted'>",
            f"<b>Top remitentes:</b> {html.escape(' • '.join(a.top_senders))}<br/>",
            f"<b>Top destinatarios:</b> {html.escape(' • '.join(a.top_recipients))}<br/>",
            f"<b>Top dominios (From):</b> {html.escape(' • '.join(a.top_sender_domains))}<br/>",
            f"<b>Top asuntos:</b> {html.escape(' • '.join(a.top_subjects))}",
            "</div>",
            "</div>",
        ]

        if a.warnings:
            parts += ["<div class='card'>", "<h3 style='margin:0 0 8px 0'>Warnings del archivo</h3><ul>"]
            parts += [f"<li>{html.escape(w)}</li>" for w in a.warnings[:50]]
            parts += ["</ul></div>"]

    parts += ["</div></body></html>"]
    filename.write_text("\n".join(parts), encoding="utf-8")


def export_pdf(artifacts: List[MboxArtifact], filename: Path) -> None:
    if SimpleDocTemplate is None:
        raise RuntimeError("reportlab no está instalado. Instala 'reportlab' para exportar PDF.")

    doc = SimpleDocTemplate(str(filename), pagesize=letter)
    styles = getSampleStyleSheet()
    elements: List[Any] = []

    generated_local = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    elements.append(Paragraph(f"{APP_NAME} — Informe forense", styles["Title"]))
    elements.append(Paragraph(f"Versión: {APP_VERSION}", styles["Normal"]))
    elements.append(Paragraph(f"Generado: {generated_local}", styles["Normal"]))
    elements.append(Paragraph(" ", styles["Normal"]))

    # Resumen general
    total_files = len(artifacts)
    total_emails = sum(a.count for a in artifacts)
    total_atts = sum(a.attachments_total for a in artifacts)
    total_dups = sum(a.duplicates_by_hash for a in artifacts)

    elements.append(Paragraph("Resumen general", styles["Heading2"]))
    elements.append(
        Paragraph(
            f"Archivos: <b>{total_files}</b> • Correos: <b>{total_emails}</b> • "
            f"Adjuntos: <b>{total_atts}</b> • Duplicados (hash): <b>{total_dups}</b>",
            styles["Normal"],
        )
    )
    elements.append(Paragraph(" ", styles["Normal"]))

    # Tabla resumen archivos
    elements.append(Paragraph("Archivos analizados", styles["Heading2"]))
    summary_data = [["Archivo", "Emails", "Primera fecha (UTC)", "Última fecha (UTC)", "Adjuntos", "Duplicados"]]
    for a in artifacts:
        summary_data.append(
            [
                Path(a.file).name,
                str(a.count),
                a.first_date_utc_iso,
                a.last_date_utc_iso,
                str(a.attachments_total),
                str(a.duplicates_by_hash),
            ]
        )

    summary_table = PdfTable(summary_data, repeatRows=1)
    summary_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ]
        )
    )
    elements.append(summary_table)
    elements.append(PageBreak())

    # Detalle por archivo
    for a in artifacts:
        elements.append(Paragraph(f"Detalle: {Path(a.file).name}", styles["Heading2"]))
        elements.append(
            Paragraph(
                f"Rango (UTC): {a.first_date_utc_iso} → {a.last_date_utc_iso}<br/>"
                f"Adjuntos: {a.attachments_total} • Duplicados (hash): {a.duplicates_by_hash}<br/>"
                f"Top remitentes: {', '.join(a.top_senders)}<br/>"
                f"Top destinatarios: {', '.join(a.top_recipients)}<br/>"
                f"Top dominios (From): {', '.join(a.top_sender_domains)}<br/>"
                f"Top asuntos: {', '.join(a.top_subjects)}",
                styles["Normal"],
            )
        )
        elements.append(Paragraph(" ", styles["Normal"]))

        data_emails = [["#", "Fecha (UTC)", "De", "Para", "Asunto", "Adjuntos", "Message-ID"]]
        for e in a.emails:
            data_emails.append(
                [
                    str(e.id),
                    (e.date_utc_iso or e.date_display)[:28],
                    (e.from_ or "")[:45],
                    (e.to or "")[:45],
                    (e.subject or "")[:60],
                    str(len(e.attachments)),
                    (e.message_id or "")[:36],
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

        if a.warnings:
            elements.append(Paragraph("Warnings del archivo:", styles["Heading3"]))
            for w in a.warnings[:40]:
                elements.append(Paragraph(f"• {html.escape(w)}", styles["Normal"]))

        elements.append(PageBreak())

    doc.build(elements)


# -----------------------------
# CLI
# -----------------------------
def build_arg_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="recovermail.py",
        description="Analizador forense de correos MBOX. Genera HTML/PDF/JSON sin modificar el archivo original.",
    )
    p.add_argument(
        "inputs",
        nargs="+",
        help="Archivos y/o carpetas a analizar (acepta glob tipo *.mbox).",
    )
    p.add_argument(
        "-o",
        "--output",
        default="informe_mbox",
        help="Prefijo de salida (sin extensión).",
    )
    p.add_argument(
        "--outdir",
        default=".",
        help="Directorio de salida (se crea si no existe).",
    )
    p.add_argument(
        "--recursive",
        action="store_true",
        help="Si pasas directorios, buscar también dentro de subcarpetas.",
    )
    p.add_argument(
        "--max-body-chars",
        type=int,
        default=DEFAULT_MAX_BODY_CHARS,
        help="Máximo de caracteres del cuerpo almacenados en HTML/JSON (0 = sin límite).",
    )
    p.add_argument(
        "--top",
        type=int,
        default=DEFAULT_TOP_N,
        help="Cuántos elementos mostrar en tops (remitentes/asuntos/dominios).",
    )
    p.add_argument(
        "--no-body",
        action="store_true",
        help="No extraer cuerpo (más rápido y JSON/HTML más ligeros).",
    )
    p.add_argument(
        "--prefer-html",
        action="store_true",
        help="Si existe, prioriza body HTML (convertido a texto) sobre text/plain.",
    )
    p.add_argument(
        "--no-pdf",
        action="store_true",
        help="No generar PDF.",
    )
    p.add_argument(
        "--no-html",
        action="store_true",
        help="No generar HTML.",
    )
    p.add_argument(
        "--no-json",
        action="store_true",
        help="No generar JSON.",
    )
    return p


def main(argv: Optional[List[str]] = None) -> int:
    console = _console()
    if console:
        print_banner(console)

    args = build_arg_parser().parse_args(argv)

    outdir = Path(args.outdir).expanduser().resolve()
    outdir.mkdir(parents=True, exist_ok=True)

    include_body = not args.no_body
    prefer_plain = not args.prefer_html

    # Reunir candidatos
    candidates = list(iter_input_paths(args.inputs, recursive=args.recursive))

    if not candidates:
        if console:
            console.print("[red]No se encontraron rutas válidas para analizar.[/red]")
        else:
            print("No se encontraron rutas válidas para analizar.", file=sys.stderr)
        return 2

    # Filtrar MBOX
    mbox_files: List[Path] = []
    missing: List[str] = []
    for p in candidates:
        if not p.exists():
            missing.append(str(p))
            continue
        if is_mbox(p):
            mbox_files.append(p)

    if not mbox_files:
        if console:
            console.print("[bold yellow]No se detectaron archivos MBOX válidos.[/bold yellow]")
        else:
            print("No se detectaron archivos MBOX válidos.", file=sys.stderr)
        return 1

    artifacts: List[MboxArtifact] = []
    total_emails = 0

    def analyze_one(path: Path) -> Optional[MboxArtifact]:
        return analyze_mbox(
            path,
            max_body_chars=args.max_body_chars,
            top_n=max(1, args.top),
            include_body=include_body,
            prefer_plain=prefer_plain,
        )

    if console and Progress is not None:
        with Progress(
            SpinnerColumn(),
            TextColumn("[bold]Analizando[/bold] {task.fields[file]}"),
            BarColumn(),
            TextColumn("{task.completed}/{task.total}"),
            TimeElapsedColumn(),
            console=console,
        ) as progress:
            task = progress.add_task("mbox", total=len(mbox_files), file="")
            for f in mbox_files:
                progress.update(task, file=f.name)
                art = analyze_one(f)
                if art:
                    artifacts.append(art)
                    total_emails += art.count
                progress.advance(task)
    else:
        for f in mbox_files:
            art = analyze_one(f)
            if art:
                artifacts.append(art)
                total_emails += art.count

    if not artifacts:
        if console:
            console.print("[bold yellow]No se pudieron extraer correos de los MBOX.[/bold yellow]")
        else:
            print("No se pudieron extraer correos de los MBOX.", file=sys.stderr)
        return 1

    # Resumen consola
    if console:
        console.print(
            f"\n[bold green]Análisis completado:[/bold green] {len(artifacts)} archivo(s), {total_emails} correo(s) procesados."
        )
        print_summary_table(console, artifacts)
    else:
        print(f"Análisis completado: {len(artifacts)} archivo(s), {total_emails} correo(s).")

    # Exportaciones
    base = args.output
    pdf_file = outdir / f"{base}.pdf"
    html_file = outdir / f"{base}.html"
    json_file = outdir / f"{base}.json"

    # JSON
    if not args.no_json:
        export_json(artifacts, json_file)

    # HTML
    if not args.no_html:
        export_html(artifacts, html_file, max_body_chars=args.max_body_chars)

    # PDF
    if not args.no_pdf:
        try:
            export_pdf(artifacts, pdf_file)
        except Exception as e:
            if console:
                console.print(f"[yellow]No se pudo generar PDF: {e}[/yellow]")
            else:
                print(f"No se pudo generar PDF: {e}", file=sys.stderr)

    # Mensajes finales
    if console:
        console.print("\n[bold green]Resultados exportados:[/bold green]")
        if not args.no_html:
            console.print(f"[cyan]► HTML:[/cyan] {html_file}")
        if not args.no_pdf:
            console.print(f"[cyan]► PDF:[/cyan]  {pdf_file}")
        if not args.no_json:
            console.print(f"[cyan]► JSON:[/cyan] {json_file}")
    else:
        print("Resultados exportados en:", outdir)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
