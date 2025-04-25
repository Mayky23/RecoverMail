import argparse
import mailbox
import os
import json
from pathlib import Path
from datetime import datetime
from rich.console import Console
from rich.table import Table
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table as PdfTable, Paragraph
from reportlab.lib.styles import getSampleStyleSheet

console = Console()

def print_banner():
    banner = r"""
 ____                                 __  __       _ _     
|  _ \ ___  ___ _____   _____ _ __   |  \/  | __ _(_) |    
| |_) / _ \/ __/ _ \ \ / / _ \ '__|  | |\/| |/ _` | | |    
|  _ <  __/ (_| (_) \ V /  __/ |     | |  | | (_| | | |    
|_| \_\___|\___\___/ \_/ \___|_|     |_|  |_|\__,_|_|_|    
                                                           
[bold blue]Herramienta Forense de Recuperación de Correos Electrónicos[/bold blue]    
[bold cyan]v3.0[/bold cyan] • Soporte MBOX • By: MARH    
"""
    console.print(banner, justify="center")

def is_mbox(filepath):
    """Detecta si un archivo es MBOX aunque no tenga extensión"""
    try:
        with open(filepath, 'rb') as f:
            first_line = f.readline().decode('utf-8', errors='ignore')
            return first_line.startswith('From ') and ' ' in first_line[5:20]
    except:
        return False

def analyze_mbox(mbox_path):
    try:
        mbox = mailbox.mbox(mbox_path)
        emails = list(mbox)
        
        if not emails:
            return None
            
        dates = []
        all_emails = []
        
        for i, msg in enumerate(emails):
            email_data = {
                'id': i+1,
                'from': msg['from'] or "Desconocido",
                'to': msg['to'] or "Desconocido",
                'subject': msg['subject'] or "Sin asunto",
                'date': msg['date'] or "Fecha desconocida",
                'cc': msg['cc'] or "",
                'content_type': msg.get_content_type() if hasattr(msg, 'get_content_type') else "Desconocido"
            }
            
            # Intentar extraer el cuerpo del mensaje
            try:
                if msg.is_multipart():
                    for part in msg.get_payload():
                        if part.get_content_type() == 'text/plain':
                            email_data['body'] = part.get_payload(decode=True).decode('utf-8', errors='replace')[:500] + "..."
                            break
                else:
                    email_data['body'] = msg.get_payload(decode=True).decode('utf-8', errors='replace')[:500] + "..."
            except:
                email_data['body'] = "No se pudo extraer el contenido"
            
            all_emails.append(email_data)
            
            if msg['date']:
                try:
                    dt = datetime.strptime(msg['date'][:25], '%a, %d %b %Y %H:%M:%S')
                    dates.append(dt)
                except ValueError:
                    continue
        
        # Datos para el resumen en terminal
        summary = {
            'count': len(emails),
            'first_date': min(dates).strftime('%Y-%m-%d %H:%M:%S') if dates else "N/D",
            'last_date': max(dates).strftime('%Y-%m-%d %H:%M:%S') if dates else "N/D",
            'subjects': [msg['subject'] or "Sin asunto" for msg in emails][:3],
            'senders': list({msg['from'] for msg in emails if msg['from']})[:3],
            'all_emails': all_emails  # Todos los correos con detalles
        }
        
        return summary
    except Exception as e:
        console.print(f"[red]Error analizando MBOX ({mbox_path}): {str(e)}[/red]")
        return None

def generate_report(artifacts):
    table = Table(title="Resumen de Análisis Forense", show_lines=True)
    table.add_column("Archivo", style="magenta")
    table.add_column("Emails", style="green")
    table.add_column("Primera fecha", style="yellow")
    table.add_column("Última fecha", style="yellow")
    table.add_column("Asuntos", style="blue")
    table.add_column("Remitentes", style="red")
    
    for art in artifacts:
        table.add_row(
            art['file'],
            str(art['count']),
            art['first_date'],
            art['last_date'],
            "\n".join(art.get('subjects', ['N/D'])),
            "\n".join(art.get('senders', ['N/D']))
        )
    
    console.print(table)
    return table

def export_pdf(artifacts, filename):
    try:
        doc = SimpleDocTemplate(filename, pagesize=letter)
        styles = getSampleStyleSheet()
        elements = []
        
        # Título del informe
        elements.append(Paragraph("Informe RecoverMail", styles['Title']))
        elements.append(Paragraph(f"Generado el {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", styles['Normal']))
        elements.append(Paragraph(" ", styles['Normal']))  # Espacio
        
        # Tabla resumen
        elements.append(Paragraph("Resumen de Archivos Analizados", styles['Heading2']))
        data_summary = [["Archivo", "Emails", "Primera fecha", "Última fecha"]]
        for art in artifacts:
            data_summary.append([
                art['file'],
                str(art['count']),
                art['first_date'],
                art['last_date']
            ])
        
        summary_table = PdfTable(data_summary)
        summary_table.setStyle([
            ('BACKGROUND', (0,0), (-1,0), '#CCCCCC'),
            ('GRID', (0,0), (-1,-1), 1, '#000000'),
            ('FONTSIZE', (0,0), (-1,-1), 10),
            ('VALIGN', (0,0), (-1,-1), 'TOP')
        ])
        elements.append(summary_table)
        elements.append(Paragraph(" ", styles['Normal']))  # Espacio
        
        # Detalles de todos los correos
        for art in artifacts:
            elements.append(Paragraph(f"Correos en archivo: {art['file']}", styles['Heading3']))
            
            # Tabla de emails
            emails = art.get('all_emails', [])
            if emails:
                data_emails = [["ID", "Remitente", "Destinatario", "Asunto", "Fecha"]]
                
                for email in emails:
                    data_emails.append([
                        str(email['id']),
                        email['from'][:30] + ('...' if len(email['from']) > 30 else ''),
                        email['to'][:30] + ('...' if len(email['to']) > 30 else ''),
                        email['subject'][:30] + ('...' if len(email['subject']) > 30 else ''),
                        email['date'][:30]
                    ])
                
                email_table = PdfTable(data_emails)
                email_table.setStyle([
                    ('BACKGROUND', (0,0), (-1,0), '#E0E0E0'),
                    ('GRID', (0,0), (-1,-1), 0.5, '#999999'),
                    ('FONTSIZE', (0,0), (-1,-1), 8),
                    ('VALIGN', (0,0), (-1,-1), 'TOP')
                ])
                elements.append(email_table)
            else:
                elements.append(Paragraph("No se encontraron correos", styles['Normal']))
            
            elements.append(Paragraph(" ", styles['Normal']))  # Espacio
        
        doc.build(elements)
        console.print(f"[green]PDF exportado: {filename}[/green]")
    except Exception as e:
        console.print(f"[red]Error generando PDF: {str(e)}[/red]")

def export_html(artifacts, filename):
    try:
        html_content = """
        <!DOCTYPE html>
        <html>
            <head>
                <title>Informe RecoverMail</title>
                <meta charset="UTF-8">
                <style>
                    body {font-family: Arial, sans-serif; margin: 0; padding: 0;}
                    .container {width: 95%; margin: 0 auto; padding: 20px;}
                    table {border-collapse: collapse; width: 100%; margin-top: 20px; margin-bottom: 30px;}
                    th, td {border: 1px solid #ddd; padding: 8px; text-align: left; font-size: 14px;}
                    th {background-color: #f2f2f2; position: sticky; top: 0;}
                    tr:nth-child(even) {background-color: #f9f9f9;}
                    .header {background-color: #004080; color: white; padding: 15px; text-align: center; margin-bottom: 20px;}
                    h2, h3 {color: #004080;}
                    .summary {margin-bottom: 30px;}
                    .email-detail {margin-bottom: 20px; padding: 10px; border: 1px solid #ddd;}
                    .email-header {background-color: #eaeaea; padding: 8px;}
                    .email-body {padding: 10px; white-space: pre-wrap; font-family: monospace; max-height: 200px; overflow-y: auto;}
                    .accordion {background-color: #eee; color: #444; cursor: pointer; padding: 18px; width: 100%; text-align: left; 
                               border: none; outline: none; transition: 0.4s; margin-bottom: 10px; font-size: 16px;}
                    .active, .accordion:hover {background-color: #ccc;}
                    .panel {padding: 0 18px; background-color: white; max-height: 0; overflow: hidden; transition: max-height 0.2s ease-out;}
                </style>
                <script>
                    function toggleAccordion(id) {
                        const panel = document.getElementById(id);
                        if (panel.style.maxHeight) {
                            panel.style.maxHeight = null;
                        } else {
                            panel.style.maxHeight = panel.scrollHeight + "px";
                        }
                    }
                </script>
            </head>
            <body>
                <div class="header">
                    <h1>Informe de Análisis Forense de Correos</h1>
                    <p>Generado el """ + datetime.now().strftime('%Y-%m-%d %H:%M:%S') + """</p>
                </div>
                
                <div class="container">
                    <div class="summary">
                        <h2>Resumen de Archivos Analizados</h2>
                        <table>
                            <tr>
                                <th>Archivo</th>
                                <th>Total Correos</th>
                                <th>Primera fecha</th>
                                <th>Última fecha</th>
                            </tr>
        """
        
        for art in artifacts:
            html_content += f"""
                            <tr>
                                <td>{art['file']}</td>
                                <td>{art['count']}</td>
                                <td>{art['first_date']}</td>
                                <td>{art['last_date']}</td>
                            </tr>
            """
        
        html_content += """
                        </table>
                    </div>
        """
        
        # Añadir todos los correos de cada archivo
        for i, art in enumerate(artifacts):
            html_content += f"""
                    <button class="accordion" onclick="toggleAccordion('panel-{i}')">
                        Ver detalles de {art['file']} ({art['count']} correos)
                    </button>
                    <div class="panel" id="panel-{i}">
                        <h3>Correos en archivo: {art['file']}</h3>
                        <table>
                            <tr>
                                <th>ID</th>
                                <th>Remitente</th>
                                <th>Destinatario</th>
                                <th>Asunto</th>
                                <th>Fecha</th>
                                <th>Acciones</th>
                            </tr>
            """
            
            emails = art.get('all_emails', [])
            for email in emails:
                # Escape HTML characters in strings
                from_escaped = email['from'].replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                to_escaped = email['to'].replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                subject_escaped = email['subject'].replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                date_escaped = email['date'].replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                body_escaped = email['body'].replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                
                email_id = f"email-{i}-{email['id']}"
                
                html_content += f"""
                            <tr>
                                <td>{email['id']}</td>
                                <td>{from_escaped}</td>
                                <td>{to_escaped}</td>
                                <td>{subject_escaped}</td>
                                <td>{date_escaped}</td>
                                <td><button onclick="toggleAccordion('{email_id}')">Ver contenido</button></td>
                            </tr>
                            <tr>
                                <td colspan="6" style="padding: 0;">
                                    <div class="panel" id="{email_id}">
                                        <div class="email-detail">
                                            <div class="email-header">
                                                <p><strong>De:</strong> {from_escaped}</p>
                                                <p><strong>Para:</strong> {to_escaped}</p>
                                                <p><strong>CC:</strong> {email.get('cc', '')}</p>
                                                <p><strong>Asunto:</strong> {subject_escaped}</p>
                                                <p><strong>Fecha:</strong> {date_escaped}</p>
                                                <p><strong>Tipo:</strong> {email.get('content_type', 'Desconocido')}</p>
                                            </div>
                                            <div class="email-body">{body_escaped}</div>
                                        </div>
                                    </div>
                                </td>
                            </tr>
                """
            
            html_content += """
                        </table>
                    </div>
            """
        
        html_content += """
                </div>
            </body>
        </html>
        """
        
        with open(filename, 'w', encoding='utf-8') as f:
            f.write(html_content)
        console.print(f"[green]HTML exportado: {filename}[/green]")
    except Exception as e:
        console.print(f"[red]Error generando HTML: {str(e)}[/red]")

def export_json(artifacts, filename):
    try:
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(artifacts, f, ensure_ascii=False, indent=2)
        console.print(f"[green]JSON exportado: {filename}[/green]")
    except Exception as e:
        console.print(f"[red]Error generando JSON: {str(e)}[/red]")

def main():
    print_banner()
    parser = argparse.ArgumentParser(description='Analizador forense de correos MBOX')
    parser.add_argument('files', nargs='+', help='Archivos MBOX a analizar')
    parser.add_argument('--output', '-o', default='informe_mbox', help='Prefijo para los archivos de salida')
    args = parser.parse_args()
    
    artifacts = []
    total_correos = 0
    
    with console.status("[bold green]Analizando archivos de correo...") as status:
        for file_path in args.files:
            path = Path(file_path)
            if not path.exists():
                console.print(f"[red]Archivo no encontrado: {file_path}[/red]")
                continue
                
            try:
                if is_mbox(file_path) or path.suffix.lower() == '.mbox':
                    console.print(f"[blue]Analizando MBOX: {file_path}[/blue]")
                    analysis = analyze_mbox(path)
                    if analysis:
                        artifacts.append({
                            'file': path.name,
                            **analysis
                        })
                        total_correos += analysis['count']
                else:
                    console.print(f"[yellow]No es un archivo MBOX válido: {file_path}[/yellow]")
            except Exception as e:
                console.print(f"[red]Error procesando archivo {file_path}: {str(e)}[/red]")
    
    if artifacts:
        console.print(f"[bold green]Análisis completado. Se encontraron [bold]{total_correos}[/bold] correos en [bold]{len(artifacts)}[/bold] archivos.[/bold green]")
        
        # Mostrar resumen en la terminal
        console.print(f"\n[bold cyan]Resumen de análisis:[/bold cyan]")
        generate_report(artifacts)
        
        # Exportar resultados completos
        pdf_file = f"{args.output}.pdf"
        html_file = f"{args.output}.html"
        json_file = f"{args.output}.json"
        
        with console.status("[bold yellow]Exportando resultados...") as status:
            export_pdf(artifacts, pdf_file)
            export_html(artifacts, html_file)
            export_json(artifacts, json_file)
        
        console.print(f"\n[bold green]Resultados exportados correctamente:[/bold green]")
        console.print(f"[cyan]► HTML: [/cyan]{html_file}")
        console.print(f"[cyan]► PDF: [/cyan]{pdf_file}")
        console.print(f"[cyan]► JSON: [/cyan]{json_file}")
    else:
        console.print("[yellow]No se encontraron archivos MBOX válidos para analizar[/yellow]")

if __name__ == "__main__":
    main()