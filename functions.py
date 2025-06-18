import pandas as pd
import os
from pathlib import Path
import argparse
from datetime import datetime

def read_excel_file(excel_file):
    """
    Liest die Excel-Datei gemäß den Angaben in excel_struktur.txt
    
    Die Funktion erkennt die Spalten anhand ihrer Bezeichner in Zeile 2,
    nicht anhand ihrer Position.
    
    Rückgabewert: DataFrame mit den relevanten Spalten und Metadaten
    """
    try:
        print(f"Lese Excel-Datei: {excel_file}")
        
        # Lese die ersten Zeilen ohne Header, um die Struktur zu verstehen
        df_raw = pd.read_excel(excel_file, header=None, nrows=5)
        
        # Extrahiere die Anzahl der Nachrichten aus der ersten Zeile, Spalte B
        messages_count_text = df_raw.iloc[0, 1]  # Zeile 1, Spalte B (0-basiert)
        
        messages_count = 0
        if isinstance(messages_count_text, str) and "(" in messages_count_text and ")" in messages_count_text:
            # Extrahiere die Zahl zwischen Klammern
            start = messages_count_text.find("(") + 1
            end = messages_count_text.find(")")
            if start > 0 and end > start:
                try:
                    messages_count = int(messages_count_text[start:end])
                except ValueError:
                    pass
        
        # Extrahiere die Spaltenbezeichner aus der zweiten Zeile
        header_row = df_raw.iloc[1].tolist()
        
        # Erstelle ein Dictionary, das die Spaltenbezeichner den Spaltenindizes zuordnet
        header_to_index = {}
        for i, header in enumerate(header_row):
            if pd.notna(header) and isinstance(header, str) and header.strip():
                header_to_index[header.strip()] = i
        
        # Definiere die benötigten Spalten laut excel_struktur.txt
        required_columns = [
            "#", "From", "To", "Direction", "Body", "Status", "Transcript",
            "Timestamp-Date", "Timestamp-Time", "Attachment #1", 
            "Attachment #1 - Details", "Deleted", "Label", "Starred message"
        ]
        
        # Überprüfe, ob alle benötigten Spalten vorhanden sind
        missing_columns = [col for col in required_columns if col not in header_to_index]
        if missing_columns:
            print(f"Warnung: Folgende Spalten fehlen in der Excel-Datei: {missing_columns}")
        
        # Lese die gesamte Excel-Datei mit den korrekten Spaltenbezeichnern
        # Wir verwenden die zweite Zeile als Header
        df = pd.read_excel(excel_file, header=1)
        
        # Erstelle ein neues DataFrame mit den benötigten Spalten
        result_df = pd.DataFrame()
        
        # Füge die benötigten Spalten hinzu, wenn sie vorhanden sind
        for col_name in required_columns:
            if col_name in df.columns:
                result_df[col_name] = df[col_name]
            else:
                result_df[col_name] = None
        
        # Füge Metadaten hinzu
        metadata = {
            "messages_count": messages_count,
            "actual_rows": len(df),
            "file_name": Path(excel_file).name,
            "import_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        
        return result_df, metadata
        
    except Exception as e:
        print(f"Fehler beim Lesen der Excel-Datei: {str(e)}")
        return None, {"error": str(e)}

import os
import re
import urllib.parse
from pathlib import Path

def is_url(text):
    """
    Überprüft, ob ein Text eine URL ist
    """
    if not text or pd.isna(text) or text == "":
        return False
        
    # Prüfe auf Dateinamen mit typischen Dateiendungen, die keine URLs sind
    file_extensions = [".jpg", ".jpeg", ".png", ".gif", ".mp4", ".mp3", ".wav", ".ogg", ".opus", ".pdf", ".doc", ".docx"]
    if any(text.lower().endswith(ext) for ext in file_extensions):
        return False
    
    # Prüfe auf Dateinamen mit Dateiendung und ohne URL-Protokoll
    if re.match(r'^[\w\-. ]+\.[a-zA-Z0-9]{2,4}$', text):
        return False
        
    # Strenge URL-Erkennung mit gängigen Protokollen
    url_pattern = re.compile(
        r'^(?:http|https|ftp|mailto|tel|file|data)s?://'
        r'(?:(?:[A-Z0-9](?:[A-Z0-9-]{0,61}[A-Z0-9])?\.)+(?:[A-Z]{2,6}\.?|[A-Z0-9-]{2,}\.?)|'
        r'localhost|'
        r'\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})'
        r'(?::\d+)?'
        r'(?:/?|[/?]\S+)$', re.IGNORECASE)
    
    # Prüfe, ob der Text dem URL-Muster entspricht
    if url_pattern.match(text):
        return True
        
    # Prüfe auf eindeutige URL-Präfixe ohne Protokoll
    if text.lower().startswith(('www.', 'http:', 'https:', 'youtu.be/', 't.co/', 'bit.ly/')):
        return True
    
    # Prüfe auf bekannte Domains
    known_domains = ['facebook.com', 'youtube.com', 'twitter.com', 'instagram.com', 'whatsapp.com']
    if any(domain in text.lower() for domain in known_domains):
        return True
        
    # Vorsichtigere URL-Kodierungsprüfung
    try:
        parsed = urllib.parse.urlparse(text)
        # Nur wenn netloc vorhanden ist oder der Pfad eindeutig eine URL ist
        if parsed.netloc and '.' in parsed.netloc:
            return True
        # Oder wenn es ein Schema hat und einen Pfad
        if parsed.scheme and parsed.path:
            return True
    except:
        pass
        
    return False

def check_attachment_exists(excel_path, attachment_name):
    """
    Überprüft, ob ein Anhang existiert
    """
    if not attachment_name or pd.isna(attachment_name) or attachment_name == "":
        return False
        
    # Prüfe, ob es sich um eine URL handelt
    if is_url(attachment_name):
        return "URL"  # Spezialwert für URLs
        
    # Verzeichnis mit der Excel-Datei
    excel_dir = Path(excel_path).parent
    
    # Verzeichnisse für Dateianhänge
    search_dirs = [
        excel_dir / 'files',                     # Standard-Verzeichnis
        excel_dir / 'instant_messages',          # Unterverzeichnis 'instant_messages'
        excel_dir                                # Hauptverzeichnis
    ]
    
    # Suche in allen möglichen Verzeichnissen
    for search_dir in search_dirs:
        if not search_dir.exists():
            continue
            
        # Suche nach der Datei in allen Unterverzeichnissen
        for root, _, files in os.walk(search_dir):
            if attachment_name in files:
                return os.path.join(root, attachment_name)
    
    return False

def categorize_attachment(attachment_name):
    """
    Kategorisiert einen Anhang basierend auf der Dateiendung oder URL
    """
    if not attachment_name or pd.isna(attachment_name) or attachment_name == "":
        return "Unbekannt"
    
    # Prüfe, ob es sich um eine URL handelt
    if is_url(attachment_name):
        # Versuche, die Domain zu extrahieren
        try:
            parsed = urllib.parse.urlparse(attachment_name)
            domain = parsed.netloc
            if not domain and attachment_name.startswith('www.'):
                domain = attachment_name.split('/', 1)[0]
                
            # Bekannte Domains kategorisieren
            if any(site in domain.lower() for site in ['youtube', 'youtu.be']):
                return "URL (YouTube)"
            elif any(site in domain.lower() for site in ['facebook', 'fb.com']):
                return "URL (Facebook)"
            elif any(site in domain.lower() for site in ['instagram', 'insta']):
                return "URL (Instagram)"
            elif any(site in domain.lower() for site in ['twitter', 'x.com', 't.co']):
                return "URL (Twitter/X)"
            elif any(site in domain.lower() for site in ['tiktok']):
                return "URL (TikTok)"
            elif any(site in domain.lower() for site in ['whatsapp']):
                return "URL (WhatsApp)"
            elif any(site in domain.lower() for site in ['amazon']):
                return "URL (Amazon)"
            elif any(site in domain.lower() for site in ['google']):
                return "URL (Google)"
            else:
                return "URL"
        except:
            return "URL"
    
    # Dateiendung extrahieren
    _, ext = os.path.splitext(attachment_name.lower())
    
    # Detaillierte Kategorien definieren
    # Bilder
    jpeg_extensions = [".jpg", ".jpeg"]
    png_extensions = [".png"]
    gif_extensions = [".gif"]
    other_image_extensions = [".bmp", ".webp", ".tiff", ".svg"]
    
    # Videos
    mp4_extensions = [".mp4"]
    other_video_extensions = [".mov", ".avi", ".mkv", ".wmv", ".webm", ".flv", ".3gp"]
    
    # Audio
    mp3_extensions = [".mp3"]
    voice_extensions = [".ogg", ".m4a", ".aac", ".opus"]
    other_audio_extensions = [".wav", ".flac", ".wma"]
    
    # Dokumente
    pdf_extensions = [".pdf"]
    office_extensions = [".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx"]
    text_extensions = [".txt", ".rtf", ".md"]
    
    # Andere Dateitypen
    archive_extensions = [".zip", ".rar", ".7z", ".tar", ".gz"]
    executable_extensions = [".exe", ".msi", ".apk"]
    
    # Kategorisierung
    if ext in jpeg_extensions:
        return "Bild (JPEG)"
    elif ext in png_extensions:
        return "Bild (PNG)"
    elif ext in gif_extensions:
        return "Bild (GIF)"
    elif ext in other_image_extensions:
        return "Bild (Andere)"
    elif ext in mp4_extensions:
        return "Video (MP4)"
    elif ext in other_video_extensions:
        return "Video (Andere)"
    elif ext in mp3_extensions:
        return "Audio (MP3)"
    elif ext in voice_extensions:
        return "Audio (Sprachnachricht)"
    elif ext in other_audio_extensions:
        return "Audio (Andere)"
    elif ext in pdf_extensions:
        return "Dokument (PDF)"
    elif ext in office_extensions:
        return "Dokument (Office)"
    elif ext in text_extensions:
        return "Dokument (Text)"
    elif ext in archive_extensions:
        return "Archiv"
    elif ext in executable_extensions:
        return "Ausführbare Datei"
    elif ext == "":
        return "Ohne Dateiendung"
    else:
        return f"Sonstige ({ext})"

def generate_statistics(df, metadata, excel_path=None, verbose=False):
    """
    Generiert Statistiken aus dem DataFrame und den Metadaten
    
    Parameter:
    - df: DataFrame mit den Daten
    - metadata: Dictionary mit Metadaten
    - excel_path: Pfad zur Excel-Datei (für Anhangsuche)
    - verbose: Wenn True, werden detaillierte Informationen zu jedem Anhang angezeigt
    """
    stats = []
    stats.append("=== Excel-Datei Statistik ===")
    stats.append(f"Dateiname: {metadata.get('file_name', 'Unbekannt')}")
    stats.append(f"Importiert am: {metadata.get('import_date', 'Unbekannt')}")
    stats.append(f"Anzahl Nachrichten (aus Header): {metadata.get('messages_count', 0)}")
    stats.append(f"Tatsächliche Anzahl Zeilen: {metadata.get('actual_rows', 0)}")
    
    # Prüfe, ob die Spalten vorhanden sind, bevor wir sie analysieren
    if df is not None:
        # Anzahl der Nachrichten mit Anhängen
        if "Attachment #1" in df.columns:
            # Erstelle eine Liste mit Informationen zu jedem Anhang
            attachment_info_list = []
            primary_attachments = 0  # Anhänge ohne Nachrichtentext
            supplementary_attachments = 0  # Anhänge mit Nachrichtentext
            
            for index, row in df.iterrows():
                attachment = row.get("Attachment #1")
                if pd.notna(attachment) and attachment:
                    line_number = row.get("#", index + 1)  # Verwende Spalte "#" oder Index+1 als Fallback
                    sender = row.get("From", "Unbekannt")
                    timestamp = row.get("Timestamp-Time", "")
                    direction = row.get("Direction", "")
                    body = row.get("Body", "")
                    
                    # Prüfe, ob ein Nachrichtentext vorhanden ist
                    has_message_text = pd.notna(body) and body and body.strip() != ""
                    attachment_type = "Ergänzung" if has_message_text else "Primär"
                    
                    # Zähle die Anhänge nach Typ
                    if has_message_text:
                        supplementary_attachments += 1
                    else:
                        primary_attachments += 1
                    
                    # Überprüfe, ob der Anhang existiert
                    attachment_path = None
                    if excel_path:
                        attachment_path = check_attachment_exists(excel_path, attachment)
                    
                    attachment_info = {
                        "line_number": line_number,
                        "attachment": attachment,
                        "sender": sender,
                        "timestamp": timestamp,
                        "direction": direction,
                        "body": body if has_message_text else "",
                        "has_message_text": has_message_text,
                        "attachment_type": attachment_type,
                        "category": categorize_attachment(attachment),
                        "exists": bool(attachment_path),
                        "path": attachment_path
                    }
                    attachment_info_list.append(attachment_info)
            
            attachments_count = len(attachment_info_list)
            stats.append(f"Nachrichten mit Anhängen: {attachments_count}")
            stats.append(f"  Primäre Anhänge (ohne Text): {primary_attachments}")
            stats.append(f"  Ergänzende Anhänge (mit Text): {supplementary_attachments}")
            
            # Zähle vorhandene und fehlende Anhänge
            if excel_path:
                url_count = 0
                existing_attachments = 0
                missing_attachments = 0
                
                for info in attachment_info_list:
                    if info["path"] == "URL":
                        url_count += 1
                    elif info["exists"]:
                        existing_attachments += 1
                    else:
                        missing_attachments += 1
                
                stats.append(f"Vorhandene Anhänge: {existing_attachments}")
                stats.append(f"URLs/Links: {url_count}")
                stats.append(f"Fehlende Anhänge: {missing_attachments}")
                
                # Kategorisiere die Anhänge
                categories = {}
                attachment_extensions = {}
                directories = {}
                
                for info in attachment_info_list:
                    # Zähle Kategorien
                    category = info["category"]
                    categories[category] = categories.get(category, 0) + 1
                    
                    # Unterscheide zwischen URLs und Dateien
                    if info["path"] == "URL" or category.startswith("URL"):
                        # Für URLs keine Dateiendung zählen
                        pass
                    else:
                        # Zähle Dateiendungen nur für echte Dateien
                        _, ext = os.path.splitext(info["attachment"].lower())
                        if ext:
                            attachment_extensions[ext] = attachment_extensions.get(ext, 0) + 1
                    
                    # Zähle Verzeichnisse nur für echte Dateien, nicht für URLs
                    if info["path"] and info["path"] != "URL":
                        directory = os.path.dirname(info["path"])
                        directories[directory] = directories.get(directory, 0) + 1
                
                # Ausgabe der Kategorien
                stats.append("\nAnhangskategorien:")
                for category, count in sorted(categories.items(), key=lambda x: x[1], reverse=True):
                    stats.append(f"  {category}: {count}")
                
                # Ausgabe der häufigsten Dateiendungen
                if attachment_extensions:
                    stats.append("\nDateiendungen:")
                    for ext, count in sorted(attachment_extensions.items(), key=lambda x: x[1], reverse=True)[:10]:  # Top 10
                        stats.append(f"  {ext}: {count}")
                
                # Analyse der Verzeichnisse, in denen Anhänge gefunden wurden
                if directories:
                    stats.append("\nVerzeichnisse mit Anhängen:")
                    for directory, count in sorted(directories.items(), key=lambda x: x[1], reverse=True)[:5]:  # Top 5
                        stats.append(f"  {directory}: {count} Dateien")
                
                # Im Verbose-Modus alle Anhänge auflisten
                if verbose and attachment_info_list:
                    stats.append("\n=== Detaillierte Anhangsübersicht ===")
                    stats.append("Zeile | Sender | Zeitstempel | Richtung | Typ | Anhang | Kategorie | Status")
                    stats.append("-" * 120)
                    
                    # Sortiere nach Zeilennummer
                    for info in sorted(attachment_info_list, key=lambda x: x["line_number"]):
                        line = info["line_number"]
                        sender = info["sender"][:20] if len(str(info["sender"])) > 20 else info["sender"]
                        timestamp = info["timestamp"][:19] if len(str(info["timestamp"])) > 19 else info["timestamp"]
                        direction = info["direction"]
                        attachment_type = info["attachment_type"]
                        attachment = info["attachment"][:30] if len(info["attachment"]) > 30 else info["attachment"]
                        category = info["category"]
                        status = "Vorhanden" if info["exists"] else "Fehlt"
                        
                        stats.append(f"{line:5} | {sender:20} | {timestamp:19} | {direction:8} | {attachment_type:8} | {attachment:30} | {category:20} | {status}")
                        
                        # Bei Nachrichten mit Text den Textanfang anzeigen
                        if info["has_message_text"]:
                            body_text = info["body"][:80] + "..." if len(info["body"]) > 80 else info["body"]
                            stats.append(f"       Text: {body_text}")
                        
                        # Bei vorhandenen Anhängen den Pfad anzeigen
                        if info["exists"]:
                            stats.append(f"       Pfad: {info['path']}")
                        
                        stats.append("-" * 120)
        
        # Anzahl der gelöschten Nachrichten
        if "Deleted" in df.columns:
            deleted_count = df[df["Deleted"] == "Yes"].shape[0]
            stats.append(f"Gelöschte Nachrichten: {deleted_count}")
        
        # Anzahl der markierten Nachrichten
        if "Starred message" in df.columns:
            starred_count = df["Starred message"].notna().sum()
            stats.append(f"Markierte Nachrichten: {starred_count}")
        
        # Verteilung der Nachrichtenrichtung
        if "Direction" in df.columns:
            direction_counts = df["Direction"].value_counts()
            stats.append("\nNachrichtenrichtung:")
            for direction, count in direction_counts.items():
                stats.append(f"  {direction}: {count}")
        
        # Verteilung des Nachrichtenstatus
        if "Status" in df.columns:
            status_counts = df["Status"].value_counts()
            stats.append("\nNachrichtenstatus:")
            for status, count in status_counts.items():
                stats.append(f"  {status}: {count}")
    
    return "\n".join(stats)

# This file is intended to be imported as a module only
