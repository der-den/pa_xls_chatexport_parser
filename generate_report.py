import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from datetime import datetime
from pathlib import Path
import re
from PIL import Image
import os
import argparse
import sys
import math
import traceback
# Audio/Video imports deaktiviert, um PyAudio-Abhängigkeit zu vermeiden
# import whisper
# import torch
# from pydub import AudioSegment
# from moviepy.editor import VideoFileClip
import tempfile

# Import der neuen Excel-Reader-Funktionalität
from excel_reader import read_excel_file, generate_statistics, is_url

# Global counters
attachment_not_found_counter = 0
attachment_found_counter = 0

def parse_participant(from_field):
    """
    Extrahiert Chat-ID und Namen aus dem From-Feld.
    Format ist typischerweise: "Name (ID)" oder nur "Name".
    """
    if not from_field or from_field == 'nan':
        return None, None
        
    # Suche nach dem Muster "Name (ID)"
    match = re.search(r'(.+?)\s*\(([^)]+)\)', from_field)
    if match:
        name = match.group(1).strip()
        chat_id = match.group(2).strip()
        return chat_id, name
    else:
        # Wenn kein Muster gefunden wurde, verwende das gesamte Feld als Namen
        return from_field, from_field

# Sichere Funktion zum Abrufen von Zellwerten aus einer DataFrame-Zeile
def safe_get_cell(row, idx, default='', verbose=False):
    try:
        if idx < len(row):
            value = str(row.iloc[idx]).strip()
            return value if value != 'nan' else default
        return default
    except Exception as e:
        if verbose:
            print(f"Fehler beim Lesen der Zelle {idx}: {e}")
        return default

def find_attachment_file(excel_path, attachment_name):
    """
    Search for an attachment file in multiple directories parallel to the Excel file.
    Returns the full path if found, "URL" if it's a URL, None otherwise.
    """
    global attachment_not_found_counter
    global attachment_found_counter
    
    #print(f"Searching for attachment: {attachment_name}")
    if not attachment_name or attachment_name == 'nan':
        return None
    
    # Prüfe, ob es sich um eine URL handelt
    if is_url(attachment_name):
        return "URL"
        
    # Get the directory containing the Excel file
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
                attachment_found_counter += 1
                return os.path.join(root, attachment_name)
    
    attachment_not_found_counter += 1
    #print(f"Warning: Attachment not found: {attachment_name}")
    return None

class ChatReport:
    def __init__(self, verbose=False, model_name="medium"):
        self.page_width, self.page_height = A4
        self.margin = 50
        self.line_height = 14
        self.y_position = self.page_height - self.margin
        self.max_image_height = 200
        self.message_count = 0
        self.current_page = 1  # Aktuelle Seite
        self.total_pages = 1   # Mindestens eine Seite
        self.verbose = verbose
        self.model_name = model_name  # Whisper model name

        # Whisper model loading deaktiviert
        print(f"Audio transcription disabled - Whisper model '{self.model_name}' not loaded")
        self._whisper_model = None
        
        # Device information
        print("Audio transcription disabled - no device needed")
        
        # Register fonts
        font_path = Path(__file__).parent / 'fonts'
        try:
            # Register DejaVuSans font for normal text
            pdfmetrics.registerFont(TTFont('DejaVuSans', str(font_path / 'DejaVuSans.ttf')))
            # Register Symbola font for emojis
            pdfmetrics.registerFont(TTFont('Symbola', str(font_path / 'Symbola.ttf')))
            print("Fonts registered successfully")
        except Exception as e:
            print(f"Error loading fonts: {e}")
            print("Some characters might not display correctly.")
    
    def add_page_number(self, canvas):
        """Add page number to current page."""
        page_text = f"Seite {self.current_page}"
        canvas.saveState()
        canvas.setFont("Helvetica", 9)
        canvas.drawRightString(self.page_width - self.margin, self.margin - 20, page_text)
        canvas.restoreState()
        
    def new_page(self, canvas):
        """Start a new page and reset position."""
        canvas.showPage()
        self.current_page += 1
        self.total_pages = max(self.total_pages, self.current_page)
        self.y_position = self.page_height - self.margin
        canvas.setFont('DejaVuSans', 12)
        self.add_page_number(canvas)
        
    def is_emoji(self, char):
        """Check if a character is an emoji using comprehensive Unicode ranges"""
        if not char:
            return False
            
        emoji_pattern = (
            "\U0001F1E0-\U0001F1FF"  # flags (iOS)
            "\U0001F300-\U0001F5FF"  # symbols & pictographs
            "\U0001F600-\U0001F64F"  # emoticons
            "\U0001F680-\U0001F6FF"  # transport & map symbols
            "\U0001F700-\U0001F77F"  # alchemical symbols
            "\U0001F780-\U0001F7FF"  # Geometric Shapes Extended
            "\U0001F800-\U0001F8FF"  # Supplemental Arrows-C
            "\U0001F900-\U0001F9FF"  # Supplemental Symbols and Pictographs
            "\U0001FA00-\U0001FA6F"  # Chess Symbols
            "\U0001FA70-\U0001FAFF"  # Symbols and Pictographs Extended-A
            "\U00002702-\U000027B0"  # Dingbats
            "\U000024C2-\U0001F251"
        )
        
        import re
        pattern = f"[{emoji_pattern}]"
        return bool(re.match(pattern, char))
        
    def draw_text_with_emojis(self, canvas, text, x, y, base_font='DejaVuSans', size=10):
        current_x = x
        for char in text:
            try:
                if self.is_emoji(char):
                    canvas.setFont('Symbola', size)
                else:
                    canvas.setFont(base_font, size)
                width = canvas.stringWidth(char, canvas._fontname, size)
                canvas.drawString(current_x, y, char)
                current_x += width
            except Exception as e:
                print(f"Error drawing character '{char}': {e}")
                canvas.setFont(base_font, size)
                current_x += size/2
        return current_x - x  # Return total width
        
    def add_chat_line(self, canvas, message_data):
        if self.y_position < self.margin + self.line_height:
            self.new_page(canvas)
            
        sender_name = str(message_data.get('sender_name', ''))
        message_body = str(message_data.get('body', ''))
        timestamp = str(message_data.get('timestamp', ''))
        is_owner = message_data.get('is_owner', False)
        is_read = str(message_data.get('Status', ''))
        read_status = "✔️ Read" if "read" in is_read.lower() else ""
        attachment = str(message_data.get('attachment', ''))
        attachment_path = message_data.get('attachment_path', '')
        
        # Debug output for attachments
        if attachment and attachment != 'nan':
            if self.verbose:
                print("\n=== New Attachment ====")
                print(f"Found attachment: {attachment}")
                print(f"Attachment path: {attachment_path}")
                
                # Prüfe, ob es sich um eine URL handelt
                if attachment_path == "URL":
                    print(f"Type: URL")
                    # Für URLs keine weiteren Prüfungen durchführen
                else:
                    # Konvertiere zu absolutem Pfad und prüfe Existenz
                    try:
                        abs_attachment_path = str(Path(attachment_path).resolve())
                        if not os.path.exists(abs_attachment_path):
                            print(f"WARNING: File does not exist:")
                            print(f"  Relative path: {attachment_path}")
                            print(f"  Absolute path: {abs_attachment_path}")
                            attachment_path = None
                        else:
                            if self.is_image_file(attachment_path):
                                print(f"Type: Image")
                            elif self.is_audio_file(attachment_path):
                                print(f"Type: Audio")
                            elif self.is_video_file(attachment_path):
                                print(f"Type: Video")
                            else:
                                print(f"Type: Unknown")
                    except:
                        print(f"WARNING: Invalid attachment path: {attachment_path}")
                        attachment_path = None
 
            
        # Calculate positions
        left_col = self.margin
        middle_col = self.margin + 100
        right_col = self.page_width - self.margin - 100
        
        # Calculate maximum width for content in middle column
        # Reduziere die maximale Breite um einen Sicherheitsabstand von 20 Punkten
        max_content_width = (right_col - middle_col) - 20
        
        # Handle message content
        y_offset = 0
        
        # Bestimme die Hintergrundfarbe
        background_color = (0.97, 0.97, 0.97) if self.message_count % 2 == 0 else (0.95, 0.95, 1.0)
        
        # Draw alternating background
        canvas.setFillColorRGB(*background_color)
            
        # Calculate total height for background
        total_height = max(24, y_offset + 12)
        if message_body and message_body != 'nan':
            words = message_body.split()
            current_line = ""
            num_lines = 1
            
            for word in words:
                test_line = current_line + " " + word if current_line else word
                width = self.calculate_text_width(canvas, test_line)
                
                if width > 200:
                    num_lines += 1
                    current_line = word
                else:
                    current_line = test_line
            
            total_height = max(total_height, num_lines * 12 + 24)
            
        # Berechne zusätzliche Höhe für Anhänge
        if attachment and attachment != 'nan':
            # Spezialfall: URL
            if attachment_path == "URL":
                # Für URLs nur eine Zeile reservieren
                total_height += 20
            elif self.is_image_file(attachment_path):
                try:
                    img = Image.open(attachment_path)
                    img_width, img_height = img.size
                    
                    # Calculate scale factors
                    width_scale = max_content_width / img_width
                    height_scale = self.max_image_height / img_height
                    scale = min(width_scale, height_scale, 1.0)
                    
                    # Calculate final height
                    final_height = img_height * scale
                    total_height += final_height + 10
                except:
                    total_height += 10
            elif self.is_audio_file(attachment_path):
                # Prüfe ob eine Transkription existiert
                transcription, is_video = self.transcribe_audio(attachment_path)
                if transcription:
                    # Höhe für Header und Abstand
                    total_height += 20
                    # Berechne Höhe für Transkriptionstext
                    words = str(transcription).split()
                    current_line = ""
                    num_lines = 0
                    
                    for word in words:
                        test_line = current_line + " " + word if current_line else word
                        width = self.calculate_text_width(canvas, test_line)
                        if width > 200:
                            num_lines += 1
                            current_line = word
                        else:
                            current_line = test_line
                    if current_line:
                        num_lines += 1
                    # Transkriptionstext + Abstand nach unten
                    total_height += num_lines * 12 + 15
                else:
                    total_height += 10
            elif self.is_video_file(attachment_path):
                # Prüfe ob eine Transkription existiert
                transcription, is_video = self.transcribe_audio(attachment_path)
                if transcription:
                    # Höhe für Header und Abstand
                    total_height += 20
                    # Berechne Höhe für Transkriptionstext
                    words = str(transcription).split()
                    current_line = ""
                    num_lines = 0
                    
                    for word in words:
                        test_line = current_line + " " + word if current_line else word
                        width = self.calculate_text_width(canvas, test_line)
                        if width > 200:
                            num_lines += 1
                            current_line = word
                        else:
                            current_line = test_line
                    if current_line:
                        num_lines += 1
                    # Transkriptionstext + Abstand nach unten
                    total_height += num_lines * 12 + 15
                else:
                    total_height += 10
            else:
                total_height += 10
                
        # Wenn der Inhalt nicht mehr auf die Seite passt, neue Seite beginnen
        if self.y_position - total_height < self.margin:
            self.new_page(canvas)
            self.y_position = self.page_height - self.margin
            # Hintergrundfarbe nach Seitenumbruch neu setzen
            canvas.setFillColorRGB(*background_color)
        
        # Draw background including timestamp area and transcription
        canvas.rect(self.margin - 10, self.y_position - total_height, 
                   self.page_width - 2 * self.margin + 20, total_height + 15,
                   fill=1, stroke=0)
        canvas.setFillColorRGB(0, 0, 0)  # Reset to black for text
        
        if is_owner:
            # Berechne zuerst die Gesamthöhe für Owner-Nachrichten
            message_height = 0
            
            # Höhe für Nachrichtentext
            if message_body and message_body != 'nan':
                words = message_body.split()
                current_line = ""
                num_lines = 0
                
                for word in words:
                    test_line = current_line + " " + word if current_line else word
                    width = self.calculate_text_width(canvas, test_line)
                    if width > 200:
                        num_lines += 1
                        current_line = word
                    else:
                        current_line = test_line
                if current_line:
                    num_lines += 1
                message_height = num_lines * 12 + 15  # Text + Abstand

            # Höhe für Anhänge
            attachment_height = 0
            if attachment and attachment != 'nan':
                if self.is_image_file(attachment_path):
                    try:
                        img = Image.open(attachment_path)
                        img_width, img_height = img.size
                        width_scale = max_content_width / img_width
                        height_scale = self.max_image_height / img_height
                        scale = min(width_scale, height_scale, 1.0)
                        attachment_height = img_height * scale + 10
                    except:
                        attachment_height = 10
                elif self.is_audio_file(attachment_path) or self.is_video_file(attachment_path):
                    transcription, is_video = self.transcribe_audio(attachment_path)
                    if transcription:
                        # Höhe für Header
                        attachment_height += 20
                        # Höhe für Transkriptionstext
                        words = str(transcription).split()
                        current_line = ""
                        num_lines = 0
                        for word in words:
                            test_line = current_line + " " + word if current_line else word
                            width = self.calculate_text_width(canvas, test_line)
                            if width > 200:
                                num_lines += 1
                                current_line = word
                            else:
                                current_line = test_line
                        if current_line:
                            num_lines += 1
                        attachment_height += num_lines * 12 + 15

            # Gesamthöhe berechnen
            total_height = max(message_height, 24) + attachment_height  # Mindestens 24 für Name und Timestamp

            # Zeichne den Hintergrund
            canvas.setFillColorRGB(*background_color)  # Setze die Hintergrundfarbe
            canvas.rect(self.margin - 10, self.y_position - total_height,
                       self.page_width - 2 * self.margin + 20, total_height + 15,
                       fill=1, stroke=0)
            canvas.setFillColorRGB(0, 0, 0)  # Reset to black for text

            # Jetzt zeichne den Inhalt
            y_offset = 0
            
            # Left: sender name and timestamp
            canvas.setFont('DejaVuSans', 10)
            canvas.drawString(left_col, self.y_position, sender_name)
            canvas.setFont('DejaVuSans', 6)
            canvas.drawString(left_col, self.y_position - 10, timestamp)
            
            # Middle: message with emoji support
            if message_body and message_body != 'nan':
                words = message_body.split()
                current_line = ""
                
                for word in words:
                    test_line = current_line + " " + word if current_line else word
                    width = self.calculate_text_width(canvas, test_line)
                    
                    if width > 200:
                        self.draw_text_with_emojis(canvas, current_line, middle_col, self.y_position - y_offset)
                        current_line = word
                        y_offset += 12
                    else:
                        current_line = test_line
                
                if current_line:
                    self.draw_text_with_emojis(canvas, current_line, middle_col, self.y_position - y_offset)
                    y_offset += 12
            
            # Handle attachment
            if attachment and attachment != 'nan':
                if self.is_image_file(attachment_path):
                    # Add some spacing before image
                    y_offset += 5
                    image_height = self.embed_image(canvas, attachment_path, middle_col, 
                                                  self.y_position - y_offset, 
                                                  self.max_image_height,
                                                  max_content_width)  # Pass max width
                    if image_height > 0:
                        y_offset += image_height + 5
                elif attachment_path == "URL":
                    # Zeige URL mit Link-Symbol an
                    y_offset += 5
                    canvas.setFont('DejaVuSans', 8)
                    canvas.drawString(middle_col, self.y_position - y_offset, f"Link: {attachment}")
                    y_offset += 10
                elif self.is_audio_file(attachment_path) or self.is_video_file(attachment_path):
                    # Transcribe audio/video and display result
                    if self.verbose:
                        print(f"Attempting to transcribe: {attachment_path}")
                    transcription, is_video = self.transcribe_audio(attachment_path)
                    if transcription:
                        y_offset += 5  # Abstand vor der Transkription
                        canvas.setFont('DejaVuSans', 8)
                        header_text = "Video Attachment, Transcription:" if is_video else "Audio Attachment, Transcription:"
                        canvas.drawString(middle_col, self.y_position - y_offset, header_text)
                        y_offset += 12
                        # Display transcription text
                        canvas.setFont('DejaVuSans', 10)
                        words = str(transcription).split()
                        current_line = ""
                        for word in words:
                            test_line = current_line + " " + word if current_line else word
                            width = self.calculate_text_width(canvas, test_line)
                            if width > 200:
                                self.draw_text_with_emojis(canvas, current_line, middle_col + 10, self.y_position - y_offset)
                                current_line = word
                                y_offset += 12
                            else:
                                current_line = test_line
                        if current_line:
                            self.draw_text_with_emojis(canvas, current_line, middle_col + 10, self.y_position - y_offset)
                            y_offset += 12
                    else:
                        canvas.setFont('DejaVuSans', 8)
                        canvas.drawString(middle_col, self.y_position - y_offset, f"[Audio: {attachment}] (Transcription failed)")
                        y_offset += 10
                elif attachment_path == "URL":
                    # Zeige URL mit Link-Symbol an
                    y_offset += 5
                    canvas.setFont('DejaVuSans', 8)
                    canvas.drawString(middle_col, self.y_position - y_offset, f"Link: {attachment}")
                    y_offset += 10
                else:
                    # Display attachment name in smaller font
                    canvas.setFont('DejaVuSans', 8)
                    canvas.drawString(middle_col, self.y_position - y_offset, f"{attachment}")
                    y_offset += 10
            
            # Right: read status
            canvas.setFont('DejaVuSans', 8)
            canvas.drawString(right_col, self.y_position, read_status)
        else:
            # Similar structure for non-owner messages
            # Left: read status
            canvas.setFont('DejaVuSans', 8)
            canvas.drawString(left_col, self.y_position, read_status)
            
            # Middle: message with emoji support
            if message_body and message_body != 'nan':
                words = message_body.split()
                current_line = ""
                
                for word in words:
                    test_line = current_line + " " + word if current_line else word
                    width = self.calculate_text_width(canvas, test_line)
                    
                    if width > 200:
                        self.draw_text_with_emojis(canvas, current_line, middle_col, self.y_position - y_offset)
                        current_line = word
                        y_offset += 12
                    else:
                        current_line = test_line
                
                if current_line:
                    self.draw_text_with_emojis(canvas, current_line, middle_col, self.y_position - y_offset)
                    y_offset += 12
            
            # Handle attachment
            if attachment and attachment != 'nan':
                if self.is_image_file(attachment_path):
                    # Add some spacing before image
                    y_offset += 5
                    image_height = self.embed_image(canvas, attachment_path, middle_col, 
                                                  self.y_position - y_offset, 
                                                  self.max_image_height,
                                                  max_content_width)  # Pass max width
                    if image_height > 0:
                        y_offset += image_height + 5
                elif attachment_path == "URL":
                    # Zeige URL mit Link-Symbol an
                    y_offset += 5
                    canvas.setFont('DejaVuSans', 8)
                    canvas.drawString(middle_col, self.y_position - y_offset, f"Link: {attachment}")
                    y_offset += 10
                elif self.is_audio_file(attachment_path) or self.is_video_file(attachment_path):
                    # Transcribe audio/video and display result
                    if self.verbose:
                        print(f"Attempting to transcribe: {attachment_path}")
                    transcription, is_video = self.transcribe_audio(attachment_path)
                    if transcription:
                        y_offset += 5  # Abstand vor der Transkription
                        canvas.setFont('DejaVuSans', 8)
                        header_text = "Video Attachment, Transcription:" if is_video else "Audio Attachment, Transcription:"
                        canvas.drawString(middle_col, self.y_position - y_offset, header_text)
                        y_offset += 12
                        # Display transcription text
                        canvas.setFont('DejaVuSans', 10)
                        words = str(transcription).split()
                        current_line = ""
                        for word in words:
                            test_line = current_line + " " + word if current_line else word
                            width = self.calculate_text_width(canvas, test_line)
                            if width > 200:
                                self.draw_text_with_emojis(canvas, current_line, middle_col + 10, self.y_position - y_offset)
                                current_line = word
                                y_offset += 12
                            else:
                                current_line = test_line
                        if current_line:
                            self.draw_text_with_emojis(canvas, current_line, middle_col + 10, self.y_position - y_offset)
                            y_offset += 12
                    else:
                        canvas.setFont('DejaVuSans', 8)
                        canvas.drawString(middle_col, self.y_position - y_offset, f"🎵 [Audio: {attachment}] (Transcription failed)")
                        y_offset += 10
                else:
                    # Display attachment name in smaller font
                    canvas.setFont('DejaVuSans', 8)
                    canvas.drawString(middle_col, self.y_position - y_offset, f"{attachment}")
                    y_offset += 10
            
            # Right: sender name and timestamp
            canvas.setFont('DejaVuSans', 10)
            canvas.drawString(right_col, self.y_position, sender_name)
            canvas.setFont('DejaVuSans', 6)
            canvas.drawString(right_col, self.y_position - 10, timestamp)
        
        # Update y_position with the maximum offset used
        self.y_position -= max(24, y_offset + 12)
        # Add some spacing between messages
        self.y_position -= 5
        self.message_count += 1  # Increment for alternating backgrounds

#pragma region is_image_file

    def is_image_file(self, filename):
        """Check if the filename has an image extension."""
        if not filename:
            return False
        return Path(filename).suffix.lower() in {'.jpg', '.jpeg', '.png', '.gif', '.bmp'}
#pragma endregion

    def is_audio_file(self, filename):
        """Check if the filename has an audio extension."""
        if not filename:
            return False
        return Path(filename).suffix.lower() in {'.mp3', '.wav', '.m4a', '.ogg', '.aac', '.m4a'}
        
    def is_video_file(self, filename):
        """Check if the filename has a video extension."""
        if not filename:
            return False
        return Path(filename).suffix.lower() in {'.mp4', '.avi', '.mov', '.mkv'}
        
    def extract_audio_from_video(self, video_path):
        """Extract audio from video file and return path to temporary audio file.
        DEAKTIVIERT: Gibt immer None, False zurück, um PyAudio-Abhängigkeit zu vermeiden.
        """
        if self.verbose:
            print(f"Audio extraction disabled for: {video_path}")
        return None, False

    def get_transcription_path(self, audio_path):
        """Get the path for the transcription file."""
        # Convert audio_path to absolute path
        abs_audio_path = Path(audio_path).resolve()
        #print(f"Audio file absolute path: {abs_audio_path}")
        
        # Get the directory containing the 'files' directory
        base_dir = abs_audio_path.parent.parent.parent  # Go up from audio file to 'files' directory, then up again to root
        #print(f"Base directory: {base_dir}")
        
        # Create transcriptions directory parallel to 'files'
        trans_dir = base_dir / 'transcriptions'
        #print(f"Creating transcriptions directory at: {trans_dir}")
        trans_dir.mkdir(exist_ok=True)
        
        # Create a filename based on the original audio filename
        audio_filename = abs_audio_path.name
        trans_filename = f"{Path(audio_filename).stem}.txt"
        final_path = trans_dir / trans_filename
        #print(f"Transcription will be saved to: {final_path}")
        
        return final_path

    def load_cached_transcription(self, trans_path):
        """Load transcription from cache if it exists."""
        try:
            if trans_path.exists():
                with open(trans_path, 'r', encoding='utf-8') as f:
                    return f.read().strip()
        except Exception as e:
            print(f"Error reading cached transcription: {e}")
        return None

    def save_transcription(self, trans_path, text):
        """Save transcription to cache."""
        try:
            with open(trans_path, 'w', encoding='utf-8') as f:
                f.write(text)
        except Exception as e:
            print(f"Error saving transcription: {e}")

    def transcribe_audio(self, file_path):
        """Transcribe audio or video file using Whisper, with caching.
        DEAKTIVIERT: Gibt immer None, False zurück, um PyAudio-Abhängigkeit zu vermeiden.
        """
        if self.verbose:
            print(f"Audio transcription disabled for: {file_path}")
        return None, False
    
    def embed_image(self, canvas, image_path, x, y, max_height, max_width):
        """Embed an image in the PDF with maximum height and width constraints."""
        if not image_path or not os.path.exists(image_path):
            print(f"Image not found: {image_path}")
            return 0
            
        try:
            img = Image.open(image_path)
            img_width, img_height = img.size
            
            # Calculate scale factors for both constraints
            width_scale = max_width / img_width
            height_scale = max_height / img_height
            
            # Use the smaller scale to ensure both constraints are met
            scale = min(width_scale, height_scale, 1.0)
            
            # Calculate final dimensions
            final_width = img_width * scale
            final_height = img_height * scale
            
            # Check if we need a new page for the image
            if y - final_height < self.margin:
                self.new_page(canvas)
                # Adjust y position to top of new page
                y = self.page_height - self.margin
            
            # Draw the image
            canvas.drawImage(image_path, x, y - final_height, 
                           width=final_width, height=final_height)
            return final_height
        except Exception as e:
            print(f"Error embedding image {image_path}: {e}")
            return 0

    def calculate_text_width(self, canvas, text):
        """Calculate the width of text considering emojis."""
        width = 0
        canvas.setFont('DejaVuSans', 10)
        for char in text:
            width += canvas.stringWidth(char, canvas._fontname, 10)
        return width

    def add_participants_header(self, canvas, participants_data):
        # Set initial position at the top of the page
        self.y_position = self.page_height - self.margin
        
        # Excel Dateiname ohne Endung
        canvas.setFont('DejaVuSans', 16)
        if isinstance(participants_data, dict):
            excel_name = Path(participants_data['excel_path']).stem
            participants_list = participants_data['participants']
        else:
            # Legacy Format: participants_data ist direkt die Liste
            excel_name = Path(participants_data[0].get('excel_path', '')).stem
            participants_list = participants_data

        # Zentriere den Dateinamen
        text_width = canvas.stringWidth(excel_name, 'DejaVuSans', 16)
        x_position = (self.page_width - text_width) / 2
        canvas.drawString(x_position, self.y_position, excel_name)
        self.y_position -= self.line_height * 2
        
        # Add title
        canvas.setFont('DejaVuSans', 11)
        canvas.drawString(self.margin, self.y_position, "Chat Participants:")
        self.y_position -= self.line_height * 2
        
        # Add participants
        canvas.setFont('DejaVuSans', 10)
        for participant in participants_list:
            name = participant.get('sender_name', '')
            is_owner = participant.get('is_owner', False)
            participant_text = f"{name} {'(OWNER)' if is_owner else ''}"
            canvas.drawString(self.margin + 20, self.y_position, participant_text)
            self.y_position -= self.line_height
        
        # Add separator line
        self.y_position -= self.line_height
        canvas.line(self.margin, self.y_position, self.page_width - self.margin, self.y_position)
        self.y_position -= self.line_height * 2

def generate_chat_report(excel_file, output_file='chat_report.pdf', verbose=False, model_name="medium"):
    """Generate a PDF report from the Excel file."""
    global attachment_found_counter, attachment_not_found_counter
    
    # Reset counters
    attachment_found_counter = 0
    attachment_not_found_counter = 0
    
    try:
        # Lese die Excel-Datei
        df, metadata = read_excel_file(excel_file)
        
        if df is None:
            print("Fehler: Die Excel-Datei konnte nicht gelesen werden.")
            return
            
        if verbose:
            print(f"Generiere PDF-Report: {output_file}")
            print(f"Excel-Datei: {excel_file}")
            print(f"Whisper-Modell: {model_name}")
            print(f"DataFrame Spalten: {len(df.columns)}")
            print(f"DataFrame Zeilen: {len(df)}")
            
        # Initialisiere den PDF-Report
        c = canvas.Canvas(output_file, pagesize=A4)
        report = ChatReport(verbose=verbose, model_name=model_name)
        
        # Spaltenindizes bestimmen (mit Fallback-Werten)
        from_col_idx = 1  # Standard: Spalte 1 für 'From'
        body_col_idx = 8  # Standard: Spalte 8 für 'Body'
        status_col_idx = 9  # Standard: Spalte 9 für 'Status'
        date_col_idx = 17  # Standard: Spalte 17 für 'Timestamp-Date'
        time_col_idx = 18  # Standard: Spalte 18 für 'Timestamp-Time'
        direction_col_idx = 6  # Standard: Spalte 6 für 'Direction'
        attachment_col_idx = 25  # Standard: Spalte 25 für 'Attachment #1'
    
    # Add participants header
    participants_data = {
        'excel_path': str(chat_data.excel_path),
        'participants': [
            {'sender_name': p.name, 'is_owner': p.is_owner}
            for p in chat_data.participants
        ]
    }
    report.add_participants_header(c, participants_data)
    
    # Process each message
    for msg in chat_data.messages:
        message_data = {
            'sender_name': msg.sender.name,
            'body': msg.body or '[Empty message]',
            'timestamp': msg.timestamp.strftime('%d.%m.%Y %H:%M:%S'),
            'is_owner': msg.sender.is_owner,
            'Status': msg.status
        }
        
        # Handle attachment if present
        if msg.attachment:
            message_data['attachment'] = msg.attachment.filename
            if msg.attachment.full_path:
                message_data['attachment_path'] = str(msg.attachment.full_path)
                if msg.attachment.type == 'audio' and msg.attachment.transcription:
                    message_data['audio_transcription'] = msg.attachment.transcription
        
        report.add_chat_line(c, message_data)
    
    # Save the PDF
    c.save()
    
    # Print attachment statistics
    stats = chat_data.attachment_stats
    print(f"Attachments found: {stats['found']}")
    if stats['not_found'] > 0:
        print(f"Attachments not found: {stats['not_found']}")

def generate_chat_report(excel_file, output_file='chat_report.pdf', verbose=False, model_name="medium"):
    """Generate a PDF report from the Excel file."""
    global attachment_found_counter, attachment_not_found_counter
    
    # Reset counters
    attachment_found_counter = 0
    attachment_not_found_counter = 0
    
    try:
        # Lese die Excel-Datei
        df, metadata = read_excel_file(excel_file)
        
        if df is None:
            print("Fehler: Die Excel-Datei konnte nicht gelesen werden.")
            return
            
        if verbose:
            print(f"Generiere PDF-Report: {output_file}")
            print(f"Excel-Datei: {excel_file}")
            print(f"Whisper-Modell: {model_name}")
            print(f"DataFrame Spalten: {len(df.columns)}")
            print(f"DataFrame Zeilen: {len(df)}")
            
        # Initialisiere den PDF-Report
        c = canvas.Canvas(output_file, pagesize=A4)
        report = ChatReport(verbose=verbose, model_name=model_name)
        
        # Spaltenindizes bestimmen (mit Fallback-Werten)
        from_col_idx = 1  # Standard: Spalte 1 für 'From'
        body_col_idx = 8  # Standard: Spalte 8 für 'Body'
        status_col_idx = 9  # Standard: Spalte 9 für 'Status'
        date_col_idx = 17  # Standard: Spalte 17 für 'Timestamp-Date'
        time_col_idx = 18  # Standard: Spalte 18 für 'Timestamp-Time'
        direction_col_idx = 6  # Standard: Spalte 6 für 'Direction'
        attachment_col_idx = 25  # Standard: Spalte 25 für 'Attachment #1'
    
        # Sichere Funktion zum Abrufen von Zellwerten innerhalb der generate_chat_report-Funktion
        def safe_get_cell_local(row, idx, default=''):
            try:
                if idx < len(row):
                    value = str(row.iloc[idx]).strip()
                    return value if value != 'nan' else default
                return default
            except Exception as e:
                if verbose:
                    print(f"Fehler beim Lesen der Zelle {idx}: {e}")
                return default
                
        # Überprüfe die Spaltenanzahl und passe die Indizes an
        if len(df.columns) <= attachment_col_idx:
            # Wenn weniger Spalten vorhanden sind, suche nach Spalten mit bestimmten Namen
            if verbose:
                print("Spaltenanzahl kleiner als erwartet, suche nach Spalten anhand der Namen...")
                
            for i, col_name in enumerate(df.columns):
                col_name_lower = str(col_name).lower()
                if 'from' in col_name_lower:
                    from_col_idx = i
                    if verbose:
                        print(f"'From'-Spalte gefunden: {i}")
                elif 'body' in col_name_lower:
                    body_col_idx = i
                    if verbose:
                        print(f"'Body'-Spalte gefunden: {i}")
                elif 'status' in col_name_lower:
                    status_col_idx = i
                    if verbose:
                        print(f"'Status'-Spalte gefunden: {i}")
                elif 'date' in col_name_lower:
                    date_col_idx = i
                    if verbose:
                        print(f"'Date'-Spalte gefunden: {i}")
                elif 'time' in col_name_lower:
                    time_col_idx = i
                    if verbose:
                        print(f"'Time'-Spalte gefunden: {i}")
                elif 'direction' in col_name_lower:
                    direction_col_idx = i
                    if verbose:
                        print(f"'Direction'-Spalte gefunden: {i}")
                elif 'attachment' in col_name_lower:
                    attachment_col_idx = i
                    if verbose:
                        print(f"'Attachment'-Spalte gefunden: {i}")
                
    if verbose:
        print(f"Verwendete Spaltenindizes:")
        print(f"  From: {from_col_idx}")
        print(f"  Body: {body_col_idx}")
        print(f"  Status: {status_col_idx}")
        print(f"  Date: {date_col_idx}")
        print(f"  Time: {time_col_idx}")
        print(f"  Direction: {direction_col_idx}")
        print(f"  Attachment: {attachment_col_idx}")
    
    # Parse participants
    participants_dict = {}
    image_attachments = []
    
    # Iterate through rows to find participants
    for _, row in df.iloc[1:].iterrows():
        try:
            from_field = safe_get_cell(row, from_col_idx)
            chat_id, name = parse_participant(from_field)
            
            if chat_id and chat_id not in participants_dict:
                participants_dict[chat_id] = {
                    'name': name,
                    'is_owner': False
                }
            
            # Check for attachments
            attachment = safe_get_cell(row, attachment_col_idx)
            if attachment != 'nan' and attachment:
                attachment_path = find_attachment_file(excel_file, attachment)
                if attachment_path and report.is_image_file(attachment_path):
                    image_attachments.append(attachment_path)
            
            # Check for owner
            direction = safe_get_cell(row, direction_col_idx).lower()
            
            if direction == 'outgoing':
                if chat_id and chat_id in participants_dict:
                    participants_dict[chat_id]['is_owner'] = True
        except Exception as e:
            if verbose:
                print(f"Fehler bei der Verarbeitung einer Zeile: {e}")
    
    if image_attachments and verbose:
        print(f"\nFound {len(image_attachments)} image attachments in chat:")
        for img in image_attachments:
            print(f"- {img}")
    
    # Convert participants_dict to list for header
    participants = []
    for info in participants_dict.values():
        participants.append({
            'sender_name': info['name'],
            'is_owner': info['is_owner']
        })
    
    # Add participants header
    report.add_participants_header(c, participants)
    
    # Process the messages
    messages = []
    for _, row in df.iloc[1:].iterrows():
        from_field = str(row.iloc[1]).strip()
        chat_id, name = parse_participant(from_field)
        
        if chat_id and chat_id in participants_dict:
            # Get body content first
            body_content = str(row.iloc[8]).strip()  # Body is in column 8
            status = str(row.iloc[9]).strip()  # Status is in column 9
            attachment = str(row.iloc[25]).strip()  # Attachment #1 is in column 25
            
            # Check for attachment if body is empty
            if body_content == 'nan' or not body_content:
                if attachment != 'nan' and attachment:
                    attachment_path = find_attachment_file(excel_file, attachment)
                    if attachment_path:
                        body_content = ""  # Don't set body content, we'll display attachment separately
                    else:
                        body_content = f"[Missing Attachment: {attachment}]"
                else:
                    body_content = "[Empty message]"
            
            # Get timestamp from Timestamp-Time column
            timestamp = str(row.iloc[18]).strip()
            # Remove (UTC+0) if present
            timestamp = timestamp.replace('(UTC+0)', '').strip()
            # Check if timestamp contains a date
            if '.' not in timestamp:
                # If no date, get it from Timestamp-Date column
                date = str(row.iloc[17]).strip()
                if date != 'nan' and date:
                    timestamp = f"{date} {timestamp}"
            
            # Check direction for owner detection
            direction = str(row.iloc[6]).strip().lower()
            is_owner = direction == 'outgoing'
            
            message_data = {
                'sender_name': name,
                'body': body_content,
                'timestamp': timestamp,
                'is_owner': is_owner,
                'Status': status,
                'attachment': attachment
            }
            
            # Add full path to attachment if it exists
            if attachment and attachment != 'nan':
                attachment_path = find_attachment_file(excel_file, attachment)
                if attachment_path:
                    message_data['attachment_path'] = attachment_path
                    # Wenn es eine Audio-Datei ist, füge Transkription hinzu
                    if report.is_audio_file(attachment_path):
                        transcription = report.transcribe_audio(attachment_path)
                        message_data['audio_transcription'] = transcription
            
            messages.append(message_data)
    
    # Initialisiere die erste Seite mit Seitennummer
    report.add_page_number(c)
    
    # Sammle Teilnehmer und füge Excel-Pfad hinzu
    participants = {
        'excel_path': excel_file,
        'participants': []
    }
    seen = set()
    for message in messages:
        sender_name = message.get('sender_name')
        if sender_name and sender_name not in seen:
            seen.add(sender_name)
            participants['participants'].append({
                'sender_name': sender_name,
                'is_owner': message.get('is_owner', False)
            })
    
    # Füge Teilnehmerliste hinzu
    report.add_participants_header(c, participants)
    
    # Process each message
    for message in messages:
        report.add_chat_line(c, message)
    
    # Save the PDF mit dem angegebenen Ausgabepfad
    c.save()
    
    # Print attachment statistics
    print(f"Attachments found: {attachment_found_counter}")
    if attachment_not_found_counter > 0:
        print(f"Attachments not found: {attachment_not_found_counter}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Analysiere WhatsApp-Export Excel-Datei und generiere optional einen PDF-Report.')
    parser.add_argument('excel_file', type=str, help='Pfad zur Excel-Datei')
    parser.add_argument('--export', '-e', action='store_true',
                       help='PDF-Report generieren (standardmäßig wird nur die Statistik angezeigt)')
    parser.add_argument('--output', '-o', type=str, default='chat_report.pdf',
                       help='Pfad zur Ausgabe-PDF-Datei (Standard: chat_report.pdf)')
    parser.add_argument('-v', '--verbose', action='store_true',
                       help='Ausführliche Ausgabe aktivieren')
    parser.add_argument('--model', '-m', type=str, default='medium',
                       choices=['tiny', 'base', 'small', 'medium', 'large'],
                       help='Whisper-Modell für die Transkription (Standard: medium)')
    
    args = parser.parse_args()
    
    # Überprüfe, ob die Excel-Datei existiert
    if not os.path.exists(args.excel_file):
        print(f"Fehler: Die Datei '{args.excel_file}' existiert nicht.")
        sys.exit(1)
    
    # Lese die Excel-Datei mit der neuen Funktion
    df, metadata = read_excel_file(args.excel_file)
    
    if df is None:
        print("Fehler: Die Excel-Datei konnte nicht gelesen werden.")
        sys.exit(1)
    
    # Zeige Statistiken an (immer)
    stats = generate_statistics(df, metadata, args.excel_file, args.verbose)
    print(stats)
    
    # Wenn PDF-Report generiert werden soll
    if args.export:
        print("\nGeneriere PDF-Report...")
        generate_chat_report(args.excel_file, args.output, args.verbose, args.model)
        print(f"PDF-Report wurde generiert: {args.output}")
