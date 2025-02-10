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
import whisper
from pydub import AudioSegment
from moviepy.editor import VideoFileClip
import tempfile

# Global counters
attachment_not_found_counter = 0
attachment_found_counter = 0

def find_attachment_file(excel_path, attachment_name):
    """
    Search for an attachment file in the 'files' directory parallel to the Excel file.
    Returns the full path if found, None otherwise.
    """
    global attachment_not_found_counter
    global attachment_found_counter
    
    #print(f"Searching for attachment: {attachment_name}")
    if not attachment_name or attachment_name == 'nan':
        return None
        
    # Get the directory containing the Excel file
    excel_dir = Path(excel_path).parent
    # Construct path to the files directory
    files_dir = excel_dir / 'files'
    
    if not files_dir.exists():
        attachment_not_found_counter += 1
        #print(f"Warning: files directory not found at {files_dir}")
        return None
    
    # Walk through all subdirectories
    for root, _, files in os.walk(files_dir):
        for file in files:
            if file == attachment_name:
                attachment_found_counter += 1
                return os.path.join(root, file)
    
    attachment_not_found_counter += 1
    #print(f"Warning: Attachment not found: {attachment_name}")
    return None

class ChatReport:
    def __init__(self, verbose=False):
        self.page_width, self.page_height = A4
        self.margin = 50
        self.line_height = 14
        self.y_position = self.page_height - self.margin
        self.max_image_height = 200
        self.message_count = 0
        self.current_page = 1  # Aktuelle Seite
        self.total_pages = 1   # Mindestens eine Seite
        self.verbose = verbose
        
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
        page_text = f"Seite {self.current_page}/{self.total_pages}"
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
                
                # Konvertiere zu absolutem Pfad und prüfe Existenz
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
            if self.is_image_file(attachment_path):
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
                    # Höhe für Header
                    total_height += 12
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
                    total_height += num_lines * 12 + 10
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
        
        # Draw background including timestamp area
        canvas.rect(self.margin - 10, self.y_position - total_height, 
                   self.page_width - 2 * self.margin + 20, total_height + 15,
                   fill=1, stroke=0)
        canvas.setFillColorRGB(0, 0, 0)  # Reset to black for text
        
        if is_owner:
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
                elif self.is_audio_file(attachment_path) or self.is_video_file(attachment_path):
                    # Transcribe audio/video and display result
                    if self.verbose:
                        print(f"Attempting to transcribe: {attachment_path}")
                    transcription, is_video = self.transcribe_audio(attachment_path)
                    if transcription:
                        canvas.setFont('DejaVuSans', 8)
                        # Verschiebe den Header-Text nach oben, indem wir vom aktuellen y_offset 3 Punkte abziehen
                        header_text = "Video Attachment, Transcription:" if is_video else "Audio Attachment, Transcription:"
                        canvas.drawString(middle_col, self.y_position - (y_offset - 3), header_text)
                        y_offset += 15
                        # Display transcription text
                        canvas.setFont('DejaVuSans', 10)
                        words = transcription.split()
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
                elif self.is_audio_file(attachment_path) or self.is_video_file(attachment_path):
                    # Transcribe audio/video and display result
                    if self.verbose:
                        print(f"Attempting to transcribe: {attachment_path}")
                    transcription, is_video = self.transcribe_audio(attachment_path)
                    if transcription:
                        canvas.setFont('DejaVuSans', 8)
                        # Verschiebe den Header-Text nach oben, indem wir vom aktuellen y_offset 3 Punkte abziehen
                        header_text = "Video Attachment, Transcription:" if is_video else "Audio Attachment, Transcription:"
                        canvas.drawString(middle_col, self.y_position - (y_offset - 3), header_text)
                        y_offset += 15
                        # Display transcription text
                        canvas.setFont('DejaVuSans', 10)
                        words = transcription.split()
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

    def is_image_file(self, filename):
        """Check if the filename has an image extension."""
        if not filename:
            return False
        return Path(filename).suffix.lower() in {'.jpg', '.jpeg', '.png', '.gif', '.bmp'}

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
        """Extract audio from video file and return path to temporary audio file."""
        if not video_path:
            return None, False
            
        # Konvertiere zu absolutem Pfad
        abs_video_path = str(Path(video_path).resolve())
        
        # Prüfe ob die Datei existiert
        if not os.path.exists(abs_video_path):
            if self.verbose:
                print(f"Video file does not exist: {abs_video_path}")
            return None, False
            
        if self.verbose:
            print(f"Processing media file: {abs_video_path}")
            print(f"File exists: {os.path.exists(abs_video_path)}")
            
        # Versuche die Datei als Audio-only MP4 zu öffnen
        try:
            audio = AudioSegment.from_file(abs_video_path)
            if self.verbose:
                print("Successfully loaded as audio file")
            # Erstelle temporäre WAV-Datei
            with tempfile.NamedTemporaryFile(suffix='.wav', delete=False) as temp_audio:
                audio.export(temp_audio.name, format='wav')
                return temp_audio.name, True
        except Exception as audio_error:
            if self.verbose:
                print(f"Not an audio-only file, trying as video: {audio_error}")
                
        # Wenn es keine Audio-Datei ist, versuche es als Video
        try:
            # Versuche das Video zu laden
            try:
                video = VideoFileClip(abs_video_path)
                if self.verbose:
                    print(f"Video Size: {video.size}")
                    print(f"Video Duration: {video.duration}")
            except Exception as e:
                if 'video_fps' in str(e):
                    if self.verbose:
                        print("Retrying with fps=30...")
                    # Wenn FPS nicht erkannt werden können, setze sie auf 30
                    video = VideoFileClip(abs_video_path, fps_source='fps')
                else:
                    raise e
            
            # Prüfe auf Audiospur
            audio = video.audio
            if audio is None:
                if self.verbose:
                    print(f"No audio track found in video: {Path(video_path).name}")
                video.close()
                return None, False
                
            # Erstelle temporäre WAV-Datei
            with tempfile.NamedTemporaryFile(suffix='.wav', delete=False) as temp_audio:
                if self.verbose:
                    print(f"Converting video to audio: {temp_audio.name}")
                audio.write_audiofile(temp_audio.name, verbose=self.verbose, logger=None)
                video.close()
                if self.verbose:
                    print("Audio extraction completed successfully")
                return temp_audio.name, True
                
        except Exception as e:
            if self.verbose:
                print(f"Error extracting audio from video: {e}")
            try:
                video.close()
            except:
                pass
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
        """Transcribe audio or video file using Whisper, with caching."""
        if not file_path or not os.path.exists(file_path):
            print(f"File not found: {file_path}")
            return None, False
            
        # Bestimme den Pfad für die Transkription
        trans_path = self.get_transcription_path(file_path)
        
        # Prüfe auf Cache
        cached_text = self.load_cached_transcription(trans_path)
        if cached_text is not None:
            if self.verbose:
                print(f"Using cached transcription for {Path(file_path).name}")
            return cached_text, False
            
        try:
            # Behandle Videos speziell
            if self.is_video_file(file_path):
                if self.verbose:
                    print(f"Processing video file: {Path(file_path).name}")
                temp_audio_path = self.extract_audio_from_video(file_path)
                if temp_audio_path is None:
                    return None, False, False
                audio_path = temp_audio_path
            else:
                audio_path = file_path
                
            if self.verbose:
                print(f"Transcribing {Path(file_path).name}...")
                
            # Load the Whisper model (using base model for speed)
            model = whisper.load_model("base")

            # Convert non-WAV files to WAV using pydub if needed
            if not self.is_video_file(file_path) and Path(audio_path).suffix.lower() != '.wav':
                audio = AudioSegment.from_file(audio_path)
                with tempfile.NamedTemporaryFile(suffix='.wav', delete=True) as temp_wav:
                    audio.export(temp_wav.name, format='wav')
                    result = model.transcribe(temp_wav.name)
            else:
                result = model.transcribe(audio_path)

            transcribed_text = result['text'].strip()
            
            # Cache the transcription
            self.save_transcription(trans_path, transcribed_text)
            
            # Cleanup temporary audio file if it was a video
            if self.is_video_file(file_path) and audio_path != file_path:
                try:
                    os.unlink(audio_path)
                except Exception as e:
                    if self.verbose:
                        print(f"Warning: Could not delete temporary audio file: {e}")
            
            return transcribed_text
            
        except Exception as e:
            print(f"Error transcribing {Path(file_path).name}: {e}")
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
        
        # Add title
        canvas.setFont('DejaVuSans', 14)
        canvas.drawString(self.margin, self.y_position, "Chat Participants:")
        self.y_position -= self.line_height * 2
        
        # Add participants
        canvas.setFont('DejaVuSans', 10)
        for participant in participants_data:
            name = participant.get('sender_name', '')
            is_owner = participant.get('is_owner', False)
            participant_text = f"{name} {'(OWNER)' if is_owner else ''}"
            canvas.drawString(self.margin + 20, self.y_position, participant_text)
            self.y_position -= self.line_height
        
        # Add separator line
        self.y_position -= self.line_height
        canvas.line(self.margin, self.y_position, self.page_width - self.margin, self.y_position)
        self.y_position -= self.line_height * 2

def generate_chat_report(excel_file, output_file, verbose=False):
    # Get the directory of the input Excel file and create output path
    input_dir = os.path.dirname(os.path.abspath(excel_file))
    base_name = os.path.splitext(os.path.basename(excel_file))[0]
    output_file = os.path.join(input_dir, f"{base_name}_chat_report.pdf")
    
    # Read the Excel file
    df = pd.read_excel(excel_file, na_values=[''], keep_default_na=False)
    
    # Create PDF
    c = canvas.Canvas(output_file, pagesize=A4)
    report = ChatReport(verbose=verbose)
    
    # Process the data to find owner and participants
    participants_dict = {}

    def parse_participant(participant_str):
        """Parse participant string into chat ID and name."""
        parts = participant_str.strip().split()
        if len(parts) >= 2:
            chat_id = parts[0]
            name = ' '.join(parts[1:])
            # Check if it's a valid chat ID (numeric)
            if chat_id.isdigit():
                return chat_id, name
        return None, participant_str

    # First pass: identify participants
    for _, row in df.iloc[1:].iterrows():
        # Check the From column (index 1) for participants
        sender = str(row.iloc[1]).strip()
        if sender and sender != 'nan' and sender != 'From' and sender != 'System Message System Message':
            chat_id, name = parse_participant(sender)
            if chat_id and name:
                if chat_id not in participants_dict:
                    participants_dict[chat_id] = {
                        'chat_id': chat_id,
                        'name': name,
                        'is_owner': False,
                        'message_direction': str(row.iloc[6]).strip()  # Direction column
                    }

    # Second pass: identify the owner and collect attachments
    image_attachments = []  # Keep track of all image attachments
    for _, row in df.iloc[1:].iterrows():
        sender = str(row.iloc[1]).strip()
        direction = str(row.iloc[6]).strip().lower()
        attachment = str(row.iloc[25]).strip()  # Attachment #1 is in column 25
        
        if attachment and attachment != 'nan':
            attachment_path = find_attachment_file(excel_file, attachment)
            if attachment_path:
                if Path(attachment_path).suffix.lower() in {'.jpg', '.jpeg', '.png', '.gif', '.bmp'}:
                    image_attachments.append(attachment_path)
        
        if sender and sender != 'System Message System Message':
            chat_id, _ = parse_participant(sender)
            if chat_id and chat_id in participants_dict:
                # Check if this is an outgoing message
                if direction == 'outgoing':
                    participants_dict[chat_id]['is_owner'] = True
                    break

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
    
    # Process each message
    for message in messages:
        report.add_chat_line(c, message)
    
    # Save the PDF
    c.save()

    print(f"Attachments not found: {attachment_not_found_counter}")
    print(f"Attachments found: {attachment_found_counter}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Generate PDF report from chat export Excel file.')
    parser.add_argument('excel_file', type=str, help='Path to the input Excel file')
    parser.add_argument('--output', '-o', type=str, default='chat_report.pdf',
                       help='Path to the output PDF file (default: chat_report.pdf)')
    parser.add_argument('-v', '--verbose', action='store_true',
                       help='Enable verbose output')
    
    args = parser.parse_args()
    
    # Read the Excel file
    df = pd.read_excel(args.excel_file, na_values=[''], keep_default_na=False)
    
    # Find the highest message number (total message count)
    message_numbers = []
    for _, row in df.iloc[1:].iterrows():
        try:
            msg_num = int(str(row.iloc[0]).strip())
            message_numbers.append(msg_num)
        except (ValueError, TypeError):
            continue
    
    total_messages = max(message_numbers) if message_numbers else 0
    print(f"Total Messages: {total_messages}")
    
    # Generate the report
    generate_chat_report(args.excel_file, args.output, verbose=args.verbose)
    print(f"PDF report has been generated: {args.output}")
