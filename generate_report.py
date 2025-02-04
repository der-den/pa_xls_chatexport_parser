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

# Read the Excel file
excel_file = 'sample/Report.xlsx'
df = pd.read_excel(excel_file, na_values=[''], keep_default_na=False)

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

class ChatReport:
    def __init__(self):
        self.message_count = 0
        self.participants = []
        self.page_height = A4[1]
        self.page_width = A4[0]
        self.margin = 50
        self.y_position = self.page_height - self.margin
        self.line_height = 15
        self.max_image_height = 120  # Maximum height for embedded images
        
        # Register fonts
        font_path = Path(__file__).parent / 'fonts'
        pdfmetrics.registerFont(TTFont('DejaVuSans', str(font_path / 'DejaVuSans.ttf')))
        pdfmetrics.registerFont(TTFont('SegoeEmoji', 'C:/Windows/Fonts/seguiemj.ttf'))
        
    def is_emoji(self, char):
        return len(char) == 1 and ord(char) > 127 or char in ['👍', '😊', '😂', '❤️', '😍', '🤣', '😅', '😁']
        
    def draw_text_with_emojis(self, canvas, text, x, y, base_font='DejaVuSans', size=10):
        current_x = x
        for char in text:
            if self.is_emoji(char):
                canvas.setFont('SegoeEmoji', size)
            else:
                canvas.setFont(base_font, size)
            width = canvas.stringWidth(char, canvas._fontname, size)
            canvas.drawString(current_x, y, char)
            current_x += width
        return current_x - x  # Return total width
        
    def new_page(self, canvas):
        canvas.setFont('DejaVuSans', 12)
        self.y_position = self.page_height - 60
        
    def add_chat_line(self, canvas, message_data):
        if self.y_position < self.margin + self.line_height:
            canvas.showPage()
            self.new_page(canvas)
            
        sender_name = str(message_data.get('sender_name', ''))
        message_body = str(message_data.get('body', ''))
        timestamp = str(message_data.get('timestamp', ''))
        is_owner = message_data.get('is_owner', False)
        is_read = str(message_data.get('Status', ''))
        read_status = "✔️ Read" if "read" in is_read.lower() else ""
        attachment = str(message_data.get('attachment', ''))
        
        # Only print if we have an image attachment
        if attachment and attachment != 'nan' and self.is_image_file(attachment):
            print(f"Found image attachment: {attachment}")
            
        # Calculate positions
        left_col = self.margin
        middle_col = self.margin + 100
        right_col = self.page_width - self.margin - 100
        
        # Handle message content
        y_offset = 0
        
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
                if self.is_image_file(attachment):
                    # Add some spacing before image
                    y_offset += 5
                    image_height = self.embed_image(canvas, attachment, middle_col, self.y_position - y_offset, self.max_image_height)
                    if image_height > 0:
                        y_offset += image_height + 5
                else:
                    # Display attachment name in smaller font
                    canvas.setFont('DejaVuSans', 8)
                    canvas.drawString(middle_col, self.y_position - y_offset, f"📎 {attachment}")
                    y_offset += 10
            
            # Right: read status
            canvas.setFont('SegoeEmoji', 8)
            canvas.drawString(right_col, self.y_position, read_status)
        else:
            # Similar structure for non-owner messages
            # Left: read status
            canvas.setFont('SegoeEmoji', 8)
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
                if self.is_image_file(attachment):
                    # Add some spacing before image
                    y_offset += 5
                    image_height = self.embed_image(canvas, attachment, middle_col, self.y_position - y_offset, self.max_image_height)
                    if image_height > 0:
                        y_offset += image_height + 5
                else:
                    # Display attachment name in smaller font
                    canvas.setFont('DejaVuSans', 8)
                    canvas.drawString(middle_col, self.y_position - y_offset, f"📎 {attachment}")
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

    def is_image_file(self, filename):
        """Check if the filename has an image extension."""
        if not filename:
            return False
        image_extensions = {'.jpg', '.jpeg', '.png', '.gif', '.bmp'}
        return Path(filename).suffix.lower() in image_extensions

    def embed_image(self, canvas, image_path, x, y, max_height):
        """Embed an image in the PDF with maximum height constraint."""
        try:
            if not os.path.exists(image_path):
                return 0  # Return 0 height if image doesn't exist
                
            img = Image.open(image_path)
            img_width, img_height = img.size
            
            # Calculate new dimensions maintaining aspect ratio
            aspect = img_width / img_height
            new_height = min(max_height, img_height)
            new_width = new_height * aspect
            
            # Draw the image
            canvas.drawImage(image_path, x, y - new_height, width=new_width, height=new_height)
            return new_height
        except Exception as e:
            print(f"Error embedding image {image_path}: {e}")
            return 0

    def calculate_text_width(self, canvas, text):
        """Calculate the width of text considering emojis."""
        width = 0
        for char in text:
            if self.is_emoji(char):
                canvas.setFont('SegoeEmoji', 10)
            else:
                canvas.setFont('DejaVuSans', 10)
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

def generate_chat_report(excel_file, output_file):
    # Read the Excel file
    df = pd.read_excel(excel_file, na_values=[''], keep_default_na=False)
    
    # Create PDF
    c = canvas.Canvas(output_file, pagesize=A4)
    report = ChatReport()
    
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
            if Path(attachment).suffix.lower() in {'.jpg', '.jpeg', '.png', '.gif', '.bmp'}:
                image_attachments.append(attachment)
        
        if sender and sender != 'System Message System Message':
            chat_id, _ = parse_participant(sender)
            if chat_id and chat_id in participants_dict:
                # Check if this is an outgoing message
                if direction == 'outgoing':
                    participants_dict[chat_id]['is_owner'] = True
                    break

    if image_attachments:
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
                    body_content = f"📎 {attachment}"
                else:
                    body_content = "no valid body"
            
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
            
            messages.append(message_data)

    # Process each message
    for message in messages:
        report.add_chat_line(c, message)
    
    # Save the PDF
    c.save()

# Generate the report
output_file = 'chat_report.pdf'
generate_chat_report(excel_file, output_file)
print(f"PDF report has been generated: {output_file}")
