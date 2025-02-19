from dataclasses import dataclass
from datetime import datetime
from typing import List, Optional
from pathlib import Path

@dataclass
class Participant:
    """Repräsentiert einen Chat-Teilnehmer"""
    chat_id: str
    name: str
    is_owner: bool = False

@dataclass
class Attachment:
    """Repräsentiert einen Anhang (Bild, Audio, Video)"""
    filename: str
    full_path: Optional[Path] = None
    type: str = "unknown"  # "image", "audio", "video"
    transcription: Optional[str] = None

@dataclass
class Message:
    """Repräsentiert eine einzelne Chat-Nachricht"""
    sender: Participant
    body: str
    timestamp: datetime
    status: str
    attachment: Optional[Attachment] = None

@dataclass
class ChatExport:
    """Hauptklasse für den Chat-Export"""
    participants: List[Participant]
    messages: List[Message]
    excel_path: Path
    
    @property
    def message_count(self) -> int:
        return len(self.messages)
    
    @property
    def attachment_stats(self) -> dict:
        """Berechnet Statistiken über Anhänge"""
        found = 0
        not_found = 0
        for msg in self.messages:
            if msg.attachment:
                if msg.attachment.full_path:
                    found += 1
                else:
                    not_found += 1
        return {
            "found": found,
            "not_found": not_found
        }

class ChatExportReader:
    """Basis-Klasse für verschiedene Chat-Export-Reader"""
    def read(self, file_path: Path) -> ChatExport:
        raise NotImplementedError

class ExcelChatExportReader(ChatExportReader):
    """Liest Chat-Exports aus Excel-Dateien"""
    def read(self, file_path: Path) -> ChatExport:
        import pandas as pd
        
        # Excel einlesen
        df = pd.read_excel(file_path, na_values=[''], keep_default_na=False)
        
        # Teilnehmer sammeln
        participants_dict = {}
        messages = []
        
        def parse_participant(participant_str: str) -> tuple[str, str]:
            """Parse participant string into chat ID and name."""
            parts = participant_str.strip().split()
            if len(parts) >= 2:
                chat_id = parts[0]
                name = ' '.join(parts[1:])
                if chat_id.isdigit():
                    return chat_id, name
            return None, participant_str

        # Erste Runde: Teilnehmer identifizieren
        for _, row in df.iloc[1:].iterrows():
            sender = str(row.iloc[1]).strip()
            if sender and sender != 'nan' and sender != 'From' and sender != 'System Message System Message':
                chat_id, name = parse_participant(sender)
                if chat_id and name:
                    if chat_id not in participants_dict:
                        participants_dict[chat_id] = Participant(
                            chat_id=chat_id,
                            name=name,
                            is_owner=str(row.iloc[6]).strip().lower() == 'outgoing'
                        )

        # Zweite Runde: Nachrichten sammeln
        for _, row in df.iloc[1:].iterrows():
            sender = str(row.iloc[1]).strip()
            chat_id, _ = parse_participant(sender)
            
            if chat_id and chat_id in participants_dict:
                body = str(row.iloc[8]).strip()  # Body ist in Spalte 8
                if body == 'nan':
                    body = ""
                
                # Timestamp verarbeiten
                timestamp = str(row.iloc[18]).strip()  # Timestamp-Time
                timestamp = timestamp.replace('(UTC+0)', '').strip()
                if '.' not in timestamp:
                    date = str(row.iloc[17]).strip()  # Timestamp-Date
                    if date != 'nan' and date:
                        timestamp = f"{date} {timestamp}"
                try:
                    timestamp_dt = datetime.strptime(timestamp, "%d.%m.%Y %H:%M:%S")
                except ValueError:
                    timestamp_dt = datetime.now()  # Fallback
                
                # Anhang verarbeiten
                attachment = None
                attachment_name = str(row.iloc[25]).strip()  # Attachment #1
                if attachment_name and attachment_name != 'nan':
                    attachment = Attachment(
                        filename=attachment_name,
                        full_path=None  # Wird später gefüllt
                    )
                
                message = Message(
                    sender=participants_dict[chat_id],
                    body=body,
                    timestamp=timestamp_dt,
                    status=str(row.iloc[9]).strip(),  # Status ist in Spalte 9
                    attachment=attachment
                )
                messages.append(message)
        
        return ChatExport(
            participants=list(participants_dict.values()),
            messages=messages,
            excel_path=Path(file_path)
        )
