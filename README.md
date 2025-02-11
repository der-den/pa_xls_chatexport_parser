# PA XLS Chat Export Parser

A Python tool for converting Excel-based PA chat exports into a readable PDF format. The tool supports embedding images and transcribing audio and video files.

## Features

- Converts Excel chat exports into a formatted PDF
- Supports emojis and special characters
- Embeds chat images directly into the PDF
- Automatically transcribes audio and video files using OpenAI Whisper
- Caches transcriptions for faster access on subsequent runs

## Installation

1. Clone repository
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Download fonts:
   ```bash
   python download_fonts.py
   ```

## Usage

```bash
python generate_report.py [-h] [-o OUTPUT] [-v] excel_file
```

### Command Line Parameters

- `excel_file`: Path to the Excel file containing chat data (required)
- `-o, --output`: Path to the output PDF file (optional, default: excel_file.pdf)
- `-v, --verbose`: Enable verbose output for debugging or scripting (optional)

### Example

```bash
python generate_report.py chat_export.xlsx
```

## File Structure

The tool expects media files (images, audio, video) to be in a `files` directory parallel to the Excel file.

## Supported Formats

- Images: .jpg, .jpeg, .png, .gif, .bmp
- Audio: .mp3, .wav, .m4a
- Video: .mp4, .avi, .mov

*.mp4 files containing only an audio track are interpreted and treated as audio files.
