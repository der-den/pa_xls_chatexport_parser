import urllib.request
import os
import zipfile
import tempfile
import shutil

def download_font(url, filename):
    print(f"Downloading {filename}...")
    urllib.request.urlretrieve(url, f"fonts/{filename}")
    print(f"Downloaded {filename}")

# DejaVu font URLs from a more reliable source
fonts = {
    'DejaVuSans.ttf': 'https://downloads.sourceforge.net/project/dejavu/dejavu/2.37/dejavu-fonts-ttf-2.37.zip',
    'DejaVuSans-Bold.ttf': 'https://downloads.sourceforge.net/project/dejavu/dejavu/2.37/dejavu-fonts-ttf-2.37.zip',
    'DejaVuSans-Oblique.ttf': 'https://downloads.sourceforge.net/project/dejavu/dejavu/2.37/dejavu-fonts-ttf-2.37.zip'
}

# Create fonts directory if it doesn't exist
os.makedirs('fonts', exist_ok=True)

# Download Noto Color Emoji font
print("Downloading Noto Color Emoji font...")
noto_url = 'https://github.com/googlefonts/noto-emoji/raw/main/fonts/NotoColorEmoji.ttf'
noto_path = os.path.join('fonts', 'NotoColorEmoji.ttf')

try:
    urllib.request.urlretrieve(noto_url, noto_path)
    print("Downloaded Noto Color Emoji font")
except Exception as e:
    print(f"Error downloading Noto font: {e}")

# Download and extract the zip file
print("Downloading DejaVu fonts...")
zip_url = 'https://downloads.sourceforge.net/project/dejavu/dejavu/2.37/dejavu-fonts-ttf-2.37.zip'
temp_zip = tempfile.mktemp('.zip')

try:
    # Download zip file
    urllib.request.urlretrieve(zip_url, temp_zip)
    print("Downloaded font package")

    # Extract required files
    with zipfile.ZipFile(temp_zip) as zip_ref:
        for zip_path in zip_ref.namelist():
            if zip_path.endswith(('.ttf',)) and 'DejaVuSans' in zip_path:
                filename = os.path.basename(zip_path)
                print(f"Extracting {filename}...")
                source = zip_ref.open(zip_path)
                target = open(os.path.join('fonts', filename), "wb")
                with source, target:
                    shutil.copyfileobj(source, target)

finally:
    # Clean up
    if os.path.exists(temp_zip):
        os.remove(temp_zip)

print("Font installation complete!")
