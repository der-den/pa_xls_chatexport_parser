import urllib.request
import os
import zipfile
import tempfile
import shutil

# Create fonts directory if it doesn't exist
os.makedirs('fonts', exist_ok=True)

# Download Symbola font
print("Downloading Symbola font...")
symbola_url = 'https://raw.githubusercontent.com/ChiefMikeK/ttf-symbola/master/Symbola-13.otf'
symbola_path = os.path.join('fonts', 'Symbola.ttf')

try:
    # Download Symbola font directly
    urllib.request.urlretrieve(symbola_url, symbola_path)
    print("Downloaded Symbola font successfully")
except Exception as e:
    print(f"Error downloading Symbola font: {e}")

# Download and extract DejaVu Sans
print("Downloading DejaVu fonts...")
zip_url = 'https://downloads.sourceforge.net/project/dejavu/dejavu/2.37/dejavu-fonts-ttf-2.37.zip'
temp_zip = tempfile.mktemp('.zip')

try:
    # Download zip file
    urllib.request.urlretrieve(zip_url, temp_zip)
    print("Downloaded DejaVu font package")

    # Extract only DejaVuSans.ttf
    with zipfile.ZipFile(temp_zip) as zip_ref:
        for zip_path in zip_ref.namelist():
            if zip_path.endswith('DejaVuSans.ttf'):
                print("Extracting DejaVuSans.ttf...")
                source = zip_ref.open(zip_path)
                target = open(os.path.join('fonts', 'DejaVuSans.ttf'), "wb")
                with source, target:
                    shutil.copyfileobj(source, target)
                break

finally:
    # Clean up
    if os.path.exists(temp_zip):
        os.remove(temp_zip)

# Remove unused font files
for file in os.listdir('fonts'):
    if file not in ['DejaVuSans.ttf', 'Symbola.ttf']:
        try:
            os.remove(os.path.join('fonts', file))
            print(f"Removed unused font: {file}")
        except Exception as e:
            print(f"Error removing {file}: {e}")

print("Font installation complete!")
