import os
import zipfile
import subprocess
from pathlib import Path

class AssetExtractor:
    def __init__(self, docx_path: str, out_dir: Path):
        self.doc_path = docx_path
        self.media_dir = out_dir / "media"
        self.embed_dir = out_dir / "embeddings"
        
        self.media_dir.mkdir(parents=True, exist_ok=True)
        self.embed_dir.mkdir(parents=True, exist_ok=True)

    def extract(self):
        with zipfile.ZipFile(self.doc_path, 'r') as z:
            for info in z.infolist():
                fname = os.path.basename(info.filename)
                
                if info.filename.startswith('word/media/'):
                    out_path = self.media_dir / fname
                    with open(out_path, "wb") as f:
                        f.write(z.read(info.filename))
                    
                    # Convert EMF to PNG using system ImageMagick
                    if fname.lower().endswith('.emf'):
                        self._convert_emf_to_png(out_path)

                elif info.filename.startswith('word/embeddings/'):
                    out_path = self.embed_dir / fname
                    with open(out_path, "wb") as f:
                        f.write(z.read(info.filename))

    def _convert_emf_to_png(self, emf_path: Path):
        """Requires ImageMagick installed on your Linux/Debian system: sudo apt install imagemagick"""
        png_path = emf_path.with_suffix('.png')
        try:
            subprocess.run(['magick', str(emf_path), str(png_path)], check=True, capture_output=True)
            os.remove(emf_path) # Clean up the EMF
            print(f"Converted {emf_path.name} to PNG.")
        except Exception as e:
            print(f"Could not convert {emf_path.name}. Is ImageMagick installed? Error: {e}")