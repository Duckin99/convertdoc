import os
import zipfile
from pathlib import Path

# Import the new Windows-native converter
from emf_to_png import EMFToPNGConverter

class AssetExtractor:
    def __init__(self, docx_path: str, out_dir: Path):
        self.doc_path = docx_path
        self.media_dir = out_dir / "media"
        self.embed_dir = out_dir / "embeddings"
        
        self.media_dir.mkdir(parents=True, exist_ok=True)
        self.embed_dir.mkdir(parents=True, exist_ok=True)
        
        # Initialize the converter
        self.emf_converter = EMFToPNGConverter()

    def extract(self):
        with zipfile.ZipFile(self.doc_path, 'r') as z:
            for info in z.infolist():
                fname = os.path.basename(info.filename)
                
                if info.filename.startswith('word/media/'):
                    out_path = self.media_dir / fname
                    with open(out_path, "wb") as f:
                        f.write(z.read(info.filename))
                    
                    # Convert EMF to PNG using Python natively
                    if fname.lower().endswith('.emf'):
                        self._convert_emf_to_png(out_path)

                elif info.filename.startswith('word/embeddings/'):
                    out_path = self.embed_dir / fname
                    with open(out_path, "wb") as f:
                        f.write(z.read(info.filename))

    def _convert_emf_to_png(self, emf_path: Path):
        """Converts EMF to PNG using native Windows APIs via python packages."""
        png_path = emf_path.with_suffix('.png')
        try:
            # Run the Python-based conversion
            self.emf_converter.emf_file_to_png_file(str(emf_path), str(png_path))
            
            # Clean up the original messy EMF file so our output folder stays clean
            os.remove(emf_path) 
            print(f"Successfully converted {emf_path.name} to PNG natively.")
        except Exception as e:
            print(f"Native conversion failed for {emf_path.name}. Error: {e}")