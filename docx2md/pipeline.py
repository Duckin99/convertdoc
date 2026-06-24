import os
from pathlib import Path
from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from .assets import AssetExtractor
from .parsers import TextParser, TableParser

class DocxPipeline:
    def __init__(self, doc_path: str, out_dir: str = "output"):
        self.doc_path = doc_path
        self.out_dir = Path(out_dir)
        self.doc = Document(doc_path)
        self.chunks = []
        
        self.assets = AssetExtractor(doc_path, self.out_dir)
        self.vision = None
        self.md_agent = None

    def attach_agents(self, vision_agent=None, md_agent=None):
        self.vision = vision_agent
        self.md_agent = md_agent

    def execute(self):
        self.assets.extract()
        for elem in self.doc.element.body:
            self._route(elem)
        self._finalize()

    def _route(self, elem):
        # 1. Standard Paragraphs
        if elem.tag.endswith('p'):
            p = Paragraph(elem, self.doc)
            self.chunks.append(TextParser.parse_p(p))
            self._find_images(elem)

        # 2. Native Tables
        elif elem.tag.endswith('tbl'):
            self.chunks.append(TableParser.parse_tbl(elem, self.doc))

        # 3. Text Boxes (Floating shapes in Word)
        box_nodes = elem.xpath('.//*[local-name()="txbxContent"]//*[local-name()="p"]')
        for node in box_nodes:
            p = Paragraph(node, self.doc)
            self.chunks.append(f"\n> [TextBox]: {TextParser.parse_p(p).strip()}\n")

        # 4. Embedded OLE Objects (.bin / .xlsx)
        for obj in elem.xpath('.//*[local-name()="OLEObject"]'):
            rId = obj.get(qn('r:id'))
            if rId:
                fname = os.path.basename(self.doc.part.related_parts[rId].partname)
                self.chunks.append(f"\n[Embedded Data: {fname}](./embeddings/{fname})\n")

    def _find_images(self, elem):
        imgs = elem.xpath('.//w:drawing//pic:pic//hlLink | .//w:drawing//pic:pic//a:blip/@r:embed')
        for rId in imgs:
            fname = os.path.basename(self.doc.part.related_parts[rId].partname)
            
            # Use PNG if we converted an EMF
            img_path = self.assets.media_dir / fname
            if fname.lower().endswith('.emf'):
                img_path = img_path.with_suffix('.png')

            if self.vision:
                print(f"Agent interpreting image: {img_path.name}...")
                self.chunks.append(self.vision.read_img(img_path))
            else:
                self.chunks.append(f"\n![Image](./media/{img_path.name})\n")

    def _finalize(self):
        raw_md = "".join(self.chunks)
        
        if self.md_agent:
            print("Running final Agent cleanup on Markdown...")
            final_md = self.md_agent.clean_markdown(raw_md)
        else:
            final_md = raw_md

        out_file = self.out_dir / "final_document.md"
        with open(out_file, "w", encoding="utf-8") as f:
            f.write(final_md)
        print(f"Success! Saved to {out_file}")