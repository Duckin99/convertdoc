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
        # 1. Text Boxes -> Render as JSON Code Blocks
        box_nodes = elem.xpath('.//*[local-name()="txbxContent"]')
        if box_nodes:
            for node in box_nodes:
                box_lines = []
                for p_node in node.xpath('.//*[local-name()="p"]'):
                    p = Paragraph(p_node, self.doc)
                    if p.text.strip():
                        # Extract raw text without markdown styling for valid JSON
                        box_lines.append(p.text)
                
                if box_lines:
                    json_str = "\n".join(box_lines)
                    self.chunks.append(f"\n```json\n{json_str}\n```\n")
            return # Exit early so we don't double-process the paragraphs inside the box

        # 2. Standard Paragraphs
        if elem.tag.endswith('p'):
            p = Paragraph(elem, self.doc)
            self.chunks.append(TextParser.parse_p(p))
            self._find_images(elem)

        # 3. Native Tables
        elif elem.tag.endswith('tbl'):
            self.chunks.append(TableParser.parse_tbl(elem, self.doc))

        # 4. Embedded OLE Objects (.bin) -> Map to Preview Images
        elif elem.xpath('.//*[local-name()="OLEObject"]'):
            # Instead of linking the .bin, we hunt for the visual preview image Word generated
            preview_img = elem.xpath('.//*[local-name()="imagedata"]/@r:id')
            if preview_img:
                self._handle_image(preview_img[0])
            else:
                # Fallback if absolutely no preview image exists
                for obj in elem.xpath('.//*[local-name()="OLEObject"]'):
                    rId = obj.get(qn('r:id'))
                    if rId:
                        fname = os.path.basename(self.doc.part.related_parts[rId].partname)
                        self.chunks.append(f"\n[Raw Embedded Data: {fname}](./embeddings/{fname})\n")

    def _handle_image(self, rId: str):
        part = self.doc.part.related_parts[rId]
        fname = os.path.basename(part.partname)
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