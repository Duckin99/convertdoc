import re
from docx.table import Table

class TextParser:
    @staticmethod
    def parse_p(p_obj) -> str:
        if not p_obj.text.strip():
            return ""

        # 1. Inline Formatting
        fmt_txt = ""
        for run in p_obj.runs:
            raw = run.text
            if not raw: continue
            
            match = re.match(r'^(\s*)(.*?)(\s*)$', raw, re.DOTALL)
            l_space, core, r_space = match.groups()

            if not core:
                fmt_txt += raw
                continue
            
            if run.bold: core = f"**{core}**"
            if run.italic: core = f"*{core}*"
            if run.underline: core = f"<u>{core}</u>"
                
            fmt_txt += f"{l_space}{core}{r_space}"

        style = p_obj.style.name.lower()
        
        # 2. Multi-Level List Detection (Checking XML directly)
        num_pr = p_obj._element.xpath('.//*[local-name()="numPr"]')
        if num_pr or 'list' in style:
            # Find the indentation level (default to 0 if not found)
            ilvl_nodes = p_obj._element.xpath('.//*[local-name()="ilvl"]/@w:val')
            lvl = int(ilvl_nodes[0]) if ilvl_nodes else 0
            
            indent = "    " * lvl
            marker = "1." if 'number' in style else "-"
            return f"{indent}{marker} {fmt_txt.strip()}\n"

        # 3. Table of Contents Handling
        if style.startswith('toc'):
            # Extract the ToC level (e.g., 'toc 2' -> level 1)
            lvl_match = re.search(r'\d+', style)
            lvl = int(lvl_match.group()) - 1 if lvl_match else 0
            indent = "    " * lvl
            
            # Remove the page numbers usually found at the end of ToC lines
            clean_txt = re.sub(r'\s+\d+$', '', fmt_txt).strip()
            anchor = clean_txt.lower().replace(' ', '-')
            return f"{indent}- [{clean_txt}](#{anchor})\n"

        # 4. Standard Headings
        if 'title' in style: return f"\n# {fmt_txt}\n"
        if 'subtitle' in style: return f"\n## {fmt_txt}\n"
        if style.startswith('heading'):
            lvl = ''.join(filter(str.isdigit, style))
            return f"\n{'#' * int(lvl)} {fmt_txt}\n" if lvl else f"\n# {fmt_txt}\n"

        return f"\n{fmt_txt}\n"

class TableParser:
    @staticmethod
    def parse_tbl(tbl_elem, doc) -> str:
        tbl = Table(tbl_elem, doc)
        if not tbl.rows: return ""
        md_tbl = []
        def clean(cell):
            txt = " ".join(p.text.strip() for p in cell.paragraphs if p.text.strip())
            return txt.replace("|", "\\|").replace("\n", " ").replace("\r", "")
        head = [clean(c) for c in tbl.rows[0].cells]
        md_tbl.append("| " + " | ".join(head) + " |")
        md_tbl.append("| " + " | ".join(["---"] * len(head)) + " |")
        for row in tbl.rows[1:]:
            cells = [clean(c) for c in row.cells]
            if any(cells): md_tbl.append("| " + " | ".join(cells) + " |")
        return "\n" + "\n".join(md_tbl) + "\n"