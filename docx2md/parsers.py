import re
from docx.table import Table

class TextParser:
    @staticmethod
    def parse_p(p_obj) -> str:
        """Parses standard paragraphs and ignores colors."""
        if not p_obj.text.strip():
            return ""

        fmt_txt = ""
        for run in p_obj.runs:
            raw = run.text
            if not raw:
                continue
            
            match = re.match(r'^(\s*)(.*?)(\s*)$', raw, re.DOTALL)
            l_space, core, r_space = match.groups()

            if not core:
                fmt_txt += raw
                continue
            
            if run.bold:
                core = f"**{core}**"
            if run.italic:
                core = f"*{core}*"
            if run.underline:
                core = f"<u>{core}</u>"
                
            fmt_txt += f"{l_space}{core}{r_space}"

        style = p_obj.style.name.lower()
        
        if 'title' in style: return f"\n# {fmt_txt}\n"
        if 'subtitle' in style: return f"\n## {fmt_txt}\n"
        if style.startswith('heading'):
            lvl = ''.join(filter(str.isdigit, style))
            return f"\n{'#' * int(lvl)} {fmt_txt}\n" if lvl else f"\n# {fmt_txt}\n"
        if 'list bullet' in style: return f"* {fmt_txt}\n"
        if 'list number' in style: return f"1. {fmt_txt}\n" 
        
        # ToC items usually have 'toc' in the style name. We return them as plain lists.
        if 'toc' in style: return f"- {fmt_txt}\n"

        return f"\n{fmt_txt}\n"

class TableParser:
    @staticmethod
    def parse_tbl(tbl_elem, doc) -> str:
        tbl = Table(tbl_elem, doc)
        if not tbl.rows:
            return ""

        md_tbl = []
        
        def clean(cell):
            txt = " ".join(p.text.strip() for p in cell.paragraphs if p.text.strip())
            return txt.replace("|", "\\|").replace("\n", " ").replace("\r", "")

        head = [clean(c) for c in tbl.rows[0].cells]
        md_tbl.append("| " + " | ".join(head) + " |")
        md_tbl.append("| " + " | ".join(["---"] * len(head)) + " |")
        
        for row in tbl.rows[1:]:
            cells = [clean(c) for c in row.cells]
            if any(cells):
                md_tbl.append("| " + " | ".join(cells) + " |")
            
        return "\n" + "\n".join(md_tbl) + "\n"