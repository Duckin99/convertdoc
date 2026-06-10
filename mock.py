import os
from docx import Document
from docx.shared import Inches
from PIL import Image, ImageDraw

def create_dummy_image(filename="test_diagram.png"):
    """Creates a simple test image to embed in the document."""
    img = Image.new('RGB', (400, 200), color=(73, 109, 137))
    d = ImageDraw.Draw(img)
    d.text((50, 90), "MOCK DIAGRAM IMAGE", fill=(255, 255, 0))
    img.save(filename)
    return filename

def generate_mock_docx(output_filename="sample_large_doc.docx"):
    doc = Document()

    # 1. Title and Text
    doc.add_heading('Mock Technical Specification', 0)
    doc.add_heading('1. Overview', level=1)
    doc.add_paragraph(
        'This is a mock document generated for testing the Python '
        'docx-to-markdown conversion pipeline. It contains multiple formatting types.'
    )

    # 2. Add an Image
    doc.add_heading('2. System Architecture', level=1)
    doc.add_paragraph('Below is a placeholder image representing a Mermaid diagram or flowchart:')
    
    img_path = create_dummy_image()
    doc.add_picture(img_path, width=Inches(3.0))
    
    # Clean up the local image file after embedding it into the docx
    os.remove(img_path)

    # 3. Add a Table
    doc.add_heading('3. API Endpoints', level=1)
    doc.add_paragraph('Here is a native Word table containing route data:')
    
    table = doc.add_table(rows=3, cols=3)
    table.style = 'Table Grid'
    
    # Headers
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Method'
    hdr_cells[1].text = 'Endpoint'
    hdr_cells[2].text = 'Description'
    
    # Row 1
    row_cells = table.rows[1].cells
    row_cells[0].text = 'GET'
    row_cells[1].text = '/api/v1/users'
    row_cells[2].text = 'Fetches all users'
    
    # Row 2
    row_cells = table.rows[2].cells
    row_cells[0].text = 'POST'
    row_cells[1].text = '/api/v1/users'
    row_cells[2].text = 'Creates a new user'

    # Save the final file
    doc.save(output_filename)
    print(f"Successfully created '{output_filename}'")

if __name__ == "__main__":
    generate_mock_docx()