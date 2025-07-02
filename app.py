import base64
import json
from flask import Flask, request, send_file, abort
from io import BytesIO
from docx import Document
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph
from docx.shared import RGBColor, Pt
import os

app = Flask(__name__)

def insert_paragraph_after(paragraph):
    new_p = OxmlElement("w:p")
    paragraph._element.addnext(new_p)
    return Paragraph(new_p, paragraph._parent)

def delete_paragraph(paragraph):
    p = paragraph._element
    p.getparent().remove(p)
    paragraph._p = paragraph._element = None

def add_blue_bullet(paragraph, text):
    paragraph.paragraph_format.left_indent = Pt(36)
    paragraph.paragraph_format.line_spacing = 1
    paragraph.paragraph_format.space_after = Pt(12)

    run_bullet = paragraph.add_run("•")
    run_bullet.font.color.rgb = RGBColor(0, 102, 204)
    run_bullet.font.size = Pt(11)

    run_spacing = paragraph.add_run("      ")
    run_text = paragraph.add_run(text)
    run_text.font.size = Pt(11)

@app.route('/modify_docx', methods=['POST'])
def modify_docx():
    try:
        data = request.get_json()
        file_base64 = data.get('file_base64')
        competences = data.get('competences', [])
        filename = data.get('filename', 'cv.docx')

        if not file_base64 or not competences:
            return abort(400, "Missing file or competences")

        # Decode base64 file content
        file_bytes = base64.b64decode(file_base64)
        file_stream = BytesIO(file_bytes)
        doc = Document(file_stream)

        # Remove old section
        found_section = False
        paras_to_delete = []
        for para in doc.paragraphs:
            if "Connaissances Métier" in para.text:
                found_section = True
                continue
            if found_section:
                if "COMPETENCES Projet" in para.text:
                    break
                paras_to_delete.append(para)
        for para in paras_to_delete:
            delete_paragraph(para)

        # Find where to insert
        insert_index = None
        for i, para in enumerate(doc.paragraphs):
            if "Connaissances Métier" in para.text:
                insert_index = i + 1
                break
        if insert_index is None:
            return abort(400, "Could not find 'Connaissances Métier' section")

        reference_para = doc.paragraphs[insert_index]
        previous_para = reference_para

        inserted_paragraphs = []
        for comp in competences:
            new_para = insert_paragraph_after(previous_para)
            add_blue_bullet(new_para, comp)
            inserted_paragraphs.append(new_para)
            previous_para = new_para

        if inserted_paragraphs:
            inserted_paragraphs[-1].paragraph_format.space_after = Pt(3)

        # Return modified docx as base64
        out_stream = BytesIO()
        doc.save(out_stream)
        out_stream.seek(0)
        encoded = base64.b64encode(out_stream.read()).decode('utf-8')

        return {"base64": encoded}

    except Exception as e:
        return abort(500, f"Error: {str(e)}")

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
