from flask import Flask, request, send_file, abort
from io import BytesIO
from docx import Document
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph
from docx.shared import RGBColor, Pt

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
    # Get uploaded file
    if 'file' not in request.files:
        return abort(400, "Missing 'file' part")
    file = request.files['file']
    if file.filename == '':
        return abort(400, "No selected file")

    # Get competencies as JSON list (or as form string split by commas)
    competencies = request.form.get('competences')
    if not competencies:
        return abort(400, "Missing 'competences' data")
    # Assume JSON list or comma-separated string
    try:
        import json
        competencies_list = json.loads(competencies)
        if not isinstance(competencies_list, list):
            raise ValueError
    except Exception:
        # fallback: comma separated string
        competencies_list = [c.strip() for c in competencies.split(',') if c.strip()]

    # Load doc from uploaded file stream
    doc = Document(file)

    # Find and remove old section paragraphs
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

    # Find insertion point (paragraph after "Connaissances Métier")
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
    for comp in competencies_list:
        new_para = insert_paragraph_after(previous_para)
        add_blue_bullet(new_para, comp)
        inserted_paragraphs.append(new_para)
        previous_para = new_para

    if inserted_paragraphs:
        inserted_paragraphs[-1].paragraph_format.space_after = Pt(3)

    # Save modified doc to bytes buffer
    out_stream = BytesIO()
    doc.save(out_stream)
    out_stream.seek(0)

    # Send modified docx as attachment
    return send_file(
        out_stream,
        as_attachment=True,
        download_name='modified.docx',
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )

if __name__ == '__main__':
    app.run(debug=True)
