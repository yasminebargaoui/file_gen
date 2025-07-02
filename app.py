from flask import Flask, request, jsonify, abort
from io import BytesIO
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from docx.shared import RGBColor, Pt
import base64
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

    run_spacing = paragraph.add_run("      ")  # Manual wide spacing
    run_text = paragraph.add_run(text)
    run_text.font.size = Pt(11)

def insert_horizontal_line_after(paragraph):
    new_p = insert_paragraph_after(paragraph)
    p = new_p._element

    # Ajouter bordure inférieure pointillée, 0.48pt, couleur #3399CC
    pPr = p.get_or_add_pPr()
    pbdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'dotted')        # Style pointillé
    bottom.set(qn('w:sz'), '4')              # Taille 0.48pt
    bottom.set(qn('w:space'), '1')           # Espacement
    bottom.set(qn('w:color'), '3399CC')      # Couleur bleu turquoise
    pbdr.append(bottom)
    pPr.append(pbdr)

    return new_p


@app.route('/modify_docx', methods=['POST'])
def modify_docx():
    try:
        data = request.get_json()
        if not data:
            return abort(400, "Missing JSON body")

        file_base64 = data.get("file_base64")
        competences = data.get("competences")
        if not file_base64 or not competences:
            return abort(400, "Missing file_base64 or competences")

        # Decode DOCX
        docx_bytes = base64.b64decode(file_base64)
        doc = Document(BytesIO(docx_bytes))

        # Supprimer ancien contenu entre "Connaissances Métier" et "COMPETENCES Projet"
        found_section = False
        paras_to_delete = []
        for para in doc.paragraphs:
            if "Connaissances Métier" in para.text:
                found_section = True
                continue
            if found_section and "COMPETENCES Projet" in para.text:
                break
            if found_section:
                paras_to_delete.append(para)
        for para in paras_to_delete:
            delete_paragraph(para)

        # Trouver le paragraphe "Connaissances Métier"
        cm_para = None
        for para in doc.paragraphs:
            if "Connaissances Métier" in para.text:
                cm_para = para
                break

        if cm_para is None:
            return abort(400, "'Connaissances Métier' not found")
        cm_para.paragraph_format.space_after = Pt(1)  # ✅ réduit l’espace sous le titre

        # Insérer le trait bleu juste après "Connaissances Métier"
        trait_para = insert_horizontal_line_after(cm_para)

        # Insérer les bullets après le trait
        previous_para = trait_para
        inserted_paragraphs = []

        for text in competences:
            new_para = insert_paragraph_after(previous_para)
            add_blue_bullet(new_para, text)
            inserted_paragraphs.append(new_para)
            previous_para = new_para

        # Réduire espace après dernier bullet
        if inserted_paragraphs:
            inserted_paragraphs[-1].paragraph_format.space_after = Pt(3)

        # Retourner le fichier modifié
        out_stream = BytesIO()
        doc.save(out_stream)
        out_stream.seek(0)
        return jsonify({
            "base64": base64.b64encode(out_stream.read()).decode("utf-8")
        })

    except Exception as e:
        return abort(500, str(e))

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
