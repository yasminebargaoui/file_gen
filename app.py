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

# --- Utilitaires ---
def insert_paragraph_after(paragraph):
    new_p = OxmlElement("w:p")
    paragraph._element.addnext(new_p)
    return Paragraph(new_p, paragraph._parent)

def delete_paragraph(paragraph):
    p = paragraph._element
    p.getparent().remove(p)
    paragraph._p = paragraph._element = None

def add_blue_bullet(paragraph, text):
    paragraph.paragraph_format.left_indent = Pt(24)
    paragraph.paragraph_format.line_spacing = 1
    paragraph.paragraph_format.space_after = Pt(2)
    run_bullet = paragraph.add_run("■")
    run_bullet.font.color.rgb = RGBColor(60, 122, 178)
    run_bullet.font.size = Pt(8)
    paragraph.add_run("      ")
    run_text = paragraph.add_run(text)
    run_text.font.size = Pt(11)

def insert_horizontal_line_after(paragraph):
    new_p = insert_paragraph_after(paragraph)
    p = new_p._element
    pPr = p.get_or_add_pPr()
    pbdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'dotted')
    bottom.set(qn('w:sz'), '4')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), '3399CC')
    pbdr.append(bottom)
    pPr.append(pbdr)
    new_p.paragraph_format.space_before = 1
    new_p.paragraph_format.space_after = Pt(10)
    return new_p

# --- Route API principale ---
@app.route('/modify_docx', methods=['POST'])
def modify_docx():
    try:
        data = request.get_json()
        if not data:
            return abort(400, "Requête JSON manquante")

        file_base64 = data.get("file_base64")
        competences = data.get("competences")

        if not file_base64 or not competences:
            return abort(400, "Champs 'file_base64' ou 'competences' manquants")

        # Lire le fichier DOCX depuis le base64
        docx_bytes = base64.b64decode(file_base64)
        doc = Document(BytesIO(docx_bytes))

        # Supprimer entre "Connaissances Métier" et "COMPETENCES Projet"
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
        cm_para = next((p for p in doc.paragraphs if "Connaissances Métier" in p.text), None)
        if not cm_para:
            return abort(400, "'Connaissances Métier' non trouvé dans le document")

        cm_para.paragraph_format.space_after = 1
        line_para = insert_horizontal_line_after(cm_para)

        previous_para = line_para
        for comp in competences:
            new_para = insert_paragraph_after(previous_para)
            new_para.paragraph_format.line_spacing = Pt(0)
            add_blue_bullet(new_para, comp)
            previous_para = new_para
        previous_para.paragraph_format.space_after = Pt(9)

        # Sauvegarde dans un buffer mémoire
        output_stream = BytesIO()
        doc.save(output_stream)
        output_stream.seek(0)

        # Encodage base64
        base64_str = base64.b64encode(output_stream.read()).decode('utf-8')

        # Découper en morceaux pour éviter les limites d’IRPA
        chunk_size = 32765  # 50 KB max par chaîne
        base64_parts = [base64_str[i:i + chunk_size] for i in range(0, len(base64_str), chunk_size)]

        return jsonify({
            "base64_parts": base64_parts
        })

    except Exception as e:
        return abort(500, f"Erreur serveur: {str(e)}")

# --- Lancement local ---
if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
