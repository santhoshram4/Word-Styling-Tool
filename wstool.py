import os
import re
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def remove_table_borders(table):
    tbl = table._element
    tblPr = tbl.xpath('w:tblPr')[0]
    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'nil') 
        tblBorders.append(border)
    tblPr.append(tblBorders)

def set_column_widths_22_78(table):
    widths = [Inches(1.43), Inches(5.07)] 
    for row in table.rows:
        for idx, width in enumerate(widths):
            row.cells[idx].width = width

def create_page_number(run):
    """Word dynamic page number field-ah inject pannum"""
    fldSimple = OxmlElement('w:fldSimple')
    fldSimple.set(qn('w:instr'), r'PAGE')
    run._r.append(fldSimple)

def setup_footer(doc, citation_name="citation"):
    """
    image_6327c0.png-la irukura maari running text-ah remove pannitu 
    'Page X of citation' format-ah set pannum.
    """
    section = doc.sections[0]
    footer = section.footer
    
    # Pathaya footer paragraphs-ah remove panrom (running text delete aagum)
    for p in footer.paragraphs:
        p_element = p._element
        p_element.getparent().remove(p_element)
    
    # Pudhu footer paragraph
    new_para = footer.add_paragraph()
    new_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    
    # "Page " text
    run = new_para.add_run("Page ")
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)
    run.font.color.rgb = RGBColor(0, 0, 128) # Blue color image-la irukura maari
    
    # Dynamic Page Number
    page_run = new_para.add_run()
    create_page_number(page_run)
    page_run.font.name = 'Times New Roman'
    page_run.font.size = Pt(11)
    page_run.font.color.rgb = RGBColor(0, 0, 128)
    
    # " of [citation]" text
    run_end = new_para.add_run(f" of {citation_name}")
    run_end.font.name = 'Times New Roman'
    run_end.font.size = Pt(11)
    run_end.font.color.rgb = RGBColor(0, 0, 128)

def apply_legal_template_final(doc):
    paragraphs = list(doc.paragraphs)
    anchor_index = -1
    
    for i, para in enumerate(paragraphs):
        clean_text = para.text.replace(" ", "").replace("-", "").upper()
        if any(kw in clean_text for kw in ["JUDGMENT", "INTRODUCTION", "BACKGROUND"]):
            anchor_index = i
            break
            
    if anchor_index != -1:
        for i in range(anchor_index):
            p = paragraphs[i]._element
            p.getparent().remove(p)

    header_para = doc.paragraphs[0].insert_paragraph_before()
    header_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    case_name = "Name of the case"
    citation_text = "citation" # Intha citation text thaan footer-laiyu varum

    run1 = header_para.add_run(case_name)
    run1.font.name = 'Times New Roman'
    run1.font.size = Pt(14) 
    run1.font.bold = True
    run1.font.color.rgb = RGBColor(0, 0, 128)
    
    header_para.add_run().add_break() 
    
    run2 = header_para.add_run(citation_text)
    run2.font.name = 'Times New Roman'
    run2.font.size = Pt(14)
    run2.font.bold = True
    run2.font.color.rgb = RGBColor(0, 0, 128)

    # Footer-ah inga setup panrom
    setup_footer(doc, citation_text)

    table_obj = doc.add_table(rows=15, cols=2)
    remove_table_borders(table_obj)
    set_column_widths_22_78(table_obj) 
    
    header_para._element.addnext(table_obj._element)
    
    labels = [
        "Reported in:", "Case No:", "Judgment Date(s):", "Hearing Date(s):",
        "Marked as:", "Country:", "Jurisdiction:", "Division:", "Judge:",
        "Bench:", "Parties:", "Appearance:", "Categories:", "Function:", "Relevant Legislation:"
    ]
    
    for r, label in enumerate(labels):
        cell_left = table_obj.cell(r, 0)
        p_l = cell_left.paragraphs[0]
        p_l.paragraph_format.space_after = Pt(4) 
        rl = p_l.add_run(label)
        rl.font.bold = True
        rl.font.name = 'Times New Roman'
        rl.font.size = Pt(11) 
        
        cell_right = table_obj.cell(r, 1)
        p_r = cell_right.paragraphs[0]
        p_r.paragraph_format.space_after = Pt(4)
        val = ""
        if label == "Reported in:": val = "Kenya Judgments Online, a LexisNexis Electronic Law Report Series"
        elif label == "Marked as:": val = "Unmarked"
        elif label == "Country:": val = "Kenya"
        elif label in ["Categories:", "Function:", "Relevant Legislation:"]: val = "NA"
        
        rr = p_r.add_run(val)
        rr.font.name = 'Times New Roman'
        rr.font.size = Pt(11)

def apply_checklist_rules(doc):
    for para in doc.paragraphs:
        if any(run.font.color.rgb == RGBColor(0, 0, 128) for run in para.runs):
            continue
        for run in para.runs:
            u, b, it = run.underline, run.bold, run.italic
            run.font.name = 'Times New Roman'
            run.font.size = Pt(11)
            run.underline, run.bold, run.italic = u, b, it

    replacements = {
        r'\bfootnote\b': 'fn', r'\bsection\b': 's', r'\bparagraph\b': 'para',
        r'\bregulation\b': 'reg', r'\bsubsection\b': 'ss', r'\brule\b': 'r',
        r' percent': '%', r'(\d+)\s%': r'\1%', r'(\d+),00': r'\1'
    }

    for para in doc.paragraphs:
        for run in para.runs:
            text = run.text.replace("  ", " ")
            for pattern, replacement in replacements.items():
                text = re.sub(pattern, replacement, text, flags=re.IGNORECASE)
            
            text = re.sub(r'(\d+)\s?centimetres', r'\1cm', text)
            text = re.sub(r'(\d+)\s?kilograms', r'\1kg', text)
            text = re.sub(r'(\d+)\s?kilometers', r'\1km', text)
            text = re.sub(r'(\d+)(st|nd|rd|th)\s([A-Z][a-z]+)\s(\d{4})', r'\1 \3 \4', text)
            run.text = text

def process_batch():
    input_dir = 'input'
    output_dir = 'output'
    if not os.path.exists(output_dir): os.makedirs(output_dir)

    for filename in os.listdir(input_dir):
        if filename.lower().endswith('.docx'):
            print(f"Processing: {filename}")
            try:
                doc = Document(os.path.join(input_dir, filename))
                apply_legal_template_final(doc)
                apply_checklist_rules(doc)
                new_name = os.path.splitext(filename)[0] + ".docx"
                doc.save(os.path.join(output_dir, new_name))
            except Exception as e:
                print(f"Error: {e}")

if __name__ == "__main__":
    process_batch()