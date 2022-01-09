import docx
from docx import Document
from docx.enum.section import WD_SECTION_START
from docx.shared import Cm, Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_BREAK
from docx.enum.style import WD_STYLE_TYPE

def break_lines(lines):
    for i in range(10 - lines):
        if i==0:
            brl=""
        else:
            brl+="\n"
    return brl
'''
Conteudo das primeiras páginas
'''
assign_line="______________________________"
work_type=["Tipo do Trabalho"]

loc_year=["Cidade","Ano de Conclusão"]
title=["Título","Subtítulo"]

school=["Nome da Escola","Curso em Curso",loc_year[0]]
names=[f"Aluno0{i}" for i in range(5)]

page01=[school,names,title,work_type,loc_year]
print(page01)
'''
Configurando as páginas
'''
doc = Document()

def margin_config():
    sections = doc.sections
    margins = [ 3.0, 2.5 ]

    for sec in sections:
        sec.top_margin = Cm(margins[0])
        sec.left_margin = Cm(margins[0])
        sec.bottom_margin = Cm(margins[1])
        sec.right_margin = Cm(margins[1])

def conf_style_font():
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(14)

def wrt_page(page):
    p = doc.add_paragraph()
    p.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    size_brk_line=len(names)
    
    conf_style_font()

    for i, idx in enumerate(page):
        aux=[]
        if i==0:
            p.add_run(f'{break_lines(size_brk_line)}')
        aux = page[i]
        for j, idx in enumerate(aux):
            aux[j]+="\n" if j!=(len(aux)-1) else break_lines(size_brk_line)
            p.add_run(aux[j].upper(),0).bold = True

if __name__ == "__main__":
    wrt_page(page01)
    doc.save("test.docx")
