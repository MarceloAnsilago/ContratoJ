from io import BytesIO
from html import escape

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer


def _linhas(texto: str) -> list[str]:
    return texto.replace("\r\n", "\n").replace("\r", "\n").split("\n")


def gerar_contrato_pdf(texto: str, titulo: str = "Contrato") -> bytes:
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=2 * cm,
        leftMargin=2 * cm,
        topMargin=2 * cm,
        bottomMargin=2 * cm,
        title=titulo,
    )

    styles = getSampleStyleSheet()
    normal = ParagraphStyle(
        "ContratoNormal",
        parent=styles["Normal"],
        fontName="Courier",
        fontSize=12,
        leading=18,
        alignment=4,
        firstLineIndent=1.25 * cm,
        spaceAfter=10,
    )
    heading = ParagraphStyle(
        "ContratoTitulo",
        parent=styles["Title"],
        fontName="Courier-Bold",
        fontSize=14,
        leading=18,
        alignment=1,
        spaceAfter=18,
    )

    elementos = []
    primeira_linha = True
    for linha in _linhas(texto):
        bloco = linha.strip()
        if not bloco:
            elementos.append(Spacer(1, 8))
            continue

        indice = 0 if primeira_linha else 1
        estilo = heading if indice == 0 else normal
        elementos.append(Paragraph(escape(bloco), estilo))
        if indice == 0:
            elementos.append(Spacer(1, 8))
            primeira_linha = False

    doc.build(elementos)
    buffer.seek(0)
    return buffer.getvalue()


def gerar_contrato_docx(texto: str, titulo: str = "Contrato") -> bytes:
    documento = Document()
    documento.core_properties.title = titulo

    estilo = documento.styles["Normal"]
    estilo.font.name = "Courier New"
    estilo.font.size = Pt(12)

    primeira_linha = True
    for linha in _linhas(texto):
        bloco = linha.strip()
        paragrafo = documento.add_paragraph()
        if not bloco:
            paragrafo.paragraph_format.space_after = Pt(4)
            continue

        indice = 0 if primeira_linha else 1
        paragrafo.alignment = WD_ALIGN_PARAGRAPH.CENTER if indice == 0 else WD_ALIGN_PARAGRAPH.JUSTIFY
        paragrafo.paragraph_format.first_line_indent = Pt(0 if indice == 0 else 36)
        paragrafo.paragraph_format.space_after = Pt(10)

        run = paragrafo.add_run(bloco)
        run.bold = indice == 0
        run.font.name = "Courier New"
        run.font.size = Pt(14 if indice == 0 else 12)
        primeira_linha = False

    buffer = BytesIO()
    documento.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()
