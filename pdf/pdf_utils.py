from io import BytesIO

import fitz
import streamlit as st
from PIL import Image, ImageFilter


def adicionar_moldura(imagem_bytes: bytes) -> BytesIO:
    pagina = Image.open(BytesIO(imagem_bytes)).convert("RGB")

    margem = 42
    deslocamento_sombra = 9
    largura = pagina.width + (margem * 2) + deslocamento_sombra
    altura = pagina.height + (margem * 2)

    fundo = Image.new("RGB", (largura, altura), "#f0f2f6")
    sombra = Image.new("RGBA", (10, pagina.height), (15, 23, 42, 50))
    sombra = sombra.filter(ImageFilter.GaussianBlur(4))
    fundo.paste(sombra, (margem + pagina.width - 2, margem), sombra)
    fundo.paste(pagina, (margem, margem))

    saida = BytesIO()
    fundo.save(saida, format="PNG")
    saida.seek(0)
    return saida


def mostrar_pdf_na_tela(pdf_bytes: bytes):
    documento = fitz.open(stream=pdf_bytes, filetype="pdf")

    for numero_pagina, pagina in enumerate(documento, start=1):
        pixmap = pagina.get_pixmap(matrix=fitz.Matrix(1.6, 1.6), alpha=False)
        imagem = adicionar_moldura(pixmap.tobytes("png"))
        st.image(imagem, caption=f"Página {numero_pagina}", width="stretch")
