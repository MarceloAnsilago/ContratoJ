from io import BytesIO

import fitz
import streamlit as st
from PIL import Image, ImageFilter, ImageOps


def adicionar_moldura(imagem_bytes: bytes) -> BytesIO:
    pagina = Image.open(BytesIO(imagem_bytes)).convert("RGB")
    pagina_com_borda = ImageOps.expand(pagina, border=2, fill="#e1e5ec")

    margem = 42
    deslocamento_sombra = 10
    largura = pagina_com_borda.width + (margem * 2) + deslocamento_sombra
    altura = pagina_com_borda.height + (margem * 2) + deslocamento_sombra

    fundo = Image.new("RGB", (largura, altura), "#f7f8fb")
    sombra = Image.new("RGBA", pagina_com_borda.size, (15, 23, 42, 30))
    sombra = sombra.filter(ImageFilter.GaussianBlur(10))
    fundo.paste(sombra, (margem + deslocamento_sombra, margem + deslocamento_sombra), sombra)
    fundo.paste(pagina_com_borda, (margem, margem))

    moldura = ImageOps.expand(fundo, border=2, fill="#d9dee8")
    saida = BytesIO()
    moldura.save(saida, format="PNG")
    saida.seek(0)
    return saida


def mostrar_pdf_na_tela(pdf_bytes: bytes):
    documento = fitz.open(stream=pdf_bytes, filetype="pdf")

    for numero_pagina, pagina in enumerate(documento, start=1):
        pixmap = pagina.get_pixmap(matrix=fitz.Matrix(1.6, 1.6), alpha=False)
        imagem = adicionar_moldura(pixmap.tobytes("png"))
        st.image(imagem, caption=f"Página {numero_pagina}", width="stretch")
