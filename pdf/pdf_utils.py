from io import BytesIO

import fitz
import streamlit as st


def mostrar_pdf_na_tela(pdf_bytes: bytes):
    documento = fitz.open(stream=pdf_bytes, filetype="pdf")

    for numero_pagina, pagina in enumerate(documento, start=1):
        pixmap = pagina.get_pixmap(matrix=fitz.Matrix(1.6, 1.6), alpha=False)
        imagem = BytesIO(pixmap.tobytes("png"))
        st.image(imagem, caption=f"Pagina {numero_pagina}", width="stretch")
