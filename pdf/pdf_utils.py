from io import BytesIO
from base64 import b64encode

import fitz
import streamlit as st


def mostrar_pdf_na_tela(pdf_bytes: bytes):
    documento = fitz.open(stream=pdf_bytes, filetype="pdf")

    for numero_pagina, pagina in enumerate(documento, start=1):
        pixmap = pagina.get_pixmap(matrix=fitz.Matrix(1.6, 1.6), alpha=False)
        imagem = BytesIO(pixmap.tobytes("png"))
        imagem_base64 = b64encode(imagem.getvalue()).decode("ascii")
        st.markdown(
            f"""
            <div class="pdf-preview-frame">
                <img src="data:image/png;base64,{imagem_base64}" alt="Página {numero_pagina}" />
            </div>
            <div class="pdf-preview-caption">Página {numero_pagina}</div>
            """,
            unsafe_allow_html=True,
        )
