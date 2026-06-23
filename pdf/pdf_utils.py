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
            <div style="
                background: #eef1f5;
                border: 2px solid #b9c3d2;
                border-radius: 8px;
                box-shadow: 0 14px 30px rgba(15, 23, 42, 0.18);
                margin: 0.75rem auto 0.35rem;
                padding: 1rem;
            ">
                <img
                    src="data:image/png;base64,{imagem_base64}"
                    alt="Página {numero_pagina}"
                    style="
                        display: block;
                        width: 100%;
                        border: 1px solid #d7dce5;
                        box-shadow: 0 2px 8px rgba(15, 23, 42, 0.1);
                    "
                />
            </div>
            <div style="
                color: #5f6b7a;
                font-size: 0.85rem;
                margin-bottom: 1rem;
                text-align: center;
            ">
                Página {numero_pagina}
            </div>
            """,
            unsafe_allow_html=True,
        )
