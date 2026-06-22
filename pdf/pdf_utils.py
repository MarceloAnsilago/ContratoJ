import base64

import streamlit as st


def mostrar_pdf_na_tela(pdf_bytes: bytes):
    base64_pdf = base64.b64encode(pdf_bytes).decode("utf-8")
    pdf_iframe = f"""
        <iframe
            src="data:application/pdf;base64,{base64_pdf}"
            width="100%"
            height="800px"
            style="border:none;">
        </iframe>
    """
    st.markdown(pdf_iframe, unsafe_allow_html=True)
