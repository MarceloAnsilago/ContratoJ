from datetime import date

import streamlit as st

from pdf.pdf_utils import mostrar_pdf_na_tela
from services.contract_builder import gerar_contrato_docx, gerar_contrato_pdf


st.set_page_config(
    page_title="Gerador de Contrato",
    page_icon="DOC",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    .stApp .block-container {
        max-width: 1280px;
        margin-left: auto;
        margin-right: auto;
        padding-left: 2rem;
        padding-right: 2rem;
    }
    .stApp .block-container h1,
    .stApp .block-container h2 {
        text-align: center;
    }
    [data-testid="stSidebar"] {
        min-width: 390px;
        max-width: 390px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


def texto_padrao(dados: dict) -> str:
    return f"""CONTRATO DE PRESTACAO DE SERVICOS

Pelo presente instrumento particular, de um lado {dados["contratante_nome"]}, inscrito(a) no CPF/CNPJ sob o no {dados["contratante_doc"]}, com endereco em {dados["contratante_endereco"]}, doravante denominado(a) CONTRATANTE, e de outro lado {dados["contratado_nome"]}, inscrito(a) no CPF/CNPJ sob o no {dados["contratado_doc"]}, com endereco em {dados["contratado_endereco"]}, doravante denominado(a) CONTRATADO(A), resolvem firmar o presente contrato.

CLAUSULA PRIMEIRA - DO OBJETO
O presente contrato tem por objeto {dados["objeto"]}.

CLAUSULA SEGUNDA - DO VALOR E FORMA DE PAGAMENTO
Pela execucao do objeto, o(a) CONTRATANTE pagara ao(a) CONTRATADO(A) o valor de {dados["valor"]}, mediante {dados["pagamento"]}.

CLAUSULA TERCEIRA - DO PRAZO
O prazo de vigencia deste contrato sera de {dados["prazo"]}, com inicio em {dados["data_inicio"]}.

CLAUSULA QUARTA - DAS OBRIGACOES
As partes comprometem-se a cumprir as condicoes ajustadas, observando boa-fe, pontualidade e responsabilidade na execucao do objeto contratado.

CLAUSULA QUINTA - DO FORO
Fica eleito o foro da comarca de {dados["foro"]} para dirimir eventuais controversias oriundas deste contrato.

E por estarem justas e contratadas, as partes assinam o presente instrumento.

{dados["local_data"]}

________________________________________
{dados["contratante_nome"]}
CONTRATANTE

________________________________________
{dados["contratado_nome"]}
CONTRATADO(A)
"""


st.title("Gerador de Contrato")

with st.sidebar:
    st.header("Dados do contrato")

    tipo_contrato = st.selectbox(
        "Tipo de contrato",
        ["Prestacao de servicos", "Locacao", "Compra e venda", "Personalizado"],
    )
    contratante_nome = st.text_input("Contratante", "NOME DO CONTRATANTE")
    contratante_doc = st.text_input("CPF/CNPJ do contratante", "000.000.000-00")
    contratante_endereco = st.text_area("Endereco do contratante", "Endereco completo do contratante")

    contratado_nome = st.text_input("Contratado(a)", "NOME DO CONTRATADO")
    contratado_doc = st.text_input("CPF/CNPJ do contratado", "000.000.000-00")
    contratado_endereco = st.text_area("Endereco do contratado", "Endereco completo do contratado")

    objeto = st.text_area("Objeto", "a prestacao dos servicos descritos pelas partes")
    valor = st.text_input("Valor", "R$ 0,00")
    pagamento = st.text_input("Forma de pagamento", "pagamento conforme acordo entre as partes")
    prazo = st.text_input("Prazo de vigencia", "12 meses")
    data_inicio = st.date_input("Data de inicio", value=date.today(), format="DD/MM/YYYY")
    foro = st.text_input("Foro", "Sao Miguel do Guapore/RO")
    local = st.text_input("Local de assinatura", "Sao Miguel do Guapore/RO")

dados = {
    "tipo_contrato": tipo_contrato,
    "contratante_nome": contratante_nome,
    "contratante_doc": contratante_doc,
    "contratante_endereco": contratante_endereco,
    "contratado_nome": contratado_nome,
    "contratado_doc": contratado_doc,
    "contratado_endereco": contratado_endereco,
    "objeto": objeto,
    "valor": valor,
    "pagamento": pagamento,
    "prazo": prazo,
    "data_inicio": data_inicio.strftime("%d/%m/%Y"),
    "foro": foro,
    "local_data": f"{local}, {date.today().strftime('%d/%m/%Y')}",
}

texto_contrato = texto_padrao(dados)

pdf_bytes = gerar_contrato_pdf(texto_contrato, titulo=f"Contrato - {tipo_contrato}")
docx_bytes = gerar_contrato_docx(texto_contrato, titulo=f"Contrato - {tipo_contrato}")

_, col_preview, _ = st.columns([1, 2, 1])

with col_preview:
    st.subheader("Pre-visualizacao")
    mostrar_pdf_na_tela(pdf_bytes)

    download_col1, download_col2 = st.columns(2)
    with download_col1:
        st.download_button(
            "Baixar PDF",
            data=pdf_bytes,
            file_name="contrato.pdf",
            mime="application/pdf",
            use_container_width=True,
        )
    with download_col2:
        st.download_button(
            "Baixar DOCX",
            data=docx_bytes,
            file_name="contrato.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
