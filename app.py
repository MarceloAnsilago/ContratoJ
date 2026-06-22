from datetime import date

import streamlit as st

from pdf.pdf_utils import mostrar_pdf_na_tela
from services.contract_builder import gerar_contrato_docx, gerar_contrato_pdf


MESES = {
    1: "janeiro",
    2: "fevereiro",
    3: "março",
    4: "abril",
    5: "maio",
    6: "junho",
    7: "julho",
    8: "agosto",
    9: "setembro",
    10: "outubro",
    11: "novembro",
    12: "dezembro",
}


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
        min-width: 430px;
        max-width: 430px;
        background: #16324f;
        border-right: 1px solid #0f253b;
    }
    [data-testid="stSidebar"] [data-testid="stSidebarContent"] {
        padding: 1.25rem 1.15rem;
    }
    [data-testid="stSidebar"] h1,
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3 {
        color: #ffffff;
        letter-spacing: 0;
    }
    [data-testid="stSidebar"] h3 {
        margin-top: 1.1rem;
        padding: 0.55rem 0.7rem;
        border-left: 4px solid #7cb7ff;
        background: #0f253b;
        border-radius: 6px;
        font-size: 1rem;
    }
    [data-testid="stSidebar"] label {
        color: #eef5ff;
        font-weight: 600;
    }
    [data-testid="stSidebar"] p {
        color: #eef5ff;
    }
    [data-testid="stSidebar"] input,
    [data-testid="stSidebar"] textarea {
        border-radius: 6px;
    }
    [data-testid="stSidebar"] [data-baseweb="radio"] {
        background: #ffffff;
        border: 1px solid #d9e1ec;
        border-radius: 8px;
        padding: 0.65rem 0.75rem;
    }
    [data-testid="stSidebar"] [data-baseweb="select"] {
        background: #ffffff;
        border-radius: 8px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


def texto_padrao(dados: dict) -> str:
    opcoes_pagamento = {
        "Cheque único para 120 dias": ("X", " ", " "),
        "Cheques pré-datados": (" ", "X", " "),
        "Cartão de crédito": (" ", " ", "X"),
    }
    opcao1, opcao2, opcao3 = opcoes_pagamento[dados["modalidade_pagamento"]]

    garantias = {
        "Nota promissória": "X" if "Nota promissória" in dados["garantias"] else " ",
        "Avalista(s)": "X" if "Avalista(s)" in dados["garantias"] else " ",
        "Hipoteca": "X" if "Hipoteca" in dados["garantias"] else " ",
        "Penhor de animais": "X" if "Penhor de animais" in dados["garantias"] else " ",
        "Outra garantia": "X" if dados["outra_garantia"].strip() else " ",
    }

    return f"""CONTRATO DE CONFISSÃO DE DÍVIDA

Pelo presente instrumento particular, de um lado:

CREDOR: {dados["credor_nome"]}, inscrito(a) no CPF/CNPJ nº {dados["credor_doc"]}, com endereço à {dados["credor_endereco"]};

e, de outro lado:

DEVEDOR: {dados["devedor_nome"]}, inscrito(a) no CPF nº {dados["devedor_cpf"]}, portador(a) do RG nº {dados["devedor_rg"]}, residente e domiciliado(a) à {dados["devedor_endereco"]};

têm entre si justo e contratado o presente CONTRATO DE CONFISSÃO DE DÍVIDA, mediante as cláusulas e condições seguintes:

CLÁUSULA PRIMEIRA - DA ORIGEM DA DÍVIDA

O DEVEDOR reconhece e confessa ser legítimo devedor da quantia de R$ {dados["valor_total"]} ({dados["valor_extenso"]}), decorrente da aquisição de animais bovinos realizada em leilão promovido em {dados["data_leilao"]}, conforme lote(s) nº {dados["lotes"]}, adquiridos pelo DEVEDOR.

CLÁUSULA SEGUNDA - DO VALOR E FORMA DE PAGAMENTO

O valor total da dívida confessada é de R$ {dados["valor_total"]} ({dados["valor_extenso"]}), referente à aquisição de animais bovinos em leilão, cujo pagamento será realizado por uma das modalidades abaixo, escolhida pelo DEVEDOR no ato da assinatura deste instrumento:

({opcao1}) Opção 1 - Cheque único para 120 dias

Pagamento mediante a entrega de 01 (um) cheque pré-datado para vencimento em 120 (cento e vinte) dias contados da data da aquisição dos animais.

({opcao2}) Opção 2 - Pagamento por meio de cheques pré-datados

O valor total será quitado mediante a entrega de 04 (quatro) cheques pré-datados, correspondentes a parcelas iguais e sucessivas, com vencimentos em:

• 30 (trinta) dias;
• 60 (sessenta) dias;
• 90 (noventa) dias;
• 120 (cento e vinte) dias.

Cada parcela corresponderá ao valor de R$ {dados["valor_parcela"]}.

({opcao3}) Opção 3 - Pagamento mediante cartão de crédito

O valor total da aquisição será pago por meio de cartão de crédito, em até 04 (quatro) parcelas mensais, conforme aprovação da operadora do cartão e condições financeiras vigentes na data da operação.

Parágrafo Único. A modalidade de pagamento escolhida pelo DEVEDOR deverá ser assinalada acima e passará a integrar este contrato para todos os efeitos legais.

CLÁUSULA TERCEIRA - DA MORA

O não pagamento de qualquer parcela em seu vencimento implicará:

I - multa de 2% (dois por cento) sobre o valor da parcela em atraso;

II - juros de mora de 1% (um por cento) ao mês, calculados proporcionalmente aos dias de atraso;

III - correção monetária pelo índice oficial aplicável.

CLÁUSULA QUARTA - DO VENCIMENTO ANTECIPADO

O atraso superior a {dados["dias_atraso"]} dias no pagamento de qualquer parcela acarretará o vencimento antecipado das parcelas vincendas, tornando exigível o saldo devedor integral, independentemente de notificação judicial ou extrajudicial.

CLÁUSULA QUINTA - DA GARANTIA

Para garantia do cumprimento das obrigações assumidas, o DEVEDOR oferece como garantia:

({garantias["Nota promissória"]}) Nota promissória;
({garantias["Avalista(s)"]}) Avalista(s);
({garantias["Hipoteca"]}) Hipoteca;
({garantias["Penhor de animais"]}) Penhor de animais;
({garantias["Outra garantia"]}) Outra garantia: {dados["outra_garantia"] or "_____________________________"}.

CLÁUSULA SEXTA - DA CONFISSÃO IRREVOGÁVEL

O DEVEDOR declara reconhecer expressamente a existência, legitimidade, liquidez e exigibilidade da dívida descrita neste instrumento, renunciando a qualquer contestação futura quanto à sua origem e valor.

CLÁUSULA SÉTIMA - DO FORO

Fica eleito o foro da Comarca de {dados["foro"]}, com renúncia a qualquer outro, por mais privilegiado que seja, para dirimir eventuais controvérsias decorrentes deste contrato.

E, por estarem justos e contratados, firmam o presente instrumento em duas vias de igual teor e forma, juntamente com duas testemunhas.

{dados["municipio_assinatura"]}, {dados["dia_assinatura"]} de {dados["mes_assinatura"]} de {dados["ano_assinatura"]}.


________________________________________
CREDOR
CPF/CNPJ: {dados["credor_doc"]}


________________________________________
DEVEDOR
CPF: {dados["devedor_cpf"]}


TESTEMUNHAS

________________________________________
1. Nome: {dados["testemunha1_nome"]}
CPF: {dados["testemunha1_cpf"]}

________________________________________
2. Nome: {dados["testemunha2_nome"]}
CPF: {dados["testemunha2_cpf"]}
"""


st.title("Contrato de Confissão de Dívida")

with st.sidebar:
    st.markdown(
        """
        <div style="
            background:#0f253b;
            color:#ffffff;
            padding:1rem;
            border-radius:8px;
            margin-bottom:1rem;
        ">
            <div style="font-size:1.05rem;font-weight:700;">Confissão de dívida</div>
            <div style="font-size:.86rem;opacity:.88;margin-top:.2rem;">
                Aquisição de bovinos em leilão
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.header("Credor")
    credor_nome = st.text_input("Nome do credor ou leiloeira", "NOME DO VENDEDOR OU LEILOEIRA")
    credor_doc = st.text_input("CPF/CNPJ do credor", "XXXXXXXXXXXX")
    credor_endereco = st.text_area("Endereço do credor", "ENDEREÇO COMPLETO")

    st.header("Devedor")
    devedor_nome = st.text_input("Nome do comprador", "NOME DO COMPRADOR")
    devedor_cpf = st.text_input("CPF do devedor", "XXXXXXXXXXXX")
    devedor_rg = st.text_input("RG do devedor", "XXXXXXXXXXXX")
    devedor_endereco = st.text_area("Endereço do devedor", "ENDEREÇO COMPLETO")

    st.header("Dívida")
    valor_total = st.text_input("Valor total", "VALOR TOTAL")
    valor_extenso = st.text_input("Valor por extenso", "VALOR POR EXTENSO")
    data_leilao = st.date_input("Data do leilão", value=date.today(), format="DD/MM/YYYY")
    lotes = st.text_input("Número dos lotes", "NÚMERO DOS LOTES")

    st.header("Pagamento")
    modalidade_pagamento = st.radio(
        "Modalidade escolhida",
        ["Cheque único para 120 dias", "Cheques pré-datados", "Cartão de crédito"],
    )
    valor_parcela = st.text_input("Valor da parcela", "VALOR DA PARCELA")

    st.header("Garantia")
    garantias = st.multiselect(
        "Garantias oferecidas",
        ["Nota promissória", "Avalista(s)", "Hipoteca", "Penhor de animais"],
    )
    outra_garantia = st.text_input("Outra garantia", "")

    st.header("Foro e assinatura")
    dias_atraso = st.number_input("Dias para vencimento antecipado", min_value=1, value=30)
    foro = st.text_input("Comarca/UF", "MUNICÍPIO/UF")
    municipio_assinatura = st.text_input("Município da assinatura", "Município")
    data_assinatura = st.date_input("Data da assinatura", value=date.today(), format="DD/MM/YYYY")

    st.header("Testemunhas")
    testemunha1_nome = st.text_input("Nome da testemunha 1", "____________________________")
    testemunha1_cpf = st.text_input("CPF da testemunha 1", "_____________________________")
    testemunha2_nome = st.text_input("Nome da testemunha 2", "____________________________")
    testemunha2_cpf = st.text_input("CPF da testemunha 2", "_____________________________")

dados = {
    "credor_nome": credor_nome,
    "credor_doc": credor_doc,
    "credor_endereco": credor_endereco,
    "devedor_nome": devedor_nome,
    "devedor_cpf": devedor_cpf,
    "devedor_rg": devedor_rg,
    "devedor_endereco": devedor_endereco,
    "valor_total": valor_total,
    "valor_extenso": valor_extenso,
    "data_leilao": data_leilao.strftime("%d/%m/%Y"),
    "lotes": lotes,
    "modalidade_pagamento": modalidade_pagamento,
    "valor_parcela": valor_parcela,
    "garantias": garantias,
    "outra_garantia": outra_garantia,
    "dias_atraso": dias_atraso,
    "foro": foro,
    "municipio_assinatura": municipio_assinatura,
    "dia_assinatura": data_assinatura.strftime("%d"),
    "mes_assinatura": MESES[data_assinatura.month],
    "ano_assinatura": data_assinatura.strftime("%Y"),
    "testemunha1_nome": testemunha1_nome,
    "testemunha1_cpf": testemunha1_cpf,
    "testemunha2_nome": testemunha2_nome,
    "testemunha2_cpf": testemunha2_cpf,
}

texto_contrato = texto_padrao(dados)

pdf_bytes = gerar_contrato_pdf(texto_contrato, titulo="Contrato de Confissão de Dívida")
docx_bytes = gerar_contrato_docx(texto_contrato, titulo="Contrato de Confissão de Dívida")

_, col_preview, _ = st.columns([1, 2, 1])

with col_preview:
    st.subheader("Pre-visualizacao")
    mostrar_pdf_na_tela(pdf_bytes)

    download_col1, download_col2 = st.columns(2)
    with download_col1:
        st.download_button(
            "Baixar PDF",
            data=pdf_bytes,
            file_name="contrato_confissao_divida.pdf",
            mime="application/pdf",
            width="stretch",
        )
    with download_col2:
        st.download_button(
            "Baixar DOCX",
            data=docx_bytes,
            file_name="contrato_confissao_divida.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            width="stretch",
        )
