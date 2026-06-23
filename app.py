from datetime import date
from decimal import Decimal, InvalidOperation

import streamlit as st
from num2words import num2words

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
    initial_sidebar_state="collapsed",
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
    [data-testid="stForm"] {
        border: 1px solid #d8dee9;
        border-radius: 8px;
        padding: 1.25rem;
        background: #ffffff;
    }
    [data-testid="stForm"] h3 {
        margin-top: 0.2rem;
        padding-bottom: 0.25rem;
        border-bottom: 1px solid #edf0f5;
        font-size: 1.05rem;
    }
    [data-testid="stForm"] label {
        font-weight: 600;
    }
    [data-testid="stForm"] input,
    [data-testid="stForm"] textarea {
        border-radius: 6px;
    }
    .pdf-preview-frame {
        background: #eef1f5;
        border: 1px solid #c8d0dc;
        border-radius: 8px;
        box-shadow: 0 12px 28px rgba(15, 23, 42, 0.16);
        margin: 0.5rem auto 0.2rem;
        padding: 1rem;
    }
    .pdf-preview-frame img {
        display: block;
        width: 100%;
        border: 1px solid #d7dce5;
        box-shadow: 0 2px 8px rgba(15, 23, 42, 0.1);
    }
    .pdf-preview-caption {
        color: #5f6b7a;
        font-size: 0.85rem;
        margin-bottom: 1rem;
        text-align: center;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


def valor_monetario_para_decimal(valor: str) -> Decimal | None:
    texto = str(valor).strip()
    if not texto:
        return None

    normalizado = "".join(caractere for caractere in texto if caractere.isdigit() or caractere in ",.-")
    if not normalizado:
        return None

    if "," in normalizado:
        normalizado = normalizado.replace(".", "").replace(",", ".")
    elif normalizado.count(".") > 1:
        partes = normalizado.split(".")
        normalizado = "".join(partes[:-1]) + "." + partes[-1]
    elif "." in normalizado and len(normalizado.rsplit(".", 1)[1]) == 3:
        normalizado = normalizado.replace(".", "")

    try:
        return Decimal(normalizado)
    except InvalidOperation:
        return None


def valor_por_extenso(valor: str) -> str:
    numero = valor_monetario_para_decimal(valor)
    if numero is None:
        return "VALOR POR EXTENSO"

    return num2words(numero, lang="pt_BR", to="currency")


def texto_padrao(dados: dict) -> str:
    def uma_linha(valor: str) -> str:
        return " ".join(str(valor).split())

    modalidades_pagamento = dados["modalidades_pagamento"]
    opcao1 = "X" if "Cheque único para 120 dias" in modalidades_pagamento else " "
    opcao2 = "X" if "Cheques pré-datados" in modalidades_pagamento else " "
    opcao3 = "X" if "Cartão de crédito" in modalidades_pagamento else " "

    garantias = {
        "Nota promissória": "X" if "Nota promissória" in dados["garantias"] else " ",
        "Avalista(s)": "X" if "Avalista(s)" in dados["garantias"] else " ",
        "Hipoteca": "X" if "Hipoteca" in dados["garantias"] else " ",
        "Penhor de animais": "X" if "Penhor de animais" in dados["garantias"] else " ",
        "Outra garantia": "X" if dados["outra_garantia"].strip() else " ",
    }

    return f"""CONTRATO DE CONFISSÃO DE DÍVIDA

Pelo presente instrumento particular, de um lado:

CREDOR: {dados["credor_nome"]}, inscrito(a) no CPF/CNPJ nº {dados["credor_doc"]}, com endereço à {uma_linha(dados["credor_endereco"])};

e, de outro lado:

DEVEDOR: {dados["devedor_nome"]}, inscrito(a) no CPF nº {dados["devedor_cpf"]}, portador(a) do RG nº {dados["devedor_rg"]}, residente e domiciliado(a) à {uma_linha(dados["devedor_endereco"])};

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

with st.form("formulario_contrato"):
    st.subheader("Dados do contrato")

    col_credor, col_devedor = st.columns(2)
    with col_credor:
        st.markdown("### Credor")
        credor_nome = st.text_input("Nome do credor ou leiloeira", "NOME DO VENDEDOR OU LEILOEIRA")
        credor_doc = st.text_input("CPF/CNPJ do credor", "XXXXXXXXXXXX")
        credor_endereco = st.text_area("Endereço do credor", "ENDEREÇO COMPLETO", height=95)

    with col_devedor:
        st.markdown("### Devedor")
        devedor_nome = st.text_input("Nome do comprador", "NOME DO COMPRADOR")
        devedor_doc_col, devedor_rg_col = st.columns(2)
        with devedor_doc_col:
            devedor_cpf = st.text_input("CPF do devedor", "XXXXXXXXXXXX")
        with devedor_rg_col:
            devedor_rg = st.text_input("RG do devedor", "XXXXXXXXXXXX")
        devedor_endereco = st.text_area("Endereço do devedor", "ENDEREÇO COMPLETO", height=95)

    col_divida, col_pagamento, col_garantia = st.columns(3)
    with col_divida:
        st.markdown("### Dívida")
        valor_total = st.text_input("Valor total", "VALOR TOTAL")
        valor_extenso = st.text_input("Valor por extenso", valor_por_extenso(valor_total), disabled=True)
        data_leilao = st.date_input("Data do leilão", value=date.today(), format="DD/MM/YYYY")
        lotes = st.text_input("Número dos lotes", "NÚMERO DOS LOTES")

    with col_pagamento:
        st.markdown("### Pagamento")
        modalidades_pagamento = st.multiselect(
            "Modalidade escolhida",
            ["Cheque único para 120 dias", "Cheques pré-datados", "Cartão de crédito"],
            default=["Cheque único para 120 dias"],
        )
        valor_parcela = st.text_input("Valor da parcela", "VALOR DA PARCELA")
        dias_atraso = st.number_input("Dias para vencimento antecipado", min_value=1, value=30)

    with col_garantia:
        st.markdown("### Garantia")
        garantias = st.multiselect(
            "Garantias oferecidas",
            ["Nota promissória", "Avalista(s)", "Hipoteca", "Penhor de animais"],
        )
        outra_garantia = st.text_input("Outra garantia", "")

    col_foro, col_testemunha1, col_testemunha2 = st.columns(3)
    with col_foro:
        st.markdown("### Foro e assinatura")
        foro = st.text_input("Comarca/UF", "MUNICÍPIO/UF")
        municipio_assinatura = st.text_input("Município da assinatura", "Município")
        data_assinatura = st.date_input("Data da assinatura", value=date.today(), format="DD/MM/YYYY")

    with col_testemunha1:
        st.markdown("### Testemunha 1")
        testemunha1_nome = st.text_input("Nome da testemunha 1", "____________________________")
        testemunha1_cpf = st.text_input("CPF da testemunha 1", "_____________________________")

    with col_testemunha2:
        st.markdown("### Testemunha 2")
        testemunha2_nome = st.text_input("Nome da testemunha 2", "____________________________")
        testemunha2_cpf = st.text_input("CPF da testemunha 2", "_____________________________")

    st.form_submit_button("Atualizar pré-visualização", width="stretch")

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
    "modalidades_pagamento": modalidades_pagamento,
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
    st.subheader("Pré-visualização")
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
