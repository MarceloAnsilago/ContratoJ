from datetime import date, timedelta
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


def numero_por_extenso(numero: int) -> str:
    return num2words(numero, lang="pt_BR")


def campo_ou_linha(valor: str, placeholder: str = "_____________________________") -> str:
    return " ".join(str(valor).split()) if str(valor).strip() else placeholder


def formatar_data(data_valor: date) -> str:
    return data_valor.strftime("%d/%m/%Y")


def texto_padrao(dados: dict) -> str:
    def negrito(valor: str) -> str:
        return f"[[B]]{valor}[[/B]]"

    modalidades_pagamento = dados["modalidades_pagamento"]
    opcao1 = negrito("X") if "Cheque único para 120 dias" in modalidades_pagamento else " "
    opcao2 = negrito("X") if "Cheques pré-datados parcelados" in modalidades_pagamento else " "
    opcao3 = negrito("X") if "Cartão de crédito" in modalidades_pagamento else " "

    return f"""CONTRATO DE CONFISSÃO DE DÍVIDA

Pelo presente instrumento particular, de um lado:

CREDOR: {negrito(campo_ou_linha(dados["credor_nome"], "NOME DO VENDEDOR OU LEILOEIRA"))}, inscrito(a) no CPF/CNPJ nº {negrito(campo_ou_linha(dados["credor_doc"], "XXXXXXXXXXXX"))}, com endereço à {negrito(campo_ou_linha(dados["credor_endereco"], "ENDEREÇO COMPLETO"))};

e, de outro lado:

DEVEDOR: {negrito(campo_ou_linha(dados["devedor_nome"]))}, inscrito(a) no CPF nº {negrito(campo_ou_linha(dados["devedor_cpf"]))}, portador(a) do RG nº {negrito(campo_ou_linha(dados["devedor_rg"]))}, residente e domiciliado(a) à {negrito(campo_ou_linha(dados["devedor_endereco"]))};

As partes acima qualificadas têm entre si justo e contratado o presente instrumento, que se regerá pelas cláusulas e condições seguintes:

CLÁUSULA PRIMEIRA - DA ORIGEM DA DÍVIDA

O DEVEDOR reconhece, confessa e assume ser legítimo devedor da quantia líquida, certa e exigível de R$ {negrito(campo_ou_linha(dados["valor_total"], "___________________"))} ({negrito(campo_ou_linha(dados["valor_extenso"], "_______________________________________________________"))}), decorrente da aquisição de animais bovinos realizada em leilão promovido em {negrito(dados["data_leilao"])}, conforme lote(s) nº {negrito(campo_ou_linha(dados["lotes"], "_______________________________"))}, adquiridos pelo DEVEDOR.

CLÁUSULA SEGUNDA - DO VALOR E FORMA DE PAGAMENTO

O valor total da dívida confessada é de R$ {negrito(campo_ou_linha(dados["valor_total"], "__________________"))} ({negrito(campo_ou_linha(dados["valor_extenso"], "________________________________________________________"))}), cujo pagamento será realizado por uma das modalidades abaixo, expressamente escolhida e assinalada pelo DEVEDOR no ato da assinatura deste instrumento:

[{opcao1}] Opção 1 - Cheque único para 120 dias:

Pagamento mediante a entrega de 01 (um) cheque pré-datado para vencimento em 120 (cento e vinte) dias contados da data da aquisição dos animais.

Dados do Cheque: Banco: {negrito(campo_ou_linha(dados["cheque_unico_banco"], "_______"))} | Agência: {negrito(campo_ou_linha(dados["cheque_unico_agencia"], "___________"))} | Conta: {negrito(campo_ou_linha(dados["cheque_unico_conta"], "____________"))} | Cheque nº: {negrito(campo_ou_linha(dados["cheque_unico_numero"], "______________"))}.

[{opcao2}] Opção 2 - Pagamento por meio de cheques pré-datados parcelados

O valor total será quitado mediante a entrega de 04 (quatro) cheques pré-datados, correspondentes a parcelas iguais e sucessivas de R$ {negrito(campo_ou_linha(dados["valor_parcela"], "VALOR DA PARCELA"))}, com os seguintes vencimentos:

{negrito(dados["parcela1_data"])} (30 dias) - Cheque nº: {negrito(campo_ou_linha(dados["parcela1_cheque"], "_______________"))} | Banco: {negrito(campo_ou_linha(dados["parcela1_banco"], "_________"))}
{negrito(dados["parcela2_data"])} (60 dias) - Cheque nº: {negrito(campo_ou_linha(dados["parcela2_cheque"], "_______________"))} | Banco: {negrito(campo_ou_linha(dados["parcela2_banco"], "_________"))}
{negrito(dados["parcela3_data"])} (90 dias) - Cheque nº: {negrito(campo_ou_linha(dados["parcela3_cheque"], "_______________"))} | Banco: {negrito(campo_ou_linha(dados["parcela3_banco"], "_________"))}
{negrito(dados["parcela4_data"])} (120 dias) - Cheque nº: {negrito(campo_ou_linha(dados["parcela4_cheque"], "_______________"))} | Banco: {negrito(campo_ou_linha(dados["parcela4_banco"], "_________"))}

[{opcao3}] Opção 3 - Pagamento mediante cartão de crédito: O valor total da aquisição será pago por meio de cartão de crédito, em até 04 (quatro) parcelas mensais, conforme aprovação da operadora do cartão e condições financeiras vigentes na data da operação.

Parágrafo primeiro. A modalidade de pagamento escolhida pelo DEVEDOR e assinalada acima integrará este contrato para todos os efeitos legais.

CLÁUSULA TERCEIRA - DA MORA E ENCARGOS INDENIZATÓRIOS

O não pagamento de qualquer parcela ou título nas datas de seus respectivos vencimentos constituirá o DEVEDOR em mora, independentemente de notificação judicial ou extrajudicial, incidindo sobre o valor do débito os seguintes encargos: I - Multa moratória e irredutível de 2% (dois por cento) sobre o valor da parcela em atraso; II - Juros de mora de 1% (um por cento) ao mês, calculados pro rata die (proporcionalmente aos dias de atraso); III - Correção monetária calculada com base na variação positiva do IGP-M/FGV (ou índice oficial que venha a substituí-lo), acumulada desde a data do vencimento até o efetivo pagamento.

Parágrafo Único. Caso o CREDOR precise recorrer a serviços advocatícios ou empresas de cobrança para o recebimento do crédito, o DEVEDOR responderá, além do principal e encargos, pelo pagamento das custas, despesas desembolsadas e honorários advocatícios, estes fixados em 10% (dez por cento) para cobrança extrajudicial e 20% (vinte por cento) em caso de ajuizamento de ação judicial.

CLÁUSULA QUARTA - DO VENCIMENTO ANTECIPADO

O atraso superior a {negrito(str(dados["dias_atraso"]))} ({negrito(dados["dias_atraso_extenso"])}) dias no pagamento de qualquer das parcelas pactuadas, ou a ocorrência de devolução por falta de fundos de qualquer dos cheques emitidos, acarretará o vencimento antecipado de todas as parcelas vincendas, tornando-se imediatamente exigível o saldo devedor integral, acrescido de todas as penalidades previstas na Cláusula Terceira, independentemente de prévia notificação ou aviso.

CLÁUSULA QUINTA - DA CONFISSÃO IRREVOGÁVEL E TÍTULO EXECUTIVO

O DEVEDOR declara reconhecer expressamente a existência, legitimidade, certeza, liquidez e exigibilidade da dívida descrita neste instrumento. Este contrato é firmado em caráter irrevogável e irretratável, constituindo-se em Título Executivo Extrajudicial, nos termos do Artigo 784, inciso III, do Código de Processo Civil brasileiro, apto a embasar Ação de Execução imediata.

CLÁUSULA SEXTA - DAS ASSINATURAS ELETRÔNICAS

As partes declaram e concordam que este contrato poderá ser assinado eletronicamente por meio de plataformas de assinatura digital, sendo as assinaturas consideradas válidas, íntegras e plenamente eficazes para todos os fins de direito, nos termos da Medida Provisória nº 2.200-2/2001 e da Lei nº 14.063/2020.

CLÁUSULA SÉTIMA - DO FORO

Fica eleito o foro da Comarca de {negrito(campo_ou_linha(dados["foro"], "São Francisco do Guaporé - RO"))}, com renúncia expressa a qualquer outro, por mais privilegiado que seja, para dirimir eventuais controvérsias decorrentes deste contrato.

E, por estarem assim justos e contratados, firmam o presente instrumento em 02 (duas) vias de igual teor e forma, na presença de 02 (duas) testemunhas abaixo assinadas.

{negrito(campo_ou_linha(dados["municipio_assinatura"], "São Francisco do Guaporé - RO"))}, {negrito(dados["dia_assinatura"])} de {negrito(dados["mes_assinatura"])} de {negrito(dados["ano_assinatura"])}.


______________________________________________
CREDOR


______________________________________________
DEVEDOR


TESTEMUNHAS:

________________________________________
Nome: {negrito(campo_ou_linha(dados["testemunha1_nome"], ""))}
CPF: {negrito(campo_ou_linha(dados["testemunha1_cpf"], ""))}

________________________________________
Nome: {negrito(campo_ou_linha(dados["testemunha2_nome"], ""))}
CPF: {negrito(campo_ou_linha(dados["testemunha2_cpf"], ""))}
"""


st.title("Contrato de Confissão de Dívida")

data_leilao_padrao = date.today()

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
        devedor_nome = st.text_input("Nome do comprador", "")
        devedor_doc_col, devedor_rg_col = st.columns(2)
        with devedor_doc_col:
            devedor_cpf = st.text_input("CPF do devedor", "")
        with devedor_rg_col:
            devedor_rg = st.text_input("RG do devedor", "")
        devedor_endereco = st.text_area("Endereço do devedor", "", height=95)

    col_divida, col_pagamento, col_assinatura = st.columns(3)
    with col_divida:
        st.markdown("### Dívida")
        valor_total = st.text_input("Valor total", "")
        valor_extenso = st.text_input("Valor por extenso", valor_por_extenso(valor_total), disabled=True)
        data_leilao = st.date_input("Data do leilão", value=data_leilao_padrao, format="DD/MM/YYYY")
        lotes = st.text_input("Número dos lotes", "")

    with col_pagamento:
        st.markdown("### Pagamento")
        modalidades_pagamento = st.multiselect(
            "Modalidade escolhida",
            ["Cheque único para 120 dias", "Cheques pré-datados parcelados", "Cartão de crédito"],
            default=["Cheque único para 120 dias"],
        )
        valor_parcela = st.text_input("Valor da parcela", "")
        dias_atraso = st.number_input("Dias para vencimento antecipado", min_value=1, value=30)

    with col_assinatura:
        st.markdown("### Foro e assinatura")
        foro = st.text_input("Comarca/UF", "São Francisco do Guaporé - RO")
        municipio_assinatura = st.text_input("Município da assinatura", "São Francisco do Guaporé - RO")
        data_assinatura = st.date_input("Data da assinatura", value=date.today(), format="DD/MM/YYYY")

    st.markdown("### Cheque único para 120 dias")
    col_cheque1, col_cheque2, col_cheque3, col_cheque4 = st.columns(4)
    with col_cheque1:
        cheque_unico_banco = st.text_input("Banco", "")
    with col_cheque2:
        cheque_unico_agencia = st.text_input("Agência", "")
    with col_cheque3:
        cheque_unico_conta = st.text_input("Conta", "")
    with col_cheque4:
        cheque_unico_numero = st.text_input("Cheque nº", "")

    st.markdown("### Cheques pré-datados parcelados")
    parcela1_padrao = data_leilao + timedelta(days=30)
    parcela2_padrao = data_leilao + timedelta(days=60)
    parcela3_padrao = data_leilao + timedelta(days=90)
    parcela4_padrao = data_leilao + timedelta(days=120)

    col_p1, col_p2 = st.columns(2)
    with col_p1:
        parcela1_data = st.date_input("Vencimento parcela 1", value=parcela1_padrao, format="DD/MM/YYYY")
        parcela1_cheque = st.text_input("Cheque nº parcela 1", "")
        parcela1_banco = st.text_input("Banco parcela 1", "")
        parcela2_data = st.date_input("Vencimento parcela 2", value=parcela2_padrao, format="DD/MM/YYYY")
        parcela2_cheque = st.text_input("Cheque nº parcela 2", "")
        parcela2_banco = st.text_input("Banco parcela 2", "")
    with col_p2:
        parcela3_data = st.date_input("Vencimento parcela 3", value=parcela3_padrao, format="DD/MM/YYYY")
        parcela3_cheque = st.text_input("Cheque nº parcela 3", "")
        parcela3_banco = st.text_input("Banco parcela 3", "")
        parcela4_data = st.date_input("Vencimento parcela 4", value=parcela4_padrao, format="DD/MM/YYYY")
        parcela4_cheque = st.text_input("Cheque nº parcela 4", "")
        parcela4_banco = st.text_input("Banco parcela 4", "")

    col_testemunha1, col_testemunha2 = st.columns(2)
    with col_testemunha1:
        st.markdown("### Testemunha 1")
        testemunha1_nome = st.text_input("Nome da testemunha 1", "")
        testemunha1_cpf = st.text_input("CPF da testemunha 1", "")

    with col_testemunha2:
        st.markdown("### Testemunha 2")
        testemunha2_nome = st.text_input("Nome da testemunha 2", "")
        testemunha2_cpf = st.text_input("CPF da testemunha 2", "")

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
    "data_leilao": formatar_data(data_leilao),
    "lotes": lotes,
    "modalidades_pagamento": modalidades_pagamento,
    "valor_parcela": valor_parcela,
    "dias_atraso": dias_atraso,
    "dias_atraso_extenso": numero_por_extenso(dias_atraso),
    "cheque_unico_banco": cheque_unico_banco,
    "cheque_unico_agencia": cheque_unico_agencia,
    "cheque_unico_conta": cheque_unico_conta,
    "cheque_unico_numero": cheque_unico_numero,
    "parcela1_data": formatar_data(parcela1_data),
    "parcela1_cheque": parcela1_cheque,
    "parcela1_banco": parcela1_banco,
    "parcela2_data": formatar_data(parcela2_data),
    "parcela2_cheque": parcela2_cheque,
    "parcela2_banco": parcela2_banco,
    "parcela3_data": formatar_data(parcela3_data),
    "parcela3_cheque": parcela3_cheque,
    "parcela3_banco": parcela3_banco,
    "parcela4_data": formatar_data(parcela4_data),
    "parcela4_cheque": parcela4_cheque,
    "parcela4_banco": parcela4_banco,
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
