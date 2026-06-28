import re
from datetime import date, timedelta
from decimal import Decimal, InvalidOperation
from io import BytesIO

import streamlit as st
from PyPDF2 import PdfReader
from docx import Document
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

OPCOES_MODALIDADES_PAGAMENTO = [
    "Cheque bancario",
    "Cartao de credito",
]

OPCOES_DIAS_PARCELA = [30, 60, 90, 120]

MAPA_MODALIDADES_LEGADAS = {
    "Cheque ????nico para 120 dias": "Cheque bancario",
    "Cheques pr????-datados parcelados": "Cheque bancario",
    "Cart????o de cr????dito": "Cartao de credito",
    "Cheque ??nico para 120 dias": "Cheque bancario",
    "Cheques pr??-datados parcelados": "Cheque bancario",
    "Cart??o de cr??dito": "Cartao de credito",
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
    .stApp {
        background:
            radial-gradient(circle at top, #f2f6f3 0%, #fbfcfb 38%, #ffffff 100%);
    }
    .stApp .block-container {
        max-width: 1280px;
        margin-left: auto;
        margin-right: auto;
        padding-left: 2rem;
        padding-right: 2rem;
    }
    .stApp .block-container h2 {
        text-align: center;
    }
    [data-testid="stForm"] {
        border: 1px solid #d8dee9;
        border-radius: 14px;
        padding: 1.25rem;
        background: rgba(255, 255, 255, 0.94);
        box-shadow: 0 18px 40px rgba(36, 52, 44, 0.07);
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
    .titulo-contrato {
        display: flex;
        align-items: center;
        justify-content: center;
        gap: 0.8rem;
        margin: 0.4rem 0 1.5rem;
        color: #24342c;
    }
    .titulo-icone {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        width: 3.25rem;
        height: 3.25rem;
        border-radius: 999px;
        background: linear-gradient(135deg, #6f8f80 0%, #8fae9c 100%);
        box-shadow: 0 12px 24px rgba(111, 143, 128, 0.22);
        font-size: 1.45rem;
    }
    .titulo-texto {
        margin: 0;
        font-size: 2.4rem;
        line-height: 1.1;
        letter-spacing: -0.03em;
    }
    .bloco-separador {
        height: 1px;
        margin: 0.75rem 0 1rem;
        border: 0;
        background: #d8dee9;
    }
    .botao-credor-wrap {
        display: flex;
        justify-content: flex-end;
        align-items: center;
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
        return ""
    return num2words(numero, lang="pt_BR", to="currency")


def valor_parcela_calculado(valor_total: str, qtd_parcelas: int) -> str:
    numero = valor_monetario_para_decimal(valor_total)
    if numero is None or qtd_parcelas <= 0:
        return ""

    parcela = (numero / Decimal(qtd_parcelas)).quantize(Decimal("0.01"))
    texto = f"{parcela:,.2f}"
    return texto.replace(",", "X").replace(".", ",").replace("X", ".")


def valor_remanescente_calculado(valor_total: str, entrada: str) -> str:
    total = valor_monetario_para_decimal(valor_total)
    valor_entrada = valor_monetario_para_decimal(entrada)

    if total is None:
        return ""

    saldo = total - (valor_entrada or Decimal("0"))
    if saldo < 0:
        saldo = Decimal("0")

    saldo = saldo.quantize(Decimal("0.01"))
    texto = f"{saldo:,.2f}"
    return texto.replace(",", "X").replace(".", ",").replace("X", ".")


def decimal_para_texto_monetario(valor: Decimal) -> str:
    valor = valor.quantize(Decimal("0.01"))
    texto = f"{valor:,.2f}"
    return texto.replace(",", "X").replace(".", ",").replace("X", ".")


def atualizar_entrada() -> None:
    entrada = str(st.session_state.get("entrada", "")).strip()
    st.session_state["entrada"] = entrada or "0"


def atualizar_valor_cheque_divisao() -> None:
    valor = str(st.session_state.get("valor_cheque_divisao", "")).strip()
    st.session_state["valor_cheque_divisao"] = valor


def gerar_tabela_vencimentos(
    data_inicial: date, qtd_parcelas: int, intervalo_dias: int, valor_total: str
) -> list[dict[str, str]]:
    valor_parcela = valor_parcela_calculado(valor_total, qtd_parcelas)
    vencimentos = []

    for indice in range(1, qtd_parcelas + 1):
        vencimento = data_inicial + timedelta(days=intervalo_dias * indice)
        vencimentos.append(
            {
                "Parcela": str(indice),
                "Vencimento": formatar_data(vencimento),
                "Valor": valor_parcela,
            }
        )

    return vencimentos


def somar_valores_tabela(vencimentos: list[dict[str, str]]) -> str:
    total = Decimal("0")
    for item in vencimentos:
        valor = valor_monetario_para_decimal(item.get("Valor", ""))
        if valor is not None:
            total += valor
    return decimal_para_texto_monetario(total)


def numero_por_extenso(numero: int) -> str:
    return num2words(numero, lang="pt_BR")


def campo_ou_linha(valor: str, placeholder: str = "_____________________________") -> str:
    return " ".join(str(valor).split()) if str(valor).strip() else placeholder


def formatar_data(data_valor: date) -> str:
    return data_valor.strftime("%d/%m/%Y")


def nome_arquivo_seguro(texto: str) -> str:
    nome_limpo = re.sub(r'[\\/:*?"<>|]+', "", str(texto)).strip()
    nome_limpo = re.sub(r"\s+", "_", nome_limpo)
    return nome_limpo or "sem_comprador"


def extrair_texto_arquivo(nome_arquivo: str, conteudo: bytes) -> str:
    extensao = nome_arquivo.lower().rsplit(".", 1)[-1]

    if extensao == "docx":
        documento = Document(BytesIO(conteudo))
        return "\n".join(paragrafo.text for paragrafo in documento.paragraphs if paragrafo.text.strip())

    if extensao == "pdf":
        leitor = PdfReader(BytesIO(conteudo))
        return "\n".join((pagina.extract_text() or "") for pagina in leitor.pages)

    return ""


def normalizar_texto_importado(texto: str) -> str:
    return re.sub(r"\s+", " ", texto).strip()


def limpar_campo_importado(valor: str) -> str:
    texto = valor.strip(" []")
    if not texto:
        return ""
    if re.fullmatch(r"[_\-.Xx\s/]+", texto):
        return ""
    if texto.upper() in {"NOME DO VENDEDOR OU LEILOEIRA", "ENDEREÇO COMPLETO"}:
        return ""
    return texto


def extrair_partes_importadas(texto: str) -> dict[str, str]:
    texto_normalizado = normalizar_texto_importado(texto)

    padrao_credor = re.compile(
        r"CREDOR:\s*(?P<nome>.*?),\s*inscrito\(a\)\s*no\s*CPF/CNPJ\s*n[ºo]\s*(?P<doc>.*?),\s*com\s*endereç[oó]\s*à\s*(?P<endereco>.*?);",
        re.IGNORECASE,
    )
    padrao_devedor = re.compile(
        r"DEVEDOR:\s*(?P<nome>.*?),\s*inscrito\(a\)\s*no\s*CPF\s*n[ºo]\s*(?P<cpf>.*?),\s*portador\(a\)\s*do\s*RG\s*n[ºo]\s*(?P<rg>.*?),\s*residente\s*e\s*domiciliado\(a\)\s*à\s*(?P<endereco>.*?);",
        re.IGNORECASE,
    )

    dados: dict[str, str] = {}

    credor = padrao_credor.search(texto_normalizado)
    if credor:
        dados["credor_nome"] = limpar_campo_importado(credor.group("nome"))
        dados["credor_doc"] = limpar_campo_importado(credor.group("doc"))
        dados["credor_endereco"] = limpar_campo_importado(credor.group("endereco"))

    devedor = padrao_devedor.search(texto_normalizado)
    if devedor:
        dados["devedor_nome"] = limpar_campo_importado(devedor.group("nome"))
        dados["devedor_cpf"] = limpar_campo_importado(devedor.group("cpf"))
        dados["devedor_rg"] = limpar_campo_importado(devedor.group("rg"))
        dados["devedor_endereco"] = limpar_campo_importado(devedor.group("endereco"))

    dados = {chave: valor for chave, valor in dados.items() if valor}

    return dados


def inicializar_estado_formulario() -> None:
    defaults = {
        "credor_nome": "NOME DO VENDEDOR OU LEILOEIRA",
        "credor_doc": "XXXXXXXXXXXX",
        "credor_endereco": "ENDEREÇO COMPLETO",
        "devedor_nome": "",
        "devedor_cpf": "",
        "devedor_rg": "",
        "devedor_endereco": "",
        "valor_total": "",
        "qtd_parcelas": 4,
        "qtd_parcelas_cartao": 4,
        "entrada": "0",
        "valor_cheque_divisao": "",
        "modalidades_pagamento": ["Cheque bancario"],
        "valor_parcela": "",
        "dias_atraso": 30,
        "dias_atraso_cartao": 30,
        "foro": "São Francisco do Guaporé - RO",
        "municipio_assinatura": "São Francisco do Guaporé – RO",
        "cheque_unico_banco": "",
        "cheque_unico_agencia": "",
        "cheque_unico_conta": "",
        "cheque_unico_numero": "",
        "cartao_credito_banco": "",
        "cartao_credito_agencia": "",
        "cartao_credito_conta": "",
        "cartao_credito_numero": "",
        "parcela1_cheque": "",
        "parcela1_banco": "",
        "parcela2_cheque": "",
        "parcela2_banco": "",
        "parcela3_cheque": "",
        "parcela3_banco": "",
        "parcela4_cheque": "",
        "parcela4_banco": "",
        "testemunha1_nome": "",
        "testemunha1_cpf": "",
        "testemunha2_nome": "",
        "testemunha2_cpf": "",
    }

    for chave, valor in defaults.items():
        st.session_state.setdefault(chave, valor)

    modalidades = st.session_state.get("modalidades_pagamento", [])
    if not isinstance(modalidades, list):
        modalidades = [modalidades] if modalidades else []

    modalidades_normalizadas = []
    for modalidade in modalidades:
        modalidade_normalizada = MAPA_MODALIDADES_LEGADAS.get(modalidade, modalidade)
        if modalidade_normalizada in OPCOES_MODALIDADES_PAGAMENTO:
            modalidades_normalizadas.append(modalidade_normalizada)

    if not modalidades_normalizadas:
        modalidades_normalizadas = ["Cheque bancario"]

    st.session_state["modalidades_pagamento"] = modalidades_normalizadas


def limpar_dados_credor() -> None:
    st.session_state["credor_nome"] = ""
    st.session_state["credor_doc"] = ""
    st.session_state["credor_endereco"] = ""


def limpar_dados_devedor() -> None:
    st.session_state["devedor_nome"] = ""
    st.session_state["devedor_cpf"] = ""
    st.session_state["devedor_rg"] = ""
    st.session_state["devedor_endereco"] = ""


inicializar_estado_formulario()


def texto_padrao(dados: dict) -> str:
    def negrito(valor: str) -> str:
        return f"[[B]]{valor}[[/B]]"

    modalidades_pagamento = dados["modalidades_pagamento"]
    opcao1 = negrito("X") if "Cheque bancario" in modalidades_pagamento else " "
    opcao2 = negrito("X") if "Cartao de credito" in modalidades_pagamento else " "
    opcao3 = " "

    return f"""CONTRATO DE CONFISSÃO DE DÍVIDA

Pelo presente instrumento particular, de um lado:

CREDOR: {negrito(campo_ou_linha(dados["credor_nome"], "NOME DO VENDEDOR OU LEILOEIRA"))}, inscrito(a) no CPF/CNPJ nº {negrito(campo_ou_linha(dados["credor_doc"], "XXXXXXXXXXXX"))}, com endereço à {negrito(campo_ou_linha(dados["credor_endereco"], "ENDEREÇO COMPLETO"))};

e, de outro lado:

DEVEDOR: {negrito(campo_ou_linha(dados["devedor_nome"], "_____________________________________________________"))}, inscrito(a) no CPF nº {negrito(campo_ou_linha(dados["devedor_cpf"], "________________________"))}, portador(a) do RG nº {negrito(campo_ou_linha(dados["devedor_rg"], "_______________________"))}, residente e domiciliado(a) à {negrito(campo_ou_linha(dados["devedor_endereco"], "_________________________________________________________________________________________________________________________________________________________"))};

As partes acima qualificadas têm entre si justo e contratado o presente instrumento, que se regerá pelas cláusulas e condições seguintes:

CLÁUSULA PRIMEIRA – DA ORIGEM DA DÍVIDA

O DEVEDOR reconhece, confessa e assume ser legítimo devedor da quantia líquida, certa e exigível de R$ {negrito(campo_ou_linha(dados["valor_total"], "___________________"))} ({negrito(campo_ou_linha(dados["valor_extenso"], "_______________________________________________________"))}), decorrente da aquisição de animais bovinos realizada em leilão promovido em {negrito(dados["data_leilao"])}, conforme lote(s) nº {negrito(campo_ou_linha(dados["lotes"], "_______________________________"))}, adquiridos pelo DEVEDOR.

CLÁUSULA SEGUNDA - DO VALOR E FORMA DE PAGAMENTO

O valor total da dívida confessada é de R$ {negrito(campo_ou_linha(dados["valor_total"], "__________________"))} ({negrito(campo_ou_linha(dados["valor_extenso"], "________________________________________________________"))}), cujo pagamento será realizado por uma das modalidades abaixo, expressamente escolhida e assinalada pelo DEVEDOR no ato da assinatura deste instrumento:

[{opcao1}] Opcao 1 - Cheque bancario:

Pagamento mediante cheque bancario, conforme dados informados neste instrumento.

Dados do Cheque: Banco: {negrito(campo_ou_linha(dados["cheque_unico_banco"], "_______"))} | Agencia: {negrito(campo_ou_linha(dados["cheque_unico_agencia"], "___________"))} | Conta: {negrito(campo_ou_linha(dados["cheque_unico_conta"], "____________"))} | Cheque n: {negrito(campo_ou_linha(dados["cheque_unico_numero"], "______________"))}.

[{opcao2}] Opcao 2 - Pagamento mediante cartao de credito:

O valor total da aquisicao sera pago por meio de cartao de credito, em parcelas sucessivas de R$ {negrito(campo_ou_linha(dados["valor_parcela_cartao"], "VALOR DA PARCELA"))}, conforme aprovacao da operadora do cartao e condicoes financeiras vigentes na data da operacao.


Parágrafo primeiro. A modalidade de pagamento escolhida pelo DEVEDOR e assinalada acima integrará este contrato para todos os efeitos legais.

CLÁUSULA TERCEIRA – DA MORA E ENCARGOS INDENIZATÓRIOS

O não pagamento de qualquer parcela ou título nas datas de seus respectivos vencimentos constituirá o DEVEDOR em mora, independentemente de notificação judicial ou extrajudicial, incidindo sobre o valor do débito os seguintes encargos: I - Multa moratória e irredutível de 2% (dois por cento) sobre o valor da parcela em atraso; II - Juros de mora de 1% (um por cento) ao mês, calculados pro rata die (proporcionalmente aos dias de atraso); III - Correção monetária calculada com base na variação positiva do IGP-M/FGV (ou índice oficial que venha a substituí-lo), acumulada desde a data do vencimento até o efetivo pagamento.

Parágrafo Único. Caso o CREDOR precise recorrer a serviços advocatícios ou empresas de cobrança para o recebimento do crédito, o DEVEDOR responderá, além do principal e encargos, pelo pagamento das custas, despesas desembolsadas e honorários advocatícios, estes fixados em 10% (dez por cento) para cobrança extrajudicial e 20% (vinte por cento) em caso de ajuizamento de ação judicial.

CLÁUSULA QUARTA – DO VENCIMENTO ANTECIPADO

O atraso superior a {negrito(str(dados["dias_atraso"]))} ({negrito(dados["dias_atraso_extenso"])}) dias no pagamento de qualquer das parcelas pactuadas, ou a ocorrência de devolução por falta de fundos de qualquer dos cheques emitidos, acarretará o vencimento antecipado de todas as parcelas vincendas, tornando-se imediatamente exigível o saldo devedor integral, acrescido de todas as penalidades previstas na Cláusula Terceira, independentemente de prévia notificação ou aviso.

CLÁUSULA QUINTA– DA CONFISSÃO IRREVOGÁVEL E TÍTULO EXECUTIVO

O DEVEDOR declara reconhecer expressamente a existência, legitimidade, certeza, liquidez e exigibilidade da dívida descrita neste instrumento. Este contrato é firmado em caráter irrevogável e irretratável, constituindo-se em Título Executivo Extrajudicial, nos termos do Artigo 784, inciso III, do Código de Processo Civil brasileiro, apto a embasar Ação de Execução imediata.

CLÁUSULA SEXTA – DAS ASSINATURAS ELETRÔNICAS

As partes declaram e concordam que este contrato poderá ser assinado eletronicamente por meio de plataformas de assinatura digital, sendo as assinaturas consideradas válidas, íntegras e plenamente eficazes para todos os fins de direito, nos termos da Medida Provisória nº 2.200-2/2001 e da Lei nº 14.063/2020.

CLÁUSULA SÉTIMA – DO FORO

Fica eleito o foro da Comarca de {negrito(campo_ou_linha(dados["foro"], "São Francisco do Guaporé - RO"))}, com renúncia expressa a qualquer outro, por mais privilegiado que seja, para dirimir eventuais controvérsias decorrentes deste contrato.

E, por estarem assim justos e contratados, firmam o presente instrumento em 02 (duas) vias de igual teor e forma, na presença de 02 (duas) testemunhas abaixo assinadas.

{negrito(campo_ou_linha(dados["municipio_assinatura"], "São Francisco do Guaporé – RO"))}, {negrito(dados["dia_assinatura"])} de {negrito(dados["mes_assinatura"])} de {negrito(dados["ano_assinatura"])}.


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


st.markdown(
    """
    <div class="titulo-contrato">
        <span class="titulo-icone">📄</span>
        <h1 class="titulo-texto">Contrato de Confissão de Dívida</h1>
    </div>
    """,
    unsafe_allow_html=True,
)

contrato_importado = st.file_uploader(
    "Importar contrato anterior para copiar credor e devedor",
    type=["docx", "pdf"],
    help="O app tenta extrair automaticamente nome, documento e endereço do credor e do devedor.",
)

if contrato_importado is not None:
    texto_importado = extrair_texto_arquivo(contrato_importado.name, contrato_importado.getvalue())
    dados_importados = extrair_partes_importadas(texto_importado)

    if dados_importados:
        st.info("Arquivo carregado. Clique no botão abaixo para copiar os dados de credor e devedor.")
        if st.button("📥 Importar dados do arquivo", type="primary"):
            for chave, valor in dados_importados.items():
                st.session_state[chave] = valor
            st.success("Dados de credor e devedor importados do contrato anterior.")
            st.rerun()
    else:
        st.warning("Não foi possível localizar automaticamente os dados de credor e devedor nesse arquivo.")

data_leilao_padrao = date.today()

with st.container():
    st.subheader("Dados do contrato")

    with st.expander("Credor e devedor", expanded=False):
        col_credor, col_devedor = st.columns(2)
        with col_credor:
            st.markdown("### Credor")
            credor_nome = st.text_input("Nome do credor ou leiloeira", key="credor_nome")
            credor_doc = st.text_input("CPF/CNPJ do credor", key="credor_doc")
            credor_endereco = st.text_area("Endereco do credor", key="credor_endereco", height=95)

        with col_devedor:
            st.markdown("### Devedor")
            devedor_nome = st.text_input("Nome do comprador", key="devedor_nome")
            devedor_doc_col, devedor_rg_col = st.columns(2)
            with devedor_doc_col:
                devedor_cpf = st.text_input("CPF do devedor", key="devedor_cpf")
            with devedor_rg_col:
                devedor_rg = st.text_input("RG do devedor", key="devedor_rg")
            devedor_endereco = st.text_area("Endereco do devedor", key="devedor_endereco", height=95)

        botoes_col1, botoes_col2, botoes_col3 = st.columns([1, 1, 1])
        with botoes_col1:
            st.markdown('<div class="botao-credor-wrap">', unsafe_allow_html=True)
            st.button("Limpar credor", width="stretch", on_click=limpar_dados_credor)
            st.markdown("</div>", unsafe_allow_html=True)
        with botoes_col2:
            st.write("")
        with botoes_col3:
            st.button("Limpar devedor", width="stretch", on_click=limpar_dados_devedor)

    with st.expander("Financeiro", expanded=False):
        st.markdown("### Leil\u00e3o e d\u00edvida")
        col_data_leilao, col_divida, col_valor_extenso = st.columns(3)
        with col_data_leilao:
            data_leilao = st.date_input("Data inicial", value=data_leilao_padrao, format="DD/MM/YYYY")
        with col_divida:
            valor_total = st.text_input("D\u00edvida", key="valor_total")
        with col_valor_extenso:
            valor_extenso = st.text_input("Valor por extenso", valor_por_extenso(valor_total), disabled=True)

        lotes = st.text_input("N\u00famero dos lotes", "")

        st.markdown("### Pagamento")
        modalidades_pagamento = st.multiselect(
            "Modalidade escolhida",
            OPCOES_MODALIDADES_PAGAMENTO,
            key="modalidades_pagamento",
        )
        col_entrada, col_valor_atual, col_valor_remanescente_extenso = st.columns(3)
        with col_entrada:
            entrada = st.text_input("Com entrada? Informe o valor:", key="entrada", on_change=atualizar_entrada)
        valor_remanescente = valor_remanescente_calculado(valor_total, entrada)
        st.session_state["valor_atual_exibicao"] = valor_remanescente
        st.session_state["valor_remanescente_extenso_exibicao"] = valor_por_extenso(valor_remanescente)
        with col_valor_atual:
            st.text_input("Valor atual", disabled=True, key="valor_atual_exibicao")
        with col_valor_remanescente_extenso:
            st.text_input("Valor remanescente por extenso", disabled=True, key="valor_remanescente_extenso_exibicao")
        mostrar_formulario_cheque = "Cheque bancario" in modalidades_pagamento
        mostrar_formulario_cartao = "Cartao de credito" in modalidades_pagamento
        valor_cheque_base = valor_remanescente
        valor_cartao_base = valor_remanescente
        valor_divisao_informado = True

        if mostrar_formulario_cheque and mostrar_formulario_cartao:
            col_valor_cheque, col_valor_cartao = st.columns(2)
            with col_valor_cheque:
                valor_cheque_divisao = st.text_input(
                    "Do valor remanescente, quanto sera pago com cheque?",
                    key="valor_cheque_divisao",
                    on_change=atualizar_valor_cheque_divisao,
                )
            valor_divisao_informado = str(valor_cheque_divisao).strip() != ""
            valor_divisao_valido = valor_divisao_informado

            remanescente_decimal = valor_monetario_para_decimal(valor_remanescente) or Decimal("0")
            valor_cheque_decimal = valor_monetario_para_decimal(valor_cheque_divisao) or Decimal("0")
            if valor_cheque_decimal < 0:
                valor_cheque_decimal = Decimal("0")
            if valor_cheque_decimal > remanescente_decimal:
                valor_divisao_valido = False
                valor_cheque_decimal = remanescente_decimal

            valor_cheque_base = decimal_para_texto_monetario(valor_cheque_decimal)
            valor_cartao_base = decimal_para_texto_monetario(remanescente_decimal - valor_cheque_decimal)
            st.session_state["valor_cartao_divisao_exibicao"] = valor_cartao_base
            with col_valor_cartao:
                st.text_input("Valor para cartao", disabled=True, key="valor_cartao_divisao_exibicao")
            if not valor_divisao_informado:
                st.info("Informe o valor que sera pago com cheque e pressione Enter para liberar os formularios de cheque e cartao.")
            elif not valor_divisao_valido:
                st.warning("O valor informado para cheque nao pode ser maior que o saldo devedor inicial/remanescente.")
        else:
            valor_divisao_valido = True

        exibir_formulario_cheque = mostrar_formulario_cheque and (
            not mostrar_formulario_cartao or (valor_divisao_informado and valor_divisao_valido)
        )
        exibir_formulario_cartao = mostrar_formulario_cartao and (
            not mostrar_formulario_cheque or (valor_divisao_informado and valor_divisao_valido)
        )

        if exibir_formulario_cheque:
            st.markdown("### Para cheque")
            col_qtd_parcelas, col_dias_parcela = st.columns(2)
            with col_qtd_parcelas:
                qtd_parcelas = st.number_input("QTD parcelas (mensais)", min_value=1, step=1, key="qtd_parcelas")
            with col_dias_parcela:
                dias_atraso = st.selectbox(
                    "Dias para vencimento de cada parcela",
                    options=OPCOES_DIAS_PARCELA,
                    key="dias_atraso",
                )
            valor_parcela = valor_parcela_calculado(valor_cheque_base, int(qtd_parcelas))

            st.markdown("### Tabela de vencimentos")
            tabela_vencimentos = gerar_tabela_vencimentos(
                data_inicial=data_leilao,
                qtd_parcelas=int(qtd_parcelas),
                intervalo_dias=int(dias_atraso),
                valor_total=valor_cheque_base,
            )
            st.dataframe(tabela_vencimentos, hide_index=True, width="stretch")
            st.caption(f"Soma das parcelas: {somar_valores_tabela(tabela_vencimentos)}")

            st.markdown("### Dados bancarios")
            col_cheque1, col_cheque2, col_cheque3, col_cheque4 = st.columns(4)
            with col_cheque1:
                cheque_unico_banco = st.text_input("Banco", key="cheque_unico_banco")
            with col_cheque2:
                cheque_unico_agencia = st.text_input("Agencia", key="cheque_unico_agencia")
            with col_cheque3:
                cheque_unico_conta = st.text_input("Conta", key="cheque_unico_conta")
            with col_cheque4:
                cheque_unico_numero = st.text_input("Cheque n", key="cheque_unico_numero")
        else:
            qtd_parcelas = st.session_state.get("qtd_parcelas", 4)
            dias_atraso = st.session_state.get("dias_atraso", 30)
            valor_parcela = valor_parcela_calculado(valor_cheque_base, int(qtd_parcelas))
            cheque_unico_banco = st.session_state.get("cheque_unico_banco", "")
            cheque_unico_agencia = st.session_state.get("cheque_unico_agencia", "")
            cheque_unico_conta = st.session_state.get("cheque_unico_conta", "")
            cheque_unico_numero = st.session_state.get("cheque_unico_numero", "")

        if exibir_formulario_cheque and exibir_formulario_cartao:
            st.markdown('<div class="bloco-separador"></div>', unsafe_allow_html=True)

        if exibir_formulario_cartao:
            st.markdown("### Para Cartao de Credito")
            col_qtd_parcelas_cartao, col_dias_parcela_cartao = st.columns(2)
            with col_qtd_parcelas_cartao:
                qtd_parcelas_cartao = st.number_input("QTD parcelas (mensais)", min_value=1, step=1, key="qtd_parcelas_cartao")
            with col_dias_parcela_cartao:
                dias_atraso_cartao = st.selectbox(
                    "Dias para vencimento de cada parcela",
                    options=OPCOES_DIAS_PARCELA,
                    key="dias_atraso_cartao",
                )
            valor_parcela_cartao = valor_parcela_calculado(valor_cartao_base, int(qtd_parcelas_cartao))

            st.markdown("### Tabela de vencimentos")
            tabela_vencimentos_cartao = gerar_tabela_vencimentos(
                data_inicial=data_leilao,
                qtd_parcelas=int(qtd_parcelas_cartao),
                intervalo_dias=int(dias_atraso_cartao),
                valor_total=valor_cartao_base,
            )
            st.dataframe(tabela_vencimentos_cartao, hide_index=True, width="stretch")
            st.caption(f"Soma das parcelas: {somar_valores_tabela(tabela_vencimentos_cartao)}")

            col_cartao1, col_cartao2, col_cartao3, col_cartao4 = st.columns(4)
            with col_cartao1:
                cartao_credito_banco = st.text_input("Banco", key="cartao_credito_banco")
            with col_cartao2:
                cartao_credito_agencia = st.text_input("Agencia", key="cartao_credito_agencia")
            with col_cartao3:
                cartao_credito_conta = st.text_input("Conta", key="cartao_credito_conta")
            with col_cartao4:
                cartao_credito_numero = st.text_input("Numero", key="cartao_credito_numero")
        else:
            qtd_parcelas_cartao = st.session_state.get("qtd_parcelas_cartao", 4)
            dias_atraso_cartao = st.session_state.get("dias_atraso_cartao", 30)
            valor_parcela_cartao = valor_parcela_calculado(valor_cartao_base, int(qtd_parcelas_cartao))
            cartao_credito_banco = st.session_state.get("cartao_credito_banco", "")
            cartao_credito_agencia = st.session_state.get("cartao_credito_agencia", "")
            cartao_credito_conta = st.session_state.get("cartao_credito_conta", "")
            cartao_credito_numero = st.session_state.get("cartao_credito_numero", "")

    with st.expander("Informacoes finais", expanded=False):
        st.markdown("### Foro e assinatura")
        foro = st.text_input("Comarca/UF", key="foro")
        municipio_assinatura = st.text_input("Municipio da assinatura", key="municipio_assinatura")
        data_assinatura = st.date_input("Data da assinatura", value=date.today(), format="DD/MM/YYYY")

        col_testemunha1, col_testemunha2 = st.columns(2)
        with col_testemunha1:
            st.markdown("### Testemunha 1")
            testemunha1_nome = st.text_input("Nome da testemunha 1", key="testemunha1_nome")
            testemunha1_cpf = st.text_input("CPF da testemunha 1", key="testemunha1_cpf")

        with col_testemunha2:
            st.markdown("### Testemunha 2")
            testemunha2_nome = st.text_input("Nome da testemunha 2", key="testemunha2_nome")
            testemunha2_cpf = st.text_input("CPF da testemunha 2", key="testemunha2_cpf")

    atualizar_preview = st.button("🔄 Atualizar pré-visualização", width="stretch")

parcela_datas = [
    data_leilao + timedelta(days=int(dias_atraso) * indice)
    for indice in range(1, 5)
]
parcela1_data, parcela2_data, parcela3_data, parcela4_data = parcela_datas
parcela1_cheque = st.session_state.get("parcela1_cheque", "")
parcela1_banco = st.session_state.get("parcela1_banco", "")
parcela2_cheque = st.session_state.get("parcela2_cheque", "")
parcela2_banco = st.session_state.get("parcela2_banco", "")
parcela3_cheque = st.session_state.get("parcela3_cheque", "")
parcela3_banco = st.session_state.get("parcela3_banco", "")
parcela4_cheque = st.session_state.get("parcela4_cheque", "")
parcela4_banco = st.session_state.get("parcela4_banco", "")

dados = {
    "credor_nome": credor_nome,
    "credor_doc": credor_doc,
    "credor_endereco": credor_endereco,
    "devedor_nome": devedor_nome,
    "devedor_cpf": devedor_cpf,
    "devedor_rg": devedor_rg,
    "devedor_endereco": devedor_endereco,
    "valor_total": valor_total,
    "entrada": entrada,
    "valor_remanescente": valor_remanescente,
    "valor_cheque_divisao": st.session_state.get("valor_cheque_divisao", "0"),
    "valor_cheque_base": valor_cheque_base,
    "valor_cartao_base": valor_cartao_base,
    "valor_extenso": valor_extenso,
    "data_leilao": formatar_data(data_leilao),
    "lotes": lotes,
    "modalidades_pagamento": modalidades_pagamento,
    "valor_parcela": valor_parcela,
    "dias_atraso": dias_atraso,
    "dias_atraso_extenso": numero_por_extenso(dias_atraso),
    "qtd_parcelas_cartao": qtd_parcelas_cartao,
    "valor_parcela_cartao": valor_parcela_cartao,
    "dias_atraso_cartao": dias_atraso_cartao,
    "cheque_unico_banco": cheque_unico_banco,
    "cheque_unico_agencia": cheque_unico_agencia,
    "cheque_unico_conta": cheque_unico_conta,
    "cheque_unico_numero": cheque_unico_numero,
    "cartao_credito_banco": cartao_credito_banco,
    "cartao_credito_agencia": cartao_credito_agencia,
    "cartao_credito_conta": cartao_credito_conta,
    "cartao_credito_numero": cartao_credito_numero,
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
nome_comprador_arquivo = nome_arquivo_seguro(devedor_nome)

pdf_bytes = gerar_contrato_pdf(texto_contrato, titulo="Contrato de Confissão de Dívida")
docx_bytes = gerar_contrato_docx(texto_contrato, titulo="Contrato de Confissão de Dívida")

_, col_preview, _ = st.columns([1, 2, 1])

with col_preview:
    st.subheader("Pré-visualização")
    mostrar_pdf_na_tela(pdf_bytes)

    download_col1, download_col2 = st.columns(2)
    with download_col1:
        st.download_button(
            "⬇️ Baixar PDF",
            data=pdf_bytes,
            file_name=f"contrato_confissao_divida_{nome_comprador_arquivo}.pdf",
            mime="application/pdf",
            width="stretch",
        )
    with download_col2:
        st.download_button(
            "⬇️ Baixar DOCX",
            data=docx_bytes,
            file_name=f"contrato_confissao_divida_{nome_comprador_arquivo}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            width="stretch",
        )
