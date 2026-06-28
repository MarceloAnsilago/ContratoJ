import re
import textwrap
import unicodedata
from datetime import date, timedelta
from decimal import Decimal, InvalidOperation
from io import BytesIO

import pandas as pd
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

OPCOES_DIAS_PARCELA = [30, 60, 90, 120, 150, 180]

MAPA_MODALIDADES_LEGADAS = {
    "Cheque ??nico para 120 dias": "Cheque bancario",
    "Cheques pr??-datados parcelados": "Cheque bancario",
    "Cart??o de cr??dito": "Cartao de credito",
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
    .streamlit-expanderContent,
    [data-testid="stExpander"] details > div,
    [data-testid="stExpander"] [data-testid="stVerticalBlock"] {
        background: #ffffff;
    }
    [data-testid="stExpander"] {
        background: #ffffff;
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


def campo_numerico_tem_letras(valor: str) -> bool:
    return bool(re.search(r"[A-Za-zÀ-ÿ]", str(valor or "")))


def exibir_warning_campo_numerico(valor: str, label: str) -> None:
    if campo_numerico_tem_letras(valor):
        st.warning(f"O campo {label.lower()} nao aceita letras.")


def sincronizar_valores_pagamento() -> None:
    valor_total = str(st.session_state.get("valor_total", "")).strip()
    entrada = str(st.session_state.get("entrada", "")).strip() or "0"

    valor_remanescente = valor_remanescente_calculado(valor_total, entrada)
    st.session_state["valor_atual_exibicao"] = valor_remanescente
    st.session_state["valor_remanescente_extenso_exibicao"] = valor_por_extenso(valor_remanescente)

    modalidades = st.session_state.get("modalidades_pagamento", [])
    if not isinstance(modalidades, list):
        modalidades = [modalidades] if modalidades else []

    remanescente_decimal = valor_monetario_para_decimal(valor_remanescente) or Decimal("0")
    valor_cheque_divisao = str(st.session_state.get("valor_cheque_divisao", "")).strip()

    if "Cheque bancario" in modalidades and "Cartao de credito" in modalidades:
        valor_cheque_decimal = valor_monetario_para_decimal(valor_cheque_divisao)
        if valor_cheque_decimal is None:
            st.session_state["valor_cartao_divisao_exibicao"] = decimal_para_texto_monetario(remanescente_decimal)
            return

        if valor_cheque_decimal < 0:
            valor_cheque_decimal = Decimal("0")

        if valor_cheque_decimal > remanescente_decimal:
            valor_cheque_decimal = remanescente_decimal

        st.session_state["valor_cartao_divisao_exibicao"] = decimal_para_texto_monetario(
            remanescente_decimal - valor_cheque_decimal
        )
    else:
        st.session_state["valor_cartao_divisao_exibicao"] = decimal_para_texto_monetario(remanescente_decimal)


def atualizar_valor_total() -> None:
    sincronizar_valores_pagamento()


def atualizar_entrada() -> None:
    sincronizar_valores_pagamento()


def atualizar_valor_cheque_divisao() -> None:
    valor = str(st.session_state.get("valor_cheque_divisao", "")).strip()
    st.session_state["valor_cheque_divisao"] = valor
    sincronizar_valores_pagamento()


def atualizar_modalidades_pagamento() -> None:
    modalidades = st.session_state.get("modalidades_pagamento", [])
    if not isinstance(modalidades, list):
        modalidades = [modalidades] if modalidades else []
    if "Cheque bancario" not in modalidades or "Cartao de credito" not in modalidades:
        st.session_state["valor_cheque_divisao"] = ""
    sincronizar_valores_pagamento()


def gerar_tabela_vencimentos(
    data_inicial: date,
    qtd_parcelas: int,
    intervalo_dias: int,
    valor_total: str,
    *,
    include_cheque_numero: bool = True,
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
                **({"Cheque(s) nº": ""} if include_cheque_numero else {}),
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


def montar_cronograma_contrato(vencimentos: list[dict[str, str]]) -> str:
    if not vencimentos:
        return "Sem cronograma informado."
    return "\n".join(
        (
            f"Parcela {item['Parcela']} - Vencimento: {item['Vencimento']} - "
            f"Valor: R$ {item['Valor']}"
            + (
                f" - Cheque(s) nº: {item['Cheque(s) nº']}"
                if 'Cheque(s) nº' in item and str(item.get('Cheque(s) nº', "")).strip()
                else ""
            )
        )
        for item in vencimentos
    )


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
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(ch for ch in texto if not unicodedata.combining(ch))
    return re.sub(r"\s+", " ", texto).strip()


def limpar_campo_importado(valor: str) -> str:
    texto = valor.strip(" []")
    if not texto:
        return ""
    if re.fullmatch(r"[_\-.Xx\s/]+", texto):
        return ""
    if texto.upper() in {
        "NOME DO VENDEDOR OU LEILOEIRA",
        "ENDERE?O COMPLETO",
        "ENDERECO COMPLETO",
        "NOME DO CREDOR",
        "NOME COMPLETO DO DEVEDOR",
        "NOME DO DEVEDOR",
        "CPF/CNPJ DO CREDOR",
        "CPF DO DEVEDOR",
        "RG DO DEVEDOR",
    }:
        return ""
    return texto


def extrair_bloco_entre(texto: str, inicio: str, fins: list[str]) -> str:
    pos_inicio = texto.find(inicio)
    if pos_inicio == -1:
        return ""
    pos_inicio += len(inicio)
    candidatos = [texto.find(fim, pos_inicio) for fim in fins if texto.find(fim, pos_inicio) != -1]
    pos_fim = min(candidatos) if candidatos else len(texto)
    return texto[pos_inicio:pos_fim].strip()


def extrair_tabela_cronograma(texto: str, rotulo: str, incluir_cheque: bool) -> list[dict[str, str]]:
    bloco = extrair_bloco_entre(
        texto,
        f"{rotulo}:",
        [
            "Opcao 2 - Cartao de credito",
            "Cartao de credito",
            "CLAUSULA TERCEIRA",
            "CLAUSULA QUARTA",
            "CLAUSULA SEGUNDA",
            "Paragrafo primeiro",
            "Paragrafo unico",
        ],
    )
    if not bloco:
        return []

    if incluir_cheque:
        padrao_linha = re.compile(
            r"Parcela\s*(?P<parcela>\d+)\s*-\s*Vencimento:\s*(?P<vencimento>\d{2}/\d{2}/\d{4})\s*-\s*Valor:\s*R\$\s*(?P<valor>[\d\.,]+)(?:\s*-\s*Cheque\(s\)\s*n(?:º|°|o)?\s*:\s*(?P<cheque>[^\n]*))?",
            re.IGNORECASE,
        )
    else:
        padrao_linha = re.compile(
            r"Parcela\s*(?P<parcela>\d+)\s*-\s*Vencimento:\s*(?P<vencimento>\d{2}/\d{2}/\d{4})\s*-\s*Valor:\s*R\$\s*(?P<valor>[\d\.,]+)",
            re.IGNORECASE,
        )

    linhas: list[dict[str, str]] = []
    for match in padrao_linha.finditer(bloco):
        item = {
            "Parcela": match.group("parcela"),
            "Vencimento": match.group("vencimento"),
            "Valor": match.group("valor"),
        }
        if incluir_cheque:
            item["Cheque(s) n?"] = limpar_campo_importado(match.group("cheque") or "")
        linhas.append(item)
    return linhas


def extrair_partes_importadas(texto: str) -> dict[str, str]:
    texto_normalizado = normalizar_texto_importado(texto)

    dados: dict[str, str] = {}

    bloco_credor = extrair_bloco_entre(
        texto_normalizado,
        "CREDOR:",
        ["DEVEDOR:", "CLAUSULA PRIMEIRA", "CL?USULA PRIMEIRA"],
    )
    if bloco_credor:
        padrao_credor = re.compile(
            r"(?P<nome>.*?),\s*inscrito\(a\)\s*no\s*CPF/CNPJ\s*n(?:º|°|o)?\s*(?P<doc>.*?),\s*com\s*endereco\s*a\s*(?P<endereco>.*?);?$",
            re.IGNORECASE,
        )
        credor = padrao_credor.search(bloco_credor)
        if credor:
            dados["credor_nome"] = limpar_campo_importado(credor.group("nome"))
            dados["credor_doc"] = limpar_campo_importado(credor.group("doc"))
            dados["credor_endereco"] = limpar_campo_importado(credor.group("endereco"))

    bloco_devedor = extrair_bloco_entre(
        texto_normalizado,
        "DEVEDOR:",
        ["CLAUSULA PRIMEIRA", "CL?USULA PRIMEIRA", "CLAUSULA SEGUNDA", "CL?USULA SEGUNDA", "E, por estarem"],
    )
    if bloco_devedor:
        padrao_devedor = re.compile(
            r"(?P<nome>.*?),\s*inscrito\(a\)\s*no\s*CPF\s*n(?:º|°|o)?\s*(?P<cpf>.*?),\s*portador\(a\)\s*do\s*RG\s*n(?:º|°|o)?\s*(?P<rg>.*?),\s*residente\s*e\s*domiciliado\(a\)\s*a\s*(?P<endereco>.*?);?$",
            re.IGNORECASE,
        )
        devedor = padrao_devedor.search(bloco_devedor)
        if devedor:
            dados["devedor_nome"] = limpar_campo_importado(devedor.group("nome"))
            dados["devedor_cpf"] = limpar_campo_importado(devedor.group("cpf"))
            dados["devedor_rg"] = limpar_campo_importado(devedor.group("rg"))
            dados["devedor_endereco"] = limpar_campo_importado(devedor.group("endereco"))

    return {chave: valor for chave, valor in dados.items() if valor}


def aplicar_dados_importados(dados_importados: dict[str, str]) -> None:
    if not dados_importados:
        return

    campos_permitidos = {
        "credor_nome",
        "credor_doc",
        "credor_endereco",
        "devedor_nome",
        "devedor_cpf",
        "devedor_rg",
        "devedor_endereco",
    }

    for chave, valor in dados_importados.items():
        if chave in campos_permitidos:
            st.session_state[chave] = valor

def inicializar_estado_formulario() -> None:
    defaults = {
        "credor_nome": "",
        "credor_doc": "",
        "credor_endereco": "",
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
        "cartao_credito_banco": "",
        "cartao_credito_agencia": "",
        "cartao_credito_conta": "",
        "testemunha1_nome": "",
        "testemunha1_cpf": "",
        "testemunha2_nome": "",
        "testemunha2_cpf": "",
    }

    for chave, valor in defaults.items():
        st.session_state.setdefault(chave, valor)

    valores_legados = {
        "credor_nome": "NOME DO VENDEDOR OU LEILOEIRA",
        "credor_doc": "XXXXXXXXXXXX",
        "credor_endereco": "ENDEREÇO COMPLETO",
    }
    for chave, valor_legado in valores_legados.items():
        if st.session_state.get(chave) == valor_legado:
            st.session_state[chave] = ""

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

    def assinatura_rotulada(rotulo: str, nome: str, padrao: str) -> str:
        conteudo = nome.strip() or padrao
        return textwrap.fill(
            f"{rotulo} - {conteudo}",
            width=58,
            subsequent_indent=" " * (len(rotulo) + 3),
        )

    def valor_documento(valor: str, padrao: str = "0,00") -> str:
        numero = valor_monetario_para_decimal(valor)
        if numero is None:
            return padrao
        return decimal_para_texto_monetario(numero)

    modalidades_pagamento = dados["modalidades_pagamento"]
    opcao1 = negrito("X") if "Cheque bancario" in modalidades_pagamento else " "
    opcao2 = negrito("X") if "Cartao de credito" in modalidades_pagamento else " "

    entrada_valor = valor_monetario_para_decimal(dados["entrada"]) or Decimal("0")
    possui_entrada = entrada_valor > 0
    resumo_entrada = (
        f"Desse total originario, foi paga entrada no valor de R$ {negrito(campo_ou_linha(valor_documento(dados['entrada']), '0,00'))} ({negrito(campo_ou_linha(valor_por_extenso(dados['entrada']), 'zero reais'))}). "
        if possui_entrada
        else ""
    )

    resumos_modalidades = []
    if "Cheque bancario" in modalidades_pagamento:
        resumos_modalidades.append(
            f"R$ {negrito(campo_ou_linha(dados['valor_cheque_base'], '0,00'))} ({negrito(campo_ou_linha(valor_por_extenso(dados['valor_cheque_base']), 'zero reais'))}) por cheque bancario"
        )
    if "Cartao de credito" in modalidades_pagamento:
        resumos_modalidades.append(
            f"R$ {negrito(campo_ou_linha(dados['valor_cartao_base'], '0,00'))} ({negrito(campo_ou_linha(valor_por_extenso(dados['valor_cartao_base']), 'zero reais'))}) por cartao de credito"
        )
    resumo_modalidades = "; ".join(resumos_modalidades) if resumos_modalidades else "sem modalidade informada"

    bloco_cheque = ""
    if "Cheque bancario" in modalidades_pagamento:
        bloco_cheque = f"""
[{opcao1}] Opção 1 - Cheque bancário:

Valor destinado ao cheque: R$ {negrito(campo_ou_linha(dados["valor_cheque_base"], "0,00"))} ({negrito(campo_ou_linha(valor_por_extenso(dados["valor_cheque_base"]), "zero reais"))}).

Cronograma de vencimentos do cheque:
{montar_cronograma_contrato(dados["tabela_vencimentos_cheque"])}

Dados do cheque: Banco: {negrito(campo_ou_linha(dados["cheque_unico_banco"], "_______"))} | Agência: {negrito(campo_ou_linha(dados["cheque_unico_agencia"], "___________"))} | Conta: {negrito(campo_ou_linha(dados["cheque_unico_conta"], "____________"))}.
"""

    bloco_cartao = ""
    if "Cartao de credito" in modalidades_pagamento:
        bloco_cartao = f"""
[{opcao2}] Opção 2 - Cartão de crédito:

Valor destinado ao cartão: R$ {negrito(campo_ou_linha(dados["valor_cartao_base"], "0,00"))} ({negrito(campo_ou_linha(valor_por_extenso(dados["valor_cartao_base"]), "zero reais"))}).

Cronograma de vencimentos do cartão:
{montar_cronograma_contrato(dados["tabela_vencimentos_cartao"])}

Dados do cartão: Banco: {negrito(campo_ou_linha(dados["cartao_credito_banco"], "_______"))} | Agência: {negrito(campo_ou_linha(dados["cartao_credito_agencia"], "___________"))} | Conta: {negrito(campo_ou_linha(dados["cartao_credito_conta"], "____________"))}.
"""

    assinatura_credor = assinatura_rotulada("Credor", dados["credor_nome"], "NOME DO CREDOR")
    assinatura_devedor = assinatura_rotulada("Devedor", dados["devedor_nome"], "NOME DO DEVEDOR")

    return f"""CONTRATO DE CONFISSÃO DE DÍVIDA

Pelo presente instrumento particular, de um lado:

CREDOR: {negrito(campo_ou_linha(dados["credor_nome"], "NOME DO VENDEDOR OU LEILOEIRA"))}, inscrito(a) no CPF/CNPJ n {negrito(campo_ou_linha(dados["credor_doc"], "XXXXXXXXXXXX"))}, com endereco a {negrito(campo_ou_linha(dados["credor_endereco"], "ENDERECO COMPLETO"))};

e, de outro lado:

DEVEDOR: {negrito(campo_ou_linha(dados["devedor_nome"], "_____________________________________________________"))}, inscrito(a) no CPF n {negrito(campo_ou_linha(dados["devedor_cpf"], "________________________"))}, portador(a) do RG n {negrito(campo_ou_linha(dados["devedor_rg"], "_______________________"))}, residente e domiciliado(a) a {negrito(campo_ou_linha(dados["devedor_endereco"], "_________________________________________________________________________________________________________________________________________________________"))};

As partes acima qualificadas têm entre si justo e contratado o presente instrumento, que se regerá pelas cláusulas e condições seguintes:

CLÁUSULA PRIMEIRA - DA ORIGEM DA DÍVIDA

O DEVEDOR reconhece que a aquisição objeto deste instrumento totalizou originalmente R$ {negrito(campo_ou_linha(valor_documento(dados["valor_total"], "___________________"), "___________________"))} ({negrito(campo_ou_linha(dados["valor_extenso"], "_______________________________________________________"))}), decorrente da aquisição de animais bovinos realizada em leilão promovido em {negrito(dados["data_leilao"])}, conforme lote(s) n {negrito(campo_ou_linha(dados["lotes"], "_______________________________"))}, adquiridos pelo DEVEDOR. {resumo_entrada}Assim, o DEVEDOR confessa e assume como saldo devedor atual a quantia líquida, certa e exigível de R$ {negrito(campo_ou_linha(dados["valor_remanescente"], "0,00"))} ({negrito(campo_ou_linha(valor_por_extenso(dados["valor_remanescente"]), "zero reais"))}), a ser paga da seguinte forma: {resumo_modalidades}.

CLÁUSULA SEGUNDA - DO VALOR E FORMA DE PAGAMENTO

O valor total da dívida confessada é de R$ {negrito(campo_ou_linha(dados["valor_remanescente"], "__________________"))} ({negrito(campo_ou_linha(valor_por_extenso(dados["valor_remanescente"]), "________________________________________________________"))}), cujo pagamento será realizado por uma das modalidades abaixo, expressamente escolhida e assinalada pelo DEVEDOR no ato da assinatura deste instrumento:
{bloco_cheque}
{bloco_cartao}
Parágrafo primeiro. A modalidade de pagamento escolhida pelo DEVEDOR e assinalada acima integrará este contrato para todos os efeitos legais.

CLÁUSULA TERCEIRA - DA MORA E ENCARGOS INDENIZATÓRIOS

O não pagamento de qualquer parcela ou título nas datas de seus respectivos vencimentos constituirá o DEVEDOR em mora, independentemente de notificação judicial ou extrajudicial, incidindo sobre o valor do débito os seguintes encargos: I - Multa moratória e irredutível de 2% (dois por cento) sobre o valor da parcela em atraso; II - Juros de mora de 1% (um por cento) ao mês, calculados pro rata die (proporcionalmente aos dias de atraso); III - Correção monetária calculada com base na variação positiva do IGP-M/FGV (ou índice oficial que venha a substituí-lo), acumulada desde a data do vencimento até o efetivo pagamento.

Parágrafo único. Caso o CREDOR precise recorrer a serviços advocatícios ou empresas de cobrança para o recebimento do crédito, o DEVEDOR responderá, além do principal e encargos, pelo pagamento das custas, despesas desembolsadas e honorários advocatícios, estes fixados em 10% (dez por cento) para cobrança extrajudicial e 20% (vinte por cento) em caso de ajuizamento de ação judicial.

CLÁUSULA QUARTA - DO VENCIMENTO ANTECIPADO

O atraso superior a {negrito(str(dados["dias_atraso"]))} ({negrito(dados["dias_atraso_extenso"])}) dias no pagamento de qualquer das parcelas pactuadas, ou a ocorrência de devolução por falta de fundos de qualquer dos cheques emitidos, acarretará o vencimento antecipado de todas as parcelas vincendas, tornando-se imediatamente exigível o saldo devedor integral, acrescido de todas as penalidades previstas na Cláusula Terceira, independentemente de prévia notificação ou aviso.

CLÁUSULA QUINTA - DA CONFISSÃO IRREVOGÁVEL E TÍTULO EXECUTIVO

O DEVEDOR declara reconhecer expressamente a existência, legitimidade, certeza, liquidez e exigibilidade da dívida descrita neste instrumento. Este contrato é firmado em caráter irrevogável e irretratável, constituindo-se em Título Executivo Extrajudicial, nos termos do Artigo 784, inciso III, do Código de Processo Civil brasileiro, apto a embasar Ação de Execução imediata.

CLÁUSULA SEXTA - DAS ASSINATURAS ELETRÔNICAS

As partes declaram e concordam que este contrato poderá ser assinado eletronicamente por meio de plataformas de assinatura digital, sendo as assinaturas consideradas válidas, íntegras e plenamente eficazes para todos os fins de direito, nos termos da Medida Provisória n 2.200-2/2001 e da Lei n 14.063/2020.

CLÁUSULA SÉTIMA - DO FORO

Fica eleito o foro da Comarca de {negrito(campo_ou_linha(dados["foro"], "Sao Francisco do Guapore - RO"))}, com renúncia expressa a qualquer outro, por mais privilegiado que seja, para dirimir eventuais controvérsias decorrentes deste contrato.

E, por estarem assim justos e contratados, firmam o presente instrumento em 02 (duas) vias de igual teor e forma, na presença de 02 (duas) testemunhas abaixo assinadas.

{negrito(campo_ou_linha(dados["municipio_assinatura"], "Sao Francisco do Guapore - RO"))}, {negrito(dados["dia_assinatura"])} de {negrito(dados["mes_assinatura"])} de {negrito(dados["ano_assinatura"])}.


______________________________________________
{assinatura_credor}


______________________________________________
{assinatura_devedor}


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
    help="O app tenta extrair automaticamente credor, devedor, financeiro e informações finais do contrato anterior.",
)

if contrato_importado is not None:
    texto_importado = extrair_texto_arquivo(contrato_importado.name, contrato_importado.getvalue())
    dados_importados = extrair_partes_importadas(texto_importado)

    if dados_importados:
        st.info("Arquivo carregado. Clique no botão abaixo para copiar os dados do contrato anterior.")
        if st.button("📥 Importar dados do arquivo", type="primary"):
            aplicar_dados_importados(dados_importados)
            for chave in (
                "tabela_vencimentos_cheque_importada",
                "tabela_vencimentos_cartao_importada",
                "tabela_vencimentos_cheque_editor",
            ):
                st.session_state.pop(chave, None)
            st.success("Dados de credor e devedor importados do contrato anterior.")
            st.rerun()
    else:
        st.warning("Não foi possível localizar automaticamente os dados do contrato nesse arquivo.")

data_leilao_padrao = date.today()

with st.container():
    st.subheader("Dados do contrato")

    with st.expander("Credor e devedor", expanded=False):
        col_credor, col_devedor = st.columns(2)
        with col_credor:
            st.markdown("### Credor")
            credor_nome = st.text_input(
                "Nome do credor ou da leiloeira",
                key="credor_nome",
                placeholder="NOME DO VENDEDOR OU LEILOEIRA",
            )
            credor_doc = st.text_input(
                "CPF/CNPJ do credor",
                key="credor_doc",
                placeholder="000.000.000-00 ou 00.000.000/0001-00",
            )
            exibir_warning_campo_numerico(credor_doc, "CPF/CNPJ do credor")
            credor_endereco = st.text_area(
                "Endereço do credor",
                key="credor_endereco",
                placeholder="ENDEREÇO COMPLETO",
                height=95,
            )

        with col_devedor:
            st.markdown("### Devedor")
            devedor_nome = st.text_input(
                "Nome completo do devedor",
                key="devedor_nome",
                placeholder="NOME COMPLETO DO DEVEDOR",
            )
            devedor_doc_col, devedor_rg_col = st.columns(2)
            with devedor_doc_col:
                devedor_cpf = st.text_input(
                    "CPF do devedor",
                    key="devedor_cpf",
                    placeholder="000.000.000-00",
                )
                exibir_warning_campo_numerico(devedor_cpf, "CPF do devedor")
            with devedor_rg_col:
                devedor_rg = st.text_input(
                    "RG do devedor",
                    key="devedor_rg",
                    placeholder="12.345.678-9 SSP/UF",
                )
            devedor_endereco = st.text_area(
                "Endereço do devedor",
                key="devedor_endereco",
                placeholder="ENDEREÇO COMPLETO",
                height=95,
            )

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
        st.markdown("### Leilão e dívida")
        col_data_leilao, col_divida, col_valor_extenso = st.columns(3)
        with col_data_leilao:
            data_leilao = st.date_input("Data inicial", value=data_leilao_padrao, format="DD/MM/YYYY", key="data_leilao")
        with col_divida:
            st.text_input("D\u00edvida", key="valor_total", on_change=atualizar_valor_total)
            valor_total = st.session_state.get("valor_total", "")
            exibir_warning_campo_numerico(valor_total, "Divida")
        with col_valor_extenso:
            valor_extenso = st.text_input("Valor por extenso", valor_por_extenso(valor_total), disabled=True)

        lotes = st.text_input("N\u00famero dos lotes", key="lotes")
        exibir_warning_campo_numerico(lotes, "Número dos lotes")

        st.markdown("### Pagamento")
        modalidades_pagamento = st.multiselect(
            "Modalidade escolhida",
            OPCOES_MODALIDADES_PAGAMENTO,
            key="modalidades_pagamento",
            on_change=atualizar_modalidades_pagamento,
        )
        col_entrada, col_valor_atual, col_valor_remanescente_extenso = st.columns(3)
        with col_entrada:
            st.text_input("Com entrada? Informe o valor:", key="entrada", on_change=atualizar_entrada)
            entrada = st.session_state.get("entrada", "0")
            exibir_warning_campo_numerico(entrada, "Com entrada? Informe o valor")
        sincronizar_valores_pagamento()
        valor_remanescente = st.session_state.get("valor_atual_exibicao", "")
        with col_valor_atual:
            st.text_input("Valor atual", disabled=True, key="valor_atual_exibicao")
        with col_valor_remanescente_extenso:
            st.text_input("Valor remanescente por extenso", disabled=True, key="valor_remanescente_extenso_exibicao")
        mostrar_formulario_cheque = "Cheque bancario" in modalidades_pagamento
        mostrar_formulario_cartao = "Cartao de credito" in modalidades_pagamento
        valor_cheque_base = valor_remanescente
        valor_cartao_base = valor_remanescente
        valor_divisao_informado = True
        tabela_vencimentos = []
        tabela_vencimentos_cartao = []

        if mostrar_formulario_cheque and mostrar_formulario_cartao:
            col_valor_cheque, col_valor_cartao = st.columns(2)
            with col_valor_cheque:
                valor_cheque_divisao = st.text_input(
                    "Do valor remanescente, quanto sera pago com cheque?",
                    key="valor_cheque_divisao",
                    on_change=atualizar_valor_cheque_divisao,
                )
                exibir_warning_campo_numerico(valor_cheque_divisao, "Do valor remanescente, quanto sera pago com cheque")
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
            tabela_vencimentos_importada = st.session_state.get("tabela_vencimentos_cheque_importada")
            if tabela_vencimentos_importada:
                tabela_vencimentos_df = pd.DataFrame(tabela_vencimentos_importada)
            else:
                tabela_vencimentos_df = pd.DataFrame(
                    gerar_tabela_vencimentos(
                        data_inicial=data_leilao,
                        qtd_parcelas=int(qtd_parcelas),
                        intervalo_dias=int(dias_atraso),
                        valor_total=valor_cheque_base,
                        include_cheque_numero=True,
                    )
                )

            if not tabela_vencimentos_df.empty and 'Cheque(s) nº' in tabela_vencimentos_df.columns:
                tabela_vencimentos_df["Cheque(s) nº"] = tabela_vencimentos_df["Cheque(s) nº"].fillna("")

            st.markdown("### Dados bancários")
            col_cheque1, col_cheque2, col_cheque3 = st.columns(3)
            with col_cheque1:
                cheque_unico_banco = st.text_input("Banco", key="cheque_unico_banco")
            with col_cheque2:
                cheque_unico_agencia = st.text_input("Agencia", key="cheque_unico_agencia")
                exibir_warning_campo_numerico(cheque_unico_agencia, "Agencia")
            with col_cheque3:
                cheque_unico_conta = st.text_input("Conta", key="cheque_unico_conta")
                exibir_warning_campo_numerico(cheque_unico_conta, "Conta")

            if not tabela_vencimentos_df.empty:
                chave_editor_cheque = f"tabela_vencimentos_cheque_editor_{int(qtd_parcelas)}_{int(dias_atraso)}_{valor_cheque_base}"
                tabela_vencimentos = st.data_editor(
                    tabela_vencimentos_df,
                    hide_index=True,
                    width="stretch",
                    disabled=["Parcela", "Vencimento", "Valor"],
                    column_config={
                        "Cheque(s) nº": st.column_config.TextColumn(
                            "Cheque(s) nº",
                            help="Informe o número do cheque correspondente a esta parcela.",
                        )
                    },
                    key=chave_editor_cheque,
                ).to_dict("records")
                st.caption(f"Soma das parcelas: {somar_valores_tabela(tabela_vencimentos)}")
        else:
            qtd_parcelas = st.session_state.get("qtd_parcelas", 4)
            dias_atraso = st.session_state.get("dias_atraso", 30)
            valor_parcela = valor_parcela_calculado(valor_cheque_base, int(qtd_parcelas))
            cheque_unico_banco = st.session_state.get("cheque_unico_banco", "")
            cheque_unico_agencia = st.session_state.get("cheque_unico_agencia", "")
            cheque_unico_conta = st.session_state.get("cheque_unico_conta", "")

        if exibir_formulario_cheque and exibir_formulario_cartao:
            st.markdown('<div class="bloco-separador"></div>', unsafe_allow_html=True)

        if exibir_formulario_cartao:
            st.markdown("### Para Cartão de Crédito")
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
            tabela_vencimentos_cartao_importada = st.session_state.get("tabela_vencimentos_cartao_importada")
            if tabela_vencimentos_cartao_importada:
                tabela_vencimentos_cartao_df = pd.DataFrame(tabela_vencimentos_cartao_importada)
            else:
                tabela_vencimentos_cartao_df = pd.DataFrame(
                    gerar_tabela_vencimentos(
                        data_inicial=data_leilao,
                        qtd_parcelas=int(qtd_parcelas_cartao),
                        intervalo_dias=int(dias_atraso_cartao),
                        valor_total=valor_cartao_base,
                        include_cheque_numero=False,
                    )
                )

            st.dataframe(tabela_vencimentos_cartao_df, hide_index=True, width="stretch")
            st.caption(f"Soma das parcelas: {somar_valores_tabela(tabela_vencimentos_cartao_df.to_dict('records'))}")

            col_cartao1, col_cartao2, col_cartao3 = st.columns(3)
            with col_cartao1:
                cartao_credito_banco = st.text_input("Banco", key="cartao_credito_banco")
            with col_cartao2:
                cartao_credito_agencia = st.text_input("Agencia", key="cartao_credito_agencia")
                exibir_warning_campo_numerico(cartao_credito_agencia, "Agencia")
            with col_cartao3:
                cartao_credito_conta = st.text_input("Conta", key="cartao_credito_conta")
                exibir_warning_campo_numerico(cartao_credito_conta, "Conta")
        else:
            qtd_parcelas_cartao = st.session_state.get("qtd_parcelas_cartao", 4)
            dias_atraso_cartao = st.session_state.get("dias_atraso_cartao", 30)
            valor_parcela_cartao = valor_parcela_calculado(valor_cartao_base, int(qtd_parcelas_cartao))
            cartao_credito_banco = st.session_state.get("cartao_credito_banco", "")
            cartao_credito_agencia = st.session_state.get("cartao_credito_agencia", "")
            cartao_credito_conta = st.session_state.get("cartao_credito_conta", "")

    with st.expander("Informações finais", expanded=False):
        st.markdown("### Foro e assinatura")
        foro = st.text_input("Comarca/UF", key="foro")
        municipio_assinatura = st.text_input("Município da assinatura", key="municipio_assinatura")
        data_assinatura = st.date_input("Data da assinatura", value=date.today(), format="DD/MM/YYYY", key="data_assinatura")

        st.markdown("### Assinantes")
        st.markdown(f"**Credor:** {credor_nome or '____________________________'}")
        st.markdown(f"**Devedor:** {devedor_nome or '____________________________'}")

        col_testemunha1, col_testemunha2 = st.columns(2)
        with col_testemunha1:
            st.markdown("### Testemunha 1")
            testemunha1_nome = st.text_input("Nome da testemunha 1", key="testemunha1_nome")
            testemunha1_cpf = st.text_input("CPF da testemunha 1", key="testemunha1_cpf")
            exibir_warning_campo_numerico(testemunha1_cpf, "CPF da testemunha 1")

        with col_testemunha2:
            st.markdown("### Testemunha 2")
            testemunha2_nome = st.text_input("Nome da testemunha 2", key="testemunha2_nome")
            testemunha2_cpf = st.text_input("CPF da testemunha 2", key="testemunha2_cpf")
            exibir_warning_campo_numerico(testemunha2_cpf, "CPF da testemunha 2")

    atualizar_preview = st.button("🔄 Atualizar pré-visualização", width="stretch")

parcela_datas = [
    data_leilao + timedelta(days=int(dias_atraso) * indice)
    for indice in range(1, 5)
]
parcela1_data, parcela2_data, parcela3_data, parcela4_data = parcela_datas
parcela1_banco = st.session_state.get("parcela1_banco", "")
parcela2_banco = st.session_state.get("parcela2_banco", "")
parcela3_banco = st.session_state.get("parcela3_banco", "")
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
    "tabela_vencimentos_cheque": tabela_vencimentos,
    "dias_atraso": dias_atraso,
    "dias_atraso_extenso": numero_por_extenso(dias_atraso),
    "qtd_parcelas_cartao": qtd_parcelas_cartao,
    "valor_parcela_cartao": valor_parcela_cartao,
    "tabela_vencimentos_cartao": tabela_vencimentos_cartao,
    "dias_atraso_cartao": dias_atraso_cartao,
    "cheque_unico_banco": cheque_unico_banco,
    "cheque_unico_agencia": cheque_unico_agencia,
    "cheque_unico_conta": cheque_unico_conta,
    "cartao_credito_banco": cartao_credito_banco,
    "cartao_credito_agencia": cartao_credito_agencia,
    "cartao_credito_conta": cartao_credito_conta,
    "parcela1_data": formatar_data(parcela1_data),
    "parcela1_banco": parcela1_banco,
    "parcela2_data": formatar_data(parcela2_data),
    "parcela2_banco": parcela2_banco,
    "parcela3_data": formatar_data(parcela3_data),
    "parcela3_banco": parcela3_banco,
    "parcela4_data": formatar_data(parcela4_data),
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
