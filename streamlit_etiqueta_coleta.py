import re
from datetime import datetime
from io import BytesIO
from pathlib import Path

import streamlit as st

try:
    from reportlab.graphics.barcode import code128
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas

    REPORTLAB_AVAILABLE = True
except Exception:
    code128 = None
    A4 = None
    canvas = None
    REPORTLAB_AVAILABLE = False

try:
    from openpyxl import load_workbook

    OPENPYXL_AVAILABLE = True
except Exception:
    load_workbook = None
    OPENPYXL_AVAILABLE = False


APP_NAME = "COLETA"
PROJETO_REDE = "REDE"
PLANILHA_BASE_CRED = "bases padrão + cred.xlsx"

DESTINOS = [
    "CTDI DO BR - SP",
    "FLEXTRONIC",
    "FEDEX CAJAMAR - SP",
    "DHL LOUVEIRA - SP",
]

PROJETOS = [
    "CIELO - POS",
    "CIELO - TEF",
    "CIELO - TRANSF",
    "FISERV",
    "MOOZ",
    "STONE",
    "PICPAY",
    "PAGBANK",
    "CTRENDS",
    "C6BANK",
    "ADYEN",
    "CLOUDWALK",
    PROJETO_REDE,
]

PREFIXOS_ROMANEIO = {
    "CIELO - POS": "1.2/",
    "CIELO - TEF": "2.2/",
    "CIELO - TRANSF": "1.3/",
    "FISERV": "34.3/",
    "MOOZ": "42.3/",
    "STONE": "41.3/",
    "PICPAY": "49.3/",
    "PAGBANK": "53.3/",
    "CTRENDS": "39.3/",
    "C6BANK": "43.3/",
    "ADYEN": "45.3/",
    "CLOUDWALK": "40.3/",
    PROJETO_REDE: "51.2/",
}

CRED_OPTIONS = [
    ("CRED369", "POLO REDE PONTA GROSSA"),
    ("CRED385", "POLO REDE SJ DOS CAMPOS"),
    ("CRED368", "POLO REDE SANTOS"),
    ("CRED372", "POLO REDE CUR PINHAIS"),
    ("CRED382", "POLO REDE LONDRINA"),
    ("CRED371", "POLO REDE CURITIBA"),
    ("CRED370", "POLO REDE MARINGA"),
    ("CRED384", "POLO REDE PINDA"),
    ("CRED383", "POLO REDE CASCAVEL"),
    ("CRED408", "POLO REDE FORTALEZA"),
    ("CRED409", "POLO REDE JUAZ DO NORTE"),
    ("CRED421", "POLO REDE BH"),
    ("CRED412", "POLO REDE JUIZ DE FORA"),
    ("CRED411", "POLO REDE TEOFILO OTONI"),
    ("CRED419", "POLO REDE GOV VALADARES"),
    ("CRED416", "POLO REDE MONTES CLAROS"),
    ("CRED422", "POLO REDE IPATINGA"),
    ("CRED425", "POLO REDE GUARULHOS"),
]
CRED_CODES = [code for code, _ in CRED_OPTIONS]

MM_TO_POINTS = 72 / 25.4
DEFAULT_CONFIG_OUTROS = {
    "largura_mm": 90.0,
    "altura_mm": 100.0,
    "espacamento_pt": 5.0,
    "escala_fonte": 2.5,
}
DEFAULT_CONFIG_REDE = {
    "largura_mm": 150.0,
    "altura_mm": 100.0,
    "espacamento_pt": 5.0,
    "escala_fonte": 1.8,
}


def _clamp(value: float, minimum: float, maximum: float) -> float:
    return max(minimum, min(value, maximum))


def _apenas_numeros(value: str) -> str:
    return re.sub(r"\D", "", value or "")


@st.cache_data(show_spinner=False)
def _carregar_origens_e_cred() -> tuple[list[str], dict[str, str], str | None]:
    path = Path(__file__).with_name(PLANILHA_BASE_CRED)
    if not path.exists():
        return [], {}, f"Planilha nao encontrada: {PLANILHA_BASE_CRED}"
    if not OPENPYXL_AVAILABLE:
        return [], {}, "Biblioteca openpyxl nao encontrada para leitura da planilha."

    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb.active
    origens: list[str] = []
    origem_para_cred: dict[str, str] = {}

    for row in ws.iter_rows(min_row=2, max_col=4, values_only=True):
        origem_raw = row[0]
        cred_raw = row[2]
        if origem_raw is None:
            continue
        origem = str(origem_raw).strip()
        if not origem:
            continue
        if origem not in origens:
            origens.append(origem)

        if cred_raw is not None and str(cred_raw).strip():
            origem_para_cred[origem] = str(cred_raw).strip().upper()

    return origens, origem_para_cred, None


def _erro_numero_obrigatorio(campo: str, valor: str, max_digitos: int | None = None) -> str | None:
    if not valor:
        return f"O campo '{campo}' e obrigatorio."
    if not valor.isdigit():
        return f"O campo '{campo}' aceita apenas numeros."
    if max_digitos and len(valor) > max_digitos:
        return f"O campo '{campo}' permite no maximo {max_digitos} digitos."
    return None


def _validar_entradas(entradas: dict) -> list[str]:
    erros: list[str] = []
    if not entradas["origem"]:
        erros.append("O campo 'Origem' e obrigatorio.")
    if not entradas["destino"]:
        erros.append("O campo 'Destino' e obrigatorio.")
    if not entradas["projeto"]:
        erros.append("O campo 'Projeto' e obrigatorio.")

    if entradas["projeto"] == PROJETO_REDE:
        tecnologia = (entradas.get("tecnologia") or "").strip().upper()
        if not tecnologia:
            erros.append("O campo 'Tecnologia' e obrigatorio.")
        elif len(tecnologia) > 3:
            erros.append("O campo 'Tecnologia' permite no maximo 3 caracteres.")
        elif not re.fullmatch(r"[A-Za-z]{1,3}", tecnologia):
            erros.append("O campo 'Tecnologia' aceita apenas letras.")

        err_nf = _erro_numero_obrigatorio("Nota Fiscal", entradas.get("nota_fiscal", ""), 8)
        if err_nf:
            erros.append(err_nf)
        err_os = _erro_numero_obrigatorio("OS", entradas.get("os", ""), 10)
        if err_os:
            erros.append(err_os)
        if not entradas.get("numero_cred"):
            erros.append("O campo 'Numero CRED' e obrigatorio.")
    else:
        err_rom = _erro_numero_obrigatorio("Romaneio", entradas.get("romaneio_sufixo", ""))
        if err_rom:
            erros.append(err_rom)
        err_nf = _erro_numero_obrigatorio("NR NF", entradas.get("nr_nf", ""))
        if err_nf:
            erros.append(err_nf)
        err_id = _erro_numero_obrigatorio("ID FEDEX", entradas.get("id_fedex", ""))
        if err_id:
            erros.append(err_id)
        err_vol = _erro_numero_obrigatorio("Volume", entradas.get("volume_total", ""), 3)
        if err_vol:
            erros.append(err_vol)
        elif int(entradas["volume_total"]) <= 0:
            erros.append("O campo 'Volume' deve ser maior que zero.")

    if entradas["largura_mm"] <= 0 or entradas["altura_mm"] <= 0:
        erros.append("Largura e altura da etiqueta devem ser maiores que zero.")
    if entradas["espacamento_linhas"] < 0:
        erros.append("Espacamento de linhas deve ser maior ou igual a zero.")
    if entradas["escala_fonte"] <= 0:
        erros.append("Escala de fonte deve ser maior que zero.")
    if entradas["ajuste_cabecalho"] < 0:
        erros.append("Ajuste cabecalho deve ser maior ou igual a zero.")
    if entradas["ajuste_rodape"] < 0:
        erros.append("Ajuste rodape deve ser maior ou igual a zero.")
    return erros


def _montar_dados_padrao(entradas: dict) -> dict:
    prefixo = PREFIXOS_ROMANEIO[entradas["projeto"]]
    romaneio = f"{prefixo}{entradas['romaneio_sufixo']}"
    total = int(entradas["volume_total"])
    total_fmt = str(total).zfill(3)
    codigo_base = _apenas_numeros(romaneio)
    data_emissao = datetime.now().strftime("%d/%m/%Y")
    id_fedex_data = f"{entradas['id_fedex']} - {data_emissao}"

    etiquetas = []
    for i in range(1, total + 1):
        atual_fmt = str(i).zfill(3)
        volume = f"{atual_fmt}/{total_fmt}"
        etiquetas.append(
            {
                "mode": "PADRAO",
                "origem": entradas["origem"],
                "destino": entradas["destino"],
                "projeto": entradas["projeto"],
                "romaneio": romaneio,
                "nr_nf": entradas["nr_nf"],
                "id_fedex_data": id_fedex_data,
                "volume": volume,
                "codigo_barras": f"{codigo_base}{atual_fmt}{total_fmt}",
            }
        )

    return {
        "mode": "PADRAO",
        "origem": entradas["origem"],
        "destino": entradas["destino"],
        "projeto": entradas["projeto"],
        "romaneio": romaneio,
        "nr_nf": entradas["nr_nf"],
        "id_fedex_data": id_fedex_data,
        "etiquetas": etiquetas,
    }


def _montar_dados_rede(entradas: dict) -> dict:
    data_emissao = datetime.now().strftime("%d/%m/%Y")
    etiqueta = {
        "mode": "REDE",
        "titulo": "OPERACAO REVERSA",
        "tecnologia": entradas["tecnologia"].strip().upper(),
        "origem": entradas["origem"],
        "destino": entradas["destino"],
        "numero_cred": entradas["numero_cred"],
        "nota_fiscal": entradas["nota_fiscal"],
        "data_emissao": data_emissao,
        "os": entradas["os"],
        "volume": "-",
    }
    return {
        "mode": "REDE",
        "origem": entradas["origem"],
        "destino": entradas["destino"],
        "projeto": entradas["projeto"],
        "etiquetas": [etiqueta],
    }


def _montar_dados(entradas: dict) -> dict:
    if entradas["projeto"] == PROJETO_REDE:
        return _montar_dados_rede(entradas)
    return _montar_dados_padrao(entradas)


def _layout_paginas_a4(largura_pt: float, altura_pt: float) -> dict | None:
    if not REPORTLAB_AVAILABLE or A4 is None:
        return None

    pagina_w, pagina_h = A4
    margem = 12 * MM_TO_POINTS
    gap = 4 * MM_TO_POINTS
    area_w = pagina_w - (2 * margem)
    area_h = pagina_h - (2 * margem)

    cols = int((area_w + gap) // (largura_pt + gap))
    rows = int((area_h + gap) // (altura_pt + gap))
    if cols < 1 or rows < 1:
        return None

    x0 = margem
    y0 = pagina_h - margem - altura_pt
    step_x = largura_pt + gap
    step_y = altura_pt + gap

    positions = []
    for row in range(rows):
        for col in range(cols):
            positions.append((x0 + col * step_x, y0 - row * step_y))

    return {"por_pagina": cols * rows, "positions": positions}


def _desenhar_etiqueta_padrao_pdf(
    c,
    x,
    y,
    largura_pt,
    altura_pt,
    dados,
    espacamento_extra,
    escala_fonte_usuario,
    ajuste_cabecalho=0.0,
    ajuste_rodape=0.0,
):
    ref_w = 105 * MM_TO_POINTS
    ref_h = 148.5 * MM_TO_POINTS
    scale = min(largura_pt / ref_w, altura_pt / ref_h)

    border = _clamp(0.85 * scale, 0.6, 1.4)
    c.setLineWidth(border)
    c.rect(x, y, largura_pt, altura_pt)

    pad = max(3.5 * MM_TO_POINTS, 6 * MM_TO_POINTS * scale)
    ax, ay = x + pad, y + pad
    aw, ah = largura_pt - (2 * pad), altura_pt - (2 * pad)

    lines = [
        ("ORIGEM", dados["origem"]),
        ("DESTINO", dados["destino"]),
        ("ROMANEIO", dados["romaneio"]),
        ("PROJETO", dados["projeto"]),
        ("NR NF", dados["nr_nf"]),
        ("VOLUME", dados["volume"]),
    ]

    title_f = _clamp(12 * scale * escala_fonte_usuario, 8, 42)
    label_f = _clamp(9.2 * scale * escala_fonte_usuario, 6, 30)
    value_f = _clamp(9.8 * scale * escala_fonte_usuario, 6, 32)
    code_f = _clamp(8.2 * scale * escala_fonte_usuario, 6, 24)
    id_f = _clamp(7.0 * scale * escala_fonte_usuario, 5.5, 18)
    gap = max(2.2, 1.8 * MM_TO_POINTS * scale) + (espacamento_extra * 1.35)
    block_gap = max(4.0, 3 * MM_TO_POINTS * scale)

    h_code = code_f * 1.45
    h_id = id_f * 1.45
    h_bar = _clamp(ah * 0.2, 10 * MM_TO_POINTS, ah * 0.26)
    gap_id_bar = max(2.2, 1.4 * MM_TO_POINTS * scale)
    gap_bar_code = max(2.2, 1.3 * MM_TO_POINTS * scale)
    h_footer = h_code + gap_bar_code + h_bar + gap_id_bar + h_id

    y_top = ay + ah
    y_title = y_top - title_f
    c.setFont("Helvetica-Bold", title_f)
    c.drawCentredString(x + largura_pt / 2, y_title, APP_NAME)

    y_div = y_title - (gap * 0.9)
    c.setLineWidth(max(0.4, border * 0.7))
    c.line(ax, y_div, ax + aw, y_div)

    header_gap = max(block_gap, gap * 0.9) + ajuste_cabecalho + (espacamento_extra * 0.65)
    y_details_top = y_div - header_gap
    y_details_base = ay + h_footer + block_gap
    available = y_details_top - y_details_base

    step = max(label_f, value_f) + gap
    needed = (len(lines) * max(label_f, value_f)) + ((len(lines) - 1) * gap)
    if needed > available and available > 0:
        factor = available / needed
        label_f *= factor
        value_f *= factor
        step = max(label_f, value_f) + (gap * factor)

    c.setFont("Helvetica-Bold", label_f)
    label_w = max(c.stringWidth(f"{k}:", "Helvetica-Bold", label_f) for k, _ in lines)
    value_x = ax + label_w + max(4, 2.4 * MM_TO_POINTS * scale)

    y_line = y_details_top - max(label_f, value_f)
    for k, v in lines:
        c.setFont("Helvetica-Bold", label_f)
        c.drawString(ax, y_line, f"{k}:")
        c.setFont("Helvetica", value_f)
        c.drawString(value_x, y_line, str(v))
        y_line -= step

    code = dados["codigo_barras"]
    target_w = aw * 0.78
    modules = max(80, (11 * len(code)) + 35)
    bar_w = _clamp(target_w / modules, 0.16, 1.6)
    bar = code128.Code128(code, barHeight=h_bar, barWidth=bar_w)
    for _ in range(20):
        if bar.width > aw * 0.82 and bar_w > 0.14:
            bar_w *= 0.95
            bar = code128.Code128(code, barHeight=h_bar, barWidth=bar_w)
            continue
        if bar.width < aw * 0.72 and bar_w < 2.0:
            bar_w *= 1.03
            bar = code128.Code128(code, barHeight=h_bar, barWidth=bar_w)
            continue
        break

    y_code = ay
    y_bar = y_code + h_code + gap_bar_code
    x_bar = ax + ((aw - bar.width) / 2)
    bar.drawOn(c, x_bar, y_bar)

    c.setFont("Helvetica", id_f)
    y_id = y_bar + h_bar + gap_id_bar
    c.drawCentredString(ax + (aw / 2), y_id, dados["id_fedex_data"])

    c.setFont("Helvetica", code_f)
    c.drawCentredString(ax + (aw / 2), y_code, code)


def _desenhar_etiqueta_rede_pdf(
    c,
    x,
    y,
    largura_pt,
    altura_pt,
    dados,
    espacamento_extra,
    escala_fonte_usuario,
    ajuste_cabecalho=0.0,
    ajuste_rodape=0.0,
):
    ref_w = 90 * MM_TO_POINTS
    ref_h = 100 * MM_TO_POINTS
    scale = min(largura_pt / ref_w, altura_pt / ref_h)

    border = _clamp(0.9 * scale, 0.6, 1.4)
    c.setLineWidth(border)
    c.rect(x, y, largura_pt, altura_pt)

    pad = max(3.0 * MM_TO_POINTS, 5.0 * MM_TO_POINTS * scale)
    ax, ay = x + pad, y + pad
    aw, ah = largura_pt - (2 * pad), altura_pt - (2 * pad)

    title_f = _clamp(10.5 * scale * escala_fonte_usuario, 7, 36)
    label_f = _clamp(8.4 * scale * escala_fonte_usuario, 6, 28)
    value_f = _clamp(8.8 * scale * escala_fonte_usuario, 6, 28)
    gap = max(2.0, 1.7 * MM_TO_POINTS * scale) + (espacamento_extra * 1.35)

    y_top = ay + ah
    y_title = y_top - title_f
    c.setFont("Helvetica-Bold", title_f)
    c.drawString(ax, y_title, "OPERACAO REVERSA")
    c.setFont("Helvetica-Bold", label_f)
    c.drawRightString(ax + aw, y_title, f"Tecnologia: {dados['tecnologia']}")

    y_div = y_title - (gap * 0.8)
    c.setLineWidth(max(0.4, border * 0.7))
    c.line(ax, y_div, ax + aw, y_div)

    fields = [
        ("Origem", dados["origem"]),
        ("Destino", dados["destino"]),
        ("Numero CRED", dados["numero_cred"]),
        ("Nota Fiscal", dados["nota_fiscal"]),
        ("Data Emissao", dados["data_emissao"]),
        ("OS", dados["os"]),
        ("Volume", "-"),
    ]

    c.setFont("Helvetica-Bold", label_f)
    label_w = max(c.stringWidth(f"{k}:", "Helvetica-Bold", label_f) for k, _ in fields)
    value_x = ax + label_w + max(5, 2.0 * MM_TO_POINTS * scale)

    header_gap = max(gap * 1.1, label_f * 1.1) + ajuste_cabecalho + (espacamento_extra * 0.65)
    y_text = y_div - header_gap
    step = max(label_f, value_f) + gap
    for k, v in fields:
        c.setFont("Helvetica-Bold", label_f)
        c.drawString(ax, y_text, f"{k}:")
        c.setFont("Helvetica", value_f)
        c.drawString(value_x, y_text, str(v))
        y_text -= step

    rodape_margem = max(gap * 1.1, label_f) + (espacamento_extra * 0.45)
    rodape_texto_y = max(ay + rodape_margem, y_text + (gap * 0.2)) + ajuste_rodape
    rodape_linha_gap = max(label_f * 1.2, gap * 1.15)
    rodape_linha_y = rodape_texto_y + rodape_linha_gap
    c.line(ax, rodape_linha_y, ax + aw, rodape_linha_y)
    c.setFont("Helvetica-Bold", label_f)
    c.drawString(ax, rodape_texto_y, f"Ordem Servico: {dados['os']}")
    c.drawRightString(ax + aw, rodape_texto_y, f"Nota Fiscal: {dados['nota_fiscal']}")


def _gerar_pdf_bytes(
    dados_lote: dict,
    largura_mm: float,
    altura_mm: float,
    espacamento_extra: float,
    escala_fonte_usuario: float,
    ajuste_cabecalho: float = 0.0,
    ajuste_rodape: float = 0.0,
) -> bytes:
    if not REPORTLAB_AVAILABLE:
        raise RuntimeError("Biblioteca reportlab nao encontrada.")

    largura_pt = largura_mm * MM_TO_POINTS
    altura_pt = altura_mm * MM_TO_POINTS
    layout = _layout_paginas_a4(largura_pt, altura_pt)
    if not layout:
        raise ValueError("Com esse tamanho de etiqueta nao cabe nenhuma unidade na folha A4.")

    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    c.setTitle(APP_NAME)

    for i, etiqueta in enumerate(dados_lote["etiquetas"]):
        slot = i % layout["por_pagina"]
        if i > 0 and slot == 0:
            c.showPage()
        x, y = layout["positions"][slot]

        if etiqueta["mode"] == "REDE":
            _desenhar_etiqueta_rede_pdf(
                c,
                x,
                y,
                largura_pt,
                altura_pt,
                etiqueta,
                espacamento_extra,
                escala_fonte_usuario,
                ajuste_cabecalho,
                ajuste_rodape,
            )
        else:
            _desenhar_etiqueta_padrao_pdf(
                c,
                x,
                y,
                largura_pt,
                altura_pt,
                etiqueta,
                espacamento_extra,
                escala_fonte_usuario,
                ajuste_cabecalho,
                ajuste_rodape,
            )

    c.showPage()
    c.save()
    return buf.getvalue()


def _init_state() -> None:
    if "resultado" not in st.session_state:
        st.session_state["resultado"] = None
    if "pdf_bytes" not in st.session_state:
        st.session_state["pdf_bytes"] = None
    if "erros" not in st.session_state:
        st.session_state["erros"] = []
    if "origem_rede_prev" not in st.session_state:
        st.session_state["origem_rede_prev"] = ""
    if "cred_code" not in st.session_state:
        st.session_state["cred_code"] = ""
    if "cfg_modo_prev" not in st.session_state:
        st.session_state["cfg_modo_prev"] = ""


def main() -> None:
    st.set_page_config(page_title=APP_NAME, layout="wide")
    _init_state()

    origens, origem_para_cred, aviso_planilha = _carregar_origens_e_cred()

    st.title(APP_NAME)
    if aviso_planilha:
        st.warning(aviso_planilha)

    st.subheader("Origem / Destino / Projeto")
    c1, c2, c3 = st.columns(3)
    origem = c1.selectbox("Origem *", [""] + origens, index=0)
    destino = c2.selectbox("Destino *", [""] + DESTINOS, index=0)
    projeto = c3.selectbox("Projeto *", [""] + PROJETOS, index=0)
    is_rede = projeto == PROJETO_REDE
    modo_cfg = PROJETO_REDE if is_rede else "OUTROS"
    cfg_padrao = DEFAULT_CONFIG_REDE if is_rede else DEFAULT_CONFIG_OUTROS
    if st.session_state.get("cfg_modo_prev") != modo_cfg:
        st.session_state["cfg_largura_mm"] = cfg_padrao["largura_mm"]
        st.session_state["cfg_altura_mm"] = cfg_padrao["altura_mm"]
        st.session_state["cfg_espacamento_pt"] = cfg_padrao["espacamento_pt"]
        st.session_state["cfg_escala_fonte"] = cfg_padrao["escala_fonte"]
        st.session_state["cfg_modo_prev"] = modo_cfg
    else:
        if "cfg_largura_mm" not in st.session_state:
            st.session_state["cfg_largura_mm"] = cfg_padrao["largura_mm"]
        if "cfg_altura_mm" not in st.session_state:
            st.session_state["cfg_altura_mm"] = cfg_padrao["altura_mm"]
        if "cfg_espacamento_pt" not in st.session_state:
            st.session_state["cfg_espacamento_pt"] = cfg_padrao["espacamento_pt"]
        if "cfg_escala_fonte" not in st.session_state:
            st.session_state["cfg_escala_fonte"] = cfg_padrao["escala_fonte"]

    tecnologia = ""
    nota_fiscal = ""
    os_num = ""
    numero_cred = ""
    romaneio_sufixo = ""
    nr_nf = ""
    id_fedex = ""
    volume_total = ""

    if is_rede:
        st.subheader("Projeto REDE")
        cred_sugerido = ""
        if origem and origem in origem_para_cred:
            cred_sugerido = origem_para_cred[origem]

        if st.session_state["origem_rede_prev"] != origem:
            st.session_state["origem_rede_prev"] = origem
            st.session_state["cred_code"] = cred_sugerido if cred_sugerido in CRED_CODES else ""

        r1, r2, r3 = st.columns(3)
        tecnologia = r1.text_input("Tecnologia * (texto, max 3)", max_chars=3).strip()
        nota_fiscal = r2.text_input("Nota Fiscal * (max 8)", max_chars=8).strip()
        os_num = r3.text_input("OS * (max 10)", max_chars=10).strip()
        r4, r5 = st.columns([1, 2])
        r4.text_input("Data Emissao", value=datetime.now().strftime("%d/%m/%Y"), disabled=True)
        numero_cred = r5.selectbox("Numero CRED *", [""] + CRED_CODES, key="cred_code")
        if cred_sugerido:
            st.caption(f"CRED sugerido pela Origem: {cred_sugerido}")
    else:
        st.subheader("Outros Projetos")
        prefixo = PREFIXOS_ROMANEIO.get(projeto, "")
        o1, o2 = st.columns([1, 2])
        o1.text_input("Prefixo Romaneio", value=prefixo, disabled=True)
        romaneio_sufixo = o2.text_input("Romaneio (numeros apos /) *", max_chars=20).strip()
        o3, o4 = st.columns(2)
        nr_nf = o3.text_input("NR NF *", max_chars=20).strip()
        id_fedex = o4.text_input("ID FEDEX *", max_chars=20).strip()
        volume_total = st.text_input("Volume (qtd total de etiquetas) *", max_chars=3).strip()

    st.subheader("Configuracao da Etiqueta")
    s1, s2, s3, s4 = st.columns(4)
    largura_mm = float(s1.number_input("Largura (mm)", min_value=1.0, key="cfg_largura_mm"))
    altura_mm = float(s2.number_input("Altura (mm)", min_value=1.0, key="cfg_altura_mm"))
    espacamento = float(
        s3.number_input("Espacamento linhas (pt)", min_value=0.0, key="cfg_espacamento_pt")
    )
    escala_fonte = float(
        s4.number_input("Escala de fonte", min_value=0.1, key="cfg_escala_fonte")
    )
    s5, s6 = st.columns(2)
    ajuste_cabecalho = float(
        s5.number_input("Ajuste cabecalho (pt)", min_value=0.0, value=3.0, step=0.5)
    )
    ajuste_rodape = float(
        s6.number_input("Ajuste rodape (pt)", min_value=0.0, value=3.0, step=0.5)
    )

    if st.button("Gerar etiqueta(s)", type="primary", use_container_width=True):
        entradas = {
            "origem": origem,
            "destino": destino,
            "projeto": projeto,
            "tecnologia": tecnologia,
            "nota_fiscal": _apenas_numeros(nota_fiscal),
            "os": _apenas_numeros(os_num),
            "numero_cred": (numero_cred or "").strip().upper(),
            "romaneio_sufixo": _apenas_numeros(romaneio_sufixo),
            "nr_nf": _apenas_numeros(nr_nf),
            "id_fedex": _apenas_numeros(id_fedex),
            "volume_total": _apenas_numeros(volume_total),
            "largura_mm": largura_mm,
            "altura_mm": altura_mm,
            "espacamento_linhas": espacamento,
            "escala_fonte": escala_fonte,
            "ajuste_cabecalho": ajuste_cabecalho,
            "ajuste_rodape": ajuste_rodape,
        }

        st.session_state["erros"] = _validar_entradas(entradas)
        st.session_state["resultado"] = None
        st.session_state["pdf_bytes"] = None

        if not st.session_state["erros"]:
            dados = _montar_dados(entradas)
            st.session_state["resultado"] = {"dados": dados, "config": entradas}
            if REPORTLAB_AVAILABLE:
                try:
                    st.session_state["pdf_bytes"] = _gerar_pdf_bytes(
                        dados,
                        largura_mm=largura_mm,
                        altura_mm=altura_mm,
                        espacamento_extra=espacamento,
                        escala_fonte_usuario=escala_fonte,
                        ajuste_cabecalho=ajuste_cabecalho,
                        ajuste_rodape=ajuste_rodape,
                    )
                except Exception as exc:
                    st.session_state["erros"].append(str(exc))
                    st.session_state["resultado"] = None

    if st.session_state["erros"]:
        for erro in st.session_state["erros"]:
            st.error(erro)

    resultado = st.session_state.get("resultado")
    if resultado:
        dados = resultado["dados"]
        etiquetas = dados["etiquetas"]
        st.success(f"Etiquetas validadas. Quantidade: {len(etiquetas)}")

        layout = _layout_paginas_a4(
            resultado["config"]["largura_mm"] * MM_TO_POINTS,
            resultado["config"]["altura_mm"] * MM_TO_POINTS,
        )
        if layout:
            por_folha = layout["por_pagina"]
            folhas = (len(etiquetas) + por_folha - 1) // por_folha
            st.info(f"Etiquetas por folha A4: {por_folha} | Folhas estimadas: {folhas}")

        st.markdown("**Preview**")
        if dados["mode"] == "REDE":
            e = etiquetas[0]
            st.code(
                (
                    f"Titulo: {e['titulo']}\n"
                    f"Tecnologia: {e['tecnologia']}\n"
                    f"Origem: {e['origem']}\n"
                    f"Destino: {e['destino']}\n"
                    f"Numero CRED: {e['numero_cred']}\n"
                    f"Nota Fiscal: {e['nota_fiscal']}\n"
                    f"Data Emissao: {e['data_emissao']}\n"
                    f"OS: {e['os']}\n"
                    "Volume: -"
                ),
                language="text",
            )
        else:
            preview = [f"{e['volume']} -> {e['codigo_barras']}" for e in etiquetas[:120]]
            if len(etiquetas) > 120:
                preview.append(f"... (mostrando 120 de {len(etiquetas)} etiquetas)")
            st.code(
                (
                    f"Origem: {dados['origem']}\n"
                    f"Destino: {dados['destino']}\n"
                    f"Projeto: {dados['projeto']}\n"
                    f"Romaneio: {dados['romaneio']}\n"
                    f"NR NF: {dados['nr_nf']}\n"
                    f"ID FEDEX: {dados['id_fedex_data']}\n"
                    f"Quantidade de etiquetas: {len(etiquetas)}\n\n"
                    "Volumes / Codigos de barras:\n"
                    + "\n".join(preview)
                ),
                language="text",
            )

        if not REPORTLAB_AVAILABLE:
            st.warning("Instale reportlab para gerar PDF: pip install reportlab")
        elif st.session_state["pdf_bytes"]:
            st.download_button(
                "Baixar etiqueta(s) em PDF",
                data=st.session_state["pdf_bytes"],
                file_name=f"etiqueta_{datetime.now():%Y%m%d_%H%M%S}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )


if __name__ == "__main__":
    main()
