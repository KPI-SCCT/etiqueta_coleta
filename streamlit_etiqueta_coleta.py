import re
from datetime import datetime
from io import BytesIO

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


APP_NAME = "COLETA"
DESTINOS = ["CTDI DO BR - SP", "FLEXTRONIC", "FEDEX CAJAMAR - SP"]
PROJETOS = ["CIELO - POS", "CIELO - TEF", "CIELO - TRANSF", "FISERV", "MOOZ", "STONE", "PICPAY", "PAGBANK", "CTRENDS", "C6BANK", "ADYEN", "CLOUDWALK"]
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
    
}
MM_TO_POINTS = 72 / 25.4


def _clamp(valor: float, minimo: float, maximo: float) -> float:
    return max(minimo, min(valor, maximo))


def _apenas_numeros(valor: str) -> str:
    return re.sub(r"\D", "", valor or "")


def _campo_obrigatorio_numerico(nome: str, valor: str) -> str | None:
    if not valor:
        return f"O campo '{nome}' e obrigatorio."
    if not valor.isdigit():
        return f"O campo '{nome}' aceita apenas numeros."
    return None


def _validar_entradas(entradas: dict) -> list[str]:
    erros: list[str] = []

    if not entradas["destino"]:
        erros.append("O campo 'Destino' e obrigatorio.")
    if not entradas["projeto"]:
        erros.append("O campo 'Projeto' e obrigatorio.")

    for nome, chave in [
        ("Romaneio", "romaneio_sufixo"),
        ("NR NF", "nr_nf"),
        ("ID FEDEX", "id_fedex"),
        ("Volume", "volume_total"),
    ]:
        erro = _campo_obrigatorio_numerico(nome, entradas[chave])
        if erro:
            erros.append(erro)

    if entradas["volume_total"] and len(entradas["volume_total"]) > 3:
        erros.append("O campo 'Volume' permite no maximo 3 digitos.")
    if (
        entradas["volume_total"]
        and entradas["volume_total"].isdigit()
        and int(entradas["volume_total"]) <= 0
    ):
        erros.append("O campo 'Volume' deve ser maior que zero.")

    if entradas["largura_mm"] <= 0 or entradas["altura_mm"] <= 0:
        erros.append("Largura e altura da etiqueta devem ser maiores que zero.")
    if entradas["espacamento_linhas"] < 0:
        erros.append("O espacamento de linhas deve ser maior ou igual a zero.")
    if entradas["escala_fonte"] <= 0:
        erros.append("A escala de fonte deve ser maior que zero.")

    return erros


def _montar_dados(entradas: dict) -> dict:
    prefixo = PREFIXOS_ROMANEIO[entradas["projeto"]]
    romaneio = f"{prefixo}{entradas['romaneio_sufixo']}"
    total_volumes = int(entradas["volume_total"])
    vol_total_fmt = str(total_volumes).zfill(3)
    base_codigo = _apenas_numeros(romaneio)
    data_emissao = datetime.now().strftime("%d/%m/%Y")
    id_fedex_data = f"{entradas['id_fedex']} - {data_emissao}"

    etiquetas = []
    for indice in range(1, total_volumes + 1):
        vol_atual_fmt = str(indice).zfill(3)
        volume_fmt = f"{vol_atual_fmt}/{vol_total_fmt}"
        etiquetas.append(
            {
                "destino": entradas["destino"],
                "projeto": entradas["projeto"],
                "romaneio": romaneio,
                "nr_nf": entradas["nr_nf"],
                "id_fedex_data": id_fedex_data,
                "volume": volume_fmt,
                "codigo_barras": f"{base_codigo}{vol_atual_fmt}{vol_total_fmt}",
            }
        )

    return {
        "destino": entradas["destino"],
        "projeto": entradas["projeto"],
        "romaneio": romaneio,
        "nr_nf": entradas["nr_nf"],
        "id_fedex": entradas["id_fedex"],
        "id_fedex_data": id_fedex_data,
        "volume_total": total_volumes,
        "etiquetas": etiquetas,
    }


def _layout_paginas_a4(largura_pt: float, altura_pt: float) -> dict | None:
    if not REPORTLAB_AVAILABLE or A4 is None:
        return None

    pagina_largura, pagina_altura = A4
    margem = 12 * MM_TO_POINTS
    gap = 4 * MM_TO_POINTS
    area_largura = pagina_largura - (2 * margem)
    area_altura = pagina_altura - (2 * margem)

    colunas = int((area_largura + gap) // (largura_pt + gap))
    linhas = int((area_altura + gap) // (altura_pt + gap))
    if colunas < 1 or linhas < 1:
        return None

    x_inicial = margem
    y_inicial = pagina_altura - margem - altura_pt
    passo_x = largura_pt + gap
    passo_y = altura_pt + gap

    posicoes = []
    for linha in range(linhas):
        for coluna in range(colunas):
            x = x_inicial + (coluna * passo_x)
            y = y_inicial - (linha * passo_y)
            posicoes.append((x, y))

    return {
        "colunas": colunas,
        "linhas": linhas,
        "por_pagina": colunas * linhas,
        "posicoes": posicoes,
    }


def _desenhar_etiqueta_pdf(
    c,
    x: float,
    y: float,
    largura_pt: float,
    altura_pt: float,
    dados: dict,
    espacamento_extra: float,
    escala_fonte_usuario: float,
) -> None:
    referencia_largura = 105 * MM_TO_POINTS
    referencia_altura = 148.5 * MM_TO_POINTS
    escala_etiqueta = min(largura_pt / referencia_largura, altura_pt / referencia_altura)

    borda = _clamp(0.85 * escala_etiqueta, 0.6, 1.4)
    c.setLineWidth(borda)
    c.rect(x, y, largura_pt, altura_pt)

    pad = max(3.5 * MM_TO_POINTS, 6 * MM_TO_POINTS * escala_etiqueta)
    area_x = x + pad
    area_y = y + pad
    area_largura = largura_pt - (2 * pad)
    area_altura = altura_pt - (2 * pad)

    linhas = [
        ("DESTINO", dados["destino"]),
        ("ROMANEIO", dados["romaneio"]),
        ("PROJETO", dados["projeto"]),
        ("NR NF", dados["nr_nf"]),
        ("VOLUME", dados["volume"]),
    ]

    fonte_titulo = _clamp(12 * escala_etiqueta * escala_fonte_usuario, 8, 28)
    fonte_label = _clamp(9.2 * escala_etiqueta * escala_fonte_usuario, 6, 20)
    fonte_valor = _clamp(9.8 * escala_etiqueta * escala_fonte_usuario, 6, 20)
    fonte_codigo = _clamp(8.2 * escala_etiqueta * escala_fonte_usuario, 6, 16)
    fonte_identificador = _clamp(7.0 * escala_etiqueta * escala_fonte_usuario, 5.5, 12)
    gap_linha = max(2.2, 1.8 * MM_TO_POINTS * escala_etiqueta) + espacamento_extra
    gap_bloco = max(4.0, 3 * MM_TO_POINTS * escala_etiqueta)

    altura_codigo = fonte_codigo * 1.45
    altura_identificador = fonte_identificador * 1.45
    altura_barcode = _clamp(area_altura * 0.2, 10 * MM_TO_POINTS, area_altura * 0.26)
    gap_identificador_barra = max(2.2, 1.4 * MM_TO_POINTS * escala_etiqueta)
    gap_barra_codigo = max(2.2, 1.3 * MM_TO_POINTS * escala_etiqueta)
    bloco_barcode_altura = (
        altura_codigo
        + gap_barra_codigo
        + altura_barcode
        + gap_identificador_barra
        + altura_identificador
    )

    y_topo = area_y + area_altura
    y_titulo = y_topo - fonte_titulo
    c.setFont("Helvetica-Bold", fonte_titulo)
    c.drawCentredString(x + (largura_pt / 2), y_titulo, APP_NAME)

    y_divisor = y_titulo - (gap_linha * 0.9)
    c.setLineWidth(max(0.4, borda * 0.7))
    c.line(area_x, y_divisor, area_x + area_largura, y_divisor)

    y_detalhes_topo = y_divisor - gap_bloco
    y_detalhes_base = area_y + bloco_barcode_altura + gap_bloco
    altura_disponivel_detalhes = y_detalhes_topo - y_detalhes_base

    passo_linha = max(fonte_label, fonte_valor) + gap_linha
    altura_necessaria = (len(linhas) * max(fonte_label, fonte_valor)) + (
        (len(linhas) - 1) * gap_linha
    )
    if altura_necessaria > altura_disponivel_detalhes and altura_disponivel_detalhes > 0:
        fator = altura_disponivel_detalhes / altura_necessaria
        fonte_label *= fator
        fonte_valor *= fator
        passo_linha = max(fonte_label, fonte_valor) + (gap_linha * fator)

    c.setFont("Helvetica-Bold", fonte_label)
    largura_labels = max(
        c.stringWidth(f"{titulo}:", "Helvetica-Bold", fonte_label) for titulo, _ in linhas
    )
    gap_label_valor = max(4, 2.4 * MM_TO_POINTS * escala_etiqueta)
    valor_x = area_x + largura_labels + gap_label_valor

    y_linha = y_detalhes_topo - max(fonte_label, fonte_valor)
    for titulo, valor in linhas:
        c.setFont("Helvetica-Bold", fonte_label)
        c.drawString(area_x, y_linha, f"{titulo}:")
        c.setFont("Helvetica", fonte_valor)
        c.drawString(valor_x, y_linha, valor)
        y_linha -= passo_linha

    codigo = dados["codigo_barras"]
    largura_alvo = area_largura * 0.78
    modulos_estimados = max(80, (11 * len(codigo)) + 35)
    bar_width = _clamp(largura_alvo / modulos_estimados, 0.16, 1.6)
    barcode = code128.Code128(codigo, barHeight=altura_barcode, barWidth=bar_width)

    for _ in range(20):
        if barcode.width > area_largura * 0.82 and bar_width > 0.14:
            bar_width *= 0.95
            barcode = code128.Code128(codigo, barHeight=altura_barcode, barWidth=bar_width)
            continue
        if barcode.width < area_largura * 0.72 and bar_width < 2.0:
            bar_width *= 1.03
            barcode = code128.Code128(codigo, barHeight=altura_barcode, barWidth=bar_width)
            continue
        break

    codigo_y = area_y
    barcode_y = codigo_y + altura_codigo + gap_barra_codigo
    barcode_x = area_x + ((area_largura - barcode.width) / 2)
    barcode.drawOn(c, barcode_x, barcode_y)

    c.setFont("Helvetica", fonte_identificador)
    id_y = barcode_y + altura_barcode + gap_identificador_barra
    c.drawCentredString(area_x + (area_largura / 2), id_y, dados["id_fedex_data"])

    c.setFont("Helvetica", fonte_codigo)
    c.drawCentredString(area_x + (area_largura / 2), codigo_y, codigo)


def _gerar_pdf_bytes(
    dados_lote: dict,
    largura_mm: float,
    altura_mm: float,
    espacamento_extra: float,
    escala_fonte_usuario: float,
) -> bytes:
    if not REPORTLAB_AVAILABLE:
        raise RuntimeError("Biblioteca reportlab nao encontrada.")

    largura_pt = largura_mm * MM_TO_POINTS
    altura_pt = altura_mm * MM_TO_POINTS

    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    c.setTitle(APP_NAME)

    layout = _layout_paginas_a4(largura_pt, altura_pt)
    if not layout:
        raise ValueError(
            "Com esse tamanho de etiqueta nao cabe nenhuma unidade na folha A4. "
            "Reduza largura/altura."
        )

    for indice, etiqueta in enumerate(dados_lote["etiquetas"]):
        indice_slot = indice % layout["por_pagina"]
        if indice > 0 and indice_slot == 0:
            c.showPage()

        x, y = layout["posicoes"][indice_slot]
        _desenhar_etiqueta_pdf(
            c,
            x,
            y,
            largura_pt,
            altura_pt,
            etiqueta,
            espacamento_extra,
            escala_fonte_usuario,
        )

    c.showPage()

    c.save()
    return buffer.getvalue()


def _inicializar_estado() -> None:
    if "resultado" not in st.session_state:
        st.session_state["resultado"] = None
    if "pdf_bytes" not in st.session_state:
        st.session_state["pdf_bytes"] = None
    if "erros" not in st.session_state:
        st.session_state["erros"] = []


def main() -> None:
    st.set_page_config(page_title=APP_NAME, layout="wide")
    _inicializar_estado()

    st.title(APP_NAME)
    st.caption("Versao Streamlit para teste e publicacao no Streamlit Cloud.")

    with st.form("form_etiqueta", clear_on_submit=False):
        col1, col2 = st.columns(2)
        destino = col1.selectbox("Destino *", [""] + DESTINOS, index=0)
        projeto = col2.selectbox("Projeto *", [""] + PROJETOS, index=0)

        prefixo_romaneio = PREFIXOS_ROMANEIO.get(projeto, "")
        colr1, colr2 = st.columns([1, 2])
        colr1.text_input("Prefixo Romaneio", value=prefixo_romaneio, disabled=True)
        romaneio_sufixo = colr2.text_input(
            "Romaneio (somente numeros apos /) *", max_chars=20
        ).strip()

        coln1, coln2 = st.columns(2)
        nr_nf = coln1.text_input("NR NF *", max_chars=20).strip()
        id_fedex = coln2.text_input("ID FEDEX *", max_chars=20).strip()

        volume_total = st.text_input(
            "Volume (qtd total de etiquetas) *", max_chars=3
        ).strip()

        st.markdown("**Configuracao da Etiqueta**")
        colc1, colc2, colc3, colc4 = st.columns(4)
        largura_mm = float(colc1.number_input("Largura (mm)", min_value=1.0, value=90.0))
        altura_mm = float(colc2.number_input("Altura (mm)", min_value=1.0, value=100.0))
        espacamento_linhas = float(
            colc3.number_input("Espaçamento linhas (pt)", min_value=0.0, value=5.0)
        )
        escala_fonte = float(colc4.number_input("Escala de fonte", min_value=0.1, value=2.5))

        gerar = st.form_submit_button("Gerar etiqueta")

    if gerar:
        entradas = {
            "destino": destino,
            "projeto": projeto,
            "romaneio_sufixo": romaneio_sufixo,
            "nr_nf": nr_nf,
            "id_fedex": id_fedex,
            "volume_total": volume_total,
            "largura_mm": largura_mm,
            "altura_mm": altura_mm,
            "espacamento_linhas": espacamento_linhas,
            "escala_fonte": escala_fonte,
        }

        st.session_state["erros"] = _validar_entradas(entradas)
        st.session_state["resultado"] = None
        st.session_state["pdf_bytes"] = None

        if st.session_state["erros"]:
            pass
        else:
            dados = _montar_dados(entradas)
            st.session_state["resultado"] = {"dados": dados, "config": entradas}
            if REPORTLAB_AVAILABLE:
                try:
                    st.session_state["pdf_bytes"] = _gerar_pdf_bytes(
                        dados,
                        largura_mm=largura_mm,
                        altura_mm=altura_mm,
                        espacamento_extra=espacamento_linhas,
                        escala_fonte_usuario=escala_fonte,
                    )
                except Exception as erro:
                    st.session_state["erros"].append(str(erro))
                    st.session_state["resultado"] = None

    if st.session_state["erros"]:
        for erro in st.session_state["erros"]:
            st.error(erro)

    resultado = st.session_state.get("resultado")
    if resultado:
        dados = resultado["dados"]
        etiquetas = dados["etiquetas"]
        st.success(f"Etiquetas validadas e prontas para PDF. Quantidade: {len(etiquetas)}.")

        largura_pt = resultado["config"]["largura_mm"] * MM_TO_POINTS
        altura_pt = resultado["config"]["altura_mm"] * MM_TO_POINTS
        layout = _layout_paginas_a4(largura_pt, altura_pt)
        if layout:
            por_folha = layout["por_pagina"]
            folhas = (len(etiquetas) + por_folha - 1) // por_folha
            st.info(
                f"Etiquetas por folha A4: {por_folha} | Folhas estimadas: {folhas}"
            )

        limite_preview = 120
        linhas_preview = [
            f"{item['volume']} -> {item['codigo_barras']}"
            for item in etiquetas[:limite_preview]
        ]
        if len(etiquetas) > limite_preview:
            linhas_preview.append(
                f"... (mostrando {limite_preview} de {len(etiquetas)} etiquetas)"
            )

        st.markdown("**Preview dos dados**")
        st.code(
            (
                f"Destino: {dados['destino']}\n"
                f"Projeto: {dados['projeto']}\n"
                f"Romaneio: {dados['romaneio']}\n"
                f"NR NF: {dados['nr_nf']}\n"
                f"ID FEDEX: {dados['id_fedex_data']}\n"
                f"Quantidade de etiquetas: {len(etiquetas)}\n\n"
                "Volumes / Codigos de barras:\n"
                + "\n".join(linhas_preview)
            ),
            language="text",
        )

        if not REPORTLAB_AVAILABLE:
            st.warning("Instale o reportlab para gerar PDF: pip install reportlab")
        else:
            nome_arquivo = f"etiqueta_{datetime.now():%Y%m%d_%H%M%S}.pdf"
            st.download_button(
                "Baixar etiqueta em PDF",
                data=st.session_state["pdf_bytes"],
                file_name=nome_arquivo,
                mime="application/pdf",
                use_container_width=True,
            )

    st.markdown("---")
    st.caption(
        "Para Streamlit Cloud, publique este arquivo no GitHub e configure as dependencias "
        "(streamlit e reportlab)."
    )


if __name__ == "__main__":
    main()
