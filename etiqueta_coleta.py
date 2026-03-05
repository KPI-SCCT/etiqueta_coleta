import re
import tempfile
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

try:
    from reportlab.graphics.barcode import code128
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas

    REPORTLAB_AVAILABLE = True
except Exception:
    REPORTLAB_AVAILABLE = False

try:
    import win32api
    import win32print

    PYWIN32_AVAILABLE = True
except Exception:
    PYWIN32_AVAILABLE = False


APP_NAME = "COLETA"
DESTINOS = ["CTDI DO BR - SP", "FLEXTRONIC", "FEDEX CAJAMAR - SP"]
PROJETOS = ["CIELO - POS", "CIELO - TEF", "FISERV - POS", "FISERV - TEF"]
PREFIXOS_ROMANEIO = {
    "CIELO - POS": "1.2/",
    "CIELO - TEF": "2.2/",
    "FISERV - POS": "34.1/",
    "FISERV - TEF": "34.2/",
}
PRINTER_DEFAULT_LABEL = "Padrao do sistema"
MM_TO_POINTS = 72 / 25.4


class EtiquetaColetaApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(APP_NAME)
        self.root.minsize(860, 680)

        self.romaneio_prefixo_var = tk.StringVar()
        self.romaneio_sufixo_var = tk.StringVar()
        self.nr_nf_var = tk.StringVar()
        self.id_fedex_var = tk.StringVar()
        self.volume_qtd_var = tk.StringVar()
        self.codigo_barras_var = tk.StringVar()
        self.etiqueta_largura_var = tk.StringVar(value="105")
        self.etiqueta_altura_var = tk.StringVar(value="148.5")
        self.espacamento_linhas_var = tk.StringVar(value="3.5")
        self.escala_fonte_var = tk.StringVar(value="1.00")
        self.impressora_var = tk.StringVar(value=PRINTER_DEFAULT_LABEL)

        self.vcmd_digitos = (self.root.register(self._validar_digitos), "%P")
        self.vcmd_volume = (self.root.register(self._validar_volume), "%P")
        self.vcmd_decimal = (self.root.register(self._validar_decimal), "%P")

        self._montar_layout()
        self._carregar_impressoras()
        self._atualizar_prefixo_romaneio()

    def _montar_layout(self) -> None:
        frame = ttk.Frame(self.root, padding=14)
        frame.grid(sticky="nsew")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        for col in range(4):
            frame.columnconfigure(col, weight=1)

        ttk.Label(
            frame, text=APP_NAME, font=("Segoe UI", 13, "bold")
        ).grid(row=0, column=0, columnspan=4, sticky="ew", pady=(0, 10))

        ttk.Label(frame, text="Destino:").grid(row=1, column=0, sticky="w")
        self.lb_destino = tk.Listbox(frame, height=len(DESTINOS), exportselection=False)
        self.lb_destino.grid(row=2, column=0, columnspan=2, sticky="nsew", padx=(0, 8))
        for destino in DESTINOS:
            self.lb_destino.insert(tk.END, destino)
        self.lb_destino.selection_set(0)

        ttk.Label(frame, text="Projeto:").grid(row=1, column=2, sticky="w")
        self.lb_projeto = tk.Listbox(frame, height=len(PROJETOS), exportselection=False)
        self.lb_projeto.grid(row=2, column=2, columnspan=2, sticky="nsew")
        for projeto in PROJETOS:
            self.lb_projeto.insert(tk.END, projeto)
        self.lb_projeto.selection_set(0)
        self.lb_projeto.bind("<<ListboxSelect>>", self._on_projeto_change)

        ttk.Label(frame, text="Romaneio:").grid(row=3, column=0, sticky="w", pady=(12, 0))
        romaneio_frame = ttk.Frame(frame)
        romaneio_frame.grid(row=4, column=0, columnspan=2, sticky="w", padx=(0, 8))
        ttk.Label(
            romaneio_frame,
            textvariable=self.romaneio_prefixo_var,
            width=7,
            anchor="w",
            font=("Segoe UI", 10, "bold"),
        ).pack(side=tk.LEFT)
        ttk.Entry(
            romaneio_frame,
            textvariable=self.romaneio_sufixo_var,
            validate="key",
            validatecommand=self.vcmd_digitos,
            width=24,
        ).pack(side=tk.LEFT, padx=(6, 0))

        ttk.Label(frame, text="NR NF:").grid(row=3, column=2, sticky="w", pady=(12, 0))
        ttk.Entry(
            frame,
            textvariable=self.nr_nf_var,
            validate="key",
            validatecommand=self.vcmd_digitos,
        ).grid(row=4, column=2, columnspan=2, sticky="ew")

        ttk.Label(frame, text="Volume (qtd total):").grid(row=5, column=0, sticky="w", pady=(12, 0))
        volume_frame = ttk.Frame(frame)
        volume_frame.grid(row=6, column=0, columnspan=2, sticky="w", padx=(0, 8))
        ttk.Entry(
            volume_frame,
            textvariable=self.volume_qtd_var,
            width=6,
            validate="key",
            validatecommand=self.vcmd_volume,
        ).pack(side=tk.LEFT)
        ttk.Label(volume_frame, text="(max 3 digitos)", font=("Segoe UI", 9)).pack(
            side=tk.LEFT, padx=(8, 0)
        )

        ttk.Label(frame, text="Tamanho etiqueta (mm):").grid(
            row=5, column=2, sticky="w", pady=(12, 0)
        )
        tamanho_frame = ttk.Frame(frame)
        tamanho_frame.grid(row=6, column=2, columnspan=2, sticky="w")
        ttk.Label(tamanho_frame, text="Largura").pack(side=tk.LEFT)
        ttk.Entry(
            tamanho_frame,
            textvariable=self.etiqueta_largura_var,
            width=8,
            validate="key",
            validatecommand=self.vcmd_decimal,
        ).pack(side=tk.LEFT, padx=(6, 12))
        ttk.Label(tamanho_frame, text="Altura").pack(side=tk.LEFT)
        ttk.Entry(
            tamanho_frame,
            textvariable=self.etiqueta_altura_var,
            width=8,
            validate="key",
            validatecommand=self.vcmd_decimal,
        ).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Label(tamanho_frame, text="Espacamento (pt)").pack(side=tk.LEFT, padx=(12, 0))
        ttk.Entry(
            tamanho_frame,
            textvariable=self.espacamento_linhas_var,
            width=6,
            validate="key",
            validatecommand=self.vcmd_decimal,
        ).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Label(tamanho_frame, text="Escala fonte").pack(side=tk.LEFT, padx=(12, 0))
        ttk.Entry(
            tamanho_frame,
            textvariable=self.escala_fonte_var,
            width=6,
            validate="key",
            validatecommand=self.vcmd_decimal,
        ).pack(side=tk.LEFT, padx=(6, 0))

        ttk.Label(frame, text="ID FEDEX:").grid(row=7, column=0, sticky="w", pady=(12, 0))
        ttk.Entry(
            frame,
            textvariable=self.id_fedex_var,
            validate="key",
            validatecommand=self.vcmd_digitos,
        ).grid(row=8, column=0, columnspan=2, sticky="ew", padx=(0, 8))

        ttk.Label(frame, text="Impressora:").grid(row=9, column=0, sticky="w", pady=(12, 0))
        impressora_frame = ttk.Frame(frame)
        impressora_frame.grid(row=10, column=0, columnspan=4, sticky="ew")
        impressora_frame.columnconfigure(0, weight=1)
        self.cmb_impressora = ttk.Combobox(
            impressora_frame,
            textvariable=self.impressora_var,
            state="readonly",
        )
        self.cmb_impressora.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(
            impressora_frame, text="Atualizar impressoras", command=self._carregar_impressoras
        ).grid(row=0, column=1, sticky="e")

        ttk.Label(frame, text="Codigo de barras:").grid(
            row=11, column=0, sticky="w", pady=(12, 0)
        )
        ttk.Entry(
            frame,
            textvariable=self.codigo_barras_var,
            state="readonly",
            font=("Consolas", 10),
        ).grid(row=12, column=0, columnspan=4, sticky="ew")

        ttk.Label(frame, text="Preview da etiqueta:").grid(
            row=13, column=0, sticky="w", pady=(12, 0)
        )
        self.txt_preview = tk.Text(frame, height=10, state="disabled", wrap="word")
        self.txt_preview.grid(row=14, column=0, columnspan=4, sticky="nsew")
        frame.rowconfigure(14, weight=1)

        botoes_frame = ttk.Frame(frame)
        botoes_frame.grid(row=15, column=0, columnspan=4, sticky="e", pady=(12, 0))
        ttk.Button(botoes_frame, text="Gerar codigo", command=self.gerar_codigo).pack(
            side=tk.LEFT, padx=6
        )
        ttk.Button(botoes_frame, text="Salvar PDF", command=self.salvar_pdf).pack(
            side=tk.LEFT, padx=6
        )
        ttk.Button(botoes_frame, text="Imprimir etiqueta", command=self.imprimir).pack(
            side=tk.LEFT, padx=6
        )

    @staticmethod
    def _validar_digitos(valor: str) -> bool:
        return valor == "" or valor.isdigit()

    @staticmethod
    def _validar_volume(valor: str) -> bool:
        return valor == "" or (valor.isdigit() and len(valor) <= 3)

    @staticmethod
    def _validar_decimal(valor: str) -> bool:
        if valor == "":
            return True
        valor = valor.replace(",", ".")
        try:
            float(valor)
            return True
        except ValueError:
            return False

    @staticmethod
    def _clamp(valor: float, minimo: float, maximo: float) -> float:
        return max(minimo, min(valor, maximo))

    def _on_projeto_change(self, _: tk.Event) -> None:
        self._atualizar_prefixo_romaneio()

    def _atualizar_prefixo_romaneio(self) -> None:
        projeto = self._valor_listbox(self.lb_projeto, "Projeto", exibir_erro=False)
        if projeto:
            self.romaneio_prefixo_var.set(PREFIXOS_ROMANEIO[projeto])

    @staticmethod
    def _apenas_numeros(valor: str) -> str:
        return re.sub(r"\D", "", valor)

    def _valor_listbox(
        self, listbox: tk.Listbox, campo: str, exibir_erro: bool = True
    ) -> str:
        selecao = listbox.curselection()
        if not selecao:
            if exibir_erro:
                messagebox.showerror("Campo obrigatorio", f"Selecione um valor para {campo}.")
            return ""
        return listbox.get(selecao[0])

    def _coletar_dados(self) -> dict | None:
        destino = self._valor_listbox(self.lb_destino, "Destino")
        projeto = self._valor_listbox(self.lb_projeto, "Projeto")
        if not destino or not projeto:
            return None

        sufixo_romaneio = self.romaneio_sufixo_var.get().strip()
        nr_nf = self.nr_nf_var.get().strip()
        id_fedex = self.id_fedex_var.get().strip()
        volume_qtd = self.volume_qtd_var.get().strip()

        if not sufixo_romaneio:
            messagebox.showerror("Campo obrigatorio", "Informe os numeros do Romaneio.")
            return None
        if not nr_nf:
            messagebox.showerror("Campo obrigatorio", "Informe o campo NR NF.")
            return None
        if not id_fedex:
            messagebox.showerror("Campo obrigatorio", "Informe o campo ID FEDEX.")
            return None
        if not volume_qtd:
            messagebox.showerror(
                "Campo obrigatorio", "Informe o campo Volume (qtd total)."
            )
            return None

        total_volumes = int(volume_qtd)
        if total_volumes <= 0:
            messagebox.showerror("Campo invalido", "O campo Volume deve ser maior que zero.")
            return None

        romaneio = f"{PREFIXOS_ROMANEIO[projeto]}{sufixo_romaneio}"
        base_codigo = self._apenas_numeros(romaneio)
        vol_total_fmt = str(total_volumes).zfill(3)
        data_emissao = datetime.now().strftime("%d/%m/%Y")
        id_fedex_data = f"{id_fedex} - {data_emissao}"

        etiquetas = []
        for indice in range(1, total_volumes + 1):
            vol_atual_fmt = str(indice).zfill(3)
            volume_fmt = f"{vol_atual_fmt}/{vol_total_fmt}"
            etiquetas.append(
                {
                    "destino": destino,
                    "projeto": projeto,
                    "romaneio": romaneio,
                    "nr_nf": nr_nf,
                    "id_fedex_data": id_fedex_data,
                    "volume": volume_fmt,
                    "codigo_barras": f"{base_codigo}{vol_atual_fmt}{vol_total_fmt}",
                }
            )

        return {
            "destino": destino,
            "projeto": projeto,
            "romaneio": romaneio,
            "nr_nf": nr_nf,
            "id_fedex": id_fedex,
            "data_emissao": data_emissao,
            "id_fedex_data": id_fedex_data,
            "volume_total": total_volumes,
            "etiquetas": etiquetas,
        }

    def _atualizar_preview(self, dados: dict) -> None:
        etiquetas = dados["etiquetas"]
        total = len(etiquetas)
        limite_preview = 80
        linhas_preview = [
            f"{item['volume']} -> {item['codigo_barras']}"
            for item in etiquetas[:limite_preview]
        ]
        if total > limite_preview:
            linhas_preview.append(
                f"... (mostrando {limite_preview} de {total} etiquetas)"
            )

        texto = (
            f"Destino: {dados['destino']}\n"
            f"Projeto: {dados['projeto']}\n"
            f"Romaneio: {dados['romaneio']}\n"
            f"NR NF: {dados['nr_nf']}\n"
            f"ID FEDEX: {dados['id_fedex_data']}\n"
            f"Quantidade de etiquetas: {total}\n\n"
            "Volumes / Codigos de barras:\n"
            + "\n".join(linhas_preview)
        )
        self.txt_preview.config(state="normal")
        self.txt_preview.delete("1.0", tk.END)
        self.txt_preview.insert("1.0", texto)
        self.txt_preview.config(state="disabled")

    def _atualizar_campo_codigo(self, dados: dict) -> None:
        etiquetas = dados["etiquetas"]
        primeiro = etiquetas[0]["codigo_barras"]
        if len(etiquetas) == 1:
            self.codigo_barras_var.set(primeiro)
        else:
            self.codigo_barras_var.set(f"{primeiro} (+{len(etiquetas) - 1})")

    def gerar_codigo(self) -> None:
        dados = self._coletar_dados()
        if not dados:
            return
        self._atualizar_campo_codigo(dados)
        self._atualizar_preview(dados)

    def _tamanho_etiqueta_points(self) -> tuple[float, float] | tuple[None, None]:
        largura_txt = self.etiqueta_largura_var.get().strip().replace(",", ".")
        altura_txt = self.etiqueta_altura_var.get().strip().replace(",", ".")
        try:
            largura_mm = float(largura_txt)
            altura_mm = float(altura_txt)
        except ValueError:
            messagebox.showerror("Tamanho invalido", "Informe largura/altura validas em mm.")
            return None, None

        if largura_mm <= 0 or altura_mm <= 0:
            messagebox.showerror(
                "Tamanho invalido", "Largura e altura da etiqueta devem ser maiores que zero."
            )
            return None, None
        return largura_mm * MM_TO_POINTS, altura_mm * MM_TO_POINTS

    def _ajustes_layout(self) -> tuple[float, float] | tuple[None, None]:
        espacamento_txt = self.espacamento_linhas_var.get().strip().replace(",", ".")
        escala_fonte_txt = self.escala_fonte_var.get().strip().replace(",", ".")
        try:
            espacamento_extra = float(espacamento_txt)
            escala_fonte = float(escala_fonte_txt)
        except ValueError:
            messagebox.showerror(
                "Ajuste invalido",
                "Informe valores validos para Espacamento e Escala fonte.",
            )
            return None, None

        if espacamento_extra < 0:
            messagebox.showerror("Ajuste invalido", "Espacamento deve ser maior ou igual a zero.")
            return None, None
        if escala_fonte <= 0:
            messagebox.showerror("Ajuste invalido", "Escala fonte deve ser maior que zero.")
            return None, None
        return espacamento_extra, escala_fonte

    def _desenhar_etiqueta_pdf(
        self,
        c: canvas.Canvas,
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

        borda = self._clamp(0.85 * escala_etiqueta, 0.6, 1.4)
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

        fonte_titulo = self._clamp(12 * escala_etiqueta * escala_fonte_usuario, 8, 28)
        fonte_label = self._clamp(9.2 * escala_etiqueta * escala_fonte_usuario, 6, 20)
        fonte_valor = self._clamp(9.8 * escala_etiqueta * escala_fonte_usuario, 6, 20)
        fonte_codigo = self._clamp(8.2 * escala_etiqueta * escala_fonte_usuario, 6, 16)
        fonte_identificador = self._clamp(
            7.0 * escala_etiqueta * escala_fonte_usuario, 5.5, 12
        )
        gap_linha = max(2.2, 1.8 * MM_TO_POINTS * escala_etiqueta) + espacamento_extra
        gap_bloco = max(4.0, 3 * MM_TO_POINTS * escala_etiqueta)

        altura_codigo = fonte_codigo * 1.45
        altura_identificador = fonte_identificador * 1.45
        altura_barcode = self._clamp(
            area_altura * 0.2, 10 * MM_TO_POINTS, area_altura * 0.26
        )
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
            if valor:
                c.setFont("Helvetica", fonte_valor)
                c.drawString(valor_x, y_linha, valor)
            y_linha -= passo_linha

        codigo = dados["codigo_barras"]
        largura_alvo = area_largura * 0.78
        modulos_estimados = max(80, (11 * len(codigo)) + 35)
        bar_width = self._clamp(largura_alvo / modulos_estimados, 0.16, 1.6)
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

        id_y = barcode_y + altura_barcode + gap_identificador_barra
        c.setFont("Helvetica", fonte_identificador)
        c.drawCentredString(area_x + (area_largura / 2), id_y, dados["id_fedex_data"])

        c.setFont("Helvetica", fonte_codigo)
        c.drawCentredString(area_x + (area_largura / 2), codigo_y, codigo)

    def _gerar_pdf_etiqueta(self, caminho_pdf: Path, dados_lote: dict) -> bool:
        if not REPORTLAB_AVAILABLE:
            messagebox.showerror(
                "Dependencia ausente",
                "Biblioteca reportlab nao encontrada.\n"
                "Instale com: pip install reportlab",
            )
            return False

        largura_pt, altura_pt = self._tamanho_etiqueta_points()
        if not largura_pt or not altura_pt:
            return False
        espacamento_extra, escala_fonte_usuario = self._ajustes_layout()
        if espacamento_extra is None or escala_fonte_usuario is None:
            return False

        try:
            caminho_pdf = Path(caminho_pdf)
            caminho_pdf.parent.mkdir(parents=True, exist_ok=True)

            _, page_h = A4
            margem = 12 * MM_TO_POINTS
            x = margem
            y = page_h - margem - altura_pt
            if y < margem:
                y = margem

            c = canvas.Canvas(str(caminho_pdf), pagesize=A4)
            c.setTitle(APP_NAME)

            for etiqueta in dados_lote["etiquetas"]:
                self._desenhar_etiqueta_pdf(
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
            return True
        except Exception as erro:
            messagebox.showerror(
                "Falha ao salvar PDF",
                "Nao foi possivel salvar o PDF no caminho selecionado.\n\n"
                f"Caminho: {caminho_pdf}\n"
                f"Detalhe: {erro}",
            )
            return False

    def salvar_pdf(self) -> None:
        dados = self._coletar_dados()
        if not dados:
            return
        self._atualizar_campo_codigo(dados)
        self._atualizar_preview(dados)

        caminho = filedialog.asksaveasfilename(
            title="Salvar etiqueta em PDF",
            defaultextension=".pdf",
            filetypes=[("Arquivos PDF", "*.pdf")],
            initialfile=f"etiqueta_{datetime.now():%Y%m%d_%H%M%S}.pdf",
        )
        if not caminho:
            return

        if self._gerar_pdf_etiqueta(Path(caminho), dados):
            messagebox.showinfo(
                "PDF gerado",
                f"Etiquetas salvas em:\n{caminho}\n\nQuantidade: {len(dados['etiquetas'])}",
            )

    def _carregar_impressoras(self) -> None:
        impressoras = [PRINTER_DEFAULT_LABEL]
        if PYWIN32_AVAILABLE:
            try:
                flags = (
                    win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
                )
                for impressora in win32print.EnumPrinters(flags):
                    nome = impressora[2]
                    if nome and nome not in impressoras:
                        impressoras.append(nome)
            except Exception:
                pass

        self.cmb_impressora["values"] = impressoras
        if self.impressora_var.get() not in impressoras:
            self.impressora_var.set(PRINTER_DEFAULT_LABEL)

    def imprimir(self) -> None:
        dados = self._coletar_dados()
        if not dados:
            return
        self._atualizar_campo_codigo(dados)
        self._atualizar_preview(dados)

        arquivo_temp = (
            Path(tempfile.gettempdir())
            / f"etiqueta_coleta_{datetime.now():%Y%m%d_%H%M%S}.pdf"
        )
        if not self._gerar_pdf_etiqueta(arquivo_temp, dados):
            return

        if not PYWIN32_AVAILABLE:
            messagebox.showwarning(
                "Impressao nao disponivel",
                "Biblioteca pywin32 nao encontrada.\n"
                f"O PDF foi gerado em:\n{arquivo_temp}\n\n"
                "Instale com: pip install pywin32",
            )
            return

        impressora = self.impressora_var.get()
        try:
            if impressora == PRINTER_DEFAULT_LABEL:
                win32api.ShellExecute(0, "print", str(arquivo_temp), None, ".", 0)
            else:
                win32api.ShellExecute(
                    0, "printto", str(arquivo_temp), f'"{impressora}"', ".", 0
                )
            messagebox.showinfo(
                "Impressao enviada",
                "Etiquetas enviadas para impressao.\n"
                f"Impressora: {impressora}\n"
                f"Quantidade: {len(dados['etiquetas'])}",
            )
        except Exception as erro:
            messagebox.showerror(
                "Falha na impressao",
                "Nao foi possivel enviar para impressao.\n"
                f"Detalhe: {erro}\n\n"
                f"PDF gerado em:\n{arquivo_temp}",
            )


def main() -> None:
    root = tk.Tk()
    EtiquetaColetaApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

