"""
Geração do relatório de auditoria em formato .docx.
Inclui gráficos de conformidade gerados com matplotlib.
"""
from __future__ import annotations

import io
import math
from datetime import date, datetime

import matplotlib
matplotlib.use("Agg")  # backend sem interface gráfica
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import pandas as pd
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.shared import Inches, Pt, RGBColor

from modules.excel_loader import AuditData
from modules.state import COL_CONFORMIDADE, COL_MOTIVO, COL_ACAO

# ---------------------------------------------------------------------------
# Paleta de cores
# ---------------------------------------------------------------------------
COR_CONFORME = "#2ecc71"
COR_NAO_CONFORME = "#e74c3c"
COR_NAO_AUDITADA = "#bdc3c7"
COR_TITULO = RGBColor(0x1A, 0x3A, 0x6A)
COR_HEADER_HEX = "1A3A6A"


# ---------------------------------------------------------------------------
# Helpers de formatação Word
# ---------------------------------------------------------------------------

def _fmt_date(dt: datetime | date | None) -> str:
    if dt is None:
        return "N/D"
    return dt.strftime("%d/%m/%Y") if isinstance(dt, (datetime, date)) else str(dt)


def _titulo(doc: Document, texto: str) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(texto)
    run.bold = True
    run.font.size = Pt(16)
    run.font.color.rgb = COR_TITULO


def _subtitulo(doc: Document, texto: str) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(texto)
    run.font.size = Pt(11)
    run.font.color.rgb = RGBColor(0x55, 0x55, 0x55)


def _heading(doc: Document, texto: str, level: int = 1) -> None:
    p = doc.add_heading(texto, level=level)
    if p.runs:
        p.runs[0].font.color.rgb = COR_TITULO


def _para(doc: Document, texto: str, bold: bool = False) -> None:
    p = doc.add_paragraph(texto)
    if bold and p.runs:
        p.runs[0].bold = True


def _set_cell_bg(cell, hex_color: str) -> None:
    ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    shading = parse_xml(f'<w:shd {ns} w:val="clear" w:color="auto" w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shading)


def _tabela_2col(
    doc: Document,
    linhas: list[tuple[str, str]],
    larguras: tuple[float, float] = (3.5, 3.5),
) -> None:
    table = doc.add_table(rows=len(linhas), cols=2)
    table.style = "Table Grid"
    for i, (label, valor) in enumerate(linhas):
        row = table.rows[i]
        row.cells[0].text = label
        row.cells[1].text = str(valor)
        row.cells[0].paragraphs[0].runs[0].bold = True
        row.cells[0].width = Inches(larguras[0])
        row.cells[1].width = Inches(larguras[1])


def _tabela_conformidade_header(table) -> None:
    hdr = table.rows[0]
    for cell in hdr.cells:
        _set_cell_bg(cell, COR_HEADER_HEX)
        if cell.paragraphs[0].runs:
            cell.paragraphs[0].runs[0].bold = True
            cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)


# ---------------------------------------------------------------------------
# Gráficos matplotlib
# ---------------------------------------------------------------------------

def _grafico_pizza(
    n_conf: int,
    n_nc: int,
    n_naud: int,
    titulo: str,
) -> io.BytesIO | None:
    """Gera gráfico de pizza de conformidade. Retorna None se não há dados auditados."""
    total_aud = n_conf + n_nc
    if total_aud == 0:
        return None

    labels, sizes, colors = [], [], []
    if n_conf > 0:
        labels.append(f"Conformes\n{n_conf} ({n_conf/total_aud*100:.1f}%)")
        sizes.append(n_conf)
        colors.append(COR_CONFORME)
    if n_nc > 0:
        labels.append(f"Não Conformes\n{n_nc} ({n_nc/total_aud*100:.1f}%)")
        sizes.append(n_nc)
        colors.append(COR_NAO_CONFORME)

    fig, ax = plt.subplots(figsize=(4.5, 3.5))
    wedges, texts = ax.pie(
        sizes, labels=labels, colors=colors, startangle=90,
        wedgeprops={"edgecolor": "white", "linewidth": 2},
    )
    for text in texts:
        text.set_fontsize(9)
    ax.set_title(titulo, fontsize=11, fontweight="bold", pad=12, color="#1A3A6A")

    if n_naud > 0:
        nota = f"* {n_naud} tarefa(s) não auditada(s) não incluída(s)"
        fig.text(0.5, 0.01, nota, ha="center", fontsize=7.5, color="#888888")

    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    buf.seek(0)
    return buf


def _grafico_barras_resumo(
    total_tarefas: int,
    total_triadas: int,
    total_nao_triadas: int,
    auditadas_triadas: int,
    auditadas_nao_triadas: int,
    conf_triadas: int,
    conf_nao_triadas: int,
) -> io.BytesIO:
    """Gráfico de barras com visão geral das estatísticas."""
    categorias = [
        "Total\nProcessadas",
        "Triadas\n(automação)",
        "Não Triadas",
        "Triadas\nAuditadas",
        "Não Triadas\nAuditadas",
    ]
    valores = [total_tarefas, total_triadas, total_nao_triadas,
               auditadas_triadas, auditadas_nao_triadas]
    cores = ["#3498db", "#2ecc71", "#e67e22", "#1abc9c", "#e74c3c"]

    fig, ax = plt.subplots(figsize=(7, 4))
    bars = ax.bar(categorias, valores, color=cores, edgecolor="white", linewidth=1.5)

    for bar, val in zip(bars, valores):
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            bar.get_height() + max(valores) * 0.01,
            str(val), ha="center", va="bottom", fontsize=10, fontweight="bold"
        )

    ax.set_ylabel("Quantidade de Tarefas", fontsize=10)
    ax.set_title("Visão Geral da Auditoria", fontsize=12, fontweight="bold", color="#1A3A6A")
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.set_ylim(0, max(valores) * 1.15)

    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    buf.seek(0)
    return buf


# ---------------------------------------------------------------------------
# Tabelas de detalhamento
# ---------------------------------------------------------------------------

def _tabela_nao_conformidades(
    doc: Document,
    df: pd.DataFrame,
    origem_label: str,
) -> None:
    df_nc = df[df[COL_CONFORMIDADE] == "Não Conforme"].copy()
    if df_nc.empty:
        _para(doc, f"Nenhuma não conformidade identificada nas tarefas {origem_label}.")
        return

    from modules.excel_loader import COL_TAREFA, COL_NUP
    headers = ["Tarefa", "NUP", "Motivo da Não Conformidade", "Ação Corretiva"]
    table = doc.add_table(rows=1 + len(df_nc), cols=4)
    table.style = "Table Grid"

    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = h
        _set_cell_bg(cell, COR_HEADER_HEX)
        if cell.paragraphs[0].runs:
            cell.paragraphs[0].runs[0].bold = True
            cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    for row_idx, (_, row) in enumerate(df_nc.iterrows(), start=1):
        r = table.rows[row_idx]
        r.cells[0].text = str(row.get(COL_TAREFA, ""))
        r.cells[1].text = str(row.get(COL_NUP, ""))
        r.cells[2].text = str(row.get(COL_MOTIVO, "") or "")
        r.cells[3].text = str(row.get(COL_ACAO, "") or "")

    for row in table.rows:
        row.cells[0].width = Inches(1.1)
        row.cells[1].width = Inches(1.5)
        row.cells[2].width = Inches(2.8)
        row.cells[3].width = Inches(2.8)


def _tabela_relacao_auditadas(
    doc: Document,
    df: pd.DataFrame,
    colunas_extras: list[str],
) -> None:
    """Lista completa das tarefas auditadas (Conformidade != 'Não auditada')."""
    from modules.excel_loader import COL_TAREFA, COL_NUP

    df_aud = df[df[COL_CONFORMIDADE] != "Não auditada"].copy()
    if df_aud.empty:
        _para(doc, "Nenhuma tarefa auditada.")
        return

    # Monta colunas: Tarefa, NUP, [extras], Conformidade
    cols_base = [COL_TAREFA, COL_NUP] + [c for c in colunas_extras if c in df.columns]
    cols_show = cols_base + [COL_CONFORMIDADE]
    headers_map = {
        COL_TAREFA: "Tarefa",
        COL_NUP: "NUP",
        COL_CONFORMIDADE: "Conformidade",
    }
    headers = [headers_map.get(c, c) for c in cols_show]

    n_cols = len(cols_show)
    table = doc.add_table(rows=1 + len(df_aud), cols=n_cols)
    table.style = "Table Grid"

    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = h
        _set_cell_bg(cell, COR_HEADER_HEX)
        if cell.paragraphs[0].runs:
            cell.paragraphs[0].runs[0].bold = True
            cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    for row_idx, (_, row) in enumerate(df_aud.iterrows(), start=1):
        r = table.rows[row_idx]
        for col_idx, col in enumerate(cols_show):
            r.cells[col_idx].text = str(row.get(col, "") or "")
        # Colorir conformidade
        conf = str(row.get(COL_CONFORMIDADE, ""))
        if conf == "Conforme":
            _set_cell_bg(r.cells[-1], "d5f5e3")
        elif conf == "Não Conforme":
            _set_cell_bg(r.cells[-1], "fadbd8")


def _section_auditoria(
    doc: Document,
    numero: str,
    titulo: str,
    df: pd.DataFrame | None,
    tipo_controle: str | None,
    tamanho_amostra: int | None,
    colunas_extras: list[str],
    origem_label: str,
) -> None:
    """Renderiza seção completa de auditoria (triadas ou não-triadas)."""
    _heading(doc, f"{numero}. {titulo}")

    if df is None or df.empty:
        _para(doc, "Nenhuma tarefa selecionada para auditoria neste ciclo.")
        doc.add_paragraph()
        return

    from modules.state import stats_df
    s = stats_df(df)
    n_naud = s["total"] - s["auditadas"]

    # Subseção: resultado quantitativo
    _heading(doc, f"{numero}.1 Resultado Quantitativo", level=2)
    linhas_stats = [
        ("Total de tarefas disponíveis", str(s["total"])),
        ("Tarefas auditadas", str(s["auditadas"])),
        ("Tarefas não auditadas (excluídas das estatísticas)", str(n_naud)),
        ("Conformes", f"{s['conformes']} ({s['pct_conf']:.1f}%)"),
        ("Não Conformes", f"{s['nao_conformes']} ({s['pct_nc']:.1f}%)"),
    ]
    if tipo_controle == "detalhado" and tamanho_amostra is not None and numero == "4":
        linhas_stats.insert(1, ("Amostra definida pela fórmula", str(tamanho_amostra)))
    _tabela_2col(doc, linhas_stats)
    doc.add_paragraph()

    # Gráfico de pizza
    grafico = _grafico_pizza(s["conformes"], s["nao_conformes"], n_naud,
                             f"Conformidade — {titulo}")
    if grafico:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        run.add_picture(grafico, width=Inches(4.0))
    doc.add_paragraph()

    # Subseção: não conformidades
    _heading(doc, f"{numero}.2 Detalhamento das Não Conformidades", level=2)
    _tabela_nao_conformidades(doc, df, origem_label)
    doc.add_paragraph()

    # Subseção: relação de auditadas
    _heading(doc, f"{numero}.3 Relação de Tarefas Auditadas", level=2)
    _tabela_relacao_auditadas(doc, df, colunas_extras)
    doc.add_paragraph()


# ---------------------------------------------------------------------------
# Geração do texto de conclusão
# ---------------------------------------------------------------------------

def _conclusao(df_tri: pd.DataFrame | None, df_nao: pd.DataFrame | None) -> str:
    from modules.state import stats_df
    s_tri = stats_df(df_tri)
    s_nao = stats_df(df_nao)

    total_aud = s_tri["auditadas"] + s_nao["auditadas"]
    total_nc = s_tri["nao_conformes"] + s_nao["nao_conformes"]

    if total_aud == 0:
        return "Nenhuma tarefa foi auditada neste ciclo."

    pct_conf = (total_aud - total_nc) / total_aud * 100

    partes = [
        f"No presente ciclo de auditoria foram examinadas {total_aud} tarefa(s), "
        f"sendo {s_tri['auditadas']} tarefa(s) triada(s) e "
        f"{s_nao['auditadas']} tarefa(s) não triada(s). "
    ]

    if total_nc == 0:
        partes.append(
            "Não foram identificadas não conformidades, demonstrando adequação "
            "das regras de negócio e dos fluxos de automação do Conecta+. "
        )
    else:
        partes.append(
            f"Foram identificadas {total_nc} não conformidade(s) "
            f"({100 - pct_conf:.1f}% do total auditado). "
            "As respectivas ações corretivas foram registradas nas seções anteriores "
            "e devem ser implementadas e verificadas no próximo ciclo. "
        )

    partes.append(
        "Recomenda-se a manutenção do controle de qualidade periódico, "
        "o registro dos resultados em NUP próprio com responsáveis e periodicidade definidos, "
        "e a revisão contínua das regras de negócio, conforme preconiza o "
        "Manual de Gerenciamento Estratégico de Contencioso (Portaria PGF/AGU n. 541/2025, seção 5)."
    )

    return " ".join(partes)


# ---------------------------------------------------------------------------
# Função principal
# ---------------------------------------------------------------------------

def gerar_relatorio(
    audit_data: AuditData,
    df_triadas: pd.DataFrame | None,
    df_nao_triadas: pd.DataFrame | None,
    tipo_controle: str | None,
    tamanho_amostra: int | None,
    responsavel: str,
    data_auditoria: date,
) -> bytes:
    """
    Gera o relatório de auditoria e retorna os bytes do arquivo .docx.
    Não realiza escrita em disco.
    """
    from modules.excel_loader import COL_CONFIG, COL_STATUS
    from modules.state import stats_df

    doc = Document()

    # Margens
    for section in doc.sections:
        section.top_margin = Inches(0.9)
        section.bottom_margin = Inches(0.9)
        section.left_margin = Inches(1.2)
        section.right_margin = Inches(1.2)

    # -----------------------------------------------------------------------
    # Cabeçalho
    # -----------------------------------------------------------------------
    _titulo(doc, "RELATÓRIO DE AUDITORIA")
    _subtitulo(doc, "Conecta+ Automação — Controle de Qualidade da Triagem")
    _subtitulo(doc, "Procuradoria-Geral Federal / Advocacia-Geral da União")
    doc.add_paragraph()

    # -----------------------------------------------------------------------
    # 1. Identificação
    # -----------------------------------------------------------------------
    _heading(doc, "1. IDENTIFICAÇÃO")
    periodo = (
        f"{_fmt_date(audit_data.periodo_inicio)} a {_fmt_date(audit_data.periodo_fim)}"
        if audit_data.periodo_inicio else "N/D"
    )
    _tabela_2col(doc, [
        ("Período Auditado", periodo),
        ("Data de Emissão do Relatório", _fmt_date(data_auditoria)),
        ("Responsável pela Auditoria", responsavel or "Não informado"),
        ("Sistema Auditado", "Conecta+ Automação — Módulo de Triagem Avançada"),
        ("Arquivo(s) Analisado(s)", audit_data.nome_arquivo),
        ("Base Normativa", "Portaria PGF/AGU n. 541/2025 — Manual de Gerenciamento Estratégico"),
    ])
    doc.add_paragraph()

    # -----------------------------------------------------------------------
    # 2. Estatísticas Gerais da Triagem
    # -----------------------------------------------------------------------
    _heading(doc, "2. ESTATÍSTICAS GERAIS DA TRIAGEM")
    _tabela_2col(doc, [
        ("Total de Tarefas Processadas pelo Sistema", str(audit_data.total_tarefas)),
        ("Tarefas Triadas (com configurações encontradas)", f"{audit_data.total_triadas} ({audit_data.pct_triadas:.1f}%)"),
        ("Tarefas Não Triadas", f"{audit_data.total_nao_triadas} ({audit_data.pct_nao_triadas:.1f}%)"),
    ])
    doc.add_paragraph()

    # Gráfico de barras visão geral
    s_tri = stats_df(df_triadas)
    s_nao = stats_df(df_nao_triadas)
    grafico_geral = _grafico_barras_resumo(
        audit_data.total_tarefas,
        audit_data.total_triadas,
        audit_data.total_nao_triadas,
        s_tri["auditadas"],
        s_nao["auditadas"],
        s_tri["conformes"],
        s_nao["conformes"],
    )
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(grafico_geral, width=Inches(5.5))
    doc.add_paragraph()

    # -----------------------------------------------------------------------
    # 3. Metodologia
    # -----------------------------------------------------------------------
    _heading(doc, "3. METODOLOGIA DO CONTROLE DE QUALIDADE")

    tipo_label = {
        "simplificado": "Controle Simplificado",
        "detalhado": "Controle Detalhado (Amostragem Estatística)",
    }.get(tipo_controle or "", tipo_controle or "Não informado")

    linhas_met = [
        ("Tipo de Controle (Tarefas Triadas)", tipo_label),
        ("Nível de Confiança", "95%"),
        ("Margem de Erro", "5%"),
    ]

    if tipo_controle == "detalhado" and tamanho_amostra is not None:
        linhas_met += [
            ("Universo Amostral", f"{audit_data.total_triadas} tarefas triadas"),
            ("Tamanho da Amostra (fórmula)", f"{tamanho_amostra} tarefas"),
            ("Fórmula Aplicada", "n = n₀ / (1 + (n₀ - 1) / N), onde n₀ = Z² · p · (1-p) / E²"),
            ("Parâmetros", "Z = 1,96 | p = 0,50 | E = 0,05"),
            ("Seleção", "Aleatória simples sem reposição"),
        ]
    if s_nao["auditadas"] > 0:
        linhas_met.append(
            ("Tarefas Não Triadas Auditadas",
             f"{s_nao['auditadas']} de {audit_data.total_nao_triadas} disponíveis")
        )

    _tabela_2col(doc, linhas_met)
    doc.add_paragraph()

    # -----------------------------------------------------------------------
    # 4. Auditoria das Tarefas Triadas
    # -----------------------------------------------------------------------
    _section_auditoria(
        doc=doc,
        numero="4",
        titulo="AUDITORIA DAS TAREFAS TRIADAS",
        df=df_triadas,
        tipo_controle=tipo_controle,
        tamanho_amostra=tamanho_amostra,
        colunas_extras=[COL_CONFIG],
        origem_label="triadas",
    )

    # -----------------------------------------------------------------------
    # 5. Auditoria das Tarefas Não Triadas
    # -----------------------------------------------------------------------
    _section_auditoria(
        doc=doc,
        numero="5",
        titulo="AUDITORIA DAS TAREFAS NÃO TRIADAS",
        df=df_nao_triadas,
        tipo_controle=None,
        tamanho_amostra=None,
        colunas_extras=[COL_STATUS],
        origem_label="não triadas",
    )

    # -----------------------------------------------------------------------
    # 6. Conclusão
    # -----------------------------------------------------------------------
    _heading(doc, "6. CONCLUSÃO")
    _para(doc, _conclusao(df_triadas, df_nao_triadas))
    doc.add_paragraph()

    # Assinatura
    doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run(f"Brasília, {_fmt_date(data_auditoria)}").italic = True

    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = p2.add_run("_" * 50)
    r2.bold = False

    p3 = doc.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r3 = p3.add_run(responsavel or "Responsável pela Auditoria")
    r3.bold = True

    # -----------------------------------------------------------------------
    # Salvar em memória
    # -----------------------------------------------------------------------
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()
