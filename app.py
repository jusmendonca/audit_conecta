"""
Auditoria Conecta+ — Aplicação principal Streamlit.
Execução: streamlit run app.py
"""
from __future__ import annotations

from datetime import date as date_type

import pandas as pd
import streamlit as st

from modules.excel_loader import (
    COL_CONFIG, COL_NUP, COL_STATUS, COL_TAREFA, COL_USUARIO,
    load_file, merge_audit_data,
)
from modules.sampling import (
    calcular_amostra, formula_descricao, selecionar_amostra, tabela_referencia,
)
from modules.state import (
    COL_ACAO, COL_CONFORMIDADE, COL_MOTIVO,
    OPCOES_CONFORMIDADE,
    get_audit_data, get_df_nao_triadas, get_df_triadas,
    init_state, preparar_df_auditoria, reset_auditoria, stats_df,
)
from modules.report import gerar_relatorio

# ---------------------------------------------------------------------------
# Configuração da página
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Auditoria Conecta+",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
)

init_state()

# ---------------------------------------------------------------------------
# CSS
# ---------------------------------------------------------------------------
st.markdown("""
<style>
    .block-container { padding-top: 1.2rem; padding-bottom: 1rem; }
    div[data-testid="stSidebarNav"] { display: none; }

    /* Cartão de tarefa */
    .task-card {
        border: 1px solid #dde3ee;
        border-left: 5px solid #1A3A6A;
        border-radius: 6px;
        padding: 0.6rem 1rem;
        margin-bottom: 0.4rem;
        background: #f8f9fb;
    }
    .task-card.conforme  { border-left-color: #27ae60; background: #f0faf4; }
    .task-card.nc        { border-left-color: #e74c3c; background: #fdf4f4; }

    .task-num    { font-size: 0.75rem; color: #888; }
    .task-title  { font-size: 0.97rem; font-weight: 700; color: #1A3A6A; margin: 1px 0; }
    .task-sub    { font-size: 0.82rem; color: #555; }
    .task-config { font-size: 0.78rem; color: #777; font-style: italic; }

    /* Banner de período */
    .periodo-box {
        background: #eaf1fb;
        border-left: 4px solid #1A3A6A;
        border-radius: 4px;
        padding: 0.5rem 1rem;
        margin: 0.4rem 0 0.8rem 0;
        font-size: 0.95rem;
    }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Sidebar — Navegação
# ---------------------------------------------------------------------------
PAGINAS = {
    "importacao":  ("📂", "1. Importação"),
    "triadas":     ("✅", "2. Tarefas Triadas"),
    "nao_triadas": ("🔍", "3. Tarefas Não Triadas"),
    "relatorio":   ("📄", "4. Relatório"),
}


def _check_icon(chave: str) -> str:
    checks = {
        "importacao":  st.session_state.get("audit_data_merged") is not None,
        "triadas":     st.session_state.get("auditoria_triadas_concluida", False),
        "nao_triadas": st.session_state.get("auditoria_nao_triadas_concluida", False),
        "relatorio":   False,
    }
    return "  ✓" if checks.get(chave) else ""


with st.sidebar:
    st.markdown("### 📋 Auditoria Conecta+")
    st.caption("Procuradoria-Geral Federal / AGU")
    st.divider()

    pagina_atual = st.session_state.get("pagina", "importacao")
    for chave, (icone, label) in PAGINAS.items():
        check = _check_icon(chave)
        btn_label = f"{icone} {label}{check}"
        if pagina_atual == chave:
            st.markdown(f"**{btn_label}**")
        else:
            if st.button(btn_label, key=f"nav_{chave}", use_container_width=True):
                st.session_state["pagina"] = chave
                st.rerun()

    st.divider()

    ad = get_audit_data()
    if ad:
        st.caption(f"📁 {ad.nome_arquivo}")
        st.caption(
            f"Total: **{ad.total_tarefas}** · "
            f"Triadas: **{ad.total_triadas}** · "
            f"Não triadas: **{ad.total_nao_triadas}**"
        )
        st.divider()

    if st.button("🔄 Nova Auditoria", use_container_width=True):
        for k in list(st.session_state.keys()):
            if k.startswith((
                "conf_tri_", "motivo_tri_", "acao_tri_",
                "conf_nao_", "motivo_nao_", "acao_nao_",
                "pag_tri", "pag_nao",
            )):
                del st.session_state[k]
        reset_auditoria()
        st.session_state["pagina"] = "importacao"
        st.session_state["audit_data_merged"] = None
        st.rerun()


# ---------------------------------------------------------------------------
# Helpers: session_state
# ---------------------------------------------------------------------------

def _inicializar_chaves(prefixo: str, df: pd.DataFrame) -> None:
    """Cria chaves de widget no session_state para cada linha do df (apenas se não existirem)."""
    for i in range(len(df)):
        for prefixo_col, col in [
            (f"conf_{prefixo}_{i}",   COL_CONFORMIDADE),
            (f"motivo_{prefixo}_{i}", COL_MOTIVO),
            (f"acao_{prefixo}_{i}",   COL_ACAO),
        ]:
            if prefixo_col not in st.session_state:
                st.session_state[prefixo_col] = df.loc[i, col]


def _sincronizar_para_df(prefixo: str, df_key: str) -> None:
    """Lê as chaves de widget e atualiza o DataFrame em session_state."""
    df = st.session_state.get(df_key)
    if df is None:
        return
    df = df.copy()
    for i in range(len(df)):
        df.at[i, COL_CONFORMIDADE] = st.session_state.get(
            f"conf_{prefixo}_{i}", OPCOES_CONFORMIDADE[0]
        )
        df.at[i, COL_MOTIVO] = st.session_state.get(f"motivo_{prefixo}_{i}", "")
        df.at[i, COL_ACAO]   = st.session_state.get(f"acao_{prefixo}_{i}", "")
    st.session_state[df_key] = df


def _stats_chaves(prefixo: str, total: int) -> dict:
    """Calcula estatísticas diretamente das chaves de widget (sempre atualizado)."""
    conformes = nc = auditadas = 0
    for i in range(total):
        v = st.session_state.get(f"conf_{prefixo}_{i}", OPCOES_CONFORMIDADE[0])
        if v != OPCOES_CONFORMIDADE[0]:
            auditadas += 1
            if v == "Conforme":
                conformes += 1
            elif v == "Não Conforme":
                nc += 1
    pct_conf = (conformes / auditadas * 100) if auditadas > 0 else 0.0
    pct_nc   = (nc / auditadas * 100) if auditadas > 0 else 0.0
    return {
        "total": total, "auditadas": auditadas,
        "conformes": conformes, "nao_conformes": nc,
        "pct_conf": pct_conf, "pct_nc": pct_nc,
    }


# ---------------------------------------------------------------------------
# Helper: cartões de auditoria
# ---------------------------------------------------------------------------

TAREFAS_POR_PAGINA = 10


def _cor_card(conf: str) -> str:
    if conf == "Conforme":
        return "conforme"
    if conf == "Não Conforme":
        return "nc"
    return ""


def _render_cartoes(
    prefixo: str,
    df: pd.DataFrame,
    df_key: str,
    mostrar_config: bool = True,
) -> None:
    """
    Renderiza cartões interativos para cada tarefa.
    - Campos de motivo/ação são sempre renderizados (disabled quando não aplicável),
      garantindo que session_state nunca perde os valores entre trocas de página.
    - Sync é feito ANTES de qualquer st.rerun() de paginação para evitar perda de dados.
    """
    total = len(df)
    pag_key = f"pag_{prefixo}"
    if pag_key not in st.session_state:
        st.session_state[pag_key] = 0

    n_pag = max(1, (total + TAREFAS_POR_PAGINA - 1) // TAREFAS_POR_PAGINA)
    if st.session_state[pag_key] >= n_pag:
        st.session_state[pag_key] = n_pag - 1

    inicio = st.session_state[pag_key] * TAREFAS_POR_PAGINA
    fim    = min(inicio + TAREFAS_POR_PAGINA, total)

    for i in range(inicio, fim):
        row        = df.loc[i]
        ck         = f"conf_{prefixo}_{i}"
        mk         = f"motivo_{prefixo}_{i}"
        ak         = f"acao_{prefixo}_{i}"
        conf_atual = st.session_state.get(ck, OPCOES_CONFORMIDADE[0])
        cor        = _cor_card(conf_atual)

        tarefa  = str(row.get(COL_TAREFA, "—"))
        nup     = str(row.get(COL_NUP, "—"))
        config  = str(row.get(COL_CONFIG, "")) if mostrar_config else ""
        usuario = str(row.get(COL_USUARIO, ""))

        st.markdown(
            f'<div class="task-card {cor}">'
            f'<div class="task-num">Tarefa {i + 1} de {total}'
            + (f" &nbsp;|&nbsp; {usuario}" if usuario else "")
            + "</div>"
            f'<div class="task-title">{tarefa}</div>'
            f'<div class="task-sub">NUP: {nup}</div>'
            + (f'<div class="task-config">Configuração: {config}</div>' if config else "")
            + "</div>",
            unsafe_allow_html=True,
        )

        col_rad, col_texto = st.columns([1, 2])

        with col_rad:
            idx_atual = (
                OPCOES_CONFORMIDADE.index(conf_atual)
                if conf_atual in OPCOES_CONFORMIDADE else 0
            )
            st.radio(
                f"Resultado (tarefa {i + 1}):",
                OPCOES_CONFORMIDADE,
                index=idx_atual,
                key=ck,
                label_visibility="collapsed",
            )

        # Campos de motivo/ação SEMPRE renderizados (disabled quando inaplicável).
        # Isso garante que session_state[mk] e session_state[ak] persistam mesmo
        # quando a tarefa não está marcada como "Não Conforme".
        is_nc = st.session_state.get(ck) == "Não Conforme"
        with col_texto:
            st.text_area(
                "Motivo da não conformidade:",
                key=mk,
                height=90,
                placeholder="Descreva o motivo…" if is_nc else "Preencha apenas se Não Conforme",
                disabled=not is_nc,
            )
            st.text_area(
                "Ação corretiva proposta:",
                key=ak,
                height=90,
                placeholder="Descreva a ação corretiva…" if is_nc else "Preencha apenas se Não Conforme",
                disabled=not is_nc,
            )

        st.divider()

    # Paginação — sync ANTES do rerun para nunca perder dados
    if n_pag > 1:
        col_ant, col_inf, col_prox = st.columns([1, 3, 1])
        with col_ant:
            if st.button(
                "← Anterior",
                disabled=st.session_state[pag_key] == 0,
                key=f"btn_ant_{prefixo}",
            ):
                _sincronizar_para_df(prefixo, df_key)
                st.session_state[pag_key] -= 1
                st.rerun()
        with col_inf:
            st.markdown(
                f"<div style='text-align:center;padding-top:0.5rem;'>"
                f"Página **{st.session_state[pag_key] + 1}** de {n_pag}"
                f" &nbsp;·&nbsp; tarefas {inicio + 1}–{fim} de {total}"
                f"</div>",
                unsafe_allow_html=True,
            )
        with col_prox:
            if st.button(
                "Próxima →",
                disabled=st.session_state[pag_key] >= n_pag - 1,
                key=f"btn_prox_{prefixo}",
            ):
                _sincronizar_para_df(prefixo, df_key)
                st.session_state[pag_key] += 1
                st.rerun()


def _barra_progresso(prefixo: str, total: int, label: str = "") -> None:
    s = _stats_chaves(prefixo, total)
    pct = s["auditadas"] / total if total > 0 else 0
    st.progress(
        pct,
        text=(
            f"{label}**{s['auditadas']}/{total}** registradas"
            f" · {s['conformes']} conformes · {s['nao_conformes']} não conformes"
        ),
    )


# ===========================================================================
# PÁGINA 1 — IMPORTAÇÃO
# ===========================================================================

def render_importacao() -> None:
    st.title("📂 Importação de Arquivo")
    st.caption(
        "Importe a planilha Excel gerada pelo módulo de Triagem Avançada do Conecta+ Automação."
    )

    col_up, col_info = st.columns([2, 1])
    with col_up:
        uploaded = st.file_uploader(
            "Selecione o(s) arquivo(s) Excel (.xlsx):",
            type=["xlsx"],
            accept_multiple_files=True,
            help="O arquivo deve conter as abas: Todas as Tarefas, Tarefas Triadas e Tarefas Não Triadas.",
        )
    with col_info:
        st.markdown("""
        **Formato esperado:**
        - Aba 1: Todas as Tarefas
        - Aba 2: Tarefas Triadas
        - Aba 3: Tarefas Não Triadas

        Colunas: ID, Tarefa, NUP, Usuário,
        Datas, Status, Configurações Encontradas
        """)

    if not uploaded:
        if get_audit_data() is not None:
            ad = get_audit_data()
            st.info(
                f"Arquivo **{ad.nome_arquivo}** já carregado. "
                "Navegue pelas etapas no menu lateral."
            )
        else:
            st.info("Selecione um ou mais arquivos Excel para começar.")
        return

    # Processar arquivos
    audit_files, erros = [], []
    for f in uploaded:
        try:
            audit_files.append(load_file(f, f.name))
        except ValueError as e:
            erros.append(str(e))

    if erros:
        for e in erros:
            st.error(e)
        return

    merged = merge_audit_data(audit_files) if len(audit_files) > 1 else audit_files[0]

    if len(audit_files) > 1:
        st.info(
            f"{len(audit_files)} arquivos consolidados. "
            "Registros duplicados foram removidos (mantida a ocorrência mais recente)."
        )

    atual = st.session_state.get("audit_data_merged")
    if atual is not None and atual.nome_arquivo != merged.nome_arquivo:
        reset_auditoria()

    st.session_state["audit_data_merged"] = merged

    # ---- Período ----
    st.divider()
    st.subheader("Resumo da Triagem")

    if merged.periodo_inicio and merged.periodo_fim:
        periodo_str = (
            f"{merged.periodo_inicio.strftime('%d/%m/%Y %H:%M')} "
            f"até {merged.periodo_fim.strftime('%d/%m/%Y %H:%M')}"
        )
    else:
        periodo_str = "Período não identificado"

    st.markdown(
        f'<div class="periodo-box">📅 <strong>Período de triagem:</strong> {periodo_str}</div>',
        unsafe_allow_html=True,
    )

    # ---- Métricas ----
    c1, c2, c3 = st.columns(3)
    c1.metric("Total de Tarefas", merged.total_tarefas)
    c2.metric(
        "Tarefas Triadas",
        merged.total_triadas,
        delta=f"{merged.pct_triadas:.1f}% do total",
        delta_color="normal",
    )
    c3.metric(
        "Tarefas Não Triadas",
        merged.total_nao_triadas,
        delta=f"{merged.pct_nao_triadas:.1f}% do total",
        delta_color="inverse",
    )

    # ---- Pré-visualização ----
    st.divider()
    tab1, tab2 = st.tabs([
        f"Tarefas Triadas ({merged.total_triadas})",
        f"Tarefas Não Triadas ({merged.total_nao_triadas})",
    ])

    cols_tri = [c for c in [COL_TAREFA, COL_NUP, COL_USUARIO, COL_STATUS, COL_CONFIG]
                if c in merged.triadas.columns]
    cols_nao = [c for c in [COL_TAREFA, COL_NUP, COL_USUARIO, COL_STATUS]
                if c in merged.nao_triadas.columns]

    with tab1:
        st.dataframe(merged.triadas[cols_tri], hide_index=True, use_container_width=True)
    with tab2:
        st.dataframe(merged.nao_triadas[cols_nao], hide_index=True, use_container_width=True)

    st.divider()
    if st.button("Iniciar Auditoria →", type="primary"):
        st.session_state["pagina"] = "triadas"
        st.rerun()


# ===========================================================================
# PÁGINA 2 — AUDITORIA DAS TAREFAS TRIADAS
# ===========================================================================

def render_auditoria_triadas() -> None:
    audit_data = get_audit_data()
    if audit_data is None:
        st.warning("Nenhum arquivo carregado. Volte à página de Importação.")
        return

    st.title("✅ Auditoria das Tarefas Triadas")

    # ---- Seleção do tipo de controle ----
    if st.session_state.get("tipo_controle") is None:

        st.markdown(
            f"**{audit_data.total_triadas}** tarefas triadas disponíveis "
            f"({audit_data.pct_triadas:.1f}% do total). "
            "Selecione o tipo de controle conforme o Manual de Gerenciamento Estratégico "
            "de Contencioso (Portaria PGF/AGU n. 541/2025, seção 5)."
        )
        st.divider()

        col_esq, col_dir = st.columns([1, 1])

        with col_esq:
            st.markdown("#### Tipo de Controle")
            tipo = st.radio(
                "Selecione:",
                ["Controle Simplificado", "Controle Detalhado (Amostragem Estatística)"],
                key="radio_tipo",
                label_visibility="collapsed",
            )
            st.markdown("""
            **Controle Simplificado** — Verificação manual das tarefas selecionadas
            pelo auditor no SuperSapiens. Indicado para fluxos bem estruturados.
            Periodicidade recomendada: **diária** (Manual, seção 5.1).

            **Controle Detalhado** — Amostragem estatística com seleção aleatória.
            Nível de confiança **95%**, margem de erro **±5%**.
            Indicado para análise mais rigorosa de conformidade.
            """)

        with col_dir:
            st.markdown("#### Tamanho da Amostra (Controle Detalhado)")
            t = audit_data.total_triadas
            if t > 0:
                n = calcular_amostra(t)
                st.metric("Tarefas a auditar", n, delta=f"{n / t * 100:.1f}% do universo")
                st.markdown(formula_descricao(t))

                with st.expander("📊 Tabela de Referência — Anexo III do Manual"):
                    df_ref = pd.DataFrame(
                        tabela_referencia(), columns=["Universo (N)", "Amostra (n)"]
                    )
                    df_ref["Calculado pela fórmula"] = df_ref["Universo (N)"].apply(calcular_amostra)
                    st.dataframe(df_ref, hide_index=True, use_container_width=True)

        st.divider()
        if st.button("Confirmar e Iniciar Auditoria →", type="primary"):
            chave = "simplificado" if tipo.startswith("Controle S") else "detalhado"
            st.session_state["tipo_controle"] = chave

            if chave == "detalhado":
                n = calcular_amostra(audit_data.total_triadas)
                st.session_state["tamanho_amostra"] = n
                df_base = selecionar_amostra(audit_data.triadas, n)
            else:
                st.session_state["tamanho_amostra"] = None
                df_base = audit_data.triadas.copy()

            colunas = [COL_TAREFA, COL_NUP, COL_USUARIO, COL_CONFIG, COL_STATUS]
            df_prep = preparar_df_auditoria(df_base, colunas)
            st.session_state["df_audit_triadas"] = df_prep

            # Limpar e inicializar chaves de widget
            for k in list(st.session_state.keys()):
                if k.startswith(("conf_tri_", "motivo_tri_", "acao_tri_", "pag_tri")):
                    del st.session_state[k]
            _inicializar_chaves("tri", df_prep)
            st.rerun()
        return

    # ---- Editor de auditoria ----
    tipo_controle = st.session_state["tipo_controle"]
    tipo_label    = "Controle Simplificado" if tipo_controle == "simplificado" else "Controle Detalhado"
    n_amostra     = st.session_state.get("tamanho_amostra")
    df            = st.session_state.get("df_audit_triadas")

    if df is None:
        st.error("Estado inconsistente. Clique em 'Nova Auditoria' no menu lateral.")
        return

    total = len(df)

    # Garantir que chaves existem para todos os itens
    _inicializar_chaves("tri", df)

    # Cabeçalho
    s = _stats_chaves("tri", total)
    col_a, col_b, col_c = st.columns([2, 1, 1])
    with col_a:
        descr = f"Amostra: **{n_amostra}** tarefas" if n_amostra else f"Total: **{total}** tarefas"
        st.markdown(f"**Tipo:** {tipo_label} · {descr}")
    with col_b:
        st.metric("Auditadas", f"{s['auditadas']}/{total}")
    with col_c:
        if s["auditadas"] > 0:
            st.metric("Conformidade", f"{s['pct_conf']:.1f}%")

    _barra_progresso("tri", total)

    if tipo_controle == "simplificado":
        st.info(
            "**Controle Simplificado:** Verifique as tarefas no SuperSapiens pelo NUP e "
            "registre o resultado abaixo. Tarefas *Não auditadas* não entram nas estatísticas.",
            icon="ℹ️",
        )
    else:
        st.info(
            f"**Controle Detalhado:** {n_amostra} tarefas selecionadas aleatoriamente. "
            "Verifique cada uma no SuperSapiens e registre o resultado.",
            icon="ℹ️",
        )

    # Renderizar cartões — sincroniza dados antes de qualquer rerun interno
    _render_cartoes("tri", df, df_key="df_audit_triadas", mostrar_config=True)

    # Sincronizar session_state → DataFrame após renderização (rerun normal)
    _sincronizar_para_df("tri", "df_audit_triadas")

    st.divider()
    col1, col2 = st.columns([2, 1])
    with col1:
        if st.button("Concluir e Avançar para Tarefas Não Triadas →", type="primary"):
            _sincronizar_para_df("tri", "df_audit_triadas")
            st.session_state["auditoria_triadas_concluida"] = True
            st.session_state["pagina"] = "nao_triadas"
            st.rerun()
    with col2:
        if st.button("↩ Trocar Tipo de Controle"):
            st.session_state["tipo_controle"]             = None
            st.session_state["df_audit_triadas"]          = None
            st.session_state["tamanho_amostra"]           = None
            st.session_state["auditoria_triadas_concluida"] = False
            for k in list(st.session_state.keys()):
                if k.startswith(("conf_tri_", "motivo_tri_", "acao_tri_", "pag_tri")):
                    del st.session_state[k]
            st.rerun()


# ===========================================================================
# PÁGINA 3 — AUDITORIA DAS TAREFAS NÃO TRIADAS
# ===========================================================================

def render_auditoria_nao_triadas() -> None:
    audit_data = get_audit_data()
    if audit_data is None:
        st.warning("Nenhum arquivo carregado. Volte à página de Importação.")
        return

    st.title("🔍 Auditoria das Tarefas Não Triadas")
    st.markdown(
        f"**{audit_data.total_nao_triadas}** tarefas não triadas disponíveis "
        f"({audit_data.pct_nao_triadas:.1f}% do total). "
        "Selecione as tarefas a auditar e registre o resultado."
    )

    # ---- Seleção ----
    if st.session_state.get("df_audit_nao_triadas") is None:
        st.divider()
        st.subheader("Seleção das Tarefas")

        nao_triadas = audit_data.nao_triadas
        cols_show = [c for c in [COL_TAREFA, COL_NUP, COL_USUARIO, COL_STATUS]
                     if c in nao_triadas.columns]
        st.dataframe(nao_triadas[cols_show], hide_index=True, use_container_width=True)
        st.divider()

        col_sel, col_opt = st.columns([1, 1])
        with col_sel:
            modo = st.radio(
                "Quais tarefas deseja auditar?",
                ["Todas as tarefas não triadas", "Seleção manual"],
                key="modo_nao_triadas",
            )
        with col_opt:
            ids_labels = [
                f"{row.get(COL_TAREFA, '')} | {row.get(COL_NUP, '')}"
                for row in nao_triadas.to_dict("records")
            ]
            sel_manual: list[str] = []
            if modo == "Seleção manual":
                sel_manual = st.multiselect(
                    "Selecione as tarefas:",
                    options=ids_labels,
                    key="multisel_nao_triadas",
                    placeholder="Digite para filtrar…",
                )

        if st.button("Abrir Editor de Auditoria →", type="primary"):
            if modo == "Todas as tarefas não triadas":
                df_base = nao_triadas.copy()
            else:
                if not sel_manual:
                    st.error("Selecione ao menos uma tarefa.")
                    return
                ids_sel = {lbl.split(" | ")[0] for lbl in sel_manual}
                df_base = nao_triadas[
                    nao_triadas[COL_TAREFA].astype(str).isin(ids_sel)
                ].copy()

            colunas = [COL_TAREFA, COL_NUP, COL_USUARIO, COL_STATUS]
            df_prep = preparar_df_auditoria(df_base, colunas)
            st.session_state["df_audit_nao_triadas"] = df_prep

            for k in list(st.session_state.keys()):
                if k.startswith(("conf_nao_", "motivo_nao_", "acao_nao_", "pag_nao")):
                    del st.session_state[k]
            _inicializar_chaves("nao", df_prep)
            st.rerun()
        return

    # ---- Editor ----
    df    = st.session_state["df_audit_nao_triadas"]
    total = len(df)

    _inicializar_chaves("nao", df)

    s = _stats_chaves("nao", total)
    col_a, col_b, col_c = st.columns([2, 1, 1])
    with col_a:
        st.markdown(f"**Tarefas no editor:** {total}")
    with col_b:
        st.metric("Auditadas", f"{s['auditadas']}/{total}")
    with col_c:
        if s["auditadas"] > 0:
            st.metric("Conformidade", f"{s['pct_conf']:.1f}%")

    _barra_progresso("nao", total)

    st.info(
        "Verifique cada tarefa no SuperSapiens pelo NUP e registre o resultado. "
        "Tarefas *Não auditadas* não entram nas estatísticas finais.",
        icon="ℹ️",
    )

    _render_cartoes("nao", df, df_key="df_audit_nao_triadas", mostrar_config=False)
    _sincronizar_para_df("nao", "df_audit_nao_triadas")

    st.divider()
    col1, col2 = st.columns([2, 1])
    with col1:
        if st.button("Concluir e Ir para Relatório →", type="primary"):
            _sincronizar_para_df("nao", "df_audit_nao_triadas")
            st.session_state["auditoria_nao_triadas_concluida"] = True
            st.session_state["pagina"] = "relatorio"
            st.rerun()
    with col2:
        if st.button("↩ Alterar Seleção"):
            st.session_state["df_audit_nao_triadas"]           = None
            st.session_state["auditoria_nao_triadas_concluida"] = False
            for k in list(st.session_state.keys()):
                if k.startswith(("conf_nao_", "motivo_nao_", "acao_nao_", "pag_nao")):
                    del st.session_state[k]
            st.rerun()


# ===========================================================================
# PÁGINA 4 — RELATÓRIO
# ===========================================================================

def render_relatorio() -> None:
    audit_data = get_audit_data()
    if audit_data is None:
        st.warning("Nenhum arquivo carregado. Volte à página de Importação.")
        return

    st.title("📄 Relatório de Auditoria")

    df_tri = get_df_triadas()
    df_nao = get_df_nao_triadas()
    s_tri  = stats_df(df_tri)
    s_nao  = stats_df(df_nao)

    # ---- Metadados ----
    st.subheader("Identificação do Relatório")
    col1, col2 = st.columns(2)
    with col1:
        responsavel = st.text_input(
            "Responsável pela auditoria:",
            value=st.session_state.get("responsavel", ""),
            placeholder="Nome completo do responsável",
            key="input_responsavel",
        )
        st.session_state["responsavel"] = responsavel
    with col2:
        data_aud = st.date_input(
            "Data da auditoria:",
            value=st.session_state.get("data_auditoria", date_type.today()),
            key="input_data_aud",
            format="DD/MM/YYYY",
        )
        st.session_state["data_auditoria"] = data_aud

    # ---- Resumo executivo ----
    st.divider()
    st.subheader("Resumo Executivo")

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Total de Tarefas", audit_data.total_tarefas)
    c2.metric("Triadas Auditadas", f"{s_tri['auditadas']}/{audit_data.total_triadas}")
    c3.metric(
        "Conformidade (triadas)",
        f"{s_tri['pct_conf']:.1f}%",
        delta=f"{s_tri['conformes']} conformes",
    )
    c4.metric("Não Triadas Auditadas", f"{s_nao['auditadas']}/{audit_data.total_nao_triadas}")
    c5.metric(
        "Conformidade (não triadas)",
        f"{s_nao['pct_conf']:.1f}%",
        delta=f"{s_nao['conformes']} conformes",
    )

    # Gráficos
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    if s_tri["auditadas"] > 0 or s_nao["auditadas"] > 0:
        fig, axes = plt.subplots(1, 2, figsize=(10, 4))

        def _pizza(ax, s: dict, titulo: str):
            if s["auditadas"] == 0:
                ax.text(0.5, 0.5, "Sem dados auditados",
                        ha="center", va="center", transform=ax.transAxes,
                        fontsize=11, color="#888")
                ax.set_title(titulo, fontsize=11, fontweight="bold")
                ax.axis("off")
                return
            labels_v, sizes_v, cores_v = [], [], []
            if s["conformes"] > 0:
                labels_v.append(f"Conformes\n{s['conformes']}")
                sizes_v.append(s["conformes"])
                cores_v.append("#2ecc71")
            if s["nao_conformes"] > 0:
                labels_v.append(f"Não Conformes\n{s['nao_conformes']}")
                sizes_v.append(s["nao_conformes"])
                cores_v.append("#e74c3c")
            nao_aud = s["total"] - s["auditadas"]
            if nao_aud > 0:
                labels_v.append(f"Não auditadas\n{nao_aud}")
                sizes_v.append(nao_aud)
                cores_v.append("#bbb")
            ax.pie(
                sizes_v, labels=labels_v, colors=cores_v,
                autopct="%1.1f%%", startangle=90,
                wedgeprops={"edgecolor": "white", "linewidth": 2},
            )
            ax.set_title(titulo, fontsize=11, fontweight="bold", color="#1A3A6A")

        _pizza(axes[0], s_tri, "Tarefas Triadas")
        _pizza(axes[1], s_nao, "Tarefas Não Triadas")
        fig.tight_layout()
        st.pyplot(fig, use_container_width=True)
        plt.close(fig)

    # ---- Não conformidades ----
    total_nc = s_tri["nao_conformes"] + s_nao["nao_conformes"]
    if total_nc > 0:
        st.divider()
        st.subheader(f"⚠️ Não Conformidades Identificadas ({total_nc})")
        dfs_nc = []
        if df_tri is not None:
            nc_tri = df_tri[df_tri[COL_CONFORMIDADE] == "Não Conforme"][
                [COL_TAREFA, COL_NUP, COL_MOTIVO, COL_ACAO]
            ].copy()
            nc_tri.insert(0, "Origem", "Triada")
            dfs_nc.append(nc_tri)
        if df_nao is not None:
            nc_nao = df_nao[df_nao[COL_CONFORMIDADE] == "Não Conforme"][
                [COL_TAREFA, COL_NUP, COL_MOTIVO, COL_ACAO]
            ].copy()
            nc_nao.insert(0, "Origem", "Não Triada")
            dfs_nc.append(nc_nao)
        if dfs_nc:
            st.dataframe(
                pd.concat(dfs_nc, ignore_index=True),
                hide_index=True, use_container_width=True,
            )
    else:
        st.success("Nenhuma não conformidade identificada nas tarefas auditadas.")

    # ---- Gerar relatório ----
    st.divider()
    col_btn1, col_btn2 = st.columns([1, 2])
    with col_btn1:
        if st.button("📥 Gerar Relatório (.docx)", type="primary", use_container_width=True):
            with st.spinner("Gerando relatório…"):
                try:
                    docx_bytes = gerar_relatorio(
                        audit_data=audit_data,
                        df_triadas=df_tri,
                        df_nao_triadas=df_nao,
                        tipo_controle=st.session_state.get("tipo_controle"),
                        tamanho_amostra=st.session_state.get("tamanho_amostra"),
                        responsavel=responsavel,
                        data_auditoria=data_aud,
                    )
                    st.session_state["relatorio_gerado"] = docx_bytes
                    st.success("Relatório gerado com sucesso!")
                except Exception as e:
                    st.error(f"Erro ao gerar o relatório: {e}")
                    raise

    if st.session_state.get("relatorio_gerado"):
        nome = f"relatorio_auditoria_{data_aud.strftime('%Y-%m-%d')}.docx"
        with col_btn2:
            st.download_button(
                label="⬇️ Baixar Relatório Word (.docx)",
                data=st.session_state["relatorio_gerado"],
                file_name=nome,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                type="primary",
            )


# ===========================================================================
# Dispatch principal
# ===========================================================================
pagina = st.session_state.get("pagina", "importacao")

if pagina == "importacao":
    render_importacao()
elif pagina == "triadas":
    render_auditoria_triadas()
elif pagina == "nao_triadas":
    render_auditoria_nao_triadas()
elif pagina == "relatorio":
    render_relatorio()
