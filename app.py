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
# Auto-persistência: salva edições pendentes de data_editors ao navegar
# ---------------------------------------------------------------------------

def _persist_editor(df_key: str, editor_key: str, indices_key: str) -> None:
    """
    Lê o delta de edição do data_editor (session_state[editor_key]) e aplica
    as mudanças de volta ao DataFrame completo (session_state[df_key]).

    O delta usa índices POSICIONAIS (0, 1, 2…) relativos às linhas exibidas,
    enquanto o DataFrame pode ter índices diferentes quando filtrado.
    indices_key armazena o mapeamento posição → índice original.
    """
    edits = st.session_state.get(editor_key, {})
    edited_rows = edits.get("edited_rows", {})

    df = st.session_state.get(df_key)
    if df is None:
        return

    if edited_rows:
        indices = st.session_state.get(indices_key, list(range(len(df))))
        df = df.copy()
        for pos_str, changes in edited_rows.items():
            pos = int(pos_str)
            if pos < len(indices):
                orig_idx = indices[pos]
                for col, val in changes.items():
                    if col in df.columns:
                        df.at[orig_idx, col] = val
        st.session_state[df_key] = df

    # Limpar estado do editor para evitar mapeamento stale
    if editor_key in st.session_state:
        del st.session_state[editor_key]


def _auto_persist_all() -> None:
    """Persiste edições pendentes de TODOS os editors. Roda no topo de cada ciclo."""
    for df_key, editor_key, indices_key in [
        ("df_audit_triadas", "editor_triadas", "_idx_triadas"),
        ("df_audit_nao_triadas", "editor_nao_triadas", "_idx_nao_triadas"),
    ]:
        if st.session_state.get(df_key) is not None:
            _persist_editor(df_key, editor_key, indices_key)


_auto_persist_all()


# ---------------------------------------------------------------------------
# Sidebar
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

        # Progresso geral
        df_tri = st.session_state.get("df_audit_triadas")
        df_nao = st.session_state.get("df_audit_nao_triadas")
        n_aud = n_total = 0
        if df_tri is not None:
            n_total += len(df_tri)
            n_aud += len(df_tri[df_tri[COL_CONFORMIDADE] != OPCOES_CONFORMIDADE[0]])
        if df_nao is not None:
            n_total += len(df_nao)
            n_aud += len(df_nao[df_nao[COL_CONFORMIDADE] != OPCOES_CONFORMIDADE[0]])
        if n_total > 0:
            st.caption(f"Progresso: **{n_aud}/{n_total}** auditadas")
            st.progress(n_aud / n_total)

        st.divider()

    if st.button("🔄 Nova Auditoria", use_container_width=True):
        # Limpar tudo
        for k in list(st.session_state.keys()):
            if k.startswith(("editor_", "_idx_", "filtro_", "busca_")):
                del st.session_state[k]
        reset_auditoria()
        st.session_state["pagina"] = "importacao"
        st.session_state["audit_data_merged"] = None
        st.rerun()


# ---------------------------------------------------------------------------
# Editor compartilhado: tabela editável com filtro e salvamento
# ---------------------------------------------------------------------------

def _render_editor(
    df_key: str,
    editor_key: str,
    indices_key: str,
    filtro_key: str,
    busca_key: str,
    column_order: list[str],
    disabled_cols: list[str],
) -> None:
    """Renderiza editor de auditoria com filtro, busca e botão de salvar."""
    df = st.session_state[df_key]
    total = len(df)
    s = stats_df(df)

    # ── Progresso ──
    pct = s["auditadas"] / total if total > 0 else 0
    st.progress(
        pct,
        text=(
            f"**{s['auditadas']}/{total}** auditadas"
            f" · {s['conformes']} conformes · {s['nao_conformes']} não conformes"
        ),
    )

    # ── Filtros ──
    def _on_filter_change():
        _persist_editor(df_key, editor_key, indices_key)

    col_f1, col_f2 = st.columns([1, 2])
    with col_f1:
        filtro = st.multiselect(
            "Filtrar por conformidade:",
            OPCOES_CONFORMIDADE,
            default=OPCOES_CONFORMIDADE,
            key=filtro_key,
            on_change=_on_filter_change,
        )
    with col_f2:
        busca = st.text_input(
            "Buscar (Tarefa ou NUP):",
            key=busca_key,
            placeholder="Digite para filtrar…",
            on_change=_on_filter_change,
        )

    # Aplicar filtros
    mask = df[COL_CONFORMIDADE].isin(filtro)
    if busca.strip():
        txt = busca.strip()
        mask = mask & (
            df[COL_TAREFA].astype(str).str.contains(txt, case=False, na=False)
            | df[COL_NUP].astype(str).str.contains(txt, case=False, na=False)
        )

    df_view = df.loc[mask]
    st.session_state[indices_key] = df_view.index.tolist()

    st.caption(f"Exibindo **{len(df_view)}** de {total} tarefas")

    if df_view.empty:
        st.info("Nenhuma tarefa corresponde ao filtro atual.")
        return

    # ── Garantir que colunas existem no df_view (para column_order) ──
    col_order = [c for c in column_order if c in df_view.columns]
    disabled = [c for c in disabled_cols if c in df_view.columns]

    # ── Editor ──
    edited = st.data_editor(
        df_view,
        key=editor_key,
        column_order=col_order,
        column_config={
            COL_TAREFA: st.column_config.TextColumn("Tarefa", width="small"),
            COL_NUP: st.column_config.TextColumn("NUP", width="medium"),
            COL_USUARIO: st.column_config.TextColumn("Usuário", width="small"),
            COL_CONFIG: st.column_config.TextColumn("Config. Encontradas", width="medium"),
            COL_STATUS: st.column_config.TextColumn("Status", width="small"),
            COL_CONFORMIDADE: st.column_config.SelectboxColumn(
                "Conformidade",
                options=OPCOES_CONFORMIDADE,
                required=True,
                width="small",
            ),
            COL_MOTIVO: st.column_config.TextColumn(
                "Motivo NC",
                width="large",
                help="Descreva o motivo da não conformidade",
            ),
            COL_ACAO: st.column_config.TextColumn(
                "Ação Corretiva",
                width="large",
                help="Descreva a ação corretiva proposta",
            ),
        },
        disabled=disabled,
        hide_index=True,
        use_container_width=True,
        num_rows="fixed",
        height=min(800, max(200, 37 + 35 * len(df_view))),
    )

    # ── Salvar ──
    col_save, col_info = st.columns([1, 3])
    with col_save:
        if st.button("💾 Salvar Alterações", type="primary", key=f"btn_save_{df_key}"):
            # Usar o retorno do editor (que já tem os índices originais)
            df_updated = st.session_state[df_key].copy()
            for col in [COL_CONFORMIDADE, COL_MOTIVO, COL_ACAO]:
                if col in edited.columns:
                    df_updated.loc[edited.index, col] = edited[col]
            st.session_state[df_key] = df_updated
            if editor_key in st.session_state:
                del st.session_state[editor_key]
            st.rerun()
    with col_info:
        s_new = stats_df(st.session_state[df_key])
        pendentes = total - s_new["auditadas"]
        if pendentes > 0:
            st.caption(f"⏳ {pendentes} tarefa(s) ainda não auditada(s)")
        else:
            st.caption("✅ Todas as tarefas foram auditadas")


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

    # Processar
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

    # Período
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

    c1, c2, c3 = st.columns(3)
    c1.metric("Total de Tarefas", merged.total_tarefas)
    c2.metric("Tarefas Triadas", merged.total_triadas,
              delta=f"{merged.pct_triadas:.1f}% do total", delta_color="normal")
    c3.metric("Tarefas Não Triadas", merged.total_nao_triadas,
              delta=f"{merged.pct_nao_triadas:.1f}% do total", delta_color="inverse")

    st.divider()
    tab1, tab2 = st.tabs([
        f"Tarefas Triadas ({merged.total_triadas})",
        f"Tarefas Não Triadas ({merged.total_nao_triadas})",
    ])
    with tab1:
        cols_tri = [c for c in [COL_TAREFA, COL_NUP, COL_USUARIO, COL_STATUS, COL_CONFIG]
                    if c in merged.triadas.columns]
        st.dataframe(merged.triadas[cols_tri], hide_index=True, use_container_width=True)
    with tab2:
        cols_nao = [c for c in [COL_TAREFA, COL_NUP, COL_USUARIO, COL_STATUS]
                    if c in merged.nao_triadas.columns]
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

    # ── Seleção do tipo de controle ──
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
            st.session_state["df_audit_triadas"] = preparar_df_auditoria(df_base, colunas)
            st.rerun()
        return

    # ── Editor ──
    tipo_controle = st.session_state["tipo_controle"]
    tipo_label = "Controle Simplificado" if tipo_controle == "simplificado" else "Controle Detalhado"
    n_amostra = st.session_state.get("tamanho_amostra")
    df = st.session_state.get("df_audit_triadas")

    if df is None:
        st.error("Estado inconsistente. Clique em 'Nova Auditoria' no menu lateral.")
        return

    descr = f"Amostra: **{n_amostra}** tarefas" if n_amostra else f"Total: **{len(df)}** tarefas"
    st.markdown(f"**Tipo:** {tipo_label} · {descr}")

    st.info(
        "Edite a coluna **Conformidade** para cada tarefa. "
        "Para não conformidades, preencha também **Motivo NC** e **Ação Corretiva**. "
        "Clique em **Salvar Alterações** para persistir.",
        icon="ℹ️",
    )

    _render_editor(
        df_key="df_audit_triadas",
        editor_key="editor_triadas",
        indices_key="_idx_triadas",
        filtro_key="filtro_conf_tri",
        busca_key="busca_tri",
        column_order=[COL_TAREFA, COL_NUP, COL_CONFIG, COL_CONFORMIDADE, COL_MOTIVO, COL_ACAO],
        disabled_cols=[COL_TAREFA, COL_NUP, COL_USUARIO, COL_CONFIG, COL_STATUS],
    )

    st.divider()
    col1, col2 = st.columns([2, 1])
    with col1:
        if st.button("Concluir e Avançar para Tarefas Não Triadas →", type="primary"):
            st.session_state["auditoria_triadas_concluida"] = True
            st.session_state["pagina"] = "nao_triadas"
            st.rerun()
    with col2:
        if st.button("↩ Trocar Tipo de Controle"):
            st.session_state["tipo_controle"] = None
            st.session_state["df_audit_triadas"] = None
            st.session_state["tamanho_amostra"] = None
            st.session_state["auditoria_triadas_concluida"] = False
            for k in list(st.session_state.keys()):
                if k.startswith(("editor_triadas", "_idx_triadas", "filtro_conf_tri", "busca_tri")):
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
        f"({audit_data.pct_nao_triadas:.1f}% do total)."
    )

    # ── Seleção ──
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
            st.session_state["df_audit_nao_triadas"] = preparar_df_auditoria(df_base, colunas)
            st.rerun()
        return

    # ── Editor ──
    st.info(
        "Edite a coluna **Conformidade** para cada tarefa. "
        "Para não conformidades, preencha também **Motivo NC** e **Ação Corretiva**. "
        "Clique em **Salvar Alterações** para persistir.",
        icon="ℹ️",
    )

    _render_editor(
        df_key="df_audit_nao_triadas",
        editor_key="editor_nao_triadas",
        indices_key="_idx_nao_triadas",
        filtro_key="filtro_conf_nao",
        busca_key="busca_nao",
        column_order=[COL_TAREFA, COL_NUP, COL_STATUS, COL_CONFORMIDADE, COL_MOTIVO, COL_ACAO],
        disabled_cols=[COL_TAREFA, COL_NUP, COL_USUARIO, COL_STATUS],
    )

    st.divider()
    col1, col2 = st.columns([2, 1])
    with col1:
        if st.button("Concluir e Ir para Relatório →", type="primary"):
            st.session_state["auditoria_nao_triadas_concluida"] = True
            st.session_state["pagina"] = "relatorio"
            st.rerun()
    with col2:
        if st.button("↩ Alterar Seleção"):
            st.session_state["df_audit_nao_triadas"] = None
            st.session_state["auditoria_nao_triadas_concluida"] = False
            for k in list(st.session_state.keys()):
                if k.startswith(("editor_nao_triadas", "_idx_nao_triadas",
                                 "filtro_conf_nao", "busca_nao")):
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

    # Metadados
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

    # Resumo executivo
    st.divider()
    st.subheader("Resumo Executivo")

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Total de Tarefas", audit_data.total_tarefas)
    c2.metric("Triadas Auditadas", f"{s_tri['auditadas']}/{audit_data.total_triadas}")
    c3.metric("Conformidade (triadas)", f"{s_tri['pct_conf']:.1f}%",
              delta=f"{s_tri['conformes']} conformes")
    c4.metric("Não Triadas Auditadas", f"{s_nao['auditadas']}/{audit_data.total_nao_triadas}")
    c5.metric("Conformidade (não triadas)", f"{s_nao['pct_conf']:.1f}%",
              delta=f"{s_nao['conformes']} conformes")

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
            ax.pie(sizes_v, labels=labels_v, colors=cores_v,
                   autopct="%1.1f%%", startangle=90,
                   wedgeprops={"edgecolor": "white", "linewidth": 2})
            ax.set_title(titulo, fontsize=11, fontweight="bold", color="#1A3A6A")

        _pizza(axes[0], s_tri, "Tarefas Triadas")
        _pizza(axes[1], s_nao, "Tarefas Não Triadas")
        fig.tight_layout()
        st.pyplot(fig, use_container_width=True)
        plt.close(fig)

    # Não conformidades
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

    # Gerar relatório
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
# Dispatch
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
