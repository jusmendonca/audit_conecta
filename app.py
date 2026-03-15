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
    .metric-box {
        background: #f0f4fb; border-radius: 8px;
        padding: 0.6rem 1rem; border-left: 4px solid #1A3A6A;
    }
    .info-box {
        background: #eaf4fb; border-left: 4px solid #2980b9;
        border-radius: 4px; padding: 0.6rem 1rem; margin-bottom: 0.5rem;
    }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Sidebar — Navegação
# ---------------------------------------------------------------------------
PAGINAS = {
    "importacao": ("📂", "1. Importação"),
    "triadas":    ("✅", "2. Triadas"),
    "nao_triadas":("🔍", "3. Não Triadas"),
    "relatorio":  ("📄", "4. Relatório"),
}


def _check_icon(chave: str) -> str:
    checks = {
        "importacao": st.session_state.get("audit_data_merged") is not None,
        "triadas":    st.session_state.get("auditoria_triadas_concluida", False),
        "nao_triadas": st.session_state.get("auditoria_nao_triadas_concluida", False),
        "relatorio":  False,
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

    # Mostrar resumo rápido se dados carregados
    ad = get_audit_data()
    if ad:
        st.caption(f"📁 {ad.nome_arquivo}")
        st.caption(f"Total: {ad.total_tarefas} | Triadas: {ad.total_triadas} | Não triadas: {ad.total_nao_triadas}")
        st.divider()

    if st.button("🔄 Nova Auditoria", use_container_width=True):
        reset_auditoria()
        st.session_state["pagina"] = "importacao"
        st.session_state["audit_data_merged"] = None
        st.rerun()


# ---------------------------------------------------------------------------
# Helper: editor de auditoria (data_editor compacto)
# ---------------------------------------------------------------------------

def _config_editor(incluir_config: bool = False, incluir_status: bool = False):
    """Retorna column_config para o st.data_editor de auditoria."""
    cfg = {
        COL_TAREFA: st.column_config.TextColumn("Tarefa", disabled=True, width="small"),
        COL_NUP: st.column_config.TextColumn("NUP", disabled=True, width="medium"),
        COL_USUARIO: st.column_config.TextColumn("Usuário", disabled=True, width="small"),
        COL_STATUS: st.column_config.TextColumn("Status", disabled=True, width="medium"),
        COL_CONFIG: st.column_config.TextColumn("Configurações Encontradas", disabled=True, width="large"),
        COL_CONFORMIDADE: st.column_config.SelectboxColumn(
            "Conformidade",
            options=OPCOES_CONFORMIDADE,
            required=True,
            width="small",
        ),
        COL_MOTIVO: st.column_config.TextColumn(
            "Motivo (se Não Conforme)",
            width="large",
            max_chars=500,
        ),
        COL_ACAO: st.column_config.TextColumn(
            "Ação Corretiva",
            width="large",
            max_chars=500,
        ),
    }
    return cfg


def _progress_bar(df: pd.DataFrame, label: str = "") -> None:
    """Exibe barra de progresso baseada em conformidade preenchida."""
    s = stats_df(df)
    total = s["total"]
    auditadas = s["auditadas"]
    pct = auditadas / total if total > 0 else 0
    st.progress(pct, text=f"{label} **{auditadas}/{total}** registradas "
                           f"({s['conformes']} conformes · {s['nao_conformes']} não conformes)")


# ===========================================================================
# PÁGINA 1 — IMPORTAÇÃO
# ===========================================================================

def render_importacao() -> None:
    st.title("📂 Importação de Arquivo")
    st.caption(
        "Importe a(s) planilha(s) Excel gerada(s) pelo módulo de Triagem Avançada "
        "do Conecta+ Automação."
    )

    col_up, col_info = st.columns([2, 1])
    with col_up:
        uploaded = st.file_uploader(
            "Arquivo(s) Excel (.xlsx)",
            type=["xlsx"],
            accept_multiple_files=True,
            help="O arquivo deve conter as abas: Todas as Tarefas, Tarefas Triadas, Tarefas Não Triadas.",
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
            st.info(f"Arquivo **{ad.nome_arquivo}** já carregado. Navegue pelas etapas no menu lateral.")
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
        st.info(f"{len(audit_files)} arquivo(s) consolidados. Duplicatas removidas (mantida ocorrência mais recente).")

    # Reiniciar se arquivo mudou
    atual = st.session_state.get("audit_data_merged")
    if atual is not None and atual.nome_arquivo != merged.nome_arquivo:
        reset_auditoria()

    st.session_state["audit_data_merged"] = merged

    # ---- Estatísticas ----
    st.divider()
    st.subheader("Resumo da Triagem")

    periodo = "N/D"
    if merged.periodo_inicio and merged.periodo_fim:
        periodo = (f"{merged.periodo_inicio.strftime('%d/%m/%Y')} – "
                   f"{merged.periodo_fim.strftime('%d/%m/%Y')}")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total de Tarefas", merged.total_tarefas)
    c2.metric("Triadas", merged.total_triadas, delta=f"{merged.pct_triadas:.1f}%", delta_color="normal")
    c3.metric("Não Triadas", merged.total_nao_triadas, delta=f"{merged.pct_nao_triadas:.1f}%", delta_color="inverse")
    c4.metric("Período", periodo)

    # ---- Visualização das abas ----
    st.divider()
    tab1, tab2 = st.tabs([
        f"Tarefas Triadas ({merged.total_triadas})",
        f"Tarefas Não Triadas ({merged.total_nao_triadas})",
    ])

    cols_tri = [c for c in [COL_TAREFA, COL_NUP, COL_USUARIO, COL_STATUS, COL_CONFIG] if c in merged.triadas.columns]
    cols_nao = [c for c in [COL_TAREFA, COL_NUP, COL_USUARIO, COL_STATUS] if c in merged.nao_triadas.columns]

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
            "Selecione o tipo de controle conforme o Manual de Gerenciamento (seção 5)."
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
            **Controle Simplificado** — Verificação manual de tarefas selecionadas pelo auditor.
            Indicado para fluxos bem estruturados com padrões definidos.
            Periodicidade recomendada: **diária** (Manual, seção 5.1).

            **Controle Detalhado** — Amostragem estatística com seleção aleatória.
            Nível de confiança **95%**, margem de erro **5%**.
            Indicado para análise mais rigorosa de conformidade.
            """)

        with col_dir:
            st.markdown("#### Cálculo da Amostra")
            t = audit_data.total_triadas
            if t > 0:
                n = calcular_amostra(t)
                st.metric("Tarefas a auditar (detalhado)", n,
                          delta=f"{n/t*100:.1f}% do universo")
                st.markdown(formula_descricao(t))

                with st.expander("📊 Tabela de Referência (Anexo III do Manual)"):
                    df_ref = pd.DataFrame(
                        tabela_referencia(),
                        columns=["Universo (N)", "Amostra (n)"]
                    )
                    df_ref["Calculado pela fórmula"] = df_ref["Universo (N)"].apply(
                        lambda x: calcular_amostra(x)
                    )
                    st.dataframe(df_ref, hide_index=True, use_container_width=True)

        st.divider()
        if st.button("Confirmar e Abrir Editor →", type="primary"):
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

    # ---- Editor de auditoria ----
    tipo_controle = st.session_state["tipo_controle"]
    tipo_label = "Controle Simplificado" if tipo_controle == "simplificado" else "Controle Detalhado"
    n_amostra = st.session_state.get("tamanho_amostra")
    df = st.session_state.get("df_audit_triadas")

    if df is None:
        st.error("Estado inconsistente. Clique em 'Nova Auditoria' no menu lateral.")
        return

    # Cabeçalho resumido
    col_a, col_b, col_c = st.columns([2, 1, 1])
    with col_a:
        st.markdown(f"**Tipo:** {tipo_label}"
                    + (f" · **Amostra:** {n_amostra} tarefas" if n_amostra else f" · **Total:** {len(df)} tarefas"))
    with col_b:
        s = stats_df(df)
        st.metric("Auditadas", f"{s['auditadas']}/{s['total']}")
    with col_c:
        if s["auditadas"] > 0:
            st.metric("Taxa de Conformidade", f"{s['pct_conf']:.1f}%")

    _progress_bar(df)

    if tipo_controle == "simplificado":
        st.info(
            "**Controle Simplificado:** Verifique as tarefas no sistema "
            "[SuperSapiens](https://supersapiens.agu.gov.br) pelo NUP e registre o resultado. "
            "Tarefas deixadas como *Não auditada* são excluídas das estatísticas.",
            icon="ℹ️",
        )
    else:
        st.info(
            f"**Controle Detalhado:** {n_amostra} tarefas selecionadas aleatoriamente. "
            "Verifique cada uma no SuperSapiens e registre o resultado.",
            icon="ℹ️",
        )

    # ---- data_editor ----
    edited = st.data_editor(
        df,
        key="editor_triadas",
        column_config=_config_editor(incluir_config=True),
        hide_index=True,
        use_container_width=True,
        num_rows="fixed",
    )
    # Auto-salvar
    st.session_state["df_audit_triadas"] = edited

    st.divider()
    col1, col2, col3 = st.columns([2, 1, 1])
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
            st.rerun()
    with col3:
        s = stats_df(edited)
        if s["auditadas"] > 0:
            st.caption(f"✅ {s['conformes']} conformes · ❌ {s['nao_conformes']} não conformes")


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
        "Registre a avaliação de cada tarefa conferida no SAPIENS."
    )

    # ---- Inicializar DataFrame se ainda não existe ----
    if st.session_state.get("df_audit_nao_triadas") is None:
        # Perguntar se audita todas ou seleciona
        st.divider()
        st.subheader("Seleção das Tarefas")

        nao_triadas = audit_data.nao_triadas
        cols_show = [c for c in [COL_TAREFA, COL_NUP, COL_USUARIO, COL_STATUS] if c in nao_triadas.columns]

        st.dataframe(nao_triadas[cols_show], hide_index=True, use_container_width=True)
        st.divider()

        col_sel, col_opt = st.columns([1, 1])
        with col_sel:
            modo = st.radio(
                "Quais tarefas auditar?",
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
                df_base = nao_triadas[nao_triadas[COL_TAREFA].astype(str).isin(ids_sel)].copy()

            colunas = [COL_TAREFA, COL_NUP, COL_USUARIO, COL_STATUS]
            st.session_state["df_audit_nao_triadas"] = preparar_df_auditoria(df_base, colunas)
            st.rerun()
        return

    # ---- Editor ----
    df = st.session_state["df_audit_nao_triadas"]
    s = stats_df(df)

    col_a, col_b, col_c = st.columns([2, 1, 1])
    with col_a:
        st.markdown(f"**Tarefas no editor:** {s['total']}")
    with col_b:
        st.metric("Auditadas", f"{s['auditadas']}/{s['total']}")
    with col_c:
        if s["auditadas"] > 0:
            st.metric("Taxa de Conformidade", f"{s['pct_conf']:.1f}%")

    _progress_bar(df)

    st.info(
        "Verifique cada tarefa no [SuperSapiens](https://supersapiens.agu.gov.br) "
        "pelo NUP e registre o resultado. Tarefas *Não auditadas* são excluídas das estatísticas.",
        icon="ℹ️",
    )

    edited = st.data_editor(
        df,
        key="editor_nao_triadas",
        column_config=_config_editor(incluir_status=True),
        hide_index=True,
        use_container_width=True,
        num_rows="fixed",
    )
    st.session_state["df_audit_nao_triadas"] = edited

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
    s_tri = stats_df(df_tri)
    s_nao = stats_df(df_nao)

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

    # ---- Resumo gráfico ----
    st.divider()
    st.subheader("Resumo Executivo")

    # Métricas
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Total de Tarefas", audit_data.total_tarefas)
    c2.metric("Triadas Auditadas", f"{s_tri['auditadas']}/{audit_data.total_triadas}")
    c3.metric("Conformes (triadas)", f"{s_tri['pct_conf']:.1f}%",
              delta=f"{s_tri['conformes']} tarefas")
    c4.metric("Não Triadas Auditadas", f"{s_nao['auditadas']}/{audit_data.total_nao_triadas}")
    c5.metric("Conformes (não triadas)", f"{s_nao['pct_conf']:.1f}%",
              delta=f"{s_nao['conformes']} tarefas")

    # Gráficos inline (usando matplotlib via st.pyplot)
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    if s_tri["auditadas"] > 0 or s_nao["auditadas"] > 0:
        fig, axes = plt.subplots(1, 2, figsize=(10, 4))

        def _pizza(ax, s: dict, titulo: str):
            if s["auditadas"] == 0:
                ax.text(0.5, 0.5, "Sem dados auditados", ha="center", va="center",
                        transform=ax.transAxes, fontsize=11, color="#888")
                ax.set_title(titulo, fontsize=11, fontweight="bold")
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
            ax.pie(sizes_v, labels=labels_v, colors=cores_v,
                   autopct="%1.1f%%", startangle=90,
                   wedgeprops={"edgecolor": "white", "linewidth": 2})
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
            nc_tri = df_tri[df_tri[COL_CONFORMIDADE] == "Não Conforme"][[COL_TAREFA, COL_NUP, COL_MOTIVO, COL_ACAO]].copy()
            nc_tri.insert(0, "Origem", "Triada")
            dfs_nc.append(nc_tri)
        if df_nao is not None:
            nc_nao = df_nao[df_nao[COL_CONFORMIDADE] == "Não Conforme"][[COL_TAREFA, COL_NUP, COL_MOTIVO, COL_ACAO]].copy()
            nc_nao.insert(0, "Origem", "Não Triada")
            dfs_nc.append(nc_nao)

        if dfs_nc:
            df_nc_all = pd.concat(dfs_nc, ignore_index=True)
            st.dataframe(df_nc_all, hide_index=True, use_container_width=True)
    else:
        st.success("Nenhuma não conformidade identificada.")

    # ---- Gerar relatório ----
    st.divider()
    col_btn1, col_btn2 = st.columns([1, 2])
    with col_btn1:
        if st.button("📥 Gerar Relatório (.docx)", type="primary", use_container_width=True):
            with st.spinner("Gerando relatório..."):
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
                    st.error(f"Erro ao gerar relatório: {e}")
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
