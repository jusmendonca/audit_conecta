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


# ---------------------------------------------------------------------------
# SUPP Login — helpers
# ---------------------------------------------------------------------------

def _supp_get_nome(client, fallback: str) -> str:
    """Tenta obter o nome do usuário via payload do JWT."""
    try:
        payload = client.payload()
        return (
            payload.get("name")
            or payload.get("username")
            or payload.get("login")
            or payload.get("sub")
            or fallback
        )
    except Exception:
        return fallback


def _supp_logout_cleanup() -> None:
    client = st.session_state.pop("supp_auth_client", None)
    if client:
        try:
            client.close()
        except Exception:
            pass
    for k in ("supp_logged_in", "supp_username", "supp_auth_client",
              "supp_login_step", "supp_totp_challenge"):
        st.session_state.pop(k, None)


def _supp_do_login(base_url: str, usuario: str, senha: str) -> None:
    """Autentica via LDAP e armazena o cliente na sessão."""
    try:
        from modules.auth import AuthClient, AuthError as _AE
        client = AuthClient(base_url=base_url.rstrip("/"))
        client.login_ldap(usuario, senha)
        st.session_state["supp_auth_client"] = client
        st.session_state["supp_logged_in"] = True
        st.session_state["supp_username"] = _supp_get_nome(client, usuario)
        st.session_state["supp_base_url"] = base_url.rstrip("/")
        st.rerun()
    except Exception as exc:
        try:
            from modules.auth import AuthError as _AE2
            msg = exc.body if isinstance(exc, _AE2) else str(exc)
        except Exception:
            msg = str(exc)
        st.error(f"Credenciais inválidas: {msg}")


def _render_login_page() -> None:
    """Tela de login SUPP via LDAP — exibida antes de qualquer outra coisa."""
    st.markdown(
        "<style>.block-container{padding-top:5rem;}</style>",
        unsafe_allow_html=True,
    )
    _, col, _ = st.columns([1, 1.5, 1])
    with col:
        st.markdown("## 📋 Auditoria Conecta+")
        st.markdown("##### Procuradoria-Geral Federal / AGU")
        st.divider()

        SUPP_URL = "https://supersapiensbackend.agu.gov.br"

        with st.form("login_page_form"):
            usuario = st.text_input("Usuário (login LDAP)", placeholder="cpf ou login")
            senha = st.text_input("Senha", type="password")
            ok = st.form_submit_button(
                "Entrar →", use_container_width=True, type="primary"
            )

        if ok:
            if not usuario.strip() or not senha:
                st.error("Preencha usuário e senha.")
            else:
                _supp_do_login(SUPP_URL, usuario.strip(), senha)


# ── Login gate: app só funciona após autenticação ──────────────────────────
if not st.session_state.get("supp_logged_in"):
    _render_login_page()
    st.stop()


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
            if k.startswith(("filtro_", "busca_", "tbl_",
                             "edit_conf_", "edit_motivo_", "edit_acao_", "btn_save_row_")):
                del st.session_state[k]
        reset_auditoria()
        st.session_state["pagina"] = "importacao"
        st.session_state["audit_data_merged"] = None
        st.rerun()

    # ── Usuário SUPP ──────────────────────────────────────────────────────────
    st.divider()
    nome = st.session_state.get("supp_username", "")
    st.caption(f"🟢 **{nome}**")
    if st.button("Sair do SUPP", use_container_width=True, key="btn_supp_logout"):
        _supp_logout_cleanup()
        st.rerun()


# ---------------------------------------------------------------------------
# SUPP — Painel de conferência
# ---------------------------------------------------------------------------

def _get_nested(d: dict, path: str):
    """Acessa valor aninhado por caminho dotted. Ex: 'setorAtual.nome'"""
    val = d
    for k in path.split("."):
        if not isinstance(val, dict):
            return None
        val = val.get(k)
    return val


def _fmt_date(s) -> str:
    """Formata ISO datetime para dd/mm/aaaa."""
    if not s:
        return "—"
    try:
        from datetime import datetime as _dt
        dt = _dt.fromisoformat(str(s).replace("Z", "+00:00"))
        return dt.strftime("%d/%m/%Y")
    except Exception:
        return str(s)[:10]



def _render_supp_panel() -> None:
    """Painel lateral de conferência — exibe dados completos da tarefa e do processo."""
    from datetime import datetime as _dt, timezone as _tz, timedelta as _td

    nup = st.session_state.get("supp_sel_nup")
    tarefa_id = st.session_state.get("supp_sel_tarefa_id")

    st.markdown("#### 🔍 Conferência SUPP")

    if not nup:
        st.info("Selecione uma linha na tabela para consultar dados no sistema.")
        return

    # ── Configuração de período (persiste na sessão) ─────────────────────────
    dias = st.slider(
        "Eventos dos últimos (dias):",
        min_value=7, max_value=365, step=7,
        key="supp_dias_eventos",
        value=st.session_state.get("supp_dias_eventos", 30),
    )

    cache_key = f"_supp_cache_{nup}"

    if cache_key not in st.session_state:
        auth = st.session_state.get("supp_auth_client")
        cache: dict = {}
        if auth:
            with st.spinner("Consultando SUPP..."):
                # 1. Tarefa
                if tarefa_id:
                    try:
                        from modules.tarefa import TarefaClient
                        tc = TarefaClient.from_auth(auth)
                        cache["tarefa"] = tc.buscar(
                            tarefa_id,
                            populate=["processo", "vinculacaoWorkflow", "especieTarefa",
                                      "usuarioResponsavel"],
                        )
                    except Exception as e:
                        cache["tarefa_erro"] = str(e)

                # 2. Processo — extraindo o ID da tarefa
                proc_id = None
                if "tarefa" in cache:
                    proc_emb = cache["tarefa"].get("processo")
                    if isinstance(proc_emb, dict):
                        proc_id = proc_emb.get("id")

                if proc_id:
                    try:
                        from modules.processo import ProcessoClient
                        pc = ProcessoClient.from_auth(auth)
                        cache["processo"] = pc.buscar(
                            proc_id,
                            populate=["especieProcesso", "setorAtual", "setorInicial",
                                      "classificacao", "criadoPor", "processoJudicial"],
                        )
                    except Exception as e:
                        cache["processo_erro"] = str(e)

                    # 3. Etiquetas do processo
                    try:
                        from modules.etiqueta import EtiquetaClient
                        ec = EtiquetaClient.from_auth(auth)
                        cache["etiquetas_processo"] = ec.listar_por_processo(proc_id)
                    except Exception as e:
                        cache["etiquetas_processo_erro"] = str(e)

                    # 4. Interessados do processo
                    try:
                        from modules.interessado import InteressadoClient
                        ic = InteressadoClient.from_auth(auth)
                        cache["interessados"] = ic.listar_por_processo(proc_id)
                    except Exception as e:
                        cache["interessados_erro"] = str(e)

                    # 5. Timeline do processo (filtrada por dias na exibição)
                    try:
                        from modules.processo import ProcessoClient as _PC2
                        cache["timeline"] = _PC2.from_auth(auth).timeline(proc_id)
                    except Exception as e:
                        cache["timeline_erro"] = str(e)

                # 6. Etiquetas da tarefa
                if tarefa_id:
                    try:
                        from modules.etiqueta import EtiquetaClient as _EC2
                        cache["etiquetas_tarefa"] = _EC2.from_auth(auth).listar_por_tarefa(tarefa_id)
                    except Exception as e:
                        cache["etiquetas_tarefa_erro"] = str(e)

        st.session_state[cache_key] = cache

    cache = st.session_state[cache_key]

    # Cabeçalho + botão atualizar
    c_nup, c_btn = st.columns([4, 1])
    with c_nup:
        st.caption(f"`{nup}`")
    with c_btn:
        if st.button("🔄", key=f"_refresh_{nup}", help="Recarregar dados do SUPP"):
            st.session_state.pop(cache_key, None)
            st.rerun()

    # ── Tarefa ────────────────────────────────────────────────────────────────
    if "tarefa" in cache:
        t = cache["tarefa"]
        st.markdown("**📋 Tarefa**")

        vw = t.get("vinculacaoWorkflow")
        workflow_txt = None
        if isinstance(vw, dict):
            wf = vw.get("workflow")
            workflow_txt = (
                (_get_nested(wf, "nome") if isinstance(wf, dict) else None)
                or f"ID {vw.get('id')}"
                + (" (concluído)" if vw.get("concluido") else "")
            )

        campos_t = [
            ("Espécie", _get_nested(t, "especieTarefa.nome") or t.get("especieTarefa")),
            ("Responsável", _get_nested(t, "usuarioResponsavel.nome")),
            ("Prazo", _fmt_date(t.get("dataHoraFinalPrazo"))),
            ("Urgente", "Sim" if t.get("urgente") else "Não"),
            ("Status", t.get("situacaoTarefa") or t.get("situacao") or t.get("status")),
            ("Fluxo", workflow_txt),
            ("Post-it", t.get("postIt")),
        ]
        for label, val in campos_t:
            if val is not None and val != "":
                st.markdown(f"**{label}:** {val}")

        # Etiquetas da tarefa
        et_t = cache.get("etiquetas_tarefa", [])
        if et_t:
            nomes_et_t = [
                _get_nested(e, "etiqueta.nome") or str(e.get("etiqueta", ""))
                for e in et_t if e
            ]
            nomes_et_t = [n for n in nomes_et_t if n]
            if nomes_et_t:
                st.markdown("**Etiquetas:** " + " · ".join(f"`{n}`" for n in nomes_et_t))
        elif "etiquetas_tarefa_erro" in cache:
            st.caption(f"⚠️ Etiquetas: {cache['etiquetas_tarefa_erro']}")

    elif "tarefa_erro" in cache:
        st.warning(f"Tarefa: {cache['tarefa_erro']}", icon="⚠️")

    # ── Processo ──────────────────────────────────────────────────────────────
    if "processo" in cache:
        p = cache["processo"]
        st.markdown("---")
        st.markdown("**📁 Processo**")

        # CNJ e classe: via processoJudicial se disponível
        pj = p.get("processoJudicial")
        cnj = None
        classe_cnj = None
        if isinstance(pj, dict):
            cnj = pj.get("numero") or pj.get("numeroFormatado") or pj.get("numeroAlternativo")
            classe_cnj = _get_nested(pj, "classeNacional.nome")
        classe_proc = _get_nested(p, "classificacao.nome") or _get_nested(p, "classificacao.nomeCompleto")

        campos_p = [
            ("NUP", p.get("NUP") or p.get("nup")),
            ("Espécie/Tipo", _get_nested(p, "especieProcesso.nome") or p.get("descricao") or p.get("assunto")),
            ("Número CNJ", cnj),
            ("Classe CNJ", classe_cnj),
            ("Classe processual", classe_proc),
            ("Setor responsável", _get_nested(p, "setorAtual.nome")),
            ("Setor inicial", _get_nested(p, "setorInicial.nome")),
            ("Status", p.get("status") or p.get("situacao")),
            ("Autuado em", _fmt_date(p.get("dataHoraAbertura") or p.get("dataHoraCriacao"))),
            ("Observação", p.get("observacao")),
        ]
        for label, val in campos_p:
            if val is not None and val != "":
                st.markdown(f"**{label}:** {val}")

        # Etiquetas do processo
        et_p = cache.get("etiquetas_processo", [])
        if et_p:
            nomes_et_p = [
                _get_nested(e, "etiqueta.nome") or str(e.get("etiqueta", ""))
                for e in et_p if e
            ]
            nomes_et_p = [n for n in nomes_et_p if n]
            if nomes_et_p:
                st.markdown("**Etiquetas:** " + " · ".join(f"`{n}`" for n in nomes_et_p))
        elif "etiquetas_processo_erro" in cache:
            st.caption(f"⚠️ Etiquetas: {cache['etiquetas_processo_erro']}")

        # Interessados
        interessados = cache.get("interessados", [])
        label_int = f"👥 Interessados ({len(interessados)})"
        if "interessados_erro" in cache:
            label_int += " ⚠️"
        with st.expander(label_int, expanded=False):
            if interessados:
                for intr in interessados:
                    nome_p = (
                        _get_nested(intr, "pessoa.nome")
                        or _get_nested(intr, "pessoa.pessoaFisica.nome")
                        or _get_nested(intr, "pessoa.pessoaJuridica.razaoSocial")
                        or str(intr.get("pessoa", "—"))
                    )
                    modalidade = _get_nested(intr, "modalidadeInteressado.valor") or ""
                    linha = f"- {nome_p}"
                    if modalidade:
                        linha += f" *({modalidade})*"
                    st.markdown(linha)
            elif "interessados_erro" in cache:
                st.warning(cache["interessados_erro"])
            else:
                st.caption("Nenhum interessado registrado.")

    elif "processo_erro" in cache:
        st.warning(f"Processo: {cache['processo_erro']}", icon="⚠️")

    # ── Timeline ──────────────────────────────────────────────────────────────
    timeline_raw = cache.get("timeline", [])
    corte = _dt.now(_tz.utc) - _td(days=dias)

    def _parse_timeline_events(raw: list) -> list[dict]:
        """Normaliza estrutura variável da timeline para lista de {data, msg}."""
        events = []
        for item in raw:
            if isinstance(item, dict):
                # Estrutura direta
                evt_date = item.get("eventDate") or item.get("dataHora") or item.get("criadoEm")
                msg = item.get("message") or item.get("mensagem") or item.get("descricao") or ""
                # Estrutura aninhada em entities
                if not evt_date and "entities" in item:
                    for sub in item.get("entities", []):
                        if isinstance(sub, dict):
                            te = sub.get("timelineEvent") or sub
                            evt_date = te.get("eventDate") or te.get("dataHora")
                            msg = te.get("message") or te.get("mensagem") or msg
                            break
                if evt_date:
                    events.append({"data": evt_date, "msg": str(msg)})
        return events

    eventos = _parse_timeline_events(timeline_raw)
    eventos_filtrados = []
    for ev in eventos:
        try:
            dt_ev = _dt.fromisoformat(str(ev["data"]).replace("Z", "+00:00"))
            if dt_ev >= corte:
                eventos_filtrados.append((dt_ev, ev["msg"]))
        except Exception:
            pass
    eventos_filtrados.sort(key=lambda x: x[0], reverse=True)

    label_tl = f"📅 Eventos — últimos {dias} dias ({len(eventos_filtrados)})"
    if "timeline_erro" in cache:
        label_tl += " ⚠️"
    with st.expander(label_tl, expanded=len(eventos_filtrados) > 0 and len(eventos_filtrados) <= 5):
        if eventos_filtrados:
            for dt_ev, msg in eventos_filtrados:
                st.markdown(f"- **{dt_ev.strftime('%d/%m/%Y')}** — {msg}")
        elif "timeline_erro" in cache:
            st.warning(cache["timeline_erro"])
        elif timeline_raw is not None:
            st.caption(f"Nenhum evento nos últimos {dias} dias.")


# ---------------------------------------------------------------------------
# Tabela interativa + editor de linha
# ---------------------------------------------------------------------------


def _render_audit_table(
    df_key: str,
    filtro_key: str,
    busca_key: str,
    column_order: list[str],
    table_key: str,
) -> tuple:
    """
    Tabela interativa com filtros e seleção de linha.
    Retorna (orig_idx, row_dict) da linha selecionada, ou (None, None).
    """
    df = st.session_state[df_key]
    total = len(df)
    s = stats_df(df)

    pct = s["auditadas"] / total if total > 0 else 0
    st.progress(
        pct,
        text=(
            f"**{s['auditadas']}/{total}** auditadas"
            f" · {s['conformes']} conformes · {s['nao_conformes']} não conformes"
        ),
    )

    col_f1, col_f2 = st.columns([1, 2])
    with col_f1:
        filtro = st.multiselect(
            "Filtrar conformidade:",
            OPCOES_CONFORMIDADE,
            default=OPCOES_CONFORMIDADE,
            key=filtro_key,
        )
    with col_f2:
        has_config = COL_CONFIG in df.columns
        busca_label = (
            "Buscar (Tarefa, NUP ou Config.):" if has_config else "Buscar (Tarefa ou NUP):"
        )
        busca = st.text_input(busca_label, key=busca_key, placeholder="Digite para filtrar…")

    mask = df[COL_CONFORMIDADE].isin(filtro)
    if busca.strip():
        txt = busca.strip()
        search_mask = (
            df[COL_TAREFA].astype(str).str.contains(txt, case=False, na=False)
            | df[COL_NUP].astype(str).str.contains(txt, case=False, na=False)
        )
        if has_config:
            search_mask = search_mask | df[COL_CONFIG].astype(str).str.contains(txt, case=False, na=False)
        mask = mask & search_mask

    df_view = df.loc[mask]
    col_order = [c for c in column_order if c in df_view.columns]

    st.caption(f"Exibindo **{len(df_view)}** de {total} tarefas — clique em uma linha para auditar")

    if df_view.empty:
        st.info("Nenhuma tarefa corresponde ao filtro atual.")
        return None, None

    event = st.dataframe(
        df_view[col_order],
        use_container_width=True,
        hide_index=True,
        on_select="rerun",
        selection_mode="single-row",
        key=table_key,
        column_config={
            COL_TAREFA: st.column_config.TextColumn("Tarefa", width="small"),
            COL_NUP: st.column_config.TextColumn("NUP", width="medium"),
            COL_USUARIO: st.column_config.TextColumn("Usuário", width="small"),
            COL_CONFIG: st.column_config.TextColumn("Config.", width="medium"),
            COL_STATUS: st.column_config.TextColumn("Status", width="small"),
            COL_CONFORMIDADE: st.column_config.TextColumn("Conformidade", width="small"),
            COL_MOTIVO: st.column_config.TextColumn("Motivo NC", width="medium"),
            COL_ACAO: st.column_config.TextColumn("Ação Corretiva", width="medium"),
        },
        height=min(600, max(200, 37 + 35 * len(df_view))),
    )

    rows = event.selection.rows
    if rows:
        orig_idx = df_view.index[rows[0]]
        row = df.loc[orig_idx].to_dict()
        st.session_state["supp_sel_nup"] = row.get(COL_NUP)
        st.session_state["supp_sel_tarefa_id"] = row.get(COL_TAREFA)
        return orig_idx, row

    return None, None


def _render_row_editor(df_key: str, orig_idx, row: dict) -> None:
    """Painel de edição dos campos de auditoria de uma linha selecionada."""
    st.markdown("#### ✏️ Auditoria")
    st.caption(f"Tarefa `{row.get(COL_TAREFA)}` · `{row.get(COL_NUP)}`")

    cur_conf = row.get(COL_CONFORMIDADE, OPCOES_CONFORMIDADE[0])
    if cur_conf not in OPCOES_CONFORMIDADE:
        cur_conf = OPCOES_CONFORMIDADE[0]

    conf = st.selectbox(
        "Conformidade:",
        OPCOES_CONFORMIDADE,
        index=OPCOES_CONFORMIDADE.index(cur_conf),
        key=f"edit_conf_{orig_idx}_{df_key}",
    )
    motivo = st.text_area(
        "Motivo NC:",
        value=row.get(COL_MOTIVO, "") or "",
        key=f"edit_motivo_{orig_idx}_{df_key}",
        height=90,
        placeholder="Descreva o motivo da não conformidade…",
    )
    acao = st.text_area(
        "Ação Corretiva:",
        value=row.get(COL_ACAO, "") or "",
        key=f"edit_acao_{orig_idx}_{df_key}",
        height=90,
        placeholder="Descreva a ação corretiva proposta…",
    )

    if st.button("💾 Salvar", type="primary", key=f"btn_save_row_{orig_idx}_{df_key}",
                 use_container_width=True):
        df = st.session_state[df_key].copy()
        df.at[orig_idx, COL_CONFORMIDADE] = conf
        df.at[orig_idx, COL_MOTIVO] = motivo
        df.at[orig_idx, COL_ACAO] = acao
        st.session_state[df_key] = df
        st.rerun()


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

    col_left, col_right = st.columns([3, 2], gap="medium")
    with col_left:
        orig_idx, row = _render_audit_table(
            df_key="df_audit_triadas",
            filtro_key="filtro_conf_tri",
            busca_key="busca_tri",
            column_order=[COL_TAREFA, COL_NUP, COL_USUARIO, COL_CONFIG, COL_STATUS,
                          COL_CONFORMIDADE, COL_MOTIVO, COL_ACAO],
            table_key="tbl_triadas",
        )
    with col_right:
        if orig_idx is not None and row is not None:
            _render_row_editor("df_audit_triadas", orig_idx, row)
            st.divider()
        _render_supp_panel()

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
                if k.startswith(("filtro_conf_tri", "busca_tri", "tbl_triadas",
                                 "edit_conf_", "edit_motivo_", "edit_acao_", "btn_save_row_")):
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
    col_left, col_right = st.columns([3, 2], gap="medium")
    with col_left:
        orig_idx, row = _render_audit_table(
            df_key="df_audit_nao_triadas",
            filtro_key="filtro_conf_nao",
            busca_key="busca_nao",
            column_order=[COL_TAREFA, COL_NUP, COL_USUARIO, COL_STATUS,
                          COL_CONFORMIDADE, COL_MOTIVO, COL_ACAO],
            table_key="tbl_nao_triadas",
        )
    with col_right:
        if orig_idx is not None and row is not None:
            _render_row_editor("df_audit_nao_triadas", orig_idx, row)
            st.divider()
        _render_supp_panel()

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
                if k.startswith(("filtro_conf_nao", "busca_nao", "tbl_nao_triadas",
                                 "edit_conf_", "edit_motivo_", "edit_acao_", "btn_save_row_")):
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
