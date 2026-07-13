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
    COL_DIST_ID, COL_DIST_NUP, COL_DIST_PROCESSO_JUDICIAL,
    COL_DIST_FONTE_DADOS, COL_DIST_USUARIO_ORIGEM, COL_DIST_SETOR_ORIGEM,
    COL_DIST_USUARIO_DESTINO, COL_DIST_SETOR_DESTINO, COL_DIST_DATA_HORA,
    load_file, merge_audit_data, load_distribution_file, detect_file_type,
)
from modules.sampling import (
    calcular_amostra, formula_descricao, selecionar_amostra, tabela_referencia,
)
from modules.state import (
    COL_ACAO, COL_CONFORMIDADE, COL_MOTIVO,
    OPCOES_CONFORMIDADE,
    get_audit_data, get_dist_data, get_df_nao_triadas, get_df_triadas, get_df_distribuicao,
    init_state, preparar_df_auditoria, reset_auditoria, stats_df,
    save_session, load_session, has_saved_session, clear_saved_session, get_session_info,
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
/* ═══════════════════════════════════════════════════════════════════
   AUDITORIA CONECTA+ — DESIGN SYSTEM
   Navy #1A3A6A · Blue #2d5fa0 · Ice #eaf1fb · Body #f7f9fc
   Border #d0dcea · Muted #7a8fad · Text #1a2a4a
   ═══════════════════════════════════════════════════════════════════ */

.block-container { padding-top: 1.2rem; padding-bottom: 1rem; }
div[data-testid="stSidebarNav"] { display: none; }

/* ── Sidebar ──────────────────────────────────────────────────────── */
div[data-testid="stSidebar"] {
    background: #f0f4fa;
    border-right: 1px solid #d0dcea;
}
div[data-testid="stSidebar"] hr { border-color: #c2d4ee; margin: 0.35rem 0; }

/* ── Tipografia ───────────────────────────────────────────────────── */
h1, h2, h3 { color: #1a2a4a !important; }

/* ── Botões primários ─────────────────────────────────────────────── */
button[kind="primary"] {
    background-color: #1A3A6A !important;
    border-color:     #1A3A6A !important;
    color: #fff !important;
}
button[kind="primary"]:hover  { background-color: #142d54 !important; border-color: #142d54 !important; }
button[kind="primary"]:active { background-color: #0f2240 !important; }

/* ── Botões secundários ───────────────────────────────────────────── */
button[kind="secondary"] {
    border-color: #c2d4ee !important;
    color: #2d5fa0 !important;
}
button[kind="secondary"]:hover {
    border-color: #2d5fa0 !important;
    background-color: #eaf1fb !important;
    color: #1A3A6A !important;
}
button[kind="tertiary"] { color: #2d5fa0 !important; }

/* ── Link buttons (st.link_button) ───────────────────────────────── */
a[data-testid="stLinkButton"] > button,
div[data-testid="stLinkButton"] > a {
    border-color: #c2d4ee !important;
    color: #2d5fa0 !important;
    font-size: 0.82rem !important;
}
a[data-testid="stLinkButton"] > button:hover,
div[data-testid="stLinkButton"] > a:hover {
    border-color: #2d5fa0 !important;
    background-color: #eaf1fb !important;
}

/* ── Barra de progresso ───────────────────────────────────────────── */
div[data-testid="stProgress"] > div {
    background-color: #dce8f5;
    border-radius: 4px;
}
div[data-testid="stProgress"] > div > div {
    background-color: #1A3A6A;
    border-radius: 4px;
}

/* ── Metrics ──────────────────────────────────────────────────────── */
div[data-testid="metric-container"] {
    background: #f7f9fc;
    border: 1px solid #d0dcea;
    border-radius: 6px;
    padding: 0.6rem 0.8rem;
}
[data-testid="stMetricLabel"] { color: #7a8fad !important; font-size: 0.78rem !important; }
[data-testid="stMetricValue"] { color: #1a2a4a !important; }
[data-testid="stMetricDelta"]  { font-size: 0.78rem !important; }

/* ── Abas ─────────────────────────────────────────────────────────── */
button[data-baseweb="tab"] { color: #7a8fad !important; }
button[data-baseweb="tab"][aria-selected="true"] {
    color: #1A3A6A !important;
    border-bottom-color: #1A3A6A !important;
}
button[data-baseweb="tab"]:hover { color: #2d5fa0 !important; }

/* ── Expanders ────────────────────────────────────────────────────── */
div[data-testid="stExpander"] details {
    border-color: #d0dcea !important;
    border-radius: 6px !important;
    background: #f7f9fc;
}
div[data-testid="stExpander"] details summary { color: #2d5fa0; }

/* ── Inputs / Selects / Textarea ──────────────────────────────────── */
div[data-baseweb="select"] > div,
div[data-baseweb="input"]  > div {
    border-color: #c2d4ee !important;
}
div[data-baseweb="select"] > div:focus-within,
div[data-baseweb="input"]  > div:focus-within {
    border-color: #1A3A6A !important;
    box-shadow: 0 0 0 2px rgba(26,58,106,.1) !important;
}
textarea {
    border-color: #c2d4ee !important;
    border-radius: 6px !important;
}
textarea:focus-visible {
    border-color: #1A3A6A !important;
    box-shadow: 0 0 0 2px rgba(26,58,106,.1) !important;
    outline: none !important;
}

/* ── Multiselect tags ─────────────────────────────────────────────── */
span[data-baseweb="tag"] {
    background-color: #2d5fa0 !important;
    border-radius: 4px !important;
}

/* ── Radio ────────────────────────────────────────────────────────── */
div[data-testid="stRadio"] > label[data-checked="true"] > div:first-child {
    background-color: #1A3A6A !important;
    border-color:     #1A3A6A !important;
}

/* ── Alertas ──────────────────────────────────────────────────────── */
div[data-testid="stAlert"] {
    border-radius: 6px !important;
}
div[data-testid="stAlert"][data-type="info"],
[data-baseweb="notification"][kind="info"] {
    background-color: #eaf1fb !important;
    border-left-color: #1A3A6A !important;
    color: #1a2a4a !important;
}
div[data-testid="stAlert"][data-type="warning"] {
    background-color: #fffbeb !important;
    border-left-color: #d97706 !important;
}
div[data-testid="stAlert"][data-type="success"] {
    background-color: #f0fdf4 !important;
    border-left-color: #16a34a !important;
}

/* ── Spinner ──────────────────────────────────────────────────────── */
div[data-testid="stSpinner"] svg { stroke: #1A3A6A; }

/* ══════════════════════════════════════════════════════════════════
   ac-* — componentes de card custom
   ══════════════════════════════════════════════════════════════════ */
.ac-card {
    border: 1px solid #d0dcea;
    border-radius: 6px;
    overflow: hidden;
    margin-bottom: 0.75rem;
    font-family: inherit;
}
.ac-card-header {
    background: #1A3A6A;
    color: #fff;
    padding: 0.5rem 0.9rem;
}
.ac-card-header-light {
    background: #eaf1fb;
    color: #1A3A6A;
    padding: 0.5rem 0.9rem;
    border-bottom: 1px solid #d0dcea;
}
.ac-card-body {
    background: #f7f9fc;
    padding: 0.65rem 0.9rem 0.3rem;
}
.ac-label {
    font-size: 0.67rem;
    font-weight: 700;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    color: #7a8fad;
    margin-bottom: 1px;
}
.ac-label-dark { color: #a8bcd4; }
.ac-value {
    font-weight: 600;
    font-size: 0.85rem;
    color: #1a2a4a;
    margin-bottom: 0.5rem;
}
.ac-value-mono { font-family: monospace; font-size: 0.82rem; }
.ac-badge {
    display: inline-block;
    background: #1A3A6A;
    color: #fff;
    font-size: 0.72rem;
    font-weight: 700;
    letter-spacing: 0.05em;
    text-transform: uppercase;
    border-radius: 4px;
    padding: 2px 8px;
    margin-right: 4px;
}
.ac-badge-light {
    background: #eaf1fb;
    color: #1A3A6A;
    border: 1px solid #c2d4ee;
}
</style>
""", unsafe_allow_html=True)




# ---------------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------------

PAGINAS_TRIAGEM = {
    "importacao":  ("📂", "1. Importação"),
    "triadas":     ("✅", "2. Tarefas Triadas"),
    "nao_triadas": ("🔍", "3. Tarefas Não Triadas"),
    "relatorio":   ("📄", "4. Relatório"),
}

PAGINAS_DISTRIBUICAO = {
    "importacao":    ("📂", "1. Importação"),
    "distribuicao":  ("📊", "2. Auditoria de Distribuição"),
    "relatorio":     ("📄", "3. Relatório"),
}


def _get_paginas() -> dict:
    tipo = st.session_state.get("tipo_relatorio")
    if tipo == "supp_distribuicao":
        return PAGINAS_DISTRIBUICAO
    return PAGINAS_TRIAGEM


def _check_icon(chave: str) -> str:
    checks = {
        "importacao": (
            st.session_state.get("audit_data_merged") is not None
            or st.session_state.get("dist_data") is not None
        ),
        "triadas":       st.session_state.get("auditoria_triadas_concluida", False),
        "nao_triadas":   st.session_state.get("auditoria_nao_triadas_concluida", False),
        "distribuicao":  st.session_state.get("auditoria_distribuicao_concluida", False),
        "relatorio":     False,
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
              "supp_login_step", "supp_totp_challenge", "supp_username_pendente"):
        st.session_state.pop(k, None)


def _supp_erro_msg(exc: Exception) -> str:
    from modules.auth import AuthError
    if isinstance(exc, AuthError):
        body = exc.body
        if isinstance(body, dict):
            return str(body.get("message") or body.get("error") or body)
        return str(body)
    return str(exc)


def _supp_finalizar_login(client, usuario: str, base_url: str) -> None:
    """Sessão autenticada: guarda o cliente e limpa o estado do 2FA."""
    st.session_state["supp_auth_client"] = client
    st.session_state["supp_logged_in"] = True
    st.session_state["supp_username"] = _supp_get_nome(client, usuario)
    st.session_state["supp_base_url"] = base_url
    st.session_state.pop("supp_totp_challenge", None)
    st.rerun()


def _supp_do_login(base_url: str, usuario: str, senha: str) -> None:
    """Etapa 1 — autentica via LDAP; se o SUPP exigir 2FA, guarda o challenge."""
    from modules.auth import AuthClient, TotpChallenge

    base_url = base_url.rstrip("/")
    client = AuthClient(base_url=base_url)
    try:
        client.login_ldap(usuario, senha)
    except TotpChallenge as totp:
        st.session_state["supp_auth_client"] = client
        st.session_state["supp_totp_challenge"] = totp.challenge
        st.session_state["supp_username_pendente"] = usuario
        st.session_state["supp_base_url"] = base_url
        try:
            client.totp_send_mail(totp.challenge)
        except Exception:
            # O código do app autenticador continua válido mesmo sem o e-mail.
            pass
        st.rerun()
    except Exception as exc:
        client.close()
        st.error(f"Credenciais inválidas: {_supp_erro_msg(exc)}")
        return

    _supp_finalizar_login(client, usuario, base_url)


def _supp_do_totp(codigo: str) -> None:
    """Etapa 2 — verifica o código de 6 dígitos e obtém o JWT final."""
    client = st.session_state.get("supp_auth_client")
    challenge = st.session_state.get("supp_totp_challenge")
    usuario = st.session_state.get("supp_username_pendente", "")
    if not client or not challenge:
        _supp_logout_cleanup()
        st.error("Sessão de login expirada. Refaça o login.")
        return

    try:
        client.totp_verify(challenge, codigo)
    except Exception as exc:
        st.error(f"Código 2FA inválido: {_supp_erro_msg(exc)}")
        return

    st.session_state.pop("supp_username_pendente", None)
    _supp_finalizar_login(client, usuario, st.session_state["supp_base_url"])


def _render_login_page() -> None:
    """Tela de login SUPP via LDAP — exibida antes de qualquer outra coisa."""
    st.markdown(
        "<style>.block-container{padding-top:5rem;}</style>",
        unsafe_allow_html=True,
    )
    _, col, _ = st.columns([1, 1.2, 1])
    with col:
        st.markdown(
            "<div class='ac-card'>"
            "<div class='ac-card-header' style='padding:1.4rem 1.4rem 1.1rem'>"
            "<div style='font-size:1.4rem;font-weight:800;letter-spacing:0.01em;margin-bottom:3px'>"
            "📋 Auditoria Conecta+</div>"
            "<div style='font-size:0.78rem;opacity:0.72'>Procuradoria-Geral Federal · AGU</div>"
            "</div>"
            "<div class='ac-card-body' style='padding:1.4rem 1.4rem 0.5rem'>",
            unsafe_allow_html=True,
        )

        SUPP_URL = "https://supersapiensbackend.agu.gov.br"

        if st.session_state.get("supp_totp_challenge"):
            st.caption(
                "Verificação em duas etapas — informe o código de 6 dígitos "
                "do seu app autenticador (também enviado por e-mail)."
            )
            with st.form("login_totp_form"):
                codigo = st.text_input(
                    "Código 2FA", max_chars=6, placeholder="000000"
                )
                ok = st.form_submit_button(
                    "Verificar →", use_container_width=True, type="primary"
                )
            voltar = st.button("← Voltar", use_container_width=True)

            st.markdown("</div></div>", unsafe_allow_html=True)

            if voltar:
                _supp_logout_cleanup()
                st.rerun()
            if ok:
                if not codigo.strip():
                    st.error("Informe o código de 6 dígitos.")
                else:
                    _supp_do_totp(codigo.strip())
            return

        with st.form("login_page_form"):
            usuario = st.text_input("Login (Rede AGU)", placeholder="login")
            senha = st.text_input("Senha", type="password")
            ok = st.form_submit_button(
                "Entrar →", use_container_width=True, type="primary"
            )

        st.markdown("</div></div>", unsafe_allow_html=True)

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
    for chave, (icone, label) in _get_paginas().items():
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
    dd = get_dist_data()
    tipo_rel = st.session_state.get("tipo_relatorio")

    if ad and tipo_rel == "conecta_triagem":
        df_tri = st.session_state.get("df_audit_triadas")
        df_nao = st.session_state.get("df_audit_nao_triadas")
        n_aud = n_total = 0
        if df_tri is not None:
            n_total += len(df_tri)
            n_aud += len(df_tri[df_tri[COL_CONFORMIDADE] != OPCOES_CONFORMIDADE[0]])
        if df_nao is not None:
            n_total += len(df_nao)
            n_aud += len(df_nao[df_nao[COL_CONFORMIDADE] != OPCOES_CONFORMIDADE[0]])

        pct_str = f"{n_aud/n_total*100:.0f}%" if n_total > 0 else "—"
        prog_bar = (
            f"<div style='background:#d0dcea;border-radius:4px;height:5px;margin-top:4px'>"
            f"<div style='background:#1A3A6A;width:{n_aud/n_total*100 if n_total else 0:.1f}%;"
            f"height:5px;border-radius:4px'></div></div>"
            if n_total > 0 else ""
        )
        st.markdown(
            f"<div class='ac-card'>"
            f"<div class='ac-card-header' style='padding:0.4rem 0.75rem;font-size:0.78rem;"
            f"white-space:nowrap;overflow:hidden;text-overflow:ellipsis'>📁 {ad.nome_arquivo}</div>"
            f"<div class='ac-card-body' style='padding:0.5rem 0.75rem'>"
            f"<div style='display:flex;justify-content:space-between;font-size:0.78rem;margin-bottom:0.3rem'>"
            f"<span><span class='ac-label'>Total</span><br><strong>{ad.total_tarefas}</strong></span>"
            f"<span><span class='ac-label'>Triadas</span><br><strong>{ad.total_triadas}</strong></span>"
            f"<span><span class='ac-label'>Não triadas</span><br><strong>{ad.total_nao_triadas}</strong></span>"
            f"<span><span class='ac-label'>Auditadas</span><br><strong>{pct_str}</strong></span>"
            f"</div>"
            f"{prog_bar}"
            f"</div></div>",
            unsafe_allow_html=True,
        )
        st.divider()

    elif dd and tipo_rel == "supp_distribuicao":
        df_dist = st.session_state.get("df_audit_distribuicao")
        n_aud = n_total = 0
        if df_dist is not None:
            n_total = len(df_dist)
            n_aud = len(df_dist[df_dist[COL_CONFORMIDADE] != OPCOES_CONFORMIDADE[0]])
        pct_str = f"{n_aud/n_total*100:.0f}%" if n_total > 0 else "—"
        prog_bar = (
            f"<div style='background:#d0dcea;border-radius:4px;height:5px;margin-top:4px'>"
            f"<div style='background:#1A3A6A;width:{n_aud/n_total*100 if n_total else 0:.1f}%;"
            f"height:5px;border-radius:4px'></div></div>"
            if n_total > 0 else ""
        )
        n_total_dist = dd.total_distribuicoes
        st.markdown(
            f"<div class='ac-card'>"
            f"<div class='ac-card-header' style='padding:0.4rem 0.75rem;font-size:0.78rem;"
            f"white-space:nowrap;overflow:hidden;text-overflow:ellipsis'>📊 {dd.nome_arquivo}</div>"
            f"<div class='ac-card-body' style='padding:0.5rem 0.75rem'>"
            f"<div style='display:flex;justify-content:space-between;font-size:0.78rem;margin-bottom:0.3rem'>"
            f"<span><span class='ac-label'>Distribuições</span><br><strong>{n_total_dist}</strong></span>"
            f"<span><span class='ac-label'>Na amostra</span><br><strong>{n_total if n_total else '—'}</strong></span>"
            f"<span><span class='ac-label'>Auditadas</span><br><strong>{pct_str}</strong></span>"
            f"</div>"
            f"{prog_bar}"
            f"</div></div>",
            unsafe_allow_html=True,
        )
        st.divider()

    if st.button("🔄 Nova Auditoria", use_container_width=True):
        for k in list(st.session_state.keys()):
            if k.startswith(("filtro_", "busca_", "tbl_",
                             "edit_conf_", "edit_motivo_", "edit_acao_", "btn_save_row_",
                             "_proc_id_cache_", "_supp_cache_")):
                del st.session_state[k]
        reset_auditoria()
        clear_saved_session()
        st.session_state["pagina"] = "importacao"
        st.session_state["audit_data_merged"] = None
        st.session_state["dist_data"] = None
        st.rerun()

    # ── Usuário SUPP ──────────────────────────────────────────────────────────
    st.divider()
    nome = st.session_state.get("supp_username", "")
    st.markdown(
        f"<div style='font-size:0.78rem;margin-bottom:0.4rem'>"
        f"<span style='color:#2ecc71'>●</span> "
        f"<span style='font-weight:600;color:#1a2a4a'>{nome}</span></div>",
        unsafe_allow_html=True,
    )
    if st.button("Sair do SUPP", use_container_width=True, key="btn_supp_logout"):
        _supp_logout_cleanup()
        st.rerun()


# ---------------------------------------------------------------------------
# SUPP — Link para visualizar processo
# ---------------------------------------------------------------------------

_SUPERSAPIENS_URL = "https://supersapiens.agu.gov.br/apps/processo/{proc_id}/visualizar/capa"



def _render_processo_link() -> None:
    """Placeholder exibido quando nenhuma linha está selecionada na tabela."""
    if not st.session_state.get("supp_sel_tarefa_id"):
        st.markdown(
            "<div class='ac-card'>"
            "<div class='ac-card-body' style='padding:1.2rem 0.9rem;text-align:center;"
            "color:#7a8fad;font-size:0.85rem'>"
            "← Selecione uma linha na tabela para auditar e abrir o processo no SuperSapiens."
            "</div></div>",
            unsafe_allow_html=True,
        )


# ---------------------------------------------------------------------------
# Tabela interativa + editor de linha
# ---------------------------------------------------------------------------


def _render_audit_table(
    df_key: str,
    filtro_key: str,
    busca_key: str,
    column_order: list[str],
    table_key: str,
    id_col: str = COL_TAREFA,
    nup_col: str = COL_NUP,
) -> tuple:
    """
    Tabela interativa com filtros e seleção de linha.
    Retorna (orig_idx, row_dict) da linha selecionada, ou (None, None).
    id_col / nup_col permitem reutilizar a tabela no modo distribuição.
    """
    df = st.session_state[df_key]
    total = len(df)
    s = stats_df(df)

    pct = s["auditadas"] / total if total > 0 else 0
    entidade = "distribuições" if id_col == COL_DIST_ID else "tarefas"
    st.progress(
        pct,
        text=(
            f"**{s['auditadas']}/{total}** {entidade} auditadas"
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
        has_setor_destino = COL_DIST_SETOR_DESTINO in df.columns
        if has_setor_destino:
            busca_label = "Buscar (Id, NUP ou Setor Destino):"
        elif has_config:
            busca_label = "Buscar (Tarefa, NUP ou Config.):"
        else:
            busca_label = "Buscar (Tarefa ou NUP):"
        busca = st.text_input(busca_label, key=busca_key, placeholder="Digite para filtrar…")

    mask = df[COL_CONFORMIDADE].isin(filtro)
    if busca.strip():
        txt = busca.strip()
        id_col_search = id_col if id_col in df.columns else None
        nup_col_search = nup_col if nup_col in df.columns else None
        search_mask = pd.Series(False, index=df.index)
        if id_col_search:
            search_mask = search_mask | df[id_col_search].astype(str).str.contains(txt, case=False, na=False)
        if nup_col_search:
            search_mask = search_mask | df[nup_col_search].astype(str).str.contains(txt, case=False, na=False)
        if has_config:
            search_mask = search_mask | df[COL_CONFIG].astype(str).str.contains(txt, case=False, na=False)
        if has_setor_destino:
            search_mask = search_mask | df[COL_DIST_SETOR_DESTINO].astype(str).str.contains(txt, case=False, na=False)
        mask = mask & search_mask

    df_view = df.loc[mask]
    col_order = [c for c in column_order if c in df_view.columns]

    label_entidade = "distribuições" if id_col == COL_DIST_ID else "tarefas"
    st.caption(f"Exibindo **{len(df_view)}** de {total} {label_entidade} — clique em uma linha para auditar")

    if df_view.empty:
        st.info("Nenhum registro corresponde ao filtro atual.")
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
            COL_DIST_ID: st.column_config.TextColumn("Id", width="small"),
            COL_DIST_NUP: st.column_config.TextColumn("NUP", width="medium"),
            COL_DIST_PROCESSO_JUDICIAL: st.column_config.TextColumn("CNJ", width="medium"),
            COL_DIST_FONTE_DADOS: st.column_config.TextColumn("Fonte", width="medium"),
            COL_DIST_SETOR_ORIGEM: st.column_config.TextColumn("Setor Origem", width="medium"),
            COL_DIST_USUARIO_DESTINO: st.column_config.TextColumn("Usuário Destino", width="medium"),
            COL_DIST_SETOR_DESTINO: st.column_config.TextColumn("Setor Destino", width="large"),
            COL_DIST_DATA_HORA: st.column_config.TextColumn("Data/Hora", width="small"),
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
        st.session_state["supp_sel_tarefa_id"] = row.get(id_col)
        return orig_idx, row

    return None, None


def _render_row_editor(df_key: str, orig_idx, row: dict) -> None:
    """Painel de edição dos campos de auditoria + link para o processo no SuperSapiens."""
    tarefa_id = row.get(COL_TAREFA)
    nup = row.get(COL_NUP)

    # ── Busca dados do processo e atividades (com cache por tarefa) ─────────
    cache_key = f"_proc_id_cache_{tarefa_id}"
    if cache_key not in st.session_state:
        auth = st.session_state.get("supp_auth_client")
        if auth:
            with st.spinner("Buscando processo e atividades..."):
                try:
                    from modules.tarefa import TarefaClient
                    tc = TarefaClient.from_auth(auth)
                    tarefa = tc.buscar(tarefa_id, populate=[
                        "processo", "especieTarefa", "usuarioResponsavel",
                        "setorResponsavel", "setorOrigem", "vinculacaoWorkflow",
                    ])
                    proc = tarefa.get("processo") or {}
                    proc_id = proc.get("id")
                    nup_fmt = proc.get("NUPFormatado") or proc.get("NUP")

                    # CNJ e classe nacional estão em processo.any.processoJudicial
                    any_ = proc.get("any") or {}
                    pj = any_.get("processoJudicial") or {}
                    cnj = pj.get("numeroFormatado") or pj.get("numero")
                    cn = pj.get("classeNacional") or {}
                    classe_nacional = cn.get("nome") if isinstance(cn, dict) else None

                    # Parte representada em processo.any.pessoaRepresentada.pessoa.nome
                    pr = any_.get("pessoaRepresentada") or {}
                    pessoa = pr.get("pessoa") or {}
                    parte = pessoa.get("nome")

                    # Metadados de workflow da tarefa
                    et = tarefa.get("especieTarefa") or {}
                    especie_tarefa = et.get("nome")
                    ur = tarefa.get("usuarioResponsavel") or {}
                    usuario_resp = ur.get("nome") or ur.get("username")
                    sr = tarefa.get("setorResponsavel") or {}
                    setor_resp = sr.get("nome")
                    setor_resp_sigla = sr.get("sigla")
                    so = tarefa.get("setorOrigem") or {}
                    setor_orig = so.get("nome")

                    # Situação: encerrada se tiver dataHoraEncerramento
                    data_encerramento = tarefa.get("dataHoraEncerramento")

                    def _fmt_dt(s: str) -> str:
                        if not s:
                            return ""
                        try:
                            from datetime import datetime as _dt
                            return _dt.fromisoformat(
                                str(s).replace("Z", "+00:00")
                            ).strftime("%d/%m/%Y %H:%M")
                        except Exception:
                            return str(s)[:16]

                    def _obj_nome(obj) -> str | None:
                        """Extrai nome/username/sigla de objeto populado."""
                        if isinstance(obj, dict):
                            return (
                                obj.get("nome")
                                or obj.get("username")
                                or obj.get("sigla")
                                or None
                            )
                        return None

                    # Busca atividades via AtividadeClient (paginação automática)
                    # IDs vindos do Excel podem ser float (ex: 12345.0) — normalizar para int
                    _tid = int(float(tarefa_id)) if tarefa_id is not None else None
                    atividades = []
                    _ativ_erro = None
                    if _tid:
                        try:
                            from modules.atividade import (
                                AtividadeClient,
                                BASE_PATH_JUDICIAL,
                                BASE_PATH_CONSULTIVO,
                                BASE_PATH_ADMINISTRATIVO,
                            )
                            # Tenta os endpoints em ordem de relevância:
                            # judicial → consultivo → administrativo
                            # Cada tentativa é independente; erros são ignorados
                            # para tentar o próximo (ex: endpoint não aplicável).
                            _ativ_raw = []
                            for _path in (
                                BASE_PATH_JUDICIAL,
                                BASE_PATH_CONSULTIVO,
                                BASE_PATH_ADMINISTRATIVO,
                            ):
                                try:
                                    with AtividadeClient(
                                        token=auth.token,
                                        base_url=auth.base_url,
                                        base_path=_path,
                                    ) as _ac:
                                        _ativ_raw = _ac.listar_por_tarefa(_tid)
                                except Exception:
                                    continue
                                if _ativ_raw:
                                    break

                            def _norm_atividade(raw: dict) -> dict:
                                # Suporte a atividade_judicial: campo pode vir
                                # diretamente ou aninhado em .atividade
                                base = raw.get("atividade") or raw
                                ea = raw.get("especieAtividade") or base.get("especieAtividade") or {}
                                au = raw.get("usuario") or base.get("usuario") or {}
                                as_ = raw.get("setor") or base.get("setor") or {}
                                desc = (
                                    raw.get("observacao")
                                    or raw.get("descricao")
                                    or base.get("observacao")
                                    or base.get("descricao")
                                    or ""
                                )
                                dt = (
                                    raw.get("dataHoraConclusao")
                                    or raw.get("criadoEm")
                                    or base.get("dataHoraConclusao")
                                    or base.get("criadoEm")
                                )
                                encerra = bool(
                                    raw.get("encerraTarefa")
                                    or base.get("encerraTarefa")
                                )
                                return {
                                    "especie":        _obj_nome(ea) or "—",
                                    "usuario":        _obj_nome(au) or "—",
                                    "setor":          as_.get("sigla") or as_.get("nome") or "",
                                    "setor_nome":     as_.get("nome") or "",
                                    "data":           _fmt_dt(dt),
                                    "descricao":      desc,
                                    "encerra_tarefa": encerra,
                                }

                            atividades = [_norm_atividade(_a) for _a in _ativ_raw]
                        except Exception as _e:
                            _ativ_erro = str(_e)

                    st.session_state[cache_key] = {
                        "proc_id":          proc_id,
                        "nup_fmt":          nup_fmt,
                        "cnj":              cnj,
                        "classe_nacional":  classe_nacional,
                        "parte":            parte,
                        "especie_tarefa":   especie_tarefa,
                        "usuario_resp":     usuario_resp,
                        "setor_resp":       setor_resp,
                        "setor_resp_sigla": setor_resp_sigla,
                        "setor_orig":       setor_orig,
                        "data_encerramento": _fmt_dt(data_encerramento) if data_encerramento else None,
                        "atividades":       atividades,
                        "atividades_erro":  _ativ_erro if not atividades else None,
                    }
                except Exception as e:
                    st.session_state[cache_key] = {"erro": str(e)}
        else:
            st.session_state[cache_key] = {}

    cached = st.session_state.get(cache_key, {})
    proc_id = cached.get("proc_id")
    nup_fmt = cached.get("nup_fmt") or nup
    cnj = cached.get("cnj")
    classe_nacional = cached.get("classe_nacional")
    parte = cached.get("parte")
    especie_tarefa = cached.get("especie_tarefa")
    usuario_resp = cached.get("usuario_resp")
    setor_resp = cached.get("setor_resp")
    setor_resp_sigla = cached.get("setor_resp_sigla")
    setor_orig = cached.get("setor_orig")
    data_encerramento = cached.get("data_encerramento")
    atividades = cached.get("atividades", [])
    atividades_erro = cached.get("atividades_erro")

    # ── Cabeçalho com dados identificadores ──────────────────────────────────
    # ── Título + botões discretos alinhados à direita ────────────────────────
    url = _SUPERSAPIENS_URL.format(proc_id=proc_id) if proc_id else None
    _c_title, _c_open, _c_refresh = st.columns([4, 3, 1])
    with _c_title:
        st.markdown("#### ✏️ Auditoria")
    with _c_open:
        if url:
            st.link_button("↗ SuperSapiens", url, use_container_width=True)
    with _c_refresh:
        if st.button("🔄", key=f"_refresh_proc_{tarefa_id}", help="Recarregar dados do processo"):
            st.session_state.pop(cache_key, None)
            st.rerun()

    # ── Card de identificação ─────────────────────────────────────────────────
    def _field(label: str, value, mono: bool = False) -> str:
        if not value:
            return ""
        val_style = (
            "font-family:monospace;font-size:0.82rem;color:#1a2a4a"
            if mono else
            "font-size:0.85rem;color:#1a2a4a"
        )
        return (
            f"<div style='margin-bottom:0.55rem'>"
            f"<div style='font-size:0.67rem;font-weight:700;letter-spacing:0.08em;"
            f"text-transform:uppercase;color:#7a8fad;margin-bottom:1px'>{label}</div>"
            f"<div style='font-weight:600;{val_style}'>{value}</div>"
            f"</div>"
        )

    ids_html = ""
    if tarefa_id or proc_id:
        id_t = (
            f"<div style='flex:1'>"
            f"<div style='font-size:0.67rem;font-weight:700;letter-spacing:0.08em;"
            f"text-transform:uppercase;color:#a8bcd4;margin-bottom:1px'>Id Tarefa</div>"
            f"<div style='font-weight:700;font-size:0.92rem;font-family:monospace'>{tarefa_id or '—'}</div>"
            f"</div>"
        )
        id_p = (
            f"<div style='flex:1'>"
            f"<div style='font-size:0.67rem;font-weight:700;letter-spacing:0.08em;"
            f"text-transform:uppercase;color:#a8bcd4;margin-bottom:1px'>Id Processo</div>"
            f"<div style='font-weight:700;font-size:0.92rem;font-family:monospace'>{proc_id or '—'}</div>"
            f"</div>"
        )
        ids_html = (
            f"<div style='background:#1A3A6A;color:#fff;padding:0.55rem 0.9rem;"
            f"display:flex;gap:1.5rem;border-radius:6px 6px 0 0'>{id_t}{id_p}</div>"
        )

    body_html = (
        _field("NUP", nup_fmt, mono=True)
        + _field("Número CNJ", cnj, mono=True)
        + _field("Classe Nacional", classe_nacional)
        + _field("Entidade Representada", parte)
    )

    st.markdown(
        f"<div style='border:1px solid #d0dcea;border-radius:6px;"
        f"overflow:hidden;margin-bottom:0.7rem'>"
        f"{ids_html}"
        f"<div style='background:#f7f9fc;padding:0.65rem 0.9rem 0.25rem'>{body_html}</div>"
        f"</div>",
        unsafe_allow_html=True,
    )
    if cached.get("erro"):
        st.caption(f"⚠️ Erro ao buscar processo: {cached['erro']}")

    # ── Fluxo de Trabalho ─────────────────────────────────────────────────────
    _tarefa_info_rows = ""
    if especie_tarefa or usuario_resp or setor_resp or setor_orig or data_encerramento is not None:
        def _fi(lbl: str, val: str) -> str:
            if not val:
                return ""
            return (
                f"<div style='min-width:0'>"
                f"<div style='font-size:0.63rem;font-weight:700;letter-spacing:0.08em;"
                f"text-transform:uppercase;color:#7a8fad;margin-bottom:1px'>{lbl}</div>"
                f"<div style='font-size:0.8rem;font-weight:600;color:#1a2a4a;"
                f"white-space:nowrap;overflow:hidden;text-overflow:ellipsis'>{val}</div>"
                f"</div>"
            )

        # Badge de situação: ABERTA (verde) ou ENCERRADA (vermelho+data)
        if data_encerramento:
            _status_badge = (
                f"<span style='background:#dc2626;color:#fff;font-size:0.65rem;"
                f"font-weight:700;padding:2px 8px;border-radius:3px;"
                f"letter-spacing:0.04em'>ENCERRADA</span>"
                f"<span style='font-size:0.75rem;color:#7a8fad;margin-left:6px'>"
                f"em {data_encerramento}</span>"
            )
        else:
            _status_badge = (
                "<span style='background:#16a34a;color:#fff;font-size:0.65rem;"
                "font-weight:700;padding:2px 8px;border-radius:3px;"
                "letter-spacing:0.04em'>ABERTA</span>"
            )

        # Label do setor responsável: "SIGLA — Nome" quando há sigla
        _setor_resp_label = ""
        if setor_resp:
            _setor_resp_label = (
                f"{setor_resp_sigla} — {setor_resp}"
                if setor_resp_sigla and setor_resp_sigla != setor_resp
                else setor_resp
            )

        _tarefa_info_rows = (
            f"<div style='margin-bottom:0.55rem'>{_status_badge}</div>"
            f"<div style='display:flex;gap:0.9rem;flex-wrap:wrap;margin-bottom:0.55rem'>"
            + _fi("Tipo de Tarefa", especie_tarefa or "")
            + _fi("Responsável", usuario_resp or "")
            + _fi("Setor Responsável", _setor_resp_label)
            + _fi("Setor Origem", setor_orig or "")
            + "</div>"
        )

    _ativ_html = ""
    if atividades:
        _items = ""
        _n = len(atividades)
        for _i, _a in enumerate(atividades):
            _encerra = _a.get("encerra_tarefa")
            _border = "#dc2626" if _encerra else (
                "#1A3A6A" if _i == 0 else "#2d5fa0" if _i < _n - 1 else "#7a8fad"
            )
            _setor = _a.get("setor") or ""
            _setor_nome = _a.get("setor_nome") or ""

            # Linha de execução: usuário · data
            _meta = " · ".join(filter(None, [_a["usuario"], _a["data"]]))

            # Setor da atividade — destacado para a atividade que encerrou a tarefa
            if _setor:
                _tooltip = f' title="{_setor_nome}"' if _setor_nome and _setor_nome != _setor else ""
                if _encerra:
                    _setor_html = (
                        f"<span style='display:inline-block;background:#fef2f2;"
                        f"color:#dc2626;border:1px solid #fca5a5;font-size:0.68rem;"
                        f"font-weight:700;padding:1px 6px;border-radius:3px;"
                        f"margin-top:2px'{_tooltip}>fechada em: {_setor}</span>"
                    )
                else:
                    _setor_html = (
                        f"<span style='display:inline-block;background:#eaf1fb;"
                        f"color:#1A3A6A;border:1px solid #c2d4ee;font-size:0.68rem;"
                        f"padding:1px 6px;border-radius:3px;"
                        f"margin-top:2px'{_tooltip}>{_setor}</span>"
                    )
            else:
                _setor_html = ""

            # Badge ENCERRA
            _encerra_badge = (
                "<span style='display:inline-block;background:#dc2626;color:#fff;"
                "font-size:0.58rem;padding:0 4px;border-radius:3px;margin-left:5px;"
                "font-weight:700;vertical-align:middle'>ENCERRA</span>"
                if _encerra else ""
            )
            _desc = (
                f"<div style='font-size:0.72rem;color:#7a8fad;margin-top:1px'>{_a['descricao']}</div>"
                if _a.get("descricao") else ""
            )
            _items += (
                f"<div style='border-left:3px solid {_border};padding:0.15rem 0 0.25rem 0.65rem;"
                f"margin-bottom:0.45rem'>"
                f"<div style='font-size:0.8rem;font-weight:700;color:#1a2a4a'>"
                f"{_a['especie']}{_encerra_badge}</div>"
                f"<div style='font-size:0.72rem;color:#7a8fad'>{_meta}</div>"
                f"{_setor_html}"
                f"{_desc}"
                f"</div>"
            )
        _ativ_html = (
            f"<div style='border-top:1px solid #d0dcea;margin-top:0.4rem;padding-top:0.55rem'>"
            f"<div style='font-size:0.63rem;font-weight:700;letter-spacing:0.08em;"
            f"text-transform:uppercase;color:#7a8fad;margin-bottom:0.45rem'>"
            f"Atividades ({_n})</div>"
            f"{_items}"
            f"</div>"
        )
    elif atividades_erro:
        _ativ_html = (
            f"<div style='border-top:1px solid #d0dcea;margin-top:0.4rem;padding-top:0.45rem'>"
            f"<div style='font-size:0.72rem;color:#e05a3a;font-style:italic'>"
            f"Erro ao carregar atividades: {atividades_erro}</div></div>"
        )
    elif cached.get("erro"):
        _ativ_html = (
            "<div style='border-top:1px solid #d0dcea;margin-top:0.4rem;padding-top:0.45rem'>"
            "<div style='font-size:0.72rem;color:#e05a3a;font-style:italic'>"
            "Não foi possível carregar as atividades.</div></div>"
        )
    else:
        _ativ_html = (
            "<div style='border-top:1px solid #d0dcea;margin-top:0.4rem;padding-top:0.45rem'>"
            "<div style='font-size:0.75rem;color:#7a8fad;font-style:italic'>"
            "Nenhuma atividade registrada.</div></div>"
        )

    st.markdown(
        f"<div style='border:1px solid #d0dcea;border-radius:6px;overflow:hidden;"
        f"margin-bottom:0.7rem'>"
        f"<div style='background:#eaf1fb;color:#1A3A6A;padding:0.35rem 0.9rem;"
        f"font-size:0.72rem;font-weight:700;letter-spacing:0.06em;"
        f"text-transform:uppercase;border-bottom:1px solid #d0dcea'>Fluxo de Trabalho</div>"
        f"<div style='background:#f7f9fc;padding:0.6rem 0.9rem 0.3rem'>"
        f"{_tarefa_info_rows}"
        f"{_ativ_html}"
        f"</div></div>",
        unsafe_allow_html=True,
    )

    # ── Campos de auditoria ───────────────────────────────────────────────────
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
        save_session()
        st.rerun()


# ===========================================================================
# PÁGINA 1 — IMPORTAÇÃO
# ===========================================================================

def render_importacao() -> None:
    st.title("📂 Importação de Arquivo")

    # ── Restauração de sessão anterior ───────────────────────────────────────
    if has_saved_session() and not st.session_state.get("session_restore_offered"):
        info = get_session_info()
        if info and info.get("nome_arquivo"):
            tipo_label = (
                "Distribuição SS" if info.get("tipo_relatorio") == "supp_distribuicao"
                else "Triagem Conecta+"
            )
            st.info(
                f"**Sessão salva encontrada** — {tipo_label}: _{info['nome_arquivo']}_  \n"
                "Deseja continuar de onde parou?"
            )
            c_sim, c_nao, _ = st.columns([1, 1, 3])
            with c_sim:
                if st.button("▶ Continuar sessão", type="primary", use_container_width=True):
                    if load_session():
                        st.session_state["session_restore_offered"] = True
                        st.rerun()
            with c_nao:
                if st.button("✕ Descartar", use_container_width=True):
                    clear_saved_session()
                    st.session_state["session_restore_offered"] = True
                    st.rerun()
            st.divider()

    # ── Tipo de relatório ─────────────────────────────────────────────────────
    tipo_rel_atual = st.session_state.get("tipo_relatorio")
    tipo_opcoes = [
        "Conecta+ — Triagem Avançada",
        "Super Sapiens — Distribuição de Tarefas",
    ]
    tipo_idx = 1 if tipo_rel_atual == "supp_distribuicao" else 0
    tipo_sel = st.radio(
        "Tipo de relatório:",
        tipo_opcoes,
        index=tipo_idx,
        horizontal=True,
        key="radio_tipo_rel",
    )
    novo_tipo = "supp_distribuicao" if tipo_sel.startswith("Super") else "conecta_triagem"
    if novo_tipo != tipo_rel_atual:
        reset_auditoria()
        st.session_state["tipo_relatorio"] = novo_tipo
        st.session_state["audit_data_merged"] = None
        st.session_state["dist_data"] = None

    st.divider()

    if novo_tipo == "conecta_triagem":
        _render_importacao_triagem()
    else:
        _render_importacao_distribuicao()


def _render_importacao_triagem() -> None:
    """Sub-renderização da importação no modo Conecta+ Triagem."""
    col_up, col_info = st.columns([2, 1])
    with col_up:
        uploaded = st.file_uploader(
            "Selecione o(s) arquivo(s) Excel (.xlsx):",
            type=["xlsx"],
            accept_multiple_files=True,
            help="O arquivo deve conter as abas: Todas as Tarefas, Tarefas Triadas e Tarefas Não Triadas.",
            key="uploader_triagem",
        )
    with col_info:
        st.markdown(
            "<div class='ac-card'>"
            "<div class='ac-card-header-light'>"
            "<span style='font-size:0.8rem;font-weight:700'>Formato esperado</span>"
            "</div>"
            "<div class='ac-card-body' style='font-size:0.82rem'>"
            "<div class='ac-label' style='margin-bottom:4px'>Abas obrigatórias</div>"
            "<div class='ac-value'>Todas as Tarefas<br>Tarefas Triadas<br>Tarefas Não Triadas</div>"
            "<div class='ac-label' style='margin-bottom:4px'>Colunas</div>"
            "<div class='ac-value' style='font-size:0.78rem'>ID · Tarefa · NUP · Usuário<br>"
            "Datas · Status · Config. Encontradas</div>"
            "</div></div>",
            unsafe_allow_html=True,
        )

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

    if merged.datas_corrigidas:
        st.warning(
            f"**{merged.datas_corrigidas} datas foram corrigidas na importação.** "
            "A planilha traz parte das datas invertidas no padrão americano "
            "(mês/dia) — por exemplo, 1º de julho gravado como 7 de janeiro. "
            "O app desfez a inversão com base nas datas não ambíguas do próprio "
            "arquivo. Confira o período abaixo antes de prosseguir; se ele não "
            "corresponder ao que você exportou, gere um novo relatório no "
            "Conecta+ e reimporte."
        )

    if len(audit_files) > 1:
        st.info(
            f"{len(audit_files)} arquivos consolidados. "
            "Registros duplicados foram removidos (mantida a ocorrência mais recente)."
        )

    atual = st.session_state.get("audit_data_merged")
    if atual is not None and atual.nome_arquivo != merged.nome_arquivo:
        reset_auditoria()

    st.session_state["audit_data_merged"] = merged
    st.session_state["tipo_relatorio"] = "conecta_triagem"
    save_session()

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
        f"<div class='ac-card'>"
        f"<div class='ac-card-header' style='display:flex;align-items:center;gap:0.6rem'>"
        f"<span style='font-size:1rem'>📅</span>"
        f"<div><div class='ac-label ac-label-dark'>Período de triagem</div>"
        f"<div style='font-weight:600;font-size:0.92rem'>{periodo_str}</div></div>"
        f"</div></div>",
        unsafe_allow_html=True,
    )

    c1, c2, c3 = st.columns(3)
    c1.metric("Total de Tarefas", merged.total_tarefas)
    c2.metric("Tarefas Triadas", merged.total_triadas,
              delta=f"{merged.pct_triadas:.1f}% do total", delta_color="normal")
    c3.metric("Tarefas Não Triadas", merged.total_nao_triadas,
              delta=f"{merged.pct_nao_triadas:.1f}% do total", delta_color="off")

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


def _render_importacao_distribuicao() -> None:
    """Sub-renderização da importação no modo Super Sapiens — Distribuição."""
    col_up, col_info = st.columns([2, 1])
    with col_up:
        uploaded = st.file_uploader(
            "Selecione o arquivo Excel (.xlsx):",
            type=["xlsx"],
            accept_multiple_files=False,
            help=(
                "Relatório 'Tarefas Judiciais Distribuídas ou Redistribuídas por um Usuário "
                "em um Período de Tempo (Detalhado)' exportado do Super Sapiens."
            ),
            key="uploader_distribuicao",
        )
    with col_info:
        st.markdown(
            "<div class='ac-card'>"
            "<div class='ac-card-header-light'>"
            "<span style='font-size:0.8rem;font-weight:700'>Formato esperado</span>"
            "</div>"
            "<div class='ac-card-body' style='font-size:0.82rem'>"
            "<div class='ac-label' style='margin-bottom:4px'>Relatório</div>"
            "<div class='ac-value' style='font-size:0.76rem'>Tarefas Distribuídas ou<br>"
            "Redistribuídas (Detalhado)</div>"
            "<div class='ac-label' style='margin-bottom:4px'>Colunas auditadas</div>"
            "<div class='ac-value' style='font-size:0.78rem'>Id · NUP · Fonte dos Dados<br>"
            "Setor Origem → Setor Destino</div>"
            "<div class='ac-label' style='margin-bottom:4px'>Foco da auditoria</div>"
            "<div class='ac-value' style='font-size:0.78rem;color:#1A3A6A;font-weight:700'>"
            "Conformidade do Setor de Destino</div>"
            "</div></div>",
            unsafe_allow_html=True,
        )

    if not uploaded:
        if get_dist_data() is not None:
            dd = get_dist_data()
            st.info(
                f"Arquivo **{dd.nome_arquivo}** já carregado. "
                "Navegue pelas etapas no menu lateral."
            )
        else:
            st.info("Selecione o arquivo Excel de Distribuição para começar.")
        return

    try:
        dd = load_distribution_file(uploaded, uploaded.name)
    except ValueError as e:
        st.error(str(e))
        return

    atual_dd = st.session_state.get("dist_data")
    if atual_dd is not None and getattr(atual_dd, "nome_arquivo", None) != dd.nome_arquivo:
        reset_auditoria()

    st.session_state["dist_data"] = dd
    st.session_state["tipo_relatorio"] = "supp_distribuicao"
    save_session()

    st.divider()
    st.subheader("Resumo do Relatório de Distribuição")

    if dd.periodo_inicio and dd.periodo_fim:
        periodo_str = (
            f"{dd.periodo_inicio.strftime('%d/%m/%Y %H:%M')} "
            f"até {dd.periodo_fim.strftime('%d/%m/%Y %H:%M')}"
        )
    else:
        periodo_str = "Período não identificado"

    extra_html = ""
    if dd.usuario_distribuidor:
        extra_html = (
            f"<div style='margin-top:0.5rem'>"
            f"<span style='font-size:0.67rem;font-weight:700;letter-spacing:0.08em;"
            f"text-transform:uppercase;color:#a8bcd4'>Distribuidor</span><br>"
            f"<span style='font-size:0.88rem;font-weight:600'>{dd.usuario_distribuidor}</span>"
            f"</div>"
        )

    st.markdown(
        f"<div class='ac-card'>"
        f"<div class='ac-card-header' style='display:flex;align-items:center;gap:0.6rem'>"
        f"<span style='font-size:1rem'>📅</span>"
        f"<div><div class='ac-label ac-label-dark'>Período de distribuição</div>"
        f"<div style='font-weight:600;font-size:0.92rem'>{periodo_str}</div>"
        f"{extra_html}</div>"
        f"</div></div>",
        unsafe_allow_html=True,
    )

    st.metric("Total de Distribuições", dd.total_distribuicoes)

    # Prévia da distribuição por setor destino
    st.divider()
    st.subheader("Prévia — Distribuição por Setor de Destino")
    if COL_DIST_SETOR_DESTINO in dd.df.columns:
        top_setores = (
            dd.df[COL_DIST_SETOR_DESTINO]
            .value_counts()
            .head(10)
            .reset_index()
            .rename(columns={"index": "Setor Destino", COL_DIST_SETOR_DESTINO: "Qtd"})
        )
        # pandas 2.x renames columns differently
        top_setores.columns = ["Setor Destino", "Qtd"]
        st.dataframe(top_setores, hide_index=True, use_container_width=True)

    st.divider()
    if st.button("Iniciar Auditoria →", type="primary"):
        st.session_state["pagina"] = "distribuicao"
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
            save_session()
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

    descr = f"Amostra: {n_amostra} tarefas" if n_amostra else f"Total: {len(df)} tarefas"
    st.markdown(
        f"<span class='ac-badge'>{tipo_label}</span>"
        f"<span class='ac-badge ac-badge-light'>{descr}</span>",
        unsafe_allow_html=True,
    )

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
        else:
            _render_processo_link()

    st.divider()
    col1, col2 = st.columns([2, 1])
    with col1:
        if st.button("Concluir e Avançar para Tarefas Não Triadas →", type="primary"):
            st.session_state["auditoria_triadas_concluida"] = True
            st.session_state["pagina"] = "nao_triadas"
            save_session()
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
            save_session()
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
        else:
            _render_processo_link()

    st.divider()
    col1, col2 = st.columns([2, 1])
    with col1:
        if st.button("Concluir e Ir para Relatório →", type="primary"):
            st.session_state["auditoria_nao_triadas_concluida"] = True
            st.session_state["pagina"] = "relatorio"
            save_session()
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
# PÁGINA — AUDITORIA DE DISTRIBUIÇÃO (Super Sapiens)
# ===========================================================================

def render_auditoria_distribuicao() -> None:
    dist_data = get_dist_data()
    if dist_data is None:
        st.warning("Nenhum arquivo carregado. Volte à página de Importação.")
        return

    st.title("📊 Auditoria de Distribuição de Tarefas")
    st.caption(
        "Verifique a conformidade do **Setor de Destino** nas distribuições realizadas no Super Sapiens."
    )

    # ── Seleção do tipo de controle ──────────────────────────────────────────
    if st.session_state.get("tipo_controle") is None:
        total = dist_data.total_distribuicoes
        st.markdown(
            f"**{total}** distribuições disponíveis no relatório. "
            "Selecione o tipo de controle conforme o Manual de Gerenciamento Estratégico "
            "(Portaria PGF/AGU n. 541/2025, seção 5)."
        )
        st.divider()

        col_esq, col_dir = st.columns([1, 1])
        with col_esq:
            st.markdown("#### Tipo de Controle")
            tipo = st.radio(
                "Selecione:",
                ["Controle Simplificado", "Controle Detalhado (Amostragem Estatística)"],
                key="radio_tipo_dist",
                label_visibility="collapsed",
            )
            st.markdown("""
            **Controle Simplificado** — Verificação de todas as distribuições (ou seleção manual).

            **Controle Detalhado** — Amostragem estatística com seleção aleatória.
            Nível de confiança **95%**, margem de erro **±5%**.
            """)
        with col_dir:
            st.markdown("#### Tamanho da Amostra (Controle Detalhado)")
            if total > 0:
                n = calcular_amostra(total)
                st.metric("Distribuições a auditar", n, delta=f"{n / total * 100:.1f}% do universo")
                st.markdown(formula_descricao(total))
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
                n = calcular_amostra(total)
                st.session_state["tamanho_amostra"] = n
                df_base = selecionar_amostra(dist_data.df, n)
            else:
                st.session_state["tamanho_amostra"] = None
                df_base = dist_data.df.copy()

            colunas = [
                COL_DIST_ID, COL_DIST_NUP, COL_DIST_PROCESSO_JUDICIAL,
                COL_DIST_FONTE_DADOS, COL_DIST_SETOR_ORIGEM,
                COL_DIST_USUARIO_DESTINO, COL_DIST_SETOR_DESTINO, COL_DIST_DATA_HORA,
            ]
            st.session_state["df_audit_distribuicao"] = preparar_df_auditoria(df_base, colunas)
            save_session()
            st.rerun()
        return

    # ── Editor ──────────────────────────────────────────────────────────────
    tipo_controle = st.session_state["tipo_controle"]
    tipo_label = "Controle Simplificado" if tipo_controle == "simplificado" else "Controle Detalhado"
    n_amostra = st.session_state.get("tamanho_amostra")
    df = st.session_state.get("df_audit_distribuicao")

    if df is None:
        st.error("Estado inconsistente. Clique em 'Nova Auditoria' no menu lateral.")
        return

    descr = f"Amostra: {n_amostra} distribuições" if n_amostra else f"Total: {len(df)} distribuições"
    st.markdown(
        f"<span class='ac-badge'>{tipo_label}</span>"
        f"<span class='ac-badge ac-badge-light'>{descr}</span>",
        unsafe_allow_html=True,
    )

    col_left, col_right = st.columns([3, 2], gap="medium")
    with col_left:
        orig_idx, row = _render_audit_table(
            df_key="df_audit_distribuicao",
            filtro_key="filtro_conf_dist",
            busca_key="busca_dist",
            column_order=[
                COL_DIST_ID, COL_DIST_NUP, COL_DIST_FONTE_DADOS,
                COL_DIST_SETOR_ORIGEM, COL_DIST_SETOR_DESTINO,
                COL_DIST_USUARIO_DESTINO, COL_DIST_DATA_HORA,
                COL_CONFORMIDADE, COL_MOTIVO, COL_ACAO,
            ],
            table_key="tbl_distribuicao",
            id_col=COL_DIST_ID,
            nup_col=COL_DIST_NUP,
        )
    with col_right:
        if orig_idx is not None and row is not None:
            _render_dist_row_editor("df_audit_distribuicao", orig_idx, row)
        else:
            st.markdown(
                "<div class='ac-card'>"
                "<div class='ac-card-body' style='padding:1.2rem 0.9rem;text-align:center;"
                "color:#7a8fad;font-size:0.85rem'>"
                "← Selecione uma linha na tabela para auditar a conformidade do setor de destino."
                "</div></div>",
                unsafe_allow_html=True,
            )

    st.divider()
    col1, col2 = st.columns([2, 1])
    with col1:
        if st.button("Concluir e Ir para Relatório →", type="primary"):
            st.session_state["auditoria_distribuicao_concluida"] = True
            st.session_state["pagina"] = "relatorio"
            save_session()
            st.rerun()
    with col2:
        if st.button("↩ Trocar Tipo de Controle"):
            st.session_state["tipo_controle"] = None
            st.session_state["df_audit_distribuicao"] = None
            st.session_state["tamanho_amostra"] = None
            st.session_state["auditoria_distribuicao_concluida"] = False
            for k in list(st.session_state.keys()):
                if k.startswith(("filtro_conf_dist", "busca_dist", "tbl_distribuicao",
                                 "edit_conf_", "edit_motivo_", "edit_acao_", "btn_save_row_")):
                    del st.session_state[k]
            st.rerun()


def _render_dist_row_editor(df_key: str, orig_idx, row: dict) -> None:
    """Painel de edição para linhas do relatório de distribuição."""
    tarefa_id = row.get(COL_DIST_ID)
    nup = row.get(COL_DIST_NUP)
    proc_judicial = row.get(COL_DIST_PROCESSO_JUDICIAL)
    fonte = row.get(COL_DIST_FONTE_DADOS)
    setor_origem = row.get(COL_DIST_SETOR_ORIGEM)
    setor_destino = row.get(COL_DIST_SETOR_DESTINO)
    usuario_destino = row.get(COL_DIST_USUARIO_DESTINO)
    data_hora = row.get(COL_DIST_DATA_HORA)

    # ── Busca dados do processo SUPP (com cache) ────────────────────────────
    cache_key = f"_proc_id_cache_{tarefa_id}"
    if cache_key not in st.session_state:
        auth = st.session_state.get("supp_auth_client")
        if auth and tarefa_id:
            with st.spinner("Buscando processo..."):
                try:
                    from modules.tarefa import TarefaClient
                    tc = TarefaClient.from_auth(auth)
                    tarefa = tc.buscar(tarefa_id, populate=[
                        "processo", "especieTarefa", "usuarioResponsavel",
                        "setorResponsavel", "setorOrigem",
                    ])
                    proc = tarefa.get("processo") or {}
                    proc_id = proc.get("id")
                    nup_fmt = proc.get("NUPFormatado") or proc.get("NUP")
                    any_ = proc.get("any") or {}
                    pj = any_.get("processoJudicial") or {}
                    cnj = pj.get("numeroFormatado") or pj.get("numero")
                    cn = pj.get("classeNacional") or {}
                    classe_nacional = cn.get("nome") if isinstance(cn, dict) else None
                    st.session_state[cache_key] = {
                        "proc_id": proc_id,
                        "nup_fmt": nup_fmt,
                        "cnj": cnj,
                        "classe_nacional": classe_nacional,
                    }
                except Exception as e:
                    st.session_state[cache_key] = {"erro": str(e)}
        else:
            st.session_state[cache_key] = {}

    cached = st.session_state.get(cache_key, {})
    proc_id = cached.get("proc_id")
    nup_fmt = cached.get("nup_fmt") or nup
    cnj = cached.get("cnj") or proc_judicial

    # ── Cabeçalho ─────────────────────────────────────────────────────────────
    url = _SUPERSAPIENS_URL.format(proc_id=proc_id) if proc_id else None
    _c_title, _c_open, _c_refresh = st.columns([4, 3, 1])
    with _c_title:
        st.markdown("#### ✏️ Auditoria")
    with _c_open:
        if url:
            st.link_button("↗ SuperSapiens", url, use_container_width=True)
    with _c_refresh:
        if tarefa_id and st.button("🔄", key=f"_refresh_dist_{tarefa_id}", help="Recarregar dados"):
            st.session_state.pop(cache_key, None)
            st.rerun()

    def _field(label: str, value, mono: bool = False, highlight: bool = False) -> str:
        if not value:
            return ""
        val_style = "font-family:monospace;font-size:0.82rem" if mono else "font-size:0.85rem"
        if highlight:
            val_style += ";color:#1A3A6A;font-weight:800"
        else:
            val_style += ";color:#1a2a4a"
        return (
            f"<div style='margin-bottom:0.55rem'>"
            f"<div style='font-size:0.67rem;font-weight:700;letter-spacing:0.08em;"
            f"text-transform:uppercase;color:#7a8fad;margin-bottom:1px'>{label}</div>"
            f"<div style='font-weight:600;{val_style}'>{value}</div>"
            f"</div>"
        )

    # ── Card de identificação ─────────────────────────────────────────────────
    id_html = (
        f"<div style='background:#1A3A6A;color:#fff;padding:0.55rem 0.9rem;"
        f"display:flex;gap:1.5rem;border-radius:6px 6px 0 0'>"
        f"<div style='flex:1'>"
        f"<div style='font-size:0.67rem;font-weight:700;letter-spacing:0.08em;"
        f"text-transform:uppercase;color:#a8bcd4;margin-bottom:1px'>Id Tarefa</div>"
        f"<div style='font-weight:700;font-size:0.92rem;font-family:monospace'>{tarefa_id or '—'}</div>"
        f"</div>"
        f"<div style='flex:1'>"
        f"<div style='font-size:0.67rem;font-weight:700;letter-spacing:0.08em;"
        f"text-transform:uppercase;color:#a8bcd4;margin-bottom:1px'>Id Processo</div>"
        f"<div style='font-weight:700;font-size:0.92rem;font-family:monospace'>{proc_id or '—'}</div>"
        f"</div>"
        f"</div>"
    )
    data_str = str(data_hora) if data_hora else ""
    if hasattr(data_hora, "strftime"):
        data_str = data_hora.strftime("%d/%m/%Y %H:%M")

    body_html = (
        _field("NUP", nup_fmt, mono=True)
        + _field("Número CNJ", cnj, mono=True)
        + _field("Fonte dos Dados", fonte)
        + _field("Setor Origem", setor_origem)
        + _field("Setor Destino", setor_destino, highlight=True)
        + _field("Usuário Destino", usuario_destino)
        + _field("Data/Hora", data_str)
    )
    st.markdown(
        f"<div style='border:1px solid #d0dcea;border-radius:6px;"
        f"overflow:hidden;margin-bottom:0.7rem'>"
        f"{id_html}"
        f"<div style='background:#f7f9fc;padding:0.65rem 0.9rem 0.25rem'>{body_html}</div>"
        f"</div>",
        unsafe_allow_html=True,
    )
    if cached.get("erro"):
        st.caption(f"⚠️ Erro ao buscar processo: {cached['erro']}")

    # ── Campos de auditoria ────────────────────────────────────────────────────
    cur_conf = row.get(COL_CONFORMIDADE, OPCOES_CONFORMIDADE[0])
    if cur_conf not in OPCOES_CONFORMIDADE:
        cur_conf = OPCOES_CONFORMIDADE[0]

    conf = st.selectbox(
        "Conformidade do Setor de Destino:",
        OPCOES_CONFORMIDADE,
        index=OPCOES_CONFORMIDADE.index(cur_conf),
        key=f"edit_conf_{orig_idx}_{df_key}",
    )
    motivo = st.text_area(
        "Motivo NC:",
        value=row.get(COL_MOTIVO, "") or "",
        key=f"edit_motivo_{orig_idx}_{df_key}",
        height=90,
        placeholder="Ex.: setor destino incompatível com a origem (TJMA → Núcleo Educação)…",
    )
    acao = st.text_area(
        "Ação Corretiva:",
        value=row.get(COL_ACAO, "") or "",
        key=f"edit_acao_{orig_idx}_{df_key}",
        height=90,
        placeholder="Ex.: redistribuir para o Núcleo correto e ajustar configuração de triagem…",
    )

    if st.button("💾 Salvar", type="primary", key=f"btn_save_row_{orig_idx}_{df_key}",
                 use_container_width=True):
        df = st.session_state[df_key].copy()
        df.at[orig_idx, COL_CONFORMIDADE] = conf
        df.at[orig_idx, COL_MOTIVO] = motivo
        df.at[orig_idx, COL_ACAO] = acao
        st.session_state[df_key] = df
        save_session()
        st.rerun()


# ===========================================================================
# PÁGINA 4 — RELATÓRIO
# ===========================================================================

def _render_relatorio_distribuicao() -> None:
    """Relatório de auditoria para o modo Super Sapiens — Distribuição."""
    from modules.report import gerar_relatorio_distribuicao

    dist_data = get_dist_data()
    if dist_data is None:
        st.warning("Nenhum arquivo carregado. Volte à página de Importação.")
        return

    st.title("📄 Relatório de Auditoria — Distribuição")

    df_dist = get_df_distribuicao()
    s_dist = stats_df(df_dist)

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

    st.divider()
    st.subheader("Resumo Executivo")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total de Distribuições", dist_data.total_distribuicoes)
    c2.metric("Na Amostra / Auditadas", f"{s_dist['auditadas']}/{s_dist['total']}")
    c3.metric("Conformes", s_dist["conformes"],
              delta=f"{s_dist['pct_conf']:.1f}%" if s_dist["auditadas"] > 0 else None)
    c4.metric("Não Conformes", s_dist["nao_conformes"],
              delta=f"{s_dist['pct_nc']:.1f}%" if s_dist["auditadas"] > 0 else None,
              delta_color="inverse")

    if s_dist["nao_conformes"] > 0:
        st.divider()
        st.subheader(f"⚠️ Não Conformidades Identificadas ({s_dist['nao_conformes']})")
        if df_dist is not None:
            nc_df = df_dist[df_dist[COL_CONFORMIDADE] == "Não Conforme"][
                [c for c in [COL_DIST_ID, COL_DIST_NUP, COL_DIST_SETOR_DESTINO,
                             COL_MOTIVO, COL_ACAO] if c in df_dist.columns]
            ].copy()
            st.dataframe(nc_df, hide_index=True, use_container_width=True)
    else:
        st.success("Nenhuma não conformidade identificada nas distribuições auditadas.")

    st.divider()
    col_btn1, col_btn2 = st.columns([1, 2])
    with col_btn1:
        if st.button("📥 Gerar Relatório (.docx)", type="primary", use_container_width=True):
            with st.spinner("Gerando relatório…"):
                try:
                    docx_bytes = gerar_relatorio_distribuicao(
                        dist_data=dist_data,
                        df_distribuicao=df_dist,
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
        nome = f"relatorio_distribuicao_{data_aud.strftime('%Y-%m-%d')}.docx"
        with col_btn2:
            st.download_button(
                label="⬇️ Baixar Relatório Word (.docx)",
                data=st.session_state["relatorio_gerado"],
                file_name=nome,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                type="primary",
            )


def render_relatorio() -> None:
    tipo_rel = st.session_state.get("tipo_relatorio")
    if tipo_rel == "supp_distribuicao":
        _render_relatorio_distribuicao()
        return

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
                cores_v.append("#16a34a")
            if s["nao_conformes"] > 0:
                labels_v.append(f"Não Conformes\n{s['nao_conformes']}")
                sizes_v.append(s["nao_conformes"])
                cores_v.append("#f59e0b")
            nao_aud = s["total"] - s["auditadas"]
            if nao_aud > 0:
                labels_v.append(f"Não auditadas\n{nao_aud}")
                sizes_v.append(nao_aud)
                cores_v.append("#c8d4e8")
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
elif pagina == "distribuicao":
    render_auditoria_distribuicao()
elif pagina == "relatorio":
    render_relatorio()
