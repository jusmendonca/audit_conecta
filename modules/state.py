"""
Gerenciamento de session_state do Streamlit.
O estado central usa DataFrames pandas com colunas de auditoria adicionadas inline.
Inclui persistência em disco para retomar auditorias entre sessões.
"""
from __future__ import annotations

import pickle
from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st

# Colunas de auditoria adicionadas aos DataFrames
COL_CONFORMIDADE = "Conformidade"
COL_MOTIVO = "Motivo da Não Conformidade"
COL_ACAO = "Ação Corretiva"

OPCOES_CONFORMIDADE = ["Não auditada", "Conforme", "Não Conforme"]

# Arquivo de sessão persistida (ao lado do app.py)
_SESSION_FILE = Path(__file__).parent.parent / "audit_session.pkl"

# Chaves que são salvas/restauradas entre sessões
_PERSIST_KEYS = [
    "tipo_relatorio",
    "pagina",
    "audit_data_merged",
    "dist_data",
    "tipo_controle",
    "tamanho_amostra",
    "df_audit_triadas",
    "auditoria_triadas_concluida",
    "df_audit_nao_triadas",
    "auditoria_nao_triadas_concluida",
    "df_audit_distribuicao",
    "auditoria_distribuicao_concluida",
    "responsavel",
    "data_auditoria",
]

# ---------------------------------------------------------------------------
# Chaves e defaults do session_state
# ---------------------------------------------------------------------------

_DEFAULTS: dict = {
    "pagina": "importacao",
    "tipo_relatorio": None,          # "conecta_triagem" | "supp_distribuicao"
    "audit_data_merged": None,
    # auditoria de triadas
    "tipo_controle": None,           # "simplificado" | "detalhado"
    "tamanho_amostra": None,         # int
    "df_audit_triadas": None,        # pd.DataFrame com colunas de auditoria
    "auditoria_triadas_concluida": False,
    # auditoria de não-triadas
    "df_audit_nao_triadas": None,    # pd.DataFrame com colunas de auditoria
    "auditoria_nao_triadas_concluida": False,
    # auditoria de distribuição
    "dist_data": None,               # DistribuicaoData
    "df_audit_distribuicao": None,   # pd.DataFrame com colunas de auditoria
    "auditoria_distribuicao_concluida": False,
    # relatório
    "responsavel": "",
    "data_auditoria": date.today(),
    "relatorio_gerado": None,
    # controle de restauração de sessão
    "session_restore_offered": False,
}


def init_state() -> None:
    """Inicializa todas as chaves com valores padrão (idempotente)."""
    for key, default in _DEFAULTS.items():
        if key not in st.session_state:
            st.session_state[key] = default


def reset_auditoria() -> None:
    """Limpa resultados de auditoria sem remover dados do arquivo carregado."""
    resetar = [
        "tipo_relatorio",
        "tipo_controle", "tamanho_amostra",
        "df_audit_triadas", "auditoria_triadas_concluida",
        "df_audit_nao_triadas", "auditoria_nao_triadas_concluida",
        "df_audit_distribuicao", "auditoria_distribuicao_concluida",
        "relatorio_gerado",
    ]
    for key in resetar:
        st.session_state[key] = _DEFAULTS[key]


# ---------------------------------------------------------------------------
# Persistência de sessão
# ---------------------------------------------------------------------------

def save_session() -> None:
    """Persiste o estado de auditoria em disco."""
    data = {k: st.session_state.get(k) for k in _PERSIST_KEYS}
    try:
        with open(_SESSION_FILE, "wb") as f:
            pickle.dump(data, f, protocol=pickle.HIGHEST_PROTOCOL)
    except Exception:
        pass  # falha silenciosa — persistência é best-effort


def load_session() -> bool:
    """
    Carrega estado salvo do disco para session_state.
    Retorna True se carregado com sucesso, False caso contrário.
    """
    if not _SESSION_FILE.exists():
        return False
    try:
        with open(_SESSION_FILE, "rb") as f:
            data = pickle.load(f)
        for k, v in data.items():
            st.session_state[k] = v
        return True
    except Exception:
        return False


def has_saved_session() -> bool:
    """Verifica se existe sessão salva em disco."""
    return _SESSION_FILE.exists()


def clear_saved_session() -> None:
    """Remove o arquivo de sessão salva."""
    try:
        _SESSION_FILE.unlink(missing_ok=True)
    except Exception:
        pass


def get_session_info() -> dict | None:
    """
    Lê metadados básicos da sessão salva (sem carregar tudo).
    Retorna dict com 'nome_arquivo' e 'pagina', ou None.
    """
    if not _SESSION_FILE.exists():
        return None
    try:
        with open(_SESSION_FILE, "rb") as f:
            data = pickle.load(f)
        ad = data.get("audit_data_merged")
        dd = data.get("dist_data")
        nome = None
        if ad is not None:
            nome = getattr(ad, "nome_arquivo", None)
        elif dd is not None:
            nome = getattr(dd, "nome_arquivo", None)
        return {
            "nome_arquivo": nome,
            "pagina": data.get("pagina", "importacao"),
            "tipo_relatorio": data.get("tipo_relatorio"),
        }
    except Exception:
        return None


# ---------------------------------------------------------------------------
# Helpers de preparação de DataFrames
# ---------------------------------------------------------------------------

def preparar_df_auditoria(df: pd.DataFrame, colunas_mostrar: list[str]) -> pd.DataFrame:
    """
    Cria uma cópia do DataFrame com apenas as colunas relevantes + colunas de auditoria.
    Preserva edições anteriores se o DataFrame já existir no estado.
    """
    colunas = [c for c in colunas_mostrar if c in df.columns]
    resultado = df[colunas].copy().reset_index(drop=True)
    resultado[COL_CONFORMIDADE] = OPCOES_CONFORMIDADE[0]   # "Não auditada"
    resultado[COL_MOTIVO] = ""
    resultado[COL_ACAO] = ""
    return resultado


# ---------------------------------------------------------------------------
# Helpers de leitura
# ---------------------------------------------------------------------------

def get_audit_data():
    return st.session_state.get("audit_data_merged")


def get_dist_data():
    return st.session_state.get("dist_data")


def get_df_triadas() -> pd.DataFrame | None:
    return st.session_state.get("df_audit_triadas")


def get_df_nao_triadas() -> pd.DataFrame | None:
    return st.session_state.get("df_audit_nao_triadas")


def get_df_distribuicao() -> pd.DataFrame | None:
    return st.session_state.get("df_audit_distribuicao")


def stats_df(df: pd.DataFrame | None) -> dict:
    """
    Retorna dicionário com estatísticas de conformidade de um DataFrame de auditoria.
    Considera apenas linhas onde Conformidade != 'Não auditada'.
    """
    if df is None or df.empty:
        return {"total": 0, "auditadas": 0, "conformes": 0, "nao_conformes": 0,
                "pct_conf": 0.0, "pct_nc": 0.0}
    auditadas = df[df[COL_CONFORMIDADE] != "Não auditada"]
    n_aud = len(auditadas)
    n_conf = (auditadas[COL_CONFORMIDADE] == "Conforme").sum()
    n_nc = (auditadas[COL_CONFORMIDADE] == "Não Conforme").sum()
    return {
        "total": len(df),
        "auditadas": n_aud,
        "conformes": int(n_conf),
        "nao_conformes": int(n_nc),
        "pct_conf": (n_conf / n_aud * 100) if n_aud > 0 else 0.0,
        "pct_nc": (n_nc / n_aud * 100) if n_aud > 0 else 0.0,
    }
