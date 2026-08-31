"""
Gerenciamento de session_state do Streamlit.
O estado central usa DataFrames pandas com colunas de auditoria adicionadas inline.
Inclui persistência em disco para retomar auditorias entre sessões.
"""
from __future__ import annotations

import dataclasses
import os
import pickle
import sys
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
    "det_data",
    "tipo_controle",
    "tamanho_amostra",
    "df_audit_triadas",
    "auditoria_triadas_concluida",
    "df_audit_nao_triadas",
    "auditoria_nao_triadas_concluida",
    "df_audit_distribuicao",
    "auditoria_distribuicao_concluida",
    "df_audit_detalhamento",
    "auditoria_detalhamento_concluida",
    "responsavel",
    "data_auditoria",
]

# ---------------------------------------------------------------------------
# Chaves e defaults do session_state
# ---------------------------------------------------------------------------

_DEFAULTS: dict = {
    "pagina": "importacao",
    "tipo_relatorio": None,          # "conecta_triagem" | "supp_distribuicao" | "detalhamento_individual"
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
    # detalhamento individual
    "det_data": None,
    "df_audit_detalhamento": None,
    "auditoria_detalhamento_concluida": False,
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
        "df_audit_detalhamento", "auditoria_detalhamento_concluida",
        "relatorio_gerado",
    ]
    for key in resetar:
        st.session_state[key] = _DEFAULTS[key]


# ---------------------------------------------------------------------------
# Persistência de sessão
# ---------------------------------------------------------------------------

def _reancorar_dataclass(obj):
    """
    Corrige uma armadilha do auto-reload de desenvolvimento do Streamlit: ao
    editar um módulo local (ex.: excel_loader.py) com a sessão já aberta, o
    Streamlit recarrega o módulo e cria um *novo* objeto de classe com o
    mesmo nome. Uma instância construída antes do reload (ex.: `det_data`
    guardado em session_state) passa a apontar para a classe *antiga* — os
    dados continuam corretos, só a identidade do tipo ficou obsoleta — e o
    pickle recusa com "it's not the same object as <classe>".

    Reconstrói `obj` usando a classe atualmente carregada, com os mesmos
    valores de campo. Se `obj` não for dataclass, ou a classe já for a
    mesma, ou a reconstrução falhar por qualquer motivo, devolve `obj` sem
    alteração — nesse caso o pickle segue e falha (ou não) por conta própria.
    """
    if not dataclasses.is_dataclass(obj) or isinstance(obj, type):
        return obj
    modulo = sys.modules.get(type(obj).__module__)
    cls_atual = getattr(modulo, type(obj).__name__, None) if modulo else None
    if cls_atual is None or cls_atual is type(obj):
        return obj
    try:
        return cls_atual(**{f.name: getattr(obj, f.name) for f in dataclasses.fields(obj)})
    except Exception:
        return obj


def save_session() -> None:
    """
    Persiste o estado de auditoria em disco, de forma atômica: grava num
    arquivo temporário e só substitui o arquivo real (`os.replace`, atômico
    no mesmo volume) se a gravação terminar por completo. Se o processo for
    interrompido no meio do caminho, o arquivo já salvo permanece intacto —
    o que se perde é, no máximo, a tentativa em andamento, nunca o que já
    estava gravado.

    Falhas ficam registradas em `st.session_state["_save_session_erro"]`
    para exibição na UI: perder a auditoria de forma silenciosa é pior do
    que incomodar o auditor com um aviso.
    """
    data = {k: _reancorar_dataclass(st.session_state.get(k)) for k in _PERSIST_KEYS}
    tmp_path = _SESSION_FILE.with_name(_SESSION_FILE.name + ".tmp")
    try:
        with open(tmp_path, "wb") as f:
            pickle.dump(data, f, protocol=pickle.HIGHEST_PROTOCOL)
        os.replace(tmp_path, _SESSION_FILE)
    except Exception as exc:
        st.session_state["_save_session_erro"] = str(exc)
        try:
            tmp_path.unlink(missing_ok=True)
        except Exception:
            pass
        return
    st.session_state.pop("_save_session_erro", None)


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
        det = data.get("det_data")
        nome = None
        if ad is not None:
            nome = getattr(ad, "nome_arquivo", None)
        elif dd is not None:
            nome = getattr(dd, "nome_arquivo", None)
        elif det is not None:
            nome = getattr(det, "nome_arquivo", None)
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


def get_det_data():
    return st.session_state.get("det_data")


def get_df_triadas() -> pd.DataFrame | None:
    return st.session_state.get("df_audit_triadas")


def get_df_nao_triadas() -> pd.DataFrame | None:
    return st.session_state.get("df_audit_nao_triadas")


def get_df_distribuicao() -> pd.DataFrame | None:
    return st.session_state.get("df_audit_distribuicao")


def get_df_detalhamento() -> pd.DataFrame | None:
    return st.session_state.get("df_audit_detalhamento")


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
