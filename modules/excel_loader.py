"""
Carregamento e parsing de planilhas Excel geradas pelo Conecta+ Automação.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from typing import IO

import pandas as pd


# ---------------------------------------------------------------------------
# Constantes de colunas
# ---------------------------------------------------------------------------

COL_ID = "ID"
COL_TAREFA = "Tarefa"
COL_NUP = "NUP"
COL_USUARIO = "Usuário"
COL_DATA_INCLUSAO = "Data Inclusão Fila"
COL_DATA_INICIO = "Data Início"
COL_DATA_FIM = "Data Fim"
COL_STATUS = "Status"
COL_CONFIG = "Configurações Encontradas"

DATE_COLS = [COL_DATA_INCLUSAO, COL_DATA_INICIO, COL_DATA_FIM]
DATE_FORMAT = "%d/%m/%Y, %H:%M:%S"

REQUIRED_SHEETS = ["Todas as Tarefas", "Tarefas Triadas", "Tarefas Não Triadas"]
REQUIRED_COLS = [COL_ID, COL_TAREFA, COL_NUP, COL_USUARIO,
                 COL_DATA_INCLUSAO, COL_DATA_INICIO, COL_DATA_FIM,
                 COL_STATUS, COL_CONFIG]


# ---------------------------------------------------------------------------
# Dataclass principal
# ---------------------------------------------------------------------------

@dataclass
class AuditData:
    nome_arquivo: str
    periodo_inicio: datetime | None
    periodo_fim: datetime | None
    todas: pd.DataFrame
    triadas: pd.DataFrame
    nao_triadas: pd.DataFrame
    total_tarefas: int
    total_triadas: int
    total_nao_triadas: int
    pct_triadas: float
    pct_nao_triadas: float


# ---------------------------------------------------------------------------
# Funções internas
# ---------------------------------------------------------------------------

def _parse_dates(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for col in DATE_COLS:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], format=DATE_FORMAT, errors="coerce")
    return df


def _detect_period(df: pd.DataFrame) -> tuple[datetime | None, datetime | None]:
    datas = []
    for col in DATE_COLS:
        if col in df.columns:
            valid = df[col].dropna()
            if not valid.empty:
                datas.extend([valid.min(), valid.max()])
    if not datas:
        return None, None
    return min(datas), max(datas)


def _validate(sheets: dict[str, pd.DataFrame], nome: str) -> None:
    for sheet in REQUIRED_SHEETS:
        if sheet not in sheets:
            raise ValueError(
                f"Arquivo '{nome}' não contém a aba '{sheet}'. "
                f"Verifique se é um arquivo exportado pelo Conecta+ Automação."
            )
    for sheet in REQUIRED_SHEETS:
        df = sheets[sheet]
        missing = [c for c in REQUIRED_COLS if c not in df.columns]
        if missing:
            raise ValueError(
                f"Aba '{sheet}' em '{nome}' não possui as colunas: {missing}. "
                f"Verifique o formato do arquivo."
            )


# ---------------------------------------------------------------------------
# API pública
# ---------------------------------------------------------------------------

def load_file(uploaded_file: IO, nome_arquivo: str = "") -> AuditData:
    """
    Lê um arquivo xlsx e retorna AuditData.
    Raises ValueError com mensagem em português se o arquivo for inválido.
    """
    nome = nome_arquivo or getattr(uploaded_file, "name", "arquivo.xlsx")
    try:
        sheets: dict[str, pd.DataFrame] = pd.read_excel(
            uploaded_file, sheet_name=None, engine="openpyxl", dtype=str
        )
    except Exception as e:
        raise ValueError(f"Não foi possível ler o arquivo '{nome}': {e}") from e

    _validate(sheets, nome)

    todas = _parse_dates(sheets["Todas as Tarefas"])
    triadas = _parse_dates(sheets["Tarefas Triadas"])
    nao_triadas = _parse_dates(sheets["Tarefas Não Triadas"])

    periodo_inicio, periodo_fim = _detect_period(todas)

    total = len(todas)
    n_tri = len(triadas)
    n_nao = len(nao_triadas)

    return AuditData(
        nome_arquivo=nome,
        periodo_inicio=periodo_inicio,
        periodo_fim=periodo_fim,
        todas=todas,
        triadas=triadas,
        nao_triadas=nao_triadas,
        total_tarefas=total,
        total_triadas=n_tri,
        total_nao_triadas=n_nao,
        pct_triadas=(n_tri / total * 100) if total > 0 else 0.0,
        pct_nao_triadas=(n_nao / total * 100) if total > 0 else 0.0,
    )


def merge_audit_data(files: list[AuditData]) -> AuditData:
    """
    Consolida múltiplos AuditData em um único, deduplicando por COL_TAREFA.
    """
    if len(files) == 1:
        return files[0]

    todas = pd.concat([f.todas for f in files], ignore_index=True)
    triadas = pd.concat([f.triadas for f in files], ignore_index=True)
    nao_triadas = pd.concat([f.nao_triadas for f in files], ignore_index=True)

    # Deduplica por ID de tarefa; mantém última ocorrência (re-execução mais recente)
    todas = todas.drop_duplicates(subset=[COL_TAREFA], keep="last").reset_index(drop=True)
    triadas = triadas.drop_duplicates(subset=[COL_TAREFA], keep="last").reset_index(drop=True)
    nao_triadas = nao_triadas.drop_duplicates(subset=[COL_TAREFA], keep="last").reset_index(drop=True)

    datas_inicio = [f.periodo_inicio for f in files if f.periodo_inicio]
    datas_fim = [f.periodo_fim for f in files if f.periodo_fim]

    total = len(todas)
    n_tri = len(triadas)
    n_nao = len(nao_triadas)

    return AuditData(
        nome_arquivo="Consolidado",
        periodo_inicio=min(datas_inicio) if datas_inicio else None,
        periodo_fim=max(datas_fim) if datas_fim else None,
        todas=todas,
        triadas=triadas,
        nao_triadas=nao_triadas,
        total_tarefas=total,
        total_triadas=n_tri,
        total_nao_triadas=n_nao,
        pct_triadas=(n_tri / total * 100) if total > 0 else 0.0,
        pct_nao_triadas=(n_nao / total * 100) if total > 0 else 0.0,
    )
