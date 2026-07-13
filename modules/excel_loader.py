"""
Carregamento e parsing de planilhas Excel.
Suporta dois formatos:
  - Conecta+ Automação (Triagem Avançada)
  - Super Sapiens (Relatório de Distribuição de Tarefas)
"""
from __future__ import annotations

import unicodedata
from dataclasses import dataclass
from datetime import datetime
from typing import IO

import pandas as pd


# ---------------------------------------------------------------------------
# Constantes de colunas — Conecta+ Triagem Avançada
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
# Um mesmo arquivo do Conecta+ mistura formatos na mesma coluna: as células
# gravadas como texto vêm em dd/mm/aaaa (com ou sem vírgula antes da hora,
# conforme a versão do export) e as gravadas como data real chegam em ISO.
# Todos os formatos são aplicados e os resultados, combinados linha a linha.
DATE_FORMATS = [
    "%d/%m/%Y, %H:%M:%S",
    "%d/%m/%Y %H:%M:%S",
    "%Y-%m-%d %H:%M:%S",
    "%d/%m/%Y",
]

REQUIRED_SHEETS = ["Todas as Tarefas", "Tarefas Triadas", "Tarefas Não Triadas"]
REQUIRED_COLS = [COL_ID, COL_TAREFA, COL_NUP, COL_USUARIO,
                 COL_DATA_INCLUSAO, COL_DATA_INICIO, COL_DATA_FIM,
                 COL_STATUS, COL_CONFIG]

# ---------------------------------------------------------------------------
# Constantes de colunas — Super Sapiens Distribuição
# ---------------------------------------------------------------------------

COL_DIST_ID = "Id"
COL_DIST_NUP = "NUP"
COL_DIST_PROCESSO_JUDICIAL = "Processo_judicial"
COL_DIST_FONTE_DADOS = "Fonte_dados"
COL_DIST_USUARIO_ORIGEM = "Usuario_origem"
COL_DIST_SETOR_ORIGEM = "Setor_origem"
COL_DIST_USUARIO_DESTINO = "Usuario_destino"
COL_DIST_SETOR_DESTINO = "Setor_destino"
COL_DIST_DATA_HORA = "DataHoraDistribuicao"

DIST_DATE_FORMAT = "%d/%m/%Y %H:%M:%S"
DIST_REQUIRED_COLS = [COL_DIST_ID, COL_DIST_NUP, COL_DIST_SETOR_DESTINO]


# ---------------------------------------------------------------------------
# Dataclasses
# ---------------------------------------------------------------------------

@dataclass
class DistribuicaoData:
    nome_arquivo: str
    periodo_inicio: datetime | None
    periodo_fim: datetime | None
    df: pd.DataFrame
    total_distribuicoes: int
    usuario_distribuidor: str | None
    params_raw: str | None  # texto bruto dos parâmetros para exibição


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
    # Células de data cuja inversão dia/mês foi desfeita na importação.
    datas_corrigidas: int = 0


# ---------------------------------------------------------------------------
# Funções internas
# ---------------------------------------------------------------------------

def _norm(nome: str) -> str:
    """Chave de comparação tolerante a acento, caixa, NBSP e espaços extras."""
    texto = unicodedata.normalize("NFKD", str(nome)).replace("\xa0", " ")
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return " ".join(texto.split()).casefold()


def _resolver(alvos: list[str], disponiveis: list[str]) -> dict[str, str]:
    """Mapeia nome canônico → nome real encontrado no arquivo (quando existir)."""
    indice = {_norm(d): d for d in disponiveis}
    return {a: indice[_norm(a)] for a in alvos if _norm(a) in indice}


def _to_datetime(serie: pd.Series) -> tuple[pd.Series, int]:
    """
    Converte a coluna e corrige a inversão dia/mês das células ISO.

    O gerador da planilha grava parte das datas como texto dd/mm/aaaa e parte
    como data já convertida (que chega em ISO). Nessa conversão, as datas cujo
    dia é ≤ 12 são lidas no padrão americano mm/dd: '01/07/2026' (1º de julho)
    vira '2026-01-07' (7 de janeiro). As de dia ≥ 13 não são ambíguas e ficam
    intactas em texto — é essa população confiável que usamos como referência.

    A troca só é aplicada quando as datas em texto comprovam a inversão, isto
    é, quando inverter aproxima os meses das células ISO dos meses reais. Se o
    arquivo não tiver datas em texto, nada é alterado.

    Retorna a série convertida e o número de células corrigidas.
    """
    bruto = serie.astype(str).str.strip().replace({"": None, "nan": None, "NaT": None})
    is_iso = bruto.str.match(r"^\d{4}-\d{2}-\d{2}", na=False)

    resultado = pd.Series(pd.NaT, index=serie.index, dtype="datetime64[ns]")
    for fmt in DATE_FORMATS:
        pendentes = resultado.isna() & bruto.notna()
        if not pendentes.any():
            break
        resultado[pendentes] = pd.to_datetime(
            bruto[pendentes], format=fmt, errors="coerce"
        )

    iso = resultado[is_iso & resultado.notna()]
    texto = resultado[~is_iso & resultado.notna()]
    if iso.empty or texto.empty:
        return resultado, 0

    # Só invertemos onde a troca é possível (dia ≤ 12) e apenas se ela aumentar
    # a coerência com os meses observados nas datas confiáveis.
    meses_reais = set(texto.dt.month.unique())
    invertivel = iso[iso.dt.day <= 12]
    if invertivel.empty:
        return resultado, 0

    trocado = pd.to_datetime(
        {
            "year": invertivel.dt.year,
            "month": invertivel.dt.day,
            "day": invertivel.dt.month,
            "hour": invertivel.dt.hour,
            "minute": invertivel.dt.minute,
            "second": invertivel.dt.second,
        }
    )
    coerencia_atual = invertivel.dt.month.isin(meses_reais).mean()
    coerencia_trocada = trocado.dt.month.isin(meses_reais).mean()
    if coerencia_trocada <= coerencia_atual:
        return resultado, 0

    resultado[invertivel.index] = trocado
    return resultado, len(invertivel)


def _parse_dates(df: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    df = df.copy()
    corrigidas = 0
    for col in DATE_COLS:
        if col in df.columns:
            df[col], n = _to_datetime(df[col])
            corrigidas += n
    return df, corrigidas


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


def _validate(sheets: dict[str, pd.DataFrame], nome: str) -> dict[str, pd.DataFrame]:
    """
    Confere abas e colunas obrigatórias e devolve as três abas do Conecta+
    já indexadas pelo nome canônico (independente de acento/caixa/espaços).
    """
    mapa = _resolver(REQUIRED_SHEETS, list(sheets))
    faltando = [s for s in REQUIRED_SHEETS if s not in mapa]
    if faltando:
        encontradas = ", ".join(f"'{s}'" for s in sheets) or "nenhuma"
        raise ValueError(
            f"Arquivo '{nome}' não contém a(s) aba(s): "
            f"{', '.join(repr(s) for s in faltando)}. "
            f"Abas encontradas: {encontradas}. "
            f"Verifique se é um arquivo exportado pelo Conecta+ Automação "
            f"e se o download foi concluído (não envie a planilha aberta no Excel)."
        )

    resolvidas: dict[str, pd.DataFrame] = {}
    for sheet in REQUIRED_SHEETS:
        df = sheets[mapa[sheet]]
        cols = _resolver(REQUIRED_COLS, list(df.columns))
        missing = [c for c in REQUIRED_COLS if c not in cols]
        if missing:
            raise ValueError(
                f"Aba '{mapa[sheet]}' em '{nome}' não possui as colunas: {missing}. "
                f"Colunas encontradas: {', '.join(str(c) for c in df.columns)}. "
                f"Verifique o formato do arquivo."
            )
        # Renomeia para os nomes canônicos usados no resto do app.
        resolvidas[sheet] = df.rename(columns={real: alvo for alvo, real in cols.items()})
    return resolvidas


# ---------------------------------------------------------------------------
# API pública
# ---------------------------------------------------------------------------

def load_file(uploaded_file: IO, nome_arquivo: str = "") -> AuditData:
    """
    Lê um arquivo xlsx e retorna AuditData.
    Raises ValueError com mensagem em português se o arquivo for inválido.
    """
    nome = nome_arquivo or getattr(uploaded_file, "name", "arquivo.xlsx")
    if hasattr(uploaded_file, "seek"):
        uploaded_file.seek(0)
    try:
        sheets: dict[str, pd.DataFrame] = pd.read_excel(
            uploaded_file, sheet_name=None, engine="openpyxl", dtype=str
        )
    except Exception as e:
        raise ValueError(f"Não foi possível ler o arquivo '{nome}': {e}") from e

    abas = _validate(sheets, nome)

    # As abas de triadas/não triadas repetem linhas de "Todas as Tarefas"; só
    # esta última entra na contagem, para não multiplicar o número exibido.
    todas, corrigidas = _parse_dates(abas["Todas as Tarefas"])
    triadas, _ = _parse_dates(abas["Tarefas Triadas"])
    nao_triadas, _ = _parse_dates(abas["Tarefas Não Triadas"])

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
        datas_corrigidas=corrigidas,
    )


def load_distribution_file(uploaded_file: IO, nome_arquivo: str = "") -> DistribuicaoData:
    """
    Lê um arquivo xlsx de Distribuição do Super Sapiens e retorna DistribuicaoData.
    O arquivo tem metadados nas primeiras linhas; o cabeçalho real está na linha
    que contém 'Id' na primeira coluna.
    Raises ValueError com mensagem em português se o arquivo for inválido.
    """
    nome = nome_arquivo or getattr(uploaded_file, "name", "arquivo.xlsx")
    if hasattr(uploaded_file, "seek"):
        uploaded_file.seek(0)
    try:
        df_raw: pd.DataFrame = pd.read_excel(
            uploaded_file, sheet_name=0, header=None, engine="openpyxl", dtype=str
        )
    except Exception as e:
        raise ValueError(f"Não foi possível ler o arquivo '{nome}': {e}") from e

    # Localiza a linha de cabeçalho (primeira coluna == "Id")
    header_row: int | None = None
    for idx in df_raw.index:
        if str(df_raw.iloc[idx, 0]).strip() == "Id":
            header_row = idx
            break

    if header_row is None:
        raise ValueError(
            f"Arquivo '{nome}' não parece ser um relatório de Distribuição do Super Sapiens. "
            "Não foi encontrado o cabeçalho 'Id, NUP, …' esperado nas primeiras linhas."
        )

    # Extrai parâmetros e usuário distribuidor das linhas de metadados
    usuario_distribuidor: str | None = None
    params_lines: list[str] = []
    for i in range(header_row):
        val = str(df_raw.iloc[i, 0]).strip()
        if val.lower().startswith("usuario:"):
            usuario_distribuidor = val.split(":", 1)[1].strip()
        if val.lower().startswith(("usuario:", "datahorainicio:", "datahorafim:")):
            params_lines.append(val)

    # Monta DataFrame de dados
    cols = [str(c).strip() for c in df_raw.iloc[header_row]]
    df_data = df_raw.iloc[header_row + 1:].copy()
    df_data.columns = cols
    df_data = df_data.dropna(how="all").reset_index(drop=True)

    missing = [c for c in DIST_REQUIRED_COLS if c not in df_data.columns]
    if missing:
        raise ValueError(
            f"Arquivo '{nome}' não contém as colunas esperadas: {missing}. "
            "Verifique se é um relatório de Distribuição exportado do Super Sapiens."
        )

    # Parseia coluna de data/hora
    periodo_inicio: datetime | None = None
    periodo_fim: datetime | None = None
    if COL_DIST_DATA_HORA in df_data.columns:
        df_data[COL_DIST_DATA_HORA] = pd.to_datetime(
            df_data[COL_DIST_DATA_HORA], format=DIST_DATE_FORMAT, errors="coerce"
        )
        valid = df_data[COL_DIST_DATA_HORA].dropna()
        if not valid.empty:
            periodo_inicio = valid.min()
            periodo_fim = valid.max()

    return DistribuicaoData(
        nome_arquivo=nome,
        periodo_inicio=periodo_inicio,
        periodo_fim=periodo_fim,
        df=df_data,
        total_distribuicoes=len(df_data),
        usuario_distribuidor=usuario_distribuidor,
        params_raw="\n".join(params_lines) if params_lines else None,
    )


def detect_file_type(uploaded_file: IO, nome_arquivo: str = "") -> str:
    """
    Detecta o tipo de arquivo Excel sem consumir o objeto de arquivo.
    Retorna "conecta_triagem" ou "supp_distribuicao".
    """
    if hasattr(uploaded_file, "seek"):
        uploaded_file.seek(0)
    try:
        sheet_names = pd.ExcelFile(uploaded_file, engine="openpyxl").sheet_names
    except Exception:
        return "conecta_triagem"
    finally:
        # reset para leituras posteriores
        if hasattr(uploaded_file, "seek"):
            uploaded_file.seek(0)

    if _resolver(REQUIRED_SHEETS, list(sheet_names)):
        return "conecta_triagem"
    return "supp_distribuicao"


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
        datas_corrigidas=sum(f.datas_corrigidas for f in files),
    )
