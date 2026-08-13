# Detalhamento Individual PGF — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Adicionar ao app um terceiro tipo de auditoria, alimentado pela planilha "Detalhamento Individual PGF" do Power BI, em que a população amostrada é o NUP e o auditor examina, no painel de conferência, as atividades realmente lançadas naquele processo no Super Sapiens.

**Architecture:** O novo tipo espelha o fluxo já existente de "Distribuição SS": um leitor em `modules/excel_loader.py` produz um dataclass, `modules/state.py` guarda o estado, `app.py` ganha três páginas (importação, auditoria, relatório) e `modules/report.py` gera o `.docx`. A novidade em relação aos fluxos atuais é o painel de conferência, que faz drill-down NUP → processo → tarefas → atividades usando os clientes HTTP já existentes.

**Tech Stack:** Python 3.9+, Streamlit, pandas, openpyxl, python-docx, httpx, pytest (novo, só para o leitor).

## Global Constraints

- Spec de referência: `docs/superpowers/specs/2026-08-13-detalhamento-individual-design.md`.
- Toda a interface e todas as mensagens de erro em **português**, no mesmo tom das existentes.
- **Uma planilha por auditoria** — não existe consolidação de múltiplos arquivos deste tipo.
- A unidade amostral é o **NUP**; o julgamento é livre (Conforme / Não Conforme + motivo + ação corretiva), idêntico aos demais fluxos.
- Amostragem: `calcular_amostra()` de `modules/sampling.py` (95% de confiança, ±5%) — não reimplementar.
- Identificador interno do tipo: a string `"detalhamento_individual"`, usada em `st.session_state["tipo_relatorio"]`.
- Planilha de referência para testes manuais: `planilhas/2026/detalhamento_individual_2025.xlsx` (12.740 linhas, 12.731 NUPs, 14.036 atividades).
- Interpretador do projeto: `.venv/Scripts/python.exe` (Windows). Rodar o app com `.venv/Scripts/python.exe -m streamlit run app.py`.
- Commits em português, prefixados com `feat:`, `test:` ou `chore:`.

---

### Task 1: Leitor da planilha

**Files:**
- Modify: `modules/excel_loader.py` (acrescentar ao final das constantes, dos dataclasses e da API pública)
- Create: `tests/test_excel_loader_detalhamento.py`
- Create: `requirements-dev.txt`

**Interfaces:**
- Consumes: `_norm`, `_resolver` (helpers privados já existentes em `modules/excel_loader.py`).
- Produces:
  - Constantes `COL_DET_NUP = "NUP"`, `COL_DET_RESPONSAVEL = "Responsável"`, `COL_DET_USUARIO = "Usuário que realizou a atividade"`, `COL_DET_ATIVIDADES = "Atividades"`.
  - `@dataclass DetalhamentoData` com os campos `nome_arquivo: str`, `usuario: str | None`, `unidade: str | None`, `regiao: str | None`, `meses: str | None`, `filtros_raw: str | None`, `df: pd.DataFrame`, `total_nups: int`, `total_atividades: int`, `total_responsaveis: int`.
  - `load_detalhamento_file(uploaded_file: IO, nome_arquivo: str = "") -> DetalhamentoData`.

- [ ] **Step 1: Criar `requirements-dev.txt`**

```
-r requirements.txt
pytest>=8.0
```

Instalar: `.venv/Scripts/python.exe -m pip install -r requirements-dev.txt`

- [ ] **Step 2: Escrever o teste que falha**

Criar `tests/test_excel_loader_detalhamento.py`. O fixture monta em memória uma planilha com a mesma forma do arquivo real: bloco de filtros em A1, linha em branco, cabeçalho, dados.

```python
"""Testes do leitor da planilha Detalhamento Individual PGF."""
from __future__ import annotations

import io

import pandas as pd
import pytest
from openpyxl import Workbook

from modules.excel_loader import (
    COL_DET_ATIVIDADES,
    COL_DET_NUP,
    COL_DET_RESPONSAVEL,
    COL_DET_USUARIO,
    load_detalhamento_file,
)

FILTROS = (
    "Filtros aplicados:\n"
    "Mês é janeiro, fevereiro\n"
    "unidades.regiao é 1ª Região\n"
    "unidades.nome é PSF EM SAO PAULO\n"
    "USUARIO é FULANO DE TAL"
)

LINHAS = [
    ("00410043865202503", "MARIA SOUZA", "FULANO DE TAL", 3),
    ("00410096579202532", "JOAO LIMA", "FULANO DE TAL", 1),
    # Mesmo NUP em duas linhas, com responsáveis diferentes: deve virar uma só.
    ("00424054601202437", "MARIA SOUZA", "FULANO DE TAL", 2),
    ("00424054601202437", "JOAO LIMA", "FULANO DE TAL", 4),
]


def _planilha(filtros: str = FILTROS, linhas=LINHAS, cabecalho=None) -> io.BytesIO:
    cabecalho = cabecalho or [
        COL_DET_NUP, COL_DET_RESPONSAVEL, COL_DET_USUARIO, COL_DET_ATIVIDADES
    ]
    wb = Workbook()
    ws = wb.active
    ws["A1"] = filtros
    ws.append([])            # linha 2 (em branco)
    ws.append([])            # placeholder; o cabeçalho vai na linha 3
    ws.append(cabecalho)
    for linha in linhas:
        ws.append(list(linha))
    # remove a linha placeholder para deixar: A1 filtros, linha 2 vazia, linha 3 cabeçalho
    ws.delete_rows(3)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def test_agrega_por_nup_somando_atividades():
    data = load_detalhamento_file(_planilha(), "detalhamento.xlsx")

    assert data.total_nups == 3
    assert data.total_atividades == 10
    assert len(data.df) == 3

    linha = data.df[data.df[COL_DET_NUP] == "00424054601202437"].iloc[0]
    assert linha[COL_DET_ATIVIDADES] == 6
    assert "MARIA SOUZA" in linha[COL_DET_RESPONSAVEL]
    assert "JOAO LIMA" in linha[COL_DET_RESPONSAVEL]


def test_extrai_filtros_aplicados():
    data = load_detalhamento_file(_planilha(), "detalhamento.xlsx")

    assert data.usuario == "FULANO DE TAL"
    assert data.unidade == "PSF EM SAO PAULO"
    assert data.regiao == "1ª Região"
    assert data.meses == "janeiro, fevereiro"
    assert data.filtros_raw is not None and "Filtros aplicados" in data.filtros_raw


def test_conta_responsaveis_distintos():
    data = load_detalhamento_file(_planilha(), "detalhamento.xlsx")
    assert data.total_responsaveis == 2


def test_cabecalho_tolera_acento_e_caixa():
    cabecalho = ["nup", "RESPONSAVEL", "usuario que realizou a atividade", "atividades"]
    data = load_detalhamento_file(_planilha(cabecalho=cabecalho), "detalhamento.xlsx")

    assert data.total_nups == 3
    assert COL_DET_NUP in data.df.columns
    assert COL_DET_ATIVIDADES in data.df.columns


def test_atividades_nao_numericas_viram_zero():
    linhas = [("00410043865202503", "MARIA SOUZA", "FULANO DE TAL", "n/d")]
    data = load_detalhamento_file(_planilha(linhas=linhas), "detalhamento.xlsx")

    assert data.total_atividades == 0
    assert data.df.iloc[0][COL_DET_ATIVIDADES] == 0


def test_sem_cabecalho_nup_levanta_erro_em_portugues():
    cabecalho = ["Id", "Outro", "Coisa", "Qualquer"]
    with pytest.raises(ValueError, match="Detalhamento Individual"):
        load_detalhamento_file(_planilha(cabecalho=cabecalho), "errado.xlsx")


def test_coluna_faltando_levanta_erro_listando_a_coluna():
    cabecalho = [COL_DET_NUP, COL_DET_RESPONSAVEL, COL_DET_USUARIO, "Outra"]
    with pytest.raises(ValueError, match="Atividades"):
        load_detalhamento_file(_planilha(cabecalho=cabecalho), "incompleto.xlsx")


def test_filtros_ausentes_nao_quebram_a_leitura():
    data = load_detalhamento_file(_planilha(filtros=""), "sem_filtros.xlsx")

    assert data.total_nups == 3
    assert data.usuario is None
    assert data.unidade is None
```

- [ ] **Step 3: Rodar o teste e confirmar que falha**

Run: `.venv/Scripts/python.exe -m pytest tests/test_excel_loader_detalhamento.py -v`
Expected: FAIL com `ImportError: cannot import name 'COL_DET_NUP' from 'modules.excel_loader'`

- [ ] **Step 4: Acrescentar as constantes e o dataclass**

Em `modules/excel_loader.py`, depois do bloco de constantes da Distribuição SS (após `DIST_REQUIRED_COLS`):

```python
# ---------------------------------------------------------------------------
# Constantes de colunas — Detalhamento Individual PGF (Power BI)
# ---------------------------------------------------------------------------

COL_DET_NUP = "NUP"
COL_DET_RESPONSAVEL = "Responsável"
COL_DET_USUARIO = "Usuário que realizou a atividade"
COL_DET_ATIVIDADES = "Atividades"

DET_REQUIRED_COLS = [COL_DET_NUP, COL_DET_RESPONSAVEL, COL_DET_USUARIO, COL_DET_ATIVIDADES]

# Chaves do bloco "Filtros aplicados:" da primeira célula, na forma "<chave> é <valor>".
DET_FILTROS = {
    "meses": "Mês",
    "regiao": "unidades.regiao",
    "unidade": "unidades.nome",
    "usuario": "USUARIO",
}
```

E, junto dos demais dataclasses:

```python
@dataclass
class DetalhamentoData:
    nome_arquivo: str
    usuario: str | None
    unidade: str | None
    regiao: str | None
    meses: str | None
    filtros_raw: str | None
    df: pd.DataFrame            # uma linha por NUP
    total_nups: int
    total_atividades: int
    total_responsaveis: int
```

- [ ] **Step 5: Implementar o parser dos filtros**

Função interna, junto dos demais helpers privados:

```python
def _parse_filtros_detalhamento(linhas: list[str]) -> dict[str, str | None]:
    """
    Lê o bloco 'Filtros aplicados:' e devolve {'usuario', 'unidade', 'regiao', 'meses'}.

    Cada filtro vem numa linha da forma '<chave> é <valor>'. A comparação da
    chave é tolerante a acento e caixa; valores ausentes ficam None.
    """
    achados: dict[str, str | None] = {campo: None for campo in DET_FILTROS}
    indice = {_norm(chave): campo for campo, chave in DET_FILTROS.items()}
    for linha in linhas:
        for sep in (" é ", " e "):
            if sep in linha:
                chave, valor = linha.split(sep, 1)
                campo = indice.get(_norm(chave))
                if campo and achados[campo] is None:
                    achados[campo] = valor.strip()
                break
    return achados
```

O separador alternativo `" e "` cobre exports em que o acento se perde.

- [ ] **Step 6: Implementar `load_detalhamento_file`**

Na seção "API pública", depois de `load_distribution_file`:

```python
def load_detalhamento_file(uploaded_file: IO, nome_arquivo: str = "") -> DetalhamentoData:
    """
    Lê a planilha 'Detalhamento Individual PGF' exportada do Power BI.

    A primeira célula traz o bloco 'Filtros aplicados:' e o cabeçalho real está
    na primeira linha cuja coluna A seja 'NUP'. Como a população auditada é o
    NUP, as linhas são agregadas por NUP: as atividades são somadas e os
    responsáveis distintos, concatenados.

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

    # Localiza a linha de cabeçalho (primeira coluna == "NUP")
    header_row: int | None = None
    for idx in df_raw.index:
        if _norm(df_raw.iloc[idx, 0]) == _norm(COL_DET_NUP):
            header_row = idx
            break

    if header_row is None:
        raise ValueError(
            f"Arquivo '{nome}' não parece ser um relatório de Detalhamento Individual "
            "do Power BI. Não foi encontrado o cabeçalho 'NUP, Responsável, …' "
            "esperado nas primeiras linhas."
        )

    # Bloco de filtros: tudo o que vem antes do cabeçalho, na coluna A.
    linhas_filtro: list[str] = []
    for i in range(header_row):
        val = df_raw.iloc[i, 0]
        if val is None or (isinstance(val, float) and pd.isna(val)):
            continue
        linhas_filtro.extend(
            parte.strip() for parte in str(val).split("\n") if parte.strip()
        )
    filtros = _parse_filtros_detalhamento(linhas_filtro)
    filtros_raw = "\n".join(linhas_filtro) if linhas_filtro else None

    cols = [str(c).strip() for c in df_raw.iloc[header_row]]
    df_data = df_raw.iloc[header_row + 1:].copy()
    df_data.columns = cols
    df_data = df_data.dropna(how="all").reset_index(drop=True)

    resolvidas = _resolver(DET_REQUIRED_COLS, list(df_data.columns))
    missing = [c for c in DET_REQUIRED_COLS if c not in resolvidas]
    if missing:
        raise ValueError(
            f"Arquivo '{nome}' não contém as colunas esperadas: {missing}. "
            f"Colunas encontradas: {', '.join(str(c) for c in df_data.columns)}. "
            "Verifique se é a planilha 'Detalhamento Individual PGF' exportada do Power BI."
        )
    df_data = df_data.rename(columns={real: alvo for alvo, real in resolvidas.items()})
    df_data = df_data[DET_REQUIRED_COLS]

    df_data[COL_DET_NUP] = df_data[COL_DET_NUP].astype(str).str.strip()
    df_data[COL_DET_ATIVIDADES] = (
        pd.to_numeric(df_data[COL_DET_ATIVIDADES], errors="coerce").fillna(0).astype(int)
    )
    total_responsaveis = df_data[COL_DET_RESPONSAVEL].dropna().nunique()

    # Uma linha por NUP: soma as atividades e junta os responsáveis distintos.
    agregado = (
        df_data.groupby(COL_DET_NUP, as_index=False, sort=False)
        .agg({
            COL_DET_RESPONSAVEL: lambda s: "; ".join(
                dict.fromkeys(v for v in s.dropna().astype(str) if v.strip())
            ),
            COL_DET_USUARIO: "first",
            COL_DET_ATIVIDADES: "sum",
        })
    )

    return DetalhamentoData(
        nome_arquivo=nome,
        usuario=filtros["usuario"],
        unidade=filtros["unidade"],
        regiao=filtros["regiao"],
        meses=filtros["meses"],
        filtros_raw=filtros_raw,
        df=agregado,
        total_nups=len(agregado),
        total_atividades=int(agregado[COL_DET_ATIVIDADES].sum()),
        total_responsaveis=int(total_responsaveis),
    )
```

- [ ] **Step 7: Rodar os testes e confirmar que passam**

Run: `.venv/Scripts/python.exe -m pytest tests/test_excel_loader_detalhamento.py -v`
Expected: 8 passed

- [ ] **Step 8: Conferir contra o arquivo real**

Run:

```bash
.venv/Scripts/python.exe -c "from modules.excel_loader import load_detalhamento_file as f; d=f(open('planilhas/2026/detalhamento_individual_2025.xlsx','rb')); print(d.total_nups, d.total_atividades, d.usuario, d.unidade)"
```

Expected: `12731 14036` seguido do usuário e da unidade do arquivo (valores não vazios).

- [ ] **Step 9: Commit**

```bash
rtk git add modules/excel_loader.py tests/test_excel_loader_detalhamento.py requirements-dev.txt
rtk git commit -m "feat: leitor da planilha Detalhamento Individual PGF"
```

---

### Task 2: Detecção do tipo de arquivo e estado da sessão

**Files:**
- Modify: `modules/excel_loader.py` (`detect_file_type`)
- Modify: `modules/state.py` (`_PERSIST_KEYS`, `_DEFAULTS`, `reset_auditoria`, helpers, `get_session_info`)
- Modify: `tests/test_excel_loader_detalhamento.py` (novos testes de detecção)

**Interfaces:**
- Consumes: `load_detalhamento_file`, `DetalhamentoData`, constantes `COL_DET_*` (Task 1).
- Produces:
  - `detect_file_type()` passa a devolver também `"detalhamento_individual"`.
  - `modules.state.get_det_data()` → `DetalhamentoData | None`.
  - `modules.state.get_df_detalhamento()` → `pd.DataFrame | None`.
  - Chaves de sessão `det_data`, `df_audit_detalhamento`, `auditoria_detalhamento_concluida`.

- [ ] **Step 1: Escrever os testes de detecção que falham**

Acrescentar ao final de `tests/test_excel_loader_detalhamento.py`:

```python
def test_detect_file_type_reconhece_detalhamento():
    from modules.excel_loader import detect_file_type

    assert detect_file_type(_planilha(), "detalhamento.xlsx") == "detalhamento_individual"


def test_detect_file_type_nao_confunde_com_distribuicao():
    from modules.excel_loader import detect_file_type

    cabecalho = ["Id", "NUP", "Setor_destino", "DataHoraDistribuicao"]
    assert detect_file_type(_planilha(cabecalho=cabecalho), "dist.xlsx") == "supp_distribuicao"
```

- [ ] **Step 2: Rodar e confirmar a falha**

Run: `.venv/Scripts/python.exe -m pytest tests/test_excel_loader_detalhamento.py -k detect -v`
Expected: FAIL — `detect_file_type` devolve `"supp_distribuicao"` para a planilha de detalhamento.

- [ ] **Step 3: Estender `detect_file_type`**

Substituir o corpo atual pela versão abaixo. A detecção do Detalhamento olha a primeira aba: precisa existir uma linha cuja coluna A seja `NUP` e, nessa linha, uma coluna `Atividades`.

```python
def detect_file_type(uploaded_file: IO, nome_arquivo: str = "") -> str:
    """
    Detecta o tipo de arquivo Excel sem consumir o objeto de arquivo.
    Retorna "conecta_triagem", "detalhamento_individual" ou "supp_distribuicao".
    """
    if hasattr(uploaded_file, "seek"):
        uploaded_file.seek(0)
    try:
        sheet_names = pd.ExcelFile(uploaded_file, engine="openpyxl").sheet_names
    except Exception:
        return "conecta_triagem"
    finally:
        if hasattr(uploaded_file, "seek"):
            uploaded_file.seek(0)

    if _resolver(REQUIRED_SHEETS, list(sheet_names)):
        return "conecta_triagem"

    # Detalhamento Individual: cabeçalho 'NUP' na coluna A e coluna 'Atividades'.
    try:
        amostra = pd.read_excel(
            uploaded_file, sheet_name=0, header=None, engine="openpyxl",
            dtype=str, nrows=15,
        )
        for idx in amostra.index:
            linha = [_norm(v) for v in amostra.iloc[idx].tolist() if v is not None]
            if linha and linha[0] == _norm(COL_DET_NUP) and _norm(COL_DET_ATIVIDADES) in linha:
                return "detalhamento_individual"
    except Exception:
        pass
    finally:
        if hasattr(uploaded_file, "seek"):
            uploaded_file.seek(0)

    return "supp_distribuicao"
```

- [ ] **Step 4: Rodar os testes**

Run: `.venv/Scripts/python.exe -m pytest tests/test_excel_loader_detalhamento.py -v`
Expected: 10 passed

- [ ] **Step 5: Acrescentar as chaves de sessão**

Em `modules/state.py`, dentro de `_PERSIST_KEYS`, logo após `"dist_data",`:

```python
    "det_data",
```

e logo após `"auditoria_distribuicao_concluida",`:

```python
    "df_audit_detalhamento",
    "auditoria_detalhamento_concluida",
```

Em `_DEFAULTS`, logo após a linha `"auditoria_distribuicao_concluida": False,`:

```python
    # detalhamento individual
    "det_data": None,
    "df_audit_detalhamento": None,
    "auditoria_detalhamento_concluida": False,
```

Conferir que `"det_data": None` não duplica outra entrada já presente no dict.

- [ ] **Step 6: Acrescentar os helpers de leitura**

Em `modules/state.py`, depois de `get_dist_data()`:

```python
def get_det_data():
    return st.session_state.get("det_data")
```

e depois de `get_df_distribuicao()`:

```python
def get_df_detalhamento() -> pd.DataFrame | None:
    return st.session_state.get("df_audit_detalhamento")
```

- [ ] **Step 7: Incluir o novo tipo em `reset_auditoria` e `get_session_info`**

Em `reset_auditoria()`, acrescentar `df_audit_detalhamento` e `auditoria_detalhamento_concluida` à lista de chaves reinicializadas, no mesmo formato usado pelas chaves de distribuição já presentes na função.

Em `get_session_info()`, substituir o bloco que resolve `nome`:

```python
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
```

- [ ] **Step 8: Verificar que o módulo importa e o estado inicializa**

Run: `.venv/Scripts/python.exe -c "import modules.state as s; print('det_data' in s._DEFAULTS, 'det_data' in s._PERSIST_KEYS, hasattr(s,'get_det_data'), hasattr(s,'get_df_detalhamento'))"`
Expected: `True True True True`

- [ ] **Step 9: Commit**

```bash
rtk git add modules/excel_loader.py modules/state.py tests/test_excel_loader_detalhamento.py
rtk git commit -m "feat: detecção e estado do tipo Detalhamento Individual"
```

---

### Task 3: Navegação e página de importação

**Files:**
- Modify: `app.py` — imports do topo; constantes de páginas (perto da linha 260); `_get_paginas()` (linha 274); `render_importacao()` (linha 1148); nova função `_render_importacao_detalhamento()` (inserir depois de `_render_importacao_distribuicao`, que termina na linha 1446)

**Interfaces:**
- Consumes: `load_detalhamento_file`, `DetalhamentoData`, `COL_DET_*` (Task 1); `get_det_data` (Task 2).
- Produces: `st.session_state["det_data"]` preenchido e `st.session_state["pagina"] = "detalhamento"`; chave de página `"detalhamento"` no menu lateral.

- [ ] **Step 1: Acrescentar os imports**

No bloco de imports de `modules.excel_loader` no topo de `app.py`, incluir:

```python
    COL_DET_ATIVIDADES,
    COL_DET_NUP,
    COL_DET_RESPONSAVEL,
    COL_DET_USUARIO,
    DetalhamentoData,
    load_detalhamento_file,
```

e, no bloco de imports de `modules.state`, incluir `get_det_data` e `get_df_detalhamento`.

- [ ] **Step 2: Declarar o mapa de páginas do novo tipo**

Logo depois de `PAGINAS_DISTRIBUICAO` (linha ~267), no mesmo formato do dicionário existente (chave interna → rótulo exibido no menu):

```python
PAGINAS_DETALHAMENTO = {
    "importacao": "📂 Importação",
    "detalhamento": "📊 Auditoria de Atividades",
    "relatorio": "📄 Relatório",
}
```

Conferir os rótulos de `PAGINAS_DISTRIBUICAO` e usar exatamente a mesma convenção de emoji e capitalização.

- [ ] **Step 3: Rotear em `_get_paginas`**

Em `_get_paginas()` (linha 274), acrescentar o novo tipo antes do retorno padrão:

```python
    if st.session_state.get("tipo_relatorio") == "detalhamento_individual":
        return PAGINAS_DETALHAMENTO
```

- [ ] **Step 4: Acrescentar a terceira opção no rádio de tipo**

Substituir o bloco das linhas 1178–1202 de `render_importacao()` por:

```python
    tipo_opcoes = [
        "Conecta+ — Triagem Avançada",
        "Super Sapiens — Distribuição de Tarefas",
        "Power BI — Detalhamento Individual PGF",
    ]
    _TIPOS = ["conecta_triagem", "supp_distribuicao", "detalhamento_individual"]
    tipo_idx = _TIPOS.index(tipo_rel_atual) if tipo_rel_atual in _TIPOS else 0
    tipo_sel = st.radio(
        "Tipo de relatório:",
        tipo_opcoes,
        index=tipo_idx,
        horizontal=True,
        key="radio_tipo_rel",
    )
    novo_tipo = _TIPOS[tipo_opcoes.index(tipo_sel)]
    if novo_tipo != tipo_rel_atual:
        reset_auditoria()
        st.session_state["tipo_relatorio"] = novo_tipo
        st.session_state["audit_data_merged"] = None
        st.session_state["dist_data"] = None
        st.session_state["det_data"] = None

    st.divider()

    if novo_tipo == "conecta_triagem":
        _render_importacao_triagem()
    elif novo_tipo == "supp_distribuicao":
        _render_importacao_distribuicao()
    else:
        _render_importacao_detalhamento()
```

- [ ] **Step 5: Rotular o tipo na oferta de restauração de sessão**

Substituir o cálculo de `tipo_label` (linhas 1155–1158) por:

```python
            tipo_label = {
                "supp_distribuicao": "Distribuição SS",
                "detalhamento_individual": "Detalhamento Individual",
            }.get(info.get("tipo_relatorio"), "Triagem Conecta+")
```

- [ ] **Step 6: Escrever `_render_importacao_detalhamento`**

Inserir logo após o fim de `_render_importacao_distribuicao` (antes de `def render_auditoria_triadas`):

```python
def _render_importacao_detalhamento() -> None:
    """Sub-renderização da importação no modo Power BI — Detalhamento Individual."""
    st.markdown(
        "Envie a planilha **Detalhamento Individual PGF** exportada do Power BI "
        "(Página Inicial → PGF → PGF - Painéis Estratégicos → Detalhamento Individual). "
        "A auditoria trata de um arquivo por vez."
    )

    arquivo = st.file_uploader(
        "Planilha de Detalhamento Individual (.xlsx)",
        type=["xlsx"],
        accept_multiple_files=False,
        key="uploader_detalhamento",
    )

    if arquivo is None:
        det_data = get_det_data()
        if det_data is not None:
            st.success(f"Arquivo carregado: **{det_data.nome_arquivo}**")
            _render_resumo_detalhamento(det_data)
            if st.button("Iniciar Auditoria →", type="primary"):
                st.session_state["pagina"] = "detalhamento"
                st.rerun()
        return

    try:
        det_data = load_detalhamento_file(arquivo, arquivo.name)
    except ValueError as e:
        st.error(str(e))
        return

    st.session_state["det_data"] = det_data
    st.success(f"Arquivo lido: **{det_data.nome_arquivo}**")
    _render_resumo_detalhamento(det_data)

    if st.button("Iniciar Auditoria →", type="primary"):
        save_session()
        st.session_state["pagina"] = "detalhamento"
        st.rerun()


def _render_resumo_detalhamento(det_data: DetalhamentoData) -> None:
    """Cartão com os filtros aplicados e as métricas da planilha."""
    st.subheader("Filtros aplicados na extração")
    col1, col2 = st.columns(2)
    with col1:
        st.markdown(f"**Usuário:** {det_data.usuario or 'N/D'}")
        st.markdown(f"**Unidade:** {det_data.unidade or 'N/D'}")
    with col2:
        st.markdown(f"**Região:** {det_data.regiao or 'N/D'}")
        st.markdown(f"**Meses:** {det_data.meses or 'N/D'}")

    if det_data.filtros_raw:
        with st.expander("Ver bloco de filtros original"):
            st.code(det_data.filtros_raw)

    st.divider()
    c1, c2, c3 = st.columns(3)
    c1.metric("NUPs (população)", det_data.total_nups)
    c2.metric("Atividades lançadas", det_data.total_atividades)
    c3.metric("Responsáveis distintos", det_data.total_responsaveis)
```

Conferir se `save_session` já está importado no topo de `app.py` (é usado em `render_auditoria_distribuicao`); se não estiver, acrescentá-lo ao import de `modules.state`.

- [ ] **Step 7: Verificar sintaxe e rodar o app**

Run: `.venv/Scripts/python.exe -m py_compile app.py`
Expected: sem saída.

Run: `.venv/Scripts/python.exe -m streamlit run app.py`
Verificar manualmente: a terceira opção aparece no rádio; ao enviar `planilhas/2026/detalhamento_individual_2025.xlsx`, o resumo mostra 12.731 NUPs, 14.036 atividades e os filtros preenchidos; o menu lateral passa a exibir as três páginas do novo fluxo. Encerrar o app com Ctrl+C.

- [ ] **Step 8: Commit**

```bash
rtk git add app.py
rtk git commit -m "feat: importação e navegação do Detalhamento Individual"
```

---

### Task 4: Página de auditoria — amostragem e tabela

**Files:**
- Modify: `app.py` — nova função `render_auditoria_detalhamento()` (inserir depois de `_render_dist_row_editor`, antes de `_render_relatorio_distribuicao`); despacho de páginas no final do arquivo (perto da linha 2200)
- Modify: `app.py:602` — `_render_audit_table`, ajuste do rótulo de entidade

**Interfaces:**
- Consumes: `get_det_data`, `COL_DET_*`, `preparar_df_auditoria`, `calcular_amostra`, `formula_descricao`, `tabela_referencia`, `selecionar_amostra`, `_render_audit_table`.
- Produces: `render_auditoria_detalhamento()`; `st.session_state["df_audit_detalhamento"]` com as colunas `NUP`, `Responsável`, `Usuário que realizou a atividade`, `Atividades`, `Conformidade`, `Motivo da Não Conformidade`, `Ação Corretiva`. Chama `_render_det_row_editor(df_key, orig_idx, row)`, implementado na Task 5.

- [ ] **Step 1: Generalizar o rótulo de entidade da tabela**

Em `_render_audit_table` (linha 621), substituir:

```python
    entidade = "distribuições" if id_col == COL_DIST_ID else "tarefas"
```

por:

```python
    if id_col == COL_DIST_ID:
        entidade = "distribuições"
    elif id_col == COL_DET_NUP:
        entidade = "NUPs"
    else:
        entidade = "tarefas"
```

- [ ] **Step 2: Escrever `render_auditoria_detalhamento`**

Inserir depois do fim de `_render_dist_row_editor`:

```python
def render_auditoria_detalhamento() -> None:
    det_data = get_det_data()
    if det_data is None:
        st.warning("Nenhum arquivo carregado. Volte à página de Importação.")
        return

    st.title("📊 Auditoria de Atividades por NUP")
    st.caption(
        "Verifique, em cada NUP sorteado, se as atividades lançadas no Super Sapiens "
        "correspondem ao que consta do processo."
    )

    # ── Seleção do tipo de controle ──────────────────────────────────────────
    if st.session_state.get("tipo_controle") is None:
        total = det_data.total_nups
        st.markdown(
            f"**{total}** NUPs com atividades lançadas por "
            f"**{det_data.usuario or 'usuário não identificado'}** "
            f"({det_data.total_atividades} atividades no total). "
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
                key="radio_tipo_det",
                label_visibility="collapsed",
            )
            st.markdown("""
            **Controle Simplificado** — Verificação de todos os NUPs (ou seleção manual).

            **Controle Detalhado** — Amostragem estatística com seleção aleatória.
            Nível de confiança **95%**, margem de erro **±5%**.
            """)
        with col_dir:
            st.markdown("#### Tamanho da Amostra (Controle Detalhado)")
            if total > 0:
                n = calcular_amostra(total)
                st.metric("NUPs a auditar", n, delta=f"{n / total * 100:.1f}% do universo")
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
                df_base = selecionar_amostra(det_data.df, n)
            else:
                st.session_state["tamanho_amostra"] = None
                df_base = det_data.df.copy()

            colunas = [
                COL_DET_NUP, COL_DET_RESPONSAVEL, COL_DET_USUARIO, COL_DET_ATIVIDADES,
            ]
            st.session_state["df_audit_detalhamento"] = preparar_df_auditoria(df_base, colunas)
            save_session()
            st.rerun()
        return

    # ── Editor ──────────────────────────────────────────────────────────────
    tipo_controle = st.session_state["tipo_controle"]
    tipo_label = "Controle Simplificado" if tipo_controle == "simplificado" else "Controle Detalhado"
    n_amostra = st.session_state.get("tamanho_amostra")
    df = st.session_state.get("df_audit_detalhamento")

    if df is None:
        st.error("Estado inconsistente. Clique em 'Nova Auditoria' no menu lateral.")
        return

    descr = f"Amostra: {n_amostra} NUPs" if n_amostra else f"Total: {len(df)} NUPs"
    st.markdown(
        f"<span class='ac-badge'>{tipo_label}</span> &nbsp; {descr}",
        unsafe_allow_html=True,
    )
    st.divider()

    col_tab, col_edit = st.columns([3, 2])
    with col_tab:
        orig_idx, row = _render_audit_table(
            df_key="df_audit_detalhamento",
            filtro_key="filtro_det",
            busca_key="busca_det",
            column_order=[
                COL_DET_NUP, COL_DET_RESPONSAVEL, COL_DET_ATIVIDADES,
                COL_CONFORMIDADE, COL_MOTIVO, COL_ACAO,
            ],
            table_key="tabela_det",
            id_col=COL_DET_NUP,
            nup_col=COL_DET_NUP,
        )
    with col_edit:
        if orig_idx is None:
            st.info("Selecione uma linha na tabela para auditar.")
        else:
            _render_det_row_editor("df_audit_detalhamento", orig_idx, row)
```

Antes de rodar, conferir nas linhas 1760–1815 como `render_auditoria_distribuicao` monta o badge, as colunas e a chamada de `_render_audit_table`, e alinhar exatamente a este padrão (nomes dos parâmetros, proporção das colunas, mensagem de "selecione uma linha", botão de concluir auditoria se existir). Replicar também o bloco que marca `auditoria_distribuicao_concluida`, trocando pela chave `auditoria_detalhamento_concluida`.

- [ ] **Step 3: Stub temporário do editor de linha**

Para que a página rode antes da Task 5, inserir logo depois:

```python
def _render_det_row_editor(df_key: str, orig_idx, row: dict) -> None:
    st.markdown("#### ✏️ Auditoria")
    st.write(row)
```

Este stub é substituído integralmente na Task 5.

- [ ] **Step 4: Ligar a página ao despacho**

No bloco final de `app.py` que despacha as páginas (`if pagina == "distribuicao": render_auditoria_distribuicao()` e similares — conferir o formato exato usado), acrescentar:

```python
    elif pagina == "detalhamento":
        render_auditoria_detalhamento()
```

- [ ] **Step 5: Verificar**

Run: `.venv/Scripts/python.exe -m py_compile app.py`
Expected: sem saída.

Run: `.venv/Scripts/python.exe -m streamlit run app.py`
Verificar manualmente: importar a planilha real, escolher Controle Detalhado, confirmar que a amostra calculada é **373** para 12.731 NUPs, que a tabela lista 373 linhas com as colunas de auditoria e que selecionar uma linha exibe o stub. Encerrar com Ctrl+C.

- [ ] **Step 6: Commit**

```bash
rtk git add app.py
rtk git commit -m "feat: página de auditoria de atividades por NUP"
```

---

### Task 5: Painel de conferência com drill-down até as atividades

**Files:**
- Modify: `modules/atividade.py` (acrescentar `from_auth`, se ausente)
- Modify: `app.py` — substituir o stub `_render_det_row_editor` pela versão final

**Interfaces:**
- Consumes: `ProcessoClient.buscar_por_nup`, `TarefaClient.listar`, `AtividadeClient.listar_por_tarefa`, `_SUPERSAPIENS_URL`, `OPCOES_CONFORMIDADE`, `COL_CONFORMIDADE`, `COL_MOTIVO`, `COL_ACAO`, `save_session`.
- Produces: `_render_det_row_editor(df_key: str, orig_idx, row: dict) -> None` completo.

- [ ] **Step 1: Conferir as assinaturas reais dos clientes**

Run:

```bash
.venv/Scripts/python.exe -c "import inspect, modules.processo as p, modules.tarefa as t, modules.atividade as a; print(inspect.signature(p.ProcessoClient.buscar_por_nup)); print(inspect.signature(t.TarefaClient.listar)); print([m for m in dir(a.AtividadeClient) if not m.startswith('_')])"
```

Anotar os nomes exatos dos parâmetros (`where`, `populate`, `limit`, `offset`) e do método de listagem de atividades. Se `AtividadeClient` não tiver `from_auth`, seguir o Step 2; se tiver, pular para o Step 3.

- [ ] **Step 2: Acrescentar `from_auth` a `AtividadeClient`**

Copiar a implementação de `from_auth` de `modules/tarefa.py` (é idêntica em todos os módulos) para `modules/atividade.py`, dentro da classe `AtividadeClient`:

```python
    @classmethod
    def from_auth(cls, auth_client) -> "AtividadeClient":
        """Instancia o cliente a partir de um AuthClient autenticado."""
        return cls(auth_client)
```

Conferir a assinatura de `AtividadeClient.__init__` e ajustar os argumentos para bater com ela — se o `__init__` receber `base_url` e `token` separadamente, replicar o que `TarefaClient.from_auth` faz.

Verificar:

Run: `.venv/Scripts/python.exe -c "from modules.atividade import AtividadeClient; print(hasattr(AtividadeClient, 'from_auth'))"`
Expected: `True`

- [ ] **Step 3: Substituir o stub por `_render_det_row_editor`**

Trocar a função-stub inteira por:

```python
def _render_det_row_editor(df_key: str, orig_idx, row: dict) -> None:
    """
    Painel de conferência de um NUP: metadados do processo, tarefas e as
    atividades lançadas em cada tarefa, seguidos do formulário de julgamento.
    """
    nup = str(row.get(COL_DET_NUP) or "").strip()
    responsavel = row.get(COL_DET_RESPONSAVEL)
    usuario_planilha = str(row.get(COL_DET_USUARIO) or "").strip()
    qtd_planilha = row.get(COL_DET_ATIVIDADES)

    # ── Drill-down NUP → processo → tarefas → atividades (com cache) ─────────
    cache_key = f"_det_cache_{nup}"
    if cache_key not in st.session_state:
        auth = st.session_state.get("supp_auth_client")
        if auth and nup:
            with st.spinner("Buscando processo, tarefas e atividades..."):
                try:
                    from modules.atividade import AtividadeClient
                    from modules.processo import ProcessoClient
                    from modules.tarefa import TarefaClient

                    pc = ProcessoClient.from_auth(auth)
                    proc = pc.buscar_por_nup(nup) or {}
                    proc_id = proc.get("id")

                    tarefas_info = []
                    if proc_id:
                        tc = TarefaClient.from_auth(auth)
                        ac = AtividadeClient.from_auth(auth)
                        tarefas = tc.listar(
                            where={"processo.id": proc_id},
                            populate=["especieTarefa", "usuarioResponsavel", "setorResponsavel"],
                            limit=100,
                        )
                        for tarefa in tarefas:
                            atividades = ac.listar_por_tarefa(
                                tarefa.get("id"),
                                populate=["especieAtividade", "usuario", "setor"],
                            )
                            tarefas_info.append({"tarefa": tarefa, "atividades": atividades})

                    st.session_state[cache_key] = {
                        "proc_id": proc_id,
                        "nup_fmt": proc.get("NUPFormatado") or proc.get("NUP") or nup,
                        "tarefas": tarefas_info,
                    }
                except Exception as e:
                    st.session_state[cache_key] = {"erro": str(e)}
        else:
            st.session_state[cache_key] = {}

    cached = st.session_state.get(cache_key, {})
    proc_id = cached.get("proc_id")
    nup_fmt = cached.get("nup_fmt") or nup
    tarefas_info = cached.get("tarefas") or []

    # ── Cabeçalho ────────────────────────────────────────────────────────────
    url = _SUPERSAPIENS_URL.format(proc_id=proc_id) if proc_id else None
    _c_title, _c_open, _c_refresh = st.columns([4, 3, 1])
    with _c_title:
        st.markdown("#### ✏️ Auditoria")
    with _c_open:
        if url:
            st.link_button("↗ SuperSapiens", url, use_container_width=True)
    with _c_refresh:
        if nup and st.button("🔄", key=f"_refresh_det_{nup}", help="Recarregar dados"):
            st.session_state.pop(cache_key, None)
            st.rerun()

    st.markdown(f"**NUP:** `{nup_fmt}`")
    st.markdown(f"**Responsável:** {responsavel or 'N/D'}")
    st.markdown(f"**Usuário auditado:** {usuario_planilha or 'N/D'}")
    st.markdown(f"**Atividades na planilha:** {qtd_planilha}")

    if cached.get("erro"):
        st.warning(f"Não foi possível consultar o Super Sapiens: {cached['erro']}")
    elif not st.session_state.get("supp_auth_client"):
        st.info("Faça login no Super Sapiens para conferir as atividades do processo.")
    elif not tarefas_info:
        st.warning("Nenhuma tarefa encontrada para este NUP no Super Sapiens.")
    else:
        total_atividades = sum(len(t["atividades"]) for t in tarefas_info)
        st.markdown(
            f"**Atividades no Super Sapiens:** {total_atividades} "
            f"em {len(tarefas_info)} tarefa(s)"
        )
        for info in tarefas_info:
            tarefa = info["tarefa"]
            especie = (tarefa.get("especieTarefa") or {}).get("nome") or "Tarefa"
            with st.expander(f"{especie} — {len(info['atividades'])} atividade(s)"):
                if not info["atividades"]:
                    st.caption("Sem atividades registradas nesta tarefa.")
                for ativ in info["atividades"]:
                    esp = (ativ.get("especieAtividade") or {}).get("nome") or "Atividade"
                    usu = (ativ.get("usuario") or {}).get("nome") or "N/D"
                    setor = (ativ.get("setor") or {}).get("nome") or "N/D"
                    data = ativ.get("dataHoraConclusao") or ativ.get("criadoEm") or "N/D"
                    obs = ativ.get("observacao") or ""
                    destaque = "🔹 " if _norm_txt(usu) == _norm_txt(usuario_planilha) else ""
                    st.markdown(
                        f"{destaque}**{esp}** — {usu} · {setor} · {data}"
                        + (f"  \n_{obs}_" if obs else "")
                    )

    st.divider()

    # ── Formulário de julgamento ─────────────────────────────────────────────
    df = st.session_state[df_key]
    conf_atual = df.at[orig_idx, COL_CONFORMIDADE]
    conf = st.radio(
        "Conformidade:",
        OPCOES_CONFORMIDADE,
        index=OPCOES_CONFORMIDADE.index(conf_atual) if conf_atual in OPCOES_CONFORMIDADE else 0,
        key=f"det_conf_{orig_idx}",
        horizontal=True,
    )
    motivo = st.text_area(
        "Motivo da não conformidade:",
        value=df.at[orig_idx, COL_MOTIVO],
        key=f"det_motivo_{orig_idx}",
        placeholder="Ex.: atividade lançada em NUP alheio ao objeto do processo",
    )
    acao = st.text_area(
        "Ação corretiva:",
        value=df.at[orig_idx, COL_ACAO],
        key=f"det_acao_{orig_idx}",
        placeholder="Ex.: solicitar retificação do lançamento ao setor responsável",
    )

    if st.button("💾 Salvar", type="primary", key=f"det_salvar_{orig_idx}"):
        df.at[orig_idx, COL_CONFORMIDADE] = conf
        df.at[orig_idx, COL_MOTIVO] = motivo
        df.at[orig_idx, COL_ACAO] = acao
        st.session_state[df_key] = df
        save_session()
        st.rerun()
```

Acrescentar, junto dos helpers do topo de `app.py`, a normalização usada no destaque:

```python
def _norm_txt(valor: str | None) -> str:
    """Normaliza nome para comparação: sem acento, sem caixa, sem espaços extras."""
    import unicodedata
    texto = unicodedata.normalize("NFKD", str(valor or ""))
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return " ".join(texto.split()).casefold()
```

Antes de rodar, comparar este formulário com o de `_render_dist_row_editor` (linhas ~1918–1972) e alinhar rótulos, chaves de widget e comportamento de salvamento ao que já existe — o padrão do arquivo vence.

- [ ] **Step 4: Ajustar os nomes de campo da API**

Os nomes `especieAtividade`, `dataHoraConclusao`, `observacao`, `usuario`, `setor` são a expectativa a partir da convenção do SUPP. Confirmar contra a spec:

Run: `rtk grep -n "especieAtividade\|dataHoraConclusao\|observacao" spec-ss/split/tag_Atividade.json`

Corrigir os nomes no código conforme a spec, caso divirjam. Ajustar também a lista de `populate` para os relacionamentos que a spec realmente expõe.

- [ ] **Step 5: Verificar**

Run: `.venv/Scripts/python.exe -m py_compile app.py modules/atividade.py`
Expected: sem saída.

Run: `.venv/Scripts/python.exe -m streamlit run app.py`
Verificar manualmente, com login SUPP ativo: selecionar uma linha da amostra, confirmar que o painel mostra o processo, as tarefas e as atividades; que as atividades do usuário auditado aparecem destacadas; que o botão 🔄 recarrega; que salvar Conforme/Não Conforme persiste ao trocar de linha e voltar. Verificar também o comportamento **sem** login: o painel deve mostrar só os dados da planilha e o convite a autenticar, sem erro. Encerrar com Ctrl+C.

- [ ] **Step 6: Commit**

```bash
rtk git add app.py modules/atividade.py
rtk git commit -m "feat: painel de conferência com drill-down até as atividades"
```

---

### Task 6: Relatório Word

**Files:**
- Modify: `modules/report.py` (nova função `gerar_relatorio_detalhamento`, depois de `gerar_relatorio_distribuicao`)
- Modify: `app.py` — nova função `_render_relatorio_detalhamento()` e roteamento em `render_relatorio()` (linha 2064)

**Interfaces:**
- Consumes: `DetalhamentoData`, `COL_DET_*`, `stats_df`, e os helpers privados de `modules/report.py` (`_titulo`, `_subtitulo`, `_heading`, `_tabela_2col`, `_fmt_date`).
- Produces: `gerar_relatorio_detalhamento(det_data, df_detalhamento, tipo_controle, tamanho_amostra, responsavel, data_auditoria) -> bytes`.

- [ ] **Step 1: Escrever `gerar_relatorio_detalhamento`**

Em `modules/report.py`, depois de `gerar_relatorio_distribuicao`. Ler essa função inteira antes de escrever e espelhar a estrutura de seções, os helpers e o estilo das tabelas.

```python
def gerar_relatorio_detalhamento(
    det_data: "DetalhamentoData",
    df_detalhamento: pd.DataFrame | None,
    tipo_controle: str | None,
    tamanho_amostra: int | None,
    responsavel: str,
    data_auditoria: date,
) -> bytes:
    """
    Gera o relatório de auditoria de atividades por NUP e retorna bytes do .docx.
    """
    from modules.excel_loader import (
        COL_DET_ATIVIDADES, COL_DET_NUP, COL_DET_RESPONSAVEL, COL_DET_USUARIO,
    )
    from modules.state import COL_ACAO, COL_CONFORMIDADE, COL_MOTIVO, stats_df

    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.9)
        section.bottom_margin = Inches(0.9)
        section.left_margin = Inches(1.2)
        section.right_margin = Inches(1.2)

    _titulo(doc, "RELATÓRIO DE AUDITORIA")
    _subtitulo(doc, "Detalhamento Individual — Conformidade das Atividades Lançadas")
    _subtitulo(doc, "Procuradoria-Geral Federal / Advocacia-Geral da União")
    doc.add_paragraph()

    # 1. Identificação
    _heading(doc, "1. IDENTIFICAÇÃO")
    _tabela_2col(doc, [
        ("Período Auditado", det_data.meses or "N/D"),
        ("Data de Emissão do Relatório", _fmt_date(data_auditoria)),
        ("Responsável pela Auditoria", responsavel or "Não informado"),
        ("Sistema Auditado", "Power BI — Detalhamento Individual PGF"),
        ("Arquivo Analisado", det_data.nome_arquivo),
        ("Usuário Auditado", det_data.usuario or "N/D"),
        ("Unidade", det_data.unidade or "N/D"),
        ("Região", det_data.regiao or "N/D"),
        ("Base Normativa", "Portaria PGF/AGU n. 541/2025 — Manual de Gerenciamento Estratégico"),
    ])
    doc.add_paragraph()

    # 2. Metodologia
    _heading(doc, "2. METODOLOGIA")
    detalhado = tipo_controle == "detalhado"
    doc.add_paragraph(
        "A população auditada é composta pelos NUPs em que o usuário registrou "
        f"atividades no período: {det_data.total_nups} processos, totalizando "
        f"{det_data.total_atividades} atividades lançadas. Como o relatório de origem "
        "não identifica individualmente cada atividade, a unidade de análise adotada "
        "foi o NUP: em cada processo selecionado verificou-se, no Super Sapiens, se as "
        "atividades registradas correspondem ao efetivamente praticado."
    )
    if detalhado:
        doc.add_paragraph(
            f"Aplicou-se controle detalhado, com amostragem aleatória simples de "
            f"{tamanho_amostra} NUPs, para nível de confiança de 95% e margem de erro "
            "de ±5%, conforme o Anexo III do Manual."
        )
    else:
        doc.add_paragraph(
            "Aplicou-se controle simplificado, com verificação da totalidade dos NUPs "
            "da população ou de seleção manual do auditor."
        )
    doc.add_paragraph()

    # 3. Resultados
    _heading(doc, "3. RESULTADOS")
    s = stats_df(df_detalhamento)
    _tabela_2col(doc, [
        ("NUPs na população", str(det_data.total_nups)),
        ("NUPs examinados", str(s["auditadas"])),
        ("Conformes", f"{s['conformes']} ({s['pct_conf']:.1f}%)"),
        ("Não conformes", f"{s['nao_conformes']} ({s['pct_nc']:.1f}%)"),
    ])
    doc.add_paragraph()

    # 4. Não conformidades
    _heading(doc, "4. NÃO CONFORMIDADES IDENTIFICADAS")
    if df_detalhamento is None or s["nao_conformes"] == 0:
        doc.add_paragraph("Não foram identificadas não conformidades na amostra examinada.")
    else:
        nc = df_detalhamento[df_detalhamento[COL_CONFORMIDADE] == "Não Conforme"]
        tabela = doc.add_table(rows=1, cols=5)
        tabela.style = "Table Grid"
        cabecalho = ["NUP", "Responsável", "Atividades", "Motivo", "Ação Corretiva"]
        for celula, texto in zip(tabela.rows[0].cells, cabecalho):
            celula.text = texto
        for _, linha in nc.iterrows():
            cells = tabela.add_row().cells
            cells[0].text = str(linha.get(COL_DET_NUP, ""))
            cells[1].text = str(linha.get(COL_DET_RESPONSAVEL, ""))
            cells[2].text = str(linha.get(COL_DET_ATIVIDADES, ""))
            cells[3].text = str(linha.get(COL_MOTIVO, ""))
            cells[4].text = str(linha.get(COL_ACAO, ""))

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()
```

Conferir como `gerar_relatorio_distribuicao` termina (nome do buffer, `return`) e usar exatamente o mesmo padrão; conferir também se o negrito do cabeçalho da tabela é aplicado por um helper próprio no arquivo e, se for, usá-lo.

- [ ] **Step 2: Escrever `_render_relatorio_detalhamento` em `app.py`**

Ler `_render_relatorio_distribuicao` (linhas 1974–2062) por inteiro e escrever a versão análoga, inserida logo depois dela:

```python
def _render_relatorio_detalhamento() -> None:
    """Relatório de auditoria para o modo Power BI — Detalhamento Individual."""
    from modules.report import gerar_relatorio_detalhamento

    det_data = get_det_data()
    if det_data is None:
        st.warning("Nenhum arquivo carregado. Volte à página de Importação.")
        return

    st.title("📄 Relatório de Auditoria — Atividades por NUP")

    df_det = get_df_detalhamento()
    s_det = stats_df(df_det)

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
    c1.metric("NUPs na população", det_data.total_nups)
    c2.metric("Na Amostra / Auditados", f"{s_det['auditadas']}/{s_det['total']}")
    c3.metric("Conformes", s_det["conformes"],
              delta=f"{s_det['pct_conf']:.1f}%" if s_det["auditadas"] > 0 else None)
    c4.metric("Não Conformes", s_det["nao_conformes"],
              delta=f"{s_det['pct_nc']:.1f}%" if s_det["auditadas"] > 0 else None,
              delta_color="inverse")

    if s_det["nao_conformes"] > 0:
        st.divider()
        st.subheader(f"⚠️ Não Conformidades Identificadas ({s_det['nao_conformes']})")
        if df_det is not None:
            nc_df = df_det[df_det[COL_CONFORMIDADE] == "Não Conforme"][
                [COL_DET_NUP, COL_DET_RESPONSAVEL, COL_DET_ATIVIDADES, COL_MOTIVO, COL_ACAO]
            ]
            st.dataframe(nc_df, hide_index=True, use_container_width=True)

    st.divider()
    if st.button("📝 Gerar Relatório (.docx)", type="primary"):
        with st.spinner("Gerando relatório..."):
            st.session_state["relatorio_gerado"] = gerar_relatorio_detalhamento(
                det_data=det_data,
                df_detalhamento=df_det,
                tipo_controle=st.session_state.get("tipo_controle"),
                tamanho_amostra=st.session_state.get("tamanho_amostra"),
                responsavel=st.session_state.get("responsavel", ""),
                data_auditoria=st.session_state.get("data_auditoria", date_type.today()),
            )

    if st.session_state.get("relatorio_gerado"):
        st.download_button(
            "⬇️ Baixar Relatório",
            data=st.session_state["relatorio_gerado"],
            file_name=f"relatorio_detalhamento_{date_type.today():%Y%m%d}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
```

Alinhar o bloco de geração/download ao que `_render_relatorio_distribuicao` já faz (nome do arquivo, rótulos dos botões).

- [ ] **Step 3: Rotear em `render_relatorio`**

Em `render_relatorio()` (linha 2064), logo após o bloco de `supp_distribuicao`:

```python
    if tipo_rel == "detalhamento_individual":
        _render_relatorio_detalhamento()
        return
```

- [ ] **Step 4: Verificar**

Run: `.venv/Scripts/python.exe -m py_compile app.py modules/report.py`
Expected: sem saída.

Run: `.venv/Scripts/python.exe -m streamlit run app.py`
Verificar manualmente: percorrer o fluxo completo — importar a planilha real, controle detalhado, julgar ao menos um NUP como Não Conforme com motivo e ação, gerar o `.docx` e abri-lo. Conferir no documento: identificação com usuário/unidade/região, metodologia citando 12.731 NUPs e a amostra de 373, resultados e a tabela de não conformidades. Encerrar com Ctrl+C.

- [ ] **Step 5: Rodar a suíte de testes**

Run: `.venv/Scripts/python.exe -m pytest tests/ -v`
Expected: 10 passed

- [ ] **Step 6: Commit**

```bash
rtk git add app.py modules/report.py
rtk git commit -m "feat: relatório Word da auditoria de atividades por NUP"
```

---

### Task 7: Documentação

**Files:**
- Modify: `CLAUDE.md`
- Modify: `README.md`

- [ ] **Step 1: Atualizar `CLAUDE.md`**

Na seção "Arquitetura", ajustar a descrição de `excel_loader.py` para mencionar os três formatos. Acrescentar, depois da seção "v2 — Integração com Super Sapiens", uma seção curta:

```markdown
## Entradas de dados

O app aceita três formatos de planilha, detectados automaticamente por
`modules/excel_loader.detect_file_type()`:

1. **Conecta+ Triagem Avançada** — abas "Todas as Tarefas", "Tarefas Triadas",
   "Tarefas Não Triadas". Unidade auditada: a tarefa.
2. **Super Sapiens — Distribuição de Tarefas** — cabeçalho iniciado em `Id`.
   Unidade auditada: a distribuição.
3. **Power BI — Detalhamento Individual PGF** — cabeçalho iniciado em `NUP`,
   com bloco "Filtros aplicados:" na primeira célula. Unidade auditada: o NUP,
   porque o relatório de origem não identifica cada atividade individualmente.
```

Na seção "Notas", substituir "Sem testes automatizados (candidato para trabalho futuro)" por: "Testes automatizados apenas para o leitor de Detalhamento Individual (`tests/`, `pytest -r requirements-dev.txt`); os demais fluxos ainda não têm cobertura."

- [ ] **Step 2: Atualizar `README.md`**

Acrescentar o terceiro formato onde os dois atuais estiverem descritos, seguindo o estilo do arquivo. Ler o README antes de editar.

- [ ] **Step 3: Commit**

```bash
rtk git add CLAUDE.md README.md
rtk git commit -m "docs: documenta a entrada Detalhamento Individual"
```

---

## Self-Review

**Cobertura da spec:**

| Seção da spec | Task |
|---|---|
| 1. Leitor (dataclass, parser de filtros, agregação por NUP) | 1 |
| 1. `detect_file_type` estendido | 2 |
| 2. Estado e navegação | 2, 3 |
| 3. Página de importação | 3 |
| 4. Página de auditoria (amostragem, tabela) | 4 |
| 4. Painel de conferência com drill-down; `AtividadeClient.from_auth` | 5 |
| 5. Relatório | 6 |
| Fora de escopo (merge, julgamento por atividade, motivos pré-definidos) | não implementado, por desenho |

**Consistência de nomes:** `COL_DET_NUP`/`COL_DET_RESPONSAVEL`/`COL_DET_USUARIO`/`COL_DET_ATIVIDADES`, `DetalhamentoData`, `load_detalhamento_file`, `get_det_data`, `get_df_detalhamento`, `render_auditoria_detalhamento`, `_render_det_row_editor`, `_render_importacao_detalhamento`, `_render_relatorio_detalhamento`, `gerar_relatorio_detalhamento`, chaves `det_data` / `df_audit_detalhamento` / `auditoria_detalhamento_concluida`, tipo `"detalhamento_individual"`, página `"detalhamento"` — usados de forma idêntica em todas as tasks.

**Ponto de atenção:** os nomes de campo da API de Atividade (`especieAtividade`, `dataHoraConclusao`, `observacao`) são expectativa por convenção; a Task 5, Step 4 os confere contra `spec-ss/split/tag_Atividade.json` antes de dar a task por concluída.
