# Detalhamento Individual PGF — nova entrada de dados

**Data:** 2026-08-13
**Status:** aprovado

## Objetivo

Adicionar um terceiro tipo de auditoria ao app, alimentado pela planilha
**Detalhamento Individual PGF 2025**, exportada do Power BI Report Server
(Página Inicial → PGF → PGF - Painéis Estratégicos → 2025 → Detalhamento
Individual PGF 2025).

A planilha lista, por usuário, os NUPs em que ele lançou atividades e quantas
atividades lançou em cada um. O objeto material da auditoria seria cada
atividade, mas a planilha não traz o identificador delas. Por isso a **unidade
amostral é o NUP**: sorteia-se um conjunto de NUPs e, em cada um, o auditor
examina as atividades lançadas em busca de registros equivocados.

## Modelo de dados da planilha

Arquivo de referência: `planilhas/2026/detalhamento_individual_2025.xlsx`
(12.740 linhas de dados, 12.731 NUPs distintos, 14.036 atividades).

| Posição | Conteúdo |
|---|---|
| Linha 0, coluna A | Bloco multilinha `Filtros aplicados:` — `Mês é …`, `unidades.regiao é …`, `unidades.nome é …`, `USUARIO é …` |
| Linha 1 | vazia |
| Linha 2 | Cabeçalho: `NUP`, `Responsável`, `Usuário que realizou a atividade`, `Atividades` |
| Linha 3+ | Dados |

Observações:

- 9 NUPs aparecem em mais de uma linha, com `Responsável` diferente.
- `Atividades` é a contagem de atividades daquele usuário naquele NUP.
- Uma auditoria trata de **um arquivo por vez** (um usuário). Não há
  consolidação de múltiplos arquivos.

## 1. Leitor (`modules/excel_loader.py`)

Novo dataclass:

```python
@dataclass
class DetalhamentoData:
    nome_arquivo: str
    usuario: str | None          # "USUARIO é …"
    unidade: str | None          # "unidades.nome é …"
    regiao: str | None           # "unidades.regiao é …"
    meses: str | None            # "Mês é …"
    filtros_raw: str | None      # bloco bruto, para exibição
    df: pd.DataFrame             # uma linha por NUP
    total_nups: int              # população
    total_atividades: int
    total_responsaveis: int
```

`load_detalhamento_file(uploaded_file, nome_arquivo="") -> DetalhamentoData`:

1. Lê a planilha com `header=None, dtype=str`.
2. Localiza a linha de cabeçalho pela primeira linha cuja coluna A seja `NUP`
   (mesma estratégia de `load_distribution_file`, que procura `Id`). Não fixar
   o índice 2.
3. Parseia o bloco de filtros das linhas anteriores ao cabeçalho, quebrando por
   `\n` e casando as chaves conhecidas de forma tolerante a acento e caixa
   (usar `_norm`).
4. Renomeia as colunas para os nomes canônicos via `_resolver`; exige
   `NUP`, `Responsável`, `Usuário que realizou a atividade`, `Atividades`.
   Ausência de qualquer uma → `ValueError` em português, no mesmo tom das
   mensagens existentes.
5. Converte `Atividades` para inteiro (`pd.to_numeric`, `errors="coerce"`,
   NaN → 0).
6. **Agrega por NUP**: soma `Atividades`, concatena responsáveis distintos com
   `; `, mantém o usuário. Resultado: uma linha por NUP, preservando o total de
   atividades.

Constantes de coluna seguindo o padrão do arquivo: `COL_DET_NUP`,
`COL_DET_RESPONSAVEL`, `COL_DET_USUARIO`, `COL_DET_ATIVIDADES`.

`detect_file_type()` passa a devolver também `"detalhamento_individual"`. Ordem
de teste: abas do Conecta+ → cabeçalho do Detalhamento (`NUP` na coluna A e
`Atividades` entre as colunas) → fallback `"supp_distribuicao"`.

Não há `merge_detalhamento_data` — um arquivo por auditoria.

## 2. Estado e navegação

`modules/state.py`:

- Chaves novas: `det_data`, `df_audit_detalhamento`,
  `auditoria_detalhamento_concluida`.
- Incluídas em `_DEFAULTS`, `_PERSIST_KEYS` e `reset_auditoria()`.
- Helper `get_det_data()`.
- `get_session_info()` reconhece o novo tipo (nome do arquivo, usuário,
  período).

`app.py`:

- O rádio de tipo na importação passa a ter três opções: Conecta+ Triagem,
  Distribuição SS e **Detalhamento Individual**.
- `_get_paginas()` devolve, para o novo tipo:
  `Importação → Auditoria de Atividades → Relatório`.

## 3. Página de importação

`_render_importacao_detalhamento()`:

- Upload de **um** arquivo `.xlsx`.
- Cartão com os filtros aplicados: usuário, unidade, região, meses.
- Métricas: **NUPs (população)**, atividades lançadas, responsáveis distintos.
- Botão que grava `det_data` no estado e avança para a auditoria.

## 4. Página de auditoria

`render_auditoria_detalhamento()`, espelhando
`render_auditoria_distribuicao()`:

- **Tipo de controle**: Simplificado (todos os NUPs) ou Detalhado (amostragem
  estatística 95% de confiança, ±5% — `calcular_amostra(12731)` ≈ 373), com a
  tabela de referência do Anexo III e a descrição da fórmula já existentes em
  `modules/sampling.py`.
- **Tabela (esquerda)**: NUP, Responsável, Usuário, Atividades, Conformidade,
  Motivo, Ação corretiva. Julgamento **livre** (Conforme / Não Conforme +
  campos de texto), idêntico aos demais fluxos.
- **Painel de conferência (direita)** — drill-down completo até as atividades:
  1. `ProcessoClient.buscar_por_nup(nup)` → metadados do processo.
  2. `TarefaClient.listar(where={"processo.id": <id>})` → tarefas do processo.
  3. `AtividadeClient.listar_por_tarefa(<tarefa_id>)` para cada tarefa →
     tipo de atividade, usuário, setor, data e observação.
  - Atividades do usuário auditado recebem destaque visual.
  - O resultado é cacheado em `session_state` por NUP, para não repetir
    chamadas ao alternar entre linhas.
  - Sem login SUPP ativo, o painel mostra apenas os dados da planilha e o
    convite a autenticar, como nos fluxos existentes.

Se `AtividadeClient` não possuir `from_auth()`, adicioná-lo seguindo a
convenção dos demais módulos.

### Custo do drill-down

São 2 + N chamadas por NUP (N = número de tarefas do processo). O carregamento
é síncrono, com spinner e cache. Otimizações (carregamento por tarefa sob
demanda) só serão feitas se a lentidão se confirmar no uso real.

## 5. Relatório

`modules/report.py` ganha a seção do novo tipo, espelhando a de distribuição:

- Identificação da auditoria e do arquivo.
- Filtros aplicados extraídos da planilha.
- Metodologia: população = NUPs com atividades lançadas pelo usuário no
  período; amostragem estatística quando o controle for detalhado.
- Resultados de conformidade (totais e percentuais).
- Tabela das não conformidades, com motivo e ação corretiva.

`render_relatorio()` em `app.py` roteia para `_render_relatorio_detalhamento()`.

## Fora de escopo

- Consolidação de múltiplas planilhas de Detalhamento Individual.
- Julgamento por atividade individual (o julgamento é por NUP).
- Motivos de não conformidade pré-definidos.
