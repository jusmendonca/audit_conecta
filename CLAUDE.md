# Auditoria Conecta+ — Guia do Projeto

## Identidade

App web **Streamlit** para auditoria de triagem de tarefas do sistema Conecta+ Automação (SUPP/SuperSapiens), conforme a Portaria PGF/AGU n. 541/2025.

- **Linguagem:** Python 3.9+
- **Framework:** Streamlit
- **HTTP client:** httpx (módulos API)
- **Configuração:** python-dotenv (`.env` → `SUPP_BASE_URL`)

## Arquitetura

```
app.py                    ← UI Streamlit (páginas do fluxo variam por tipo de entrada)
modules/
  ├── config.py           ← BASE_URL via env SUPP_BASE_URL
  ├── auth.py             ← Autenticação JWT (login, refresh, TOTP 2FA)
  ├── processo.py         ← Processos (CRUD, timeline, NUP, juntadas)
  ├── tarefa.py           ← Tarefas (listagem, filtros, workflow)
  ├── atividade.py        ← Atividades lançadas nas tarefas
  ├── interessado.py      ← Interessados do processo
  ├── documento.py        ← Documentos + ComponenteDigital
  ├── componente_digital.py ← Arquivos binários (upload/download/assinatura)
  ├── etiqueta.py         ← Vinculação de etiquetas
  ├── catalogo_etiqueta.py ← Catálogo de etiquetas por setor
  ├── setor.py            ← Setores/departamentos
  ├── lotacao.py          ← Lotações de usuários
  ├── area_trabalho.py    ← Áreas de trabalho pessoais
  ├── folder.py           ← Pastas do usuário
  ├── bookmark.py         ← Destaques/anotações em documentos
  ├── formulario_ia.py    ← Formulários IA e triagem
  ├── tipo_documento.py   ← Tipos de documento
  ├── excel_loader.py     ← Importação das planilhas Excel (três formatos; ver "Entradas de dados")
  ├── sampling.py         ← Amostragem estatística (95% confiança, ±5%)
  ├── state.py            ← Estado da sessão Streamlit
  └── report.py           ← Geração de relatório Word (.docx)
```

## Convenções dos módulos API

Todos os módulos SUPP seguem o mesmo padrão:

- Docstring inicial: `hermes/<nome>.py` com "Endpoints cobertos:" listando cada método/rota
- Classe por entidade (ex: `ProcessoClient`, `DocumentoClient`)
- HTTP via `httpx` síncrono com Bearer token
- `from_auth(auth_client)` — class method para instanciar a partir de um `AuthClient`
- Métodos padrão: `listar()`, `listar_todos()`, `contar()`, `buscar()`, `criar()`, `atualizar()`, `atualizar_parcial()`, `deletar()`
- Filtros via `where` (dict JSON), `populate` (lista de relações), `order` (dict), `limit/offset`

## Referência da API (spec)

```
spec-ss/split/
  ├── index.md          ← Índice de 430 tags (consultar primeiro)
  ├── tag_*.json        ← Spec OpenAPI por tag/entidade
  ├── meta.json         ← Metadados de autenticação e servidor
  ├── components.json   ← Schemas reutilizáveis
  └── search_index.json ← Índice de busca textual
```

**Workflow para criar novo módulo:** consultar `index.md` → localizar a tag → ler `tag_<Nome>.json`.

## v2 — Integração com Super Sapiens

**Princípio:** a importação de dados continua via planilhas Excel. A integração com o SUPP é exclusivamente para **conferência** — o auditor clica no NUP e abre os detalhes no painel lateral.

### Roadmap

1. **Login** — Autenticação no SUPP dentro do app (usar `modules/auth.py`)
2. **Links NUP** — NUPs clicáveis nas tabelas de auditoria, abrindo detalhes
3. **Painel de conferência** — Painel lateral direito com metadados da tarefa, processo, documentos, etiquetas, movimentos e eventos
4. **Navegação completa** — Drill-down entre tarefa → processo → documentos → componentes digitais

## Entradas de dados

O app aceita três formatos de planilha, detectados por
`modules/excel_loader.detect_file_type()` na ordem abaixo:

1. **Conecta+ Triagem Avançada** (`"conecta_triagem"`) — reconhecido pelas abas
   "Todas as Tarefas", "Tarefas Triadas" e "Tarefas Não Triadas". Unidade
   auditada: a tarefa. É o único formato com consolidação de múltiplos arquivos
   (`merge_audit_data()`, deduplicando por `COL_TAREFA`).
2. **Power BI — Detalhamento Individual PGF** (`"detalhamento_individual"`) —
   reconhecido por uma linha, entre as 15 primeiras, cuja coluna A seja `NUP` e
   que contenha também a coluna `Atividades`. Uma planilha por auditoria.
3. **Super Sapiens — Distribuição de Tarefas** (`"supp_distribuicao"`) — retorno
   padrão quando nenhum dos anteriores casa. Unidade auditada: a distribuição.

O tipo detectado fica em `st.session_state["tipo_relatorio"]` e determina o
conjunto de páginas exibido por `_get_paginas()` em `app.py`:
`PAGINAS_TRIAGEM` (4 páginas), `PAGINAS_DISTRIBUICAO` e `PAGINAS_DETALHAMENTO`
(3 páginas cada).

### Detalhamento Individual PGF

Relatório extraído do Power BI Report Server em
Página Inicial → PGF → PGF - Painéis Estratégicos → 2025 →
Detalhamento Individual PGF 2025.

- As linhas acima do cabeçalho trazem o bloco "Filtros aplicados:", lido por
  `load_detalhamento_file()` para preencher `usuario`, `unidade`, `regiao` e
  `meses` de `DetalhamentoData` (chaves em `DET_FILTROS`). Esse bloco é usado
  apenas como metadado — não participa da detecção do tipo.
- Colunas obrigatórias: `COL_DET_NUP`, `COL_DET_RESPONSAVEL`, `COL_DET_USUARIO`
  ("Usuário que realizou a atividade") e `COL_DET_ATIVIDADES`.
- **A unidade amostral é o NUP, não a atividade** — o relatório de origem não
  expõe o id de cada atividade. As linhas são agregadas em uma por NUP (soma de
  atividades, responsáveis distintos concatenados) e o auditor examina cada NUP
  sorteado em busca de atividades lançadas de forma equivocada.
- A amostragem reaproveita `modules/sampling.py` (95% de confiança, ±5%).
- Julgamento livre por NUP: Conforme / Não Conforme, com motivo e ação corretiva
  em texto livre.
- Estado em `modules/state.py`: `det_data`, `df_audit_detalhamento`,
  `auditoria_detalhamento_concluida`, com os acessores `get_det_data()` e
  `get_df_detalhamento()`.
- Relatório Word por `modules.report.gerar_relatorio_detalhamento()`.
- Planilhas de exemplo em `planilhas/2026/`:
  `detalhamento_individual_2025.xlsx` (12.731 NUPs, 14.036 atividades, amostra
  373) e `detalhamento_individual_2026.xlsx` (10.275 NUPs, 12.107 atividades,
  amostra 371).

## Comandos

```bash
pip install -r requirements.txt
streamlit run app.py

# Testes (requer pip install -r requirements-dev.txt)
pytest tests/ -q
```

## Notas

- Testes automatizados cobrem apenas o leitor do Detalhamento Individual
  (`tests/test_excel_loader_detalhamento.py`); os demais fluxos ainda não têm cobertura
- `.env` necessário com `SUPP_BASE_URL` para módulos API
- Docstrings usam `hermes/` como nome do pacote (artefato histórico; o pacote real é `modules/`)
