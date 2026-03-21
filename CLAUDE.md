# Auditoria Conecta+ — Guia do Projeto

## Identidade

App web **Streamlit** para auditoria de triagem de tarefas do sistema Conecta+ Automação (SUPP/SuperSapiens), conforme a Portaria PGF/AGU n. 541/2025.

- **Linguagem:** Python 3.9+
- **Framework:** Streamlit
- **HTTP client:** httpx (módulos API)
- **Configuração:** python-dotenv (`.env` → `SUPP_BASE_URL`)

## Arquitetura

```
app.py                    ← UI Streamlit (fluxo de auditoria em 4 páginas)
modules/
  ├── config.py           ← BASE_URL via env SUPP_BASE_URL
  ├── auth.py             ← Autenticação JWT (login, refresh, TOTP 2FA)
  ├── processo.py         ← Processos (CRUD, timeline, NUP, juntadas)
  ├── tarefa.py           ← Tarefas (listagem, filtros, workflow)
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
  ├── excel_loader.py     ← Importação e consolidação de planilhas Excel
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

## Comandos

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Notas

- Sem testes automatizados (candidato para trabalho futuro)
- `.env` necessário com `SUPP_BASE_URL` para módulos API
- Docstrings usam `hermes/` como nome do pacote (artefato histórico; o pacote real é `modules/`)
