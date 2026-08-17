# Auditoria Conecta+ (v2 — Integração Super Sapiens)

Aplicação web desenvolvida em **Streamlit** para auditoria de triagem de tarefas do sistema **Conecta+ Automação** (SuperSapiens/SUPP), conforme a Portaria PGF/AGU n. 541/2025 — Manual de Gerenciamento Estratégico de Contencioso.

## Funcionalidades

### v1 (herdadas)

- **Importação** de planilhas Excel geradas pelo módulo de Triagem Avançada do Conecta+, com suporte a múltiplos arquivos (consolidação automática)
- **Auditoria de Tarefas Triadas** com dois modos de controle:
  - *Controle Simplificado* — verificação manual de todas as tarefas triadas
  - *Controle Detalhado* — amostragem estatística com nível de confiança de 95% e margem de erro de ±5%
- **Auditoria de Tarefas Não Triadas** com seleção total ou manual
- **Registro de conformidade** por tarefa (Conforme / Não Conforme / Não Avaliado), com campos de motivo e ação corretiva
- **Relatório em Word (.docx)** com resumo executivo, gráficos e lista de não conformidades

### v2 (implementadas)

> A importação de dados permanece via planilhas Excel. A integração com o SUPP é exclusivamente para **conferência** dos dados durante a auditoria.

- **Login no SUPP** — autenticação integrada ao Super Sapiens via JWT (login na Rede AGU)
- **Painel de conferência** — ao selecionar uma tarefa, busca automaticamente os metadados do processo via API: NUP formatado, número CNJ, classe nacional e entidade representada
- **Link direto ao SuperSapiens** — botão para abrir o processo no Super Sapiens em nova aba, sem necessidade de copiar NUP
- **Design system azul/cinza** — interface reformulada com paleta de cores frias (azul marinho `#1A3A6A` + cinza `#f7f9fc`), cards estruturados e uso restrito do vermelho a ações destrutivas

### Auditoria de Atividades por NUP (Detalhamento Individual PGF)

Terceiro formato de entrada aceito pelo app, além da planilha do Conecta+ e do relatório de Distribuição de Tarefas do Super Sapiens. O tipo é reconhecido automaticamente na importação.

- **Importação de um arquivo por auditoria** — a planilha corresponde a um usuário; não há consolidação de múltiplos arquivos, ao contrário do fluxo do Conecta+
- **Leitura dos filtros da extração** — usuário, unidade, região e meses são lidos do bloco "Filtros aplicados:" da planilha e exibidos antes de iniciar a auditoria
- **Amostragem pelo NUP** — a unidade amostral é o NUP, não a atividade, porque o relatório de origem não identifica cada atividade individualmente. As linhas da planilha são agregadas em uma por NUP (atividades somadas, responsáveis distintos concatenados)
- **Controle Simplificado ou Detalhado** — o Detalhado usa a mesma amostragem estatística dos demais fluxos (95% de confiança, ±5%)
- **Painel de conferência com drill-down** — para cada NUP sorteado, o app localiza o processo no Super Sapiens e percorre NUP → processo → tarefas → atividades, mostrando quantas atividades foram lançadas e quantas são do usuário auditado, para comparação com a quantidade informada na planilha
- **Registro de conformidade por NUP** — Conforme ou Não Conforme, com motivo e ação corretiva em texto livre
- **Relatório em Word (.docx)** — Identificação, Metodologia, Resultados, Não Conformidades Identificadas e Conclusão

## Estrutura do Projeto

```
audit_conecta/
├── app.py                        # Aplicação principal Streamlit
├── CLAUDE.md                     # Guia do projeto para Claude Code
├── requirements.txt              # Dependências Python
├── modules/
│   ├── config.py                 # Configuração (SUPP_BASE_URL)
│   ├── auth.py                   # Autenticação JWT (login, refresh, TOTP)
│   ├── processo.py               # Processos (CRUD, timeline, NUP)
│   ├── tarefa.py                 # Tarefas (listagem, filtros, workflow)
│   ├── atividade.py              # Atividades (administrativas, judiciais, consultivas)
│   ├── interessado.py            # Interessados do processo
│   ├── documento.py              # Documentos e componentes digitais
│   ├── componente_digital.py     # Arquivos binários (upload/download)
│   ├── etiqueta.py               # Vinculação de etiquetas
│   ├── catalogo_etiqueta.py      # Catálogo de etiquetas por setor
│   ├── setor.py                  # Setores/departamentos
│   ├── lotacao.py                # Lotações de usuários
│   ├── area_trabalho.py          # Áreas de trabalho pessoais
│   ├── folder.py                 # Pastas do usuário
│   ├── bookmark.py               # Destaques/anotações em documentos
│   ├── formulario_ia.py          # Formulários IA e triagem
│   ├── tipo_documento.py         # Tipos de documento
│   ├── excel_loader.py           # Importação das planilhas (três formatos)
│   ├── sampling.py               # Amostragem estatística
│   ├── state.py                  # Gerenciamento de estado da sessão
│   └── report.py                 # Geração do relatório Word
├── planilhas/                    # Exemplos de planilhas para importação
└── spec-ss/
    └── split/                    # Especificação OpenAPI do SUPP (430 tags)
        ├── index.md              # Índice geral das tags e rotas
        ├── tag_*.json            # Spec por tag/entidade
        ├── meta.json             # Metadados de autenticação
        └── components.json       # Schemas reutilizáveis
```

## Requisitos

- Python 3.9+
- Dependências listadas em `requirements.txt`:
  - `streamlit >= 1.40.0`
  - `pandas >= 2.2.0`
  - `openpyxl >= 3.1.0`
  - `python-docx >= 1.1.0`
  - `matplotlib >= 3.8.0`
  - `httpx >= 0.27.0` (cliente HTTP para API do SUPP)
  - `python-dotenv >= 1.0.0` (variáveis de ambiente)

## Instalação

```bash
# Clone o repositório
git clone https://github.com/jusmendonca/audit_conecta.git
cd audit_conecta

# Crie e ative o ambiente virtual
python -m venv .venv
source .venv/bin/activate   # Linux/macOS
.venv\Scripts\activate      # Windows

# Instale as dependências
pip install -r requirements.txt
```

## Configuração

Crie um arquivo `.env` na raiz do projeto:

```env
SUPP_BASE_URL=https://supersapiensbackend.agu.gov.br
```

> O `.env` é necessário apenas para as funcionalidades v2 (conferência via API). A auditoria via planilhas funciona sem ele.

## Execução

```bash
streamlit run app.py
```

A aplicação estará disponível em `http://localhost:8501`.

## Formatos de Planilha de Entrada

O app reconhece o formato do arquivo automaticamente na importação.

### Conecta+ Triagem Avançada

O arquivo Excel deve conter três abas:

| Aba | Descrição |
|-----|-----------|
| `Todas as Tarefas` | Lista completa de tarefas do período |
| `Tarefas Triadas` | Tarefas que passaram pela triagem |
| `Tarefas Não Triadas` | Tarefas pendentes de triagem |

Colunas esperadas: `ID`, `Tarefa`, `NUP`, `Usuário`, datas de criação/conclusão, `Status`, `Configurações Encontradas`.

### Detalhamento Individual PGF (Power BI)

Relatório exportado do Power BI Report Server, em Página Inicial → PGF → PGF - Painéis Estratégicos → 2025 → Detalhamento Individual PGF 2025.

As primeiras linhas trazem o bloco "Filtros aplicados:" (usuário, unidade, região e meses da extração). O cabeçalho começa na primeira linha cuja coluna A seja `NUP`, com as colunas:

| Coluna | Descrição |
|--------|-----------|
| `NUP` | Número único de protocolo do processo |
| `Responsável` | Responsável pelo processo |
| `Usuário que realizou a atividade` | Usuário auditado |
| `Atividades` | Quantidade de atividades lançadas |

Um mesmo NUP pode aparecer em mais de uma linha; a importação agrega tudo em uma linha por NUP, somando as atividades e concatenando os responsáveis distintos. Cada auditoria usa um único arquivo, correspondente a um usuário.

Exemplos em `planilhas/2026/`: `detalhamento_individual_2025.xlsx` (12.731 NUPs, 14.036 atividades — amostra de 373) e `detalhamento_individual_2026.xlsx` (10.275 NUPs, 12.107 atividades — amostra de 371).

### Super Sapiens — Distribuição de Tarefas

Relatório de distribuição exportado do Super Sapiens, com as colunas `Id`, `NUP` e os campos de usuário/setor de origem e destino.

## Contexto Normativo

Esta ferramenta apoia o controle interno da triagem realizado pela Procuradoria-Geral Federal (PGF/AGU), conforme previsto na seção 5 da Portaria PGF/AGU n. 541/2025.
