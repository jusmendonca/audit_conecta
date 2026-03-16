# Auditoria Conecta+

Aplicação web desenvolvida em **Streamlit** para auditoria de triagem de tarefas do sistema **Conecta+ Automação** (SuperSapiens), conforme a Portaria PGF/AGU n. 541/2025 — Manual de Gerenciamento Estratégico de Contencioso.

## Funcionalidades

- **Importação** de planilhas Excel geradas pelo módulo de Triagem Avançada do Conecta+, com suporte a múltiplos arquivos (consolidação automática)
- **Auditoria de Tarefas Triadas** com dois modos de controle:
  - *Controle Simplificado* — verificação manual de todas as tarefas triadas
  - *Controle Detalhado* — amostragem estatística com nível de confiança de 95% e margem de erro de ±5%
- **Auditoria de Tarefas Não Triadas** com seleção total ou manual
- **Registro de conformidade** por tarefa (Conforme / Não Conforme / Não Avaliado), com campos de motivo e ação corretiva
- **Relatório em Word (.docx)** com resumo executivo, gráficos e lista de não conformidades

## Estrutura do Projeto

```
audit_conecta/
├── app.py              # Aplicação principal Streamlit
├── modules/
│   ├── excel_loader.py # Leitura e consolidação das planilhas Excel
│   ├── sampling.py     # Cálculo de amostragem estatística
│   ├── state.py        # Gerenciamento de estado e dados de auditoria
│   └── report.py       # Geração do relatório Word
├── planilhas/          # Exemplos de planilhas para importação
└── requirements.txt
```

## Requisitos

- Python 3.9+
- Dependências listadas em `requirements.txt`:
  - `streamlit >= 1.40.0`
  - `pandas >= 2.2.0`
  - `openpyxl >= 3.1.0`
  - `python-docx >= 1.1.0`
  - `matplotlib >= 3.8.0`

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

## Execução

```bash
streamlit run app.py
```

A aplicação estará disponível em `http://localhost:8501`.

## Formato da Planilha de Entrada

O arquivo Excel deve conter três abas:

| Aba | Descrição |
|-----|-----------|
| `Todas as Tarefas` | Lista completa de tarefas do período |
| `Tarefas Triadas` | Tarefas que passaram pela triagem |
| `Tarefas Não Triadas` | Tarefas pendentes de triagem |

Colunas esperadas: `ID`, `Tarefa`, `NUP`, `Usuário`, datas de criação/conclusão, `Status`, `Configurações Encontradas`.

## Contexto Normativo

Esta ferramenta apoia o controle interno da triagem realizado pela Procuradoria-Geral Federal (PGF/AGU), conforme previsto na seção 5 da Portaria PGF/AGU n. 541/2025.
