# VR Automation App (Vale-Refeição)

Este repositório contém uma aplicação em Python com interface Streamlit para orquestrar o cálculo mensal de VR (Vale‑Refeição) a partir de planilhas Excel, gerando um arquivo final padronizado com totais e formatação. Abaixo você encontra instruções em Português e, ao final, a versão em Inglês.

---

## 🇧🇷 Português

### Visão geral

- Front-end: página única em Streamlit para upload das bases, seleção de datas (início/fim), execução com logs em tempo real e download dos resultados.
- Back-end: scripts Python que consolidam bases, aplicam regras de negócio e geram a planilha final do VR com totais e formatação.
- Logs: a execução grava auditoria em `run.log` e exibe o progresso no app.

Arquivos principais:

- `app.py`: UI Streamlit (tema claro, logo, botões suaves, logs com spinner).
- `consolidar_bases.py`: consolida as bases de entrada e chama o agente de VR.
- `vr_agent.py`: ferramentas de leitura/gravação Excel e (opcional) agente LangChain+Gemini para orquestrar cálculos.
- Pasta `xlsx/`: arquivos de entrada e saídas (inclui `BaseConsolidada.xlsx` e o arquivo final `VR_MENSAL_MM.YYYY_FINAL.xlsx`).

### Tecnologias

- Python 3.11+
- Streamlit (UI)
- pandas, openpyxl (manipulação de dados/Excel)
- LangChain + Google Gemini (para orquestração com LLM)
- python-dotenv, holidays

### Regras de negócio aplicadas (VR)

- Competência: segue o mês/ano da “Data de Fim” selecionada na UI.
- Dias trabalhados: baseia-se nos dias úteis por sindicato fornecidos nas planilhas, com descontos por férias/ausências quando presentes.
- Regra de desligamento: se “DATA DEMISSÃO” for menor ou igual ao dia 15 do mês de competência, o VR do colaborador é zerado no mês (observação registrada).
- Cálculos de valores:
  - TOTAL_Prestador = Dias trabalhados × Valor Diário VR
  - Custo empresa = 80% do total
  - Desconto profissional = 20% do total
- Layout e formatação do Excel final:
  - Linha 1: rótulos acima dos totais (negrito + fundo claro):
    - Coluna TOTAL_Prestador: “Total_Custo”
    - Coluna Custo empresa: “Total_Custo_Empresa”
    - Coluna Desconto profissional: “Total_Custo_Funcionario”
  - Linha 2: os totais (soma da coluna)
  - Linha 3: cabeçalhos
  - Linha 4+: dados
  - Formatação monetária BRL aplicada a “Valor Diario VR”, “TOTAL_Prestador”, “Custo empresa”, “Desconto profissional”.
- Nomenclatura dinâmica do arquivo final: `VR_MENSAL_MM.YYYY_FINAL.xlsx` (por ex.: `VR_MENSAL_05.2025_FINAL.xlsx`).

#### Resumo das regras de negócio (VR)

- Conferência de datas de início e fim do contrato no mês.
- Exclusão de colaboradores em férias (parcial ou integral conforme regra do sindicato).
- Ajustes para datas “quebradas” (ex.: admissões no meio do mês e desligamentos).
- Cálculo do número exato de dias a serem comprados para cada pessoa.
- Geração de um layout de compra a ser enviado ao fornecedor.
- Considerar as regras vigentes decorrentes dos acordos coletivos de cada sindicato.
- Base única consolidada: reunir e consolidar informações de bases separadas em uma única base final para:
  - Ativos
  - Férias
  - Desligados
  - Base cadastral (admitidos do mês)
  - Base sindicato x valor
  - Dias úteis por colaborador
- Tratamento de exclusões: remover da base final todos os profissionais com:
  - cargos de diretores, estagiários e aprendizes;
  - afastamentos em geral (ex.: licença maternidade);
  - atuação no exterior.
  - Observação: guiar-se pela matrícula nas planilhas.
- Validar e corrigir:
  - datas inconsistentes ou “quebradas”;
  - campos faltantes;
  - férias mal preenchidas;
  - aplicação correta de feriados estaduais e municipais.
- Cálculo automatizado do benefício (com base na planilha):
  - quantidade de dias úteis por colaborador (considerando dias úteis de cada sindicato, férias, afastamentos e data de desligamento);
  - regra de desligamento: se houver comunicado “OK” até dia 15, não considerar para pagamento; se informado após dia 15, compra proporcional. Verificar pela matrícula a elegibilidade ao benefício (vide regras de exclusão);
  - valor total de Vale Refeição (VR) por colaborador, de acordo com o valor vigente do sindicato ao qual o profissional está vinculado, garantindo cálculo correto e vigente.

### Estrutura de pastas e arquivos esperados

- Coloque as planilhas de entrada na pasta `xlsx/` com os nomes abaixo:
  - `ATIVOS.xlsx`
  - `ADMISSÃO ABRIL.xlsx`
  - `FÉRIAS.xlsx`
  - `DESLIGADOS.xlsx`
  - `Base sindicato x valor.xlsx`
  - `Base dias uteis.xlsx` (pode ter uma linha extra de cabeçalho; o script trata)
  - (opcionais para filtros) `AFASTAMENTOS.xlsx`, `EXTERIOR.xlsx`
  - (modelo) `VR MENSAL 05.2025.xlsx` (usado para referência de layout)
- Saídas:
  - `xlsx/BaseConsolidada.xlsx`
  - `xlsx/VR_MENSAL_MM.YYYY_FINAL.xlsx`

### Configuração do ambiente

Pré-requisitos:

- Python 3.11+
- Pip recente

Passos (Windows, cmd.exe):

1. Criar e ativar um ambiente virtual

```
python -m venv .venv
.\.venv\Scripts\activate
```

2. Instalar dependências

```
pip install -r requirements.txt
```

3. Configurar a chave do Google Gemini (para usar o agente LLM)

- Opção A: variável de ambiente `GEMINI_API_KEY` configurada no sistema/sessão;
- Opção B: arquivo `.env` na raiz contendo `GEMINI_API_KEY=...`;
- Opção C: arquivo `gemini_api_key.txt` na raiz contendo a chave ou a linha `GEMINI_API_KEY=...`.

Observação Streamlit Cloud: Se você publicar esta aplicação no Streamlit Community Cloud, defina `GEMINI_API_KEY` em `st.secrets` (App settings > Edit secrets). O código prioriza `st.secrets` e só depois verifica variáveis de ambiente e arquivos locais.

Observação: Sem a chave, a UI e consolidação funcionam; apenas a etapa do agente LLM não executará.

Dicas (cmd.exe):

- Sessão atual: `set GEMINI_API_KEY=SUACHAVE`
- Persistente (usuário): `setx GEMINI_API_KEY "SUACHAVE"`

### Como executar

Opção 1 — UI (recomendado):

```
streamlit run app.py
```

- Abra o link mostrado (ex.: http://localhost:8501)
- Faça upload das planilhas em `xlsx/` (ou use a UI para salvar arquivos), selecione as datas e clique em “Executar Cálculo”.
- Acompanhe os logs e baixe as saídas geradas.

Opção 2 — Linha de comando (end‑to‑end):

```
python consolidar_bases.py --inicio YYYY-MM-DD --fim YYYY-MM-DD
```

- Gera `BaseConsolidada.xlsx` e chama o `vr_agent` para criar o arquivo final dinâmico.

Execução direta do agente (avançado):

```
python vr_agent.py
```

- Requer `GEMINI_API_KEY` válido e os arquivos no lugar.

Tarefas do VS Code (alternativo):

- Run consolidar_bases
- Run VR Agent / Run VR Agent (populate/blank+total/verify rule)
- Run end-to-end (consolidar->vr_agent)

### Logs e auditoria

- Arquivo `run.log` acumula auditoria e mensagens do agente.
- A UI faz streaming tanto do `stdout` do processo quanto do `run.log` para facilitar o acompanhamento.

### Solução de problemas

- “Permission denied” ao salvar Excel: feche a planilha no Excel e reexecute; o código tem fallback para salvar com nome alternativo em caso de lock.
- Erro de chave Gemini ausente: configure `GEMINI_API_KEY` (veja Configuração do ambiente).
- Arquivos ausentes/nome incorreto: confirme os nomes exatos listados acima em `xlsx/`.
- Formatação não aparece: abra o arquivo no Excel (não no visualizador online simplificado) para ver estilos e moeda BRL.

### Desenvolvimento e testes

- Teste automatizado do layout do Excel: `test_excel_layout.py` gera arquivos de teste e valida rótulos na linha 1 e totais na linha 2 para ambos os caminhos de escrita (pandas e tabela).

---

## 🇺🇸 English

### Overview

- Front-end: a single-page Streamlit app to upload input spreadsheets, pick start/end dates, run the pipeline with live logs, and download outputs.
- Back-end: Python scripts that consolidate sources, apply business rules, and generate the final VR workbook with totals and formatting.
- Logging: execution appends to `run.log` and shows progress in the UI.

Key files:

- `app.py`: Streamlit UI (light theme, logo, soft buttons, spinner, readable logs).
- `consolidar_bases.py`: consolidates inputs and triggers the VR agent.
- `vr_agent.py`: Excel I/O tools and (optional) LangChain+Gemini agent leveraging Pandas.
- `xlsx/`: input and output spreadsheets, including `BaseConsolidada.xlsx` and the final `VR_MENSAL_MM.YYYY_FINAL.xlsx`.

### Technologies

- Python 3.11+
- Streamlit
- pandas, openpyxl
- LangChain + Google Gemini (optional)
- python-dotenv (optional), holidays (optional)

### Applied business rules (VR)

- Competência (month/year) follows the selected “End Date”.
- Worked days: derived from per‑union business days in the provided spreadsheets, with deductions for vacations/absences when available.
- Termination rule: if “DATA DEMISSÃO” (termination date) is on/before the 15th of the competence month, the collaborator’s VR is zeroed that month (with an audit note).
- Value calculations:
  - TOTAL_Prestador = Worked days × Daily VR value
  - Custo empresa = 80% of total
  - Desconto profissional = 20% of total
- Final Excel layout and formatting:
  - Row 1: labels above totals (bold + light background):
    - Column TOTAL_Prestador: “Total_Custo”
    - Column Custo empresa: “Total_Custo_Empresa”
    - Column Desconto profissional: “Total_Custo_Funcionario”
  - Row 2: totals (column sums)
  - Row 3: headers
  - Row 4+: data
  - BRL currency format applied to monetary columns.
- Dynamic final filename: `VR_MENSAL_MM.YYYY_FINAL.xlsx` (e.g., `VR_MENSAL_05.2025_FINAL.xlsx`).

### Folder structure and expected files

Place inputs in `xlsx/` with the exact names:

- `ATIVOS.xlsx`, `ADMISSÃO ABRIL.xlsx`, `FÉRIAS.xlsx`, `DESLIGADOS.xlsx`
- `Base sindicato x valor.xlsx`, `Base dias uteis.xlsx`
- Optional filters: `AFASTAMENTOS.xlsx`, `EXTERIOR.xlsx`
- Template: `VR MENSAL 05.2025.xlsx`

Outputs:

- `xlsx/BaseConsolidada.xlsx`
- `xlsx/VR_MENSAL_MM.YYYY_FINAL.xlsx`

### Environment setup

Prereqs:

- Python 3.11+
- Recent pip

Steps (Windows, cmd.exe):

1. Create and activate a virtual environment

```
python -m venv .venv
.\.venv\Scripts\activate
```

2. Install dependencies

```
pip install -r requirements.txt
```

3. Set Google Gemini API key (to enable the LLM agent)

- Option A: environment variable `GEMINI_API_KEY`
- Option B: `.env` file with `GEMINI_API_KEY=...`
- Option C: `gemini_api_key.txt` containing the key or `GEMINI_API_KEY=...`

Note: Without the key, UI and consolidation still work; only the LLM agent step is skipped.

### How to run

Option 1 — Streamlit UI (recommended):

```
streamlit run app.py
```

- Open the provided URL (e.g., http://localhost:8501), upload xlsx files, select dates, run, and download results.

Option 2 — Command line (end‑to‑end):

```
python consolidar_bases.py --inicio YYYY-MM-DD --fim YYYY-MM-DD
```

- Produces `BaseConsolidada.xlsx` and triggers `vr_agent` to write the dynamic final workbook.

Direct agent run (advanced):

```
python vr_agent.py
```

- Requires a valid `GEMINI_API_KEY` and inputs in place.

VS Code tasks (alternative):

- Run consolidar_bases
- Run VR Agent / Run VR Agent (populate/blank+total/verify rule)
- Run end-to-end (consolidar->vr_agent)

### Logs and auditing

- `run.log` stores the audit trail and agent messages; the UI streams both process stdout and `run.log`.

### Troubleshooting

- “Permission denied” saving Excel: close the workbook in Excel and re-run; code includes a fallback to a timestamped file if locked.
- Missing Gemini key: configure `GEMINI_API_KEY` (see setup).
- Missing/incorrect filenames: ensure exact names in `xlsx/`.
- Missing formatting: open the file in desktop Excel to see styles and BRL currency format.

### Development and tests

- Automated Excel layout test: `test_excel_layout.py` creates test workbooks and validates row‑1 labels and row‑2 totals for both writer paths.

---

Boa execução! / Happy processing!
