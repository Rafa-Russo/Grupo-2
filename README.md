# Impacto da criminalidade no mercado de hospedagem em São Paulo

## Início rápido

Pré-requisitos: Python 3.12+ e [uv](https://docs.astral.sh/uv/).

```powershell
uv sync
uv run ingest-listings
```

O segundo comando lê as planilhas em `data/crime_data/` e gera `data/crimes.csv`.
O terceiro lê os CSVs em `data/listings_data/` e gera `data/listings.csv`.
Os arquivos grandes em `data/` não são versionados. Para executar o projeto a partir de um clone, coloque os dados recebidos pelo grupo nos caminhos esperados:

```text
data/
	crime_data/SPDadosCriminais_*.xlsx
	crimes.csv
	listings.csv
```

Para abrir a análise no VS Code, selecione o ambiente `.venv` como kernel e abra `exploration/Russo/EDA.ipynb`.

## Estrutura

```text
data/                 Dados locais, não versionados
exploration/*/        Pastas para cada um explorar
src/ingestion/        Código de ingestão reutilizável
```

## Colaboração

Não inclua CSVs, planilhas ou outros arquivos grandes nos commits.
