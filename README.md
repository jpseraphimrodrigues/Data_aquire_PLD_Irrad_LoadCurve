# Data Acquire: PLD, Irradiância e Curva de Carga

Repositório para aquisição, tratamento e visualização de séries horárias de:

- PLD da CCEE;
- irradiância e geração fotovoltaica do PVGIS;
- curva de carga do ONS.

## Estrutura

- `src/`: scripts de aquisição, tratamento e visualização;
- `data/`: dados baixados ou processados localmente (não versionados por padrão);
- `figures/`: gráficos gerados (não versionados por padrão).

## Fontes

Os scripts documentam as fontes oficiais utilizadas. Consulte os comentários e as URLs de cada arquivo antes de publicar os dados no GitHub, verificando licença e condições de redistribuição.

## Instalação com uv

```bash
uv sync
```

O projeto usa o [uv](https://docs.astral.sh/uv/) como gerenciador principal de ambiente e dependências. O arquivo `uv.lock`, quando gerado, deverá ser versionado para garantir instalações reproduzíveis.

Para executar um script usando o ambiente do projeto:

```bash
uv run python src/2026_PLD_anual.py
```

## Compatibilidade com pip

O `requirements.txt` é mantido para facilitar o compartilhamento com pessoas que não utilizam uv:

```bash
python -m pip install -r requirements.txt
```

Quando as dependências forem alteradas, mantenha `requirements.txt` sincronizado com o `pyproject.toml`. O arquivo `pyproject.toml` é a fonte principal da configuração do projeto.

Os scripts ainda preservam seus nomes e caminhos originais. A padronização dos caminhos de entrada e saída será feita em uma etapa posterior.
