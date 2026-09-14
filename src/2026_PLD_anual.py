# ============================================================
# DOWNLOAD E TRATAMENTO DO PLD HORÁRIO DE 2022
# Fonte: Dados Abertos da CCEE
#
# Objetivos:
# 1. Baixar todos os registros horários do PLD de 2022;
# 2. Selecionar o submercado Sudeste;
# 3. Construir corretamente a data a partir de:
#       MES_REFERENCIA + DIA
# 4. Interpretar corretamente a hora;
# 5. Verificar lacunas, duplicidades e valores ausentes;
# 6. Gerar Excel e CSV;
# 7. Produzir resumos anual, mensal, diário e horário.
#
# Bibliotecas necessárias:
#
# python -m pip install pandas openpyxl
# ============================================================

from pathlib import Path
import json
import time
import urllib.error
import urllib.parse
import urllib.request

import pandas as pd


# ============================================================
# 1. CONFIGURAÇÕES PRINCIPAIS
# ============================================================

ANO_ANALISE = 2022

# Resource ID oficial do PLD horário de 2022
RESOURCE_ID_CCEE = "723cf7e6-6c29-4da6-aa39-e4c8804baf65"

# Na base da CCEE, o nome pode aparecer como "SUDESTE".
# O código também reconhece automaticamente a sigla "SE".
SUBMERCADO_ANALISE = "SUDESTE"

# Quantidade de registros por página da API
LIMITE_REGISTROS_POR_REQUISICAO = 5000

# Pequeno intervalo entre as páginas
INTERVALO_ENTRE_REQUISICOES_SEGUNDOS = 0.2

# Tempo máximo por requisição
TIMEOUT_SEGUNDOS = 120

# Pasta para salvar os resultados
PASTA_SAIDA = Path("resultados_pld_ccee")

# Nome base dos arquivos
NOME_BASE_ARQUIVO = (
    f"PLD_Horario_CCEE_{ANO_ANALISE}_{SUBMERCADO_ANALISE}"
)

# Salvar também em CSV
SALVAR_CSV = True

# Incluir a base de todos os submercados no Excel
SALVAR_BASE_COMPLETA = True

# Quantidade de dias de maior variação a destacar
QUANTIDADE_DIAS_MAIOR_VARIACAO = 20


# ============================================================
# 2. MONTAGEM DA URL DA API
# ============================================================

def montar_url_api(
    resource_id: str,
    limite: int,
    offset: int,
) -> str:
    """
    Monta a URL da API CKAN da CCEE.
    """

    url_base = (
        "https://dadosabertos.ccee.org.br/"
        "api/3/action/datastore_search"
    )

    parametros = {
        "resource_id": resource_id,
        "limit": limite,
        "offset": offset,
    }

    return (
        url_base
        + "?"
        + urllib.parse.urlencode(parametros)
    )


# ============================================================
# 3. REALIZAÇÃO DE UMA REQUISIÇÃO
# ============================================================

def realizar_requisicao_api(
    url: str,
) -> dict:
    """
    Executa uma requisição HTTP e devolve o conteúdo JSON.
    """

    requisicao = urllib.request.Request(
        url,
        headers={
            "User-Agent": (
                "Mozilla/5.0 "
                "Pesquisa-Academica-PLD-CCEE"
            ),
            "Accept": "application/json",
        },
    )

    try:
        with urllib.request.urlopen(
            requisicao,
            timeout=TIMEOUT_SEGUNDOS,
        ) as resposta:

            conteudo = resposta.read().decode("utf-8")

    except urllib.error.HTTPError as erro:

        try:
            mensagem = erro.read().decode(
                "utf-8",
                errors="replace",
            )
        except Exception:
            mensagem = str(erro)

        raise RuntimeError(
            "\nErro HTTP ao acessar a API da CCEE.\n"
            f"Código: {erro.code}\n"
            f"Mensagem: {mensagem[:2000]}"
        ) from erro

    except urllib.error.URLError as erro:
        raise RuntimeError(
            "\nNão foi possível acessar a API da CCEE.\n"
            f"Motivo: {erro.reason}"
        ) from erro

    except TimeoutError as erro:
        raise RuntimeError(
            "A requisição à CCEE excedeu o tempo máximo."
        ) from erro

    try:
        dados_json = json.loads(conteudo)

    except json.JSONDecodeError as erro:
        raise RuntimeError(
            "A resposta da CCEE não pôde ser convertida para JSON."
        ) from erro

    if not dados_json.get("success", False):
        raise RuntimeError(
            "A API da CCEE indicou falha na consulta.\n"
            f"Resposta: {dados_json}"
        )

    if "result" not in dados_json:
        raise RuntimeError(
            "A resposta da API não contém o campo 'result'."
        )

    return dados_json


# ============================================================
# 4. DOWNLOAD PAGINADO
# ============================================================

def baixar_todos_registros_ccee(
    resource_id: str,
    limite_por_requisicao: int,
) -> pd.DataFrame:
    """
    Baixa todos os registros do recurso da CCEE.
    """

    todos_registros = []

    offset = 0
    total_registros = None
    numero_pagina = 1

    print("=" * 75)
    print("DOWNLOAD DO PLD HORÁRIO — CCEE")
    print("=" * 75)
    print(f"Ano:                         {ANO_ANALISE}")
    print(f"Resource ID:                 {resource_id}")
    print(
        f"Limite por requisição:       "
        f"{limite_por_requisicao}"
    )
    print()

    while True:

        url = montar_url_api(
            resource_id=resource_id,
            limite=limite_por_requisicao,
            offset=offset,
        )

        print(
            f"Baixando página {numero_pagina} "
            f"— offset {offset}..."
        )

        dados_json = realizar_requisicao_api(url)

        resultado = dados_json["result"]

        registros_pagina = resultado.get(
            "records",
            [],
        )

        if total_registros is None:
            total_registros = resultado.get(
                "total",
                0,
            )

            print(
                f"Total informado pela API: "
                f"{total_registros} registros"
            )
            print()

        if not registros_pagina:
            break

        todos_registros.extend(registros_pagina)

        quantidade_recebida = len(registros_pagina)

        print(
            f"Registros recebidos nesta página: "
            f"{quantidade_recebida}"
        )

        print(
            f"Registros acumulados: "
            f"{len(todos_registros)}"
        )
        print()

        offset += quantidade_recebida
        numero_pagina += 1

        if len(todos_registros) >= total_registros:
            break

        if quantidade_recebida < limite_por_requisicao:
            break

        time.sleep(
            INTERVALO_ENTRE_REQUISICOES_SEGUNDOS
        )

    if not todos_registros:
        raise RuntimeError(
            "Nenhum registro foi retornado pela API da CCEE."
        )

    df = pd.DataFrame(todos_registros)

    print("Download concluído.")
    print(
        f"Total efetivamente baixado: "
        f"{len(df)} registros"
    )
    print()

    return df


# ============================================================
# 5. PADRONIZAÇÃO DAS COLUNAS
# ============================================================

def padronizar_nomes_colunas(
    df: pd.DataFrame,
) -> pd.DataFrame:
    """
    Padroniza os nomes das colunas para letras maiúsculas.
    """

    df = df.copy()

    df.columns = [
        str(coluna)
        .strip()
        .upper()
        for coluna in df.columns
    ]

    colunas_obrigatorias = [
        "MES_REFERENCIA",
        "SUBMERCADO",
        "PERIODO_COMERCIALIZACAO",
        "DIA",
        "HORA",
        "PLD_HORA",
    ]

    colunas_ausentes = [
        coluna
        for coluna in colunas_obrigatorias
        if coluna not in df.columns
    ]

    if colunas_ausentes:
        raise RuntimeError(
            "A base da CCEE não contém todas as colunas esperadas.\n"
            f"Colunas ausentes: {colunas_ausentes}\n"
            f"Colunas encontradas: {df.columns.tolist()}"
        )

    return df


# ============================================================
# 6. CONVERSÃO DE NÚMEROS
# ============================================================

def converter_numero(
    serie: pd.Series,
) -> pd.Series:
    """
    Converte valores numéricos escritos com ponto ou vírgula.
    """

    serie_texto = (
        serie
        .astype(str)
        .str.strip()
        .str.replace("R$", "", regex=False)
        .str.replace(" ", "", regex=False)
    )

    mascara_com_virgula = (
        serie_texto.str.contains(
            ",",
            regex=False,
            na=False,
        )
    )

    serie_texto.loc[mascara_com_virgula] = (
        serie_texto.loc[mascara_com_virgula]
        .str.replace(".", "", regex=False)
        .str.replace(",", ".", regex=False)
    )

    return pd.to_numeric(
        serie_texto,
        errors="coerce",
    )


# ============================================================
# 7. NORMALIZAÇÃO DO SUBMERCADO
# ============================================================

def normalizar_submercado(
    valor: str,
) -> str:
    """
    Padroniza diferentes nomes possíveis dos submercados.
    """

    texto = (
        str(valor)
        .strip()
        .upper()
    )

    equivalencias = {
        "SE": "SUDESTE",
        "SUDESTE": "SUDESTE",
        "SUDESTE/CENTRO-OESTE": "SUDESTE",
        "SUDESTE CENTRO-OESTE": "SUDESTE",
        "SE/CO": "SUDESTE",

        "S": "SUL",
        "SUL": "SUL",

        "NE": "NORDESTE",
        "NORDESTE": "NORDESTE",

        "N": "NORTE",
        "NORTE": "NORTE",
    }

    return equivalencias.get(
        texto,
        texto,
    )


# ============================================================
# 8. TRATAMENTO DA BASE
# ============================================================

def tratar_base_ccee(
    df: pd.DataFrame,
) -> pd.DataFrame:
    """
    Trata os registros baixados da CCEE.

    A data é construída obrigatoriamente por:

        MES_REFERENCIA + DIA

    Exemplo:

        MES_REFERENCIA = 202201
        DIA = 1
        DATA = 2022-01-01
    """

    df = padronizar_nomes_colunas(df)

    # --------------------------------------------------------
    # Submercado
    # --------------------------------------------------------

    df["SUBMERCADO_ORIGINAL"] = (
        df["SUBMERCADO"]
        .astype(str)
        .str.strip()
        .str.upper()
    )

    df["SUBMERCADO"] = (
        df["SUBMERCADO_ORIGINAL"]
        .apply(normalizar_submercado)
    )

    # --------------------------------------------------------
    # Conversão numérica
    # --------------------------------------------------------

    df["HORA"] = converter_numero(
        df["HORA"]
    )

    df["PLD_HORA"] = converter_numero(
        df["PLD_HORA"]
    )

    df["PERIODO_COMERCIALIZACAO"] = converter_numero(
        df["PERIODO_COMERCIALIZACAO"]
    )

    quantidade_horas_invalidas = (
        df["HORA"].isna().sum()
    )

    quantidade_pld_invalidos = (
        df["PLD_HORA"].isna().sum()
    )

    if quantidade_horas_invalidas > 0:
        raise RuntimeError(
            f"Foram encontradas {quantidade_horas_invalidas} "
            "horas inválidas."
        )

    if quantidade_pld_invalidos > 0:
        raise RuntimeError(
            f"Foram encontrados {quantidade_pld_invalidos} "
            "valores inválidos de PLD."
        )

    # ========================================================
    # CONSTRUÇÃO CORRETA DA DATA
    # ========================================================

    # MES_REFERENCIA deve ser AAAAMM
    df["MES_REFERENCIA_TEXTO"] = (
        df["MES_REFERENCIA"]
        .astype(str)
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
        .str.replace("-", "", regex=False)
        .str.replace("/", "", regex=False)
        .str.zfill(6)
    )

    # DIA deve ser DD
    df["DIA_TEXTO"] = (
        df["DIA"]
        .astype(str)
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
        .str.zfill(2)
    )

    # Exemplo:
    # 202201 + 01 = 20220101
    df["DATA_TEXTO"] = (
        df["MES_REFERENCIA_TEXTO"]
        + df["DIA_TEXTO"]
    )

    df["data"] = pd.to_datetime(
        df["DATA_TEXTO"],
        format="%Y%m%d",
        errors="coerce",
    )

    quantidade_datas_invalidas = (
        df["data"].isna().sum()
    )

    if quantidade_datas_invalidas > 0:

        exemplos_invalidos = (
            df.loc[
                df["data"].isna(),
                [
                    "MES_REFERENCIA",
                    "DIA",
                    "MES_REFERENCIA_TEXTO",
                    "DIA_TEXTO",
                    "DATA_TEXTO",
                ],
            ]
            .drop_duplicates()
            .head(20)
        )

        raise RuntimeError(
            f"Foram encontradas {quantidade_datas_invalidas} "
            "datas inválidas.\n\n"
            "Exemplos:\n"
            f"{exemplos_invalidos.to_string(index=False)}"
        )

    print("=" * 75)
    print("CONSTRUÇÃO DAS DATAS")
    print("=" * 75)
    print(f"Primeira data: {df['data'].min()}")
    print(f"Última data:   {df['data'].max()}")
    print(f"Datas inválidas: {quantidade_datas_invalidas}")
    print()

    # ========================================================
    # INTERPRETAÇÃO DA COLUNA HORA
    # ========================================================

    hora_minima = df["HORA"].min()
    hora_maxima = df["HORA"].max()

    valores_hora = sorted(
        df["HORA"]
        .dropna()
        .unique()
        .tolist()
    )

    print("=" * 75)
    print("INTERPRETAÇÃO DA COLUNA HORA")
    print("=" * 75)
    print(f"Menor valor encontrado: {hora_minima}")
    print(f"Maior valor encontrado: {hora_maxima}")
    print(f"Valores encontrados:    {valores_hora}")
    print()

    # Caso 0–23
    if (
        hora_minima >= 0
        and hora_maxima <= 23
        and 0 in valores_hora
    ):

        print(
            "A coluna HORA foi interpretada como 0 a 23."
        )

        df["hora_inicio"] = df["HORA"]

    # Caso 1–24
    elif (
        hora_minima >= 1
        and hora_maxima <= 24
    ):

        print(
            "A coluna HORA foi interpretada como 1 a 24."
        )

        df["hora_inicio"] = (
            df["HORA"] - 1
        )

    else:

        raise RuntimeError(
            "A coluna HORA apresenta valores fora dos "
            "intervalos esperados.\n"
            f"Mínimo: {hora_minima}\n"
            f"Máximo: {hora_maxima}"
        )

    df["hora_inicio"] = (
        df["hora_inicio"]
        .round()
        .astype(int)
    )

    print()

    # ========================================================
    # CONSTRUÇÃO DA DATA E HORA
    # ========================================================

    df["data_hora"] = (
        df["data"]
        + pd.to_timedelta(
            df["hora_inicio"],
            unit="h",
        )
    )

    # --------------------------------------------------------
    # Colunas auxiliares
    # --------------------------------------------------------

    df["ano"] = df["data_hora"].dt.year
    df["mes"] = df["data_hora"].dt.month
    df["dia"] = df["data_hora"].dt.day
    df["hora"] = df["data_hora"].dt.hour
    df["dia_do_ano"] = (
        df["data_hora"].dt.dayofyear
    )

    df["dia_da_semana_numero"] = (
        df["data_hora"].dt.dayofweek
    )

    nomes_dias = {
        0: "Segunda-feira",
        1: "Terça-feira",
        2: "Quarta-feira",
        3: "Quinta-feira",
        4: "Sexta-feira",
        5: "Sábado",
        6: "Domingo",
    }

    df["dia_da_semana"] = (
        df["dia_da_semana_numero"]
        .map(nomes_dias)
    )

    df["chave_data_hora"] = (
        df["data_hora"]
        .dt.strftime("%Y-%m-%d %H:00:00")
    )

    # Nome explícito da unidade
    df["PLD_R$_MWh"] = df["PLD_HORA"]

    # --------------------------------------------------------
    # Ordenação
    # --------------------------------------------------------

    df = (
        df.sort_values(
            [
                "data_hora",
                "SUBMERCADO",
            ]
        )
        .reset_index(drop=True)
    )

    return df


# ============================================================
# 9. FILTRO DO ANO E SUBMERCADO
# ============================================================

def filtrar_ano_submercado(
    df: pd.DataFrame,
    ano: int,
    submercado: str,
) -> pd.DataFrame:
    """
    Filtra o ano e o submercado desejados.
    """

    submercado_normalizado = normalizar_submercado(
        submercado
    )

    submercados_disponiveis = sorted(
        df["SUBMERCADO"]
        .dropna()
        .unique()
        .tolist()
    )

    print("=" * 75)
    print("SUBMERCADOS ENCONTRADOS")
    print("=" * 75)
    print(submercados_disponiveis)
    print()

    if submercado_normalizado not in submercados_disponiveis:
        raise RuntimeError(
            f"O submercado '{submercado_normalizado}' "
            "não foi encontrado.\n"
            f"Disponíveis: {submercados_disponiveis}"
        )

    filtro = (
        (df["ano"] == ano)
        & (
            df["SUBMERCADO"]
            == submercado_normalizado
        )
    )

    df_filtrado = df.loc[filtro].copy()

    df_filtrado = (
        df_filtrado
        .sort_values("data_hora")
        .reset_index(drop=True)
    )

    if df_filtrado.empty:
        raise RuntimeError(
            "Nenhum registro permaneceu após o filtro."
        )

    return df_filtrado


# ============================================================
# 10. VERIFICAÇÃO DA INTEGRIDADE TEMPORAL
# ============================================================

def verificar_serie_horaria(
    df: pd.DataFrame,
    ano: int,
) -> None:
    """
    Verifica lacunas, duplicidades e valores inválidos.
    """

    inicio_ano = pd.Timestamp(
        f"{ano}-01-01 00:00:00"
    )

    fim_ano = pd.Timestamp(
        f"{ano}-12-31 23:00:00"
    )

    indice_esperado = pd.date_range(
        start=inicio_ano,
        end=fim_ano,
        freq="h",
    )

    indice_obtido = pd.DatetimeIndex(
        df["data_hora"]
    )

    quantidade_esperada = len(indice_esperado)
    quantidade_obtida = len(df)

    duplicados = (
        df["data_hora"]
        .duplicated()
        .sum()
    )

    horarios_ausentes = (
        indice_esperado
        .difference(indice_obtido)
    )

    horarios_excedentes = (
        indice_obtido
        .difference(indice_esperado)
    )

    valores_pld_ausentes = (
        df["PLD_R$_MWh"]
        .isna()
        .sum()
    )

    valores_pld_negativos = (
        df["PLD_R$_MWh"] < 0
    ).sum()

    print("=" * 75)
    print("VERIFICAÇÃO DA SÉRIE HORÁRIA DO PLD")
    print("=" * 75)
    print(f"Ano analisado:              {ano}")
    print(f"Registros esperados:        {quantidade_esperada}")
    print(f"Registros obtidos:          {quantidade_obtida}")
    print(f"Horários duplicados:        {duplicados}")
    print(f"Horários ausentes:          {len(horarios_ausentes)}")
    print(f"Horários excedentes:        {len(horarios_excedentes)}")
    print(f"Valores de PLD ausentes:    {valores_pld_ausentes}")
    print(f"Valores de PLD negativos:   {valores_pld_negativos}")
    print(f"Primeiro registro:          {df['data_hora'].min()}")
    print(f"Último registro:            {df['data_hora'].max()}")
    print()

    if len(horarios_ausentes) > 0:

        print("Primeiros horários ausentes:")

        for horario in horarios_ausentes[:20]:
            print(f"  - {horario}")

        print()

    if len(horarios_excedentes) > 0:

        print("Primeiros horários excedentes:")

        for horario in horarios_excedentes[:20]:
            print(f"  - {horario}")

        print()

    serie_consistente = (
        quantidade_obtida == quantidade_esperada
        and duplicados == 0
        and len(horarios_ausentes) == 0
        and len(horarios_excedentes) == 0
        and valores_pld_ausentes == 0
    )

    if serie_consistente:
        print(
            "A série horária do PLD está completa e consistente."
        )
    else:
        print(
            "ATENÇÃO: a série apresenta inconsistências."
        )

    print()


# ============================================================
# 11. VERIFICAÇÃO DAS CONTAGENS MENSAIS
# ============================================================

def verificar_contagem_mensal(
    df: pd.DataFrame,
    ano: int,
) -> pd.DataFrame:
    """
    Compara a quantidade mensal obtida com a esperada.
    """

    contagem_obtida = (
        df.groupby("mes")
        .size()
        .rename("horas_obtidas")
        .reset_index()
    )

    registros_esperados = []

    for mes in range(1, 13):

        inicio_mes = pd.Timestamp(
            year=ano,
            month=mes,
            day=1,
        )

        if mes == 12:
            inicio_mes_seguinte = pd.Timestamp(
                year=ano + 1,
                month=1,
                day=1,
            )
        else:
            inicio_mes_seguinte = pd.Timestamp(
                year=ano,
                month=mes + 1,
                day=1,
            )

        quantidade_horas = int(
            (
                inicio_mes_seguinte
                - inicio_mes
            ).total_seconds()
            / 3600
        )

        registros_esperados.append(
            {
                "mes": mes,
                "horas_esperadas": quantidade_horas,
            }
        )

    df_esperado = pd.DataFrame(
        registros_esperados
    )

    verificacao = pd.merge(
        df_esperado,
        contagem_obtida,
        on="mes",
        how="left",
    )

    verificacao["horas_obtidas"] = (
        verificacao["horas_obtidas"]
        .fillna(0)
        .astype(int)
    )

    verificacao["diferenca"] = (
        verificacao["horas_obtidas"]
        - verificacao["horas_esperadas"]
    )

    nomes_meses = {
        1: "Janeiro",
        2: "Fevereiro",
        3: "Março",
        4: "Abril",
        5: "Maio",
        6: "Junho",
        7: "Julho",
        8: "Agosto",
        9: "Setembro",
        10: "Outubro",
        11: "Novembro",
        12: "Dezembro",
    }

    verificacao["nome_mes"] = (
        verificacao["mes"]
        .map(nomes_meses)
    )

    verificacao["status"] = (
        verificacao["diferenca"]
        .apply(
            lambda valor: (
                "OK"
                if valor == 0
                else "INCONSISTENTE"
            )
        )
    )

    ordem_colunas = [
        "mes",
        "nome_mes",
        "horas_esperadas",
        "horas_obtidas",
        "diferenca",
        "status",
    ]

    verificacao = verificacao[
        ordem_colunas
    ]

    print("=" * 75)
    print("CONTAGEM MENSAL")
    print("=" * 75)
    print(
        verificacao.to_string(index=False)
    )
    print()

    return verificacao


# ============================================================
# 12. ORGANIZAÇÃO DAS COLUNAS
# ============================================================

def organizar_colunas(
    df: pd.DataFrame,
) -> pd.DataFrame:
    """
    Organiza as colunas em ordem intuitiva.
    """

    colunas_prioritarias = [
        "data_hora",
        "chave_data_hora",
        "data",
        "ano",
        "mes",
        "dia",
        "hora",
        "dia_do_ano",
        "dia_da_semana_numero",
        "dia_da_semana",

        "SUBMERCADO",
        "SUBMERCADO_ORIGINAL",

        "PLD_R$_MWh",
        "PLD_HORA",

        "MES_REFERENCIA",
        "MES_REFERENCIA_TEXTO",
        "DIA",
        "DIA_TEXTO",
        "DATA_TEXTO",

        "HORA",
        "hora_inicio",

        "PERIODO_COMERCIALIZACAO",
        "_ID",
    ]

    colunas_existentes = [
        coluna
        for coluna in colunas_prioritarias
        if coluna in df.columns
    ]

    demais_colunas = [
        coluna
        for coluna in df.columns
        if coluna not in colunas_existentes
    ]

    return df[
        colunas_existentes + demais_colunas
    ].copy()


# ============================================================
# 13. RESUMO ANUAL
# ============================================================

def criar_resumo_anual(
    df: pd.DataFrame,
) -> pd.DataFrame:
    """
    Cria indicadores estatísticos anuais.
    """

    indice_minimo = (
        df["PLD_R$_MWh"]
        .idxmin()
    )

    indice_maximo = (
        df["PLD_R$_MWh"]
        .idxmax()
    )

    linha_minimo = df.loc[indice_minimo]
    linha_maximo = df.loc[indice_maximo]

    media = df["PLD_R$_MWh"].mean()
    desvio = df["PLD_R$_MWh"].std()

    coeficiente_variacao = (
        desvio / media
        if media != 0
        else pd.NA
    )

    resumo = pd.DataFrame(
        {
            "Indicador": [
                "Ano",
                "Submercado",
                "Quantidade de registros",
                "PLD médio anual [R$/MWh]",
                "PLD mediano anual [R$/MWh]",
                "Desvio-padrão [R$/MWh]",
                "PLD mínimo [R$/MWh]",
                "Data e hora do PLD mínimo",
                "PLD máximo [R$/MWh]",
                "Data e hora do PLD máximo",
                "Primeiro quartil [R$/MWh]",
                "Terceiro quartil [R$/MWh]",
                "Coeficiente de variação",
            ],
            "Valor": [
                ANO_ANALISE,
                SUBMERCADO_ANALISE,
                len(df),
                media,
                df["PLD_R$_MWh"].median(),
                desvio,
                linha_minimo["PLD_R$_MWh"],
                linha_minimo["data_hora"],
                linha_maximo["PLD_R$_MWh"],
                linha_maximo["data_hora"],
                df["PLD_R$_MWh"].quantile(0.25),
                df["PLD_R$_MWh"].quantile(0.75),
                coeficiente_variacao,
            ],
        }
    )

    return resumo


# ============================================================
# 14. RESUMO MENSAL
# ============================================================

def criar_resumo_mensal(
    df: pd.DataFrame,
) -> pd.DataFrame:
    """
    Cria estatísticas mensais do PLD.
    """

    resumo = (
        df.groupby(
            ["ano", "mes"],
            as_index=False,
        )
        .agg(
            quantidade_horas=(
                "PLD_R$_MWh",
                "size",
            ),
            pld_medio_R_MWh=(
                "PLD_R$_MWh",
                "mean",
            ),
            pld_mediano_R_MWh=(
                "PLD_R$_MWh",
                "median",
            ),
            pld_minimo_R_MWh=(
                "PLD_R$_MWh",
                "min",
            ),
            pld_maximo_R_MWh=(
                "PLD_R$_MWh",
                "max",
            ),
            desvio_padrao_R_MWh=(
                "PLD_R$_MWh",
                "std",
            ),
        )
    )

    nomes_meses = {
        1: "Janeiro",
        2: "Fevereiro",
        3: "Março",
        4: "Abril",
        5: "Maio",
        6: "Junho",
        7: "Julho",
        8: "Agosto",
        9: "Setembro",
        10: "Outubro",
        11: "Novembro",
        12: "Dezembro",
    }

    resumo["nome_mes"] = (
        resumo["mes"]
        .map(nomes_meses)
    )

    ordem_colunas = [
        "ano",
        "mes",
        "nome_mes",
        "quantidade_horas",
        "pld_medio_R_MWh",
        "pld_mediano_R_MWh",
        "pld_minimo_R_MWh",
        "pld_maximo_R_MWh",
        "desvio_padrao_R_MWh",
    ]

    return resumo[
        ordem_colunas
    ]


# ============================================================
# 15. RESUMO DIÁRIO
# ============================================================

def criar_resumo_diario(
    df: pd.DataFrame,
) -> pd.DataFrame:
    """
    Cria indicadores diários do PLD.
    """

    resumo = (
        df.groupby(
            "data",
            as_index=False,
        )
        .agg(
            quantidade_horas=(
                "PLD_R$_MWh",
                "size",
            ),
            pld_medio_R_MWh=(
                "PLD_R$_MWh",
                "mean",
            ),
            pld_mediano_R_MWh=(
                "PLD_R$_MWh",
                "median",
            ),
            pld_minimo_R_MWh=(
                "PLD_R$_MWh",
                "min",
            ),
            pld_maximo_R_MWh=(
                "PLD_R$_MWh",
                "max",
            ),
            amplitude_diaria_R_MWh=(
                "PLD_R$_MWh",
                lambda serie: (
                    serie.max()
                    - serie.min()
                ),
            ),
            desvio_padrao_R_MWh=(
                "PLD_R$_MWh",
                "std",
            ),
        )
    )

    resumo["serie_diaria_completa"] = (
        resumo["quantidade_horas"] == 24
    )

    return (
        resumo
        .sort_values("data")
        .reset_index(drop=True)
    )


# ============================================================
# 16. PERFIL MÉDIO HORÁRIO
# ============================================================

def criar_perfil_medio_horario(
    df: pd.DataFrame,
) -> pd.DataFrame:
    """
    Calcula estatísticas por hora do dia.
    """

    perfil = (
        df.groupby(
            "hora",
            as_index=False,
        )
        .agg(
            quantidade_registros=(
                "PLD_R$_MWh",
                "size",
            ),
            pld_medio_R_MWh=(
                "PLD_R$_MWh",
                "mean",
            ),
            pld_mediano_R_MWh=(
                "PLD_R$_MWh",
                "median",
            ),
            pld_minimo_R_MWh=(
                "PLD_R$_MWh",
                "min",
            ),
            pld_maximo_R_MWh=(
                "PLD_R$_MWh",
                "max",
            ),
            desvio_padrao_R_MWh=(
                "PLD_R$_MWh",
                "std",
            ),
        )
    )

    return perfil


# ============================================================
# 17. DIAS DE MAIOR VARIAÇÃO
# ============================================================

def identificar_dias_maior_variacao(
    resumo_diario: pd.DataFrame,
    quantidade: int,
) -> pd.DataFrame:
    """
    Seleciona os dias com maior amplitude intradiária.
    """

    return (
        resumo_diario
        .sort_values(
            "amplitude_diaria_R_MWh",
            ascending=False,
        )
        .head(quantidade)
        .reset_index(drop=True)
    )


# ============================================================
# 18. RESUMO DOS EXTREMOS HORÁRIOS
# ============================================================

def identificar_maiores_menores_pld(
    df: pd.DataFrame,
    quantidade: int = 50,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Retorna as horas com maiores e menores valores de PLD.
    """

    colunas = [
        "data_hora",
        "SUBMERCADO",
        "PLD_R$_MWh",
        "mes",
        "dia",
        "hora",
        "dia_da_semana",
    ]

    maiores = (
        df[colunas]
        .sort_values(
            "PLD_R$_MWh",
            ascending=False,
        )
        .head(quantidade)
        .reset_index(drop=True)
    )

    menores = (
        df[colunas]
        .sort_values(
            "PLD_R$_MWh",
            ascending=True,
        )
        .head(quantidade)
        .reset_index(drop=True)
    )

    return maiores, menores


# ============================================================
# 19. EXPORTAÇÃO
# ============================================================

def salvar_resultados(
    df_submercado: pd.DataFrame,
    df_base_completa: pd.DataFrame,
    resumo_anual: pd.DataFrame,
    resumo_mensal: pd.DataFrame,
    resumo_diario: pd.DataFrame,
    perfil_horario: pd.DataFrame,
    dias_maior_variacao: pd.DataFrame,
    verificacao_mensal: pd.DataFrame,
    maiores_pld: pd.DataFrame,
    menores_pld: pd.DataFrame,
) -> None:
    """
    Salva os resultados em Excel e CSV.
    """

    PASTA_SAIDA.mkdir(
        parents=True,
        exist_ok=True,
    )

    caminho_excel = (
        PASTA_SAIDA
        / f"{NOME_BASE_ARQUIVO}.xlsx"
    )

    caminho_csv = (
        PASTA_SAIDA
        / f"{NOME_BASE_ARQUIVO}.csv"
    )

    with pd.ExcelWriter(
        caminho_excel,
        engine="openpyxl",
    ) as escritor:

        df_submercado.to_excel(
            escritor,
            sheet_name="PLD_Horario_SUDESTE",
            index=False,
        )

        resumo_anual.to_excel(
            escritor,
            sheet_name="Resumo_Anual",
            index=False,
        )

        resumo_mensal.to_excel(
            escritor,
            sheet_name="Resumo_Mensal",
            index=False,
        )

        resumo_diario.to_excel(
            escritor,
            sheet_name="Resumo_Diario",
            index=False,
        )

        perfil_horario.to_excel(
            escritor,
            sheet_name="Perfil_Medio_Horario",
            index=False,
        )

        dias_maior_variacao.to_excel(
            escritor,
            sheet_name="Dias_Maior_Variacao",
            index=False,
        )

        verificacao_mensal.to_excel(
            escritor,
            sheet_name="Verificacao_Mensal",
            index=False,
        )

        maiores_pld.to_excel(
            escritor,
            sheet_name="Maiores_PLD",
            index=False,
        )

        menores_pld.to_excel(
            escritor,
            sheet_name="Menores_PLD",
            index=False,
        )

        if SALVAR_BASE_COMPLETA:

            df_base_completa.to_excel(
                escritor,
                sheet_name="Todos_Submercados",
                index=False,
            )

    if SALVAR_CSV:

        df_submercado.to_csv(
            caminho_csv,
            index=False,
            sep=";",
            decimal=",",
            encoding="utf-8-sig",
        )

    print("=" * 75)
    print("ARQUIVOS GERADOS")
    print("=" * 75)
    print(f"Excel:\n{caminho_excel.resolve()}")
    print()

    if SALVAR_CSV:
        print(f"CSV:\n{caminho_csv.resolve()}")
        print()


# ============================================================
# 20. FUNÇÃO PRINCIPAL
# ============================================================

def executar() -> None:
    """
    Executa todo o processamento.
    """

    df_base_bruta = baixar_todos_registros_ccee(
        resource_id=RESOURCE_ID_CCEE,
        limite_por_requisicao=(
            LIMITE_REGISTROS_POR_REQUISICAO
        ),
    )

    print("=" * 75)
    print("COLUNAS RECEBIDAS")
    print("=" * 75)
    print(df_base_bruta.columns.tolist())
    print()

    df_base_tratada = tratar_base_ccee(
        df_base_bruta
    )

    df_pld = filtrar_ano_submercado(
        df=df_base_tratada,
        ano=ANO_ANALISE,
        submercado=SUBMERCADO_ANALISE,
    )

    verificar_serie_horaria(
        df=df_pld,
        ano=ANO_ANALISE,
    )

    verificacao_mensal = verificar_contagem_mensal(
        df=df_pld,
        ano=ANO_ANALISE,
    )

    df_base_tratada = organizar_colunas(
        df_base_tratada
    )

    df_pld = organizar_colunas(
        df_pld
    )

    resumo_anual = criar_resumo_anual(
        df_pld
    )

    resumo_mensal = criar_resumo_mensal(
        df_pld
    )

    resumo_diario = criar_resumo_diario(
        df_pld
    )

    perfil_horario = criar_perfil_medio_horario(
        df_pld
    )

    dias_maior_variacao = (
        identificar_dias_maior_variacao(
            resumo_diario=resumo_diario,
            quantidade=QUANTIDADE_DIAS_MAIOR_VARIACAO,
        )
    )

    maiores_pld, menores_pld = (
        identificar_maiores_menores_pld(
            df=df_pld,
            quantidade=50,
        )
    )

    print("=" * 75)
    print("PRIMEIROS REGISTROS")
    print("=" * 75)
    print(
        df_pld[
            [
                "data_hora",
                "SUBMERCADO",
                "PLD_R$_MWh",
            ]
        ].head(10)
    )
    print()

    print("=" * 75)
    print("ÚLTIMOS REGISTROS")
    print("=" * 75)
    print(
        df_pld[
            [
                "data_hora",
                "SUBMERCADO",
                "PLD_R$_MWh",
            ]
        ].tail(10)
    )
    print()

    print("=" * 75)
    print("RESUMO ANUAL")
    print("=" * 75)
    print(
        resumo_anual.to_string(index=False)
    )
    print()

    salvar_resultados(
        df_submercado=df_pld,
        df_base_completa=df_base_tratada,
        resumo_anual=resumo_anual,
        resumo_mensal=resumo_mensal,
        resumo_diario=resumo_diario,
        perfil_horario=perfil_horario,
        dias_maior_variacao=dias_maior_variacao,
        verificacao_mensal=verificacao_mensal,
        maiores_pld=maiores_pld,
        menores_pld=menores_pld,
    )

    print("=" * 75)
    print("PROCESSAMENTO CONCLUÍDO")
    print("=" * 75)


# ============================================================
# 21. INÍCIO DO PROGRAMA
# ============================================================

if __name__ == "__main__":
    executar()