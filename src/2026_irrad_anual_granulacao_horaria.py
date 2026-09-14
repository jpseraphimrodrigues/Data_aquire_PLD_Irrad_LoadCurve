# ============================================================
# CURVA ANUAL HORÁRIA DE IRRADIÂNCIA E GERAÇÃO FOTOVOLTAICA
# Fonte: PVGIS 5.3
#
# Objetivo:
# 1. Baixar dados horários de irradiância solar;
# 2. Estimar a geração de um sistema fotovoltaico;
# 3. Converter o horário UTC para o horário de Brasília;
# 4. Selecionar exatamente o ano civil desejado;
# 5. Exportar os resultados para Excel e CSV;
# 6. Preparar uma série horária para cruzamento com o PLD.
#
# Bibliotecas necessárias:
#
# python -m pip install requests pandas openpyxl
# ============================================================

from pathlib import Path

import pandas as pd
import requests


# ============================================================
# 1. CONFIGURAÇÕES PRINCIPAIS
# ============================================================

# Coordenadas aproximadas de Sorocaba/SP
LATITUDE = -23.5015
LONGITUDE = -47.4526

# Ano desejado para a curva anual
ANO_ANALISE = 2022

# Fuso horário local
FUSO_HORARIO_LOCAL = "America/Sao_Paulo"

# Potência de referência do sistema fotovoltaico
#
# Utilizando 1 kWp, a curva obtida pode ser escalada depois
# para qualquer potência instalada.
POTENCIA_PICO_KW = 1.0

# Inclinação dos módulos em relação ao plano horizontal
INCLINACAO_GRAUS = 23.0

# Convenção do PVGIS:
#
#   0° = Sul
# -90° = Leste
#  90° = Oeste
# 180° = Norte
#
# Para o hemisfério sul, normalmente os módulos são orientados
# para o norte.
AZIMUTE_PVGIS = 180.0

# Perdas globais estimadas do sistema fotovoltaico
PERDAS_PERCENTUAIS = 14.0

# Tipo de módulo fotovoltaico
TECNOLOGIA_FV = "crystSi"

# Tipo de instalação
LOCAL_INSTALACAO = "free"

# Tempo máximo da requisição
TIMEOUT_SEGUNDOS = 180

# Pasta de saída
PASTA_SAIDA = Path("resultados_irradiancia_pvgis")

# Nome dos arquivos
NOME_BASE_ARQUIVO = (
    f"Curva_Horaria_Irradiancia_Geracao_PVGIS_{ANO_ANALISE}"
)

# Salvar também em CSV
SALVAR_CSV = True


# ============================================================
# 2. REQUISIÇÃO À API DO PVGIS
# ============================================================

def requisitar_dados_pvgis(
    latitude: float,
    longitude: float,
    ano_inicial: int,
    ano_final: int,
    potencia_pico_kw: float,
    inclinacao_graus: float,
    azimute_pvgis: float,
    perdas_percentuais: float,
) -> dict:
    """
    Faz a requisição dos dados horários ao PVGIS 5.3.

    Retorna a resposta JSON completa.
    """

    url = "https://re.jrc.ec.europa.eu/api/v5_3/seriescalc"

    parametros = {
        # Localização
        "lat": latitude,
        "lon": longitude,

        # Período
        "startyear": ano_inicial,
        "endyear": ano_final,

        # Cálculo fotovoltaico
        "pvcalculation": 1,
        "peakpower": potencia_pico_kw,
        "pvtechchoice": TECNOLOGIA_FV,
        "mountingplace": LOCAL_INSTALACAO,
        "loss": perdas_percentuais,

        # Geometria dos módulos
        "trackingtype": 0,
        "angle": inclinacao_graus,
        "aspect": azimute_pvgis,

        # Componentes da irradiância
        "components": 1,

        # Horizonte
        "usehorizon": 1,

        # Formato da saída
        "outputformat": "json",
    }

    print("=" * 75)
    print("REQUISIÇÃO AO PVGIS")
    print("=" * 75)
    print(f"Latitude:                  {latitude}")
    print(f"Longitude:                 {longitude}")
    print(f"Período solicitado:        {ano_inicial} a {ano_final}")
    print(f"Potência de referência:    {potencia_pico_kw} kWp")
    print(f"Inclinação:                {inclinacao_graus}°")
    print(f"Azimute PVGIS:             {azimute_pvgis}°")
    print(f"Perdas do sistema:         {perdas_percentuais}%")
    print()

    try:
        resposta = requests.get(
            url,
            params=parametros,
            timeout=TIMEOUT_SEGUNDOS,
        )

    except requests.exceptions.Timeout as erro:
        raise RuntimeError(
            "A requisição ao PVGIS excedeu o tempo máximo."
        ) from erro

    except requests.exceptions.ConnectionError as erro:
        raise RuntimeError(
            "Não foi possível conectar ao PVGIS. "
            "Verifique a conexão com a internet."
        ) from erro

    except requests.exceptions.RequestException as erro:
        raise RuntimeError(
            f"Erro durante a requisição ao PVGIS: {erro}"
        ) from erro

    if resposta.status_code != 200:
        raise RuntimeError(
            "\nErro retornado pelo PVGIS.\n"
            f"Status HTTP: {resposta.status_code}\n"
            f"Resposta:\n{resposta.text[:2000]}"
        )

    try:
        dados_json = resposta.json()

    except ValueError as erro:
        raise RuntimeError(
            "A resposta do PVGIS não pôde ser convertida para JSON."
        ) from erro

    if "outputs" not in dados_json:
        raise RuntimeError(
            "A resposta do PVGIS não contém o campo 'outputs'."
        )

    if "hourly" not in dados_json["outputs"]:
        raise RuntimeError(
            "A resposta não contém a série horária esperada."
        )

    print("Requisição concluída com sucesso.")
    print()

    return dados_json


# ============================================================
# 3. CRIAÇÃO E TRATAMENTO DO DATAFRAME
# ============================================================

def criar_dataframe_pvgis(
    dados_json: dict,
    potencia_pico_kw: float,
) -> pd.DataFrame:
    """
    Converte a resposta do PVGIS para um DataFrame organizado.

    O arredondamento é feito primeiro em UTC. Isso evita erros
    relacionados ao antigo horário de verão brasileiro.
    """

    dados_horarios = dados_json["outputs"]["hourly"]

    df = pd.DataFrame(dados_horarios)

    if df.empty:
        raise RuntimeError(
            "O PVGIS retornou uma tabela vazia."
        )

    if "time" not in df.columns:
        raise RuntimeError(
            "A coluna temporal 'time' não foi encontrada."
        )

    # --------------------------------------------------------
    # Conversão da data original do PVGIS
    # --------------------------------------------------------
    #
    # Exemplo:
    # 20200101:0003
    #
    # O horário retornado pelo PVGIS está em UTC.

    df["data_hora_utc_original"] = pd.to_datetime(
        df["time"],
        format="%Y%m%d:%H%M",
        errors="coerce",
        utc=True,
    )

    quantidade_datas_invalidas = (
        df["data_hora_utc_original"].isna().sum()
    )

    if quantidade_datas_invalidas > 0:
        raise RuntimeError(
            f"Foram encontradas {quantidade_datas_invalidas} "
            "datas inválidas."
        )

    # --------------------------------------------------------
    # Ajuste para horas inteiras
    # --------------------------------------------------------
    #
    # O PVGIS pode fornecer horários como:
    #
    # 00:03
    # 01:03
    # 02:03
    #
    # O arredondamento é feito em UTC, antes da conversão para
    # o horário local, evitando ambiguidades de horário de verão.

    df["data_hora_utc"] = (
        df["data_hora_utc_original"]
        .dt.floor("h")
    )

    # --------------------------------------------------------
    # Conversão para o horário de Brasília
    # --------------------------------------------------------

    df["data_hora_local_original"] = (
        df["data_hora_utc_original"]
        .dt.tz_convert(FUSO_HORARIO_LOCAL)
    )

    df["data_hora_local"] = (
        df["data_hora_utc"]
        .dt.tz_convert(FUSO_HORARIO_LOCAL)
    )

    # --------------------------------------------------------
    # Conversão das colunas numéricas
    # --------------------------------------------------------

    colunas_numericas = [
        "P",
        "Gb(i)",
        "Gd(i)",
        "Gr(i)",
        "H_sun",
        "T2m",
        "WS10m",
        "Int",
    ]

    for coluna in colunas_numericas:
        if coluna in df.columns:
            df[coluna] = pd.to_numeric(
                df[coluna],
                errors="coerce",
            )

    # --------------------------------------------------------
    # Potência e energia fotovoltaica
    # --------------------------------------------------------

    if "P" not in df.columns:
        raise RuntimeError(
            "A coluna de potência fotovoltaica 'P' não foi encontrada."
        )

    df["potencia_fv_w"] = (
        df["P"]
        .fillna(0.0)
        .clip(lower=0.0)
    )

    df["potencia_fv_kw"] = (
        df["potencia_fv_w"] / 1000.0
    )

    # Como cada registro representa uma hora:
    #
    # energia [kWh] = potência média [kW] × 1 hora
    df["energia_fv_kwh"] = (
        df["potencia_fv_kw"] * 1.0
    )

    # Curva normalizada pela potência nominal
    if potencia_pico_kw > 0:
        df["geracao_fv_pu"] = (
            df["potencia_fv_kw"]
            / potencia_pico_kw
        )
    else:
        df["geracao_fv_pu"] = 0.0

    df["geracao_fv_pu"] = (
        df["geracao_fv_pu"]
        .fillna(0.0)
        .clip(lower=0.0)
    )

    # --------------------------------------------------------
    # Irradiância no plano dos módulos
    # --------------------------------------------------------

    componentes_irradiancia = [
        "Gb(i)",
        "Gd(i)",
        "Gr(i)",
    ]

    if all(
        coluna in df.columns
        for coluna in componentes_irradiancia
    ):
        df["irradiancia_plano_w_m2"] = (
            df["Gb(i)"].fillna(0.0)
            + df["Gd(i)"].fillna(0.0)
            + df["Gr(i)"].fillna(0.0)
        )

    else:
        df["irradiancia_plano_w_m2"] = pd.NA

    # --------------------------------------------------------
    # Ordenação
    # --------------------------------------------------------

    df = (
        df.sort_values("data_hora_utc")
        .reset_index(drop=True)
    )

    return df


# ============================================================
# 4. FILTRO DO ANO CIVIL LOCAL
# ============================================================

def filtrar_ano_civil_local(
    df: pd.DataFrame,
    ano: int,
) -> pd.DataFrame:
    """
    Seleciona exatamente o ano civil desejado no horário local.

    Exemplo para 2020:

    início: 01/01/2020 00:00
    fim:    antes de 01/01/2021 00:00
    """

    inicio_ano = pd.Timestamp(
        f"{ano}-01-01 00:00:00",
        tz=FUSO_HORARIO_LOCAL,
    )

    inicio_ano_seguinte = pd.Timestamp(
        f"{ano + 1}-01-01 00:00:00",
        tz=FUSO_HORARIO_LOCAL,
    )

    filtro = (
        (df["data_hora_local"] >= inicio_ano)
        & (df["data_hora_local"] < inicio_ano_seguinte)
    )

    df_ano = df.loc[filtro].copy()

    df_ano = (
        df_ano.sort_values("data_hora_utc")
        .reset_index(drop=True)
    )

    return df_ano


# ============================================================
# 5. CRIAÇÃO DE COLUNAS AUXILIARES
# ============================================================

def adicionar_colunas_auxiliares(
    df: pd.DataFrame,
) -> pd.DataFrame:
    """
    Adiciona colunas úteis para análises posteriores com o PLD.
    """

    df = df.copy()

    df["ano"] = df["data_hora_local"].dt.year
    df["mes"] = df["data_hora_local"].dt.month
    df["dia"] = df["data_hora_local"].dt.day
    df["hora"] = df["data_hora_local"].dt.hour
    df["dia_do_ano"] = df["data_hora_local"].dt.dayofyear
    df["dia_da_semana_numero"] = (
        df["data_hora_local"].dt.dayofweek
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

    # Data sem a informação de hora
    df["data_local"] = (
        df["data_hora_local"].dt.date
    )

    # Identificador textual compatível com cruzamentos
    df["chave_data_hora"] = (
        df["data_hora_local"]
        .dt.strftime("%Y-%m-%d %H:00:00")
    )

    return df


# ============================================================
# 6. VERIFICAÇÃO DA SÉRIE HORÁRIA
# ============================================================

def verificar_serie_horaria(
    df: pd.DataFrame,
    ano: int,
) -> None:
    """
    Verifica quantidade de registros, duplicidades e lacunas.
    """

    inicio_utc = df["data_hora_utc"].min()
    fim_utc = df["data_hora_utc"].max()

    if pd.isna(inicio_utc) or pd.isna(fim_utc):
        raise RuntimeError(
            "Não foi possível identificar o período da série."
        )

    quantidade_obtida = len(df)

    # Para anos sem horário de verão:
    # ano comum = 8.760
    # ano bissexto = 8.784
    quantidade_teorica = (
        8784
        if pd.Timestamp(f"{ano}-12-31").dayofyear == 366
        else 8760
    )

    duplicados_utc = (
        df["data_hora_utc"]
        .duplicated()
        .sum()
    )

    duplicados_local_com_fuso = (
        df["data_hora_local"]
        .duplicated()
        .sum()
    )

    print("=" * 75)
    print("VERIFICAÇÃO DA SÉRIE HORÁRIA")
    print("=" * 75)
    print(f"Ano analisado:                   {ano}")
    print(f"Registros teóricos:              {quantidade_teorica}")
    print(f"Registros obtidos:               {quantidade_obtida}")
    print(f"Duplicados no horário UTC:       {duplicados_utc}")
    print(
        f"Duplicados no horário local:     "
        f"{duplicados_local_com_fuso}"
    )
    print(f"Primeiro horário local:          {df['data_hora_local'].min()}")
    print(f"Último horário local:            {df['data_hora_local'].max()}")
    print()

    if quantidade_obtida == quantidade_teorica:
        print("A quantidade de registros está correta.")
    else:
        print(
            "ATENÇÃO: a quantidade de registros difere do valor "
            "teórico esperado."
        )

    if duplicados_utc == 0:
        print("Não existem horários UTC duplicados.")
    else:
        print("ATENÇÃO: existem horários UTC duplicados.")

    print()


# ============================================================
# 7. ORGANIZAÇÃO DAS COLUNAS
# ============================================================

def organizar_colunas(
    df: pd.DataFrame,
) -> pd.DataFrame:
    """
    Organiza as colunas em uma ordem mais intuitiva.
    """

    colunas_prioritarias = [
        # Datas e identificação
        "data_hora_local",
        "data_hora_local_original",
        "data_hora_utc",
        "data_hora_utc_original",
        "time",
        "chave_data_hora",
        "data_local",
        "ano",
        "mes",
        "dia",
        "hora",
        "dia_do_ano",
        "dia_da_semana_numero",
        "dia_da_semana",

        # Geração fotovoltaica
        "potencia_fv_w",
        "potencia_fv_kw",
        "energia_fv_kwh",
        "geracao_fv_pu",

        # Irradiância
        "irradiancia_plano_w_m2",
        "Gb(i)",
        "Gd(i)",
        "Gr(i)",

        # Variáveis meteorológicas
        "H_sun",
        "T2m",
        "WS10m",
        "Int",
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
# 8. RESUMO ANUAL
# ============================================================

def criar_resumo_anual(
    df: pd.DataFrame,
    ano: int,
    potencia_pico_kw: float,
) -> pd.DataFrame:
    """
    Cria uma tabela com os principais indicadores anuais.
    """

    energia_anual_kwh = (
        df["energia_fv_kwh"].sum()
    )

    potencia_maxima_kw = (
        df["potencia_fv_kw"].max()
    )

    irradiancia_media_w_m2 = (
        df["irradiancia_plano_w_m2"].mean()
    )

    irradiancia_maxima_w_m2 = (
        df["irradiancia_plano_w_m2"].max()
    )

    horas_equivalentes_anuais = (
        energia_anual_kwh / potencia_pico_kw
        if potencia_pico_kw > 0
        else 0.0
    )

    fator_capacidade = (
        energia_anual_kwh
        / (potencia_pico_kw * len(df))
        if potencia_pico_kw > 0 and len(df) > 0
        else 0.0
    )

    horas_com_geracao = (
        (df["potencia_fv_kw"] > 0).sum()
    )

    resumo = pd.DataFrame(
        {
            "Indicador": [
                "Ano analisado",
                "Latitude",
                "Longitude",
                "Fuso horário",
                "Potência de referência [kWp]",
                "Inclinação [graus]",
                "Azimute PVGIS [graus]",
                "Perdas do sistema [%]",
                "Registros horários",
                "Horas com geração fotovoltaica",
                "Energia anual estimada [kWh]",
                "Horas equivalentes anuais [h]",
                "Fator de capacidade",
                "Potência máxima estimada [kW]",
                "Irradiância média no plano [W/m²]",
                "Irradiância máxima no plano [W/m²]",
            ],
            "Valor": [
                ano,
                LATITUDE,
                LONGITUDE,
                FUSO_HORARIO_LOCAL,
                potencia_pico_kw,
                INCLINACAO_GRAUS,
                AZIMUTE_PVGIS,
                PERDAS_PERCENTUAIS,
                len(df),
                horas_com_geracao,
                energia_anual_kwh,
                horas_equivalentes_anuais,
                fator_capacidade,
                potencia_maxima_kw,
                irradiancia_media_w_m2,
                irradiancia_maxima_w_m2,
            ],
        }
    )

    return resumo


# ============================================================
# 9. RESUMO MENSAL
# ============================================================

def criar_resumo_mensal(
    df: pd.DataFrame,
) -> pd.DataFrame:
    """
    Cria uma tabela mensal de geração e irradiância.
    """

    resumo_mensal = (
        df.groupby(
            ["ano", "mes"],
            as_index=False,
        )
        .agg(
            energia_fv_kwh=("energia_fv_kwh", "sum"),
            potencia_maxima_kw=("potencia_fv_kw", "max"),
            irradiancia_media_w_m2=(
                "irradiancia_plano_w_m2",
                "mean",
            ),
            irradiancia_maxima_w_m2=(
                "irradiancia_plano_w_m2",
                "max",
            ),
            horas_com_geracao=(
                "potencia_fv_kw",
                lambda serie: int((serie > 0).sum()),
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

    resumo_mensal["nome_mes"] = (
        resumo_mensal["mes"]
        .map(nomes_meses)
    )

    colunas = [
        "ano",
        "mes",
        "nome_mes",
        "energia_fv_kwh",
        "potencia_maxima_kw",
        "irradiancia_media_w_m2",
        "irradiancia_maxima_w_m2",
        "horas_com_geracao",
    ]

    return resumo_mensal[colunas]


# ============================================================
# 10. REMOÇÃO DO TIMEZONE PARA EXCEL
# ============================================================

def remover_timezone_para_exportacao(
    df: pd.DataFrame,
) -> pd.DataFrame:
    """
    Remove o timezone das colunas datetime.

    O Excel não aceita datas com timezone. A hora mostrada é
    mantida; somente a informação formal do fuso é retirada.
    """

    df_exportacao = df.copy()

    for coluna in df_exportacao.columns:
        if isinstance(
            df_exportacao[coluna].dtype,
            pd.DatetimeTZDtype,
        ):
            df_exportacao[coluna] = (
                df_exportacao[coluna]
                .dt.tz_localize(None)
            )

    return df_exportacao


# ============================================================
# 11. EXPORTAÇÃO PARA EXCEL E CSV
# ============================================================

def salvar_resultados(
    df: pd.DataFrame,
    resumo_anual: pd.DataFrame,
    resumo_mensal: pd.DataFrame,
    pasta_saida: Path,
    nome_base: str,
    salvar_csv: bool,
) -> None:
    """
    Salva a série horária, o resumo anual e o resumo mensal.
    """

    pasta_saida.mkdir(
        parents=True,
        exist_ok=True,
    )

    caminho_excel = (
        pasta_saida / f"{nome_base}.xlsx"
    )

    caminho_csv = (
        pasta_saida / f"{nome_base}.csv"
    )

    df_excel = remover_timezone_para_exportacao(df)

    with pd.ExcelWriter(
        caminho_excel,
        engine="openpyxl",
    ) as escritor:

        df_excel.to_excel(
            escritor,
            sheet_name="Serie_Horaria",
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

    print("=" * 75)
    print("ARQUIVOS GERADOS")
    print("=" * 75)
    print(f"Excel:\n{caminho_excel.resolve()}")
    print()

    if salvar_csv:
        df_csv = remover_timezone_para_exportacao(df)

        df_csv.to_csv(
            caminho_csv,
            index=False,
            sep=";",
            decimal=",",
            encoding="utf-8-sig",
        )

        print(f"CSV:\n{caminho_csv.resolve()}")
        print()


# ============================================================
# 12. FUNÇÃO PRINCIPAL
# ============================================================

def executar() -> None:
    """
    Executa todo o processamento.
    """

    # Como o horário de Sorocaba está atrás do UTC, precisamos
    # solicitar o ano analisado e o ano seguinte.
    #
    # Exemplo:
    #
    # Para obter até 31/12/2020 23:00 no horário de Brasília,
    # precisamos dos registros de 01/01/2021 em UTC.
    #
    # Não é necessário solicitar o ano anterior.
    ano_inicial_requisicao = ANO_ANALISE
    ano_final_requisicao = ANO_ANALISE + 1

    dados_json = requisitar_dados_pvgis(
        latitude=LATITUDE,
        longitude=LONGITUDE,
        ano_inicial=ano_inicial_requisicao,
        ano_final=ano_final_requisicao,
        potencia_pico_kw=POTENCIA_PICO_KW,
        inclinacao_graus=INCLINACAO_GRAUS,
        azimute_pvgis=AZIMUTE_PVGIS,
        perdas_percentuais=PERDAS_PERCENTUAIS,
    )

    df_solar = criar_dataframe_pvgis(
        dados_json=dados_json,
        potencia_pico_kw=POTENCIA_PICO_KW,
    )

    df_solar = filtrar_ano_civil_local(
        df=df_solar,
        ano=ANO_ANALISE,
    )

    df_solar = adicionar_colunas_auxiliares(
        df=df_solar,
    )

    df_solar = organizar_colunas(
        df=df_solar,
    )

    verificar_serie_horaria(
        df=df_solar,
        ano=ANO_ANALISE,
    )

    resumo_anual = criar_resumo_anual(
        df=df_solar,
        ano=ANO_ANALISE,
        potencia_pico_kw=POTENCIA_PICO_KW,
    )

    resumo_mensal = criar_resumo_mensal(
        df=df_solar,
    )

    print("=" * 75)
    print("PRIMEIROS REGISTROS")
    print("=" * 75)
    print(
        df_solar[
            [
                "data_hora_local",
                "potencia_fv_kw",
                "energia_fv_kwh",
                "geracao_fv_pu",
                "irradiancia_plano_w_m2",
            ]
        ].head(10)
    )
    print()

    print("=" * 75)
    print("ÚLTIMOS REGISTROS")
    print("=" * 75)
    print(
        df_solar[
            [
                "data_hora_local",
                "potencia_fv_kw",
                "energia_fv_kwh",
                "geracao_fv_pu",
                "irradiancia_plano_w_m2",
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
        df=df_solar,
        resumo_anual=resumo_anual,
        resumo_mensal=resumo_mensal,
        pasta_saida=PASTA_SAIDA,
        nome_base=NOME_BASE_ARQUIVO,
        salvar_csv=SALVAR_CSV,
    )

    print("=" * 75)
    print("PROCESSAMENTO CONCLUÍDO")
    print("=" * 75)


# ============================================================
# 13. INÍCIO DO PROGRAMA
# ============================================================

if __name__ == "__main__":
    executar()