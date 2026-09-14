
#Visualizador temporal integrado — Irradiância, Carga ONS e PLD CCEE
#==================================================================


from pathlib import Path
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

try:
    import xlwt
except ImportError:
    xlwt = None


# =============================================================================
# 1. PARÂMETROS DO USUÁRIO
# =============================================================================

ARQUIVO_IRRADIANCIA = Path("resultados_irradiancia_pvgis\\Curva_Horaria_Irradiancia_Geracao_PVGIS_2022.xlsx")
ARQUIVO_CARGA = Path("resultados_curva_carga_ons\\Curva_Carga_Horaria_ONS_2022_SUDESTE_CENTRO_OESTE.xlsx")
ARQUIVO_PLD = Path("resultados_pld_ccee\\PLD_Horario_CCEE_2022_SUDESTE.xlsx")

ABA_IRRADIANCIA = "Serie_Horaria"
ABA_CARGA = "Carga_Horaria_SE"
ABA_PLD = "PLD_Horario_SUDESTE"

# -----------------------------------------------------------------------------
# PERÍODO DE ANÁLISE
# -----------------------------------------------------------------------------
# Para plotar o ano inteiro:
# DATA_INICIAL = None
# DATA_FINAL = None
#
# Exemplo — uma semana:
# DATA_INICIAL = "2022-07-01"
# DATA_FINAL = "2022-07-07"
#
# Exemplo — um único dia:
# DATA_INICIAL = "2022-07-15"
# DATA_FINAL = "2022-07-15"

DATA_INICIAL = None
DATA_FINAL = None

# Normalização do gráfico conjunto:
#   "anual"   -> cada série é dividida pelo máximo do ano inteiro.
#                Recomendado para comparar diferentes períodos entre si.
#   "periodo" -> cada série é dividida pelo máximo do período selecionado.
MODO_NORMALIZACAO = "anual"

# Exibição e salvamento
EXIBIR_GRAFICOS = True
SALVAR_GRAFICOS = True
PASTA_SAIDA = Path("graficos_irradiancia_carga_pld")
DPI_SAIDA = 300

# -----------------------------------------------------------------------------
# EXPORTAÇÃO DOS DADOS DO PERÍODO SELECIONADO
# -----------------------------------------------------------------------------
# Gera um único arquivo .xls com três planilhas compatíveis com os códigos
# de simulação e análise econômica:
#   CurvasPV              -> curva FV / irradiância
#   LoadShape48_Cargas    -> curva de carga
#   CurvasPLD             -> PLD horário
GERAR_ARQUIVO_DADOS = True
ARQUIVO_DADOS = Path("dados.xls")


# =============================================================================
# 2. LEITURA DAS BASES
# =============================================================================

def ler_base_irradiancia(caminho: Path) -> pd.DataFrame:
    """Lê apenas as colunas necessárias da base PVGIS."""

    df = pd.read_excel(
        caminho,
        sheet_name=ABA_IRRADIANCIA,
        usecols=[
            "data_hora_local",
            "irradiancia_plano_w_m2",
            "potencia_fv_kw",
            "geracao_fv_pu",
        ],
    )

    df = df.rename(columns={"data_hora_local": "data_hora"})
    df["data_hora"] = pd.to_datetime(df["data_hora"])

    return df.sort_values("data_hora").reset_index(drop=True)


def ler_base_carga(caminho: Path) -> pd.DataFrame:
    """Lê apenas as colunas necessárias da curva horária do ONS."""

    df = pd.read_excel(
        caminho,
        sheet_name=ABA_CARGA,
        usecols=[
            "data_hora",
            "carga_mwmed",
            "carga_pu_pico",
        ],
    )

    df["data_hora"] = pd.to_datetime(df["data_hora"])

    return df.sort_values("data_hora").reset_index(drop=True)


def ler_base_pld(caminho: Path) -> pd.DataFrame:
    """Lê apenas as colunas necessárias da base horária de PLD da CCEE."""

    df = pd.read_excel(
        caminho,
        sheet_name=ABA_PLD,
        usecols=[
            "data_hora",
            "PLD_R$_MWh",
        ],
    )

    df["data_hora"] = pd.to_datetime(df["data_hora"])

    return df.sort_values("data_hora").reset_index(drop=True)


# =============================================================================
# 3. VALIDAÇÃO E SINCRONIZAÇÃO TEMPORAL
# =============================================================================

def validar_base(df: pd.DataFrame, nome_base: str) -> None:
    """Executa verificações simples antes do cruzamento das séries."""

    if df.empty:
        raise ValueError(f"A base '{nome_base}' está vazia.")

    if df["data_hora"].isna().any():
        raise ValueError(f"A base '{nome_base}' possui datas/horas inválidas.")

    if df["data_hora"].duplicated().any():
        duplicadas = df.loc[df["data_hora"].duplicated(), "data_hora"]
        raise ValueError(
            f"A base '{nome_base}' possui timestamps duplicados. "
            f"Primeiros casos: {duplicadas.head().tolist()}"
        )


def sincronizar_bases(
    df_irradiancia: pd.DataFrame,
    df_carga: pd.DataFrame,
    df_pld: pd.DataFrame,
) -> pd.DataFrame:
    """Une as três bases utilizando o timestamp horário como chave."""

    dados = df_irradiancia.merge(
        df_carga,
        on="data_hora",
        how="inner",
        validate="one_to_one",
    )

    dados = dados.merge(
        df_pld,
        on="data_hora",
        how="inner",
        validate="one_to_one",
    )

    dados = dados.sort_values("data_hora").reset_index(drop=True)

    if dados.empty:
        raise ValueError("Não houve coincidência temporal entre as três bases.")

    return dados


def imprimir_resumo_sincronizacao(
    df_irradiancia: pd.DataFrame,
    df_carga: pd.DataFrame,
    df_pld: pd.DataFrame,
    dados_integrados: pd.DataFrame,
) -> None:
    """Mostra no terminal um resumo rápido da qualidade do cruzamento."""

    print("\n" + "=" * 78)
    print("VALIDAÇÃO DAS BASES HORÁRIAS")
    print("=" * 78)
    print(f"Irradiância/PVGIS : {len(df_irradiancia):5d} registros")
    print(f"Carga/ONS         : {len(df_carga):5d} registros")
    print(f"PLD/CCEE          : {len(df_pld):5d} registros")
    print(f"Registros comuns  : {len(dados_integrados):5d} registros")
    print(
        "Período integrado  : "
        f"{dados_integrados['data_hora'].min()} até "
        f"{dados_integrados['data_hora'].max()}"
    )

    tamanho_minimo = min(len(df_irradiancia), len(df_carga), len(df_pld))

    if len(dados_integrados) == tamanho_minimo:
        print("Sincronização      : OK — todos os timestamps disponíveis coincidem.")
    else:
        print(
            "ATENÇÃO           : existem timestamps presentes em uma base "
            "e ausentes em outra."
        )


# =============================================================================
# 4. FILTRO DO PERÍODO
# =============================================================================

def filtrar_periodo(
    dados: pd.DataFrame,
    data_inicial=None,
    data_final=None,
) -> pd.DataFrame:
    """
    Filtra o intervalo escolhido.

    Se a data final for informada apenas como AAAA-MM-DD, o filtro inclui
    automaticamente todas as 24 horas daquele dia.
    """

    dados_filtrados = dados.copy()

    if data_inicial is not None:
        inicio = pd.Timestamp(data_inicial)
        dados_filtrados = dados_filtrados[
            dados_filtrados["data_hora"] >= inicio
        ]

    if data_final is not None:
        fim = pd.Timestamp(data_final)

        # Se o usuário informou somente uma data, inclui o dia inteiro.
        if fim == fim.normalize():
            fim = fim + pd.Timedelta(days=1) - pd.Timedelta(microseconds=1)

        dados_filtrados = dados_filtrados[
            dados_filtrados["data_hora"] <= fim
        ]

    dados_filtrados = dados_filtrados.reset_index(drop=True)

    if dados_filtrados.empty:
        raise ValueError(
            "O período selecionado não possui dados. "
            "Verifique DATA_INICIAL e DATA_FINAL."
        )

    return dados_filtrados


# =============================================================================
# 5. EXPORTAÇÃO DO PERÍODO SELECIONADO PARA dados.xls
# =============================================================================

def _escrever_cabecalhos_xls(planilha, cabecalhos):
    """Escreve os cabeçalhos com formatação simples no arquivo .xls."""
    estilo_cabecalho = xlwt.easyxf("font: bold on; align: horiz center")

    for coluna, cabecalho in enumerate(cabecalhos):
        planilha.write(0, coluna, cabecalho, estilo_cabecalho)


def _ajustar_larguras_xls(planilha, larguras):
    """Ajusta larguras aproximadas das colunas do arquivo .xls."""
    for coluna, largura_caracteres in enumerate(larguras):
        planilha.col(coluna).width = int(largura_caracteres * 256)


def exportar_dados_periodo_xls(
    dados_periodo: pd.DataFrame,
    caminho_saida: Path,
) -> None:
    """
    Gera um único arquivo ``dados.xls`` com três planilhas.

    As duas primeiras colunas de cada planilha preservam exatamente a estrutura
    dos arquivos de entrada utilizados pelos códigos técnicos/econômicos:

    CurvasPV
        Curva_PV_kW | Hora

    LoadShape48_Cargas
        Hora | Carregamento_pu

    CurvasPLD
        Curva_PLD | Hora

    Colunas adicionais são incluídas apenas para rastreabilidade do período
    selecionado e não interferem nos códigos atuais, que leem as colunas
    esperadas por nome ou posição.
    """
    if xlwt is None:
        raise ImportError(
            "Para gerar o arquivo dados.xls é necessário instalar a biblioteca "
            "'xlwt'. Execute uma única vez: pip install xlwt"
        )

    if dados_periodo.empty:
        raise ValueError("Não há dados no período selecionado para exportar.")

    caminho_saida = Path(caminho_saida)
    if caminho_saida.parent != Path("."):
        caminho_saida.parent.mkdir(parents=True, exist_ok=True)

    # O formato .xls (BIFF8) possui limite de 65.536 linhas por planilha.
    # As bases anuais horárias utilizadas aqui ficam muito abaixo desse limite.
    if len(dados_periodo) + 1 > 65536:
        raise ValueError(
            "O período selecionado excede o limite de linhas do formato .xls."
        )

    dados_exportacao = dados_periodo.sort_values("data_hora").reset_index(drop=True)

    livro = xlwt.Workbook(encoding="utf-8")
    estilo_data = xlwt.easyxf(num_format_str="dd/mm/yyyy hh:mm")

    # -------------------------------------------------------------------------
    # PLANILHA 1 - CURVA FV / IRRADIÂNCIA
    # -------------------------------------------------------------------------
    planilha_pv = livro.add_sheet("CurvasPV")
    cabecalhos_pv = [
        "Curva_PV_kW",
        "Hora",
        "Data_Hora",
        "Irradiancia_W_m2",
        "Geracao_PV_pu",
    ]
    _escrever_cabecalhos_xls(planilha_pv, cabecalhos_pv)
    _ajustar_larguras_xls(planilha_pv, [16, 8, 20, 20, 16])

    for indice, linha in dados_exportacao.iterrows():
        linha_excel = indice + 1
        data_hora = pd.Timestamp(linha["data_hora"])

        planilha_pv.write(linha_excel, 0, float(linha["potencia_fv_kw"]))
        planilha_pv.write(linha_excel, 1, int(data_hora.hour))
        planilha_pv.write(linha_excel, 2, data_hora.to_pydatetime(), estilo_data)
        planilha_pv.write(
            linha_excel,
            3,
            float(linha["irradiancia_plano_w_m2"]),
        )
        planilha_pv.write(linha_excel, 4, float(linha["geracao_fv_pu"]))

    # -------------------------------------------------------------------------
    # PLANILHA 2 - CURVA DE CARGA
    # -------------------------------------------------------------------------
    planilha_carga = livro.add_sheet("LoadShape48_Cargas")
    cabecalhos_carga = [
        "Hora",
        "Carregamento_pu",
        "Data_Hora",
        "Carga_MWmed",
    ]
    _escrever_cabecalhos_xls(planilha_carga, cabecalhos_carga)
    _ajustar_larguras_xls(planilha_carga, [8, 18, 20, 16])

    for indice, linha in dados_exportacao.iterrows():
        linha_excel = indice + 1
        data_hora = pd.Timestamp(linha["data_hora"])

        planilha_carga.write(linha_excel, 0, int(data_hora.hour))
        planilha_carga.write(linha_excel, 1, float(linha["carga_pu_pico"]))
        planilha_carga.write(
            linha_excel,
            2,
            data_hora.to_pydatetime(),
            estilo_data,
        )
        planilha_carga.write(linha_excel, 3, float(linha["carga_mwmed"]))

    # -------------------------------------------------------------------------
    # PLANILHA 3 - PLD
    # -------------------------------------------------------------------------
    planilha_pld = livro.add_sheet("CurvasPLD")
    cabecalhos_pld = [
        "Curva_PLD",
        "Hora",
        "Data_Hora",
    ]
    _escrever_cabecalhos_xls(planilha_pld, cabecalhos_pld)
    _ajustar_larguras_xls(planilha_pld, [16, 8, 20])

    for indice, linha in dados_exportacao.iterrows():
        linha_excel = indice + 1
        data_hora = pd.Timestamp(linha["data_hora"])

        planilha_pld.write(linha_excel, 0, float(linha["PLD_R$_MWh"]))
        planilha_pld.write(linha_excel, 1, int(data_hora.hour))
        planilha_pld.write(
            linha_excel,
            2,
            data_hora.to_pydatetime(),
            estilo_data,
        )

    livro.save(str(caminho_saida))

    print("\n" + "=" * 78)
    print("ARQUIVO DE DADOS DO PERÍODO")
    print("=" * 78)
    print(f"Arquivo gerado     : {caminho_saida.resolve()}")
    print(f"Registros exportados: {len(dados_exportacao)} horas")
    print("Planilhas           : CurvasPV | LoadShape48_Cargas | CurvasPLD")

    if len(dados_exportacao) == 24:
        print(
            "Compatibilidade     : OK - 24 registros; pronto para uso direto "
            "nos códigos atuais de simulação diária."
        )
    else:
        print(
            "Observação          : o arquivo contém mais/menos de 24 horas. "
            "Os dados foram exportados integralmente, mas os códigos atuais "
            "de simulação diária devem receber um período de 24 horas para "
            "uso direto sem adaptação."
        )


# =============================================================================
# 6. NORMALIZAÇÃO PARA O GRÁFICO CONJUNTO
# =============================================================================

def normalizar_por_maximo(serie: pd.Series, maximo_referencia: float) -> pd.Series:
    """Normaliza uma série para pu usando um valor máximo de referência."""

    if pd.isna(maximo_referencia) or maximo_referencia <= 0:
        raise ValueError("O máximo utilizado na normalização deve ser positivo.")

    return serie / maximo_referencia


def preparar_series_normalizadas(
    dados_periodo: pd.DataFrame,
    dados_anuais: pd.DataFrame,
    modo: str = "anual",
) -> pd.DataFrame:
    """Cria as três séries normalizadas utilizadas no quarto gráfico."""

    modo = modo.lower().strip()

    if modo not in {"anual", "periodo"}:
        raise ValueError("MODO_NORMALIZACAO deve ser 'anual' ou 'periodo'.")

    referencia = dados_anuais if modo == "anual" else dados_periodo

    max_irradiancia = referencia["irradiancia_plano_w_m2"].max()
    max_carga = referencia["carga_mwmed"].max()
    max_pld = referencia["PLD_R$_MWh"].max()

    normalizado = dados_periodo[["data_hora"]].copy()

    normalizado["irradiancia_pu"] = normalizar_por_maximo(
        dados_periodo["irradiancia_plano_w_m2"],
        max_irradiancia,
    )

    normalizado["carga_pu"] = normalizar_por_maximo(
        dados_periodo["carga_mwmed"],
        max_carga,
    )

    normalizado["pld_pu"] = normalizar_por_maximo(
        dados_periodo["PLD_R$_MWh"],
        max_pld,
    )

    return normalizado


# =============================================================================
# 7. FORMATAÇÃO DO EIXO TEMPORAL
# =============================================================================

def configurar_eixo_tempo(ax, dados_periodo: pd.DataFrame) -> None:
    """Escolhe automaticamente uma formatação adequada ao tamanho do período."""

    inicio = dados_periodo["data_hora"].min()
    fim = dados_periodo["data_hora"].max()
    duracao = fim - inicio

    if duracao <= pd.Timedelta(days=2):
        locator = mdates.HourLocator(interval=2)
        formatter = mdates.DateFormatter("%d/%m\n%H:%M")

    elif duracao <= pd.Timedelta(days=14):
        locator = mdates.DayLocator(interval=1)
        formatter = mdates.DateFormatter("%d/%m")

    elif duracao <= pd.Timedelta(days=90):
        locator = mdates.WeekdayLocator(interval=1)
        formatter = mdates.DateFormatter("%d/%m")

    elif duracao <= pd.Timedelta(days=200):
        locator = mdates.MonthLocator(interval=1)
        formatter = mdates.DateFormatter("%b/%Y")

    else:
        locator = mdates.MonthLocator(interval=1)
        formatter = mdates.DateFormatter("%b")

    ax.xaxis.set_major_locator(locator)
    ax.xaxis.set_major_formatter(formatter)
    ax.tick_params(axis="x", rotation=0)


# =============================================================================
# 8. FUNÇÕES DE PLOTAGEM
# =============================================================================

def finalizar_grafico(fig, ax, dados_periodo, titulo, nome_arquivo):
    """Aplica elementos comuns e opcionalmente salva a figura."""

    configurar_eixo_tempo(ax, dados_periodo)
    ax.set_title(titulo)
    ax.set_xlabel("Data / hora")
    ax.grid(True, alpha=0.25)

    fig.tight_layout()

    if SALVAR_GRAFICOS:
        PASTA_SAIDA.mkdir(parents=True, exist_ok=True)
        caminho_saida = PASTA_SAIDA / nome_arquivo
        fig.savefig(caminho_saida, dpi=DPI_SAIDA, bbox_inches="tight")
        print(f"Gráfico salvo: {caminho_saida}")


def plotar_irradiancia(dados_periodo: pd.DataFrame):
    """Gráfico 1 — irradiância horária."""

    fig, ax = plt.subplots(figsize=(13, 5))

    ax.plot(
        dados_periodo["data_hora"],
        dados_periodo["irradiancia_plano_w_m2"],
        linewidth=1.0,
        label="Irradiância",
    )

    ax.set_ylabel("Irradiância no plano [W/m²]")
    ax.set_ylim(bottom=0)
    ax.legend()

    finalizar_grafico(
        fig,
        ax,
        dados_periodo,
        "Irradiância solar horária — PVGIS",
        "01_irradiancia.png",
    )

    return fig


def plotar_carga(dados_periodo: pd.DataFrame):
    """Gráfico 2 — curva de carga do ONS."""

    fig, ax = plt.subplots(figsize=(13, 5))

    ax.plot(
        dados_periodo["data_hora"],
        dados_periodo["carga_mwmed"],
        linewidth=1.0,
        label="Carga SE/CO",
    )

    ax.set_ylabel("Carga [MWmed]")
    ax.set_ylim(bottom=0)
    ax.legend()

    finalizar_grafico(
        fig,
        ax,
        dados_periodo,
        "Curva de carga horária — ONS — Sudeste/Centro-Oeste",
        "02_carga_ons.png",
    )

    return fig


def plotar_pld(dados_periodo: pd.DataFrame):
    """Gráfico 3 — PLD horário."""

    fig, ax = plt.subplots(figsize=(13, 5))

    ax.plot(
        dados_periodo["data_hora"],
        dados_periodo["PLD_R$_MWh"],
        linewidth=1.0,
        label="PLD Sudeste",
    )

    ax.set_ylabel("PLD [R$/MWh]")
    ax.set_ylim(bottom=0)
    ax.legend()

    finalizar_grafico(
        fig,
        ax,
        dados_periodo,
        "Preço de Liquidação das Diferenças — CCEE — Sudeste",
        "03_pld_ccee.png",
    )

    return fig


def plotar_series_conjuntas(
    dados_periodo: pd.DataFrame,
    dados_anuais: pd.DataFrame,
):
    """
    Gráfico 4 — irradiância, carga e PLD no mesmo eixo em valores pu.

    A normalização evita que as diferenças de unidade e ordem de grandeza
    escondam alguma das curvas.
    """

    dados_pu = preparar_series_normalizadas(
        dados_periodo=dados_periodo,
        dados_anuais=dados_anuais,
        modo=MODO_NORMALIZACAO,
    )

    fig, ax = plt.subplots(figsize=(14, 6))

    ax.plot(
        dados_pu["data_hora"],
        dados_pu["irradiancia_pu"],
        linewidth=1.1,
        label="Irradiância [pu]",
    )

    ax.plot(
        dados_pu["data_hora"],
        dados_pu["carga_pu"],
        linewidth=1.1,
        label="Carga [pu]",
    )

    ax.plot(
        dados_pu["data_hora"],
        dados_pu["pld_pu"],
        linewidth=1.1,
        label="PLD [pu]",
    )

    ax.set_ylabel("Valor normalizado [pu]")
    ax.set_ylim(bottom=0)
    ax.legend(ncol=3)

    titulo = (
        "Irradiância × Carga × PLD — séries horárias normalizadas "
        f"(referência: {MODO_NORMALIZACAO})"
    )

    finalizar_grafico(
        fig,
        ax,
        dados_periodo,
        titulo,
        "04_irradiancia_carga_pld_normalizados.png",
    )

    return fig


# =============================================================================
# 9. RESUMO DO PERÍODO SELECIONADO
# =============================================================================

def imprimir_resumo_periodo(dados_periodo: pd.DataFrame) -> None:
    """Exibe alguns valores úteis do período escolhido."""

    print("\n" + "=" * 78)
    print("PERÍODO SELECIONADO")
    print("=" * 78)
    print(f"Início             : {dados_periodo['data_hora'].min()}")
    print(f"Fim                : {dados_periodo['data_hora'].max()}")
    print(f"Horas analisadas   : {len(dados_periodo)}")
    print()
    print(
        "Irradiância máxima: "
        f"{dados_periodo['irradiancia_plano_w_m2'].max():.2f} W/m²"
    )
    print(
        "Carga média        : "
        f"{dados_periodo['carga_mwmed'].mean():.2f} MWmed"
    )
    print(
        "Carga máxima       : "
        f"{dados_periodo['carga_mwmed'].max():.2f} MWmed"
    )
    print(
        "PLD médio          : "
        f"R$ {dados_periodo['PLD_R$_MWh'].mean():.2f}/MWh"
    )
    print(
        "PLD máximo         : "
        f"R$ {dados_periodo['PLD_R$_MWh'].max():.2f}/MWh"
    )


# =============================================================================
# 10. PROGRAMA PRINCIPAL
# =============================================================================

def main():
    # -------------------------------------------------------------------------
    # Leitura
    # -------------------------------------------------------------------------
    df_irradiancia = ler_base_irradiancia(ARQUIVO_IRRADIANCIA)
    df_carga = ler_base_carga(ARQUIVO_CARGA)
    df_pld = ler_base_pld(ARQUIVO_PLD)

    # -------------------------------------------------------------------------
    # Validação
    # -------------------------------------------------------------------------
    validar_base(df_irradiancia, "Irradiância/PVGIS")
    validar_base(df_carga, "Carga/ONS")
    validar_base(df_pld, "PLD/CCEE")

    # -------------------------------------------------------------------------
    # Sincronização
    # -------------------------------------------------------------------------
    dados_anuais = sincronizar_bases(
        df_irradiancia,
        df_carga,
        df_pld,
    )

    imprimir_resumo_sincronizacao(
        df_irradiancia,
        df_carga,
        df_pld,
        dados_anuais,
    )

    # -------------------------------------------------------------------------
    # Recorte temporal
    # -------------------------------------------------------------------------
    dados_periodo = filtrar_periodo(
        dados=dados_anuais,
        data_inicial=DATA_INICIAL,
        data_final=DATA_FINAL,
    )

    imprimir_resumo_periodo(dados_periodo)

    # -------------------------------------------------------------------------
    # Exportação do período selecionado para um único arquivo .xls
    # -------------------------------------------------------------------------
    if GERAR_ARQUIVO_DADOS:
        exportar_dados_periodo_xls(
            dados_periodo=dados_periodo,
            caminho_saida=ARQUIVO_DADOS,
        )

    # -------------------------------------------------------------------------
    # Quatro gráficos
    # -------------------------------------------------------------------------
    plotar_irradiancia(dados_periodo)
    plotar_carga(dados_periodo)
    plotar_pld(dados_periodo)
    plotar_series_conjuntas(dados_periodo, dados_anuais)

    # -------------------------------------------------------------------------
    # Exibição
    # -------------------------------------------------------------------------
    if EXIBIR_GRAFICOS:
        plt.show()
    else:
        plt.close("all")


if __name__ == "__main__":
    main()
