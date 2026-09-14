# ============================================================
# CURVA DE CARGA HORÁRIA REAL — ONS — ANO 2022
#
# Fonte oficial:
# https://ons-aws-prod-opendata.s3.amazonaws.com/
# dataset/curva-carga-ho/CURVA_CARGA_2022.csv
#
# Objetivos:
# 1. Baixar a curva de carga horária oficial do ONS;
# 2. Detectar automaticamente separador, codificação e decimal;
# 3. Padronizar os nomes das colunas;
# 4. Filtrar o subsistema Sudeste/Centro-Oeste;
# 5. Auditar cobertura temporal, duplicidades e valores inválidos;
# 6. Criar curvas normalizadas para aplicação no OpenDSS;
# 7. Gerar resumos anual, mensal, diário e por hora do dia;
# 8. Exportar os resultados para Excel e CSV.
#
# Dependências:
# python -m pip install pandas openpyxl
# ============================================================

from __future__ import annotations

from io import BytesIO, StringIO
from pathlib import Path
import csv
import hashlib
import json
import re
import unicodedata
import urllib.error
import urllib.request

import pandas as pd


# ============================================================
# 1. CONFIGURAÇÕES
# ============================================================

ANO_ANALISE = 2022

# URL oficial do arquivo anual disponibilizado pelo ONS
URL_ARQUIVO_ONS = (
    "https://ons-aws-prod-opendata.s3.amazonaws.com/"
    "dataset/curva-carga-ho/CURVA_CARGA_2022.csv"
)

# ONS normalmente utiliza:
# N, NE, S e SE
#
# Para Sudeste/Centro-Oeste, utilize "SE".
SUBSISTEMA_ANALISE = "SE"

# Nome usado nos arquivos de saída
NOME_SUBSISTEMA_SAIDA = "SUDESTE_CENTRO_OESTE"

# Pasta de saída
PASTA_SAIDA = Path("resultados_curva_carga_ons")

# Arquivo bruto baixado
NOME_ARQUIVO_BRUTO = f"CURVA_CARGA_{ANO_ANALISE}_ONS.csv"

# Arquivos tratados
NOME_BASE_SAIDA = (
    f"Curva_Carga_Horaria_ONS_{ANO_ANALISE}_"
    f"{NOME_SUBSISTEMA_SAIDA}"
)

SALVAR_CSV_TRATADO = True
SALVAR_BASE_COMPLETA = True
SALVAR_ARQUIVO_BRUTO = True

# Tempo máximo da requisição
TIMEOUT_SEGUNDOS = 180

# Em caso de nova publicação do arquivo pelo ONS:
# True  -> sempre baixa novamente;
# False -> usa o arquivo local, caso já exista.
FORCAR_NOVO_DOWNLOAD = False

# Define se qualquer inconsistência temporal encerra o programa.
#
# Para pesquisa e cruzamento com PLD/PVGIS, recomenda-se True.
INTERROMPER_SE_INCONSISTENTE = True

# Quantidade de extremos incluídos nas abas de auditoria
QUANTIDADE_EXTREMOS = 50

# Curva OpenDSS:
# "pico"    -> carga / carga máxima anual
# "media"   -> carga / carga média anual
# "minmax"  -> (carga - mínimo) / (máximo - mínimo)
METODO_NORMALIZACAO_PRINCIPAL = "pico"


# ============================================================
# 2. FUNÇÕES AUXILIARES DE TEXTO
# ============================================================

def remover_acentos(texto: str) -> str:
    """Remove acentos e sinais diacríticos."""
    normalizado = unicodedata.normalize("NFKD", str(texto))
    return "".join(
        caractere
        for caractere in normalizado
        if not unicodedata.combining(caractere)
    )


def normalizar_nome_coluna(nome: str) -> str:
    """
    Padroniza nomes de colunas:
    - remove acentos;
    - converte para minúsculas;
    - troca caracteres especiais por "_";
    - remove "_" repetidos.
    """
    texto = remover_acentos(nome).strip().lower()
    texto = re.sub(r"[^a-z0-9]+", "_", texto)
    texto = re.sub(r"_+", "_", texto)
    return texto.strip("_")


def normalizar_texto(valor: object) -> str:
    """Padroniza textos usados em filtros."""
    if pd.isna(valor):
        return ""

    texto = remover_acentos(str(valor)).strip().upper()
    texto = re.sub(r"\s+", " ", texto)
    return texto


def normalizar_codigo_subsistema(valor: object) -> str:
    """
    Converte diferentes representações para:
    N, NE, S ou SE.
    """
    texto = normalizar_texto(valor)

    equivalencias = {
        "N": "N",
        "NORTE": "N",

        "NE": "NE",
        "NORDESTE": "NE",

        "S": "S",
        "SUL": "S",

        "SE": "SE",
        "SUDESTE": "SE",
        "SUDESTE/CENTRO-OESTE": "SE",
        "SUDESTE / CENTRO-OESTE": "SE",
        "SUDESTE CENTRO-OESTE": "SE",
        "SE/CO": "SE",
        "SE-CO": "SE",
    }

    return equivalencias.get(texto, texto)


# ============================================================
# 3. DOWNLOAD DO ARQUIVO
# ============================================================

def calcular_sha256(conteudo: bytes) -> str:
    """Calcula o hash SHA-256 do arquivo baixado."""
    return hashlib.sha256(conteudo).hexdigest()


def baixar_arquivo_ons(
    url: str,
    caminho_local: Path,
    forcar_download: bool,
) -> tuple[bytes, str]:
    """
    Baixa o arquivo oficial do ONS.

    Retorna
    -------
    conteudo : bytes
        Conteúdo integral do arquivo.

    origem : str
        "download" ou "arquivo_local".
    """

    if caminho_local.exists() and not forcar_download:
        print("Arquivo bruto já existe. Utilizando cópia local.")
        conteudo = caminho_local.read_bytes()
        return conteudo, "arquivo_local"

    print("Baixando arquivo oficial do ONS...")

    requisicao = urllib.request.Request(
        url,
        headers={
            "User-Agent": (
                "Mozilla/5.0 "
                "Pesquisa-Academica-Curva-Carga-ONS"
            ),
            "Accept": (
                "text/csv,text/plain,application/octet-stream,*/*"
            ),
        },
    )

    try:
        with urllib.request.urlopen(
            requisicao,
            timeout=TIMEOUT_SEGUNDOS,
        ) as resposta:
            conteudo = resposta.read()

    except urllib.error.HTTPError as erro:
        mensagem = ""

        try:
            mensagem = erro.read().decode(
                "utf-8",
                errors="replace",
            )
        except Exception:
            mensagem = str(erro)

        raise RuntimeError(
            "\nErro HTTP durante o download do ONS.\n"
            f"Código: {erro.code}\n"
            f"Mensagem: {mensagem[:2000]}"
        ) from erro

    except urllib.error.URLError as erro:
        raise RuntimeError(
            "\nNão foi possível acessar o servidor do ONS.\n"
            f"Motivo: {erro.reason}"
        ) from erro

    except TimeoutError as erro:
        raise RuntimeError(
            "O download excedeu o tempo máximo."
        ) from erro

    if not conteudo:
        raise RuntimeError(
            "O arquivo baixado está vazio."
        )

    caminho_local.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    if SALVAR_ARQUIVO_BRUTO:
        caminho_local.write_bytes(conteudo)

    return conteudo, "download"


# ============================================================
# 4. DETECÇÃO DA CODIFICAÇÃO E DO FORMATO CSV
# ============================================================

def detectar_codificacao(conteudo: bytes) -> str:
    """
    Testa codificações comuns dos arquivos públicos brasileiros.
    """

    candidatas = [
        "utf-8-sig",
        "utf-8",
        "latin-1",
        "cp1252",
    ]

    for codificacao in candidatas:
        try:
            conteudo.decode(codificacao)
            return codificacao
        except UnicodeDecodeError:
            continue

    raise RuntimeError(
        "Não foi possível identificar a codificação do arquivo."
    )


def detectar_separador(texto: str) -> str:
    """
    Detecta o delimitador por csv.Sniffer e por contagem,
    com suporte a ';', ',' e tabulação.
    """

    amostra = texto[:10000]

    try:
        dialecto = csv.Sniffer().sniff(
            amostra,
            delimiters=";,\t|",
        )
        return dialecto.delimiter

    except csv.Error:
        primeira_linha = amostra.splitlines()[0]

        contagens = {
            ";": primeira_linha.count(";"),
            ",": primeira_linha.count(","),
            "\t": primeira_linha.count("\t"),
            "|": primeira_linha.count("|"),
        }

        separador = max(
            contagens,
            key=contagens.get,
        )

        if contagens[separador] == 0:
            raise RuntimeError(
                "Não foi possível identificar o separador do CSV."
            )

        return separador


def ler_csv_ons(
    conteudo: bytes,
) -> tuple[pd.DataFrame, dict]:
    """
    Lê o CSV com detecção automática de:
    - codificação;
    - separador;
    - tipo de decimal.
    """

    codificacao = detectar_codificacao(conteudo)
    texto = conteudo.decode(codificacao)
    separador = detectar_separador(texto)

    # Primeira leitura como texto evita perda de informação.
    df = pd.read_csv(
        StringIO(texto),
        sep=separador,
        dtype=str,
        keep_default_na=False,
        na_values=[],
        engine="python",
    )

    metadados_leitura = {
        "codificacao_detectada": codificacao,
        "separador_detectado": repr(separador),
        "quantidade_linhas_brutas": len(df),
        "quantidade_colunas_brutas": len(df.columns),
        "colunas_originais": json.dumps(
            df.columns.tolist(),
            ensure_ascii=False,
        ),
    }

    return df, metadados_leitura


# ============================================================
# 5. IDENTIFICAÇÃO DAS COLUNAS
# ============================================================

def localizar_coluna(
    colunas: list[str],
    aliases: list[str],
    descricao: str,
) -> str:
    """
    Localiza uma coluna por lista de aliases normalizados.
    """

    mapa = {
        normalizar_nome_coluna(coluna): coluna
        for coluna in colunas
    }

    for alias in aliases:
        alias_normalizado = normalizar_nome_coluna(alias)

        if alias_normalizado in mapa:
            return mapa[alias_normalizado]

    raise RuntimeError(
        f"Não foi possível localizar a coluna de {descricao}.\n"
        f"Aliases procurados: {aliases}\n"
        f"Colunas encontradas: {colunas}"
    )


def identificar_colunas_ons(
    df: pd.DataFrame,
) -> dict[str, str]:
    """
    Identifica as colunas centrais do conjunto Curva de Carga.
    """

    colunas = df.columns.tolist()

    coluna_id = localizar_coluna(
        colunas,
        aliases=[
            "id_subsistema",
            "idsubsistema",
            "cod_subsistema",
            "codigo_subsistema",
        ],
        descricao="identificador do subsistema",
    )

    coluna_nome = localizar_coluna(
        colunas,
        aliases=[
            "nom_subsistema",
            "nome_subsistema",
            "subsistema",
        ],
        descricao="nome do subsistema",
    )

    coluna_data_hora = localizar_coluna(
        colunas,
        aliases=[
            "din_instante",
            "data_hora",
            "datahora",
            "instante",
            "datetime",
        ],
        descricao="data e hora",
    )

    coluna_carga = localizar_coluna(
        colunas,
        aliases=[
            # Nome usado na Curva de Carga Horária
            "val_cargaenergiahomwmed",

            # Variações observadas ou plausíveis
            "val_cargaenergiahora_mwmed",
            "val_cargaenergiahora",
            "val_cargaenergiamwmed",
            "carga_mwmed",
            "carga_mw",
            "valor_carga",
        ],
        descricao="valor da carga horária",
    )

    return {
        "id_subsistema": coluna_id,
        "nom_subsistema": coluna_nome,
        "data_hora": coluna_data_hora,
        "carga": coluna_carga,
    }


# ============================================================
# 6. CONVERSÃO NUMÉRICA ROBUSTA
# ============================================================

def converter_numero_robusto(
    serie: pd.Series,
) -> pd.Series:
    """
    Converte números nos formatos:
    33108.8615
    33108,8615
    33.108,8615
    33,108.8615
    """

    def converter_valor(valor: object) -> float | None:
        if pd.isna(valor):
            return None

        texto = str(valor).strip()

        if texto == "":
            return None

        texto = (
            texto
            .replace("R$", "")
            .replace("MWmed", "")
            .replace("MW", "")
            .replace(" ", "")
        )

        possui_ponto = "." in texto
        possui_virgula = "," in texto

        if possui_ponto and possui_virgula:
            # O último símbolo é tratado como separador decimal.
            if texto.rfind(",") > texto.rfind("."):
                texto = texto.replace(".", "")
                texto = texto.replace(",", ".")
            else:
                texto = texto.replace(",", "")

        elif possui_virgula:
            # Vírgula isolada é tratada como decimal.
            texto = texto.replace(",", ".")

        try:
            return float(texto)

        except ValueError:
            return None

    return serie.apply(converter_valor)


# ============================================================
# 7. TRATAMENTO DA BASE
# ============================================================

def tratar_base_ons(
    df_bruto: pd.DataFrame,
) -> tuple[pd.DataFrame, dict]:
    """
    Padroniza e trata a base completa do ONS.
    """

    colunas = identificar_colunas_ons(df_bruto)

    df = pd.DataFrame(
        {
            "id_subsistema_original": (
                df_bruto[colunas["id_subsistema"]]
                .astype(str)
                .str.strip()
            ),
            "nom_subsistema_original": (
                df_bruto[colunas["nom_subsistema"]]
                .astype(str)
                .str.strip()
            ),
            "data_hora_original": (
                df_bruto[colunas["data_hora"]]
                .astype(str)
                .str.strip()
            ),
            "carga_original": (
                df_bruto[colunas["carga"]]
                .astype(str)
                .str.strip()
            ),
        }
    )

    # Subsistema
    df["id_subsistema"] = (
        df["id_subsistema_original"]
        .apply(normalizar_codigo_subsistema)
    )

    df["nom_subsistema"] = (
        df["nom_subsistema_original"]
        .apply(normalizar_texto)
    )

    # Data e hora
    #
    # A primeira tentativa usa formato misto.
    df["data_hora"] = pd.to_datetime(
        df["data_hora_original"],
        format="mixed",
        dayfirst=False,
        errors="coerce",
    )

    # Segunda tentativa, com dayfirst=True, apenas nas falhas.
    mascara_invalida = df["data_hora"].isna()

    if mascara_invalida.any():
        df.loc[
            mascara_invalida,
            "data_hora",
        ] = pd.to_datetime(
            df.loc[
                mascara_invalida,
                "data_hora_original",
            ],
            format="mixed",
            dayfirst=True,
            errors="coerce",
        )

    # Carga
    df["carga_mwmed"] = converter_numero_robusto(
        df["carga_original"]
    )

    # Auditoria inicial
    datas_invalidas = int(
        df["data_hora"].isna().sum()
    )

    cargas_invalidas = int(
        df["carga_mwmed"].isna().sum()
    )

    if datas_invalidas > 0:
        exemplos = (
            df.loc[
                df["data_hora"].isna(),
                "data_hora_original",
            ]
            .drop_duplicates()
            .head(20)
            .tolist()
        )

        raise RuntimeError(
            f"Foram encontradas {datas_invalidas} datas inválidas.\n"
            f"Exemplos: {exemplos}"
        )

    if cargas_invalidas > 0:
        exemplos = (
            df.loc[
                df["carga_mwmed"].isna(),
                "carga_original",
            ]
            .drop_duplicates()
            .head(20)
            .tolist()
        )

        raise RuntimeError(
            f"Foram encontradas {cargas_invalidas} cargas inválidas.\n"
            f"Exemplos: {exemplos}"
        )

    # Colunas auxiliares
    df["ano"] = df["data_hora"].dt.year
    df["mes"] = df["data_hora"].dt.month
    df["dia"] = df["data_hora"].dt.day
    df["hora"] = df["data_hora"].dt.hour
    df["dia_do_ano"] = df["data_hora"].dt.dayofyear
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

    df["data"] = df["data_hora"].dt.normalize()

    df["chave_data_hora"] = (
        df["data_hora"]
        .dt.strftime("%Y-%m-%d %H:00:00")
    )

    df = (
        df.sort_values(
            ["data_hora", "id_subsistema"]
        )
        .reset_index(drop=True)
    )

    diagnostico = {
        "coluna_id_subsistema": colunas["id_subsistema"],
        "coluna_nome_subsistema": colunas["nom_subsistema"],
        "coluna_data_hora": colunas["data_hora"],
        "coluna_carga": colunas["carga"],
        "datas_invalidas": datas_invalidas,
        "cargas_invalidas": cargas_invalidas,
    }

    return df, diagnostico


# ============================================================
# 8. FILTRO DO ANO E SUBSISTEMA
# ============================================================

def filtrar_serie(
    df: pd.DataFrame,
    ano: int,
    subsistema: str,
) -> pd.DataFrame:
    """
    Filtra o ano e o subsistema solicitados.
    """

    subsistema_normalizado = normalizar_codigo_subsistema(
        subsistema
    )

    subsistemas_disponiveis = sorted(
        df["id_subsistema"]
        .dropna()
        .unique()
        .tolist()
    )

    print("=" * 78)
    print("SUBSISTEMAS ENCONTRADOS")
    print("=" * 78)
    print(subsistemas_disponiveis)
    print()

    if subsistema_normalizado not in subsistemas_disponiveis:
        raise RuntimeError(
            f"O subsistema '{subsistema_normalizado}' "
            "não foi encontrado.\n"
            f"Disponíveis: {subsistemas_disponiveis}"
        )

    filtro = (
        (df["ano"] == ano)
        & (df["id_subsistema"] == subsistema_normalizado)
    )

    resultado = (
        df.loc[filtro]
        .copy()
        .sort_values("data_hora")
        .reset_index(drop=True)
    )

    if resultado.empty:
        raise RuntimeError(
            "Nenhum registro permaneceu após o filtro."
        )

    return resultado


# ============================================================
# 9. CURVAS NORMALIZADAS
# ============================================================

def adicionar_curvas_normalizadas(
    df: pd.DataFrame,
) -> pd.DataFrame:
    """
    Adiciona três formas de normalização.

    carga_pu_pico:
        carga / pico anual

    carga_pu_media:
        carga / média anual

    carga_pu_minmax:
        (carga - mínimo) / (máximo - mínimo)
    """

    df = df.copy()

    carga_maxima = df["carga_mwmed"].max()
    carga_minima = df["carga_mwmed"].min()
    carga_media = df["carga_mwmed"].mean()

    if carga_maxima <= 0:
        raise RuntimeError(
            "A carga máxima não é positiva."
        )

    if carga_media <= 0:
        raise RuntimeError(
            "A carga média não é positiva."
        )

    df["carga_pu_pico"] = (
        df["carga_mwmed"] / carga_maxima
    )

    df["carga_pu_media"] = (
        df["carga_mwmed"] / carga_media
    )

    amplitude = carga_maxima - carga_minima

    if amplitude > 0:
        df["carga_pu_minmax"] = (
            (df["carga_mwmed"] - carga_minima)
            / amplitude
        )
    else:
        df["carga_pu_minmax"] = 0.0

    metodos = {
        "pico": "carga_pu_pico",
        "media": "carga_pu_media",
        "minmax": "carga_pu_minmax",
    }

    if METODO_NORMALIZACAO_PRINCIPAL not in metodos:
        raise RuntimeError(
            "METODO_NORMALIZACAO_PRINCIPAL inválido.\n"
            f"Use uma destas opções: {list(metodos)}"
        )

    df["loadshape_opendss_pu"] = (
        df[metodos[METODO_NORMALIZACAO_PRINCIPAL]]
    )

    return df


# ============================================================
# 10. AUDITORIA TEMPORAL E FÍSICA
# ============================================================

def auditar_serie_horaria(
    df: pd.DataFrame,
    ano: int,
) -> tuple[pd.DataFrame, dict]:
    """
    Verifica:
    - 8.760 ou 8.784 horas;
    - primeiro e último horário;
    - duplicidades;
    - lacunas;
    - valores nulos;
    - cargas não positivas;
    - frequência horária;
    - 24 registros por dia.
    """

    inicio = pd.Timestamp(
        year=ano,
        month=1,
        day=1,
        hour=0,
    )

    fim = pd.Timestamp(
        year=ano,
        month=12,
        day=31,
        hour=23,
    )

    indice_esperado = pd.date_range(
        start=inicio,
        end=fim,
        freq="h",
    )

    indice_obtido = pd.DatetimeIndex(
        df["data_hora"]
    )

    horarios_ausentes = indice_esperado.difference(
        indice_obtido
    )

    horarios_excedentes = indice_obtido.difference(
        indice_esperado
    )

    quantidade_esperada = len(indice_esperado)
    quantidade_obtida = len(df)

    duplicados = int(
        df["data_hora"].duplicated().sum()
    )

    valores_nulos = int(
        df["carga_mwmed"].isna().sum()
    )

    valores_negativos = int(
        (df["carga_mwmed"] < 0).sum()
    )

    valores_iguais_zero = int(
        (df["carga_mwmed"] == 0).sum()
    )

    # Verifica saltos diferentes de uma hora
    diferencas = (
        df["data_hora"]
        .sort_values()
        .diff()
        .dropna()
    )

    intervalos_irregulares = int(
        (diferencas != pd.Timedelta(hours=1)).sum()
    )

    # Verificação diária
    contagem_diaria = (
        df.groupby("data")
        .size()
        .rename("horas_obtidas")
        .reset_index()
    )

    contagem_diaria["horas_esperadas"] = 24
    contagem_diaria["diferenca"] = (
        contagem_diaria["horas_obtidas"]
        - contagem_diaria["horas_esperadas"]
    )

    contagem_diaria["status"] = (
        contagem_diaria["diferenca"]
        .apply(
            lambda valor: (
                "OK" if valor == 0 else "INCONSISTENTE"
            )
        )
    )

    dias_incompletos = int(
        (contagem_diaria["status"] != "OK").sum()
    )

    consistente = (
        quantidade_obtida == quantidade_esperada
        and duplicados == 0
        and len(horarios_ausentes) == 0
        and len(horarios_excedentes) == 0
        and valores_nulos == 0
        and valores_negativos == 0
        and valores_iguais_zero == 0
        and intervalos_irregulares == 0
        and dias_incompletos == 0
    )

    diagnostico = {
        "ano": ano,
        "registros_esperados": quantidade_esperada,
        "registros_obtidos": quantidade_obtida,
        "primeiro_registro": df["data_hora"].min(),
        "ultimo_registro": df["data_hora"].max(),
        "horarios_duplicados": duplicados,
        "horarios_ausentes": len(horarios_ausentes),
        "horarios_excedentes": len(horarios_excedentes),
        "valores_nulos": valores_nulos,
        "valores_negativos": valores_negativos,
        "valores_iguais_zero": valores_iguais_zero,
        "intervalos_irregulares": intervalos_irregulares,
        "dias_incompletos": dias_incompletos,
        "serie_consistente": consistente,
        "primeiros_horarios_ausentes": (
            [str(valor) for valor in horarios_ausentes[:20]]
        ),
        "primeiros_horarios_excedentes": (
            [str(valor) for valor in horarios_excedentes[:20]]
        ),
    }

    print("=" * 78)
    print("AUDITORIA DA SÉRIE HORÁRIA")
    print("=" * 78)

    for chave, valor in diagnostico.items():
        if chave not in {
            "primeiros_horarios_ausentes",
            "primeiros_horarios_excedentes",
        }:
            print(f"{chave:30}: {valor}")

    print()

    if consistente:
        print(
            "A série horária está completa e consistente."
        )
    else:
        print(
            "ATENÇÃO: a série apresentou inconsistências."
        )

        if diagnostico["primeiros_horarios_ausentes"]:
            print("\nPrimeiros horários ausentes:")
            for valor in diagnostico[
                "primeiros_horarios_ausentes"
            ]:
                print(f"  - {valor}")

        if diagnostico["primeiros_horarios_excedentes"]:
            print("\nPrimeiros horários excedentes:")
            for valor in diagnostico[
                "primeiros_horarios_excedentes"
            ]:
                print(f"  - {valor}")

    print()

    if INTERROMPER_SE_INCONSISTENTE and not consistente:
        raise RuntimeError(
            "A execução foi interrompida porque a série "
            "horária não passou na auditoria."
        )

    return contagem_diaria, diagnostico


# ============================================================
# 11. VERIFICAÇÃO MENSAL
# ============================================================

def criar_verificacao_mensal(
    df: pd.DataFrame,
    ano: int,
) -> pd.DataFrame:
    """Compara as horas obtidas e esperadas em cada mês."""

    obtido = (
        df.groupby("mes")
        .size()
        .rename("horas_obtidas")
        .reset_index()
    )

    esperado = []

    for mes in range(1, 13):
        inicio_mes = pd.Timestamp(
            year=ano,
            month=mes,
            day=1,
        )

        if mes == 12:
            inicio_seguinte = pd.Timestamp(
                year=ano + 1,
                month=1,
                day=1,
            )
        else:
            inicio_seguinte = pd.Timestamp(
                year=ano,
                month=mes + 1,
                day=1,
            )

        horas = int(
            (
                inicio_seguinte - inicio_mes
            ).total_seconds()
            / 3600
        )

        esperado.append(
            {
                "mes": mes,
                "horas_esperadas": horas,
            }
        )

    verificacao = pd.merge(
        pd.DataFrame(esperado),
        obtido,
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
                "OK" if valor == 0 else "INCONSISTENTE"
            )
        )
    )

    return verificacao[
        [
            "mes",
            "nome_mes",
            "horas_esperadas",
            "horas_obtidas",
            "diferenca",
            "status",
        ]
    ]


# ============================================================
# 12. RESUMOS ESTATÍSTICOS
# ============================================================

def criar_resumo_anual(
    df: pd.DataFrame,
) -> pd.DataFrame:
    """Cria indicadores anuais da curva de carga."""

    indice_minimo = df["carga_mwmed"].idxmin()
    indice_maximo = df["carga_mwmed"].idxmax()

    linha_minimo = df.loc[indice_minimo]
    linha_maximo = df.loc[indice_maximo]

    media = df["carga_mwmed"].mean()
    desvio = df["carga_mwmed"].std()

    fator_carga = (
        media / linha_maximo["carga_mwmed"]
        if linha_maximo["carga_mwmed"] > 0
        else pd.NA
    )

    energia_anual_gwh = (
        df["carga_mwmed"].sum() / 1000
    )

    return pd.DataFrame(
        {
            "Indicador": [
                "Ano",
                "Subsistema",
                "Registros horários",
                "Carga média [MWmed]",
                "Carga mediana [MWmed]",
                "Carga mínima [MWmed]",
                "Data/hora da carga mínima",
                "Carga máxima [MWmed]",
                "Data/hora da carga máxima",
                "Desvio-padrão [MWmed]",
                "Primeiro quartil [MWmed]",
                "Terceiro quartil [MWmed]",
                "Fator de carga médio/pico",
                "Energia anual aproximada [GWh]",
                "Método da curva principal OpenDSS",
            ],
            "Valor": [
                ANO_ANALISE,
                SUBSISTEMA_ANALISE,
                len(df),
                media,
                df["carga_mwmed"].median(),
                linha_minimo["carga_mwmed"],
                linha_minimo["data_hora"],
                linha_maximo["carga_mwmed"],
                linha_maximo["data_hora"],
                desvio,
                df["carga_mwmed"].quantile(0.25),
                df["carga_mwmed"].quantile(0.75),
                fator_carga,
                energia_anual_gwh,
                METODO_NORMALIZACAO_PRINCIPAL,
            ],
        }
    )


def criar_resumo_mensal(
    df: pd.DataFrame,
) -> pd.DataFrame:
    """Cria indicadores mensais."""

    resumo = (
        df.groupby(
            ["ano", "mes"],
            as_index=False,
        )
        .agg(
            quantidade_horas=(
                "carga_mwmed",
                "size",
            ),
            carga_media_mwmed=(
                "carga_mwmed",
                "mean",
            ),
            carga_mediana_mwmed=(
                "carga_mwmed",
                "median",
            ),
            carga_minima_mwmed=(
                "carga_mwmed",
                "min",
            ),
            carga_maxima_mwmed=(
                "carga_mwmed",
                "max",
            ),
            desvio_padrao_mwmed=(
                "carga_mwmed",
                "std",
            ),
            energia_mensal_gwh=(
                "carga_mwmed",
                lambda serie: serie.sum() / 1000,
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
        resumo["mes"].map(nomes_meses)
    )

    return resumo[
        [
            "ano",
            "mes",
            "nome_mes",
            "quantidade_horas",
            "carga_media_mwmed",
            "carga_mediana_mwmed",
            "carga_minima_mwmed",
            "carga_maxima_mwmed",
            "desvio_padrao_mwmed",
            "energia_mensal_gwh",
        ]
    ]


def criar_resumo_diario(
    df: pd.DataFrame,
) -> pd.DataFrame:
    """Cria indicadores diários."""

    resumo = (
        df.groupby(
            "data",
            as_index=False,
        )
        .agg(
            quantidade_horas=(
                "carga_mwmed",
                "size",
            ),
            carga_media_mwmed=(
                "carga_mwmed",
                "mean",
            ),
            carga_minima_mwmed=(
                "carga_mwmed",
                "min",
            ),
            carga_maxima_mwmed=(
                "carga_mwmed",
                "max",
            ),
            amplitude_diaria_mw=(
                "carga_mwmed",
                lambda serie: (
                    serie.max() - serie.min()
                ),
            ),
            desvio_padrao_mwmed=(
                "carga_mwmed",
                "std",
            ),
            energia_diaria_gwh=(
                "carga_mwmed",
                lambda serie: serie.sum() / 1000,
            ),
        )
    )

    resumo["serie_diaria_completa"] = (
        resumo["quantidade_horas"] == 24
    )

    return resumo.sort_values(
        "data"
    ).reset_index(drop=True)


def criar_perfil_medio_horario(
    df: pd.DataFrame,
) -> pd.DataFrame:
    """Calcula o perfil estatístico médio das 24 horas."""

    return (
        df.groupby(
            "hora",
            as_index=False,
        )
        .agg(
            quantidade_registros=(
                "carga_mwmed",
                "size",
            ),
            carga_media_mwmed=(
                "carga_mwmed",
                "mean",
            ),
            carga_mediana_mwmed=(
                "carga_mwmed",
                "median",
            ),
            carga_minima_mwmed=(
                "carga_mwmed",
                "min",
            ),
            carga_maxima_mwmed=(
                "carga_mwmed",
                "max",
            ),
            desvio_padrao_mwmed=(
                "carga_mwmed",
                "std",
            ),
            loadshape_media_pu=(
                "carga_pu_pico",
                "mean",
            ),
        )
    )


def identificar_extremos(
    df: pd.DataFrame,
    quantidade: int,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Seleciona os maiores e menores valores horários."""

    colunas = [
        "data_hora",
        "id_subsistema",
        "nom_subsistema",
        "carga_mwmed",
        "carga_pu_pico",
        "mes",
        "dia",
        "hora",
        "dia_da_semana",
    ]

    maiores = (
        df[colunas]
        .sort_values(
            "carga_mwmed",
            ascending=False,
        )
        .head(quantidade)
        .reset_index(drop=True)
    )

    menores = (
        df[colunas]
        .sort_values(
            "carga_mwmed",
            ascending=True,
        )
        .head(quantidade)
        .reset_index(drop=True)
    )

    return maiores, menores


# ============================================================
# 13. ORGANIZAÇÃO DAS COLUNAS
# ============================================================

def organizar_colunas(
    df: pd.DataFrame,
) -> pd.DataFrame:
    """Organiza as colunas principais."""

    prioritarias = [
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

        "id_subsistema",
        "nom_subsistema",

        "carga_mwmed",
        "carga_pu_pico",
        "carga_pu_media",
        "carga_pu_minmax",
        "loadshape_opendss_pu",

        "id_subsistema_original",
        "nom_subsistema_original",
        "data_hora_original",
        "carga_original",
    ]

    existentes = [
        coluna
        for coluna in prioritarias
        if coluna in df.columns
    ]

    restantes = [
        coluna
        for coluna in df.columns
        if coluna not in existentes
    ]

    return df[existentes + restantes].copy()


# ============================================================
# 14. RELATÓRIO DE METADADOS E AUDITORIA
# ============================================================

def criar_tabela_metadados(
    metadados_leitura: dict,
    diagnostico_tratamento: dict,
    diagnostico_auditoria: dict,
    origem: str,
    sha256: str,
) -> pd.DataFrame:
    """Consolida metadados de rastreabilidade."""

    dados = {
        "Fonte": "ONS — Dados Abertos",
        "URL": URL_ARQUIVO_ONS,
        "Ano": ANO_ANALISE,
        "Subsistema filtrado": SUBSISTEMA_ANALISE,
        "Origem da leitura": origem,
        "SHA256 do arquivo": sha256,
        "Método de normalização principal": (
            METODO_NORMALIZACAO_PRINCIPAL
        ),
    }

    dados.update(metadados_leitura)
    dados.update(diagnostico_tratamento)
    dados.update(diagnostico_auditoria)

    linhas = []

    for chave, valor in dados.items():
        if isinstance(valor, (list, dict)):
            valor = json.dumps(
                valor,
                ensure_ascii=False,
                default=str,
            )

        linhas.append(
            {
                "Campo": chave,
                "Valor": valor,
            }
        )

    return pd.DataFrame(linhas)


# ============================================================
# 15. EXPORTAÇÃO
# ============================================================

def salvar_resultados(
    df_subsistema: pd.DataFrame,
    df_base_completa: pd.DataFrame,
    resumo_anual: pd.DataFrame,
    resumo_mensal: pd.DataFrame,
    resumo_diario: pd.DataFrame,
    perfil_horario: pd.DataFrame,
    verificacao_mensal: pd.DataFrame,
    verificacao_diaria: pd.DataFrame,
    maiores_cargas: pd.DataFrame,
    menores_cargas: pd.DataFrame,
    metadados: pd.DataFrame,
) -> tuple[Path, Path | None]:
    """Salva Excel e CSV."""

    PASTA_SAIDA.mkdir(
        parents=True,
        exist_ok=True,
    )

    caminho_excel = (
        PASTA_SAIDA / f"{NOME_BASE_SAIDA}.xlsx"
    )

    caminho_csv = (
        PASTA_SAIDA / f"{NOME_BASE_SAIDA}.csv"
    )

    with pd.ExcelWriter(
        caminho_excel,
        engine="openpyxl",
    ) as escritor:

        df_subsistema.to_excel(
            escritor,
            sheet_name="Carga_Horaria_SE",
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

        verificacao_mensal.to_excel(
            escritor,
            sheet_name="Verificacao_Mensal",
            index=False,
        )

        verificacao_diaria.to_excel(
            escritor,
            sheet_name="Verificacao_Diaria",
            index=False,
        )

        maiores_cargas.to_excel(
            escritor,
            sheet_name="Maiores_Cargas",
            index=False,
        )

        menores_cargas.to_excel(
            escritor,
            sheet_name="Menores_Cargas",
            index=False,
        )

        metadados.to_excel(
            escritor,
            sheet_name="Metadados_Auditoria",
            index=False,
        )

        if SALVAR_BASE_COMPLETA:
            df_base_completa.to_excel(
                escritor,
                sheet_name="Todos_Subssistemas",
                index=False,
            )

    caminho_csv_gerado = None

    if SALVAR_CSV_TRATADO:
        df_subsistema.to_csv(
            caminho_csv,
            index=False,
            sep=";",
            decimal=",",
            encoding="utf-8-sig",
        )
        caminho_csv_gerado = caminho_csv

    return caminho_excel, caminho_csv_gerado


# ============================================================
# 16. FUNÇÃO PRINCIPAL
# ============================================================

def executar() -> None:
    """Executa o fluxo completo."""

    print("=" * 78)
    print("CURVA DE CARGA HORÁRIA REAL — ONS")
    print("=" * 78)
    print(f"Ano:                       {ANO_ANALISE}")
    print(f"Subsistema:                {SUBSISTEMA_ANALISE}")
    print(f"Normalização principal:    {METODO_NORMALIZACAO_PRINCIPAL}")
    print()

    PASTA_SAIDA.mkdir(
        parents=True,
        exist_ok=True,
    )

    caminho_bruto = (
        PASTA_SAIDA / NOME_ARQUIVO_BRUTO
    )

    conteudo, origem = baixar_arquivo_ons(
        url=URL_ARQUIVO_ONS,
        caminho_local=caminho_bruto,
        forcar_download=FORCAR_NOVO_DOWNLOAD,
    )

    sha256 = calcular_sha256(conteudo)

    print(f"Tamanho do arquivo:        {len(conteudo):,} bytes")
    print(f"SHA-256:                   {sha256}")
    print()

    df_bruto, metadados_leitura = ler_csv_ons(
        conteudo
    )

    print("=" * 78)
    print("LEITURA DO CSV")
    print("=" * 78)
    print(f"Linhas:                    {len(df_bruto)}")
    print(f"Colunas:                   {len(df_bruto.columns)}")
    print(f"Codificação:               {metadados_leitura['codificacao_detectada']}")
    print(f"Separador:                 {metadados_leitura['separador_detectado']}")
    print(f"Nomes das colunas:         {df_bruto.columns.tolist()}")
    print()

    df_completo, diagnostico_tratamento = (
        tratar_base_ons(df_bruto)
    )

    print("=" * 78)
    print("PERÍODO DA BASE COMPLETA")
    print("=" * 78)
    print(f"Primeiro instante:         {df_completo['data_hora'].min()}")
    print(f"Último instante:           {df_completo['data_hora'].max()}")
    print(f"Registros tratados:        {len(df_completo)}")
    print()

    df_carga = filtrar_serie(
        df=df_completo,
        ano=ANO_ANALISE,
        subsistema=SUBSISTEMA_ANALISE,
    )

    df_carga = adicionar_curvas_normalizadas(
        df_carga
    )

    verificacao_diaria, diagnostico_auditoria = (
        auditar_serie_horaria(
            df=df_carga,
            ano=ANO_ANALISE,
        )
    )

    verificacao_mensal = criar_verificacao_mensal(
        df=df_carga,
        ano=ANO_ANALISE,
    )

    resumo_anual = criar_resumo_anual(
        df_carga
    )

    resumo_mensal = criar_resumo_mensal(
        df_carga
    )

    resumo_diario = criar_resumo_diario(
        df_carga
    )

    perfil_horario = criar_perfil_medio_horario(
        df_carga
    )

    maiores_cargas, menores_cargas = (
        identificar_extremos(
            df=df_carga,
            quantidade=QUANTIDADE_EXTREMOS,
        )
    )

    df_carga = organizar_colunas(df_carga)
    df_completo = organizar_colunas(df_completo)

    metadados = criar_tabela_metadados(
        metadados_leitura=metadados_leitura,
        diagnostico_tratamento=diagnostico_tratamento,
        diagnostico_auditoria=diagnostico_auditoria,
        origem=origem,
        sha256=sha256,
    )

    print("=" * 78)
    print("PRIMEIROS REGISTROS")
    print("=" * 78)
    print(
        df_carga[
            [
                "data_hora",
                "id_subsistema",
                "carga_mwmed",
                "loadshape_opendss_pu",
            ]
        ].head(10).to_string(index=False)
    )
    print()

    print("=" * 78)
    print("ÚLTIMOS REGISTROS")
    print("=" * 78)
    print(
        df_carga[
            [
                "data_hora",
                "id_subsistema",
                "carga_mwmed",
                "loadshape_opendss_pu",
            ]
        ].tail(10).to_string(index=False)
    )
    print()

    print("=" * 78)
    print("RESUMO ANUAL")
    print("=" * 78)
    print(resumo_anual.to_string(index=False))
    print()

    caminho_excel, caminho_csv = salvar_resultados(
        df_subsistema=df_carga,
        df_base_completa=df_completo,
        resumo_anual=resumo_anual,
        resumo_mensal=resumo_mensal,
        resumo_diario=resumo_diario,
        perfil_horario=perfil_horario,
        verificacao_mensal=verificacao_mensal,
        verificacao_diaria=verificacao_diaria,
        maiores_cargas=maiores_cargas,
        menores_cargas=menores_cargas,
        metadados=metadados,
    )

    print("=" * 78)
    print("ARQUIVOS GERADOS")
    print("=" * 78)
    print(f"Excel: {caminho_excel.resolve()}")

    if caminho_csv is not None:
        print(f"CSV:   {caminho_csv.resolve()}")

    if SALVAR_ARQUIVO_BRUTO:
        print(f"Bruto: {caminho_bruto.resolve()}")

    print()
    print("=" * 78)
    print("PROCESSAMENTO CONCLUÍDO")
    print("=" * 78)


# ============================================================
# 17. INÍCIO DO PROGRAMA
# ============================================================

if __name__ == "__main__":
    executar()