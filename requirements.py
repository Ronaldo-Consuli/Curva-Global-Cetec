"""
Carta controle NIR DS-3 — Curvas Globais — comparação NIR x Método Convencional (Ref)
=================================================================================

Estrutura esperada da planilha:
- Coluna A: Tempo de análise
- Coluna B: Nome do produto
- Coluna E: Número de amostra
- A partir da coluna I: pares de ensaio NIR / "Ref <ensaio>"
- Diferença utilizada em todos os indicadores: REF - NIR

O app:
1. Filtra produto, ensaio e verificação.
2. Cria uma sequência de Verificação a cada 30 pares por Produto + Ensaio sem renumerar os dados.
3. Aplica os limites da planilha "Limites curva global.xlsx" por produto/ensaio;
   nos demais casos, calcula LIC/LSC por Carta I (amplitude móvel).
4. Exibe os KPIs existentes e uma segunda linha centralizada com os cartões Bias Atuais, Ajustes e Novas Bias.
5. Gera a seção Carta Controle com LIC/LSC tracejados, sem o selo do método de limites.
6. Exibe boletim analítico simplificado com Id, Data/hora, Ref, NIR,
   Ref - NIR, LIC, LSC e Status, com cabeçalhos e células centralizados.
7. Não exibe informações de estrutura detectada nem rodapé informativo no app.
8. Exporta um HTML autônomo e interativo com filtros de Produto, Ensaio e Verificação.
9. Mantém na barra lateral somente o campo para importar até duas planilhas Excel.
10. Consolida automaticamente 1 ou 2 arquivos compatíveis, remove linhas exatamente
    duplicadas e mantém a sequência cronológica de 30 pares por Produto + Ensaio.
11. Controla Bias Atual, Ajuste e Nova Bias por Produto + Ensaio + Verificação.
    O ajuste só ocorre ao completar 30 pares e quando algum critério de controle estiver em alerta.

Dependências:
    pip install streamlit pandas numpy plotly openpyxl

Execução:
    streamlit run carta_controle_curva_global.py
"""

from __future__ import annotations

import io
import json
import re
import unicodedata
from html import escape
from dataclasses import dataclass
from datetime import datetime

import numpy as np
import pandas as pd
import plotly.graph_objects as go
from plotly.offline import get_plotlyjs
import streamlit as st


# -----------------------------------------------------------------------------
# Configurações da planilha
# -----------------------------------------------------------------------------
SHEET_PREFERIDA = "Amostras"
IDX_DATA = 0
IDX_PRODUTO = 1
IDX_TIPO_AMOSTRA = 3
IDX_AMOSTRA = 4
IDX_EQUIPAMENTO = 6
IDX_INICIO_ENSAIOS = 8
PARES_POR_VERIFICACAO = 30

COL_ANALISTA_CANDIDATAS = ["Analista"]
COL_TURNO_CANDIDATAS = ["Turno"]

# Identidade visual baseada no boletim HTML de referência
VERDE = "#00983A"
AZUL = "#004881"
AZUL_ESCURO = "#083050"
NAVY = "#143257"
AMARELO = "#FFCB05"
VERMELHO = "#C0392B"
MUTED = "#6B7785"
FUNDO = "#E9EDF1"
CARD = "#FFFFFF"
BORDA = "#E0E5EA"
GRID = "#EEF1F4"

# -----------------------------------------------------------------------------
# Limites fixos por produto / ensaio — planilha "Limites curva global.xlsx"
# -----------------------------------------------------------------------------
# A diferença utilizada no app é sempre REF - NIR. Os limites abaixo são
# simétricos em torno de zero e são aplicados automaticamente quando a
# combinação Produto + Ensaio selecionada estiver cadastrada nesta tabela.
# Para combinações não cadastradas, o app mantém o cálculo por Carta I.
LIMITES_FIXOS: dict[str, dict[str, float]] = {
    "creme levedura global": {
        "proteina": 23.0,
        "levedo": 5.0,
        "etanol": 0.6,
        "massa seca": 1.6,
    },
    "caldo clarificado pre evaporado global": {
        "brix": 0.3,
        "pol": 0.4,
    },
    "vinho volante global": {
        "brix": 0.3,
        "ph": 0.2,
        "glicerol": 2.6,
        "acidez": 0.4,
        "etanol": 0.3,
        "arrt": 0.02,
    },
    "vinho bruto global": {
        "levedo": 2.0,
        "etanol": 0.3,
    },
    "xarope global": {
        "brix": 0.7,
        "pol": 1.2,
    },
}

# Nomes reais encontrados nas planilhas podem trazer complementos como ano,
# unidade, planta ou descrição. Os aliases abaixo ligam esses nomes à regra
# canônica sem alterar os limites solicitados.
ALIASES_PRODUTOS: dict[str, tuple[str, ...]] = {
    "creme levedura global": ("creme levedura global",),
    "caldo clarificado pre evaporado global": (
        "caldo clarificado pre evaporado global",
    ),
    "vinho volante global": ("vinho volante global",),
    "vinho bruto global": ("vinho bruto global",),
    "xarope global": ("xarope global",),
}

ALIASES_ENSAIOS: dict[str, tuple[str, ...]] = {
    "pol": ("pol",),
    "brix": ("brix",),
    "proteina": ("proteina", "proteinas"),
    "ph": ("ph",),
    "acidez": ("acidez",),
    "levedo": ("levedo", "conc levedo"),
    "etanol": ("etanol",),
    "massa seca": ("massa seca", "materia seca"),
    "glicerol": ("glicerol",),
    # Na planilha de limites consta ARRT; na base Curva Global_02 o par
    # correspondente está identificado como ART HPLC.
    "arrt": ("arrt", "art hplc"),
}



# -----------------------------------------------------------------------------
# Bias atuais informadas em 27/08/2026
# -----------------------------------------------------------------------------
# A data-base ancora os valores fornecidos na última verificação existente até
# este momento. Verificações posteriores recebem automaticamente a Nova Bias da
# verificação anterior, sem deslocar o ponto inicial quando novos dados forem
# acrescentados à planilha.
BIAS_DATA_BASE = pd.Timestamp("2026-08-27 14:30:00")

BIAS_INICIAIS: dict[str, dict[str, float]] = {
    "acucar cor 400 a 1800": {
        "cor": -122.0,
        "pol": 0.01,
        "umidade": -0.01,
        "cinzas": -0.009,
    },
    "bagaco": {
        "pol": -0.23,
        "ar": -0.10,
        "umidade": -0.30,
    },
    "creme de levedura": {
        "proteina": 8.9,
    },
    "levedura seca": {
        "densidade": -0.03,
        "umidade": -1.73,
        "ph": 0.15,
        "proteina": -3.0,
    },
    "meis": {
        "pol": 0.79,
        "brix": 0.28,
    },
    "mosto": {
        "acidez": -0.03,
        "art": 1.00,
        "ph": -0.18,
        "brix": 0.22,
    },
    "torta usm": {
        "pol": -0.14,
        "ar": 0.05,
        "umidade": -0.63,
    },
}

ALIASES_PRODUTOS_BIAS: dict[str, tuple[str, ...]] = {
    "acucar cor 400 a 1800": ("acucar cor 400 a 1800",),
    "bagaco": ("bagaco",),
    "creme de levedura": ("creme de levedura", "creme levedura"),
    "levedura seca": ("levedura seca",),
    "meis": ("meis",),
    "mosto": ("mosto",),
    # Aceita também "Torta" para manter compatibilidade com planilhas que não
    # trazem o sufixo USM no nome do produto.
    "torta usm": ("torta usm", "torta"),
}

@dataclass(frozen=True)
class ParEnsaio:
    nome_exibicao: str
    coluna_nir: str
    coluna_ref: str


def limpar_nome_ensaio(nome: str) -> str:
    """Cria um nome amigável sem alterar o cabeçalho original."""
    texto = str(nome).strip().replace("_", " ")
    texto = re.sub(r"\s+", " ", texto).strip()
    return texto


def normalizar_chave(texto: str) -> str:
    """Normaliza acentos, caixa e separadores para comparação robusta."""
    base = unicodedata.normalize("NFKD", str(texto))
    base = "".join(c for c in base if not unicodedata.combining(c))
    base = base.casefold().replace("_", " ")
    base = re.sub(r"[^a-z0-9]+", " ", base)
    return re.sub(r"\s+", " ", base).strip()


def _alias_presente(valor_norm: str, alias: str) -> bool:
    """Retorna True quando todos os termos do alias aparecem como palavras inteiras."""
    termos_valor = set(valor_norm.split())
    termos_alias = set(normalizar_chave(alias).split())
    return bool(termos_alias) and termos_alias.issubset(termos_valor)


def resolver_produto_limite(produto: str) -> str | None:
    """Resolve o nome real da planilha para uma chave canônica de limite fixo.

    A resolução aceita complementos como ``- 2026`` ou ``USM``, porém utiliza
    aliases explícitos. Assim, ``Açúcar_Cor inf. 400 UI`` não recebe os limites
    de ``Açúcar cor 400 a 1800`` e ``Creme levedura`` não é confundido com
    ``Levedura``.
    """
    produto_norm = normalizar_chave(produto)

    candidatos: list[tuple[int, int, str, str]] = []
    for chave, aliases in ALIASES_PRODUTOS.items():
        for alias in aliases:
            alias_norm = normalizar_chave(alias)
            if _alias_presente(produto_norm, alias_norm):
                candidatos.append(
                    (len(alias_norm.split()), len(alias_norm), chave, alias_norm)
                )

    if not candidatos:
        return None

    # O alias mais específico vence. Ex.: "creme levedura" antes de "levedura".
    candidatos.sort(reverse=True)
    return candidatos[0][2]


def resolver_ensaio_limite(ensaio: str, chaves_permitidas: list[str]) -> str | None:
    """Resolve o ensaio sem confundir nomes próximos, como AR e ART."""
    ensaio_norm = normalizar_chave(ensaio)

    candidatos: list[tuple[int, int, str]] = []
    for chave in chaves_permitidas:
        aliases = ALIASES_ENSAIOS.get(chave, (chave,))
        for alias in aliases:
            alias_norm = normalizar_chave(alias)
            if _alias_presente(ensaio_norm, alias_norm):
                candidatos.append((len(alias_norm.split()), len(alias_norm), chave))

    if not candidatos:
        return None

    candidatos.sort(reverse=True)
    return candidatos[0][2]


def resolver_produto_bias(produto: str) -> str | None:
    """Resolve somente os produtos que possuem Bias inicial cadastrada."""
    produto_norm = normalizar_chave(produto)
    candidatos: list[tuple[int, int, str]] = []
    for chave, aliases in ALIASES_PRODUTOS_BIAS.items():
        for alias in aliases:
            alias_norm = normalizar_chave(alias)
            if _alias_presente(produto_norm, alias_norm):
                candidatos.append((len(alias_norm.split()), len(alias_norm), chave))
    if not candidatos:
        return None
    candidatos.sort(reverse=True)
    return candidatos[0][2]


def obter_bias_inicial(produto: str, ensaio: str) -> float:
    """Retorna a Bias atual informada; usa 0,00 quando ainda não há cadastro."""
    produto_chave = resolver_produto_bias(produto)
    if produto_chave is None:
        return 0.0

    ensaios = BIAS_INICIAIS[produto_chave]
    ensaio_chave = resolver_ensaio_limite(ensaio, list(ensaios))
    if ensaio_chave is None:
        return 0.0
    return float(ensaios[ensaio_chave])


def obter_limites_predefinidos(produto: str, ensaio: str) -> tuple[float, float, str] | None:
    """Retorna LIC, LSC e descrição do limite fixo de Produto + Ensaio.

    Os limites são sempre simétricos em torno de zero porque a variável da carta
    é ``REF - NIR``. Nomes reais da planilha podem conter ano, unidade ou
    descrição adicional; a correspondência é feita por aliases explícitos.
    """
    produto_chave = resolver_produto_limite(produto)
    if produto_chave is None:
        return None

    limites_produto = LIMITES_FIXOS[produto_chave]
    ensaio_chave = resolver_ensaio_limite(ensaio, list(limites_produto))
    if ensaio_chave is None:
        return None

    amplitude = float(limites_produto[ensaio_chave])
    lic, lsc = -amplitude, amplitude
    return lic, lsc, (
        f"Limite fixo · {limpar_nome_ensaio(produto)} / "
        f"{limpar_nome_ensaio(ensaio)} ±{amplitude:g}"
    )


def nome_seguro_arquivo(texto: str) -> str:
    texto = re.sub(r"[^A-Za-z0-9À-ÿ._ -]+", "-", str(texto))
    texto = re.sub(r"\s+", "_", texto).strip("._-")
    return texto or "dados"


def obter_coluna(df: pd.DataFrame, candidatos: list[str]) -> str | None:
    mapa = {str(c).strip().casefold(): c for c in df.columns}
    for nome in candidatos:
        achada = mapa.get(nome.strip().casefold())
        if achada is not None:
            return achada
    return None


def detectar_pares_ensaio(df: pd.DataFrame) -> list[ParEnsaio]:
    """Detecta pares Ensaio / Ref Ensaio somente a partir da coluna I."""
    colunas = list(df.columns)
    candidatos = colunas[IDX_INICIO_ENSAIOS:]

    refs_por_chave: dict[str, str] = {}
    for col in candidatos:
        nome = str(col).strip()
        if nome.casefold().startswith("ref "):
            refs_por_chave[nome[4:].strip().casefold()] = col

    pares: list[ParEnsaio] = []
    nomes_usados: set[str] = set()

    for col in candidatos:
        nome = str(col).strip()
        if nome.casefold().startswith("ref "):
            continue

        ref = refs_por_chave.get(nome.casefold())
        if ref is None:
            continue

        exibicao = limpar_nome_ensaio(nome)
        if exibicao.casefold() in nomes_usados:
            exibicao = f"{exibicao} [{nome}]"
        nomes_usados.add(exibicao.casefold())

        pares.append(ParEnsaio(exibicao, col, ref))

    return pares


def ler_planilha(arquivo) -> tuple[pd.DataFrame, str]:
    """Lê a aba Amostras; se não existir, usa a primeira aba."""
    conteudo = arquivo.getvalue() if hasattr(arquivo, "getvalue") else arquivo.read()
    buffer = io.BytesIO(conteudo)
    xls = pd.ExcelFile(buffer, engine="openpyxl")
    aba = SHEET_PREFERIDA if SHEET_PREFERIDA in xls.sheet_names else xls.sheet_names[0]
    buffer.seek(0)
    df = pd.read_excel(buffer, sheet_name=aba, engine="openpyxl")
    return df, aba


def _nome_arquivo_upload(arquivo, indice: int) -> str:
    nome = getattr(arquivo, "name", None)
    return str(nome) if nome else f"arquivo_{indice}.xlsx"


def consolidar_planilhas(arquivos: list) -> tuple[pd.DataFrame, str, int]:
    """Lê e consolida 1 ou 2 planilhas Excel com a mesma estrutura.

    A primeira planilha define a ordem/cabeçalhos canônicos. A segunda precisa
    possuir as mesmas colunas na mesma ordem após normalização de acentos, caixa
    e espaços. Linhas exatamente duplicadas entre os arquivos são removidas para
    evitar dupla contagem de um histórico repetido.

    Retorna: dataframe consolidado, descrição das abas e nº de duplicatas removidas.
    """
    arquivos = list(arquivos or [])
    if not arquivos:
        raise ValueError("Nenhuma planilha foi selecionada.")
    if len(arquivos) > 2:
        raise ValueError("Selecione no máximo dois arquivos Excel.")

    bases: list[pd.DataFrame] = []
    abas: list[str] = []
    colunas_canonicas: list[str] | None = None
    chaves_canonicas: list[str] | None = None

    for indice, arquivo in enumerate(arquivos, start=1):
        df_atual, aba = ler_planilha(arquivo)
        nome_arquivo = _nome_arquivo_upload(arquivo, indice)

        if df_atual.empty:
            raise ValueError(f"O arquivo '{nome_arquivo}' não possui dados na aba '{aba}'.")

        colunas_atual = [str(c).strip() for c in df_atual.columns]
        chaves_atual = [normalizar_chave(c) for c in colunas_atual]

        if len(set(chaves_atual)) != len(chaves_atual):
            raise ValueError(
                f"O arquivo '{nome_arquivo}' possui cabeçalhos duplicados após normalização. "
                "Revise os nomes das colunas antes de consolidar."
            )

        if colunas_canonicas is None:
            colunas_canonicas = colunas_atual
            chaves_canonicas = chaves_atual
        else:
            if chaves_atual != chaves_canonicas:
                esperado = ", ".join(colunas_canonicas)
                encontrado = ", ".join(colunas_atual)
                raise ValueError(
                    f"A estrutura de '{nome_arquivo}' é diferente da primeira planilha. "
                    "Os dois arquivos precisam ter as mesmas colunas e na mesma ordem. "
                    f"Esperado: [{esperado}]. Encontrado: [{encontrado}]."
                )

        df_atual = df_atual.copy()
        df_atual.columns = colunas_canonicas
        bases.append(df_atual)
        abas.append(f"{nome_arquivo}: {aba}")

    consolidado = pd.concat(bases, ignore_index=True, sort=False)
    total_antes = len(consolidado)
    consolidado = consolidado.drop_duplicates(keep="first").reset_index(drop=True)
    duplicatas_removidas = total_antes - len(consolidado)

    # A ordenação final por data é feita dentro de montar_dados_ensaio para cada
    # Produto + Ensaio, preservando assim a sequência correta das verificações.
    descricao_abas = " · ".join(abas)
    return consolidado, descricao_abas, duplicatas_removidas


def mascara_valida(
    serie_nir: pd.Series,
    serie_ref: pd.Series,
    ignorar_marcador: bool,
    marcador: float,
) -> tuple[pd.Series, pd.Series, pd.Series]:
    nir = pd.to_numeric(serie_nir, errors="coerce")
    ref = pd.to_numeric(serie_ref, errors="coerce")
    mascara = nir.notna() & ref.notna()

    if ignorar_marcador:
        mascara &= ~np.isclose(nir, marcador, equal_nan=False)
        mascara &= ~np.isclose(ref, marcador, equal_nan=False)

    return nir, ref, mascara


def calcular_limites_i_mr(diferencas: pd.Series) -> tuple[float, float, float, float]:
    """Carta de Individuais: sigma = MR médio / 1,128; limites = média ± 3 sigma."""
    valores = pd.to_numeric(diferencas, errors="coerce").dropna()
    if valores.empty:
        return np.nan, np.nan, np.nan, np.nan

    centro = float(valores.mean())
    if len(valores) < 2:
        return centro, np.nan, np.nan, np.nan

    mr_medio = float(valores.diff().abs().dropna().mean())
    sigma = mr_medio / 1.128 if np.isfinite(mr_medio) else np.nan
    lic = centro - 3 * sigma if np.isfinite(sigma) else np.nan
    lsc = centro + 3 * sigma if np.isfinite(sigma) else np.nan
    return centro, lic, lsc, sigma


def calcular_limites_3s(diferencas: pd.Series) -> tuple[float, float, float, float]:
    valores = pd.to_numeric(diferencas, errors="coerce").dropna()
    if valores.empty:
        return np.nan, np.nan, np.nan, np.nan

    centro = float(valores.mean())
    if len(valores) < 2:
        return centro, np.nan, np.nan, np.nan

    sigma = float(valores.std(ddof=1))
    return centro, centro - 3 * sigma, centro + 3 * sigma, sigma


def montar_dados_ensaio(
    df_produto: pd.DataFrame,
    par: ParEnsaio,
    ignorar_marcador: bool,
    marcador: float,
) -> pd.DataFrame:
    """Monta a base de um ensaio com REF - NIR e metadados."""
    nir, ref, mascara = mascara_valida(
        df_produto[par.coluna_nir],
        df_produto[par.coluna_ref],
        ignorar_marcador,
        marcador,
    )

    base = df_produto.loc[mascara].copy()
    if base.empty:
        return base

    col_data = df_produto.columns[IDX_DATA]
    col_produto = df_produto.columns[IDX_PRODUTO]
    col_amostra = df_produto.columns[IDX_AMOSTRA] if len(df_produto.columns) > IDX_AMOSTRA else None
    col_equip = df_produto.columns[IDX_EQUIPAMENTO] if len(df_produto.columns) > IDX_EQUIPAMENTO else None
    col_tipo = df_produto.columns[IDX_TIPO_AMOSTRA] if len(df_produto.columns) > IDX_TIPO_AMOSTRA else None
    col_analista = obter_coluna(df_produto, COL_ANALISTA_CANDIDATAS)
    col_turno = obter_coluna(df_produto, COL_TURNO_CANDIDATAS)

    base["DataHora"] = pd.to_datetime(base[col_data], errors="coerce", dayfirst=True)
    base["Produto"] = base[col_produto].astype(str).str.strip()
    base["Ensaio"] = par.nome_exibicao
    base["NIR"] = nir.loc[mascara].astype(float)
    base["REF"] = ref.loc[mascara].astype(float)
    base["Diferenca"] = base["REF"] - base["NIR"]
    base["Numero_Amostra"] = base[col_amostra].astype(str) if col_amostra else ""
    base["Equipamento"] = base[col_equip].astype(str) if col_equip else ""
    base["Tipo_Amostra"] = base[col_tipo].astype(str) if col_tipo else ""
    base["Analista"] = base[col_analista].astype(str) if col_analista else ""
    base["Turno"] = base[col_turno].astype(str) if col_turno else ""

    base["_ordem_original"] = np.arange(len(base))
    base = base.sort_values(["DataHora", "_ordem_original"], na_position="last", kind="stable").reset_index(drop=True)
    base["Id"] = np.arange(1, len(base) + 1)

    # A sequência de verificação é permanente dentro de cada combinação
    # Produto + Ensaio: pares 1–30 = Verificação 1, 31–60 = Verificação 2,
    # 61–90 = Verificação 3, e assim sucessivamente. Como essa classificação
    # é criada ANTES dos filtros de data, a numeração não é reiniciada quando
    # o usuário recorta o período no painel.
    base["Verificacao"] = ((base["Id"] - 1) // PARES_POR_VERIFICACAO + 1).astype(int)
    base["Posicao_Verificacao"] = ((base["Id"] - 1) % PARES_POR_VERIFICACAO + 1).astype(int)
    base["Verificacao_Rotulo"] = "Verificação " + base["Verificacao"].astype(str)
    base["Pares_na_Verificacao"] = base.groupby("Verificacao")["Id"].transform("count").astype(int)
    base["Verificacao_Completa"] = base["Pares_na_Verificacao"] >= PARES_POR_VERIFICACAO
    base["Status_Verificacao"] = np.where(base["Verificacao_Completa"], "Completa", "Em andamento")
    return base


def aplicar_classificacoes(
    dados: pd.DataFrame,
    centro: float,
    lic: float,
    lsc: float,
) -> pd.DataFrame:
    saida = dados.copy()
    dif = saida["Diferenca"]

    dentro = pd.Series(True, index=saida.index)
    if np.isfinite(lic):
        dentro &= dif >= lic
    else:
        dentro &= False
    if np.isfinite(lsc):
        dentro &= dif <= lsc
    else:
        dentro &= False

    zero = np.isclose(dif.to_numpy(dtype=float), 0.0, atol=1e-12, rtol=0)
    sinal = np.where(zero, "Igual a 0", np.where(dif.to_numpy() < 0, "Menor que 0", "Maior que 0"))

    saida["LC"] = centro
    saida["LIC"] = lic
    saida["LSC"] = lsc
    saida["Dentro_Faixa"] = dentro
    saida["Status_Faixa"] = np.where(dentro, "Dentro da faixa", "Fora da faixa")
    saida["Sinal_Diferenca"] = sinal
    return saida


def indicadores(dados: pd.DataFrame) -> dict[str, float | int]:
    n = len(dados)
    if n == 0:
        return {
            "n": 0,
            "media": np.nan,
            "pct_dentro": np.nan,
            "pct_menor": np.nan,
            "pct_maior": np.nan,
            "pct_zero": np.nan,
        }

    dif = dados["Diferenca"].to_numpy(dtype=float)
    zero = np.isclose(dif, 0.0, atol=1e-12, rtol=0)
    menor = (dif < 0) & ~zero
    maior = (dif > 0) & ~zero

    return {
        "n": n,
        "media": float(np.mean(dif)),
        "pct_dentro": float(dados["Dentro_Faixa"].mean() * 100),
        "pct_menor": float(menor.mean() * 100),
        "pct_maior": float(maior.mean() * 100),
        "pct_zero": float(zero.mean() * 100),
    }


def filtrar_verificacao(dados: pd.DataFrame, verificacao: str | int) -> pd.DataFrame:
    """Filtra o bloco sequencial de 30 pares sem renumerar os dados."""
    if verificacao == "Todos":
        return dados.copy()
    try:
        numero = int(verificacao)
    except (TypeError, ValueError):
        return dados.copy()
    return dados[dados["Verificacao"] == numero].copy()


def criar_grafico_powerbi(
    dados: pd.DataFrame,
    ensaio: str,
    unidade: str,
    lic: float,
    lsc: float,
    centro: float,
    mostrar_lc: bool,
    destacar_desvios: bool = True,
) -> go.Figure:
    """Carta de controle com a identidade visual do boletim HTML de referência."""
    fig = go.Figure()

    titulo_unidade = f" {unidade}" if unidade.strip() else ""
    status_fora = ~dados["Dentro_Faixa"].to_numpy(dtype=bool)
    marker_colors = [VERMELHO if (destacar_desvios and fora) else AZUL for fora in status_fora]
    marker_sizes = [9 if (destacar_desvios and fora) else 7 for fora in status_fora]

    # O eixo X exibe o "Tempo de análise" completo (data + horário).
    # Para evitar sobreposição visual quando várias análises acontecem no mesmo
    # dia ou em horários muito próximos, cada observação recebe uma posição
    # sequencial exclusiva no gráfico. Os rótulos do eixo e o hover continuam
    # mostrando exclusivamente a data/hora real da coluna Tempo de análise.
    x_tempo = pd.to_datetime(dados["DataHora"], errors="coerce")
    x_pos = np.arange(len(dados), dtype=int)

    # Mantém o eixo legível mesmo com muitos pontos: exibe até 12 referências
    # de data/hora, distribuídas ao longo de toda a série.
    if len(dados) <= 12:
        tick_idx = x_pos
    else:
        tick_idx = np.unique(np.linspace(0, len(dados) - 1, 12, dtype=int))
    tick_text = [
        x_tempo.iloc[i].strftime("%d/%m/%Y<br>%H:%M:%S") if pd.notna(x_tempo.iloc[i]) else "Sem data/hora"
        for i in tick_idx
    ]

    fig.add_trace(
        go.Scatter(
            x=x_pos,
            y=dados["Diferenca"],
            mode="lines+markers",
            name="Diferença Ref - NIR",
            line={"color": AZUL, "width": 2.6},
            marker={
                "color": marker_colors,
                "size": marker_sizes,
                "line": {"color": "#FFFFFF", "width": 1.1},
            },
            customdata=np.column_stack(
                [
                    dados["Id"],
                    dados["Numero_Amostra"].astype(str),
                    dados["DataHora"].dt.strftime("%d/%m/%Y %H:%M:%S").fillna("-"),
                    dados["NIR"],
                    dados["REF"],
                    dados["Status_Faixa"],
                ]
            ),
            hovertemplate=(
                "<b>Tempo de análise: %{customdata[2]}</b><br>"
                "Id: %{customdata[0]}<br>"
                "Amostra: %{customdata[1]}<br>"
                "NIR: %{customdata[3]:.4f}<br>"
                "Ref: %{customdata[4]:.4f}<br>"
                "<b>Ref - NIR: %{y:.4f}</b><br>"
                "%{customdata[5]}<extra></extra>"
            ),
        )
    )

    xmin = int(x_pos.min()) if len(x_pos) else 0
    xmax = int(x_pos.max()) if len(x_pos) else 1

    if np.isfinite(lsc):
        fig.add_trace(
            go.Scatter(
                x=[xmin, xmax], y=[lsc, lsc], mode="lines", name="LSC",
                line={"color": VERMELHO, "width": 1.8, "dash": "dot"}, hoverinfo="skip"
            )
        )
        fig.add_annotation(
            x=xmax, y=lsc, text=f"LSC {lsc:.4f}", showarrow=False, xanchor="right", yshift=10,
            font={"family": "JetBrains Mono", "size": 10, "color": VERMELHO},
            bgcolor="rgba(255,255,255,0.86)", borderpad=2,
        )

    if np.isfinite(lic):
        fig.add_trace(
            go.Scatter(
                x=[xmin, xmax], y=[lic, lic], mode="lines", name="LIC",
                line={"color": VERMELHO, "width": 1.8, "dash": "dot"}, hoverinfo="skip"
            )
        )
        fig.add_annotation(
            x=xmax, y=lic, text=f"LIC {lic:.4f}", showarrow=False, xanchor="right", yshift=-10,
            font={"family": "JetBrains Mono", "size": 10, "color": VERMELHO},
            bgcolor="rgba(255,255,255,0.86)", borderpad=2,
        )

    if mostrar_lc and np.isfinite(centro):
        fig.add_trace(
            go.Scatter(
                x=[xmin, xmax], y=[centro, centro], mode="lines", name="LC",
                line={"color": VERDE, "width": 1.7, "dash": "dash"}, hoverinfo="skip"
            )
        )

    fig.add_hline(y=0, line_width=1, line_dash="dot", line_color="#AEB8C2")

    fig.update_layout(
        title={
            "text": f"{ensaio} · Diferença Ref − NIR{titulo_unidade}",
            "x": 0.0, "xanchor": "left",
            "font": {"family": "Montserrat", "size": 17, "color": NAVY},
        },
        xaxis_title="Tempo de análise",
        yaxis_title=f"Diferença{titulo_unidade}",
        height=500,
        margin={"l": 55, "r": 30, "t": 62, "b": 95},
        hovermode="closest",
        plot_bgcolor="#FFFFFF",
        paper_bgcolor="#FFFFFF",
        showlegend=False,
        font={"family": "Montserrat", "color": "#2B3542"},
        hoverlabel={"bgcolor": "#FFFFFF", "bordercolor": BORDA, "font": {"family": "JetBrains Mono"}},
    )
    fig.update_xaxes(
        showgrid=True, gridcolor=GRID, griddash="dot", zeroline=False,
        title_font={"family": "JetBrains Mono", "size": 11, "color": MUTED},
        tickfont={"family": "JetBrains Mono", "size": 10, "color": MUTED},
        linecolor=BORDA,
        tickmode="array",
        tickvals=tick_idx.tolist(),
        ticktext=tick_text,
        tickangle=-35,
        automargin=True,
    )
    fig.update_yaxes(
        showgrid=True, gridcolor=GRID, griddash="dot", zeroline=False,
        title_font={"family": "JetBrains Mono", "size": 11, "color": MUTED},
        tickfont={"family": "JetBrains Mono", "size": 10, "color": MUTED},
        linecolor=BORDA,
    )
    return fig




def fmt_num(valor: float, casas: int = 2) -> str:
    if not np.isfinite(valor):
        return "—"
    return f"{valor:.{casas}f}".replace(".", ",")


def fmt_pct(valor: float) -> str:
    if not np.isfinite(valor):
        return "—"
    return f"{valor:.1f}%".replace(".", ",")


def status_kpi_intervalo(valor: float, lic: float, lsc: float) -> str:
    """Classifica o Limite de Controle 1 pela faixa LIC–LSC ativa do ensaio."""
    if not (np.isfinite(valor) and np.isfinite(lic) and np.isfinite(lsc)):
        return "neutro"
    return "alerta" if valor < lic or valor > lsc else "ok"


def status_kpi_minimo(valor: float, minimo: float = 70.0) -> str:
    """Classifica o Limite de Controle 2: abaixo de 70% é alerta."""
    if not np.isfinite(valor):
        return "neutro"
    return "alerta" if valor < minimo else "ok"


def status_kpi_maximo(valor: float, maximo: float = 70.0) -> str:
    """Classifica o Limite de Controle 3: acima de 70% é alerta."""
    if not np.isfinite(valor):
        return "neutro"
    return "alerta" if valor > maximo else "ok"


def criterio_exige_ajuste(kpis: dict[str, float | int], lic: float, lsc: float) -> bool:
    """True quando qualquer um dos critérios existentes do painel está em alerta."""
    media_dif = float(kpis["media"])
    pct_dentro = float(kpis["pct_dentro"])
    pct_menor = float(kpis["pct_menor"])
    pct_maior = float(kpis["pct_maior"])
    pct_zero = float(kpis["pct_zero"])
    return any(
        status == "alerta"
        for status in (
            status_kpi_intervalo(media_dif, lic, lsc),
            status_kpi_minimo(pct_dentro, 70.0),
            status_kpi_maximo(pct_menor, 70.0),
            status_kpi_maximo(pct_maior, 70.0),
            status_kpi_maximo(pct_zero, 70.0),
        )
    )


def obter_verificacao_base_bias(dados: pd.DataFrame) -> int:
    """Última verificação existente na data em que as Bias iniciais foram informadas."""
    if dados.empty:
        return 1

    datas = pd.to_datetime(dados["DataHora"], errors="coerce")
    anteriores = dados.loc[datas.notna() & (datas <= BIAS_DATA_BASE), "Verificacao"]
    if not anteriores.empty:
        return int(pd.to_numeric(anteriores, errors="coerce").dropna().max())

    # Fallback para planilhas sem data/hora válida.
    return int(pd.to_numeric(dados["Verificacao"], errors="coerce").dropna().max())


def calcular_historico_bias(
    dados: pd.DataFrame,
    produto: str,
    ensaio: str,
) -> pd.DataFrame:
    """Calcula Bias Atual -> Ajuste -> Nova Bias para cada verificação.

    A Bias informada entra na última verificação existente até BIAS_DATA_BASE.
    Somente verificações completas (30 pares) podem gerar ajuste. Se qualquer
    critério do painel estiver em alerta, o ajuste é a média de REF - NIR.
    A Nova Bias passa a ser a Bias Atual da verificação seguinte.
    """
    colunas = [
        "Verificacao", "Bias_Atual", "Ajuste_Bias", "Nova_Bias",
        "Bias_Verificacao_Base", "Bias_Verificacao_Completa",
        "Bias_Ajuste_Aplicado",
    ]
    if dados.empty:
        return pd.DataFrame(columns=colunas)

    verificacoes = sorted(
        pd.to_numeric(dados["Verificacao"], errors="coerce").dropna().astype(int).unique()
    )
    if not verificacoes:
        return pd.DataFrame(columns=colunas)

    verificacao_base = obter_verificacao_base_bias(dados)
    bias_inicial = obter_bias_inicial(produto, ensaio)
    limite_fixo = obter_limites_predefinidos(produto, ensaio)

    registros: list[dict[str, object]] = []
    bias_corrente = float(bias_inicial)
    iniciou = False

    for numero in verificacoes:
        bloco = dados[dados["Verificacao"] == numero].copy()
        completa = len(bloco) >= PARES_POR_VERIFICACAO

        if numero < verificacao_base:
            registros.append({
                "Verificacao": numero,
                "Bias_Atual": np.nan,
                "Ajuste_Bias": np.nan,
                "Nova_Bias": np.nan,
                "Bias_Verificacao_Base": verificacao_base,
                "Bias_Verificacao_Completa": completa,
                "Bias_Ajuste_Aplicado": False,
            })
            continue

        if not iniciou:
            bias_corrente = float(bias_inicial)
            iniciou = True

        if limite_fixo is not None:
            lic_bloco, lsc_bloco, _ = limite_fixo
            centro_bloco = 0.0
        else:
            centro_bloco, lic_bloco, lsc_bloco, _ = calcular_limites_i_mr(bloco["Diferenca"])

        ajuste = 0.0
        ajuste_aplicado = False
        if completa:
            bloco_classificado = aplicar_classificacoes(bloco, centro_bloco, lic_bloco, lsc_bloco)
            kpis_bloco = indicadores(bloco_classificado)
            ajuste_aplicado = criterio_exige_ajuste(kpis_bloco, lic_bloco, lsc_bloco)
            if ajuste_aplicado:
                ajuste = float(kpis_bloco["media"])

        nova_bias = float(bias_corrente + ajuste)
        registros.append({
            "Verificacao": numero,
            "Bias_Atual": float(bias_corrente),
            "Ajuste_Bias": float(ajuste),
            "Nova_Bias": nova_bias,
            "Bias_Verificacao_Base": verificacao_base,
            "Bias_Verificacao_Completa": completa,
            "Bias_Ajuste_Aplicado": ajuste_aplicado,
        })
        bias_corrente = nova_bias

    return pd.DataFrame(registros, columns=colunas)


def obter_estado_bias(
    historico: pd.DataFrame,
    verificacao: str | int,
) -> dict[str, float | int | bool]:
    """Retorna o estado dos cartões; 'Todos' usa a última verificação."""
    if historico.empty:
        return {
            "bias_atual": 0.0, "ajuste": 0.0, "nova_bias": 0.0,
            "verificacao": 0, "completa": False, "ajuste_aplicado": False,
        }

    if verificacao == "Todos":
        linha = historico.sort_values("Verificacao").iloc[-1]
    else:
        alvo = historico[historico["Verificacao"] == int(verificacao)]
        linha = alvo.iloc[0] if not alvo.empty else historico.sort_values("Verificacao").iloc[-1]

    def _num(valor):
        return float(valor) if pd.notna(valor) and np.isfinite(float(valor)) else np.nan

    return {
        "bias_atual": _num(linha["Bias_Atual"]),
        "ajuste": _num(linha["Ajuste_Bias"]),
        "nova_bias": _num(linha["Nova_Bias"]),
        "verificacao": int(linha["Verificacao"]),
        "completa": bool(linha["Bias_Verificacao_Completa"]),
        "ajuste_aplicado": bool(linha["Bias_Ajuste_Aplicado"]),
    }


def fmt_bias(valor: float) -> str:
    """Exibe Bias com 2 casas e preserva a terceira quando necessária."""
    if not np.isfinite(valor):
        return "—"
    casas = 3 if not np.isclose(valor, round(valor, 2), atol=1e-12, rtol=0) else 2
    return f"{valor:.{casas}f}".replace(".", ",")


def html_kpi(rotulo: str, valor: str, detalhe: str = "", status: str = "ok") -> str:
    status_seguro = status if status in {"ok", "alerta", "neutro"} else "neutro"
    return (
        f'<div class="qmp-kpi qmp-kpi-{status_seguro}">'
        f'<div class="qmp-kpi-rotulo">{escape(rotulo)}</div>'
        f'<div class="qmp-kpi-valor">{escape(valor)}</div>'
        f'<div class="qmp-kpi-detalhe">{escape(detalhe)}</div>'
        '</div>'
    )


def html_cartao_titulo(rotulo: str, valor: str = "0,00") -> str:
    """Cartão visual para Bias Atual, Ajuste ou Nova Bias."""
    return (
        '<div class="qmp-kpi qmp-kpi-titulo">'
        f'<div class="qmp-kpi-rotulo">{escape(rotulo)}</div>'
        f'<div class="qmp-kpi-valor">{escape(valor)}</div>'
        '</div>'
    )


def html_header(produto: str, ensaio: str, unidade: str, periodo_txt: str, n: int, aba: str) -> str:
    # Cabeçalho simplificado conforme solicitado: exibe somente o título.
    return """
    <div class="qmp-header">
      <div class="qmp-header-texto">
        <div class="qmp-titulo">Carta controle NIR DS-3 — Curvas Globais</div>
      </div>
    </div>
    """



def montar_base_html_interativa(
    df: pd.DataFrame,
    pares: list[ParEnsaio],
    ignorar_marcador: bool,
    marcador: float,
) -> pd.DataFrame:
    """Monta a base completa embutida no HTML para os filtros interativos."""
    col_produto = df.columns[IDX_PRODUTO]
    produtos = sorted(df[col_produto].dropna().astype(str).str.strip().unique().tolist())
    partes: list[pd.DataFrame] = []

    for produto in produtos:
        df_produto = df[df[col_produto].astype(str).str.strip() == produto].copy()
        for par in pares:
            dados = montar_dados_ensaio(df_produto, par, ignorar_marcador, marcador)
            if dados.empty:
                continue

            limite_fixo = obter_limites_predefinidos(produto, par.nome_exibicao)
            if limite_fixo is not None:
                lic_fixo, lsc_fixo, metodo = limite_fixo
            else:
                lic_fixo, lsc_fixo, metodo = np.nan, np.nan, "Carta I (amplitude móvel)"

            historico_bias = calcular_historico_bias(dados, produto, par.nome_exibicao)

            bloco = dados[
                ["Id", "Verificacao", "DataHora", "Produto", "Ensaio", "NIR", "REF", "Diferenca"]
            ].copy()
            bloco["LIC_Fixo"] = lic_fixo
            bloco["LSC_Fixo"] = lsc_fixo
            bloco["Metodo_Limite"] = metodo
            bloco = bloco.merge(
                historico_bias[
                    [
                        "Verificacao", "Bias_Atual", "Ajuste_Bias", "Nova_Bias",
                        "Bias_Verificacao_Base", "Bias_Verificacao_Completa",
                        "Bias_Ajuste_Aplicado",
                    ]
                ],
                on="Verificacao",
                how="left",
            )
            partes.append(bloco)

    if not partes:
        return pd.DataFrame(
            columns=[
                "Id", "Verificacao", "DataHora", "Produto", "Ensaio", "NIR", "REF",
                "Diferenca", "LIC_Fixo", "LSC_Fixo", "Metodo_Limite",
                "Bias_Atual", "Ajuste_Bias", "Nova_Bias", "Bias_Verificacao_Base",
                "Bias_Verificacao_Completa", "Bias_Ajuste_Aplicado",
            ]
        )

    base = pd.concat(partes, ignore_index=True)
    base["DataHora"] = base["DataHora"].apply(
        lambda v: v.isoformat() if pd.notna(v) and hasattr(v, "isoformat") else None
    )
    return base


def montar_html_relatorio(
    base_interativa: pd.DataFrame,
    produto_inicial: str,
    ensaio_inicial: str,
    verificacao_inicial: str | int,
    atualizado_em: datetime,
) -> bytes:
    """Gera HTML autônomo com filtros de Produto, Ensaio e Verificação.

    Todo o conjunto de pares válidos é incorporado ao arquivo. Os filtros, KPIs,
    limites, gráfico e boletim são recalculados localmente no navegador, sem
    depender do Streamlit depois do download.
    """
    registros: list[dict[str, object]] = []
    if not base_interativa.empty:
        for row in base_interativa.to_dict(orient="records"):
            registro: dict[str, object] = {}
            for chave, valor in row.items():
                if isinstance(valor, (np.integer,)):
                    registro[chave] = int(valor)
                elif isinstance(valor, (np.floating, float)):
                    registro[chave] = float(valor) if np.isfinite(valor) else None
                elif pd.isna(valor):
                    registro[chave] = None
                else:
                    registro[chave] = valor
            registros.append(registro)

    dados_json = json.dumps(registros, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    inicial_json = json.dumps(
        {
            "produto": produto_inicial,
            "ensaio": ensaio_inicial,
            "verificacao": str(verificacao_inicial),
        },
        ensure_ascii=False,
    ).replace("</", "<\\/")
    atualizado_txt = atualizado_em.strftime("%d/%m/%Y %H:%M:%S")
    plotly_js = get_plotlyjs()

    html_template = r"""<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Carta controle NIR DS-3 — Curvas Globais</title>
<style>
:root { --verde:__VERDE__; --azul:__AZUL__; --navy:__NAVY__; --vermelho:__VERMELHO__; --muted:__MUTED__; --fundo:__FUNDO__; --borda:__BORDA__; }
* { box-sizing:border-box; }
body { margin:0; padding:24px; background:var(--fundo); color:#2B3542; font-family:Arial,Helvetica,sans-serif; }
.wrap { max-width:1500px; margin:0 auto; }
.header { background:linear-gradient(118deg,var(--verde) 0%,var(--azul) 64%,#083050 100%); color:#fff; border-radius:12px; padding:24px 30px; }
.header h1 { margin:0 0 8px; font-size:30px; }
.meta { font-size:12px; opacity:.92; line-height:1.6; }
.card { background:#fff; border:1px solid var(--borda); border-radius:10px; box-shadow:0 1px 2px rgba(12,34,64,.06); padding:16px; margin-top:16px; overflow:auto; }
h2 { margin:0 0 4px; color:var(--navy); font-size:18px; }
.note { color:var(--muted); font-size:12px; margin-bottom:12px; }
.filtros-grid { display:grid; grid-template-columns:repeat(3,minmax(0,1fr)); gap:14px; margin-top:12px; }
.filtro label { display:block; color:#4B5868; font-size:12px; font-weight:700; margin-bottom:6px; }
.filtro select { width:100%; height:42px; border:1px solid var(--borda); border-radius:7px; background:#fff; color:#2B3542; padding:0 10px; font-size:13px; }
.kpis { display:grid; grid-template-columns:repeat(5,minmax(0,1fr)); gap:10px; margin:16px 0 10px; }
.kpis-secundarios { width:60%; margin:0 auto 16px; display:grid; grid-template-columns:repeat(3,minmax(0,1fr)); gap:10px; }
.kpi { background:#fff; border:1px solid var(--borda); border-radius:10px; box-shadow:0 1px 2px rgba(12,34,64,.06); padding:14px; text-align:center; }
.kpi.titulo { min-height:96px; display:flex; flex-direction:column; align-items:center; justify-content:center; gap:8px; }
.kpi .label { color:var(--muted); font-size:12px; min-height:30px; }
.kpi .value { color:var(--navy); font-size:21px; font-weight:700; margin-top:6px; }
.kpi.titulo .label { min-height:0; color:#3E3E3E; font-size:13px; font-weight:600; }
.kpi.titulo .value { color:var(--navy); margin-top:0; }
.kpi.ok .value { color:#009A44; }
.kpi.alerta .value { color:var(--vermelho); }
.kpi.neutro .value { color:var(--muted); }
#grafico { width:100%; min-height:455px; }
table { width:100%; border-collapse:collapse; font-size:12px; table-layout:fixed; }
th { background:var(--navy); color:#fff; text-align:center; padding:9px 8px; position:sticky; top:0; z-index:2; }
td { border-bottom:1px solid #E8ECF0; padding:8px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; text-align:center; }
td.status { font-weight:700; }
td.dentro { color:#00702C; background:#E9F5EE; }
td.fora { color:var(--vermelho); background:#FDECEA; }
.table-wrap { max-height:520px; overflow:auto; border:1px solid var(--borda); border-radius:8px; }
@media (max-width:900px) {
  body { padding:10px; }
  .filtros-grid { grid-template-columns:1fr; }
  .kpis { grid-template-columns:1fr 1fr; }
  .kpis-secundarios { width:100%; grid-template-columns:1fr; }
}
</style>
<script>__PLOTLY_JS__</script>
</head>
<body>
<div class="wrap">
  <div class="header">
    <h1>Carta controle NIR DS-3 — Curvas Globais</h1>
    <div class="meta" id="metaRelatorio">Atualizado em __ATUALIZADO__</div>
  </div>

  <div class="card">
    <h2>Filtros das análises</h2>
    <div class="note">Os filtros abaixo atualizam os cartões, a Carta Controle e o Boletim analítico.</div>
    <div class="filtros-grid">
      <div class="filtro"><label for="filtroProduto">Produto</label><select id="filtroProduto"></select></div>
      <div class="filtro"><label for="filtroEnsaio">Ensaio</label><select id="filtroEnsaio"></select></div>
      <div class="filtro"><label for="filtroVerificacao">Verificação</label><select id="filtroVerificacao"></select></div>
    </div>
  </div>

  <div class="kpis">
    <div class="kpi neutro" id="kpiMedia"><div class="label">Média de Dif</div><div class="value">—</div></div>
    <div class="kpi neutro" id="kpiDentro"><div class="label">% Dentro da Faixa</div><div class="value">—</div></div>
    <div class="kpi neutro" id="kpiMenor"><div class="label">Diferença &lt; 0 (%)</div><div class="value">—</div></div>
    <div class="kpi neutro" id="kpiMaior"><div class="label">Diferença &gt; 0 (%)</div><div class="value">—</div></div>
    <div class="kpi neutro" id="kpiZero"><div class="label">Diferença = 0 (%)</div><div class="value">—</div></div>
  </div>

  <div class="kpis-secundarios">
    <div class="kpi titulo" id="kpiBiasAtual"><div class="label">Bias Atuais</div><div class="value">0,00</div></div>
    <div class="kpi titulo" id="kpiAjuste"><div class="label">Ajustes</div><div class="value">0,00</div></div>
    <div class="kpi titulo" id="kpiNovaBias"><div class="label">Novas Bias</div><div class="value">0,00</div></div>
  </div>

  <div class="card">
    <h2>Carta Controle</h2>
    <div class="note">Cada ponto representa Ref − NIR. Pontos fora dos limites são destacados em vermelho.</div>
    <div id="grafico"></div>
  </div>

  <div class="card">
    <h2>Boletim analítico</h2>
    <div class="note" id="notaBoletim">0 observações · diferença calculada como Ref − NIR</div>
    <div class="table-wrap">
      <table>
        <thead><tr><th>Id</th><th>Data/hora</th><th>Ref</th><th>NIR</th><th>Ref - NIR</th><th>LIC</th><th>LSC</th><th>Status</th></tr></thead>
        <tbody id="corpoTabela"></tbody>
      </table>
    </div>
  </div>
</div>

<script>
const BASE = __DADOS_JSON__;
const INICIAL = __INICIAL_JSON__;
const CORES = {azul:"__AZUL__", vermelho:"__VERMELHO__", verde:"__VERDE__", navy:"__NAVY__", muted:"__MUTED__", borda:"__BORDA__"};
const PARES_POR_VERIFICACAO = 30;

const selProduto = document.getElementById('filtroProduto');
const selEnsaio = document.getElementById('filtroEnsaio');
const selVerificacao = document.getElementById('filtroVerificacao');

function unicos(arr) { return [...new Set(arr)]; }
function ordenarTexto(arr) { return [...arr].sort((a,b) => String(a).localeCompare(String(b), 'pt-BR', {numeric:true})); }
function preencherSelect(select, valores, preferido) {
  select.innerHTML = '';
  valores.forEach(v => {
    const opt = document.createElement('option');
    opt.value = String(v.value ?? v);
    opt.textContent = String(v.label ?? v);
    select.appendChild(opt);
  });
  const desejado = String(preferido ?? '');
  if ([...select.options].some(o => o.value === desejado)) select.value = desejado;
}

function fmtNum(v, casas=2) {
  return Number.isFinite(v) ? v.toLocaleString('pt-BR', {minimumFractionDigits:casas, maximumFractionDigits:casas}) : '—';
}
function fmtPct(v) { return Number.isFinite(v) ? fmtNum(v, 1) + '%' : '—'; }
function fmtTabela(v) { return Number.isFinite(v) ? fmtNum(v, 4) : '—'; }
function fmtBias(v) {
  if (!Number.isFinite(v)) return '—';
  const arred2 = Math.round(v * 100) / 100;
  const casas = Math.abs(v - arred2) > 1e-12 ? 3 : 2;
  return v.toLocaleString('pt-BR', {minimumFractionDigits:casas, maximumFractionDigits:casas});
}
function fmtData(iso) {
  if (!iso) return '—';
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return String(iso);
  const p = n => String(n).padStart(2,'0');
  return `${p(d.getDate())}/${p(d.getMonth()+1)}/${d.getFullYear()} ${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}`;
}
function media(valores) { return valores.length ? valores.reduce((a,b)=>a+b,0) / valores.length : NaN; }

function calcularLimites(linhas) {
  if (!linhas.length) return {centro:NaN, lic:NaN, lsc:NaN, metodo:'—'};
  const licFixo = linhas[0].LIC_Fixo;
  const lscFixo = linhas[0].LSC_Fixo;
  if (Number.isFinite(licFixo) && Number.isFinite(lscFixo)) {
    return {centro:0, lic:licFixo, lsc:lscFixo, metodo:linhas[0].Metodo_Limite || 'Limite fixo'};
  }
  const dif = linhas.map(r => r.Diferenca).filter(Number.isFinite);
  const centro = media(dif);
  if (dif.length < 2) return {centro, lic:NaN, lsc:NaN, metodo:'Carta I (amplitude móvel)'};
  const mr = [];
  for (let i=1; i<dif.length; i++) mr.push(Math.abs(dif[i] - dif[i-1]));
  const mrMedio = media(mr);
  const sigma = Number.isFinite(mrMedio) ? mrMedio / 1.128 : NaN;
  return {
    centro,
    lic:Number.isFinite(sigma) ? centro - 3*sigma : NaN,
    lsc:Number.isFinite(sigma) ? centro + 3*sigma : NaN,
    metodo:'Carta I (amplitude móvel)'
  };
}

function comStatus(linhas, limites) {
  return linhas.map(r => {
    const dentro = Number.isFinite(limites.lic) && Number.isFinite(limites.lsc) && r.Diferenca >= limites.lic && r.Diferenca <= limites.lsc;
    return {...r, Dentro_Faixa:dentro, Status_Faixa:dentro ? 'Dentro da faixa' : 'Fora da faixa'};
  });
}

function statusIntervalo(v, lic, lsc) {
  if (![v,lic,lsc].every(Number.isFinite)) return 'neutro';
  return (v < lic || v > lsc) ? 'alerta' : 'ok';
}
function statusMinimo(v, minimo=70) { return Number.isFinite(v) ? (v < minimo ? 'alerta' : 'ok') : 'neutro'; }
function statusMaximo(v, maximo=70) { return Number.isFinite(v) ? (v > maximo ? 'alerta' : 'ok') : 'neutro'; }
function setKpi(id, valor, status) {
  const el = document.getElementById(id);
  el.className = `kpi ${status}`;
  el.querySelector('.value').textContent = valor;
}

function atualizarOpcoesEnsaios(preferido=null) {
  const produto = selProduto.value;
  const ensaios = ordenarTexto(unicos(BASE.filter(r => r.Produto === produto).map(r => r.Ensaio)));
  preencherSelect(selEnsaio, ensaios, preferido);
}

function atualizarOpcoesVerificacoes(preferido='Todos') {
  const produto = selProduto.value;
  const ensaio = selEnsaio.value;
  const numeros = [...new Set(BASE.filter(r => r.Produto === produto && r.Ensaio === ensaio).map(r => Number(r.Verificacao)))].sort((a,b)=>a-b);
  const opcoes = [{value:'Todos', label:`Todas (${numeros.length} verificações)`}];
  numeros.forEach(numero => {
    const bloco = BASE.filter(r => r.Produto === produto && r.Ensaio === ensaio && Number(r.Verificacao) === numero);
    const inicio = (numero - 1) * PARES_POR_VERIFICACAO + 1;
    const fim = bloco.length ? Math.max(...bloco.map(r => Number(r.Id))) : numero * PARES_POR_VERIFICACAO;
    const status = bloco.length >= PARES_POR_VERIFICACAO ? 'completa' : 'em andamento';
    opcoes.push({value:String(numero), label:`Verificação ${numero} · pares ${inicio}–${fim} · ${bloco.length}/${PARES_POR_VERIFICACAO} · ${status}`});
  });
  preencherSelect(selVerificacao, opcoes, preferido);
}

function linhasSelecionadas() {
  const produto = selProduto.value;
  const ensaio = selEnsaio.value;
  const verificacao = selVerificacao.value;
  return BASE.filter(r => r.Produto === produto && r.Ensaio === ensaio && (verificacao === 'Todos' || String(r.Verificacao) === verificacao));
}

function atualizarMeta(linhas, limites) {
  const datas = linhas.map(r => r.DataHora).filter(Boolean).map(v => new Date(v)).filter(d => !Number.isNaN(d.getTime())).sort((a,b)=>a-b);
  const periodo = datas.length ? `${fmtData(datas[0].toISOString()).slice(0,10)} — ${fmtData(datas[datas.length-1].toISOString()).slice(0,10)}` : 'sem data válida';
  const verificacao = selVerificacao.value === 'Todos' ? 'Todas as verificações' : `Verificação ${selVerificacao.value}`;
  document.getElementById('metaRelatorio').innerHTML = `Produto: ${escapeHtml(selProduto.value)} &nbsp;·&nbsp; Ensaio: ${escapeHtml(selEnsaio.value)} &nbsp;·&nbsp; ${escapeHtml(verificacao)}<br>Período ${periodo} &nbsp;·&nbsp; Método de limites: ${escapeHtml(limites.metodo)} &nbsp;·&nbsp; Atualizado em __ATUALIZADO__`;
}

function escapeHtml(texto) {
  return String(texto ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[c]));
}

function atualizarBias(linhas) {
  const elAtual = document.getElementById('kpiBiasAtual').querySelector('.value');
  const elAjuste = document.getElementById('kpiAjuste').querySelector('.value');
  const elNova = document.getElementById('kpiNovaBias').querySelector('.value');

  if (!linhas.length) {
    elAtual.textContent = '—';
    elAjuste.textContent = '—';
    elNova.textContent = '—';
    return;
  }

  let alvo;
  if (selVerificacao.value === 'Todos') {
    const ultima = Math.max(...linhas.map(r => Number(r.Verificacao)));
    alvo = linhas.find(r => Number(r.Verificacao) === ultima);
  } else {
    alvo = linhas[0];
  }

  elAtual.textContent = fmtBias(alvo?.Bias_Atual);
  elAjuste.textContent = fmtBias(alvo?.Ajuste_Bias);
  elNova.textContent = fmtBias(alvo?.Nova_Bias);
}

function atualizarKpis(linhas, limites) {
  const n = linhas.length;
  const dif = linhas.map(r => r.Diferenca).filter(Number.isFinite);
  const mediaDif = media(dif);
  const pctDentro = n ? linhas.filter(r => r.Dentro_Faixa).length / n * 100 : NaN;
  const pctMenor = n ? linhas.filter(r => r.Diferenca < 0 && Math.abs(r.Diferenca) > 1e-12).length / n * 100 : NaN;
  const pctMaior = n ? linhas.filter(r => r.Diferenca > 0 && Math.abs(r.Diferenca) > 1e-12).length / n * 100 : NaN;
  const pctZero = n ? linhas.filter(r => Math.abs(r.Diferenca) <= 1e-12).length / n * 100 : NaN;

  setKpi('kpiMedia', fmtNum(mediaDif,2), statusIntervalo(mediaDif, limites.lic, limites.lsc));
  setKpi('kpiDentro', fmtPct(pctDentro), statusMinimo(pctDentro));
  setKpi('kpiMenor', fmtPct(pctMenor), statusMaximo(pctMenor));
  setKpi('kpiMaior', fmtPct(pctMaior), statusMaximo(pctMaior));
  setKpi('kpiZero', fmtPct(pctZero), statusMaximo(pctZero));
}

function atualizarGrafico(linhas, limites) {
  // O eixo mostra Tempo de análise (data + horário), mas usa uma posição
  // sequencial exclusiva para cada observação. Isso impede que pontos do mesmo
  // dia ou de horários muito próximos fiquem desenhados praticamente no mesmo X.
  const x = linhas.map((_, i) => i);
  const y = linhas.map(r => r.Diferenca);
  const cores = linhas.map(r => r.Dentro_Faixa ? CORES.azul : CORES.vermelho);
  const tamanhos = linhas.map(r => r.Dentro_Faixa ? 7 : 9);
  const custom = linhas.map(r => [fmtData(r.DataHora), r.Id, r.NIR, r.REF, r.Status_Faixa]);
  const traces = [{
    x, y, mode:'lines+markers', name:'Diferença Ref - NIR',
    line:{color:CORES.azul, width:2.6},
    marker:{color:cores, size:tamanhos, line:{color:'#FFFFFF', width:1.1}},
    customdata:custom,
    hovertemplate:'<b>Tempo de análise: %{customdata[0]}</b><br>Id: %{customdata[1]}<br>NIR: %{customdata[2]:.4f}<br>Ref: %{customdata[3]:.4f}<br><b>Ref - NIR: %{y:.4f}</b><br>%{customdata[4]}<extra></extra>'
  }];

  const xmin = x.length ? x[0] : 0;
  const xmax = x.length ? x[x.length - 1] : 1;

  // Mostra no máximo 12 marcações de data/hora no eixo para não poluir a leitura.
  const maxTicks = 12;
  const tickvals = [];
  if (linhas.length <= maxTicks) {
    for (let i = 0; i < linhas.length; i++) tickvals.push(i);
  } else {
    for (let j = 0; j < maxTicks; j++) {
      const idx = Math.round(j * (linhas.length - 1) / (maxTicks - 1));
      if (!tickvals.includes(idx)) tickvals.push(idx);
    }
  }
  const ticktext = tickvals.map(i => {
    const txt = fmtData(linhas[i]?.DataHora);
    if (txt === '—') return 'Sem data/hora';
    const partes = txt.split(' ');
    return partes.length >= 2 ? `${partes[0]}<br>${partes[1]}` : txt;
  });
  if (Number.isFinite(limites.lsc)) traces.push({x:[xmin,xmax], y:[limites.lsc,limites.lsc], mode:'lines', name:'LSC', line:{color:CORES.vermelho,width:1.8,dash:'dot'}, hoverinfo:'skip'});
  if (Number.isFinite(limites.lic)) traces.push({x:[xmin,xmax], y:[limites.lic,limites.lic], mode:'lines', name:'LIC', line:{color:CORES.vermelho,width:1.8,dash:'dot'}, hoverinfo:'skip'});

  const annotations = [];
  if (Number.isFinite(limites.lsc)) annotations.push({x:xmax,y:limites.lsc,text:`LSC ${limites.lsc.toFixed(4)}`,showarrow:false,xanchor:'right',yshift:10,font:{size:10,color:CORES.vermelho},bgcolor:'rgba(255,255,255,.86)'});
  if (Number.isFinite(limites.lic)) annotations.push({x:xmax,y:limites.lic,text:`LIC ${limites.lic.toFixed(4)}`,showarrow:false,xanchor:'right',yshift:-10,font:{size:10,color:CORES.vermelho},bgcolor:'rgba(255,255,255,.86)'});

  const layout = {
    title:{text:`${selEnsaio.value} · Diferença Ref − NIR`,x:0,xanchor:'left',font:{size:17,color:CORES.navy}},
    xaxis:{title:'Tempo de análise',type:'linear',tickmode:'array',tickvals:tickvals,ticktext:ticktext,tickangle:-35,automargin:true,showgrid:true,gridcolor:'#EEF1F4',zeroline:false,linecolor:CORES.borda},
    yaxis:{title:'Diferença',showgrid:true,gridcolor:'#EEF1F4',zeroline:false,linecolor:CORES.borda},
    height:500, margin:{l:55,r:30,t:62,b:95}, hovermode:'closest', plot_bgcolor:'#FFFFFF', paper_bgcolor:'#FFFFFF',
    showlegend:false, font:{color:'#2B3542'}, annotations,
    shapes:[{type:'line',x0:xmin,x1:xmax,y0:0,y1:0,line:{color:'#AEB8C2',width:1,dash:'dot'}}]
  };
  Plotly.react('grafico', traces, layout, {displaylogo:false,responsive:true});
}

function atualizarTabela(linhas, limites) {
  const corpo = document.getElementById('corpoTabela');
  corpo.innerHTML = linhas.map(r => {
    const cls = r.Dentro_Faixa ? 'dentro' : 'fora';
    return `<tr><td>${r.Id}</td><td>${fmtData(r.DataHora)}</td><td>${fmtTabela(r.REF)}</td><td>${fmtTabela(r.NIR)}</td><td>${fmtTabela(r.Diferenca)}</td><td>${fmtTabela(limites.lic)}</td><td>${fmtTabela(limites.lsc)}</td><td class="status ${cls}">${r.Status_Faixa}</td></tr>`;
  }).join('');
  const fora = linhas.filter(r => !r.Dentro_Faixa).length;
  document.getElementById('notaBoletim').textContent = `${linhas.length} observações · ${fora} fora dos limites de controle · diferença calculada como Ref − NIR`;
}

function atualizarTudo() {
  const brutas = linhasSelecionadas();
  const limites = calcularLimites(brutas);
  const linhas = comStatus(brutas, limites);
  atualizarMeta(linhas, limites);
  atualizarKpis(linhas, limites);
  atualizarBias(linhas);
  atualizarGrafico(linhas, limites);
  atualizarTabela(linhas, limites);
}

function iniciar() {
  const produtos = ordenarTexto(unicos(BASE.map(r => r.Produto)));
  preencherSelect(selProduto, produtos, INICIAL.produto);
  atualizarOpcoesEnsaios(INICIAL.ensaio);
  atualizarOpcoesVerificacoes(INICIAL.verificacao || 'Todos');
  atualizarTudo();
}

selProduto.addEventListener('change', () => { atualizarOpcoesEnsaios(); atualizarOpcoesVerificacoes('Todos'); atualizarTudo(); });
selEnsaio.addEventListener('change', () => { atualizarOpcoesVerificacoes('Todos'); atualizarTudo(); });
selVerificacao.addEventListener('change', atualizarTudo);

iniciar();
</script>
</body>
</html>"""

    html = (
        html_template
        .replace("__VERDE__", VERDE)
        .replace("__AZUL__", AZUL)
        .replace("__NAVY__", NAVY)
        .replace("__VERMELHO__", VERMELHO)
        .replace("__MUTED__", MUTED)
        .replace("__FUNDO__", FUNDO)
        .replace("__BORDA__", BORDA)
        .replace("__PLOTLY_JS__", plotly_js)
        .replace("__ATUALIZADO__", atualizado_txt)
        .replace("__DADOS_JSON__", dados_json)
        .replace("__INICIAL_JSON__", inicial_json)
    )
    return html.encode("utf-8")

def main() -> None:
    st.set_page_config(page_title="Carta controle NIR DS-3 — Curvas Globais", page_icon="📈", layout="wide")

    css_app = (
        "<style>\n:root {"
        + f"--verde:{VERDE}; --azul:{AZUL}; --azul-escuro:{AZUL_ESCURO}; --navy:{NAVY}; "
        + f"--amarelo:{AMARELO}; --vermelho:{VERMELHO}; --muted:{MUTED}; "
        + f"--fundo:{FUNDO}; --card:{CARD}; --borda:{BORDA}; --grid:{GRID};"
        + "}\n"
        + r"""
        @import url('https://fonts.googleapis.com/css2?family=Montserrat:ital,wght@0,400;0,500;0,600;0,700;0,800;1,700;1,800&family=JetBrains+Mono:wght@400;500;600;700&display=swap');
        html, body, [class*="css"] {font-family:'Montserrat',system-ui,-apple-system,'Segoe UI',sans-serif;}
        .stApp {background:var(--fundo); color:#2B3542;}
        .block-container {max-width:1560px; padding-top:1.15rem; padding-bottom:2.2rem; padding-left:2rem; padding-right:2rem;}
        [data-testid="stSidebar"] {background:#F5F7F9; border-right:1px solid var(--borda);}
        [data-testid="stSidebar"] * {font-family:'Montserrat',sans-serif;}
        h1,h2,h3 {color:var(--navy); letter-spacing:-0.02em;}
        label, .stCaption {color:var(--muted)!important;}
        div[data-baseweb="select"] > div, input {border-radius:6px!important;}
        [data-testid="stDateInput"] input {font-family:'JetBrains Mono',monospace; font-size:12px;}

        .qmp-header {
            background:linear-gradient(118deg,var(--verde) 0%,var(--azul) 64%,var(--azul-escuro) 100%);
            padding:24px 34px 20px; border-radius:12px; display:flex; align-items:center;
            justify-content:space-between; gap:24px; box-shadow:0 1px 2px rgba(12,34,64,.10); margin-bottom:16px;
        }
        .qmp-header-texto {display:flex; flex-direction:column; gap:5px;}
        .qmp-eyebrow {font-family:'JetBrains Mono',monospace; font-size:10.5px; font-weight:600; letter-spacing:.16em; text-transform:uppercase; color:var(--amarelo);}
        .qmp-titulo {font-size:31px; font-weight:800; color:#fff; line-height:1.05; letter-spacing:-.02em;}
        .qmp-titulo em {font-style:italic; font-weight:800;}
        .qmp-meta {font-family:'JetBrains Mono',monospace; font-size:10.5px; font-weight:500; letter-spacing:.04em; text-transform:uppercase; color:rgba(255,255,255,.84);}
        .qmp-slashes {display:flex;gap:7px;flex-shrink:0;}
        .qmp-slashes span {display:block;width:13px;height:52px;transform:skewX(-18deg);background:rgba(255,255,255,.30);}
        .qmp-slashes span:nth-child(2) {background:rgba(255,255,255,.16);}
        .qmp-slashes span:nth-child(3) {background:var(--amarelo);}

        [data-testid="stVerticalBlockBorderWrapper"] {
            background:#fff; border:1px solid var(--borda)!important; border-radius:12px!important;
            box-shadow:0 1px 2px rgba(12,34,64,.06); overflow:hidden;
        }
        .qmp-section-title {font-size:17px;font-weight:800;color:var(--navy);letter-spacing:-.01em;margin-bottom:2px;}
        .qmp-note {font-size:11.5px;color:var(--muted);}
        .qmp-rot {font-family:'JetBrains Mono',monospace;font-size:10px;font-weight:600;letter-spacing:.14em;text-transform:uppercase;color:var(--muted);}

        .qmp-kpi {
            background:#FFFFFF; border:1px solid var(--borda); border-radius:10px;
            box-shadow:0 1px 2px rgba(12,34,64,.05);
            padding:14px 12px 12px; min-height:126px; display:flex; flex-direction:column;
            gap:8px; justify-content:flex-start; align-items:center; text-align:center;
        }
        .qmp-kpi-rotulo {
            font-size:12px; font-weight:500; letter-spacing:0; text-transform:none;
            color:#3E3E3E; line-height:1.25; min-height:30px; display:flex; align-items:flex-end; justify-content:center;
        }
        .qmp-kpi-valor {
            font-family:'Montserrat',system-ui,-apple-system,'Segoe UI',sans-serif;
            font-size:22px; font-weight:600; letter-spacing:0; font-variant-numeric:tabular-nums;
            line-height:1.05; transition:color .15s ease;
        }
        .qmp-kpi-detalhe {
            font-family:'Montserrat',system-ui,-apple-system,'Segoe UI',sans-serif;
            font-size:14px; font-weight:400; color:#6A6A6A; min-height:18px; line-height:1.2;
        }
        .qmp-kpi-ok .qmp-kpi-valor {color:#009A44;}
        .qmp-kpi-alerta .qmp-kpi-valor {color:var(--vermelho);}
        .qmp-kpi-neutro .qmp-kpi-valor {color:var(--muted);}
        .qmp-kpi-titulo {min-height:96px; justify-content:center;}
        .qmp-kpi-titulo .qmp-kpi-rotulo {min-height:0; align-items:center; font-size:13px; font-weight:600;}
        .qmp-kpi-titulo .qmp-kpi-valor {color:var(--navy);}

        [data-testid="stDownloadButton"] button {background:var(--navy);color:#fff;border:1px solid var(--navy);border-radius:6px;font-weight:700;}
        [data-testid="stDownloadButton"] button:hover {background:var(--azul);border-color:var(--azul);color:#fff;}
        [data-testid="stDataFrame"] {border:1px solid var(--borda);border-radius:8px;overflow:hidden;}
        .qmp-table-wrap {
            width:100%; max-height:520px; overflow:auto; border:1px solid var(--borda);
            border-radius:8px; background:#fff; margin-top:10px;
        }
        .qmp-table {width:100%; border-collapse:collapse; table-layout:fixed; font-size:12px;}
        .qmp-table th {
            position:sticky; top:0; z-index:2; background:var(--navy); color:#fff;
            text-align:center!important; vertical-align:middle!important; padding:9px 7px;
            font-weight:700; border-bottom:1px solid var(--borda);
        }
        .qmp-table td {
            text-align:center!important; vertical-align:middle!important; padding:8px 7px;
            border-bottom:1px solid #E8ECF0; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;
            font-family:'JetBrains Mono',monospace;
        }
        .qmp-table td.qmp-status {font-family:'Montserrat',sans-serif;font-weight:700;}
        .qmp-table td.qmp-dentro {color:#00702C;background:#E9F5EE;}
        .qmp-table td.qmp-fora {color:var(--vermelho);background:#FDECEA;}
        </style>
        """
    )
    st.markdown(css_app, unsafe_allow_html=True)

    arquivos = st.sidebar.file_uploader(
        "Planilhas Excel",
        type=["xlsx", "xlsm"],
        accept_multiple_files=True,
        help=(
            "Selecione 1 ou 2 arquivos. O app procura a aba 'Amostras', valida se "
            "as estruturas são compatíveis e consolida os pares NIR/Ref automaticamente."
        ),
    )

    if not arquivos:
        st.markdown(
            """
            <div class="qmp-header">
              <div class="qmp-header-texto">
                <div class="qmp-titulo">Carta controle NIR DS-3 — Curvas Globais</div>
              </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.info("Carregue uma ou duas planilhas Excel na barra lateral para iniciar.")
        st.stop()

    if len(arquivos) > 2:
        st.error("Selecione no máximo dois arquivos Excel.")
        st.stop()

    try:
        df, aba, duplicatas_removidas = consolidar_planilhas(arquivos)
    except Exception as exc:
        st.error(f"Não foi possível consolidar as planilhas: {exc}")
        st.stop()

    if duplicatas_removidas > 0:
        st.info(
            f"Consolidação concluída: {duplicatas_removidas} linha(s) exatamente duplicada(s) "
            "foram ignoradas para evitar dupla contagem."
        )

    if df.shape[1] <= IDX_INICIO_ENSAIOS:
        st.error("A planilha não possui colunas suficientes para localizar os ensaios a partir da coluna I.")
        st.stop()

    col_produto = df.columns[IDX_PRODUTO]
    produtos = sorted(df[col_produto].dropna().astype(str).str.strip().unique().tolist())
    pares = detectar_pares_ensaio(df)

    if not produtos:
        st.error("Nenhum produto foi encontrado na coluna B.")
        st.stop()
    if not pares:
        st.error("Nenhum par NIR / Ref foi encontrado a partir da coluna I.")
        st.stop()

    # Tratamento padronizado: não é mais exibido na barra lateral.
    ignorar_marcador = True
    marcador = 42.0
    unidade = ""
    mostrar_lc = False

    header_placeholder = st.empty()

    # ------------------------------------------------------------------
    # Painel de filtros das análises
    # ------------------------------------------------------------------
    # Os três filtros operacionais ficam na mesma faixa: Produto, Ensaio e
    # Verificação. Cada verificação corresponde a um bloco cronológico fixo
    # de até 30 pares NIR/Ref, definido antes de qualquer recorte da análise.
    with st.container(border=True):
        st.markdown('<div class="qmp-section-title">Filtros das análises</div>', unsafe_allow_html=True)
        st.markdown('<div class="qmp-note">Selecione o produto, o ensaio e a verificação.</div>', unsafe_allow_html=True)

        f1, f2, f3 = st.columns([1.05, 1.45, 1.15])
        with f1:
            produto = st.selectbox("Produto", produtos, key="produto_sel")

        df_produto = df[df[col_produto].astype(str).str.strip() == produto].copy()
        pares_disponiveis: list[ParEnsaio] = []
        contagens: dict[str, int] = {}
        for p in pares:
            _, _, mask = mascara_valida(df_produto[p.coluna_nir], df_produto[p.coluna_ref], ignorar_marcador, marcador)
            n = int(mask.sum())
            if n > 0:
                pares_disponiveis.append(p)
                contagens[p.nome_exibicao] = n

        if not pares_disponiveis:
            st.warning("Não há pares NIR/Ref válidos para o produto selecionado.")
            st.stop()

        rotulos = [f"{p.nome_exibicao} ({contagens[p.nome_exibicao]} pares)" for p in pares_disponiveis]
        if st.session_state.get("ensaio_sel") not in rotulos:
            st.session_state["ensaio_sel"] = rotulos[0]
        with f2:
            rotulo_escolhido = st.selectbox("Ensaio", rotulos, key="ensaio_sel")
        par = pares_disponiveis[rotulos.index(rotulo_escolhido)]

        dados = montar_dados_ensaio(df_produto, par, ignorar_marcador, marcador)
        if dados.empty:
            st.warning("Não há dados válidos para a combinação selecionada.")
            st.stop()

        numeros_verificacao = sorted(dados["Verificacao"].dropna().astype(int).unique().tolist())
        opcoes_verificacao: list[str | int] = ["Todos"] + numeros_verificacao
        if st.session_state.get("verificacao_sel") not in opcoes_verificacao:
            st.session_state["verificacao_sel"] = "Todos"

        def formatar_verificacao(v: str | int) -> str:
            if v == "Todos":
                return f"Todas ({len(numeros_verificacao)} verificações)"
            numero = int(v)
            inicio = (numero - 1) * PARES_POR_VERIFICACAO + 1
            fim_teorico = numero * PARES_POR_VERIFICACAO
            bloco = dados[dados["Verificacao"] == numero]
            quantidade = len(bloco)
            fim_real = int(bloco["Id"].max()) if quantidade else fim_teorico
            status = "completa" if quantidade >= PARES_POR_VERIFICACAO else "em andamento"
            return f"Verificação {numero} · pares {inicio}–{fim_real} · {quantidade}/{PARES_POR_VERIFICACAO} · {status}"

        with f3:
            verificacao = st.selectbox(
                "Verificação",
                opcoes_verificacao,
                key="verificacao_sel",
                format_func=formatar_verificacao,
                help=(
                    f"Cada verificação agrupa {PARES_POR_VERIFICACAO} pares sequenciais: "
                    "1–30 = Verificação 1, 31–60 = Verificação 2, 61–90 = Verificação 3..."
                ),
            )

    # Os demais controles permanecem com comportamento padrão: período completo,
    # destaque de desvios ativo e limites fixos quando cadastrados; nos demais
    # casos, utiliza-se Carta I (amplitude móvel) sobre o recorte selecionado.
    dados_periodo = dados.copy()
    destacar_desvios = True

    limite_predefinido = obter_limites_predefinidos(produto, par.nome_exibicao)
    if limite_predefinido is not None:
        _, _, metodo_limites = limite_predefinido
    else:
        metodo_limites = "Carta I (amplitude móvel)"

    # ------------------------------------------------------------------
    # Verificação selecionada e limites
    # ------------------------------------------------------------------
    dados_analise = filtrar_verificacao(dados_periodo, verificacao)
    if dados_analise.empty:
        st.warning("A verificação selecionada não possui pares válidos.")
        st.stop()

    if limite_predefinido is not None:
        lic, lsc, _ = limite_predefinido
        centro = 0.0
        sigma = np.nan
    else:
        centro, lic, lsc, sigma = calcular_limites_i_mr(dados_analise["Diferenca"])

    dados_exibidos = aplicar_classificacoes(dados_analise, centro, lic, lsc)

    # Histórico completo de Bias é calculado sobre a sequência integral de
    # verificações do Produto + Ensaio, independentemente do filtro atual.
    historico_bias = calcular_historico_bias(dados, produto, par.nome_exibicao)
    estado_bias = obter_estado_bias(historico_bias, verificacao)

    primeira_data = dados_exibidos["DataHora"].min()
    ultima_data = dados_exibidos["DataHora"].max()
    atualizado_em = datetime.now()
    periodo_txt = (
        f"Período {primeira_data.strftime('%d/%m/%y')} — {ultima_data.strftime('%d/%m/%y')}"
        if pd.notna(primeira_data) and pd.notna(ultima_data) else "Período sem data válida"
    )

    header_placeholder.markdown(
        html_header(produto, par.nome_exibicao, unidade, periodo_txt, len(dados_exibidos), aba),
        unsafe_allow_html=True,
    )

    # ------------------------------------------------------------------
    # KPIs — somente os cartões dos limites de controle
    # ------------------------------------------------------------------
    kpis = indicadores(dados_exibidos)

    media_dif = float(kpis["media"])
    pct_dentro = float(kpis["pct_dentro"])
    pct_menor = float(kpis["pct_menor"])
    pct_maior = float(kpis["pct_maior"])
    pct_zero = float(kpis["pct_zero"])

    # Regras de formatação condicional solicitadas:
    # LC1: média fora de LIC–LSC = vermelho; dentro da faixa = verde.
    # LC2: % dentro da faixa < 70% = vermelho; >= 70% = verde.
    # LC3: cada percentual > 70% = vermelho; <= 70% = verde.
    status_lc1 = status_kpi_intervalo(media_dif, lic, lsc)
    status_lc2 = status_kpi_minimo(pct_dentro, 70.0)
    status_lc3_menor = status_kpi_maximo(pct_menor, 70.0)
    status_lc3_maior = status_kpi_maximo(pct_maior, 70.0)
    status_lc3_zero = status_kpi_maximo(pct_zero, 70.0)

    k1, k2, k3, k4, k5 = st.columns(5)
    with k1:
        st.markdown(
            html_kpi("Média de Dif", fmt_num(media_dif, 2), "Limite de Controle 1", status_lc1),
            unsafe_allow_html=True,
        )
    with k2:
        st.markdown(
            html_kpi("% Dentro da Faixa", fmt_pct(pct_dentro), "Limite de Controle 2", status_lc2),
            unsafe_allow_html=True,
        )
    with k3:
        st.markdown(
            html_kpi("Diferença < 0 (%)", fmt_pct(pct_menor), "Limite de Controle 3", status_lc3_menor),
            unsafe_allow_html=True,
        )
    with k4:
        st.markdown(
            html_kpi("Diferença > 0 (%)", fmt_pct(pct_maior), "Limite de Controle 3", status_lc3_maior),
            unsafe_allow_html=True,
        )
    with k5:
        st.markdown(
            html_kpi("Diferença = 0 (%)", fmt_pct(pct_zero), "Limite de Controle 3", status_lc3_zero),
            unsafe_allow_html=True,
        )

    # Segunda linha de cartões: ocupa as três colunas centrais da página.
    st.write("")
    _, b1, b2, b3, _ = st.columns(5)
    with b1:
        st.markdown(
            html_cartao_titulo("Bias Atuais", fmt_bias(float(estado_bias["bias_atual"]))),
            unsafe_allow_html=True,
        )
    with b2:
        st.markdown(
            html_cartao_titulo("Ajustes", fmt_bias(float(estado_bias["ajuste"]))),
            unsafe_allow_html=True,
        )
    with b3:
        st.markdown(
            html_cartao_titulo("Novas Bias", fmt_bias(float(estado_bias["nova_bias"]))),
            unsafe_allow_html=True,
        )

    # ------------------------------------------------------------------
    # Carta
    # ------------------------------------------------------------------
    st.write("")
    with st.container(border=True):
        st.markdown('<div class="qmp-section-title">Carta Controle</div>', unsafe_allow_html=True)
        st.markdown('<div class="qmp-note">Cada ponto representa Ref − NIR. Pontos fora dos limites são destacados em vermelho.</div>', unsafe_allow_html=True)

        fig = criar_grafico_powerbi(
            dados=dados_exibidos,
            ensaio=par.nome_exibicao,
            unidade=unidade,
            lic=lic,
            lsc=lsc,
            centro=centro,
            mostrar_lc=mostrar_lc,
            destacar_desvios=destacar_desvios,
        )
        st.plotly_chart(fig, use_container_width=True, config={"displaylogo": False})

    # ------------------------------------------------------------------
    # Boletim analítico / tabela
    # ------------------------------------------------------------------
    st.write("")
    with st.container(border=True):
        st.markdown('<div class="qmp-section-title">Boletim analítico</div>', unsafe_allow_html=True)
        n_fora = int((~dados_exibidos["Dentro_Faixa"]).sum())
        st.markdown(
            f'<div class="qmp-note">{len(dados_exibidos)} observações · {n_fora} fora dos limites de controle · diferença calculada como Ref − NIR</div>',
            unsafe_allow_html=True,
        )

        tabela = dados_exibidos[
            ["Id", "DataHora", "REF", "NIR", "Diferenca", "LIC", "LSC", "Status_Faixa"]
        ].copy()
        tabela.columns = ["Id", "Data/hora", "Ref", "NIR", "Ref - NIR", "LIC", "LSC", "Status"]

        # Renderização HTML própria para garantir centralização real de cabeçalhos
        # e valores em todas as colunas do boletim analítico.
        def fmt_tabela_num(v):
            if pd.isna(v):
                return "—"
            try:
                return f"{float(v):.4f}".replace(".", ",")
            except (TypeError, ValueError):
                return escape(str(v))

        linhas_tabela = []
        for _, row in tabela.iterrows():
            data_hora = row["Data/hora"]
            if pd.isna(data_hora):
                data_txt = "—"
            elif hasattr(data_hora, "strftime"):
                data_txt = data_hora.strftime("%d/%m/%Y %H:%M:%S")
            else:
                data_txt = escape(str(data_hora))

            status = str(row["Status"])
            status_cls = "qmp-fora" if status == "Fora da faixa" else "qmp-dentro"
            linhas_tabela.append(
                "<tr>"
                f"<td>{escape(str(row['Id']))}</td>"
                f"<td>{data_txt}</td>"
                f"<td>{fmt_tabela_num(row['Ref'])}</td>"
                f"<td>{fmt_tabela_num(row['NIR'])}</td>"
                f"<td>{fmt_tabela_num(row['Ref - NIR'])}</td>"
                f"<td>{fmt_tabela_num(row['LIC'])}</td>"
                f"<td>{fmt_tabela_num(row['LSC'])}</td>"
                f'<td class="qmp-status {status_cls}">{escape(status)}</td>'
                "</tr>"
            )

        tabela_boletim_html = (
            '<div class="qmp-table-wrap"><table class="qmp-table">'
            '<thead><tr>'
            '<th>Id</th><th>Data/hora</th><th>Ref</th><th>NIR</th>'
            '<th>Ref - NIR</th><th>LIC</th><th>LSC</th><th>Status</th>'
            '</tr></thead><tbody>'
            + "".join(linhas_tabela)
            + '</tbody></table></div>'
        )
        st.markdown(tabela_boletim_html, unsafe_allow_html=True)

    # ------------------------------------------------------------------
    # Exportação — somente HTML
    # ------------------------------------------------------------------
    base_html_interativa = montar_base_html_interativa(df, pares, ignorar_marcador, marcador)
    html_relatorio = montar_html_relatorio(
        base_interativa=base_html_interativa,
        produto_inicial=produto,
        ensaio_inicial=par.nome_exibicao,
        verificacao_inicial=verificacao,
        atualizado_em=atualizado_em,
    )

    st.write("")
    with st.container(border=True):
        st.markdown('<div class="qmp-section-title">Exportação</div>', unsafe_allow_html=True)
        st.markdown('<div class="qmp-note">Baixe um HTML interativo com filtros de Produto, Ensaio e Verificação.</div>', unsafe_allow_html=True)
        st.download_button(
            "Baixar relatório (HTML)",
            data=html_relatorio,
            file_name="carta_controle_interativa.html",
            mime="text/html",
            use_container_width=True,
        )


if __name__ == "__main__":
    main()
