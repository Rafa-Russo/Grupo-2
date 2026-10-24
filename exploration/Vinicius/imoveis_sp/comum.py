"""Caminhos, estilo dos gráficos e funções compartilhadas pelos notebooks 01 a 05."""

import os
import re
import ssl
import unicodedata
from pathlib import Path

import certifi
import matplotlib.pyplot as plt
import numpy as np
from matplotlib.colors import LinearSegmentedColormap
from matplotlib.ticker import FuncFormatter

RAIZ = next(p for p in [Path.cwd(), *Path.cwd().parents] if (p / "pyproject.toml").exists())
CACHE = RAIZ / "data" / "SP" / "cache"
ANOS_COMPLETOS = [2022, 2023, 2024, 2025]

# ── estilo ────────────────────────────────────────────────────────────────
AZUL, LARANJA, AQUA, VIOLETA = "#2a78d6", "#eb6834", "#1baf7a", "#4a3aa7"
SUPERFICIE, TINTA, TINTA2, MUDO, GRADE = "#fcfcfb", "#0b0b0b", "#52514e", "#898781", "#e1e0d9"
DIVERGENTE = LinearSegmentedColormap.from_list("div", [AZUL, "#f0efec", "#e34948"])

plt.rcParams.update({
    "figure.facecolor": SUPERFICIE, "axes.facecolor": SUPERFICIE, "savefig.facecolor": SUPERFICIE,
    "figure.dpi": 110, "font.family": "sans-serif",
    "font.sans-serif": ["Segoe UI", "DejaVu Sans"], "font.size": 10,
    "axes.titlesize": 12, "axes.titleweight": "bold", "axes.titlecolor": TINTA,
    "axes.titlelocation": "left", "axes.titlepad": 12,
    "axes.labelcolor": TINTA2, "axes.labelsize": 10,
    "axes.edgecolor": "#c3c2b7", "axes.linewidth": 0.8,
    "axes.spines.top": False, "axes.spines.right": False,
    "axes.grid": True, "axes.axisbelow": True,
    "grid.color": GRADE, "grid.linewidth": 0.8,
    "xtick.color": MUDO, "ytick.color": MUDO, "xtick.labelsize": 9, "ytick.labelsize": 9,
    "legend.frameon": False, "lines.linewidth": 2, "lines.markersize": 8,
})


def reais(valor, _=None):
    return f"R$ {valor:,.0f}".replace(",", ".")


FORMATO_REAIS = FuncFormatter(reais)
FORMATO_PERCENTUAL = FuncFormatter(lambda valor, _: f"{valor:.0%}")

# ── texto ─────────────────────────────────────────────────────────────────
ABREVIACOES_BAIRRO = {
    "VL": "VILA", "JD": "JARDIM", "PQ": "PARQUE", "CJ": "CONJUNTO", "CID": "CIDADE",
    "S": "SAO", "STO": "SANTO", "STA": "SANTA", "CH": "CHACARA", "AL": "ALTO",
}


def sem_acento(texto):
    normalizado = unicodedata.normalize("NFKD", str(texto))
    return normalizado.encode("ascii", "ignore").decode().upper().strip()


def normaliza_bairro(nome):
    palavras = re.sub(r"[^A-Z0-9 ]", " ", sem_acento(nome)).split()
    return " ".join(ABREVIACOES_BAIRRO.get(p, p) for p in palavras) or None


def em_unicos(serie, funcao):
    """Aplica `funcao` uma vez por valor distinto (bem mais rápido em colunas repetitivas)."""
    return serie.map({valor: funcao(valor) for valor in serie.dropna().unique()})


# ── crime ─────────────────────────────────────────────────────────────────
VIOLENTOS = ("ROUBO", "LATROCINIO", "HOMICIDIO", "LESAO CORPORAL", "ESTUPRO", "EXTORSAO", "SEQUESTRO")


def grupo_crime(natureza):
    """Trânsito/culposo é testado primeiro: senão 'homicídio culposo' contaria como violento."""
    if "ACIDENTE DE TRANSITO" in natureza or "CULPOS" in natureza:
        return "trânsito / culposo"
    if natureza.startswith("FURTO"):
        return "patrimonial sem violência"
    if any(termo in natureza for termo in VIOLENTOS):
        return "violento"
    return "outros"


def perfil_por_unidade(transacoes, crimes, unidade):
    """Preço mediano e as quatro medidas de crime por `unidade` (bairro ou distrito).

    Crime: média anual dos anos completos. Mantém unidades com ≥ 30 transações e ≥ 30 crimes/ano.
    """
    crime = (
        crimes[crimes["ano"].isin(ANOS_COMPLETOS)]
        .groupby([unidade, "grupo"], observed=True)
        .size()
        .unstack(fill_value=0)
        .div(len(ANOS_COMPLETOS))
    )
    crime.index = crime.index.astype(str)

    preco = transacoes.groupby(unidade).agg(
        preco_m2_mediano=("preco_m2", "median"),
        n_transacoes=("preco_m2", "size"),
        area_mediana=("area_construida", "median"),
        idade_mediana=("idade", "median"),
    )

    perfil = preco.join(crime, how="inner")
    perfil["crimes_total"] = perfil[crime.columns].sum(axis=1)
    perfil["violentos"] = perfil["violento"]
    perfil["share_violento"] = perfil["violentos"] / perfil["crimes_total"]
    perfil["violentos_por_1k"] = perfil["violentos"] / perfil["n_transacoes"] * 1000

    perfil = perfil[perfil["n_transacoes"].ge(30) & perfil["crimes_total"].ge(30)].copy()
    for coluna in ["crimes_total", "violentos", "violentos_por_1k"]:
        perfil[f"log_{coluna}"] = np.log10(1 + perfil[coluna])
    return perfil


MEDIDAS_CRIME = {
    "crimes totais / ano (log)": "log_crimes_total",
    "crimes violentos / ano (log)": "log_violentos",
    "% de crimes que são violentos": "share_violento",
    "violentos por mil transações (log)": "log_violentos_por_1k",
}

FORMULA_OLS = (
    "np.log(preco_m2_mediano) ~ log_violentos + share_violento"
    " + np.log(area_mediana) + idade_mediana + np.log(n_transacoes)"
)


# ── rede ──────────────────────────────────────────────────────────────────
def contexto_https():
    """SSL com certifi; GEOSAMPA_CA_BUNDLE (ou SSL_CERT_FILE…) acrescenta a CA de redes com proxy."""
    contexto = ssl.create_default_context(cafile=certifi.where())
    caminho_ca = next(
        (os.environ[v] for v in ("GEOSAMPA_CA_BUNDLE", "SSL_CERT_FILE", "REQUESTS_CA_BUNDLE", "CURL_CA_BUNDLE")
         if os.environ.get(v)),
        None,
    )
    if caminho_ca:
        caminho_ca = Path(caminho_ca).expanduser()
        if not caminho_ca.is_file():
            raise FileNotFoundError(f"Arquivo de certificados não encontrado: {caminho_ca}")
        contexto.load_verify_locations(cafile=str(caminho_ca))
    return contexto
