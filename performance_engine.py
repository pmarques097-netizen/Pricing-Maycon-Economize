from __future__ import annotations

import hashlib
import pickle
import re
import threading
import time
from functools import lru_cache
from pathlib import Path
from typing import Iterable

import pandas as pd


BASE_DIR = Path(__file__).resolve().parent
CACHE_DIR = BASE_DIR / "_cache_pricing"
FILE_CACHE_DIR = CACHE_DIR / "arquivos"
FOLDER_CACHE_DIR = CACHE_DIR / "pastas"
FILE_CACHE_DIR.mkdir(parents=True, exist_ok=True)
FOLDER_CACHE_DIR.mkdir(parents=True, exist_ok=True)

_CACHE_LOCK = threading.RLock()
SUPPORTED_PATTERNS = ("*.xlsx", "*.xls", "*.csv", "*.parquet")

# Estatísticas apenas diagnósticas; não participam de nenhum cálculo de negócio.
_PERF_STATS = {
    "arquivo_hit": 0,
    "arquivo_miss": 0,
    "pasta_hit": 0,
    "pasta_miss": 0,
    "tempo_leitura_fisica_s": 0.0,
}


def limpar_colunas(df: pd.DataFrame) -> pd.DataFrame:
    if not isinstance(df, pd.DataFrame) or df.empty:
        return df
    out = df.copy()
    out.columns = out.columns.astype(str).str.strip()
    return out


def arquivos_atuais_da_pasta(nome_pasta: str) -> list[Path]:
    """Retorna somente arquivos válidos que existem na pasta atual."""
    pasta = BASE_DIR / str(nome_pasta)
    if not pasta.exists() or not pasta.is_dir():
        return []

    arquivos: list[Path] = []
    for padrao in SUPPORTED_PATTERNS:
        arquivos.extend(pasta.glob(padrao))

    ignorados = {
        "analise_pricing.xlsx",
        "analise_pricing.csv",
    }

    return sorted(
        [
            arq
            for arq in arquivos
            if not arq.name.startswith("~$")
            and arq.name.lower() not in ignorados
            and not arq.name.lower().startswith("ignorado_")
        ],
        key=lambda p: p.name.lower(),
    )


def _assinatura_arquivo(arquivo: Path) -> str:
    try:
        st = arquivo.stat()
        return f"{arquivo.resolve()}|{st.st_size}|{st.st_mtime_ns}"
    except OSError:
        return f"{arquivo.resolve()}|indisponivel"


def assinatura_pasta(nome_pasta: str) -> str:
    """Assinatura baseada em nome, tamanho e data de modificação."""
    partes: list[str] = []
    for arq in arquivos_atuais_da_pasta(nome_pasta):
        try:
            stat = arq.stat()
            partes.append(f"{arq.name}|{stat.st_size}|{stat.st_mtime_ns}")
        except OSError:
            continue

    if not partes:
        return "vazio"

    return hashlib.sha256("||".join(partes).encode("utf-8")).hexdigest()


def _hash_kwargs(kwargs: dict) -> str:
    # repr ordenado preserva opções de leitura sem depender de JSON serializável.
    partes = [f"{k}={repr(kwargs[k])}" for k in sorted(kwargs)]
    return hashlib.sha256("|".join(partes).encode("utf-8")).hexdigest()


def _file_cache_path(arquivo: Path, header: int, kwargs: dict) -> Path:
    chave = hashlib.sha256(
        (
            "file_v824|"
            + _assinatura_arquivo(arquivo)
            + f"|header={header}|opts={_hash_kwargs(kwargs)}"
        ).encode("utf-8")
    ).hexdigest()
    return FILE_CACHE_DIR / f"{chave}.pkl"


def _folder_cache_path(nome_pasta: str, header: int, assinatura: str) -> Path:
    chave = hashlib.sha256(
        f"folder_v824|{nome_pasta}|{header}|{assinatura}".encode("utf-8")
    ).hexdigest()
    return FOLDER_CACHE_DIR / f"{chave}.pkl"


def _ler_csv(arquivo: Path, **kwargs) -> pd.DataFrame:
    if kwargs:
        return pd.read_csv(arquivo, **kwargs)

    erros: list[Exception] = []
    for encoding in ("utf-8-sig", "utf-8", "latin1"):
        try:
            return pd.read_csv(
                arquivo,
                sep=None,
                engine="python",
                encoding=encoding,
                low_memory=False,
            )
        except Exception as exc:
            erros.append(exc)
    raise erros[-1]


def _ler_arquivo_fisico(arquivo: Path, header: int = 0, **kwargs) -> pd.DataFrame:
    """
    Leitor físico por formato. Esta é a única parte que efetivamente abre
    o arquivo original quando o cache daquele arquivo ainda não existe.
    """
    ext = arquivo.suffix.lower()

    if ext == ".xls":
        try:
            import xlrd  # noqa: F401
        except ImportError as exc:
            raise RuntimeError(
                "Dependência xlrd ausente. Execute instalar_dependencias.bat "
                "ou executar_pricing.bat para instalar automaticamente."
            ) from exc

        opts = dict(kwargs)
        opts.setdefault("header", header)
        opts.setdefault("engine", "xlrd")
        return pd.read_excel(arquivo, **opts)

    if ext in {".xlsx", ".xlsm"}:
        opts = dict(kwargs)
        opts.setdefault("header", header)
        opts.setdefault("engine", "openpyxl")
        return pd.read_excel(arquivo, **opts)

    if ext == ".csv":
        return _ler_csv(arquivo, **kwargs)

    if ext == ".parquet":
        return pd.read_parquet(arquivo, **kwargs)

    raise ValueError(f"Formato não suportado: {arquivo.name}")


def ler_arquivo_cacheado(
    arquivo: str | Path,
    header: int = 0,
    **kwargs,
) -> pd.DataFrame:
    """
    Cache incremental por ARQUIVO.

    Se apenas um novo arquivo entrar em VENDA_TESTE, somente ele é aberto.
    Todos os arquivos antigos são recuperados do cache binário individual.
    A função sempre devolve uma cópia para impedir mutação acidental do cache.
    """
    arquivo = Path(arquivo)
    caminho = _file_cache_path(arquivo, int(header), kwargs)

    with _CACHE_LOCK:
        if caminho.exists():
            try:
                with caminho.open("rb") as f:
                    obj = pickle.load(f)
                if isinstance(obj, pd.DataFrame):
                    _PERF_STATS["arquivo_hit"] += 1
                    return obj.copy()
            except Exception:
                try:
                    caminho.unlink()
                except OSError:
                    pass

    _PERF_STATS["arquivo_miss"] += 1
    t0 = time.perf_counter()
    resultado = _ler_arquivo_fisico(arquivo, int(header), **kwargs)
    _PERF_STATS["tempo_leitura_fisica_s"] += time.perf_counter() - t0

    if isinstance(resultado, pd.DataFrame):
        with _CACHE_LOCK:
            try:
                tmp = caminho.with_suffix(".tmp")
                with tmp.open("wb") as f:
                    pickle.dump(resultado, f, protocol=pickle.HIGHEST_PROTOCOL)
                tmp.replace(caminho)
            except Exception as exc:
                print(f"[performance_engine] Cache de arquivo não salvo: {exc}")
        return resultado.copy()

    return resultado


def _ler_pasta_incremental(nome_pasta: str, header: int = 0) -> pd.DataFrame:
    bases: list[pd.DataFrame] = []

    for arquivo in arquivos_atuais_da_pasta(nome_pasta):
        try:
            df = limpar_colunas(ler_arquivo_cacheado(arquivo, header=header))
            if not isinstance(df, pd.DataFrame):
                continue
            df["Arquivo_Origem"] = arquivo.name
            bases.append(df)
        except Exception as exc:
            print(f"[performance_engine] Erro ao ler {arquivo.name}: {exc}")

    if not bases:
        arquivos_existentes = arquivos_atuais_da_pasta(nome_pasta)
        if arquivos_existentes:
            nomes = ", ".join(a.name for a in arquivos_existentes[:5])
            raise RuntimeError(
                f"Nenhum arquivo da pasta {nome_pasta} pôde ser lido. "
                f"Arquivos encontrados: {nomes}. Verifique dependências Excel."
            )
        return pd.DataFrame()

    return limpar_colunas(pd.concat(bases, ignore_index=True, sort=False))


@lru_cache(maxsize=24)
def _carregar_memoria(
    nome_pasta: str,
    header: int,
    assinatura: str,
) -> pd.DataFrame:
    caminho = _folder_cache_path(nome_pasta, header, assinatura)

    with _CACHE_LOCK:
        if caminho.exists():
            try:
                with caminho.open("rb") as f:
                    objeto = pickle.load(f)
                if isinstance(objeto, pd.DataFrame):
                    _PERF_STATS["pasta_hit"] += 1
                    return objeto
            except Exception:
                try:
                    caminho.unlink()
                except OSError:
                    pass

    _PERF_STATS["pasta_miss"] += 1
    resultado = _ler_pasta_incremental(nome_pasta, header)

    with _CACHE_LOCK:
        try:
            temporario = caminho.with_suffix(".tmp")
            with temporario.open("wb") as f:
                pickle.dump(resultado, f, protocol=pickle.HIGHEST_PROTOCOL)
            temporario.replace(caminho)
        except Exception as exc:
            print(f"[performance_engine] Cache da pasta não salvo: {exc}")

    return resultado


def carregar_pasta_excel(nome_pasta: str, header: int = 0) -> pd.DataFrame:
    assinatura = assinatura_pasta(nome_pasta)
    return _carregar_memoria(
        str(nome_pasta),
        int(header),
        assinatura,
    ).copy()


def _normalizar_ean(valor: object) -> str:
    texto = re.sub(r"\D", "", str(valor).replace(".0", ""))
    return texto.strip()


def _localizar_coluna(df: pd.DataFrame, nomes: Iterable[str]) -> str | None:
    mapa = {str(c).strip().lower(): str(c) for c in df.columns}
    for nome in nomes:
        chave = str(nome).strip().lower()
        if chave in mapa:
            return mapa[chave]
    return None


def deduplicar_pesquisa_mercado(df: pd.DataFrame) -> pd.DataFrame:
    if not isinstance(df, pd.DataFrame) or df.empty:
        return df

    out = limpar_colunas(df)

    col_ean = _localizar_coluna(
        out,
        ["EAN", "EAN (GTIN)", "GTIN", "Código de Barras", "Codigo de Barras"],
    )
    col_farmacia = _localizar_coluna(
        out,
        ["Farmácia", "Farmacia", "Loja", "Nome Fantasia", "Estabelecimento"],
    )
    col_data = _localizar_coluna(
        out,
        ["Data Emissão", "Data_Emissao", "Data Pesquisa", "Data_Pesquisa", "Data"],
    )
    col_preco = _localizar_coluna(
        out,
        ["Preço (R$)", "Preço", "Preco", "Preco_Venda"],
    )

    if col_ean:
        out["_EAN_DEDUP"] = out[col_ean].map(_normalizar_ean)

    if col_data:
        out["_DATA_DEDUP"] = pd.to_datetime(
            out[col_data],
            errors="coerce",
            dayfirst=True,
        )

    subset: list[str] = []
    for coluna in ("_EAN_DEDUP", col_farmacia, "_DATA_DEDUP", col_preco):
        if coluna and coluna in out.columns:
            subset.append(coluna)

    if subset:
        if "_DATA_DEDUP" in out.columns:
            out = out.sort_values("_DATA_DEDUP", ascending=False)
        out = out.drop_duplicates(subset=subset, keep="first")
    else:
        out = out.drop_duplicates(keep="first")

    return out.drop(
        columns=["_EAN_DEDUP", "_DATA_DEDUP"],
        errors="ignore",
    ).reset_index(drop=True)


def carregar_historico() -> pd.DataFrame:
    historico = carregar_pasta_excel("VENDA_TESTE", header=0)
    if historico.empty:
        return historico

    historico = deduplicar_pesquisa_mercado(historico)

    if "Data Emissão" in historico.columns:
        historico["Data Emissão"] = pd.to_datetime(
            historico["Data Emissão"],
            errors="coerce",
            dayfirst=True,
        )

    return historico


def carregar_compra() -> pd.DataFrame:
    compra = carregar_pasta_excel("COMPRA_TESTE", header=2)
    if compra.empty:
        return compra

    return limpar_colunas(compra).rename(
        columns={
            "Rótulos de Linha": "Marca",
            "Soma de Valor Líquido": "Valor_Liquido",
            "PART%": "Participacao",
            "GT%": "Acumulado",
        }
    )


def carregar_venda_rede() -> pd.DataFrame:
    return carregar_pasta_excel("VENDA_FINAL_TESTE", header=0)


def carregar_estoque() -> pd.DataFrame:
    return carregar_pasta_excel("ESTOQUE_TESTE", header=0)


def identificar_rede(nome: object) -> str:
    nome_original = str(nome).strip()
    nome_base = nome_original.upper()

    regras = {
        "RAIADROGASIL": "Drogasil",
        "DROGASIL": "Drogasil",
        "DROGA RAIA": "Droga Raia",
        "RAIA": "Droga Raia",
        "PAGUE MENOS": "Pague Menos",
        "ULTRAPOPULAR": "Ultra Popular",
        "ULTRA POPULAR": "Ultra Popular",
        "SAO JOAO": "São João",
        "SÃO JOÃO": "São João",
        "PANVEL": "Panvel",
        "NISSEI": "Nissei",
        "EXTRAFARMA": "Extrafarma",
        "PACHECO": "Pacheco",
        "VENANCIO": "Venancio",
        "VENÂNCIO": "Venancio",
        "DROGARIA SAO PAULO": "Drogaria São Paulo",
        "DROGARIA SÃO PAULO": "Drogaria São Paulo",
        "DROGARIAS PACHECO": "Pacheco",
        "PRECO POPULAR": "Preço Popular",
        "PREÇO POPULAR": "Preço Popular",
        "INDIANA": "Indiana",
        "ARAUJO": "Araujo",
        "ARAÚJO": "Araujo",
        "DROGAL": "Drogal",
        "DROGASMIL": "Drogasmil",
        "GLOBO": "Globo",
        "ZANOL": "Zanol e Thomaz",
        "THOMAZ": "Zanol e Thomaz",
        "TRIANGULO": "Triangulo",
        "TRIÂNGULO": "Triangulo",
        "BRASIFARMA": "Brasifarma",
    }

    for chave, rede in regras.items():
        if chave in nome_base:
            return rede

    remover = [
        "FARMACIA", "FARMÁCIA", "DROGARIA", "DROGARIAS", "DROGA",
        "MEDICAMENTOS", "MEDICAMENTO", "COMERCIO", "COMÉRCIO",
        "PRODUTOS", "FARMACEUTICOS", "FARMACÊUTICOS", "PERFUMARIA",
        "PERFUMARIAS", "COSMETICOS", "COSMÉTICOS", "LTDA", "EIRELI",
        "ME", "EPP", "SA", "S/A", "S.A.", "MATRIZ", "FILIAL",
        "LOJA", "CIA", "COMPANHIA",
    ]

    resumo = nome_base
    for palavra in remover:
        resumo = resumo.replace(palavra, " ")

    resumo = " ".join(resumo.split())
    if not resumo:
        resumo = nome_original

    return " ".join(resumo.title().split()[:3])


def curva_abc(df: pd.DataFrame) -> pd.DataFrame:
    if not isinstance(df, pd.DataFrame) or df.empty:
        return pd.DataFrame()

    produto_col = (
        "Descricao_Unica"
        if "Descricao_Unica" in df.columns
        else "Produto"
        if "Produto" in df.columns
        else None
    )

    if not produto_col:
        return pd.DataFrame()

    valor_col = None
    for candidato in (
        "Venda_Final_Faturamento",
        "Valor Total",
        "Valor_Total",
        "Faturamento",
        "Ganho_Potencial",
    ):
        if candidato in df.columns:
            valor_col = candidato
            break

    if not valor_col:
        return pd.DataFrame()

    base = df.copy()
    base[valor_col] = pd.to_numeric(base[valor_col], errors="coerce").fillna(0)

    ranking = (
        base.groupby(produto_col, dropna=False)[valor_col]
        .sum()
        .reset_index()
        .rename(columns={produto_col: "Produto", valor_col: "Valor_ABC"})
        .sort_values("Valor_ABC", ascending=False)
    )

    total = ranking["Valor_ABC"].sum()
    if total <= 0:
        ranking["Perc_Acum"] = 0.0
        ranking["ABC"] = "C"
        return ranking

    ranking["Perc_Acum"] = ranking["Valor_ABC"].cumsum() / total
    ranking["ABC"] = ranking["Perc_Acum"].apply(
        lambda x: "A" if x <= 0.80 else "B" if x <= 0.95 else "C"
    )
    return ranking


def auditoria_pesquisa_atual() -> pd.DataFrame:
    registros: list[dict[str, object]] = []

    for arquivo in arquivos_atuais_da_pasta("VENDA_TESTE"):
        item: dict[str, object] = {
            "Arquivo": arquivo.name,
            "Tamanho_MB": round(arquivo.stat().st_size / (1024 * 1024), 3),
            "Modificado_em": pd.to_datetime(arquivo.stat().st_mtime, unit="s"),
            "Linhas": 0,
            "Status": "OK",
        }

        try:
            amostra = ler_arquivo_cacheado(arquivo, header=0)
            item["Linhas"] = int(len(amostra))
        except Exception as exc:
            item["Status"] = f"ERRO: {exc}"

        registros.append(item)

    return pd.DataFrame(registros)


def limpar_caches_antigos(max_arquivos: int = 240) -> None:
    """
    Mantém espaço controlado sem apagar arquivos recentes necessários à leitura
    incremental. O limite maior é proposital porque VENDA_TESTE possui dezenas
    de arquivos e cada um pode ter seu próprio cache.
    """
    try:
        arquivos = sorted(
            list(FILE_CACHE_DIR.glob("*.pkl")) + list(FOLDER_CACHE_DIR.glob("*.pkl")),
            key=lambda p: p.stat().st_mtime,
            reverse=True,
        )
        for arquivo in arquivos[max_arquivos:]:
            try:
                arquivo.unlink()
            except OSError:
                pass
    except Exception:
        pass


def diagnosticar_bases_principais():
    resultado = {}
    configuracao = {
        "VENDA_TESTE": 0,
        "COMPRA_TESTE": 2,
        "ESTOQUE_TESTE": 0,
        "VENDA_FINAL_TESTE": 0,
    }

    for pasta, header in configuracao.items():
        try:
            df = carregar_pasta_excel(pasta, header=header)
            resultado[pasta] = {
                "ok": True,
                "linhas": int(len(df)),
                "colunas": int(len(df.columns)),
            }
        except Exception as exc:
            resultado[pasta] = {
                "ok": False,
                "linhas": 0,
                "colunas": 0,
                "erro": str(exc),
            }

    return resultado


def diagnostico_performance() -> dict:
    """Resumo técnico para futura tela administrativa."""
    return {
        **_PERF_STATS,
        "cache_arquivos": len(list(FILE_CACHE_DIR.glob("*.pkl"))),
        "cache_pastas": len(list(FOLDER_CACHE_DIR.glob("*.pkl"))),
        "cache_mb": round(
            sum(p.stat().st_size for p in CACHE_DIR.rglob("*.pkl")) / (1024 * 1024),
            2,
        ),
    }
