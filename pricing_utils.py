"""
Camada de compatibilidade do Eirox Pricing.

O dashboard continua importando as mesmas funções de pricing_utils,
mas toda a leitura e o cache ficam isolados em performance_engine.py.
"""

from performance_engine import (
    assinatura_pasta,
    arquivos_atuais_da_pasta,
    auditoria_pesquisa_atual,
    carregar_compra,
    carregar_estoque,
    carregar_historico,
    carregar_pasta_excel,
    carregar_venda_rede,
    curva_abc,
    deduplicar_pesquisa_mercado,
    identificar_rede,
    limpar_caches_antigos,
    limpar_colunas,
)

__all__ = [
    "assinatura_pasta",
    "arquivos_atuais_da_pasta",
    "auditoria_pesquisa_atual",
    "carregar_compra",
    "carregar_estoque",
    "carregar_historico",
    "carregar_pasta_excel",
    "carregar_venda_rede",
    "curva_abc",
    "deduplicar_pesquisa_mercado",
    "identificar_rede",
    "limpar_caches_antigos",
    "limpar_colunas",
]
