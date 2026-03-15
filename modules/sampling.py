"""
Fórmula de amostragem estatística conforme Manual de Gerenciamento PGF (seção 5).

Fórmula para populações finitas (padrão: confiança 95%, margem de erro 5%):

    n₀ = Z² · p · (1 - p) / E²   → tamanho para população infinita
    n  = n₀ / (1 + (n₀ - 1) / N) → correção para população finita (N)

Onde:
    Z = 1,96  (valor crítico para nível de confiança de 95%)
    p = 0,50  (proporção esperada — valor conservador que maximiza a amostra)
    E = 0,05  (margem de erro de 5%)
    N = total de tarefas a auditar
"""
from __future__ import annotations

import math

import pandas as pd


def calcular_amostra(
    N: int,
    Z: float = 1.96,
    p: float = 0.50,
    E: float = 0.05,
) -> int:
    """
    Calcula o tamanho da amostra para população finita N.

    Parâmetros:
        N  — tamanho da população (total de tarefas triadas)
        Z  — valor-z para o nível de confiança (padrão 1,96 = 95%)
        p  — proporção esperada (padrão 0,5 = máxima variabilidade)
        E  — margem de erro (padrão 0,05 = 5%)

    Retorna min(N, ceil(n)) — audita tudo se a fórmula resultar em n ≥ N.
    """
    if N <= 0:
        raise ValueError("N deve ser maior que zero.")
    n0 = (Z ** 2 * p * (1 - p)) / (E ** 2)   # tamanho para pop. infinita
    n = n0 / (1 + (n0 - 1) / N)               # correção finita
    return min(N, math.ceil(n))


def formula_descricao(N: int, Z: float = 1.96, p: float = 0.50, E: float = 0.05) -> str:
    """Retorna texto explicativo da fórmula com os valores substituídos."""
    n0 = (Z ** 2 * p * (1 - p)) / (E ** 2)
    n = n0 / (1 + (n0 - 1) / N)
    result = min(N, math.ceil(n))
    return (
        f"n₀ = {Z}² × {p} × (1 - {p}) / {E}² = **{n0:.2f}**\n\n"
        f"n = {n0:.2f} / (1 + ({n0:.2f} - 1) / {N}) = **{n:.2f}** → **{result} tarefas**"
    )


def selecionar_amostra(df: pd.DataFrame, n: int, seed: int | None = None) -> pd.DataFrame:
    """
    Seleciona n linhas aleatórias de df sem reposição.
    Retorna o DataFrame com as linhas selecionadas, ordenadas pelo índice original.
    """
    n = min(n, len(df))
    return df.sample(n=n, random_state=seed).sort_index().reset_index(drop=True)


def tabela_referencia() -> list[tuple[int, int]]:
    """Tabela de referência do Anexo III do Manual (Z=1,96, p=0,5, E=5%)."""
    return [
        (50, 45), (100, 80), (200, 132), (300, 169), (400, 196),
        (500, 218), (600, 235), (700, 249), (800, 260), (900, 270),
        (1000, 278), (1500, 306), (2000, 323), (2500, 333), (3000, 341),
    ]
