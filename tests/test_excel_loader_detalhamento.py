"""Testes do leitor da planilha Detalhamento Individual PGF."""
from __future__ import annotations

import io

import pandas as pd
import pytest
from openpyxl import Workbook

from modules.excel_loader import (
    COL_DET_ATIVIDADES,
    COL_DET_NUP,
    COL_DET_RESPONSAVEL,
    COL_DET_USUARIO,
    load_detalhamento_file,
)

FILTROS = (
    "Filtros aplicados:\n"
    "Mês é janeiro, fevereiro\n"
    "unidades.regiao é 1ª Região\n"
    "unidades.nome é PSF EM SAO PAULO\n"
    "USUARIO é FULANO DE TAL"
)

LINHAS = [
    ("00410043865202503", "MARIA SOUZA", "FULANO DE TAL", 3),
    ("00410096579202532", "JOAO LIMA", "FULANO DE TAL", 1),
    # Mesmo NUP em duas linhas, com responsáveis diferentes: deve virar uma só.
    ("00424054601202437", "MARIA SOUZA", "FULANO DE TAL", 2),
    ("00424054601202437", "JOAO LIMA", "FULANO DE TAL", 4),
]


def _planilha(filtros: str = FILTROS, linhas=LINHAS, cabecalho=None) -> io.BytesIO:
    cabecalho = cabecalho or [
        COL_DET_NUP, COL_DET_RESPONSAVEL, COL_DET_USUARIO, COL_DET_ATIVIDADES
    ]
    wb = Workbook()
    ws = wb.active
    ws["A1"] = filtros
    ws.append([])            # linha 2 (em branco)
    ws.append([])            # placeholder; o cabeçalho vai na linha 3
    ws.append(cabecalho)
    for linha in linhas:
        ws.append(list(linha))
    # remove a linha placeholder para deixar: A1 filtros, linha 2 vazia, linha 3 cabeçalho
    ws.delete_rows(3)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def test_agrega_por_nup_somando_atividades():
    data = load_detalhamento_file(_planilha(), "detalhamento.xlsx")

    assert data.total_nups == 3
    assert data.total_atividades == 10
    assert len(data.df) == 3

    linha = data.df[data.df[COL_DET_NUP] == "00424054601202437"].iloc[0]
    assert linha[COL_DET_ATIVIDADES] == 6
    assert "MARIA SOUZA" in linha[COL_DET_RESPONSAVEL]
    assert "JOAO LIMA" in linha[COL_DET_RESPONSAVEL]


def test_extrai_filtros_aplicados():
    data = load_detalhamento_file(_planilha(), "detalhamento.xlsx")

    assert data.usuario == "FULANO DE TAL"
    assert data.unidade == "PSF EM SAO PAULO"
    assert data.regiao == "1ª Região"
    assert data.meses == "janeiro, fevereiro"
    assert data.filtros_raw is not None and "Filtros aplicados" in data.filtros_raw


def test_conta_responsaveis_distintos():
    data = load_detalhamento_file(_planilha(), "detalhamento.xlsx")
    assert data.total_responsaveis == 2


def test_cabecalho_tolera_acento_e_caixa():
    cabecalho = ["nup", "RESPONSAVEL", "usuario que realizou a atividade", "atividades"]
    data = load_detalhamento_file(_planilha(cabecalho=cabecalho), "detalhamento.xlsx")

    assert data.total_nups == 3
    assert COL_DET_NUP in data.df.columns
    assert COL_DET_ATIVIDADES in data.df.columns


def test_atividades_nao_numericas_viram_zero():
    linhas = [("00410043865202503", "MARIA SOUZA", "FULANO DE TAL", "n/d")]
    data = load_detalhamento_file(_planilha(linhas=linhas), "detalhamento.xlsx")

    assert data.total_atividades == 0
    assert data.df.iloc[0][COL_DET_ATIVIDADES] == 0


def test_sem_cabecalho_nup_levanta_erro_em_portugues():
    cabecalho = ["Id", "Outro", "Coisa", "Qualquer"]
    with pytest.raises(ValueError, match="Detalhamento Individual"):
        load_detalhamento_file(_planilha(cabecalho=cabecalho), "errado.xlsx")


def test_coluna_faltando_levanta_erro_listando_a_coluna():
    cabecalho = [COL_DET_NUP, COL_DET_RESPONSAVEL, COL_DET_USUARIO, "Outra"]
    with pytest.raises(ValueError, match="Atividades"):
        load_detalhamento_file(_planilha(cabecalho=cabecalho), "incompleto.xlsx")


def test_filtros_ausentes_nao_quebram_a_leitura():
    data = load_detalhamento_file(_planilha(filtros=""), "sem_filtros.xlsx")

    assert data.total_nups == 3
    assert data.usuario is None
    assert data.unidade is None
