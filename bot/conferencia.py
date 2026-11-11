"""Confere se o relatório exportado respeita os filtros pedidos ao TMS.

Linhas fora do esperado (outro tipo, outro status, data fora do período) não
vão para a planilha e aparecem como aviso no log.
"""
from __future__ import annotations

import logging
from datetime import date, datetime

import pandas as pd

from .planilha import _comparavel, _sem_acento, valor_celula

log = logging.getLogger(__name__)


def _coluna(df: pd.DataFrame, nome: str) -> str | None:
    return next((c for c in df.columns if _sem_acento(c) == _sem_acento(nome)), None)


def _data(v) -> date | None:
    v = valor_celula(v)
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, str):
        for fmt in ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%d/%m/%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
            try:
                return datetime.strptime(v.strip(), fmt).date()
            except ValueError:
                pass
    return None


def conferir_relatorio(df: pd.DataFrame, filtros: dict, de: date, ate: date,
                       colunas: dict | None = None) -> pd.DataFrame:
    """Retorna só as linhas que batem com os filtros e registra o resultado no log."""
    if df.empty:
        log.info("Conferência: relatório vazio para %s a %s.", f"{de:%d/%m/%Y}", f"{ate:%d/%m/%Y}")
        return df
    colunas = colunas or {}
    regras = []  # (descrição, coluna, função que diz se o valor está ok)

    col_tipo = _coluna(df, colunas.get("tipo", "Tipo Solicitação"))
    if col_tipo and filtros.get("tipo_solicitacao"):
        alvo = _comparavel(filtros["tipo_solicitacao"])
        regras.append(("tipo", col_tipo, lambda v: _comparavel(valor_celula(v)) == alvo))

    col_status = _coluna(df, colunas.get("status", "Status Solicitação"))
    if col_status and filtros.get("status"):
        aceitos = {_comparavel(s) for s in filtros["status"]}
        regras.append(("status", col_status, lambda v: _comparavel(valor_celula(v)) in aceitos))

    col_data = _coluna(df, colunas.get("data", "Inclusão"))
    if col_data:
        regras.append(("data de inclusão", col_data, lambda v: (d := _data(v)) is not None and de <= d <= ate))

    if not regras:
        log.warning("Conferência: não encontrei as colunas de tipo/status/inclusão no relatório.")
        return df

    ok = pd.Series(True, index=df.index)
    for descricao, col, teste in regras:
        bate = df[col].map(teste)
        for i in df.index[~bate]:
            log.warning("Conferência: linha %d com %s fora do filtro (%s = %r) — não será gravada.",
                        i + 2, descricao, col, valor_celula(df.at[i, col]))
        ok &= bate

    conferidas = ", ".join(f"{d} ({c})" for d, c, _ in regras)
    if ok.all():
        log.info("Conferência OK: as %d linha(s) batem com os filtros — %s.", len(df), conferidas)
    else:
        log.warning("Conferência: %d de %d linha(s) fora dos filtros foram descartadas.", (~ok).sum(), len(df))
    return df[ok].reset_index(drop=True)
