"""Define o período a extrair e lembra a última data processada com sucesso.

Assim, se o PC ficar desligado (fim de semana, férias), a próxima execução
busca de uma vez todos os dias que faltaram, até ontem.
"""
from __future__ import annotations

import json
import logging
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path

log = logging.getLogger(__name__)


@dataclass(frozen=True)
class Periodo:
    de: date
    ate: date

    @property
    def dias(self) -> int:
        return (self.ate - self.de).days + 1

    def __str__(self) -> str:
        if self.de == self.ate:
            return self.de.strftime("%d/%m/%Y")
        return f"{self.de:%d/%m/%Y} a {self.ate:%d/%m/%Y} ({self.dias} dias)"


class Estado:
    """Arquivo JSON com a última data processada com sucesso."""

    def __init__(self, arquivo: str | Path):
        self.arquivo = Path(arquivo)

    def ultima_data_ok(self) -> date | None:
        try:
            dados = json.loads(self.arquivo.read_text(encoding="utf-8"))
            return datetime.strptime(dados["ultima_data_ok"], "%Y-%m-%d").date()
        except FileNotFoundError:
            return None
        except Exception as e:
            log.warning("Arquivo de estado ilegível (%s): %s — ignorando.", self.arquivo, e)
            return None

    def registrar_sucesso(self, ate: date) -> None:
        atual = self.ultima_data_ok()
        if atual and atual >= ate:  # nunca retrocede (ex.: reprocessar um dia antigo)
            return
        self.arquivo.write_text(
            json.dumps({"ultima_data_ok": ate.isoformat(), "atualizado_em": datetime.now().isoformat(timespec="seconds")},
                       ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        log.info("Última data processada com sucesso: %s", ate.strftime("%d/%m/%Y"))


def calcular_periodo(ultima_ok: date | None, hoje: date | None = None, max_dias: int = 31) -> Periodo | None:
    """Período pendente até ontem. None se já estiver tudo processado."""
    ontem = (hoje or date.today()) - timedelta(days=1)
    de = ultima_ok + timedelta(days=1) if ultima_ok else ontem
    if de > ontem:
        return None
    if (ontem - de).days + 1 > max_dias:
        novo_de = ontem - timedelta(days=max_dias - 1)
        log.warning("Faltam dados desde %s, mas o limite é %d dias: buscando a partir de %s. "
                    "Para o restante, rode: python main.py --de %s --ate %s",
                    de.strftime("%d/%m/%Y"), max_dias, novo_de.strftime("%d/%m/%Y"),
                    de.strftime("%d/%m/%Y"), (novo_de - timedelta(days=1)).strftime("%d/%m/%Y"))
        de = novo_de
    return Periodo(de, ontem)
