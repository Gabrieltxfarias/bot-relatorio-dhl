"""Leitura da configuração (config.yaml)."""
from __future__ import annotations

from datetime import date, timedelta
from pathlib import Path

import yaml
from dotenv import load_dotenv

RAIZ = Path(__file__).resolve().parent.parent


def carregar_config(caminho: str | Path | None = None) -> dict:
    caminho = Path(caminho) if caminho else RAIZ / "config.yaml"
    with open(caminho, encoding="utf-8") as f:
        cfg = yaml.safe_load(f)
    load_dotenv(RAIZ / ".env")  # só usado se não houver credencial no Windows
    return cfg


def data_referencia(dias_atras: int = 1, hoje: date | None = None) -> date:
    """Dia usado no filtro (padrão: dia anterior)."""
    return (hoje or date.today()) - timedelta(days=dias_atras)
