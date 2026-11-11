"""Leitura da configuração (config.yaml)."""
from __future__ import annotations

from pathlib import Path

from dotenv import load_dotenv
from ruamel.yaml import YAML

RAIZ = Path(__file__).resolve().parent.parent


def carregar_config(caminho: str | Path | None = None) -> dict:
    caminho = Path(caminho) if caminho else RAIZ / "config.yaml"
    with open(caminho, encoding="utf-8") as f:
        cfg = ler_yaml(f)
    load_dotenv(RAIZ / ".env")  # só usado se não houver credencial no Windows
    return cfg



def ler_yaml(fluxo) -> dict:
    """Lê YAML com ruamel.yaml em modo puro (não precisa de compilador C++)."""
    return YAML(typ="safe", pure=True).load(fluxo) or {}
