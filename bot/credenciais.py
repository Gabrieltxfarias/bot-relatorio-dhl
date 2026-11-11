"""Credenciais do TMS guardadas fora do código.

Ordem de busca:
1. Gerenciador de Credenciais do Windows (via `keyring`) — recomendado;
   a senha fica criptografada pela sua conta do Windows.
2. Arquivo .env (TMS_USUARIO / TMS_SENHA) — alternativa simples.
"""
from __future__ import annotations

import logging
import os

log = logging.getLogger(__name__)

SERVICO = "bot_relatorio_dhl_tms"
CHAVE_USUARIO = "__usuario__"


def salvar(usuario: str, senha: str) -> None:
    import keyring

    keyring.set_password(SERVICO, CHAVE_USUARIO, usuario)
    keyring.set_password(SERVICO, usuario, senha)


def apagar() -> None:
    import keyring

    usuario = keyring.get_password(SERVICO, CHAVE_USUARIO)
    if usuario:
        keyring.delete_password(SERVICO, usuario)
        keyring.delete_password(SERVICO, CHAVE_USUARIO)


def obter() -> tuple[str, str]:
    """Retorna (usuario, senha); strings vazias se nada estiver configurado."""
    try:
        import keyring

        usuario = keyring.get_password(SERVICO, CHAVE_USUARIO)
        if usuario:
            senha = keyring.get_password(SERVICO, usuario) or ""
            if senha:
                return usuario, senha
    except Exception as e:  # keyring indisponível nesta máquina
        log.debug("keyring indisponível: %s", e)
    return os.getenv("TMS_USUARIO", ""), os.getenv("TMS_SENHA", "")
