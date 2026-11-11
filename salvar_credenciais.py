"""Grava (ou apaga) o login do TMS no Gerenciador de Credenciais do Windows.

Uso:
    python salvar_credenciais.py           # pede usuário e senha
    python salvar_credenciais.py --apagar  # remove o que foi salvo

A senha é digitada sem aparecer na tela e não é gravada em nenhum arquivo.
"""
import getpass
import sys

from bot import credenciais


def main() -> None:
    if "--apagar" in sys.argv:
        credenciais.apagar()
        print("Credenciais removidas.")
        return
    usuario = input("Usuário do TMS (ex.: login/e-mail DHL): ").strip()
    senha = getpass.getpass("Senha (não aparece ao digitar): ")
    if not usuario or not senha:
        print("Usuário e senha são obrigatórios. Nada foi salvo.")
        sys.exit(1)
    credenciais.salvar(usuario, senha)
    u, s = credenciais.obter()
    print("Credenciais salvas com sucesso." if (u, s) == (usuario, senha) else "Falha ao confirmar o salvamento.")


if __name__ == "__main__":
    main()
