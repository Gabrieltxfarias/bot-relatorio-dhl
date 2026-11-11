"""Confere se este PC tem tudo o que o bot precisa.

Uso (na pasta do projeto, com o ambiente ativado):
    .venv\\Scripts\\activate
    python verificar_ambiente.py
"""
from __future__ import annotations

import importlib
import os
import platform
import subprocess
import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parent
NOME_TAREFA = "Bot Relatorio DHL - Servicos Acessorios"
resultados: list[tuple[str, str]] = []

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


def ok(item: str, detalhe: str = "") -> None:
    resultados.append(("OK", item))
    print(f"  [OK]    {item}" + (f" — {detalhe}" if detalhe else ""))


def falha(item: str, como_resolver: str) -> None:
    resultados.append(("FALHA", item))
    print(f"  [FALHA] {item}\n          -> {como_resolver}")


def aviso(item: str, detalhe: str) -> None:
    resultados.append(("AVISO", item))
    print(f"  [AVISO] {item}\n          -> {detalhe}")


def secao(titulo: str) -> None:
    print(f"\n== {titulo}")


# ------------------------------------------------------------------ Python
def checar_python() -> None:
    secao("Python")
    v = sys.version_info
    if v >= (3, 10):
        ok("Versão do Python", f"{v.major}.{v.minor}.{v.micro}")
    else:
        falha(f"Python {v.major}.{v.minor} é antigo", "Instale o Python 3.12 em python.org (marque 'Add python.exe to PATH').")
    if "onedrive" in str(RAIZ).lower():
        aviso("O projeto está dentro do OneDrive",
              "Mova a pasta do projeto para fora (ex.: C:\\Users\\<usuario>\\BotRPS). Senão o OneDrive vai "
              "sincronizar o .venv (milhares de arquivos) e a sessão de login do navegador (perfil_navegador) para a nuvem.")
    if sys.prefix != getattr(sys, "base_prefix", sys.prefix):
        ok("Ambiente virtual (.venv) ativo")
    else:
        aviso("Ambiente virtual não está ativo",
              "Rode '.venv\\Scripts\\activate' antes (ou crie com 'python -m venv .venv').")


def checar_pacotes() -> bool:
    secao("Bibliotecas")
    pacotes = {"playwright": "playwright", "pandas": "pandas", "openpyxl": "openpyxl", "lxml": "lxml",
               "ruamel.yaml": "ruamel.yaml", "dotenv": "python-dotenv", "keyring": "keyring"}
    faltando = []
    for modulo, nome in pacotes.items():
        try:
            m = importlib.import_module(modulo)
            ok(nome, getattr(m, "__version__", ""))
        except Exception:
            faltando.append(nome)
    if faltando:
        falha(f"Faltando: {', '.join(faltando)}", "Rode: pip install -r requirements.txt")
    return not faltando


# ------------------------------------------------------------------ config
def carregar_config() -> dict | None:
    secao("Configuração")
    try:
        sys.path.insert(0, str(RAIZ))
        from bot.config import carregar_config as carregar

        cfg = carregar()
        ok("config.yaml lido")
        return cfg
    except Exception as e:
        falha("config.yaml com erro", f"{e}")
        return None


def checar_navegador(cfg: dict) -> None:
    secao("Navegador")
    nav = cfg.get("navegador", {})
    canal = nav.get("canal") or ""
    try:
        from playwright.sync_api import sync_playwright

        with sync_playwright() as pw:
            opcoes = {"headless": True}
            if nav.get("executavel"):
                opcoes["executable_path"] = nav["executavel"]
            elif canal:
                opcoes["channel"] = canal
            navegador = pw.chromium.launch(**opcoes)
            versao = navegador.version
            navegador.close()
        ok(f"Navegador '{nav.get('executavel') or canal or 'chromium'}' abre", versao)
    except Exception as e:
        primeira = str(e).strip().splitlines()[0][:150]
        falha(f"Não consegui abrir o navegador '{canal or 'chromium'}'",
              ("Confira se o Microsoft Edge está instalado e abre normalmente; "
               if canal == "msedge" else f"Rode: playwright install {canal or 'chromium'}; ")
              + f"(erro: {primeira})")
    perfil = RAIZ / nav.get("pasta_perfil", "perfil_navegador")
    if perfil.exists():
        ok("Sessão do navegador salva", str(perfil))
    else:
        aviso("Ainda não há sessão salva",
              "Normal antes da 1ª execução: rode 'python main.py' com a janela visível e faça o login.")


def checar_credenciais() -> None:
    secao("Login do TMS")
    try:
        from bot import credenciais

        usuario, senha = credenciais.obter()
        if usuario and senha:
            ok("Credenciais salvas", f"usuário: {usuario}")
        else:
            aviso("Nenhuma credencial salva",
                  "Rode 'python salvar_credenciais.py' (ou faça o login manual na 1ª execução).")
    except Exception as e:
        falha("Gerenciador de credenciais indisponível", f"{e}")


def checar_rede(cfg: dict) -> None:
    secao("Rede / VPN")
    import urllib.error
    import urllib.request
    from urllib.parse import urlparse

    url = cfg.get("urls", {}).get("inicial", "")
    host = urlparse(url).hostname
    try:
        urllib.request.urlopen(urllib.request.Request(f"https://{host}/", method="HEAD"), timeout=15)
        ok(f"TMS acessível ({host})")
    except urllib.error.HTTPError as e:  # respondeu (mesmo que com erro/redirect de login) = rede ok
        ok(f"TMS acessível ({host})", f"HTTP {e.code}")
    except Exception as e:
        falha(f"Não consegui acessar {host}", f"Conecte a VPN/rede da DHL e tente de novo. ({e.__class__.__name__}: {e})")


def checar_planilha(cfg: dict) -> None:
    secao("Planilha (OneDrive)")
    p = cfg.get("planilha", {})
    bruto = str(p.get("arquivo", ""))
    dentro_do_projeto = False
    try:
        from bot.caminho_planilha import resolver_caminho_planilha

        if bruto.lower().startswith("http"):
            caminho = resolver_caminho_planilha(bruto, RAIZ)
            ok("Link do SharePoint encontrado sincronizado no PC", str(caminho))
        else:
            caminho = Path(os.path.expandvars(os.path.expanduser(bruto)))
            dentro_do_projeto = not caminho.is_absolute()
            if dentro_do_projeto:
                caminho = RAIZ / caminho
                aviso("planilha.arquivo aponta para dentro do projeto",
                      f"Hoje grava em {caminho}. Para a planilha online, coloque o link do SharePoint.")
    except Exception as e:
        falha("Planilha do SharePoint não encontrada no PC", str(e))
        return
    if not caminho.parent.exists() and dentro_do_projeto:
        return  # a pasta dentro do projeto é criada pelo bot
    if not caminho.parent.exists():
        falha("Pasta da planilha não existe", f"{caminho.parent} — confira o caminho no config.yaml.")
        return
    if not caminho.exists():
        aviso("Planilha ainda não existe", f"Será criada na 1ª execução: {caminho}")
        return
    ok("Planilha encontrada", str(caminho))
    try:
        from bot.planilha import DestinoExcelLocal, ErroPlanilha

        d = DestinoExcelLocal(caminho, p.get("aba") or "")
        d._checar_itens_nao_suportados()
        cab, linhas = d.ler()
        ok("Planilha legível", f"{len(linhas)} linha(s), colunas: {', '.join(str(c) for c in cab[:4])}...")
        try:
            with open(caminho, "r+b"):
                ok("Planilha livre para gravação")
        except PermissionError:
            aviso("Planilha aberta/travada agora", "Feche-a no Excel do PC antes do horário do bot.")
    except ErroPlanilha as e:
        falha("Planilha com problema", str(e))
    except Exception as e:
        falha("Não consegui ler a planilha", f"{e}")


def checar_agendamento() -> None:
    secao("Agendamento")
    if platform.system() != "Windows":
        aviso("Não é Windows", "A verificação do Agendador de Tarefas só funciona no Windows.")
        return
    r = subprocess.run(["schtasks", "/Query", "/TN", NOME_TAREFA, "/FO", "LIST", "/V"],
                       capture_output=True, text=True, errors="replace")
    if r.returncode != 0:
        aviso("Tarefa agendada não encontrada",
              "Rode: powershell -ExecutionPolicy Bypass -File .\\agendar_tarefa.ps1")
        return
    info = {}
    for linha in r.stdout.splitlines():
        if ":" in linha:
            k, _, v = linha.partition(":")
            info.setdefault(k.strip().lower(), v.strip())
    proxima = next((v for k, v in info.items() if "próxima" in k or "next run" in k or "proxima" in k), "?")
    ultimo = next((v for k, v in info.items() if "último resultado" in k or "last result" in k or "ultimo resultado" in k), "?")
    ok("Tarefa agendada existe", f"próxima execução: {proxima} | último resultado: {ultimo}")
    if (RAIZ / ".venv" / "Scripts" / "python.exe").exists():
        ok(".venv encontrado para a tarefa")
    else:
        falha(".venv não encontrado", "A tarefa usa .venv\\Scripts\\python.exe. Crie com 'python -m venv .venv'.")


def checar_estado() -> None:
    secao("Histórico")
    try:
        from bot.periodo import Estado

        ultima = Estado(RAIZ / "estado.json").ultima_data_ok()
        if ultima:
            ok("Última data processada com sucesso", ultima.strftime("%d/%m/%Y"))
        else:
            aviso("O bot ainda não rodou com sucesso", "Normal antes da 1ª execução.")
    except Exception as e:
        aviso("Não consegui ler estado.json", str(e))


def main() -> int:
    print(f"Verificando o ambiente do bot em {RAIZ}  ({platform.system()} {platform.release()})")
    checar_python()
    if not checar_pacotes():
        print("\nInstale as bibliotecas e rode a verificação de novo.")
        return 1
    cfg = carregar_config()
    if cfg:
        checar_navegador(cfg)
        checar_credenciais()
        checar_rede(cfg)
        checar_planilha(cfg)
    checar_agendamento()
    checar_estado()

    n_falha = sum(1 for s, _ in resultados if s == "FALHA")
    n_aviso = sum(1 for s, _ in resultados if s == "AVISO")
    print("\n" + "=" * 60)
    if n_falha:
        print(f"Resultado: {n_falha} falha(s) e {n_aviso} aviso(s). Corrija as falhas acima.")
    elif n_aviso:
        print(f"Resultado: tudo essencial OK, com {n_aviso} aviso(s) para conferir.")
    else:
        print("Resultado: tudo OK! O bot está pronto para rodar.")
    return 1 if n_falha else 0


if __name__ == "__main__":
    sys.exit(main())
