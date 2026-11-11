"""Converte o valor de `planilha.arquivo` (link do SharePoint ou caminho) no
arquivo local que o OneDrive sincroniza.

Bibliotecas do SharePoint sincronizadas aparecem no PC em lugares como:
  - %USERPROFILE%\\<Empresa>\\<Site> - <Biblioteca>\\...      (botão "Sincronizar")
  - %OneDriveCommercial%\\<Pasta do atalho>\\...             ("Adicionar atalho a Meus arquivos")
  - %OneDriveCommercial%\\...                                (arquivos do próprio OneDrive)
"""
from __future__ import annotations

import logging
import os
from pathlib import Path
from urllib.parse import parse_qs, unquote, urlparse

from .planilha import ErroPlanilha

log = logging.getLogger(__name__)

COMO_SINCRONIZAR = (
    "Abra o link da planilha no navegador, vá na pasta da biblioteca (Documentos) e clique em "
    "'Sincronizar' ou em 'Adicionar atalho a Meus arquivos'. Aguarde o OneDrive mostrar a pasta "
    "no Explorer e rode de novo. Alternativa: coloque em planilha.arquivo o caminho local do "
    "arquivo (Explorer > Shift + botão direito no arquivo > 'Copiar como caminho')."
)


def resolver_caminho_planilha(valor: str, raiz: Path, ambiente: dict | None = None) -> Path:
    valor = (valor or "").strip().strip('"')
    if valor.lower().startswith(("http://", "https://")):
        return localizar_arquivo_sincronizado(valor, ambiente)
    caminho = Path(os.path.expandvars(os.path.expanduser(valor)))
    if not caminho.is_absolute():  # dentro do projeto (ex.: saida/): cria a pasta
        caminho = raiz / caminho
        caminho.parent.mkdir(parents=True, exist_ok=True)
    return caminho


def partes_do_link(url: str) -> tuple[str, list[str]]:
    """Retorna (site, caminho_dentro_da_biblioteca) a partir do link."""
    u = urlparse(url)
    caminho = u.path
    q = parse_qs(u.query)
    if "id" in q:  # links de visualização de pasta: .../AllItems.aspx?id=/sites/X/Shared Documents/...
        caminho = q["id"][0]
    partes = [unquote(p) for p in caminho.split("/") if p]
    # links "/:x:/r/sites/..." têm um prefixo antes de /sites/: começa em sites/teams/personal
    inicio = next((i for i, p in enumerate(partes) if p in ("sites", "teams", "personal")), None)
    partes = partes[inicio:] if inicio is not None else []
    if len(partes) >= 3 and partes[0] in ("sites", "teams", "personal"):
        site = partes[1]
        resto = partes[3:]  # pula a biblioteca ("Shared Documents", "Documents"...)
        if partes[0] == "personal":
            site = ""
        if resto and resto[-1].lower().endswith((".xlsx", ".xlsm")):
            return site, resto
    raise ErroPlanilha(
        "Não consegui entender o link da planilha. Use o link da barra de endereço do navegador "
        "(com o nome do arquivo .xlsx no final) ou, no Excel, Arquivo > Informações > Copiar caminho. "
        "Links curtos de compartilhamento (…/:x:/g/…) não trazem o caminho do arquivo."
    )


def _raizes(site: str, ambiente: dict) -> list[Path]:
    """Pastas onde bibliotecas sincronizadas costumam ficar."""
    raizes: list[Path] = []
    for var in ("OneDriveCommercial", "OneDrive"):
        if ambiente.get(var):
            raizes.append(Path(ambiente[var]))
    perfil = Path(ambiente.get("USERPROFILE") or ambiente.get("HOME") or Path.home())
    site_l = site.lower()
    try:
        for nivel1 in perfil.iterdir():
            if not nivel1.is_dir() or nivel1.name.startswith("."):
                continue
            if site_l and nivel1.name.lower().startswith(site_l):
                raizes.append(nivel1)
            if nivel1.name.lower().startswith("onedrive"):
                raizes.append(nivel1)
            try:
                for nivel2 in nivel1.iterdir():  # %USERPROFILE%\<Empresa>\<Site> - <Biblioteca>
                    if nivel2.is_dir() and site_l and nivel2.name.lower().startswith(site_l):
                        raizes.append(nivel2)
            except (PermissionError, OSError):
                continue
    except (PermissionError, OSError):
        pass
    # atalhos ficam um nível abaixo do OneDrive
    extras = []
    for r in raizes:
        if r.name.lower().startswith("onedrive"):
            try:
                extras += [p for p in r.iterdir() if p.is_dir()]
            except (PermissionError, OSError):
                pass
    vistos, unicas = set(), []
    for r in raizes + extras:
        chave = str(r).lower()
        if chave not in vistos and r.exists():
            vistos.add(chave)
            unicas.append(r)
    return unicas


def localizar_arquivo_sincronizado(url: str, ambiente: dict | None = None) -> Path:
    ambiente = dict(os.environ) if ambiente is None else ambiente
    site, resto = partes_do_link(url)
    raizes = _raizes(site, ambiente)
    # tenta o caminho completo e os finais dele (atalho pode apontar para uma subpasta)
    for raiz in raizes:
        for i in range(len(resto)):
            candidato = raiz.joinpath(*resto[i:])
            if candidato.is_file():
                log.info("Planilha do SharePoint encontrada no PC: %s", candidato)
                return candidato
    raise ErroPlanilha(
        f"A planilha '{resto[-1]}' não foi encontrada sincronizada neste PC "
        f"(procurei em: {', '.join(str(r) for r in raizes) or 'nenhuma pasta do OneDrive encontrada'}). "
        + COMO_SINCRONIZAR
    )
