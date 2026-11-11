"""Leitura do arquivo exportado pelo TMS, filtro de linhas e alimentação da planilha.

A gravação é separada em duas partes:
- `planejar_gravacao`: decide o que é linha nova e o que é atualização
  (não depende de onde a planilha está);
- um "destino" que aplica o plano: `DestinoExcelLocal` (arquivo .xlsx) hoje,
  e um destino online (planilha na nuvem) com a mesma interface.
"""
from __future__ import annotations

import logging
import os
import re
import shutil
import tempfile
import time as _time
import unicodedata
import zipfile
from dataclasses import dataclass, field
from datetime import date, datetime
from io import StringIO
from pathlib import Path

import pandas as pd
from openpyxl import Workbook, load_workbook

log = logging.getLogger(__name__)

FORMATO_DATA_HORA = "dd/mm/yyyy hh:mm:ss"


# ============================================================ leitura
def ler_exportacao(arquivo: str | Path) -> pd.DataFrame:
    """Lê o arquivo exportado mantendo os tipos (datas, números, textos).

    Também aceita "Excel" que na verdade é HTML/CSV, comum em sistemas ASP.NET.
    Todas as colunas do relatório são mantidas, mesmo as vazias, para a
    planilha ficar com a mesma disposição do TMS.
    """
    arquivo = Path(arquivo)
    erros = []
    for leitor in (_ler_excel, _ler_html, _ler_csv):
        try:
            df = leitor(arquivo)
            if df is not None and not df.empty:
                df.columns = [str(c).strip() for c in df.columns]
                vazias_sem_nome = [c for c in df.columns if c.startswith("Unnamed") and df[c].isna().all()]
                df = df.drop(columns=vazias_sem_nome).dropna(how="all").reset_index(drop=True)
                log.info("Arquivo lido com %s: %d linha(s), colunas %s", leitor.__name__, len(df), list(df.columns))
                return df
        except Exception as e:  # tenta o próximo formato
            erros.append(f"{leitor.__name__}: {e}")
    if erros:
        log.warning("Nenhuma linha lida de %s (%s)", arquivo, "; ".join(erros))
    return pd.DataFrame()


def _ler_excel(arquivo: Path) -> pd.DataFrame:
    return pd.read_excel(arquivo)


def _ler_html(arquivo: Path) -> pd.DataFrame:
    tabelas = pd.read_html(StringIO(_texto(arquivo)))
    return max(tabelas, key=len)


def _ler_csv(arquivo: Path) -> pd.DataFrame:
    return pd.read_csv(StringIO(_texto(arquivo)), sep=None, engine="python")


def _texto(arquivo: Path) -> str:
    bruto = arquivo.read_bytes()
    for cod in ("utf-8-sig", "cp1252"):
        try:
            return bruto.decode(cod)
        except UnicodeDecodeError:
            continue
    return bruto.decode("latin-1", errors="replace")


# ============================================================ normalização
def _sem_acento(s) -> str:
    s = unicodedata.normalize("NFD", str(s))
    return " ".join("".join(c for c in s if not unicodedata.combining(c)).lower().split())


def _comparavel(s) -> str:
    """Texto para comparação: sem acento, minúsculo, traços e espaços uniformes."""
    s = re.sub(r"[‐-―−]", "-", _sem_acento("" if s is None else s))
    return re.sub(r"\s*-\s*", " - ", s).strip()


def valor_celula(v):
    """Converte valores do pandas/numpy para tipos que planilhas entendem."""
    if v is None:
        return None
    try:
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(v, pd.Timestamp):
        return v.to_pydatetime()
    if hasattr(v, "item"):  # numpy int/float/bool
        v = v.item()
    if isinstance(v, float) and v.is_integer():
        return int(v)
    if isinstance(v, str):
        v = v.strip()
        return v or None
    return v


def chave_texto(v) -> str:
    """Representação estável de um valor para comparar chaves entre execuções.

    Ex.: datetime(2026,11,11,11,5,24), "11/11/2026 11:05:24" e "2026-11-11 11:05:24"
    viram o mesmo texto; 5518, 5518.0 e "5518" também.
    """
    v = valor_celula(v)
    if v is None:
        return ""
    if isinstance(v, datetime):
        return v.strftime("%d/%m/%Y %H:%M:%S")
    if isinstance(v, date):
        return v.strftime("%d/%m/%Y 00:00:00")
    if isinstance(v, str):
        for fmt in ("%d/%m/%Y %H:%M:%S", "%Y-%m-%d %H:%M:%S", "%d/%m/%Y %H:%M", "%Y-%m-%dT%H:%M:%S"):
            try:
                return datetime.strptime(v, fmt).strftime("%d/%m/%Y %H:%M:%S")
            except ValueError:
                pass
        try:
            f = float(v.replace(",", ".")) if re.fullmatch(r"-?\d+([.,]\d+)?", v) else None
            if f is not None:
                return str(int(f)) if f.is_integer() else str(f)
        except ValueError:
            pass
        return _comparavel(v)
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() else str(v)
    return str(v)


# ============================================================ filtro
def filtrar_linhas(df: pd.DataFrame, coluna: str, valores: list[str]) -> pd.DataFrame:
    """Mantém só as linhas cuja `coluna` seja um dos `valores` (ignora acento,
    maiúsculas, espaços extras e tipo de traço). Erro se a coluna não existir,
    para nunca gravar dados sem o filtro aplicado."""
    if df.empty or not valores:
        return df
    col = next((c for c in df.columns if _sem_acento(c) == _sem_acento(coluna)), None)
    if col is None:
        raise ValueError(f"Coluna '{coluna}' não encontrada no relatório exportado. Colunas: {list(df.columns)}")
    alvos = {_comparavel(v) for v in valores}
    mascara = df[col].map(lambda x: _comparavel(valor_celula(x))).isin(alvos)
    filtrado = df[mascara].reset_index(drop=True)
    descartadas = df.loc[~mascara, col].fillna("(vazio)").value_counts()
    log.info("Filtro '%s': %d de %d linha(s) mantida(s).", col, len(filtrado), len(df))
    for servico, qtd in descartadas.items():
        log.info("  descartado: %s (%d)", servico, qtd)
    return filtrado


# ============================================================ plano de gravação
@dataclass
class Plano:
    cabecalho: list[str]                       # cabeçalho final da planilha
    cabecalho_mudou: bool
    novas: list[list] = field(default_factory=list)                 # linhas a acrescentar
    atualizacoes: dict[int, list] = field(default_factory=dict)     # nº da linha -> linha completa


def _resolver_chave(desejadas: list[str], colunas: list[str]) -> list[str]:
    """Encontra as colunas-chave configuradas, tolerando acentos/maiúsculas e
    nomes parciais ("Inclusão" casa com "Data Inclusão")."""
    achadas = []
    for d in desejadas:
        alvo = _sem_acento(d)
        exatas = [c for c in colunas if _sem_acento(c) == alvo]
        parciais = [c for c in colunas if alvo in _sem_acento(c)]
        escolha = (exatas or parciais or [None])[0]
        if escolha is None:
            log.warning("Coluna-chave '%s' não existe no relatório (colunas: %s). "
                        "Comparando a linha inteira para evitar duplicatas.", d, colunas)
            return []
        achadas.append(escolha)
    if achadas:
        log.info("Identificando lançamentos por: %s", achadas)
    return achadas


def planejar_gravacao(
    df: pd.DataFrame,
    cabecalho_atual: list[str],
    linhas_atuais: list[list],
    colunas_chave: list[str] | None,
    primeira_linha_dados: int = 2,
) -> Plano:
    """Compara o relatório com o que já está na planilha.

    - lançamento novo (chave inédita) -> acrescenta;
    - lançamento existente com algum valor diferente (ex.: status) -> atualiza;
    - igual -> ignora.
    """
    cabecalho = [c for c in cabecalho_atual if c not in (None, "")] or list(df.columns)
    mudou = not cabecalho_atual or any(c not in cabecalho for c in df.columns)
    for col in df.columns:  # coluna nova no relatório vai para o fim
        if col not in cabecalho:
            cabecalho.append(col)
    idx = {c: i for i, c in enumerate(cabecalho)}

    chave_cols = _resolver_chave(colunas_chave or [], list(df.columns)) or list(df.columns)

    def chave_linha(linha: list) -> tuple:
        return tuple(chave_texto(linha[idx[c]]) if idx[c] < len(linha) else "" for c in chave_cols)

    existentes: dict[tuple, int] = {}
    linhas_por_numero: dict[int, list] = {}
    for n, linha in enumerate(linhas_atuais, start=primeira_linha_dados):
        linha = list(linha) + [None] * (len(cabecalho) - len(linha))
        if all(valor_celula(v) is None for v in linha):
            continue
        existentes.setdefault(chave_linha(linha), n)
        linhas_por_numero[n] = linha

    plano = Plano(cabecalho=cabecalho, cabecalho_mudou=mudou)
    for _, reg in df.iterrows():
        nova = [valor_celula(reg[c]) if c in df.columns else None for c in cabecalho]
        chave = chave_linha(nova)
        if chave in existentes:
            n = existentes[chave]
            atual = plano.atualizacoes.get(n, linhas_por_numero[n])
            if [chave_texto(v) for v in atual] != [chave_texto(v) for v in nova]:
                plano.atualizacoes[n] = nova
            continue
        plano.novas.append(nova)
        existentes[chave] = -len(plano.novas)  # evita duplicar dentro do mesmo relatório
    return plano


# ============================================================ destinos
class ErroPlanilha(RuntimeError):
    pass


class DestinoExcelLocal:
    """Planilha .xlsx no disco — inclusive a cópia sincronizada do OneDrive/SharePoint.

    Ao salvar no arquivo da pasta do OneDrive, o próprio OneDrive envia a
    alteração para a planilha online.
    """

    # Itens que o openpyxl não preserva: gravar apagaria esses elementos.
    ITENS_NAO_SUPORTADOS = {"xl/charts/": "gráficos", "xl/drawings/": "imagens/formas",
                            "xl/pivotTables/": "tabelas dinâmicas"}

    def __init__(self, arquivo: str | Path, aba: str = "", pasta_backup: str | Path | None = None,
                 backups_mantidos: int = 30, espera_arquivo_aberto_s: int = 600):
        self.arquivo = Path(os.path.expandvars(os.path.expanduser(str(arquivo))))
        self.aba = aba
        self.pasta_backup = Path(pasta_backup) if pasta_backup else None
        self.backups_mantidos = backups_mantidos
        self.espera_arquivo_aberto_s = espera_arquivo_aberto_s

    def __str__(self):
        return f"{self.arquivo} [{self.aba or 'primeira aba'}]"

    def _abrir(self):
        if self.arquivo.exists():
            wb = load_workbook(self.arquivo)
        else:
            if not self.arquivo.parent.exists():
                raise ErroPlanilha(f"Pasta da planilha não existe: {self.arquivo.parent}. "
                                   "Confira o caminho em planilha.arquivo no config.yaml.")
            wb = Workbook()
            if self.aba:
                wb.active.title = self.aba
        if self.aba:
            if self.aba not in wb.sheetnames and self.arquivo.exists():
                raise ErroPlanilha(f"Aba '{self.aba}' não existe em {self.arquivo.name}. Abas: {wb.sheetnames}")
            ws = wb[self.aba] if self.aba in wb.sheetnames else wb.active
        else:
            ws = wb.worksheets[0]
        return wb, ws

    def ler(self) -> tuple[list[str], list[list]]:
        if not self.arquivo.exists():
            log.info("Planilha ainda não existe; será criada: %s", self.arquivo)
            return [], []
        _, ws = self._abrir()
        linhas = [list(r) for r in ws.iter_rows(values_only=True)]
        if not linhas or all(v is None for v in linhas[0]):
            return [], []
        return [c for c in linhas[0]], linhas[1:]

    def aplicar(self, plano: Plano) -> None:
        self._checar_itens_nao_suportados()
        self._esperar_arquivo_livre()
        self._backup()
        wb, ws = self._abrir()
        if plano.cabecalho_mudou:
            for i, col in enumerate(plano.cabecalho, 1):
                ws.cell(1, i, col)
        for n, linha in plano.atualizacoes.items():
            for i, v in enumerate(linha, 1):
                self._escrever(ws.cell(n, i), v)
        proxima = self._ultima_linha_com_dados(ws) + 1
        for linha in plano.novas:
            for i, v in enumerate(linha, 1):
                self._escrever(ws.cell(proxima, i), v)
            proxima += 1
        self._estender_tabelas(ws)
        self._salvar(wb)

    # ---------------------------------------------------------- proteções
    def _checar_itens_nao_suportados(self) -> None:
        if not self.arquivo.exists():
            return
        with zipfile.ZipFile(self.arquivo) as z:
            achados = {nome for prefixo, nome in self.ITENS_NAO_SUPORTADOS.items()
                       if any(n.startswith(prefixo) for n in z.namelist())}
        if achados:
            raise ErroPlanilha(
                f"A planilha tem {', '.join(sorted(achados))}, que seriam apagados ao gravar. "
                "Mantenha só os dados nela e faça gráficos/dinâmicas em outro arquivo que leia esta planilha.")

    def _esperar_arquivo_livre(self) -> None:
        """Se a planilha estiver aberta no Excel do PC (arquivo travado), espera."""
        if not self.arquivo.exists():
            return
        fim = _time.time() + self.espera_arquivo_aberto_s
        avisou = False
        while True:
            try:
                with open(self.arquivo, "r+b"):
                    return
            except PermissionError:
                if _time.time() >= fim:
                    raise ErroPlanilha(f"A planilha continua aberta/travada: {self.arquivo}. "
                                       "Feche-a no Excel do computador e rode o bot de novo.")
                if not avisou:
                    log.warning("Planilha aberta no Excel deste PC. Aguardando ser fechada "
                                "(até %d min)...", self.espera_arquivo_aberto_s // 60)
                    avisou = True
                _time.sleep(15)

    def _backup(self) -> None:
        if not self.pasta_backup or not self.arquivo.exists():
            return
        self.pasta_backup.mkdir(parents=True, exist_ok=True)
        destino = self.pasta_backup / f"{self.arquivo.stem}_{datetime.now():%Y%m%d_%H%M%S}{self.arquivo.suffix}"
        shutil.copy2(self.arquivo, destino)
        antigos = sorted(self.pasta_backup.glob(f"{self.arquivo.stem}_*{self.arquivo.suffix}"))
        for velho in antigos[:-self.backups_mantidos]:
            velho.unlink(missing_ok=True)
        log.info("Backup da planilha: %s", destino)

    def _salvar(self, wb) -> None:
        """Salva num temporário e copia por cima: o arquivo sincronizado nunca
        fica pela metade se algo falhar no meio da gravação."""
        with tempfile.TemporaryDirectory() as tmp:
            temp = Path(tmp) / self.arquivo.name
            wb.save(temp)
            shutil.copyfile(temp, self.arquivo)

    def _estender_tabelas(self, ws) -> None:
        """Se os dados estiverem numa Tabela do Excel, inclui as linhas novas nela."""
        ultima = self._ultima_linha_com_dados(ws)
        for tabela in ws.tables.values():
            inicio, fim = tabela.ref.split(":")
            col_fim = re.match(r"[A-Z]+", fim).group()
            if int(re.search(r"\d+", fim).group()) < ultima:
                tabela.ref = f"{inicio}:{col_fim}{ultima}"
                if tabela.autoFilter is not None:
                    tabela.autoFilter.ref = tabela.ref

    @staticmethod
    def _escrever(cel, v):
        cel.value = v
        if isinstance(v, datetime):
            cel.number_format = FORMATO_DATA_HORA
        elif isinstance(v, int) and not isinstance(v, bool):
            cel.number_format = "0"

    @staticmethod
    def _ultima_linha_com_dados(ws) -> int:
        for n in range(ws.max_row, 0, -1):
            if any(c.value is not None for c in ws[n]):
                return n
        return 0


def alimentar_planilha(df: pd.DataFrame, destino, colunas_chave: list[str] | None = None) -> tuple[int, int]:
    """Grava o relatório no destino. Retorna (linhas_novas, linhas_atualizadas)."""
    if df.empty:
        log.info("Relatório sem linhas — nada a gravar.")
        return 0, 0
    cabecalho, linhas = destino.ler()
    plano = planejar_gravacao(df, cabecalho, linhas, colunas_chave)
    if plano.novas or plano.atualizacoes or plano.cabecalho_mudou:
        destino.aplicar(plano)
    log.info("Planilha %s: %d linha(s) nova(s), %d atualizada(s).", destino, len(plano.novas), len(plano.atualizacoes))
    return len(plano.novas), len(plano.atualizacoes)
