"""Bot de extração do relatório de Serviços Acessórios (TMS DHL).

Uso:
    python main.py                               # tudo que falta até ontem (recupera dias perdidos)
    python main.py --data 10/11/2026             # um dia específico
    python main.py --de 01/11/2026 --ate 10/11/2026   # período específico
    python main.py --arquivo rel.xlsx            # só alimenta a planilha com um arquivo já baixado

Sem datas, o bot lembra a última data processada com sucesso (estado.json) e
busca de uma vez todos os dias que faltam até ontem. Se já estiver em dia,
encerra sem abrir o navegador.
"""
from __future__ import annotations

import argparse
import logging
import sys
import time
from datetime import date, datetime, timedelta
from pathlib import Path

from bot.caminho_planilha import resolver_caminho_planilha
import pandas as pd

from bot.conferencia import conferir_relatorio
from bot.config import RAIZ, carregar_config
from bot.periodo import Estado, Periodo, calcular_periodo
from bot.planilha import DestinoExcelLocal, alimentar_planilha, filtrar_linhas, ler_exportacao


def configurar_log() -> None:
    (RAIZ / "logs").mkdir(exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler(RAIZ / "logs" / f"bot_{datetime.now():%Y%m}.log", encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )


def _data(txt: str) -> date:
    return datetime.strptime(txt, "%d/%m/%Y").date()


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--data", help="Um dia específico dd/mm/aaaa")
    ap.add_argument("--de", help="Início de um período específico dd/mm/aaaa")
    ap.add_argument("--ate", help="Fim de um período específico dd/mm/aaaa")
    ap.add_argument("--arquivo", help="Pula o navegador e usa um arquivo já exportado")
    ap.add_argument("--config", help="Caminho alternativo do config.yaml")
    ap.add_argument("--headless", action="store_true", help="Roda sem abrir janela")
    args = ap.parse_args()

    configurar_log()
    log = logging.getLogger("main")
    cfg = carregar_config(args.config)
    if args.headless:
        cfg["navegador"]["headless"] = True

    estado = Estado(RAIZ / "estado.json")
    ultima_ok = estado.ultima_data_ok()
    manual = bool(args.data or args.de or args.ate)
    if args.data:
        periodo = Periodo(_data(args.data), _data(args.data))
    elif args.de or args.ate:
        if not (args.de and args.ate):
            ap.error("Use --de e --ate juntos.")
        periodo = Periodo(_data(args.de), _data(args.ate))
        if periodo.de > periodo.ate:
            ap.error("--de não pode ser depois de --ate.")
    else:
        periodo = calcular_periodo(ultima_ok, max_dias=cfg["filtros"].get("max_dias_recuperacao", 31))
        if periodo is None:
            log.info("Nada a fazer: dados já processados até %s.", ultima_ok.strftime("%d/%m/%Y"))
            return 0

    log.info("=== Início | período %s ===", periodo)
    if ultima_ok and not manual and periodo.dias > 1:
        log.info("Recuperando dias perdidos desde a última execução (%s).", ultima_ok.strftime("%d/%m/%Y"))

    tentativas = 1 if args.arquivo else max(1, cfg.get("execucao", {}).get("tentativas", 3))
    intervalo_min = cfg.get("execucao", {}).get("intervalo_tentativas_min", 20)
    for tentativa in range(1, tentativas + 1):
        try:
            lidas, novas, atualizadas = processar(cfg, periodo, args.arquivo)
            # Atualiza a "última data ok" só com dias COMPLETOS (até ontem) e sem deixar
            # buraco para trás. Ex.: testar com a data de hoje não marca hoje como feito.
            concluido_ate = min(periodo.ate, date.today() - timedelta(days=1))
            sem_buraco = ultima_ok is None or periodo.de <= ultima_ok + timedelta(days=1)
            if concluido_ate >= periodo.de and sem_buraco:
                estado.registrar_sucesso(concluido_ate)
            log.info("=== Fim | %s | %d linha(s) lida(s), %d nova(s), %d atualizada(s) ===",
                     periodo, lidas, novas, atualizadas)
            return 0
        except Exception:
            log.exception("Falha na tentativa %d de %d", tentativa, tentativas)
            if tentativa < tentativas:
                log.info("Nova tentativa em %d min (ex.: VPN ainda conectando).", intervalo_min)
                time.sleep(intervalo_min * 60)
    log.error("Todas as tentativas falharam. O período %s será buscado de novo na próxima execução.", periodo)
    return 1


def processar(cfg: dict, periodo: Periodo, arquivo: str | None) -> tuple[int, int, int]:
    """Extrai (ou usa o arquivo dado), filtra e grava na planilha."""
    if not arquivo:
        from bot.navegador import BotTMS

        with BotTMS(cfg) as bot:
            arquivo = bot.executar(periodo.de, periodo.ate)

    log = logging.getLogger("main")
    if arquivo is None:  # pesquisa sem resultados: nada para exportar
        log.info("Relatório do TMS: nenhum (a pesquisa não retornou lançamentos).")
        df = pd.DataFrame()
    else:
        df = ler_exportacao(arquivo)
        log.info("Relatório do TMS: %s", Path(arquivo).resolve())
    total = len(df)
    df = conferir_relatorio(df, cfg["filtros"], periodo.de, periodo.ate)
    fora_filtros = total - len(df)
    p = cfg["planilha"]
    antes_servico = len(df)
    for filtro in p.get("filtrar_linhas") or []:
        df = filtrar_linhas(df, filtro["coluna"], filtro["valores"])
    outro_servico = antes_servico - len(df)
    destino = DestinoExcelLocal(
        resolver_caminho_planilha(p["arquivo"], RAIZ),
        p.get("aba") or "",
        pasta_backup=RAIZ / "backups",
        backups_mantidos=p.get("backups_mantidos", 30),
        espera_arquivo_aberto_s=p.get("espera_arquivo_aberto_min", 10) * 60,
    )
    novas, atualizadas = alimentar_planilha(df, destino, p.get("colunas_chave"))
    ja_existiam = len(df) - novas - atualizadas

    log.info("-------- RESUMO %s --------", periodo)
    log.info("Linhas no relatório do TMS ............ %d", total)
    log.info("Descartadas (fora dos filtros) ........ %d", fora_filtros)
    log.info("Descartadas (outro serviço) ........... %d", outro_servico)
    log.info("Já estavam na planilha (ignoradas) .... %d", ja_existiam)
    log.info("Atualizadas na planilha ............... %d", atualizadas)
    log.info("NOVAS gravadas na planilha ............ %d", novas)
    if total == 0:
        log.info("Conclusão: o TMS não retornou lançamentos para o período. Nada a gravar.")
    elif novas == 0 and atualizadas == 0 and ja_existiam > 0:
        log.info("Conclusão: tudo o que veio do TMS JÁ ESTAVA na planilha. Nada foi alterado.")
    elif novas == 0 and atualizadas == 0:
        log.info("Conclusão: nenhuma linha do relatório era do serviço/filtros esperados. Nada a gravar.")
    else:
        log.info("Conclusão: planilha atualizada com %d nova(s) e %d atualizada(s).", novas, atualizadas)
    return total, novas, atualizadas

if __name__ == "__main__":
    sys.exit(main())
