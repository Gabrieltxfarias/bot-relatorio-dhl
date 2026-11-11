"""Bot de extração do relatório de Serviços Acessórios (TMS DHL).

Uso:
    python main.py                      # dia anterior
    python main.py --data 10/11/2026    # data específica
    python main.py --arquivo rel.xlsx   # só alimenta a planilha com um arquivo já baixado
"""
from __future__ import annotations

import argparse
import logging
import os
import sys
from datetime import datetime

from bot.config import RAIZ, carregar_config, data_referencia
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


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--data", help="Data de referência dd/mm/aaaa (padrão: dia anterior)")
    ap.add_argument("--arquivo", help="Pula o navegador e usa um arquivo já exportado")
    ap.add_argument("--config", help="Caminho alternativo do config.yaml")
    ap.add_argument("--headless", action="store_true", help="Roda sem abrir janela")
    args = ap.parse_args()

    configurar_log()
    log = logging.getLogger("main")
    cfg = carregar_config(args.config)
    if args.headless:
        cfg["navegador"]["headless"] = True

    data_ref = (
        datetime.strptime(args.data, "%d/%m/%Y").date()
        if args.data
        else data_referencia(cfg["filtros"].get("dias_atras", 1))
    )
    log.info("=== Início | data de referência %s ===", data_ref.strftime("%d/%m/%Y"))

    try:
        if args.arquivo:
            arquivo = args.arquivo
        else:
            from bot.navegador import BotTMS

            with BotTMS(cfg) as bot:
                arquivo = bot.executar(data_ref)

        df = ler_exportacao(arquivo)
        p = cfg["planilha"]
        for filtro in p.get("filtrar_linhas") or []:
            df = filtrar_linhas(df, filtro["coluna"], filtro["valores"])
        caminho = os.path.expandvars(os.path.expanduser(p["arquivo"]))
        if not os.path.isabs(caminho):  # caminho dentro do projeto (ex.: saida/): cria a pasta
            (RAIZ / caminho).parent.mkdir(parents=True, exist_ok=True)
        destino = DestinoExcelLocal(
            RAIZ / caminho,
            p.get("aba") or "",
            pasta_backup=RAIZ / "backups",
            backups_mantidos=p.get("backups_mantidos", 30),
            espera_arquivo_aberto_s=p.get("espera_arquivo_aberto_min", 10) * 60,
        )
        novas, atualizadas = alimentar_planilha(df, destino, p.get("colunas_chave"))
        log.info("=== Fim | %d linha(s) lida(s), %d nova(s), %d atualizada(s) ===", len(df), novas, atualizadas)
        return 0
    except Exception:
        log.exception("Falha na execução")
        return 1


if __name__ == "__main__":
    sys.exit(main())
