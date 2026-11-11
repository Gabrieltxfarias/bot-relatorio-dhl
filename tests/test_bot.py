import os
from urllib.parse import urlparse
from datetime import date

import pandas as pd
import pytest
from openpyxl import load_workbook

from bot.config import carregar_config
from bot.periodo import Estado, Periodo, calcular_periodo
from bot.planilha import DestinoExcelLocal, alimentar_planilha, filtrar_linhas, ler_exportacao


def test_periodo_normal_e_dia_anterior():
    hoje = date(2026, 11, 11)
    assert calcular_periodo(None, hoje) == Periodo(date(2026, 11, 10), date(2026, 11, 10))
    assert calcular_periodo(date(2026, 11, 9), hoje) == Periodo(date(2026, 11, 10), date(2026, 11, 10))
    assert calcular_periodo(date(2026, 11, 10), hoje) is None          # já em dia
    assert calcular_periodo(None, date(2026, 1, 1)).de == date(2025, 12, 31)


def test_periodo_recupera_fim_de_semana_e_respeita_limite():
    # última execução ok foi sexta (dados de quinta); PC só ligou na segunda
    assert calcular_periodo(date(2026, 11, 5), date(2026, 11, 9)) == Periodo(date(2026, 11, 6), date(2026, 11, 8))
    p = calcular_periodo(date(2026, 9, 1), date(2026, 11, 11), max_dias=31)
    assert p.dias == 31 and p.ate == date(2026, 11, 10)


def test_estado_nunca_retrocede(tmp_path):
    e = Estado(tmp_path / "estado.json")
    assert e.ultima_data_ok() is None
    e.registrar_sucesso(date(2026, 11, 10))
    e.registrar_sucesso(date(2026, 11, 1))
    assert e.ultima_data_ok() == date(2026, 11, 10)


def test_main_recupera_dias_e_depois_fica_em_dia(tmp_path, monkeypatch):
    import main as programa
    from tests.mock_tms import gerar_exportacao

    exp = tmp_path / "exp.xlsx"
    exp.write_bytes(gerar_exportacao("10/11/2026"))
    monkeypatch.setattr(programa, "RAIZ", tmp_path)
    Estado(tmp_path / "estado.json").registrar_sucesso(date.today() - __import__("datetime").timedelta(days=4))

    def rodar(*args):
        monkeypatch.setattr("sys.argv", ["main.py", *args])
        return programa.main()

    assert rodar("--arquivo", str(exp)) == 0                     # busca os 3 dias que faltavam
    assert Estado(tmp_path / "estado.json").ultima_data_ok() == date.today() - __import__("datetime").timedelta(days=1)
    planilha = tmp_path / "saida" / "InclusaoServicosAcessorios.xlsx"
    assert load_workbook(planilha).worksheets[0].max_row == 4
    assert rodar("--arquivo", str(exp)) == 0                     # já em dia: nada a fazer
    assert load_workbook(planilha).worksheets[0].max_row == 4



CHAVE = ["Documento Logística Original", "Inclusão"]


def _exportacao(tmp_path, dia="10/11/2026"):
    from tests.mock_tms import gerar_exportacao

    arq = tmp_path / "export.xlsx"
    arq.write_bytes(gerar_exportacao(dia))
    return ler_exportacao(arq)


def test_le_exportacao_mantem_colunas_e_tipos(tmp_path):
    from tests.mock_tms import COLUNAS

    df = _exportacao(tmp_path)
    assert list(df.columns) == COLUNAS                 # colunas vazias (Viagem, Status Geração...) mantidas
    assert len(df) == 4


def test_planilha_igual_ao_modelo_e_sem_duplicar(tmp_path):
    from datetime import datetime

    from tests.mock_tms import COLUNAS, GELO

    df = filtrar_linhas(_exportacao(tmp_path), "Serviço", [GELO])
    destino = DestinoExcelLocal(tmp_path / "online.xlsx")
    assert alimentar_planilha(df, destino, CHAVE) == (3, 0)
    assert alimentar_planilha(df, destino, CHAVE) == (0, 0)       # rodou de novo: nada muda

    ws = load_workbook(tmp_path / "online.xlsx").worksheets[0]
    assert [c.value for c in ws[1]] == COLUNAS                     # mesma disposição do TMS, sem colunas extras
    assert ws.max_row == 4
    assert ws["B2"].value == 1111 and ws["E2"].value == 940       # números continuam números
    assert ws["L2"].value == datetime(2026, 11, 10, 11, 5, 24)    # data/hora real
    assert ws["L2"].number_format == "dd/mm/yyyy hh:mm:ss"

    # status mudou num lançamento já gravado -> atualiza a linha
    df.loc[0, "Status Solicitação"] = "Aprovada"
    assert alimentar_planilha(df, destino, CHAVE) == (0, 1)
    ws = load_workbook(tmp_path / "online.xlsx").worksheets[0]
    assert ws["F2"].value == "Aprovada" and ws.max_row == 4

    # dia seguinte: linhas novas vão para o fim
    df2 = filtrar_linhas(_exportacao(tmp_path, "11/11/2026"), "Serviço", [GELO])
    assert alimentar_planilha(df2, destino, CHAVE) == (3, 0)
    assert load_workbook(tmp_path / "online.xlsx").worksheets[0].max_row == 7


def test_chave_reconhece_data_em_texto_ou_datetime():
    from datetime import datetime

    from bot.planilha import chave_texto

    assert chave_texto(datetime(2026, 11, 11, 11, 5, 24)) == chave_texto("11/11/2026 11:05:24") \
        == chave_texto("2026-11-11 11:05:24")
    assert chave_texto(5518) == chave_texto(5518.0) == chave_texto("5518")


def test_planilha_existente_com_formatacao_preservada(tmp_path):
    from openpyxl import Workbook
    from openpyxl.styles import Font

    from tests.mock_tms import COLUNAS, GELO

    arq = tmp_path / "online.xlsx"
    wb = Workbook()
    wb.active.title = "Minha Aba"
    wb.active.append(COLUNAS)
    for c in wb.active[1]:
        c.font = Font(bold=True)
    wb.save(arq)
    df = filtrar_linhas(_exportacao(tmp_path), "Serviço", [GELO])
    assert alimentar_planilha(df, DestinoExcelLocal(arq), CHAVE) == (3, 0)
    ws = load_workbook(arq)["Minha Aba"]
    assert ws["A1"].font.bold and ws.max_row == 4


GELO = "Troca de Gelo por Volume - Cliente/Transportador"


def test_filtra_servico():
    df = pd.DataFrame({"Servico": [
        "Troca de Gelo por Volume - Cliente/Transportador",
        "Complemento de Carga - Cliente/Transportador",
        "TROCA DE GELO POR VOLUME – CLIENTE/TRANSPORTADOR ",   # caixa alta, traço longo, espaço
        "Ajudante - Cliente/Transportador",
        None,
    ]})
    assert len(filtrar_linhas(df, "Serviço", [GELO])) == 2


def test_filtro_sem_coluna_da_erro():
    with pytest.raises(ValueError):
        filtrar_linhas(pd.DataFrame({"Viagem": ["1"]}), "Serviço", [GELO])


def test_le_excel_que_e_html(tmp_path):
    arq = tmp_path / "r.xls"
    arq.write_text("<table><tr><th>A</th></tr><tr><td>x</td></tr></table>", encoding="utf-8")
    assert ler_exportacao(arq)["A"].tolist() == ["x"]


@pytest.mark.e2e
def test_fluxo_completo_no_tms_falso(tmp_path, monkeypatch):
    from bot import navegador
    from tests import mock_tms

    srv, base = mock_tms.iniciar()
    try:
        cfg = carregar_config()
        for chave, url in cfg["urls"].items():
            u = urlparse(url)
            cfg["urls"][chave] = f"{base}{u.path}" + (f"?{u.query}" if u.query else "")
        cfg["urls"]["inicial"] = f"{base}/novidades.aspx"
        cfg["navegador"].update(headless=True, canal="", executavel=os.getenv("PW_EXECUTAVEL", ""),
                                pasta_perfil=str(tmp_path / "perfil"), timeout_login_s=30, timeout_relatorio_s=30)
        monkeypatch.setattr(navegador, "RAIZ", tmp_path)
        monkeypatch.setattr(navegador.credenciais, "obter", lambda: ("usuario.teste", "senha.teste"))

        with navegador.BotTMS(cfg) as bot:
            arquivo = bot.executar(date(2026, 11, 10))

        assert mock_tms.Estado.ultimo_filtro == {
            "tipo": "4",                      # Receita de Documento
            "status": ["10", "11"],           # Aguardando Aprovação + Aprovada (Em Elaboração removida)
            "de": "10/11/2026", "ate": "10/11/2026",
            "usuario": ["Gabriel Teixeira De Farias"],
        }
        df = filtrar_linhas(ler_exportacao(arquivo), "Serviço", [GELO])
        assert len(df) == 3 and set(df["Serviço"]) == {GELO}
        destino = tmp_path / "consolidado.xlsx"
        assert alimentar_planilha(df, DestinoExcelLocal(destino), CHAVE) == (3, 0)

        # 2ª execução: sessão salva no perfil, sem precisar de login
        monkeypatch.setattr(navegador.credenciais, "obter", lambda: ("", ""))
        with navegador.BotTMS(cfg) as bot:
            bot.abrir_e_logar()
            assert bot.page.url.endswith("/novidades.aspx")
            assert "/login" not in bot.page.url
    finally:
        srv.shutdown()


def test_onedrive_tabela_excel_estendida_e_backup(tmp_path):
    from openpyxl import Workbook
    from openpyxl.worksheet.table import Table

    from tests.mock_tms import COLUNAS, GELO

    arq = tmp_path / "OneDrive" / "online.xlsx"
    arq.parent.mkdir()
    wb = Workbook()
    wb.active.append(COLUNAS)
    wb.active.append([None] * 12 + ["x"])
    wb.active.add_table(Table(displayName="Lancamentos", ref="A1:M2"))
    wb.save(arq)

    destino = DestinoExcelLocal(arq, pasta_backup=tmp_path / "backups")
    df = filtrar_linhas(_exportacao(tmp_path), "Serviço", [GELO])
    assert alimentar_planilha(df, destino, CHAVE) == (3, 0)
    ws = load_workbook(arq).worksheets[0]
    assert ws.tables["Lancamentos"].ref == "A1:M5"
    assert len(list((tmp_path / "backups").glob("online_*.xlsx"))) == 1


def test_recusa_planilha_com_grafico(tmp_path):
    from openpyxl import Workbook
    from openpyxl.chart import BarChart, Reference

    from bot.planilha import ErroPlanilha
    from tests.mock_tms import COLUNAS, GELO

    arq = tmp_path / "online.xlsx"
    wb = Workbook()
    wb.active.append(COLUNAS)
    graf = BarChart()
    graf.add_data(Reference(wb.active, min_col=5, min_row=1, max_row=2))
    wb.active.add_chart(graf, "O2")
    wb.save(arq)
    antes = arq.read_bytes()
    df = filtrar_linhas(_exportacao(tmp_path), "Serviço", [GELO])
    with pytest.raises(ErroPlanilha, match="gráficos"):
        alimentar_planilha(df, DestinoExcelLocal(arq), CHAVE)
    assert arq.read_bytes() == antes          # nada foi alterado


def test_espera_planilha_aberta_no_excel(tmp_path, monkeypatch):
    import builtins

    from bot import planilha
    from tests.mock_tms import GELO

    arq = tmp_path / "online.xlsx"
    df = filtrar_linhas(_exportacao(tmp_path), "Serviço", [GELO])
    alimentar_planilha(df.iloc[:1], DestinoExcelLocal(arq), CHAVE)

    tentativas = {"n": 0}
    abrir_original = builtins.open

    def abrir_travado(caminho, modo="r", *a, **k):
        if str(caminho) == str(arq) and modo == "r+b" and tentativas["n"] < 2:
            tentativas["n"] += 1
            raise PermissionError("arquivo em uso")
        return abrir_original(caminho, modo, *a, **k)

    monkeypatch.setattr(builtins, "open", abrir_travado)
    monkeypatch.setattr(planilha._time, "sleep", lambda s: None)
    assert alimentar_planilha(df, DestinoExcelLocal(arq), CHAVE) == (2, 0)
    assert tentativas["n"] == 2
