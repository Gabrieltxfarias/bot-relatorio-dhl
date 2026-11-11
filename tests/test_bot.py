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
    from io import StringIO

    from ruamel.yaml import YAML

    from bot.config import ler_yaml

    cfg = ler_yaml(open(programa.RAIZ / "config.yaml", encoding="utf-8"))
    cfg["planilha"]["arquivo"] = "saida/InclusaoServicosAcessorios.xlsx"
    saida = StringIO()
    YAML(typ="safe", pure=True).dump(cfg, saida)
    (tmp_path / "config.yaml").write_text(saida.getvalue(), encoding="utf-8")
    monkeypatch.setattr(programa, "RAIZ", tmp_path)
    Estado(tmp_path / "estado.json").registrar_sucesso(date.today() - __import__("datetime").timedelta(days=4))

    def rodar(*args):
        monkeypatch.setattr("sys.argv", ["main.py", "--config", str(tmp_path / "config.yaml"), *args])
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
        cfg["rotulos"]["usuario_solicitante"] = "Solicitante"   # rótulo parcial também deve funcionar
        cfg["navegador"].update(headless=True, canal="", executavel=os.getenv("PW_EXECUTAVEL", ""),
                                pasta_perfil=str(tmp_path / "perfil"), timeout_login_s=30, timeout_relatorio_s=30)
        monkeypatch.setattr(navegador, "RAIZ", tmp_path)
        monkeypatch.setattr(navegador.credenciais, "obter", lambda: ("usuario.teste", "senha.teste"))

        with navegador.BotTMS(cfg) as bot:
            arquivo = bot.executar(date(2026, 11, 10))
        assert (tmp_path / "conferencia" / "2026-11-10_1_filtros_preenchidos.png").exists()
        assert (tmp_path / "conferencia" / "2026-11-10_2_resultado_pesquisa.png").exists()

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

        # dia sem lançamentos: pesquisa vazia não é erro
        with navegador.BotTMS(cfg) as bot:
            assert bot.executar(date(2026, 1, 1)) is None
        assert (tmp_path / "conferencia" / "2026-01-01_2_resultado_pesquisa.png").exists()

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


LINK = ("https://dpdhl.sharepoint.com/sites/DHL-AIR-GRU/Shared%20Documents/Processos%20-%20A%C3%A9reo%20Operacional/"
        "Emiss%C3%A3o%20(LastMile,%20RPS,%20Minuta)/REPORT%20RPS%20TROCA%20DE%20GELO.xlsx?web=1")
RESTO = ["Processos - Aéreo Operacional", "Emissão (LastMile, RPS, Minuta)", "REPORT RPS TROCA DE GELO.xlsx"]


def test_link_sharepoint_partes():
    from bot.caminho_planilha import partes_do_link

    assert partes_do_link(LINK) == ("DHL-AIR-GRU", RESTO)
    assert partes_do_link(LINK.replace("/sites/", "/:x:/r/sites/")) == ("DHL-AIR-GRU", RESTO)


def _cria(caminho):
    caminho.parent.mkdir(parents=True, exist_ok=True)
    caminho.write_bytes(b"x")
    return caminho


def test_link_acha_biblioteca_sincronizada(tmp_path):
    from bot.caminho_planilha import resolver_caminho_planilha

    perfil = tmp_path / "Users" / "gabriel"
    arq = _cria(perfil.joinpath("Deutsche Post DHL", "DHL-AIR-GRU - Documents", *RESTO))
    (perfil / "OneDrive - Deutsche Post DHL").mkdir(parents=True)
    amb = {"USERPROFILE": str(perfil), "OneDriveCommercial": str(perfil / "OneDrive - Deutsche Post DHL")}
    assert resolver_caminho_planilha(LINK, tmp_path, amb) == arq


def test_link_acha_atalho_no_onedrive(tmp_path):
    from bot.caminho_planilha import resolver_caminho_planilha

    perfil = tmp_path / "Users" / "gabriel"
    od = perfil / "OneDrive - Deutsche Post DHL"
    arq = _cria(od.joinpath(*RESTO[1:]))          # atalho para a pasta "Emissão (...)"
    amb = {"USERPROFILE": str(perfil), "OneDriveCommercial": str(od)}
    assert resolver_caminho_planilha(LINK, tmp_path, amb) == arq


def test_link_nao_sincronizado_explica(tmp_path):
    from bot.caminho_planilha import resolver_caminho_planilha
    from bot.planilha import ErroPlanilha

    perfil = tmp_path / "Users" / "gabriel"
    perfil.mkdir(parents=True)
    with pytest.raises(ErroPlanilha, match="Sincronizar"):
        resolver_caminho_planilha(LINK, tmp_path, {"USERPROFILE": str(perfil)})


def test_main_rodar_dia_de_hoje_nao_marca_como_concluido(tmp_path, monkeypatch):
    import datetime as dt

    from ruamel.yaml import YAML
    from io import StringIO

    import main as programa
    from bot.config import ler_yaml
    from tests.mock_tms import gerar_exportacao

    cfg = ler_yaml(open(programa.RAIZ / "config.yaml", encoding="utf-8"))
    cfg["planilha"]["arquivo"] = "saida/teste.xlsx"
    buf = StringIO()
    YAML(typ="safe", pure=True).dump(cfg, buf)
    (tmp_path / "config.yaml").write_text(buf.getvalue(), encoding="utf-8")
    exp = tmp_path / "exp.xlsx"
    hoje = dt.date.today()
    exp.write_bytes(gerar_exportacao(hoje.strftime("%d/%m/%Y")))
    monkeypatch.setattr(programa, "RAIZ", tmp_path)

    def rodar(*args):
        monkeypatch.setattr("sys.argv", ["main.py", "--config", str(tmp_path / "config.yaml"), *args])
        return programa.main()

    assert rodar("--data", hoje.strftime("%d/%m/%Y"), "--arquivo", str(exp)) == 0
    assert Estado(tmp_path / "estado.json").ultima_data_ok() is None      # hoje ainda não acabou
    ontem = hoje - dt.timedelta(days=1)
    assert rodar("--de", ontem.strftime("%d/%m/%Y"), "--ate", hoje.strftime("%d/%m/%Y"), "--arquivo", str(exp)) == 0
    assert Estado(tmp_path / "estado.json").ultima_data_ok() == ontem     # só até ontem


def test_conferencia_descarta_linhas_fora_do_filtro(tmp_path):
    from bot.conferencia import conferir_relatorio

    df = _exportacao(tmp_path, "10/11/2026")
    df.loc[1, "Status Solicitação"] = "Reprovada"
    df.loc[2, "Tipo Solicitação"] = "Receita de Viagem"
    df.loc[3, "Inclusão"] = pd.Timestamp("2026-11-09 08:00:00")
    filtros = {"tipo_solicitacao": "Receita de Documento", "status": ["Aguardando Aprovação", "Aprovada"]}
    ok = conferir_relatorio(df, filtros, date(2026, 11, 10), date(2026, 11, 10))
    assert len(ok) == 1
    tudo_certo = conferir_relatorio(_exportacao(tmp_path, "10/11/2026"), filtros, date(2026, 11, 10), date(2026, 11, 10))
    assert len(tudo_certo) == 4
