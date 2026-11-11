"""Servidor local que imita as telas do TMS (Benner), só para testar o bot.

Reproduz: login, troca de papel por URL, postbacks ASP.NET (__doPostBack),
campos select2 (estático e com busca "AJAX"), e a barra de processos
assíncronos (#header_task_bar) que avisa quando a exportação termina.
"""
from __future__ import annotations

import html
import threading
import time
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

STATIC = Path(__file__).parent / "static"
P = "ctl00$Main$K_INCLUSODESERVIOSACESSRIOS$FilterControl$GERAL$"
ID = "ctl00_Main_K_INCLUSODESERVIOSACESSRIOS_FilterControl_GERAL_"
URL_PAPEL = "/Roles/Select/K_ANALISTASERVICOSACESSORIOS"
URL_GRID = "/cli/tms/operacional/pages/a/ServicosAcessorios/InclusaoServicos/Grid.aspx"
USUARIOS = ["Gabriel Silva", "Gabriel Teixeira De Farias", "Maria Souza"]

LOGIN = """<form method='post' action='/login'>
<label>Usuário <input name='username'></label>
<label>Senha <input type='password' name='password'></label>
<button>Entrar</button></form>"""


def pagina(corpo: str) -> str:
    return f"""<html><head><meta charset='utf-8'>
<link rel='stylesheet' href='/static/select2.min.css'>
<script src='/static/jquery.min.js'></script><script src='/static/select2.min.js'></script>
</head><body>
<ul><li class="dropdown dropdown-extended dropdown-tasks" id="header_task_bar">
  <a href="#" class="dropdown-toggle" onclick="$('#tasks').toggle();return false">▤ <span class="badge"></span></a>
  <ul id="tasks" class="dropdown-menu extended tasks" style="display:none">
    <li class="external"><h3><span class="bold">Processos</span></h3>
      <a href="javascript:Benner.AsyncProcesses.clearReadyProcesses();">Limpar</a></li>
    <li><div class="slimScrollDiv"><ul class="dropdown-menu-list scroller"></ul></div></li>
  </ul></li></ul>
{corpo}
<div id="ampl-survey" class="ampl-survey" style="position:fixed;inset:0;z-index:9999;background:rgba(0,0,0,.4)">
  <button class="ampl-option">Concordo</button></div>
<script>
window.Benner = {{AsyncProcesses: {{clearReadyProcesses: function() {{ $.get('/processos/limpar'); }}}}}};
setInterval(function() {{ $.get('/processos', function(h) {{ $('#header_task_bar ul.dropdown-menu-list').html(h); }}); }}, 700);
</script></body></html>"""


class Estado:
    ultimo_filtro: dict = {}
    processos: list[dict] = []
    lock = threading.Lock()


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    # ---------------------------------------------------------------- helpers
    def _enviar(self, corpo: bytes, tipo="text/html; charset=utf-8", status=200, headers=None):
        self.send_response(status)
        self.send_header("Content-Type", tipo)
        for k, v in (headers or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(corpo)

    def _html(self, corpo: str):
        self._enviar(pagina(corpo).encode("utf-8"))

    def _redirect(self, para: str, cookie: str | None = None):
        self.send_response(302)
        if cookie:
            self.send_header("Set-Cookie", f"{cookie}; Path=/; Max-Age=86400")
        self.send_header("Location", para)
        self.end_headers()

    def _cookie(self, nome: str) -> str | None:
        for parte in (self.headers.get("Cookie") or "").split(";"):
            k, _, v = parte.strip().partition("=")
            if k == nome:
                return v
        return None

    # ------------------------------------------------------------------- GET
    def do_GET(self):
        u = urlparse(self.path)
        if u.path.startswith("/static/"):
            arq = STATIC / Path(u.path).name
            tipo = "text/css" if arq.suffix == ".css" else "application/javascript"
            return self._enviar(arq.read_bytes(), tipo)
        if u.path == "/login":
            return self._html(LOGIN)
        if self._cookie("sessao") != "ok":
            return self._redirect("/login")
        if u.path == "/novidades.aspx":
            papel = "Analista de Serviços Acessórios (DHL)" if self._cookie("papel") else "Analista Comercial (DHL)"
            return self._html(f"<header>{papel}</header><h1>Painel de Novidades</h1>")
        if u.path == URL_PAPEL:
            return self._redirect("/novidades.aspx", "papel=K_ANALISTASERVICOSACESSORIOS")
        if u.path == URL_GRID:
            return self._grid({})
        if u.path == "/processos":
            return self._enviar(self._html_processos().encode("utf-8"))
        if u.path == "/processos/limpar":
            with Estado.lock:
                Estado.processos = [p for p in Estado.processos if time.time() < p["fim"]]
            return self._enviar(b"ok")
        if u.path.startswith("/download/"):
            return self._download()
        self._enviar(b"nao encontrado", status=404)

    # ------------------------------------------------------------------ POST
    def do_POST(self):
        tam = int(self.headers.get("Content-Length", 0))
        form = parse_qs(self.rfile.read(tam).decode(), keep_blank_values=True)
        u = urlparse(self.path)
        if u.path == "/login":
            if form.get("username") == ["usuario.teste"] and form.get("password") == ["senha.teste"]:
                return self._redirect("/novidades.aspx", "sessao=ok")
            return self._html(LOGIN + "<p>Credenciais inválidas</p>")
        if u.path == URL_GRID and self._cookie("sessao") == "ok":
            if form.get("__EVENTTARGET", [""])[0].endswith("TIPOSOLICITACAO"):
                time.sleep(1.5)
            return self._grid(form)
        self._enviar(b"erro", status=400)

    # ------------------------------------------------------------------ telas
    def _grid(self, form: dict):
        if self._cookie("papel") != "K_ANALISTASERVICOSACESSORIOS":
            return self._html("<h1>Acesso negado: papel incorreto</h1>")
        alvo = form.get("__EVENTTARGET", [""])[0]
        arg = form.get("__EVENTARGUMENT", [""])[0]
        filtro = {
            "tipo": form.get(P + "TIPOSOLICITACAOselect", ["1"])[0],
            "status": form.get(P + "STATUS", []),
            "de": form.get(P + "DATAINCLUSAODE_DATE", ["10/11/2026"])[0],
            "ate": form.get(P + "DATAINCLUSAOATE_DATE", ["11/11/2026"])[0],
            "usuario": form.get(P + "USUARIOSOLICITANTE", []),
        }
        grid = ""
        if alvo.endswith("FilterButton") or arg == "ExportToExcel":
            Estado.ultimo_filtro = filtro
            grid = """<div id="ctl00_Main_K_INCLUSODESERVIOSACESSRIOS_SimpleGrid_ctl01_id">
              <div class="btn extra-options-button" onclick="$('#menu').toggle()">⋮</div></div>
              <ul id="menu" style="display:none"><li id="ExportToExcel" role="presentation">
              <a href="javascript:__doPostBack('ctl00$Main$K_INCLUSODESERVIOSACESSRIOS','ExportToExcel')">Exportar para Excel</a></li></ul>
              <table><tr><td>2 registros</td></tr></table>"""
        if arg == "ExportToExcel":
            with Estado.lock:
                n = len(Estado.processos) + 1
                Estado.processos.append({"id": n, "fim": time.time() + 3, "filtro": filtro})

        def opt(valor, texto, sel):
            return f"<option value='{valor}'{' selected' if sel else ''}>{texto}</option>"

        tipos = [("-1", "(nenhum)"), ("1", "Todos"), ("2", "Despesa de Viagem"), ("3", "Receita de Viagem"),
                 ("4", "Receita de Documento"), ("5", "Despesa e Receita de Viagem")]
        status = [("10", "Aguardando Aprovação"), ("11", "Aprovada"), ("12", "Cancelada"),
                  ("13", "Em Elaboração"), ("14", "Reprovada")]
        users_sel = "".join(opt(u, u, True) for u in filtro["usuario"])
        corpo = f"""
<header>Analista de Serviços Acessórios (DHL)</header>
<form method="post" id="aspnetForm">
<input type="hidden" name="__EVENTTARGET"><input type="hidden" name="__EVENTARGUMENT">
<div data-field="TIPOSOLICITACAO" data-label="Tipo de Solicita&amp;#231;&amp;#227;o">
  <select name="{P}TIPOSOLICITACAOselect" id="{ID}TIPOSOLICITACAOselect" class="form-control"
    onchange="__doPostBack('ctl00$Main$K_INCLUSODESERVIOSACESSRIOS$FilterControl$GERAL$TIPOSOLICITACAO','')">
  {''.join(opt(v, t, v == filtro['tipo']) for v, t in tipos)}</select></div>
<div data-field="STATUSAPROVACAOFIN" data-label="Status Aprova&amp;#231;&amp;#227;o Financeira">
  <select name="{P}STATUSFIN" class="s2" multiple><option>Pendente</option><option>Aprovada</option></select></div>
<div class="row"><label class="control-label">Status</label><div class="col">
  <select name="{P}STATUS" id="status" multiple>
  {''.join(opt(v, t, True) for v, t in status if v in filtro['status'] or (not form and v in ('10', '13')))}</select></div></div>
<div data-field="DATAINCLUSAODE" data-label="Data Inclus&amp;#227;o De" data-type="date"><div class="input-group">
  <input name="{P}DATAINCLUSAODE_DATE" id="{ID}DATAINCLUSAODE_DATE" value="{html.escape(filtro['de'])}" maxlength="10" class="form-control datepicker"></div></div>
<div data-field="DATAINCLUSAOATE" data-label="Data Inclus&amp;#227;o At&amp;#233;" data-type="date"><div class="input-group">
  <input name="{P}DATAINCLUSAOATE_DATE" id="{ID}DATAINCLUSAOATE_DATE" value="{html.escape(filtro['ate'])}" maxlength="10" class="form-control datepicker"></div></div>
<div data-field="USUARIOSOLICITANTE" data-label="Usu&amp;#225;rio Solicitante">
  <select name="{P}USUARIOSOLICITANTE" id="usr" multiple>{users_sel}</select></div>
<a id="ctl00_Main_K_INCLUSODESERVIOSACESSRIOS_FilterControl_FilterButton" title="Filtrar" class="btn btn-primary filter-button"
   href="javascript:__doPostBack('ctl00$Main$K_INCLUSODESERVIOSACESSRIOS$FilterControl$FilterButton','')">🔍</a>
{grid}
</form>
<script>
function __doPostBack(t, a) {{ var f = document.getElementById('aspnetForm');
  f.__EVENTTARGET.value = t; f.__EVENTARGUMENT.value = a; f.submit(); }}
$('.s2').select2({{width: '300px'}});
var STATUS = {[{"id": v, "text": t} for v, t in status]!r};
$('#status').select2({{width: '300px', closeOnSelect: false,
  ajax: {{ transport: function(params, ok) {{
    var q = (params.data.term || '').toLowerCase();
    setTimeout(function() {{ ok({{results: STATUS.filter(s => s.text.toLowerCase().includes(q))}}); }}, 300);
  }} }} }});
var USUARIOS = {USUARIOS!r};
$('#usr').select2({{width: '300px', placeholder: 'Todos', minimumInputLength: 3,
  ajax: {{ delay: 300, transport: function(params, ok) {{
    var q = (params.data.term || '').toLowerCase();
    setTimeout(function() {{ ok({{results: USUARIOS.filter(u => u.toLowerCase().includes(q)).map(u => ({{id: u, text: u}}))}}); }}, 400);
  }} }} }});
</script>"""
        self._html(corpo)

    def _html_processos(self) -> str:
        itens = []
        with Estado.lock:
            for p in Estado.processos:
                if time.time() < p["fim"]:
                    itens.append("<li><a href='#'><span class='desc'><i class='fa fa-spinner'></i> Exportar para Excel"
                                 " - InclusãodeServiçosAcessórios<br><i>Processando</i></span></a></li>")
                else:
                    itens.append(f"<li><a href='/download/{p['id']}'><span class='desc'><i class='fa fa-check process-success'></i>"
                                 " Exportar para Excel - InclusãodeServiçosAcessórios<br>"
                                 "<i class='process-success'>Concluído</i></span></a></li>")
        return "".join(itens)

    def _download(self):
        n = int(self.path.rsplit("/", 1)[1])
        f = next(p["filtro"] for p in Estado.processos if p["id"] == n)
        self._enviar(gerar_exportacao(f["de"]),
                     "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                     headers={"Content-Disposition": "attachment; filename=InclusaodeServicosAcessorios.xlsx"})


COLUNAS = ["Viagem", "Documento Logística Original", "Tipo Solicitação", "Serviço", "Valor",
           "Status Solicitação", "Status Geração", "Status Complementar", "Fluxo", "Processo Atual",
           "Alçada", "Inclusão", "Usuário"]
GELO = "Troca de Gelo por Volume - Cliente/Transportador"


def gerar_exportacao(dia: str) -> bytes:
    """Gera um .xlsx no mesmo formato do 'Exportar para Excel' do TMS (dados fictícios)."""
    from datetime import datetime as dt
    from io import BytesIO

    from openpyxl import Workbook

    d = dt.strptime(dia, "%d/%m/%Y")
    linhas = [
        (None, 1111, "Receita de Documento", GELO, 940, "Aguardando Aprovação", d.replace(hour=11, minute=5, second=24)),
        (None, 1111, "Receita de Documento", GELO, 1410, "Aguardando Aprovação", d.replace(hour=11, minute=4, second=44)),
        (None, 2222, "Receita de Documento", "Complemento de Carga - Cliente/Transportador", 188, "Aprovada",
         d.replace(hour=12, minute=0, second=1)),
        ("2026/0000001-99", 3333, "Receita de Documento", GELO, 500, "Aprovada", d.replace(hour=15, minute=30, second=0)),
    ]
    wb = Workbook()
    ws = wb.active
    ws.title = "InclusãodeServiçosAcessórios_1"
    ws.append(COLUNAS)
    for viagem, doc, tipo, serv, valor, status, inclusao in linhas:
        ws.append([viagem, doc, tipo, serv, valor, status, None, None, "Emissão da Receita (Operação & CS)",
                   "Aprovação de CS", "Receita - CLIENTE TESTE", inclusao, "usuario.teste"])
        ws.cell(ws.max_row, 2).number_format = "0"
        ws.cell(ws.max_row, 5).number_format = "0"
        ws.cell(ws.max_row, 12).number_format = "dd/mm/yyyy hh:mm:ss"
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def iniciar() -> tuple[ThreadingHTTPServer, str]:
    Estado.processos, Estado.ultimo_filtro = [], {}
    srv = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    return srv, f"http://127.0.0.1:{srv.server_address[1]}"
