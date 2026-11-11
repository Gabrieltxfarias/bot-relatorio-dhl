"""Automação da tela de Inclusão de Serviços Acessórios do TMS (Benner) com Playwright."""
from __future__ import annotations

import logging
import re
import time
from datetime import date, datetime
from pathlib import Path
from urllib.parse import urlparse

from playwright.sync_api import BrowserContext, Download, Locator, Page, TimeoutError as PWTimeout, sync_playwright

from . import credenciais
from .config import RAIZ

log = logging.getLogger(__name__)

# Acha o bloco <div data-field> de um filtro pelo rótulo. O data-label do Benner
# vem com entidades HTML duplamente codificadas ("Inclus&amp;#227;o"), por isso
# decodificamos duas vezes e comparamos sem acentos/maiúsculas.
JS_ACHAR_CAMPO = """(rotulo) => {
  const dec = s => { const t = document.createElement('textarea'); t.innerHTML = s; return t.value; };
  const norm = s => dec(dec(s || '')).normalize('NFD').replace(/[\\u0300-\\u036f]/g, '')
                      .replace(/[:*]/g, '').toLowerCase().replace(/\\s+/g, ' ').trim();
  const alvo = norm(rotulo);
  const marca = 'campo-' + alvo.replace(/[^a-z0-9]+/g, '-');
  const marcar = el => { el.setAttribute('data-bot-campo', marca); return marca; };

  // 1) blocos <div data-field=... data-label=...> do Benner
  let exato = null, parcial = null;
  for (const d of document.querySelectorAll('[data-field]')) {
    if (!d.querySelector('select, .select2')) continue;
    const nomes = [norm(d.getAttribute('data-label')), norm(d.getAttribute('data-field'))];
    if (nomes.includes(alvo)) { exato = exato || d; }
    else if (nomes.some(n => n.includes(alvo))) { parcial = parcial || d; }
  }
  if (exato) return marcar(exato);

  // 2) pelo texto do rótulo na tela (exato; depois contendo): o 1º campo que vem depois dele
  const textoProprio = el => norm([...el.childNodes].filter(n => n.nodeType === 3).map(n => n.textContent).join(' '));
  const candidatos = [...document.querySelectorAll('label, span, td, th, div, b, strong')]
    .filter(el => el.offsetParent !== null && textoProprio(el));
  const ehCampo = el => el.tagName === 'SELECT' || (el.classList && el.classList.contains('select2-container'))
    || (el.tagName === 'INPUT' && el.classList.contains('select2-offscreen'));
  const depois = r => {
    const w = document.createTreeWalker(document.body, NodeFilter.SHOW_ELEMENT);
    w.currentNode = r;
    for (let i = 0, el; i < 300 && (el = w.nextNode()); i++) {
      if (ehCampo(el)) return el.parentElement;
    }
    return null;
  };
  for (const teste of [t => t === alvo, t => t.includes(alvo)]) {
    for (const r of candidatos.filter(el => teste(textoProprio(el)))) {
      const bloco = depois(r);
      if (bloco) return marcar(bloco);
    }
  }
  // 3) por último, nome parcial (ex.: "Solicitante" -> "Usuário Solicitante")
  return parcial ? marcar(parcial) : null;
}"""

# Seleciona opções de um <select> (select2) pelo texto, se todas existirem nele.
JS_SELECIONAR_POR_TEXTO = """(el, textos) => {
  const sel = el.querySelector('select');
  if (!sel) return false;
  const norm = s => s.normalize('NFD').replace(/[\\u0300-\\u036f]/g, '').toLowerCase().trim();
  const opts = [...sel.options];
  const escolhidas = textos.map(t => opts.find(o => norm(o.text) === norm(t)));
  if (escolhidas.some(o => !o)) return false;
  const valores = escolhidas.map(o => o.value);
  if (window.jQuery) { jQuery(sel).val(sel.multiple ? valores : valores[0]).trigger('change'); }
  else { opts.forEach(o => o.selected = valores.includes(o.value)); sel.dispatchEvent(new Event('change', {bubbles: true})); }
  return true;
}"""


# A pesquisa de satisfação (Amplifique.me) abre por cima da tela e pode
# bloquear os cliques; escondemos via CSS em toda página carregada.
JS_ESCONDER_POPUPS = """(() => {
  const css = '#ampl-survey, .ampl-survey { display: none !important; visibility: hidden !important; }';
  const aplicar = () => { if (document.getElementById('bot-esconder-popups')) return;
    const s = document.createElement('style'); s.id = 'bot-esconder-popups'; s.textContent = css;
    (document.head || document.documentElement).appendChild(s); };
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', aplicar); else aplicar();
})();"""


class ErroBot(RuntimeError):
    pass


class BotTMS:
    def __init__(self, cfg: dict):
        self.cfg = cfg
        self.nav = cfg["navegador"]
        self.urls = cfg["urls"]
        self.sel = {k: (v or "").strip() for k, v in (cfg.get("seletores") or {}).items()}
        self.pasta_downloads = RAIZ / "downloads"
        self.pasta_logs = RAIZ / "logs"
        self.pasta_downloads.mkdir(exist_ok=True)
        self.pasta_logs.mkdir(exist_ok=True)
        self._downloads: list[Download] = []
        self._pw = None
        self.ctx: BrowserContext | None = None
        self.page: Page | None = None

    # ------------------------------------------------------------------ ciclo
    def __enter__(self):
        self._pw = sync_playwright().start()
        opcoes = dict(
            user_data_dir=str(RAIZ / self.nav["pasta_perfil"]),
            headless=self.nav.get("headless", False),
            accept_downloads=True,
            viewport={"width": 1440, "height": 900},
            locale="pt-BR",
        )
        if self.nav.get("executavel"):
            opcoes["executable_path"] = self.nav["executavel"]
        elif self.nav.get("canal"):
            opcoes["channel"] = self.nav["canal"]
        self.ctx = self._pw.chromium.launch_persistent_context(**opcoes)
        self.ctx.set_default_timeout(self.nav.get("timeout_padrao_ms", 30000))
        self.ctx.on("page", lambda p: p.on("download", lambda d: self._downloads.append(d)))
        for p in self.ctx.pages:
            p.on("download", lambda d: self._downloads.append(d))
        self.ctx.add_init_script(JS_ESCONDER_POPUPS)
        self.page = self.ctx.pages[0] if self.ctx.pages else self.ctx.new_page()
        return self

    def __exit__(self, tipo, exc, tb):
        if exc and self.page:
            self.registrar_tela("erro")
        try:
            self.ctx.close()
        finally:
            self._pw.stop()

    def registrar_tela(self, nome: str) -> None:
        """Salva screenshot + HTML da página em logs/ (útil para ajustar seletores)."""
        base = self.pasta_logs / f"{datetime.now():%Y%m%d_%H%M%S}_{nome}"
        try:
            self.page.screenshot(path=f"{base}.png", full_page=True)
            Path(f"{base}.html").write_text(self.page.content(), encoding="utf-8")
            log.info("Tela registrada em %s.png/.html", base)
        except Exception:
            pass

    def print_conferencia(self, nome: str) -> None:
        """Print da tela em conferencia/ para você ver o que o bot preencheu."""
        pasta = RAIZ / "conferencia"
        pasta.mkdir(exist_ok=True)
        caminho = pasta / f"{nome}.png"
        try:
            self._fechar_listas_select2()
            self.page.screenshot(path=str(caminho), full_page=True)
            log.info("Print para conferência: %s", caminho)
        except Exception as e:
            log.warning("Não consegui salvar o print %s: %s", caminho, e)
        limite = time.time() - 60 * 86400  # mantém 60 dias
        for velho in pasta.glob("*.png"):
            if velho.stat().st_mtime < limite:
                velho.unlink(missing_ok=True)

    # ------------------------------------------------------------ utilitários
    def _esperar(self) -> None:
        """Aguarda postbacks/carregamentos do ASP.NET."""
        try:
            self.page.wait_for_load_state("networkidle", timeout=20000)
        except PWTimeout:
            pass

    def _ir(self, url: str) -> None:
        self.page.goto(url, wait_until="domcontentloaded")
        self._esperar()

    def _clicar_postback(self, loc: Locator) -> None:
        """Clica num botão que dispara __doPostBack e espera a página recarregar.

        Se o postback for parcial (UpdatePanel/AJAX) e não houver navegação,
        apenas aguarda a rede ficar ociosa.
        """
        try:
            with self.page.expect_navigation(timeout=15000, wait_until="domcontentloaded"):
                loc.click()
        except PWTimeout:
            pass
        self._esperar()

    def _el(self, chave: str) -> Locator:
        loc = self.page.locator(self.sel[chave]).first
        try:
            loc.wait_for(state="visible")
        except PWTimeout:
            raise ErroBot(f"Elemento '{chave}' ({self.sel[chave]}) não apareceu na tela.")
        return loc

    # ------------------------------------------------------------------ login
    def abrir_e_logar(self) -> None:
        url = self.urls["inicial"]
        host = urlparse(url).hostname
        log.info("Abrindo %s", url)
        self._ir(url)
        if self._logado(host):
            log.info("Sessão salva reaproveitada — já logado.")
            return

        log.info("Login necessário.")
        usuario, senha = credenciais.obter()
        fim = time.time() + self.nav.get("timeout_login_s", 300)
        avisou = False
        while time.time() < fim:
            if self._logado(host):
                log.info("Login concluído.")
                return
            if usuario and senha:
                self._tentar_login_automatico(usuario, senha)
            elif not avisou:
                log.warning("Sem credenciais salvas: faça o login na janela aberta "
                            "(ou rode 'python salvar_credenciais.py').")
                avisou = True
            self.page.wait_for_timeout(2000)
        raise ErroBot("Tempo de login esgotado (MFA pendente ou credenciais incorretas?).")

    def _logado(self, host: str) -> bool:
        if urlparse(self.page.url).hostname != host:
            return False
        try:
            return self.page.locator("input[type=password]:visible").count() == 0
        except Exception:
            return False

    def _tentar_login_automatico(self, usuario: str, senha: str) -> None:
        """Preenche telas de login comuns (formulário simples ou Microsoft SSO, em etapas)."""
        p = self.page
        campo_usuario = p.locator(
            "input[type=email]:visible, input[name='loginfmt']:visible, input[name*='user' i]:visible, "
            "input[id*='user' i]:visible, input[name*='login' i]:visible"
        ).first
        campo_senha = p.locator("input[type=password]:visible").first
        try:
            if campo_usuario.count() and not campo_usuario.input_value():
                campo_usuario.fill(usuario)
                if not campo_senha.count():
                    campo_usuario.press("Enter")
                    self._esperar()
                    return
            if campo_senha.count() and not campo_senha.input_value():
                campo_senha.fill(senha)
                campo_senha.press("Enter")
                self._esperar()
                return
            manter = p.get_by_role("button", name=re.compile(r"^(sim|yes)$", re.I))  # "Continuar conectado?"
            if manter.count() and manter.first.is_visible():
                manter.first.click()
                self._esperar()
        except Exception as e:
            log.debug("Etapa de login ignorada: %s", e)

    # ----------------------------------------------------------------- passos
    def escolher_papel(self) -> None:
        log.info("Selecionando papel: %s", self.urls["papel"])
        self._ir(self.urls["papel"])

    def abrir_tela(self) -> None:
        log.info("Abrindo Inclusão de Serviços Acessórios")
        self._ir(self.urls["tela"])
        self._el("botao_pesquisar")
        papel = self.cfg.get("papel_nome", "")
        if papel and not self.page.get_by_text(papel, exact=False).count():
            self.registrar_tela("papel_nao_confirmado")
            raise ErroBot(f"O papel '{papel}' não aparece no cabeçalho. A troca de papel falhou?")

    def preencher_filtros(self, data_de: date, data_ate: date) -> None:
        f = self.cfg["filtros"]
        fmt = f.get("formato_data", "%d/%m/%Y")
        de_txt, ate_txt = data_de.strftime(fmt), data_ate.strftime(fmt)
        log.info("Filtros: %s | %s | %s a %s | %s", f["tipo_solicitacao"], f["status"],
                 de_txt, ate_txt, f["usuario_solicitante"])

        self._selecionar_com_postback("tipo_solicitacao", f["tipo_solicitacao"])
        self._select2("status_campo", "status", list(f["status"]))
        self._data("data_de", de_txt)
        self._data("data_ate", ate_txt)
        self._select2("usuario_campo", "usuario_solicitante", [f["usuario_solicitante"]])

    def _selecionar_com_postback(self, chave: str, texto: str) -> None:
        """Seleciona uma opção num <select> que pode recarregar a página (AutoPostBack)."""
        for _ in range(2):
            campo = self._el(chave)
            try:
                with self.page.expect_navigation(timeout=6000, wait_until="domcontentloaded"):
                    campo.select_option(label=texto)
                log.info("A tela recarregou após escolher '%s'.", texto)
            except PWTimeout:
                pass  # não recarregou
            self._esperar()
            self._el("botao_pesquisar")  # tela pronta de novo
            atual = self._el(chave).evaluate("e => e.options[e.selectedIndex] && e.options[e.selectedIndex].text")
            if (atual or "").strip().lower() == texto.strip().lower():
                return
        raise ErroBot(f"Não consegui manter '{texto}' selecionado em {chave} (ficou '{atual}').")

    def _data(self, chave: str, texto: str) -> None:
        campo = self._el(chave)
        campo.fill(texto)
        campo.dispatch_event("change")
        campo.press("Escape")  # fecha o datepicker
        if campo.input_value() != texto:
            raise ErroBot(f"Não consegui preencher {chave} com {texto} (ficou '{campo.input_value()}').")

    def _bloco(self, chave_sel: str, chave_rotulo: str) -> Locator:
        if self.sel.get(chave_sel):
            return self._el(chave_sel)
        rotulo = self.cfg["rotulos"][chave_rotulo]
        fim = time.time() + 20  # a tela pode estar recarregando (postback)
        marca = None
        while time.time() < fim:
            try:
                marca = self.page.evaluate(JS_ACHAR_CAMPO, rotulo)
            except Exception:  # contexto destruído no meio de um recarregamento
                marca = None
            if marca:
                break
            self.page.wait_for_timeout(1000)
        if not marca:
            self.registrar_tela(f"campo_{chave_rotulo}_nao_encontrado")
            raise ErroBot(f"Campo '{rotulo}' não encontrado. Preencha seletores.{chave_sel} no config.yaml.")
        bloco = self.page.locator(f"[data-bot-campo='{marca}']").first
        log.info("Campo '%s' encontrado.", rotulo)
        return bloco

    def _select2(self, chave_sel: str, chave_rotulo: str, textos: list[str]) -> None:
        """Preenche um campo select2 (múltipla escolha ou com busca)."""
        bloco = self._bloco(chave_sel, chave_rotulo)
        # Limpa seleções anteriores
        remover = bloco.locator(".select2-selection__choice__remove")
        while remover.count():
            remover.first.click(force=True)
            self._fechar_listas_select2()
        # 1) opções já carregadas no <select>: seleciona direto
        if bloco.evaluate(JS_SELECIONAR_POR_TEXTO, textos):
            self._esperar()
            return
        # 2) busca digitando (select2 com opções carregadas sob demanda)
        rx_carregando = re.compile(r"(buscando|searching|carregando|loading|digite|please enter)", re.I)
        for texto in textos:
            self._fechar_listas_select2()
            busca = bloco.locator(".select2-search__field").first
            if busca.count():
                busca.focus()  # foco direto: a lista aberta pode estar cobrindo a caixa
            else:  # select2 simples: abre e usa a busca do dropdown
                bloco.locator(".select2-selection").first.click(force=True)
                busca = self.page.locator(".select2-container--open .select2-search__field").first
                busca.focus()
            busca.fill("")
            busca.press_sequentially(texto, delay=40)
            opcoes = self.page.locator(".select2-container--open .select2-results__option")
            exata = opcoes.filter(has_text=re.compile(rf"^\s*{re.escape(texto)}\s*$", re.I))
            contem = opcoes.filter(has_text=re.compile(re.escape(texto), re.I))
            try:
                contem.first.wait_for(state="visible", timeout=15000)
            except PWTimeout:
                vistos = [t for t in opcoes.all_inner_texts() if not rx_carregando.search(t)]
                raise ErroBot(f"'{texto}' não apareceu na lista. Opções vistas: {vistos[:10]}")
            alvo = exata.first if exata.count() else contem.first
            if alvo.get_attribute("aria-selected") == "true":
                log.info("'%s' já estava selecionado.", texto)
            else:
                alvo.click()
                self._esperar()
            busca.fill("")
        self._fechar_listas_select2()

    def _fechar_listas_select2(self) -> None:
        """Fecha listas do select2 que ficam abertas depois de escolher uma opção."""
        if self.page.locator(".select2-container--open").count():
            self.page.keyboard.press("Escape")
            self.page.wait_for_timeout(200)
        if self.page.locator(".select2-container--open").count():
            self.page.evaluate("() => { if (window.jQuery) jQuery('select').each(function () {"
                               " try { if (jQuery(this).data('select2')) jQuery(this).select2('close'); }"
                               " catch (e) {} }); }")

    def pesquisar(self) -> None:
        log.info("Pesquisando...")
        self._clicar_postback(self._el("botao_pesquisar"))

    def exportar_excel(self, nome: str) -> Path:
        barra = self.sel["barra_processos"]
        itens = f"{barra} ul.dropdown-menu-list > li"
        pronto = self.sel.get("processo_concluido") or ".process-success"
        erro = self.sel.get("processo_erro") or ".process-error, .process-fail, .process-danger"
        # Limpa processos já concluídos, para o próximo item pronto ser o nosso
        self.page.evaluate("() => { try { Benner.AsyncProcesses.clearReadyProcesses(); } catch (e) {} }")
        self.page.wait_for_timeout(1000)
        antes = self._prontos(itens, pronto)

        log.info("Exportando para Excel...")
        self._el("tres_pontos").click()
        self._downloads.clear()
        self._clicar_postback(self._el("exportar_excel"))

        log.info("Aguardando a exportação terminar (barra de processos)...")
        fim = time.time() + self.nav.get("timeout_relatorio_s", 900)
        ultimo_status, clicou_em = "", 0.0
        while time.time() < fim:
            if self._downloads:
                return self._salvar_download(self._downloads[-1], nome)
            status = " | ".join(" ".join(t.split()) for t in self.page.locator(itens).all_inner_texts() if t.strip())
            if status != ultimo_status:
                log.info("Barra de processos: %s", status or "(vazia)")
                ultimo_status = status
            if self.page.locator(itens).filter(has=self.page.locator(erro)).count():
                self.registrar_tela("erro_exportacao")
                raise ErroBot(f"O TMS indicou erro na exportação: {status}")
            prontos = self._prontos(itens, pronto)
            # (re)tenta clicar a cada 20s, caso o 1º clique não tenha iniciado o download
            if len(prontos) > len(antes) and time.time() - clicou_em > 20:
                novos = [i for i, t in prontos if t not in {t for _, t in antes}]
                idx = novos[0] if novos else prontos[0][0]
                log.info("Exportação concluída. Baixando...")
                self._clicar_item_processo(itens, idx)
                clicou_em = time.time()
            self.page.wait_for_timeout(1500)
        self.registrar_tela("timeout_exportacao")
        raise ErroBot("Tempo esgotado aguardando a exportação (veja o HTML salvo em logs/).")

    def _prontos(self, seletor_itens: str, seletor_pronto: str) -> list[tuple[int, str]]:
        """Itens da barra de processos marcados como concluídos: [(índice, texto)]."""
        loc = self.page.locator(seletor_itens)
        res = []
        for i in range(loc.count()):
            el = loc.nth(i)
            if el.locator(seletor_pronto).count() and "excel" in el.inner_text().lower():
                res.append((i, " ".join(el.inner_text().split())))
        return res

    def _clicar_item_processo(self, seletor_itens: str, idx: int) -> None:
        item = self.page.locator(seletor_itens).nth(idx)
        if not item.is_visible():  # abre o menu da barra de processos
            self.page.locator(f"{self.sel['barra_processos']} > a.dropdown-toggle").click()
            try:
                item.wait_for(state="visible", timeout=5000)
            except PWTimeout:
                pass
        alvo = item.locator("a").first if item.locator("a").count() else item.locator(".desc").first
        alvo.click(force=True)

    def _salvar_download(self, dl: Download, nome: str) -> Path:
        ext = Path(dl.suggested_filename).suffix or ".xlsx"
        destino = self.pasta_downloads / f"relatorio_{nome}_{datetime.now():%Y%m%d_%H%M%S}{ext}"
        dl.save_as(str(destino))
        log.info("Relatório baixado: %s", destino)
        return destino

    # --------------------------------------------------------------- completo
    def executar(self, data_de: date, data_ate: date | None = None) -> Path:
        data_ate = data_ate or data_de
        self.abrir_e_logar()
        self.escolher_papel()
        self.abrir_tela()
        nome = f"{data_de:%Y-%m-%d}" + ("" if data_de == data_ate else f"_a_{data_ate:%Y-%m-%d}")
        self.preencher_filtros(data_de, data_ate)
        self.print_conferencia(f"{nome}_1_filtros_preenchidos")
        self.pesquisar()
        self.print_conferencia(f"{nome}_2_resultado_pesquisa")
        return self.exportar_excel(nome)
