# -*- coding: utf-8 -*-
"""
RPA Download DCTF
=================
Baixa, de uma vez, os recibos de entrega e as declarações da DCTFWeb no e-CAC.

Como funciona
1. A pessoa escolhe o período, a pasta de destino e clica em "Iniciar download".
2. O robô abre o Microsoft Edge (ou o Google Chrome) na página de login do e-CAC.
3. A própria pessoa faz o login (gov.br com certificado digital ou CPF e senha).
   O robô nunca vê nem guarda senha ou certificado.
4. O robô abre a DCTFWeb, pesquisa o período e baixa os PDFs com os nomes
   "Recibo - MM-AAAA.pdf" e "DCTF - MM-AAAA.pdf" (o Windows não aceita "/" em
   nome de arquivo, por isso a competência usa hífen).
5. No fim, grava um relatório em Excel na mesma pasta.

Se alguma tela do e-CAC mudar, ajuste o bloco ECAC logo abaixo ou coloque um
arquivo config.json ao lado do programa com as chaves que quiser trocar.

Criado por Diego Verciani.
"""

import base64
import datetime as dt
import json
import os
import queue
import re
import sys
import threading
import time
import traceback
from dataclasses import dataclass, field
from pathlib import Path

APP_NOME = "RPA Download DCTF"
APP_VERSAO = "1.0"
AUTOR = "Diego Verciani"

# --------------------------------------------------------------------------
# Identidade visual (mesma paleta e fonte da vitrine de aplicações)
# --------------------------------------------------------------------------
COR = {
    "navy": "#061A2F", "navy2": "#0B2948", "navy3": "#12385B",
    "card": "#0C2239", "line": "#233C55", "gold": "#E5AC32",
    "gold_hover": "#F0BD4F", "white": "#F5F7FA", "gray": "#AAB8C8",
    "gray2": "#8496AB", "erro": "#E07A70", "ok": "#6FCF97",
}

# --------------------------------------------------------------------------
# Pontos de calibração do e-CAC / DCTFWeb
# Todos os textos são expressões regulares, sem diferenciar maiúsculas.
# --------------------------------------------------------------------------
ECAC = {
    "url_login": "https://cav.receita.fazenda.gov.br/autenticacao/login",
    # Endereço que indica que o login terminou (página inicial do e-CAC)
    "regex_logado": r"cav\.receita\.fazenda\.gov\.br/ecac",
    "regex_tela_login": r"/autenticacao|sso\.acesso\.gov\.br|acesso\.gov\.br/login",
    # Caminho até a DCTFWeb. Se descobrir o endereço direto, preencha url_dctfweb.
    "url_dctfweb": "",
    "menu_declaracoes": [r"^\s*Declara[cç][oõ]es e Demonstrativos\s*$"],
    "servico_dctfweb": [r"Assinar e Transmitir DCTFWeb", r"^\s*DCTFWeb\s*$"],
    # Textos que só aparecem na tela de consulta da DCTFWeb
    "marcas_tela_dctfweb": [r"Rela[cç][aã]o de Declara[cç][oõ]es",
                            r"Sou Procurador e desejo visualizar"],
    # Filtro de período de apuração
    "campo_periodo": r"per[ií]odo|apura[cç]|compet|\bPA\b",
    "campo_periodo_excluir": r"transmiss|recibo|processo|outorg|data",
    "botao_pesquisar": [r"^\s*Pesquisar\s*$", r"^\s*Consultar\s*$",
                        r"^\s*Buscar\s*$", r"^\s*Filtrar\s*$"],
    "sem_resultados": r"nenhum(a)? (registro|declara)|n[aã]o (foram )?encontrad",
    # Situações aceitas (a declaração vigente é a "Ativa")
    "situacoes_aceitas": ["Ativa"],
    # Ações por linha da lista. "linha" = procurado dentro da linha;
    # "tela" = procurado na tela inteira, caso a ação abra uma tela de detalhe.
    "acao_recibo": {
        "linha": [r"visualizar\s+recibo", r"download\s+recibo", r"\brecibo\b"],
        "tela": [r"download\s+recibo", r"baixar\s+recibo", r"imprimir\s+recibo",
                 r"\brecibo\b"],
        "excluir": r"xml|darf|extrato|retific|excluir|editar|transmit",
    },
    "acao_declaracao": {
        "linha": [r"declara[cç][aã]o\s+completa", r"visualizar\s+declara[cç][aã]o",
                  r"\bdeclara[cç][aã]o\b"],
        "tela": [r"declara[cç][aã]o\s+completa", r"baixar\s+declara",
                 r"imprimir\s+declara", r"gerar\s+pdf", r"\bimprimir\b"],
        "excluir": r"xml|darf|extrato|retific|excluir|editar|transmit|procurador",
    },
    # Botão que abre o menu de serviços de uma linha, se as ações ficarem escondidas
    "menu_da_linha": [r"^\s*servi[cç]os\s*$", r"^\s*a[cç][oõ]es\s*$",
                      r"^\s*op[cç][oõ]es\s*$", r"more_vert|more_horiz",
                      r"^\s*(⋮|…|\.\.\.)\s*$"],
    "seletor_menu_aberto": "[role=menu], .dropdown-menu.show, .dropdown-menu.open, "
                           ".mat-menu-panel, .mat-mdc-menu-panel, .p-menu, .p-tieredmenu, "
                           ".cdk-overlay-pane, .popover",
    "botao_voltar": [r"^\s*Voltar\s*$", r"^\s*Retornar\s*$"],
    "proxima_pagina": [r"pr[oó]xima\s*p[aá]gina", r"^\s*pr[oó]xim[ao]\s*$",
                       r"next\s*page", r"^\s*›\s*$", r"^\s*»\s*$", r"^\s*>\s*$"],
    "captcha": r"hcaptcha\.com|recaptcha|captcha",
    # Navegadores tentados, em ordem. "chromium" = navegador do Playwright (testes)
    "navegadores": ["msedge", "chrome"],
    "headless": False,
    "tempo_pdf": 45,
}

CATEGORIAS = [
    (r"13[ºo°]?\s*sal[aá]rio", "13º Salário"),
    (r"espet[aá]culo\s+desportivo", "Espetáculo Desportivo"),
    (r"aferi[cç][aã]o", "Aferição de Obra"),
    (r"reclamat[oó]ria", "Reclamatória Trabalhista"),
    (r"\banual\b", "Anual"),
    (r"\bgeral\b", "Geral"),
]
SITUACOES = ["Ativa", "Retificada", "Em andamento", "Excluída", "Cancelada",
             "Faseamento", "Suspensa"]

RE_COMP = re.compile(r"(?<![\d/])(0[1-9]|1[0-2])/(\d{4})(?![\d/])")
RE_ANO = re.compile(r"(?<![\d/.,-])(20\d{2})(?![\d/.,-])")


# --------------------------------------------------------------------------
# Utilidades
# --------------------------------------------------------------------------
def recurso(*partes):
    """Caminho de um arquivo empacotado (funciona no .exe do PyInstaller)."""
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base.joinpath(*partes)


def pasta_do_programa():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def pasta_de_dados():
    base = os.environ.get("LOCALAPPDATA") or str(Path.home() / ".local" / "share")
    p = Path(base) / APP_NOME
    p.mkdir(parents=True, exist_ok=True)
    return p


def carregar_config():
    """Mescla o config.json (se existir ao lado do programa) com o padrão."""
    cfg = json.loads(json.dumps(ECAC))
    candidatos = [pasta_do_programa() / "config.json"]
    extra = os.environ.get("RPA_DCTF_CONFIG")
    if extra:
        candidatos.append(Path(extra))
    for c in candidatos:
        try:
            if c.exists():
                cfg.update(json.loads(c.read_text(encoding="utf-8")))
        except Exception:
            pass
    return cfg


def ler_competencia(texto):
    """'01/2026' -> (2026, 1). Devolve None se inválida."""
    m = re.fullmatch(r"\s*(0?[1-9]|1[0-2])\s*/\s*(\d{4})\s*", texto or "")
    if not m:
        return None
    return int(m.group(2)), int(m.group(1))


def meses_do_periodo(ini, fim):
    a, m = ini
    out = []
    while (a, m) <= fim:
        out.append((a, m))
        m += 1
        if m > 12:
            a, m = a + 1, 1
    return out


class Cancelado(Exception):
    pass


class ErroRobo(Exception):
    pass


@dataclass
class Linha:
    texto: str
    competencia: str            # "01/2026" ou "13/2025" (anual)
    ano: int
    mes: int                    # 13 = anual / 13º salário
    categoria: str
    situacao: str

    @property
    def chave(self):
        return self.texto

    @property
    def comp_arquivo(self):
        return f"{self.mes:02d}-{self.ano}"


@dataclass
class Pedido:
    """Pergunta do robô para a janela; o robô espera a resposta."""
    evento: threading.Event = field(default_factory=threading.Event)
    resposta: str = ""

    def responder(self, valor):
        self.resposta = valor
        self.evento.set()

    def respondido(self):
        return self.evento.is_set()


class Ponte:
    """Canal entre o robô (thread de trabalho) e a janela (thread principal)."""

    def __init__(self):
        self.fila = queue.Queue()
        self.cancelar = threading.Event()

    def log(self, msg, tipo="info"):
        self.fila.put(("log", msg, tipo))

    def status(self, msg):
        self.fila.put(("status", msg))

    def progresso(self, atual, total):
        self.fila.put(("prog", atual, total))

    def pedir(self, tipo, titulo, texto, botoes):
        p = Pedido()
        self.fila.put(("pedido", tipo, titulo, texto, botoes, p))
        return p

    def login_detectado(self):
        self.fila.put(("login_ok",))


# --------------------------------------------------------------------------
# Scripts executados dentro da página
# --------------------------------------------------------------------------
JS_ACHAR = r"""
(root, cfg) => {
  root = root || document;
  if (cfg.raiz) {
    const r = Array.from(document.querySelectorAll(cfg.raiz)).find(e => {
      const b = e.getBoundingClientRect(); return b.width > 0 && b.height > 0;
    });
    if (!r) return null;
    root = r;
  }
  const sel ='a,button,[role=button],[role=menuitem],[role=link],[role=tab],' +
    'input[type=button],input[type=submit],img,i,mat-icon,span,svg,li,label,div,' +
    '[title],[aria-label],[mattooltip],[ptooltip],[data-original-title],[ngbtooltip]';
  const cands = Array.from(root.querySelectorAll(sel));
  if (root !== document && root.matches && root.matches(sel)) cands.unshift(root);
  const vis = e => {
    const r = e.getBoundingClientRect(); const s = getComputedStyle(e);
    return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none';
  };
  const inativo = e => {
    const c = e.closest('button,a,li,[role=button]') || e;
    return c.disabled || c.getAttribute('aria-disabled') === 'true' ||
           /(^|\s)(disabled|p-disabled|mat-button-disabled)(\s|$)/.test(c.className || '');
  };
  const partes = e => [e.getAttribute('title'), e.getAttribute('aria-label'), e.getAttribute('alt'),
      e.getAttribute('mattooltip'), e.getAttribute('ptooltip'), e.getAttribute('data-original-title'),
      e.getAttribute('ngbtooltip'), e.getAttribute('ng-reflect-message'),
      (e.tagName === 'INPUT' ? e.value : ''),
      (e.children.length <= 3 ? e.innerText : '')]
    .filter(Boolean).map(s => s.replace(/\s+/g, ' ').trim()).filter(Boolean);
  const txt = e => partes(e).join(' | ');
  const ex = cfg.excluir ? new RegExp(cfg.excluir, 'i') : null;
  const clicavel = 'a,button,[role=button],[role=menuitem],[role=link],[role=tab],input,[onclick]';
  for (const p of cfg.padroes) {
    const re = new RegExp(p, 'i');
    let melhor = null, nota = 1e9;
    for (const e of cands) {
      const t = txt(e);
      if (!t || t.length > 140 || !partes(e).some(s => re.test(s))) continue;
      if (ex && ex.test(t)) continue;
      if (!vis(e)) continue;
      if (cfg.ativos && inativo(e)) continue;
      // texto mais curto vence; elementos clicáveis têm preferência
      const n = t.length + (e.closest(clicavel) ? 0 : 1000);
      if (n < nota) { melhor = e; nota = n; }
    }
    if (melhor) {
      const alvo = melhor.closest(clicavel) || melhor;
      const id = 'rpa' + Math.random().toString(36).slice(2, 10);
      alvo.setAttribute('data-rpa-alvo', id);
      return {id: id, texto: txt(melhor)};
    }
  }
  return null;
}
"""

JS_LINHAS = r"""
() => {
  const out = [];
  document.querySelectorAll('[data-rpa-linha]').forEach(r => r.removeAttribute('data-rpa-linha'));
  const rows = Array.from(document.querySelectorAll('tr, [role=row], mat-row'));
  rows.forEach((r, i) => {
    if (r.querySelector('th') && !r.querySelector('td')) return;
    const rect = r.getBoundingClientRect();
    if (rect.width === 0 || rect.height === 0) return;
    const t = (r.innerText || '').replace(/\s+/g, ' ').trim();
    if (!t || t.length > 600) return;
    r.setAttribute('data-rpa-linha', String(i));
    out.push({i: i, t: t});
  });
  return out;
}
"""

JS_CAMPOS_PERIODO = r"""
(cfg) => {
  const ins = Array.from(document.querySelectorAll('input')).filter(e => {
    const t = (e.getAttribute('type') || 'text').toLowerCase();
    if (!['text', 'month', 'tel', 'search', 'date'].includes(t)) return false;
    const r = e.getBoundingClientRect();
    return r.width > 0 && r.height > 0 && !e.disabled && !e.readOnly;
  });
  const rotulo = e => {
    let s = [e.placeholder, e.name, e.id, e.getAttribute('aria-label'),
             e.getAttribute('formcontrolname'), e.title].filter(Boolean).join(' ');
    if (e.id) {
      const l = document.querySelector('label[for="' + CSS.escape(e.id) + '"]');
      if (l) s += ' ' + l.innerText;
    }
    const lab = e.closest('label'); if (lab) s += ' ' + lab.innerText;
    let p = e.parentElement;
    for (let k = 0; k < 3 && p; k++, p = p.parentElement) {
      const t = p.innerText || ''; if (t.length < 90) s += ' ' + t;
    }
    return s.replace(/\s+/g, ' ');
  };
  const re = new RegExp(cfg.campo, 'i'), ex = new RegExp(cfg.excluir, 'i');
  const achados = ins.filter(e => { const l = rotulo(e); return re.test(l) && !ex.test(l); });
  const ini = achados.find(e => /inicial|in[ií]cio|\bde\b/i.test(rotulo(e))) || achados[0];
  const fim = achados.find(e => e !== ini && /final|\bfim\b|at[eé]/i.test(rotulo(e))) ||
              achados.find(e => e !== ini);
  document.querySelectorAll('[data-rpa-campo]').forEach(e => e.removeAttribute('data-rpa-campo'));
  if (ini) ini.setAttribute('data-rpa-campo', 'ini');
  if (fim) fim.setAttribute('data-rpa-campo', 'fim');
  return {ini: ini ? (ini.getAttribute('type') || 'text') : null,
          fim: fim ? (fim.getAttribute('type') || 'text') : null};
}
"""

JS_BLOB_B64 = r"""
async (u) => {
  const r = await fetch(u); const b = new Uint8Array(await r.arrayBuffer());
  let s = ''; for (let i = 0; i < b.length; i += 0x8000)
    s += String.fromCharCode.apply(null, b.subarray(i, i + 0x8000));
  return btoa(s);
}
"""


# --------------------------------------------------------------------------
# Robô
# --------------------------------------------------------------------------
class Robo:
    def __init__(self, ponte, inicio, fim, pasta, baixar_recibo=True,
                 baixar_declaracao=True, cfg=None):
        self.ui = ponte
        self.cfg = cfg or carregar_config()
        self.inicio, self.fim = inicio, fim
        self.pasta = Path(pasta)
        self.docs = []
        if baixar_recibo:
            self.docs.append(("recibo", "Recibo", "acao_recibo"))
        if baixar_declaracao:
            self.docs.append(("declaracao", "DCTF", "acao_declaracao"))
        self.resultados = []
        self.usados = set()
        self.encontradas = set()
        self.ctx = None
        self.page = None
        self.frame = None
        self.total = max(1, len(meses_do_periodo(inicio, fim)) * max(1, len(self.docs)))
        self.feitos = 0

    # ---------------- infraestrutura ----------------
    def log(self, msg, tipo="info"):
        self.ui.log(msg, tipo)

    def _checar(self):
        if self.ui.cancelar.is_set():
            raise Cancelado()

    def _paginas(self):
        return [p for p in self.ctx.pages if not p.is_closed()]

    def _pausa(self, ms=300):
        """Espera mantendo o navegador respondendo (o Playwright precisa disso)."""
        pags = self._paginas()
        if pags:
            try:
                pags[0].wait_for_timeout(ms)
                return
            except Exception:
                pass
        time.sleep(ms / 1000)

    def _esperar(self, func, segundos):
        fim = time.time() + segundos
        while time.time() < fim:
            self._checar()
            try:
                r = func()
                if r:
                    return r
            except Exception:
                pass
            self._pausa(500)
        return None

    def _perguntar(self, tipo, titulo, texto, botoes, vigia=None):
        """Mostra uma caixa na janela e espera a resposta, sem travar o navegador."""
        pedido = self.ui.pedir(tipo, titulo, texto, botoes)
        avisado = False
        while not pedido.respondido():
            self._checar()
            if vigia and not avisado:
                try:
                    if vigia():
                        avisado = True
                        self.ui.login_detectado()
                except Exception:
                    pass
            self._pausa(500)
        if pedido.resposta == "cancelar":
            raise Cancelado()
        return pedido.resposta

    def _assistido(self, titulo, texto):
        self.log(titulo + " — aguardando você no navegador.", "aviso")
        self._perguntar("assistido", titulo, texto, ["Continuar", "Cancelar"])

    def _abrir_navegador(self, pw):
        perfil = Path(os.environ.get("RPA_DCTF_PERFIL") or (pasta_de_dados() / "navegador"))
        perfil.mkdir(parents=True, exist_ok=True)
        ultimo = None
        for canal in self.cfg["navegadores"]:
            try:
                ctx = pw.chromium.launch_persistent_context(
                    str(perfil),
                    channel=None if canal == "chromium" else canal,
                    headless=bool(self.cfg.get("headless")),
                    accept_downloads=True,
                    no_viewport=True,
                    args=["--start-maximized"],
                )
                nome = {"msedge": "Microsoft Edge", "chrome": "Google Chrome"}.get(canal, canal)
                self.log(f"Navegador aberto: {nome}.")
                return ctx
            except Exception as e:
                ultimo = e
                if "already in use" in str(e) or "lock" in str(e).lower():
                    raise ErroRobo("O navegador do robô já está aberto. Feche a janela "
                                   "do navegador aberta pelo robô e tente de novo.")
        raise ErroRobo("Não encontrei o Microsoft Edge nem o Google Chrome neste "
                       f"computador. Detalhe técnico: {ultimo}")

    def _diagnostico(self, motivo):
        """Salva prints e o HTML das telas para ajustar o robô depois."""
        try:
            pasta = pasta_de_dados() / "diagnostico" / dt.datetime.now().strftime("%Y%m%d-%H%M%S")
            pasta.mkdir(parents=True, exist_ok=True)
            (pasta / "motivo.txt").write_text(motivo, encoding="utf-8")
            for i, p in enumerate(self._paginas()):
                try:
                    p.screenshot(path=str(pasta / f"aba{i}.png"))
                except Exception:
                    pass
                for j, fr in enumerate(p.frames):
                    try:
                        (pasta / f"aba{i}_quadro{j}.html").write_text(
                            f"<!-- {fr.url} -->\n" + fr.content(), encoding="utf-8")
                    except Exception:
                        pass
            self.log(f"Diagnóstico salvo em {pasta}", "aviso")
        except Exception:
            pass

    # ---------------- busca de elementos ----------------
    def _achar(self, alvo, padroes, excluir=None, ativos=False, raiz=None):
        """Procura um elemento clicável. alvo = frame ou locator de uma linha.
        raiz = seletor CSS opcional para limitar a busca (ex.: menu aberto)."""
        cfg = {"padroes": padroes, "excluir": excluir, "ativos": ativos, "raiz": raiz}
        try:
            if hasattr(alvo, "evaluate_all"):      # locator (linha)
                return alvo.evaluate(JS_ACHAR, cfg)
            return alvo.evaluate(f"(cfg) => ({JS_ACHAR})(document, cfg)", cfg)
        except Exception:
            return None

    def _achar_em_tudo(self, padroes, excluir=None, ativos=False):
        for p in reversed(self._paginas()):
            for fr in p.frames:
                r = self._achar(fr, padroes, excluir, ativos)
                if r:
                    return fr, r
        return None, None

    def _clicar_marcado(self, frame, marca, nova_aba_ms=1500):
        """Clica no elemento marcado e acompanha se abriu outra aba."""
        antes = set(self._paginas())
        frame.locator(f'[data-rpa-alvo="{marca["id"]}"]').first.click(timeout=10000)
        self._pausa(nova_aba_ms)
        novas = [p for p in self._paginas() if p not in antes]
        if novas:
            self.page = novas[-1]
            try:
                self.page.wait_for_load_state("domcontentloaded", timeout=30000)
            except Exception:
                pass

    @staticmethod
    def _clique_simples(frame, marca):
        """Clique sem esperar a página carregar (usado quando o clique gera um PDF)."""
        frame.locator(f'[data-rpa-alvo="{marca["id"]}"]').first.click(timeout=10000, no_wait_after=True)

    def _clicar_texto(self, padroes, abrir_antes=None, segundos=20):
        fim = time.time() + segundos
        tentou_menu = False
        while time.time() < fim:
            self._checar()
            fr, m = self._achar_em_tudo(padroes)
            if m:
                self._clicar_marcado(fr, m)
                return True
            if abrir_antes and not tentou_menu:
                fr, m = self._achar_em_tudo(abrir_antes)
                if m:
                    self._clicar_marcado(fr, m, 800)
                    tentou_menu = True
                    continue
            self._pausa(700)
        return False

    def _tem_texto(self, frame, padrao):
        try:
            loc = frame.get_by_text(re.compile(padrao, re.I))
            for i in range(min(loc.count(), 5)):
                if loc.nth(i).is_visible():
                    return True
        except Exception:
            pass
        return False

    def _tem_captcha(self):
        for p in self._paginas():
            for fr in p.frames:
                if re.search(self.cfg["captcha"], fr.url or "", re.I):
                    try:
                        if fr.frame_element().is_visible():
                            return True
                    except Exception:
                        return True
        return False

    def _vigiar_captcha(self):
        if self._tem_captcha():
            self._assistido("Verificação de segurança",
                            "O e-CAC pediu uma verificação (\"não sou robô\").\n\n"
                            "Resolva no navegador e clique em Continuar.")

    def _sessao_expirada(self):
        return any(re.search(self.cfg["regex_tela_login"], p.url or "", re.I)
                   for p in self._paginas()) and not self._logado()

    # ---------------- etapas ----------------
    def _logado(self):
        for p in self._paginas():
            u = p.url or ""
            if re.search(self.cfg["regex_logado"], u, re.I) and \
                    not re.search(self.cfg["regex_tela_login"], u, re.I):
                return p
        return None

    def _login(self):
        self.ui.status("Aguardando o seu login no e-CAC...")
        self.log("Abrindo a página de login do e-CAC...")
        self.page.goto(self.cfg["url_login"], wait_until="domcontentloaded", timeout=60000)
        self._perguntar(
            "login", "Faça login no e-CAC",
            "1. Na janela do navegador que abriu, entre com a sua conta gov.br "
            "(certificado digital ou CPF e senha).\n"
            "2. Se for procurador, troque o perfil para a empresa desejada.\n"
            "3. Com a página inicial do e-CAC aberta, clique em Continuar.\n\n"
            "O robô não vê nem guarda a sua senha.",
            ["Continuar", "Cancelar"], vigia=self._logado)
        p = self._logado()
        if p:
            self.page = p
            self.log("Login confirmado.", "ok")
        else:
            self.log("Não identifiquei a página inicial do e-CAC; seguindo assim mesmo.", "aviso")

    def _frame_dctfweb(self):
        for p in reversed(self._paginas()):
            for fr in p.frames:
                if any(self._tem_texto(fr, m) for m in self.cfg["marcas_tela_dctfweb"]):
                    return fr
        return None

    def _abrir_dctfweb(self):
        self.ui.status("Abrindo a DCTFWeb...")
        self.log("Abrindo a DCTFWeb...")
        fr = self._frame_dctfweb()
        if not fr:
            if self.cfg.get("url_dctfweb"):
                self.page.goto(self.cfg["url_dctfweb"], wait_until="domcontentloaded",
                               timeout=60000)
            else:
                self._clicar_texto(self.cfg["servico_dctfweb"],
                                   abrir_antes=self.cfg["menu_declaracoes"], segundos=25)
            fr = self._esperar(self._frame_dctfweb, 40)
        if not fr:
            self._vigiar_captcha()
            self._diagnostico("Tela da DCTFWeb não encontrada")
            self._assistido(
                "Abra a DCTFWeb",
                "Não consegui abrir a DCTFWeb sozinho.\n\nNo navegador, vá em "
                "Declarações e Demonstrativos > Assinar e Transmitir DCTFWeb. "
                "Quando a lista de declarações aparecer, clique em Continuar.")
            fr = self._esperar(self._frame_dctfweb, 30)
            if not fr:
                raise ErroRobo("A tela de consulta da DCTFWeb não foi encontrada.")
        self.frame = fr
        self.page = fr.page
        self.log("DCTFWeb aberta.", "ok")

    def _digitar_competencia(self, loc, tipo, ano, mes):
        valor = f"{mes:02d}/{ano}"
        loc.click()
        loc.press("Control+A")
        loc.press("Backspace")
        if tipo == "month":
            loc.fill(f"{ano}-{mes:02d}")
        elif tipo == "date":
            loc.fill(f"{ano}-{mes:02d}-01")
        else:
            loc.press_sequentially(f"{mes:02d}{ano}", delay=60)
            if re.sub(r"\D", "", loc.input_value()) != f"{mes:02d}{ano}":
                loc.fill(valor)
        loc.press("Tab")

    def _pesquisar(self):
        ini_a, ini_m = self.inicio
        fim_a, fim_m = self.fim
        self.ui.status("Pesquisando as declarações do período...")
        fr = self.frame
        campos = None
        try:
            campos = fr.evaluate(JS_CAMPOS_PERIODO, {"campo": self.cfg["campo_periodo"],
                                                    "excluir": self.cfg["campo_periodo_excluir"]})
        except Exception:
            pass
        feito = False
        if campos and campos.get("ini") and campos.get("fim"):
            try:
                self._digitar_competencia(fr.locator('[data-rpa-campo="ini"]').first,
                                          campos["ini"], ini_a, ini_m)
                self._digitar_competencia(fr.locator('[data-rpa-campo="fim"]').first,
                                          campos["fim"], fim_a, fim_m)
                m = self._achar(fr, self.cfg["botao_pesquisar"])
                if m:
                    self._clicar_marcado(fr, m, 800)
                    feito = True
            except Exception as e:
                self.log(f"Não consegui preencher o filtro sozinho ({e}).", "aviso")
        if not feito:
            self._diagnostico("Filtro de período não encontrado")
            self._assistido(
                "Pesquise o período",
                f"Na DCTFWeb, preencha o período de apuração de {ini_m:02d}/{ini_a} "
                f"a {fim_m:02d}/{fim_a}, clique em Pesquisar e, quando a lista "
                "aparecer, clique em Continuar.")
            self.frame = self._frame_dctfweb() or self.frame
        # espera a lista carregar
        self._pausa(1500)
        try:
            self.frame.page.wait_for_load_state("networkidle", timeout=10000)
        except Exception:
            pass
        self._esperar(lambda: self._linhas() or self._tem_texto(self.frame, self.cfg["sem_resultados"]), 40)
        self.log(f"Pesquisa de {ini_m:02d}/{ini_a} a {fim_m:02d}/{fim_a} feita.")

    def _interpretar(self, texto):
        m = RE_COMP.search(texto)
        categoria = ""
        for rx, nome in CATEGORIAS:
            if re.search(rx, texto, re.I):
                categoria = nome
                break
        if m:
            ano, mes = int(m.group(2)), int(m.group(1))
        else:
            if categoria not in ("13º Salário", "Anual"):
                return None
            limpo = re.sub(r"\d{2}/\d{2}/\d{4}(\s+\d{2}:\d{2}(:\d{2})?)?", " ", texto)
            a = RE_ANO.search(limpo)
            if not a:
                return None
            ano, mes = int(a.group(1)), 13
        situacao = "?"
        for s in SITUACOES:
            if re.search(r"\b" + re.escape(s) + r"\b", texto, re.I):
                situacao = s
                break
        return Linha(texto=texto, competencia=f"{mes:02d}/{ano}", ano=ano, mes=mes,
                     categoria=categoria, situacao=situacao)

    def _no_periodo(self, l):
        if l.mes == 13:
            return self.inicio[0] <= l.ano <= self.fim[0]
        return self.inicio <= (l.ano, l.mes) <= self.fim

    def _linhas(self):
        try:
            brutas = self.frame.evaluate(JS_LINHAS)
        except Exception:
            return []
        out = []
        for b in brutas:
            l = self._interpretar(b["t"])
            if l:
                out.append(l)
        return out

    def _localizar_linha(self, alvo):
        """Encontra de novo a linha na tela (a lista pode ter sido redesenhada)."""
        try:
            brutas = self.frame.evaluate(JS_LINHAS)
        except Exception:
            return None
        for b in brutas:
            if b["t"] == alvo.texto:
                return self.frame.locator(f'[data-rpa-linha="{b["i"]}"]').first
        return None

    def _nome_arquivo(self, prefixo, l):
        base = f"{prefixo} - {l.comp_arquivo}"
        nome = base
        if (prefixo, nome) in self.usados and l.categoria and l.categoria != "Geral":
            nome = f"{base} - {l.categoria}"
        n = 2
        while (prefixo, nome) in self.usados:
            nome = f"{base} ({n})"
            n += 1
        self.usados.add((prefixo, nome))
        return nome + ".pdf"

    # ---------------- captura do PDF ----------------
    def _capturar_pdf(self, clicar, destino):
        capt = {"download": None, "resp": [], "json": [], "paginas": [], "lidos": set()}

        def on_resp(r):
            try:
                h = r.headers
                ct = (h.get("content-type") or "").lower()
                cd = (h.get("content-disposition") or "").lower()
                url = r.url.lower().split("?")[0]
                if "pdf" in ct or ".pdf" in cd or url.endswith(".pdf"):
                    capt["resp"].append(r)
                elif r.request.resource_type in ("xhr", "fetch") and \
                        ("json" in ct or "text" in ct or "octet" in ct):
                    capt["json"].append(r)
            except Exception:
                pass

        def on_dl(d):
            capt["download"] = d

        def on_page(p):
            capt["paginas"].append(p)
            p.on("download", on_dl)

        self.ctx.on("response", on_resp)
        self.ctx.on("page", on_page)
        pags_antes = self._paginas()
        for p in pags_antes:
            p.on("download", on_dl)
        try:
            clicar()
            fim = time.time() + float(self.cfg.get("tempo_pdf", 45))
            while time.time() < fim:
                self._checar()
                if capt["download"]:
                    capt["download"].save_as(str(destino))
                    if self._pdf_valido(destino):
                        return True
                    capt["download"] = None
                for r in list(capt["resp"]):
                    if id(r) in capt["lidos"]:
                        continue
                    capt["lidos"].add(id(r))
                    try:
                        corpo = r.body()
                    except Exception:
                        continue
                    if corpo[:5] == b"%PDF-":
                        destino.write_bytes(corpo)
                        return True
                for r in list(capt["json"]):
                    if id(r) in capt["lidos"]:
                        continue
                    capt["lidos"].add(id(r))
                    try:
                        txt = r.text()
                    except Exception:
                        continue
                    m = re.search(r"JVBERi0[A-Za-z0-9+/=\\]{100,}", txt)
                    if m:
                        corpo = base64.b64decode(m.group(0).replace("\\/", "/").replace("\\", ""))
                        if corpo[:5] == b"%PDF-":
                            destino.write_bytes(corpo)
                            return True
                for p in list(capt["paginas"]):
                    if p.is_closed() or id(p) in capt["lidos"]:
                        continue
                    u = p.url or ""
                    try:
                        if u.startswith("blob:"):
                            capt["lidos"].add(id(p))
                            corpo = base64.b64decode(p.evaluate(JS_BLOB_B64, u))
                            if corpo[:5] == b"%PDF-":
                                destino.write_bytes(corpo)
                                return True
                        elif u.startswith("http"):
                            capt["lidos"].add(id(p))
                            resp = self.ctx.request.get(u, timeout=30000)
                            corpo = resp.body()
                            if corpo[:5] == b"%PDF-":
                                destino.write_bytes(corpo)
                                return True
                    except Exception:
                        pass
                self._pausa(300)
            return False
        finally:
            try:
                self.ctx.remove_listener("response", on_resp)
                self.ctx.remove_listener("page", on_page)
            except Exception:
                pass
            for p in self._paginas():
                try:
                    p.remove_listener("download", on_dl)
                except Exception:
                    pass
            for p in capt["paginas"]:
                try:
                    if not p.is_closed() and p is not self.frame.page:
                        p.close()
                except Exception:
                    pass

    @staticmethod
    def _pdf_valido(caminho):
        try:
            with open(caminho, "rb") as f:
                return f.read(5) == b"%PDF-"
        except Exception:
            return False

    # ---------------- download de uma linha ----------------
    def _executar_acao(self, linha, acao, destino):
        fr = self.frame
        excl = acao.get("excluir")
        marca = self._achar(linha, acao["linha"], excl)
        if not marca:
            menu = self._achar(linha, self.cfg["menu_da_linha"])
            if menu:
                self._clicar_marcado(fr, menu, 600)
                # primeiro dentro do menu que abriu; depois na tela toda
                marca = self._achar(fr, acao["linha"], excl, raiz=self.cfg["seletor_menu_aberto"]) \
                    or self._achar(fr, acao["linha"], excl)
        if not marca:
            return False, "botão não encontrado na linha"
        ok = self._capturar_pdf(lambda: self._clique_simples(fr, marca), destino)
        if ok:
            return True, ""
        # a ação pode ter aberto uma tela de detalhe com o botão do PDF
        fr2, m2 = self._achar_em_tudo(acao["tela"], excl, ativos=True)
        if m2:
            ok = self._capturar_pdf(lambda: self._clique_simples(fr2, m2), destino)
            if ok:
                return True, ""
        return False, "o PDF não foi gerado"

    def _lista_visivel(self):
        fr = self._frame_dctfweb()
        if fr:
            self.frame = fr
            if self._linhas():
                return True
        return False

    def _voltar_para_lista(self):
        if self._lista_visivel():
            return
        # fecha abas extras e volta na aba principal
        principal = self.frame.page if self.frame else self.page
        for p in self._paginas():
            if p is not principal and len(self._paginas()) > 1:
                try:
                    p.close()
                except Exception:
                    pass
        if self._lista_visivel():
            return
        fr, m = self._achar_em_tudo(self.cfg["botao_voltar"])
        if m:
            try:
                self._clicar_marcado(fr, m, 1000)
            except Exception:
                pass
            if self._esperar(self._lista_visivel, 10):
                return
        try:
            principal.go_back(wait_until="domcontentloaded", timeout=15000)
            if self._esperar(self._lista_visivel, 10):
                return
        except Exception:
            pass
        self.log("Refazendo a pesquisa para voltar à lista...", "aviso")
        self._abrir_dctfweb()
        self._pesquisar()

    def _baixar_linha(self, l):
        self.encontradas.add((l.ano, l.mes))
        for doc, prefixo, chave_acao in self.docs:
            self._checar()
            self._vigiar_captcha()
            nome = self._nome_arquivo(prefixo, l)
            destino = self.pasta / nome
            rotulo = "recibo" if doc == "recibo" else "declaração"
            self.ui.status(f"Baixando {rotulo} de {l.competencia}...")
            ok, obs = False, ""
            for tentativa in (1, 2):
                linha = self._localizar_linha(l)
                if linha is None:
                    self._voltar_para_lista()
                    linha = self._localizar_linha(l)
                if linha is None:
                    obs = "linha não encontrada na lista"
                    break
                try:
                    ok, obs = self._executar_acao(linha, self.cfg[chave_acao], destino)
                except Cancelado:
                    raise
                except Exception as e:
                    ok, obs = False, f"erro: {str(e).splitlines()[0][:150]}"
                self._voltar_para_lista()
                if ok:
                    break
                if self._sessao_expirada():
                    self._assistido("Sessão encerrada",
                                    "O e-CAC encerrou a sessão. Faça login de novo no "
                                    "navegador e clique em Continuar.")
                    self._abrir_dctfweb()
                    self._pesquisar()
            if ok:
                self.log(f"✓ {nome}", "ok")
            else:
                self.log(f"✗ {rotulo} {l.competencia}: {obs}", "erro")
                self._diagnostico(f"{rotulo} {l.competencia}: {obs}")
            self.resultados.append({
                "competencia": l.competencia, "documento": "Recibo de entrega" if doc == "recibo" else "Declaração",
                "arquivo": nome if ok else "", "categoria": l.categoria, "situacao": l.situacao,
                "status": "Baixado" if ok else "Falhou", "obs": obs})
            self.feitos += 1
            self.ui.progresso(min(self.feitos, self.total), self.total)

    def _proxima_pagina(self):
        m = self._achar(self.frame, self.cfg["proxima_pagina"], ativos=True)
        if not m:
            return False
        antes = [x.texto for x in self._linhas()]
        self._clicar_marcado(self.frame, m, 800)
        mudou = self._esperar(lambda: [x.texto for x in self._linhas()] != antes, 15)
        return bool(mudou)

    def _processar_lista(self):
        vistos = set()
        paginas = 1
        for _ in range(5000):
            self._checar()
            pendente = None
            for l in self._linhas():
                if l.chave in vistos:
                    continue
                aceita = l.situacao in self.cfg["situacoes_aceitas"] or l.situacao == "?"
                if not (self._no_periodo(l) and aceita):
                    vistos.add(l.chave)
                    continue
                pendente = l
                break
            if pendente:
                vistos.add(pendente.chave)
                self._baixar_linha(pendente)
                continue
            if paginas >= 200 or not self._proxima_pagina():
                break
            paginas += 1

    def _registrar_ausentes(self):
        for a, m in meses_do_periodo(self.inicio, self.fim):
            if (a, m) not in self.encontradas:
                self.resultados.append({
                    "competencia": f"{m:02d}/{a}", "documento": "-", "arquivo": "",
                    "categoria": "", "situacao": "", "status": "Sem declaração ativa",
                    "obs": "nenhuma declaração ativa encontrada na pesquisa"})

    def executar(self):
        from playwright.sync_api import sync_playwright
        self.pasta.mkdir(parents=True, exist_ok=True)
        with sync_playwright() as pw:
            self.ctx = self._abrir_navegador(pw)
            try:
                self.page = self._paginas()[0] if self._paginas() else self.ctx.new_page()
                self._login()
                self._abrir_dctfweb()
                self._pesquisar()
                self._processar_lista()
                self._registrar_ausentes()
                self.ui.progresso(self.total, self.total)
            finally:
                try:
                    self.ctx.close()
                except Exception:
                    pass
        return self.resultados


# --------------------------------------------------------------------------
# Relatório em Excel
# --------------------------------------------------------------------------
def gerar_relatorio(pasta, resultados, inicio, fim, situacao_final):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    wb = Workbook()
    ws = wb.active
    ws.title = "Downloads"
    navy = PatternFill("solid", fgColor=COR["navy"][1:])
    gold = Font(bold=True, color=COR["gold"][1:], size=14)
    branco = Font(bold=True, color="FFFFFF")
    fina = Side(style="thin", color="D0D7E1")

    ws["A1"] = f"{APP_NOME} — relatório de execução"
    ws["A1"].font = gold
    ws["A1"].fill = navy
    for col in "BCDEFG":
        ws[f"{col}1"].fill = navy
    ws["A2"] = f"Período: {inicio[1]:02d}/{inicio[0]} a {fim[1]:02d}/{fim[0]}"
    ws["A3"] = f"Gerado em: {dt.datetime.now():%d/%m/%Y %H:%M}   |   Situação: {situacao_final}"
    ws["A4"] = f"Pasta: {pasta}"

    cab = ["Competência", "Documento", "Arquivo", "Categoria", "Situação da declaração",
           "Status", "Observação"]
    ws.append([])
    ws.append(cab)
    lin_cab = ws.max_row
    for i, _ in enumerate(cab, start=1):
        c = ws.cell(row=lin_cab, column=i)
        c.fill = navy
        c.font = branco
        c.alignment = Alignment(vertical="center")
    def ordem(r):
        mm, aa = r["competencia"].split("/")
        doc = {"Recibo de entrega": 0, "Declaração": 1}.get(r["documento"], 2)
        return (int(aa), int(mm), doc)
    for r in sorted(resultados, key=ordem):
        ws.append([r["competencia"], r["documento"], r["arquivo"], r["categoria"],
                   r["situacao"], r["status"], r["obs"]])
        st = ws.cell(row=ws.max_row, column=6)
        st.font = Font(bold=True, color={"Baixado": "1E7B4F", "Falhou": "B5574F"}.get(r["status"], "8A6D1F"))
    for row in ws.iter_rows(min_row=lin_cab, max_row=ws.max_row, max_col=len(cab)):
        for c in row:
            c.border = Border(top=fina, bottom=fina, left=fina, right=fina)
    for col, larg in zip("ABCDEFG", [13, 19, 34, 22, 22, 20, 48]):
        ws.column_dimensions[col].width = larg
    ws.freeze_panes = ws.cell(row=lin_cab + 1, column=1)

    resumo = wb.create_sheet("Resumo")
    resumo.append(["Indicador", "Quantidade"])
    for c in resumo[1]:
        c.fill = navy
        c.font = branco
    ini_dados, fim_dados = lin_cab + 1, max(ws.max_row, lin_cab + 1)
    faixa_doc = f"Downloads!$B${ini_dados}:$B${fim_dados}"
    faixa_st = f"Downloads!$F${ini_dados}:$F${fim_dados}"
    resumo.append(["Recibos baixados", f'=COUNTIFS({faixa_doc},"Recibo de entrega",{faixa_st},"Baixado")'])
    resumo.append(["Declarações baixadas", f'=COUNTIFS({faixa_doc},"Declaração",{faixa_st},"Baixado")'])
    resumo.append(["Downloads com falha", f'=COUNTIF({faixa_st},"Falhou")'])
    resumo.append(["Competências sem declaração ativa", f'=COUNTIF({faixa_st},"Sem declaração ativa")'])
    resumo.column_dimensions["A"].width = 36
    resumo.column_dimensions["B"].width = 14

    nome = f"Relatório RPA DCTF - {dt.datetime.now():%Y-%m-%d %H%M}.xlsx"
    caminho = Path(pasta) / nome
    wb.save(caminho)
    return caminho


# --------------------------------------------------------------------------
# Janela
# --------------------------------------------------------------------------
def registrar_fontes():
    """Carrega a fonte Instrument Sans só para este programa (Windows)."""
    if not sys.platform.startswith("win"):
        return
    try:
        import ctypes
        FR_PRIVATE = 0x10
        for f in ("InstrumentSans-Regular.ttf", "InstrumentSans-Bold.ttf"):
            p = recurso("assets", f)
            if p.exists():
                ctypes.windll.gdi32.AddFontResourceExW(str(p), FR_PRIVATE, 0)
    except Exception:
        pass


def ativar_dpi():
    if sys.platform.startswith("win"):
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass


def main(iniciar=True):
    ativar_dpi()
    registrar_fontes()
    import tkinter as tk
    from tkinter import filedialog, font as tkfont, ttk

    class Botao(tk.Label):
        """Botão plano com cor de destaque (o tk.Button ignora cores em alguns sistemas)."""

        def __init__(self, pai, texto, comando, primario=True, **kw):
            self.primario = primario
            self.cores = (COR["gold"], COR["gold_hover"], COR["navy"]) if primario else \
                (COR["navy3"], COR["line"], COR["white"])
            super().__init__(pai, text=texto, bg=self.cores[0], fg=self.cores[2],
                             cursor="hand2", padx=16, pady=9, **kw)
            self.comando = comando
            self.ativo = True
            self.bind("<Enter>", lambda e: self.ativo and self.config(bg=self.cores[1]))
            self.bind("<Leave>", lambda e: self.config(bg=self.cores[0]))
            self.bind("<Button-1>", lambda e: self.ativo and self.comando())

        def estilo(self, primario, texto=None):
            self.cores = (COR["gold"], COR["gold_hover"], COR["navy"]) if primario else \
                (COR["navy3"], COR["line"], COR["white"])
            self.config(bg=self.cores[0], fg=self.cores[2])
            if texto:
                self.config(text=texto)

    class App:
        def __init__(self):
            self.root = tk.Tk()
            self.root.title(APP_NOME)
            self.root.configure(bg=COR["navy"])
            self.root.resizable(False, False)
            try:
                if sys.platform.startswith("win"):
                    self.root.iconbitmap(str(recurso("assets", "icone.ico")))
                else:
                    self._icone = tk.PhotoImage(file=str(recurso("assets", "icone.png")))
                    self.root.iconphoto(True, self._icone)
            except Exception:
                pass
            familias = set(tkfont.families())
            fam = next((f for f in ("Instrument Sans", "Segoe UI", "Helvetica") if f in familias),
                       "TkDefaultFont")
            self.f = {
                "titulo": (fam, 22, "bold"), "sub": (fam, 10), "rot": (fam, 9, "bold"),
                "campo": (fam, 11), "btn": (fam, 11, "bold"), "txt": (fam, 10),
                "log": (fam, 9), "rodape": (fam, 9), "dlg_t": (fam, 13, "bold"),
            }
            self.ponte = None
            self.trabalhando = False
            self.dialogo = None
            self._montar()
            self.root.after(120, self._processar_fila)

        # ---------- layout ----------
        def _montar(self):
            r = self.root
            topo = tk.Frame(r, bg=COR["navy"])
            topo.pack(fill="x", padx=28, pady=(26, 6))
            ic = tk.Canvas(topo, width=46, height=46, bg=COR["navy"], highlightthickness=0)
            ic.create_rectangle(1, 1, 45, 45, fill=COR["navy2"], outline=COR["line"])
            ic.create_polygon(23, 9, 37, 23, 23, 37, 9, 23, outline=COR["gold"], fill="", width=3)
            ic.pack(side="left", padx=(0, 14))
            txt = tk.Frame(topo, bg=COR["navy"])
            txt.pack(side="left", fill="x")
            tk.Label(txt, text=APP_NOME, font=self.f["titulo"], fg=COR["white"],
                     bg=COR["navy"]).pack(anchor="w")
            tk.Frame(txt, bg=COR["gold"], width=64, height=3).pack(anchor="w", pady=(2, 6))
            tk.Label(txt, text="Recibos de entrega e declarações da DCTFWeb, direto do e-CAC",
                     font=self.f["sub"], fg=COR["gray"], bg=COR["navy"]).pack(anchor="w")

            card = tk.Frame(r, bg=COR["card"], highlightbackground=COR["line"],
                            highlightthickness=1)
            card.pack(fill="x", padx=28, pady=(14, 10))
            corpo = tk.Frame(card, bg=COR["card"])
            corpo.pack(fill="x", padx=20, pady=18)

            hoje = dt.date.today()
            ant = (hoje.replace(day=1) - dt.timedelta(days=1))
            self.v_ini = tk.StringVar(value=f"01/{ant.year}")
            self.v_fim = tk.StringVar(value=f"{ant.month:02d}/{ant.year}")
            self.v_pasta = tk.StringVar(value="")
            self.v_rec = tk.BooleanVar(value=True)
            self.v_dec = tk.BooleanVar(value=True)

            linha = tk.Frame(corpo, bg=COR["card"])
            linha.pack(fill="x")
            for i, (rot, var) in enumerate((("COMPETÊNCIA INICIAL", self.v_ini),
                                            ("COMPETÊNCIA FINAL", self.v_fim))):
                col = tk.Frame(linha, bg=COR["card"])
                col.pack(side="left", fill="x", expand=True, padx=(0, 12) if i == 0 else 0)
                tk.Label(col, text=rot, font=self.f["rot"], fg=COR["gray2"],
                         bg=COR["card"]).pack(anchor="w")
                self._entrada(col, var).pack(fill="x", pady=(5, 0), ipady=6)
            tk.Label(corpo, text="Formato MM/AAAA. A DCTFWeb anual (13º salário) do período "
                                 "entra junto e é salva como 13-AAAA.", font=self.f["log"], fg=COR["gray2"],
                     bg=COR["card"]).pack(anchor="w", pady=(6, 14))

            tk.Label(corpo, text="PASTA DE DESTINO", font=self.f["rot"], fg=COR["gray2"],
                     bg=COR["card"]).pack(anchor="w")
            lp = tk.Frame(corpo, bg=COR["card"])
            lp.pack(fill="x", pady=(5, 14))
            e = self._entrada(lp, self.v_pasta)
            e.config(state="readonly", readonlybackground=COR["navy2"])
            e.pack(side="left", fill="x", expand=True, ipady=6)
            Botao(lp, "Escolher…", self._escolher_pasta, primario=False,
                  font=self.f["txt"]).pack(side="left", padx=(10, 0))

            tk.Label(corpo, text="DOCUMENTOS", font=self.f["rot"], fg=COR["gray2"],
                     bg=COR["card"]).pack(anchor="w")
            lc = tk.Frame(corpo, bg=COR["card"])
            lc.pack(fill="x", pady=(4, 16))
            for texto, var in (("Recibos de entrega", self.v_rec), ("Declarações (DCTF)", self.v_dec)):
                tk.Checkbutton(lc, text=texto, variable=var, font=self.f["txt"], fg=COR["white"],
                               bg=COR["card"], activebackground=COR["card"],
                               activeforeground=COR["white"], selectcolor=COR["navy3"],
                               highlightthickness=0, bd=0, cursor="hand2").pack(side="left", padx=(0, 18))

            self.btn = Botao(corpo, "Iniciar download", self._iniciar, font=self.f["btn"])
            self.btn.pack(fill="x")

            st = tk.Frame(r, bg=COR["navy"])
            st.pack(fill="both", expand=True, padx=28)
            self.l_status = tk.Label(st, text="Pronto para começar.", font=self.f["txt"],
                                     fg=COR["white"], bg=COR["navy"], anchor="w")
            self.l_status.pack(fill="x", pady=(4, 6))
            estilo = ttk.Style()
            try:
                estilo.theme_use("clam")
            except Exception:
                pass
            estilo.configure("Ouro.Horizontal.TProgressbar", troughcolor=COR["navy2"],
                             background=COR["gold"], bordercolor=COR["line"],
                             lightcolor=COR["gold"], darkcolor=COR["gold"], thickness=8)
            self.prog = ttk.Progressbar(st, style="Ouro.Horizontal.TProgressbar",
                                        mode="determinate", maximum=100)
            self.prog.pack(fill="x")
            caixa = tk.Frame(st, bg=COR["line"])
            caixa.pack(fill="both", expand=True, pady=(10, 0))
            self.t_log = tk.Text(caixa, height=8, width=64, bg=COR["navy2"], fg=COR["gray"],
                                 font=self.f["log"], relief="flat", bd=0, padx=10, pady=8,
                                 insertbackground=COR["gold"], wrap="word", state="disabled",
                                 highlightthickness=0)
            self.t_log.pack(fill="both", expand=True, padx=1, pady=1)
            for tag, cor in (("ok", COR["ok"]), ("erro", COR["erro"]), ("aviso", COR["gold"]),
                             ("info", COR["gray"])):
                self.t_log.tag_configure(tag, foreground=cor)

            rod = tk.Frame(r, bg=COR["navy"])
            rod.pack(fill="x", padx=28, pady=(14, 14))
            tk.Frame(rod, bg=COR["line"], height=1).pack(fill="x", pady=(0, 10))
            tk.Label(rod, text=f"criado por {AUTOR}", font=self.f["rodape"], fg=COR["gray2"],
                     bg=COR["navy"]).pack()

            r.update_idletasks()
            w, h = max(600, r.winfo_reqwidth()), r.winfo_reqheight()
            x = (r.winfo_screenwidth() - w) // 2
            y = max(20, (r.winfo_screenheight() - h) // 2 - 30)
            r.geometry(f"{w}x{h}+{x}+{y}")

        def _entrada(self, pai, var):
            return tk.Entry(pai, textvariable=var, font=self.f["campo"], bg=COR["navy2"],
                            fg=COR["white"], insertbackground=COR["gold"], relief="flat",
                            highlightthickness=1, highlightbackground=COR["line"],
                            highlightcolor=COR["gold"], disabledbackground=COR["navy2"])

        # ---------- ações ----------
        def _escolher_pasta(self):
            p = filedialog.askdirectory(title="Escolha a pasta onde os PDFs serão salvos",
                                        parent=self.root)
            if p:
                self.v_pasta.set(os.path.normpath(p))

        def _log(self, msg, tipo="info"):
            self.t_log.config(state="normal")
            self.t_log.insert("end", f"{dt.datetime.now():%H:%M:%S}  {msg}\n", tipo)
            self.t_log.see("end")
            self.t_log.config(state="disabled")

        def _iniciar(self):
            if self.trabalhando:
                if self.ponte:
                    self.ponte.cancelar.set()
                    self._log("Cancelando...", "aviso")
                    if self.dialogo:
                        self.dialogo.fechar("cancelar")
                return
            ini, fim = ler_competencia(self.v_ini.get()), ler_competencia(self.v_fim.get())
            if not ini or not fim:
                self._aviso("Competência inválida", "Use o formato MM/AAAA, por exemplo 01/2026.")
                return
            if ini > fim:
                self._aviso("Período invertido", "A competência inicial é maior que a final.")
                return
            if not (self.v_rec.get() or self.v_dec.get()):
                self._aviso("Nada para baixar", "Marque recibos, declarações ou os dois.")
                return
            if not self.v_pasta.get():
                self._escolher_pasta()
                if not self.v_pasta.get():
                    return
            self.trabalhando = True
            self.btn.estilo(False, "Cancelar")
            self.prog.config(value=0)
            self.ponte = Ponte()
            self.params = dict(inicio=ini, fim=fim, pasta=self.v_pasta.get(),
                               baixar_recibo=self.v_rec.get(), baixar_declaracao=self.v_dec.get())
            self._log(f"Início: {ini[1]:02d}/{ini[0]} a {fim[1]:02d}/{fim[0]}.", "aviso")
            threading.Thread(target=self._rodar, daemon=True).start()

        def _rodar(self):
            robo = None
            situacao, erro = "Concluído", None
            try:
                robo = Robo(self.ponte, **self.params)
                robo.executar()
            except Cancelado:
                situacao = "Cancelado"
            except ErroRobo as e:
                situacao, erro = "Interrompido", str(e)
            except Exception as e:
                situacao = "Interrompido"
                erro = f"{e}\n\n{traceback.format_exc(limit=3)}"
            rel = None
            res = robo.resultados if robo else []
            if res:
                try:
                    rel = gerar_relatorio(self.params["pasta"], res, self.params["inicio"],
                                          self.params["fim"], situacao)
                except Exception as e:
                    self.ponte.log(f"Não foi possível gravar o relatório: {e}", "erro")
            self.ponte.fila.put(("fim", situacao, res, rel, erro))

        def _processar_fila(self):
            try:
                while True:
                    item = self.ponte.fila.get_nowait() if self.ponte else None
                    if item is None:
                        break
                    tipo = item[0]
                    if tipo == "log":
                        self._log(item[1], item[2])
                    elif tipo == "status":
                        self.l_status.config(text=item[1])
                    elif tipo == "prog":
                        self.prog.config(maximum=item[2], value=item[1])
                    elif tipo == "pedido":
                        _, t, titulo, texto, botoes, pedido = item
                        self.dialogo = Dialogo(self, titulo, texto, botoes, pedido,
                                               login=(t == "login"))
                    elif tipo == "login_ok":
                        if self.dialogo:
                            self.dialogo.login_ok()
                    elif tipo == "fim":
                        self._fim(*item[1:])
            except queue.Empty:
                pass
            self.root.after(120, self._processar_fila)

        def _fim(self, situacao, res, rel, erro):
            self.trabalhando = False
            self.btn.estilo(True, "Iniciar download")
            if self.dialogo:
                self.dialogo.fechar("cancelar")
            rec = sum(1 for r in res if r["documento"] == "Recibo de entrega" and r["status"] == "Baixado")
            dec = sum(1 for r in res if r["documento"] == "Declaração" and r["status"] == "Baixado")
            falhas = sum(1 for r in res if r["status"] == "Falhou")
            sem = sum(1 for r in res if r["status"] == "Sem declaração ativa")
            self.l_status.config(text=f"{situacao}: {rec} recibo(s) e {dec} declaração(ões).")
            if erro:
                self._log(erro.splitlines()[0], "erro")
            if rel:
                self._log(f"Relatório: {rel.name}", "aviso")
            texto = f"{rec} recibo(s) e {dec} declaração(ões) baixados."
            if falhas:
                texto += f"\n{falhas} download(s) falharam — veja o relatório."
            if sem:
                texto += f"\n{sem} competência(s) sem declaração ativa."
            if erro:
                texto += f"\n\nMotivo: {erro.splitlines()[0]}"
            Dialogo(self, situacao, texto, ["Abrir pasta", "Fechar"], Pedido(),
                    ao_responder=lambda r: r == "Abrir pasta" and self._abrir_pasta())

        def _abrir_pasta(self):
            p = self.params.get("pasta")
            try:
                if sys.platform.startswith("win"):
                    os.startfile(p)
                elif sys.platform == "darwin":
                    os.system(f'open "{p}"')
                else:
                    os.system(f'xdg-open "{p}" >/dev/null 2>&1 &')
            except Exception:
                pass

        def _aviso(self, titulo, texto):
            Dialogo(self, titulo, texto, ["OK"], Pedido())

        def rodar(self):
            self.root.mainloop()

    class Dialogo:
        """Caixa de diálogo com a identidade visual."""

        def __init__(self, app, titulo, texto, botoes, pedido, login=False, ao_responder=None):
            self.app, self.pedido, self.ao_responder = app, pedido, ao_responder
            self.login, self.detectado, self.cliques = login, False, 0
            t = self.top = tk.Toplevel(app.root)
            t.title(titulo)
            t.configure(bg=COR["card"])
            t.resizable(False, False)
            t.transient(app.root)
            t.attributes("-topmost", True)
            t.protocol("WM_DELETE_WINDOW", lambda: self.fechar(botoes[-1].lower()))
            tk.Frame(t, bg=COR["gold"], height=4).pack(fill="x", side="top")
            corpo = tk.Frame(t, bg=COR["card"])
            corpo.pack(fill="both", padx=22, pady=18)
            tk.Label(corpo, text=titulo, font=app.f["dlg_t"], fg=COR["white"], bg=COR["card"],
                     anchor="w").pack(fill="x")
            tk.Label(corpo, text=texto, font=app.f["txt"], fg=COR["gray"], bg=COR["card"],
                     justify="left", wraplength=380, anchor="w").pack(fill="x", pady=(10, 12))
            self.l_info = tk.Label(corpo, text="Aguardando o login no navegador..." if login else "",
                                   font=app.f["log"], fg=COR["gold"], bg=COR["card"], anchor="w")
            if login:
                self.l_info.pack(fill="x", pady=(0, 10))
            lb = tk.Frame(corpo, bg=COR["card"])
            lb.pack(fill="x")
            for i, b in enumerate(botoes):
                Botao(lb, b, lambda b=b: self._clique(b), primario=(i == 0),
                      font=app.f["btn"]).pack(side="left", padx=(0, 10))
            t.update_idletasks()
            w, h = t.winfo_reqwidth(), t.winfo_reqheight()
            if login:   # canto superior direito, para não tapar o formulário de login
                x, y = t.winfo_screenwidth() - w - 40, 60
            else:
                x = app.root.winfo_rootx() + (app.root.winfo_width() - w) // 2
                y = app.root.winfo_rooty() + 120
            t.geometry(f"{w}x{h}+{max(0, x)}+{max(0, y)}")
            t.lift()
            t.focus_force()

        def login_ok(self):
            self.detectado = True
            self.l_info.config(text="✓ Login identificado. Clique em Continuar.", fg=COR["ok"])

        def _clique(self, botao):
            if self.login and botao == "Continuar" and not self.detectado and self.cliques == 0:
                self.cliques += 1
                self.l_info.config(text="Ainda não vi a página inicial do e-CAC. Termine o login "
                                        "e clique de novo (ou clique outra vez para seguir assim).")
                return
            self.fechar("cancelar" if botao == "Cancelar" else botao)

        def fechar(self, resposta):
            try:
                self.top.destroy()
            except Exception:
                pass
            if self.app.dialogo is self:
                self.app.dialogo = None
            if not self.pedido.respondido():
                self.pedido.responder(resposta)
            if self.ao_responder:
                self.ao_responder(resposta)

    app = App()
    if iniciar:
        app.rodar()
    return app


if __name__ == "__main__":
    main()
