"""
Automação de Férias — versão refatorada e otimizada.

Uso:
    python main.py planilha.xlsx            # execução normal
    python main.py planilha.xlsx --sec      # Modo para levantamento por secretaria
    python main.py planilha.xlsx --perf     # com relatório de desempenho
    python main.py planilha.xlsx --verbose  # logs detalhados de janelas
    python main.py planilha.xlsx --fluxo-legado
        # mantém as DUAS fases originais (consulta e, depois, matrícula/status)

Principais ideias da refatoração (detalhes no relatório de entrega):
  * Localização de janelas por Windows API (ctypes) com cache de handle
    validado, em vez de Desktop(backend="win32").windows() a cada verificação.
  * Polling com intervalo progressivo (20 ms -> 200 ms) em vez de 1 s fixo.
  * Matrícula e status são lidos no Menu logo após a pesquisa, na mesma
    passagem da consulta (elimina a segunda pesquisa de todos os servidores).
    Quem ficar sem matrícula é reprocessado por um passe de recuperação.
  * Salvamento em lotes (a cada N alterações ou T segundos) e sempre no fim,
    de forma atômica (arquivo temporário + os.replace).
  * Profiling opcional (--perf) com custo praticamente nulo quando desligado.
"""

import argparse
import ctypes
import os
import re
import sys
import time
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from functools import lru_cache, wraps
from pathlib import Path

import pyautogui
import pyperclip
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# A importação do pywinauto também aplica a configuração de DPI que o
# código original já utilizava (coordenadas consistentes com o PyAutoGUI).
import pywinauto  # noqa: F401
from pywinauto.controls.hwndwrapper import HwndWrapper


# ============================================================
# CONFIGURAÇÕES
# ============================================================

TITULO_MENU = "Menu: Arquivos / Cadastro de Funcionários"
TITULO_CADASTRO = "Cadastro de Férias"
TITULO_PESQUISA = "Pesquisando dados dos funcionários"

ROTULOS_JANELAS = {
    TITULO_MENU: "Menu",
    TITULO_CADASTRO: "Cadastro",
    TITULO_PESQUISA: "Pesquisa",
}

# Coordenadas (preservadas do original).
X_INICIAL = 583
Y_INICIAL = 841

X_CONSULTA = 594
Y_CONSULTA = 727

X_DATA = 983
Y_DATA = 494

# Coordenadas da rotina de recuperação (preservadas).
X_RECUP_1, Y_RECUP_1 = 895, 360
X_RECUP_2, Y_RECUP_2 = 1300, 399

ROSA_CLARO = PatternFill(
    fill_type="solid",
    start_color="FFFFC0CB",
    end_color="FFFFC0CB",
)

# --- Esperas por janelas -------------------------------------------------
TEMPO_LIMITE = 60            # espera máxima por Menu/Cadastro
TIMEOUT_PESQUISA = 15        # espera máxima pela janela de pesquisa
TIMEOUT_FECHAMENTO = 3       # espera (não fatal) pelo fechamento de uma janela
POLL_INICIAL = 0.02          # primeiro intervalo de polling
POLL_MAXIMO = 0.20           # teto do intervalo de polling de janelas
POLL_FATOR = 1.6             # crescimento do intervalo a cada tentativa

# --- Esperas ligadas à interface ----------------------------------------
INTERVALO = 0.3              # usado apenas na rotina de recuperação (rara)
ESPERA_TECLAS = 0.1          # entre Ctrl+V -> Enter -> Enter na pesquisa
ESPERA_APOS_CLIQUE_DATA = 0.12
TIMEOUT_COPIA_RAPIDA = 0.8   # 1ª tentativa de Ctrl+C
TIMEOUT_COPIA_LENTA = 2.0    # 2ª tentativa (mesmo limite do original)
PAUSA_PYAUTOGUI = 0.02       # o padrão do PyAutoGUI é 0.1 s após CADA chamada

# --- Leitura do Menu -------------------------------------------------------
TENTATIVAS_LEITURA_MENU = 3
ESPERA_RELEITURA_MENU = 0.15

# --- Salvamento ---------------------------------------------------------------
LOTE_SALVAMENTO = 5          # salva a cada N pessoas alteradas...
INTERVALO_SALVAMENTO = 60.0  # ...ou após T segundos com alterações pendentes

VERBOSO = False


# ============================================================
# PROFILING (ativado somente com --perf)
# ============================================================

PERF_ATIVO = False
_TEMPOS = {}


def _registrar_tempo(nome, duracao):
    registro = _TEMPOS.setdefault(nome, [0, 0.0])
    registro[0] += 1
    registro[1] += duracao


def cronometrar(nome=None):
    """Decorator que mede o tempo total de uma função (só com --perf)."""

    def decorador(func):
        rotulo = nome or func.__name__

        @wraps(func)
        def wrapper(*args, **kwargs):
            if not PERF_ATIVO:
                return func(*args, **kwargs)

            inicio = time.perf_counter()

            try:
                return func(*args, **kwargs)
            finally:
                _registrar_tempo(rotulo, time.perf_counter() - inicio)

        return wrapper

    return decorador


class _Medidor:
    """Context manager de medição (usado só quando --perf está ativo)."""

    __slots__ = ("nome", "inicio")

    def __init__(self, nome):
        self.nome = nome
        self.inicio = 0.0

    def __enter__(self):
        self.inicio = time.perf_counter()

    def __exit__(self, *_):
        _registrar_tempo(self.nome, time.perf_counter() - self.inicio)
        return False


class _MedidorNulo:
    __slots__ = ()

    def __enter__(self):
        return None

    def __exit__(self, *_):
        return False


_MEDIDOR_NULO = _MedidorNulo()


def medir(nome):
    """with medir("etapa"): ... — sem custo relevante se --perf estiver off."""
    return _Medidor(nome) if PERF_ATIVO else _MEDIDOR_NULO


def imprimir_relatorio_tempos(tempo_total=None):
    if not PERF_ATIVO or not _TEMPOS:
        return

    print("\n" + "=" * 78)
    print(" RELATÓRIO DE DESEMPENHO (--perf)")
    print(" Obs.: tempos aninhados se sobrepõem; não some as linhas.")
    print("=" * 78)

    print(
        f"{'Etapa':<42}"
        f"{'Chamadas':>9}"
        f"{'Total (s)':>11}"
        f"{'Média (s)':>11}"
        f"{'% exec':>8}"
    )

    for nome, (chamadas, total) in sorted(
        _TEMPOS.items(),
        key=lambda item: item[1][1],
        reverse=True,
    ):
        media = total / chamadas if chamadas else 0.0
        pct = f"{100 * total / tempo_total:>5.1f}" if tempo_total else f"{'-':>5}"

        print(
            f"{nome:<42}"
            f"{chamadas:>9}"
            f"{total:>11.2f}"
            f"{media:>11.3f}"
            f"{pct:>8}"
        )

    print("=" * 78)


# ============================================================
# LOG
# ============================================================

def log(mensagem=""):
    print(mensagem)


def log_detalhe(mensagem):
    if VERBOSO:
        print(mensagem)


# ============================================================
# WINDOWS API (ctypes)
# ============================================================

class JanelaNaoEncontrada(RuntimeError):
    """A janela esperada não apareceu dentro do tempo limite."""


if sys.platform == "win32":
    from ctypes import wintypes

    _user32 = ctypes.WinDLL("user32", use_last_error=True)
    _ENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

    _user32.EnumWindows.argtypes = [_ENUMPROC, wintypes.LPARAM]
    _user32.EnumWindows.restype = wintypes.BOOL
    _user32.EnumChildWindows.argtypes = [wintypes.HWND, _ENUMPROC, wintypes.LPARAM]
    _user32.EnumChildWindows.restype = wintypes.BOOL
    _user32.IsWindow.argtypes = [wintypes.HWND]
    _user32.IsWindow.restype = wintypes.BOOL
    _user32.IsWindowVisible.argtypes = [wintypes.HWND]
    _user32.IsWindowVisible.restype = wintypes.BOOL
    _user32.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
    _user32.GetWindowTextW.restype = ctypes.c_int
    _user32.GetClassNameW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
    _user32.GetClassNameW.restype = ctypes.c_int
    _user32.GetWindowRect.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.RECT)]
    _user32.GetWindowRect.restype = wintypes.BOOL
    _user32.GetForegroundWindow.argtypes = []
    _user32.GetForegroundWindow.restype = wintypes.HWND
    _user32.SetForegroundWindow.argtypes = [wintypes.HWND]
    _user32.SetForegroundWindow.restype = wintypes.BOOL
    _user32.SendMessageTimeoutW.argtypes = [
        wintypes.HWND,
        wintypes.UINT,
        wintypes.WPARAM,
        ctypes.c_void_p,
        wintypes.UINT,
        wintypes.UINT,
        ctypes.POINTER(ctypes.c_size_t),
    ]
    _user32.SendMessageTimeoutW.restype = ctypes.c_ssize_t
else:  # permite importar/compilar fora do Windows (testes)
    _user32 = None

_WM_GETTEXT = 0x000D
_SMTO_ABORTIFHUNG = 0x0002

# Buffers reutilizados (a automação é sequencial, sem threads).
_TAM_BUF = 512
_BUF_TEXTO = ctypes.create_unicode_buffer(_TAM_BUF)
_BUF_CLASSE = ctypes.create_unicode_buffer(128)


def _titulo_janela(hwnd):
    tamanho = _user32.GetWindowTextW(hwnd, _BUF_TEXTO, _TAM_BUF)
    return _BUF_TEXTO.value.strip() if tamanho else ""


def _texto_controle(hwnd):
    """Texto de um controle (de qualquer processo) via WM_GETTEXT com timeout."""
    resultado = ctypes.c_size_t(0)
    _BUF_TEXTO.value = ""
    ok = _user32.SendMessageTimeoutW(
        hwnd,
        _WM_GETTEXT,
        _TAM_BUF,
        _BUF_TEXTO,
        _SMTO_ABORTIFHUNG,
        500,
        ctypes.byref(resultado),
    )
    return _BUF_TEXTO.value.strip() if ok else ""


def _enumerar_primeira_janela(trecho_casefold):
    """Primeira janela visível (ordem Z) cujo título contém o trecho."""
    achada = []

    def callback(hwnd, _):
        if not _user32.IsWindowVisible(hwnd):
            return True
        if _user32.GetWindowTextW(hwnd, _BUF_TEXTO, _TAM_BUF) and (
            trecho_casefold in _BUF_TEXTO.value.casefold()
        ):
            achada.append(hwnd)
            return False
        return True

    _user32.EnumWindows(_ENUMPROC(callback), 0)
    return achada[0] if achada else None


def _handle_ainda_valido(hwnd, trecho_casefold):
    """Confere se o handle em cache continua sendo a mesma janela."""
    if not _user32.IsWindow(hwnd) or not _user32.IsWindowVisible(hwnd):
        return False

    if not _user32.GetWindowTextW(hwnd, _BUF_TEXTO, _TAM_BUF):
        return False

    return trecho_casefold in _BUF_TEXTO.value.casefold()


_CACHE_HANDLES = {}      # título esperado -> hwnd
_CASEFOLD = {}           # título esperado -> título em casefold


def _trecho(titulo_esperado):
    trecho = _CASEFOLD.get(titulo_esperado)
    if trecho is None:
        trecho = _CASEFOLD[titulo_esperado] = titulo_esperado.casefold()
    return trecho


@cronometrar("localizar_janela")
def localizar_janela(titulo_esperado):
    """
    Localiza uma janela pelo título (ou trecho), sem diferenciar
    maiúsculas/minúsculas. Retorna o handle (int) ou None.

    O handle em cache só é reutilizado depois de validado (existe, está
    visível e o título ainda confere); caso contrário, enumera as janelas.
    """
    trecho = _trecho(titulo_esperado)

    hwnd = _CACHE_HANDLES.get(titulo_esperado)

    if hwnd is not None:
        if _handle_ainda_valido(hwnd, trecho):
            return hwnd
        del _CACHE_HANDLES[titulo_esperado]

    hwnd = _enumerar_primeira_janela(trecho)

    if hwnd is not None:
        _CACHE_HANDLES[titulo_esperado] = hwnd
        log_detalhe(f"[OK] Janela identificada: {_titulo_janela(hwnd)}")

    return hwnd


def _esperar_condicao(
    condicao,
    timeout,
    inicial=POLL_INICIAL,
    maximo=POLL_MAXIMO,
):
    """
    Polling com intervalo progressivo. A condição é verificada
    IMEDIATAMENTE antes de qualquer espera. Retorna o primeiro valor
    verdadeiro da condição ou None ao estourar o tempo.
    """
    limite = time.monotonic() + timeout
    intervalo = inicial

    while True:
        valor = condicao()

        if valor:
            return valor

        restante = limite - time.monotonic()

        if restante <= 0:
            return None

        time.sleep(min(intervalo, restante))
        intervalo = min(intervalo * POLL_FATOR, maximo)


@cronometrar("aguardar_janela")
def aguardar_janela(titulo_esperado, timeout=TEMPO_LIMITE):
    """
    Aguarda a janela aparecer.
    Retorna o handle ou None se atingir o limite.
    """
    rotulo = ROTULOS_JANELAS.get(titulo_esperado, titulo_esperado)

    with medir(f"espera janela: {rotulo}"):
        hwnd = localizar_janela(titulo_esperado)

        if hwnd is not None:
            return hwnd

        log_detalhe(f"[INFO] Aguardando a janela: {titulo_esperado}")

        hwnd = _esperar_condicao(
            lambda: localizar_janela(titulo_esperado),
            timeout,
        )

    if hwnd is None:
        log(
            f"[ERRO] A janela '{titulo_esperado}' "
            f"não apareceu em {timeout} segundos."
        )

    return hwnd


def exigir_janela(titulo_esperado, timeout=TEMPO_LIMITE):
    """Como aguardar_janela, mas levanta JanelaNaoEncontrada."""
    hwnd = aguardar_janela(titulo_esperado, timeout)

    if hwnd is None:
        raise JanelaNaoEncontrada(
            f"Janela '{titulo_esperado}' não encontrada em {timeout}s."
        )

    return hwnd


# Se uma janela nunca "some" para o nosso critério, esperar seu fechamento
# só desperdiçaria tempo: após o primeiro estouro, a espera é desativada.
_ESPERA_FECHAMENTO_CONFIAVEL = {}


def aguardar_janela_fechada(titulo_esperado, timeout=TIMEOUT_FECHAMENTO):
    """
    Espera (sem ser fatal) a janela desaparecer. Garante que a interface
    terminou a transição antes do próximo clique.
    """
    if not _ESPERA_FECHAMENTO_CONFIAVEL.get(titulo_esperado, True):
        return False

    rotulo = ROTULOS_JANELAS.get(titulo_esperado, titulo_esperado)

    with medir(f"espera fechar: {rotulo}"):
        fechada = _esperar_condicao(
            lambda: localizar_janela(titulo_esperado) is None,
            timeout,
        )

    if not fechada:
        _ESPERA_FECHAMENTO_CONFIAVEL[titulo_esperado] = False
        log(
            f"[AVISO] '{rotulo}' continuou aberta após {timeout}s; "
            "não esperarei mais pelo seu fechamento."
        )
        return False

    return True


def cadastro_aberto(hwnd):
    """True se o handle ainda é a janela de Cadastro de Férias."""
    if hwnd is None:
        return False

    return _handle_ainda_valido(hwnd, _trecho(TITULO_CADASTRO))


@cronometrar("focar_janela")
def focar_janela(hwnd):
    """Dá foco à janela (pula a chamada se ela já está em primeiro plano)."""
    if _user32.GetForegroundWindow() == hwnd:
        return

    try:
        HwndWrapper(hwnd).set_focus()
    except Exception:
        _user32.SetForegroundWindow(hwnd)


# ============================================================
# INTERAÇÃO COM A INTERFACE (PyAutoGUI)
# ============================================================

def clicar(x, y):
    with medir("ui: pyautogui"):
        pyautogui.click(x, y)


def pressionar(tecla):
    with medir("ui: pyautogui"):
        pyautogui.press(tecla)


def atalho(*teclas):
    with medir("ui: pyautogui"):
        pyautogui.hotkey(*teclas)


# ============================================================
# ÁREA DE TRANSFERÊNCIA
# ============================================================

def _colocar_na_area_de_transferencia(texto, tentativas=3):
    """pyperclip.copy com retentativa (a área pode estar bloqueada)."""
    for tentativa in range(tentativas):
        try:
            pyperclip.copy(texto)
            return
        except Exception:
            if tentativa == tentativas - 1:
                raise
            time.sleep(0.05)


def _ler_area_de_transferencia():
    try:
        conteudo = pyperclip.paste()
    except Exception:
        return ""

    return conteudo.strip() if conteudo else ""


@cronometrar("ler_conteudo_copiado")
def ler_conteudo_copiado(timeout=3):
    """Polling progressivo da área de transferência (timeout monotônico)."""
    return _esperar_condicao(
        _ler_area_de_transferencia,
        timeout,
        inicial=0.02,
        maximo=0.1,
    ) or ""


@cronometrar("copiar_resultado")
def copiar_resultado(timeout=2):
    """
    Copia o conteúdo selecionado. A área é limpa antes do Ctrl+C para nunca
    confundir a cópia nova com um resultado antigo.
    Retorna string vazia se não conseguir copiar.
    """
    _colocar_na_area_de_transferencia("")
    atalho("ctrl", "c")

    conteudo = ler_conteudo_copiado(timeout=timeout)

    if conteudo:
        log(f"Conteúdo copiado: {conteudo}")
    else:
        log("Nenhum conteúdo foi copiado.")

    return conteudo


@cronometrar("recuperar_copia")
def recuperar_copia(sessao, numero_linha):
    """
    Executa a recuperação e marca o resultado (coluna C) em rosa-claro.
    """
    log("Executando recuperação da cópia...")

    sessao.planilha.cell(numero_linha, 3).fill = ROSA_CLARO
    sessao.salvador.marcar()

    pressionar("esc")
    time.sleep(INTERVALO)

    clicar(X_RECUP_1, Y_RECUP_1)
    time.sleep(INTERVALO)

    clicar(X_RECUP_2, Y_RECUP_2)
    time.sleep(INTERVALO)

    log("Recuperação concluída. Continuando com Ctrl+C.")


@cronometrar("consultar_resultado")
def consultar_resultado(sessao, numero_linha):
    """
    Tenta copiar o resultado; se falhar, executa a recuperação e tenta
    novamente. A data copiada é convertida para o dia anterior.
    """
    log("Tentando copiar o resultado...")

    clicar(X_DATA, Y_DATA)
    clicar(X_DATA, Y_DATA)
    time.sleep(ESPERA_APOS_CLIQUE_DATA)

    # Duas tentativas de Ctrl+C antes de recorrer à recuperação. O tempo
    # total de espera (0,8 + 2,0 s) é maior que o do original (0,3 + 2,0 s),
    # então nenhuma falha é declarada mais cedo do que antes.
    conteudo = copiar_resultado(timeout=TIMEOUT_COPIA_RAPIDA)

    if not conteudo:
        conteudo = copiar_resultado(timeout=TIMEOUT_COPIA_LENTA)

    adicional = 0
    if not conteudo:
        log("Falha na primeira cópia. Iniciando recuperação...")

        recuperar_copia(sessao, numero_linha)

        hwnd_menu = exigir_janela(TITULO_MENU)
        focar_janela(hwnd_menu)

        log("Tentando Ctrl+C novamente...")
        conteudo = copiar_resultado(timeout=TIMEOUT_COPIA_LENTA)

        adicional = 1

    if not conteudo:
        log("Não foi possível copiar o resultado após a recuperação.")
        return ""

    try:
        data = datetime.strptime(conteudo, "%d/%m/%Y")
    except ValueError:
        log(f"O conteúdo copiado não é uma data válida: {conteudo!r}")
        return ""

    resultado = (data - timedelta(days=adicional)).strftime("%d/%m/%Y")

    log(f"Data ajustada: {resultado}")

    return resultado


# ============================================================
# LEITURA DO MENU (matrícula e status)
# ============================================================

def _escanear_menu(hwnd_menu):
    """
    Percorre os controles do Menu UMA vez (EnumChildWindows) e devolve
    (texto_da_matricula, texto_do_status), usando os mesmos critérios de
    classe e posição do original. Encerra assim que acha os dois.
    """
    achados = {"matricula": None, "status": None}
    rect = wintypes.RECT()

    def callback(hwnd, _):
        _user32.GetClassNameW(hwnd, _BUF_CLASSE, 128)
        classe = _BUF_CLASSE.value

        if classe == "TelDBKeyEdit" and achados["matricula"] is None:
            if _user32.GetWindowRect(hwnd, ctypes.byref(rect)) and (
                270 <= rect.top <= 305
            ):
                texto = _texto_controle(hwnd)

                if texto.isdecimal():
                    achados["matricula"] = texto

        elif classe == "TComboBox" and achados["status"] is None:
            if _user32.GetWindowRect(hwnd, ctypes.byref(rect)) and (
                854 <= rect.left <= 860 and 525 <= rect.top <= 535
            ):
                texto = _texto_controle(hwnd)

                if texto:
                    achados["status"] = texto

        return not (achados["matricula"] and achados["status"])

    _user32.EnumChildWindows(hwnd_menu, _ENUMPROC(callback), 0)

    return achados["matricula"], achados["status"]


def _diagnosticar_menu(hwnd_menu):
    """Lista os TelDBKeyEdit do Menu (só quando a matrícula não é achada)."""
    log("Controles TelDBKeyEdit encontrados:")
    rect = wintypes.RECT()

    def callback(hwnd, _):
        _user32.GetClassNameW(hwnd, _BUF_CLASSE, 128)

        if _BUF_CLASSE.value == "TelDBKeyEdit":
            _user32.GetWindowRect(hwnd, ctypes.byref(rect))
            log(
                f"Texto={_texto_controle(hwnd)!r} | "
                f"Posição=({rect.left}, {rect.top}) | "
                f"Tamanho=({rect.right - rect.left}x{rect.bottom - rect.top})"
            )

        return True

    _user32.EnumChildWindows(hwnd_menu, _ENUMPROC(callback), 0)


@cronometrar("ler_dados_menu")
def ler_dados_menu(diagnostico=False):
    """
    Lê matrícula (int) e status (str) do Menu numa única varredura.
    Substitui ler_matricula_menu() + ler_status_menu() do original, que
    percorriam todos os controles duas vezes.
    Retorna (matricula | None, status | None).
    """
    try:
        hwnd_menu = aguardar_janela(TITULO_MENU, timeout=3)

        if hwnd_menu is None:
            log("Janela do menu não encontrada.")
            return None, None

        texto_matricula, status = _escanear_menu(hwnd_menu)

        if texto_matricula is None:
            log("Campo da matrícula não identificado.")

            if diagnostico:
                _diagnosticar_menu(hwnd_menu)

            matricula = None
        else:
            matricula = int(texto_matricula)
            log(f"Matrícula identificada: {matricula}")

        if status:
            log(f"Status identificado: {status}")
        else:
            log("Campo de status não identificado.")

        return matricula, status

    except Exception as erro:
        log(f"Erro ao ler dados do menu: {erro}")
        return None, None


# ============================================================
# SALVAMENTO
# ============================================================

class Salvador:
    """
    Salva o workbook em lotes e de forma atômica:
      * só salva se houve alteração;
      * salva a cada LOTE_SALVAMENTO alterações ou INTERVALO_SALVAMENTO s;
      * grava num arquivo temporário e troca com os.replace (uma queda no
        meio do salvamento não corrompe a planilha original);
      * se o arquivo estiver aberto no Excel (PermissionError), avisa uma
        vez, mantém as alterações em memória e tenta de novo no próximo lote.
    """

    def __init__(self, workbook, caminho):
        self.workbook = workbook
        self.caminho = Path(caminho)
        self.pendentes = 0
        self.ultimo_salvamento = time.monotonic()
        self._avisou_bloqueio = False

    def marcar(self, quantidade=1):
        self.pendentes += quantidade

    def salvar_se_preciso(self):
        if not self.pendentes:
            return True

        decorrido = time.monotonic() - self.ultimo_salvamento

        if self.pendentes >= LOTE_SALVAMENTO or decorrido >= INTERVALO_SALVAMENTO:
            return self.salvar()

        return True

    @cronometrar("salvar_workbook")
    def salvar(self, forcar=False):
        if not self.pendentes and not forcar:
            return True

        temporario = self.caminho.with_name(
            f"{self.caminho.stem}.~tmp{self.caminho.suffix}"
        )

        try:
            self.workbook.save(temporario)
            os.replace(temporario, self.caminho)

        except (PermissionError, OSError) as erro:
            try:
                temporario.unlink(missing_ok=True)
            except OSError:
                pass

            # Evita insistir a cada pessoa: tenta de novo só no próximo lote.
            self.ultimo_salvamento = time.monotonic()

            if not self._avisou_bloqueio:
                self._avisou_bloqueio = True
                log(
                    f"\n[AVISO] Não foi possível salvar a planilha ({erro}). "
                    "Feche-a no Excel; as alterações continuam em memória "
                    "e novas tentativas serão feitas automaticamente."
                )

            return False

        self.pendentes = 0
        self.ultimo_salvamento = time.monotonic()
        self._avisou_bloqueio = False
        log_detalhe("[OK] Planilha salva.")
        return True

    def salvar_copia_de_seguranca(self):
        """Último recurso quando o arquivo original não pôde ser gravado."""
        carimbo = datetime.now().strftime("%Y%m%d_%H%M%S")
        destino = self.caminho.with_name(
            f"{self.caminho.stem}_recuperado_{carimbo}{self.caminho.suffix}"
        )
        self.workbook.save(destino)
        return destino


# ============================================================
# SESSÃO (estado compartilhado da execução)
# ============================================================

@dataclass
class Sessao:
    planilha: object
    salvador: Salvador
    ler_menu_junto: bool = True
    ultima_matricula: object = None
    matriculas: set = field(default_factory=set)   # linhas com matrícula lida
    status: dict = field(default_factory=dict)     # linha -> status do Menu


# ============================================================
# FLUXO DE PESQUISA / MATRÍCULA
# ============================================================

@cronometrar("pesquisar_servidor")
def pesquisar_servidor(nome, proximo=False):
    """
    Menu -> Pesquisa -> cola o nome -> Enter, Enter -> volta ao Menu.
    Ao retornar, o Menu está com foco e exibindo o servidor pesquisado.
    Levanta JanelaNaoEncontrada se alguma janela não aparecer.
    """
    hwnd = exigir_janela(TITULO_MENU)
    focar_janela(hwnd)

    log("Abrindo pesquisa...")
    clicar(X_INICIAL, Y_INICIAL)

    hwnd = exigir_janela(TITULO_PESQUISA, timeout=TIMEOUT_PESQUISA)
    focar_janela(hwnd)

    log(f"Pesquisando servidor: {nome}")

    _colocar_na_area_de_transferencia(nome)
    atalho("ctrl", "v")
    
    time.sleep(0.1)
    time.sleep(ESPERA_TECLAS)
    
    pressionar("enter")

    if proximo: pressionar("down")

    time.sleep(ESPERA_TECLAS)
    pressionar("enter")

    # Estado em vez de tempo fixo: o Menu só está pronto depois que a
    # janela de pesquisa sumiu (o Menu existe o tempo todo por trás dela).
    aguardar_janela_fechada(TITULO_PESQUISA)

    hwnd = exigir_janela(TITULO_MENU)
    focar_janela(hwnd)


def registrar_dados_menu(nome, numero_linha, sessao):
    """
    Lê matrícula/status no Menu e grava a matrícula na coluna A.
    O status é guardado para ser aplicado na coluna E depois de
    verificar_prazo() (que limpa essa coluna), exatamente como no original.
    Retorna True se obteve a matrícula.
    """
    matricula = status = None

    for tentativa in range(TENTATIVAS_LEITURA_MENU):
        ultima = tentativa == TENTATIVAS_LEITURA_MENU - 1

        matricula, status = ler_dados_menu(diagnostico=ultima)

        # Matrícula igual à do servidor anterior pode ser leitura de tela
        # ainda não atualizada: relê alguns instantes depois.
        if matricula is not None and matricula != sessao.ultima_matricula:
            break

        if not ultima:
            time.sleep(ESPERA_RELEITURA_MENU)

    if matricula is not None:
        sessao.ultima_matricula = matricula
        sessao.matriculas.add(numero_linha)
        sessao.planilha.cell(numero_linha, 1).value = matricula
        sessao.salvador.marcar()

        log(
            f"Matrícula {matricula} preenchida na coluna A, "
            f"linha {numero_linha}."
        )
    else:
        log(f"Não foi possível obter a matrícula de {nome}.")

    if status:
        sessao.status[numero_linha] = status

    return matricula is not None


@cronometrar("lerMatricula")
def lerMatricula(nome, numero_linha, sessao):
    """
    Passe de recuperação/legado: pesquisa o servidor e lê matrícula/status.
    Usado para quem ficou sem matrícula e com --fluxo-legado.
    """
    log(f"\nLendo matrícula de: {nome}")

    try:
        pesquisar_servidor(nome)
    except JanelaNaoEncontrada as erro:
        log(f"{erro}")
        return False

    return registrar_dados_menu(nome, numero_linha, sessao)


def fechar_cadastro(hwnd_cadastro, forcar=False):
    """Fecha o Cadastro com Esc (se ainda aberto) e espera ele sumir."""
    if forcar or cadastro_aberto(hwnd_cadastro):
        log("Fechando Cadastro de Férias com Esc...")
        pressionar("esc")

    aguardar_janela_fechada(TITULO_CADASTRO)


def get_status(nome, numero_linha, sessao):
    if sessao.ler_menu_junto:
        try:
            registrar_dados_menu(nome, numero_linha, sessao)
        except Exception as erro:
            log(f"Erro ao ler matrícula/status de {nome}: {erro}")

    return sessao.status.get(numero_linha, "")


@cronometrar("processar_pessoa")
def processar_pessoa(nome, numero_linha, sessao):
    log(f"\n{'=' * 50}")
    log(f"Linha {numero_linha}: {nome}")

    # ETAPAS 1-4: Menu -> Pesquisa -> pesquisar -> Menu.
    pesquisar_servidor(nome)

    # Matrícula e status já estão na tela: lê agora, evitando uma segunda
    # pesquisa deste mesmo servidor na fase seguinte.
    
    status = get_status(nome, numero_linha, sessao)

    if status.strip() == "03 - Demitido":     

        pesquisar_servidor(nome, True)
        status = get_status(nome, numero_linha, sessao)

        print(f"Deveria estar na {TITULO_MENU}")

        if status.strip() == "03 - Demitido":
            sessao.planilha.cell(numero_linha, 3).value = None
            sessao.planilha.cell(numero_linha, 4).value = None
            sessao.salvador.marcar()

            log(f"Servidor demitido: {nome}. Pulando processamento.")
            return None

    # ETAPA 5: abrir Cadastro de Férias.
    log("Abrindo Cadastro de Férias...")
    clicar(X_CONSULTA, Y_CONSULTA)

    hwnd_cadastro = exigir_janela(TITULO_CADASTRO)
    focar_janela(hwnd_cadastro)

    # ETAPA 6: consultar o resultado.
    conteudo = consultar_resultado(sessao, numero_linha)

    if not conteudo:
        log(
            f"Nenhum conteúdo copiado após as tentativas para: {nome}. "
            "Prosseguindo para a próxima pessoa."
        )
        fechar_cadastro(hwnd_cadastro, forcar=True)
        return None

    # ETAPA 7: fechar o cadastro (se a recuperação já o fechou, não há Esc).
    fechar_cadastro(hwnd_cadastro)

    hwnd_menu = exigir_janela(TITULO_MENU)
    focar_janela(hwnd_menu)

    log(f"Processamento concluído: {nome}")
    return conteudo


# ============================================================
# EXCEL: DATAS E PRAZOS
# ============================================================

_RE_BARRA = re.compile(r"\b\d{2}/\d{2}/\d{4}\b")
_RE_PONTO = re.compile(r"\b\d{2}\.\d{2}\.\d{4}\b")


@lru_cache(maxsize=None)
def _parse_data(texto, formato):
    """strptime com cache (None se inválida). Datas se repetem muito."""
    try:
        return datetime.strptime(texto, formato)
    except ValueError:
        return None


@cronometrar("extrair_periodo")
def extrair_periodo(valor):
    """
    Aceita uma data real do Excel ou texto com uma data.

    Retorna (data_inicial, data_final) ou None.
    A data inicial é o dia seguinte à data extraída (exceto 28/02 de ano
    bissexto, que vai para 01/03 — regra especial preservada).
    A data final é a mesma data no ano seguinte.
    """

    if valor is None or not str(valor).strip():
        return None

    if isinstance(valor, (datetime, date)):
        data_extraida = valor

    else:
        texto = str(valor).strip()

        if "/" in texto:
            formato = "%d/%m/%Y"
            correspondencia = _RE_BARRA.search(texto)
        else:
            formato = "%d.%m.%Y"
            correspondencia = _RE_PONTO.search(texto)

        if not correspondencia:
            log(f"Nenhuma data reconhecida: {texto}")
            return None

        data_extraida = _parse_data(correspondencia.group(0), formato)

        if data_extraida is None:
            log(f"Data inválida: {texto}")
            return None

    # Regra especial: ignorar 29 de fevereiro.
    if (
        data_extraida.month == 2
        and data_extraida.day == 28
        and data_extraida.year % 4 == 0
        and (
            data_extraida.year % 100 != 0
            or data_extraida.year % 400 == 0
        )
    ):
        data_inicial = data_extraida.replace(month=3, day=1)
    else:
        data_inicial = data_extraida + timedelta(days=1)

    try:
        data_final = data_extraida.replace(year=data_extraida.year + 1)
    except ValueError:
        # 29 de fevereiro em ano não bissexto.
        data_final = data_extraida.replace(
            year=data_extraida.year + 1,
            day=28,
        )

    return data_inicial, data_final


@cronometrar("interpretar_excel_final")
def interpretar_excel_final(planilha, data_limite=None):
    """
    Coluna C (resultado) -> coluna D (períodos) e limpa a coluna E.

    Quando data_limite é informada, acrescenta os períodos anuais
    completos até essa data, separados por '; '.
    """
    log("\nInterpretando resultados da coluna C...")

    interpretadas = 0

    for celula_c, celula_d, celula_e in planilha.iter_rows(
        min_col=3,
        max_col=5,
    ):
        resultado = celula_c.value

        if resultado is None or not str(resultado).strip():
            continue

        periodo = extrair_periodo(resultado)

        if not periodo:
            continue

        data_inicial, data_final = periodo
        periodos = []

        while data_limite is None or data_final.date() <= data_limite:
            periodos.append(
                f"{data_inicial:%d.%m.%Y} a {data_final:%d.%m.%Y}"
            )

            if data_limite is None:
                break

            data_inicial = data_final + timedelta(days=1)

            try:
                data_final = data_final.replace(
                    year=data_final.year + 1
                )
            except ValueError:
                data_final = data_final.replace(
                    year=data_final.year + 1,
                    day=28,
                )

        celula_d.value = "; ".join(periodos) if periodos else None
        celula_e.value = None

        if periodos:
            interpretadas += 1

        log_detalhe(f"Linha {celula_c.row}: {celula_d.value}")

    log(f"Interpretação do Excel concluída ({interpretadas} períodos).")


@cronometrar("verificar_prazo")
def verificar_prazo(planilha):
    """
    Coluna D: período aquisitivo interpretado.
    Coluna E: "Incompleto" quando o fim do período cai a partir do
    primeiro dia do próximo mês.
    """
    hoje = datetime.today()
    mes_atual = (hoje.year, hoje.month)

    incompletos = 0

    for _celula_c, celula_d, celula_e in planilha.iter_rows(min_col=3, max_col=5):
        celula_e.value = None

        periodo = celula_d.value

        if not isinstance(periodo, str):
            continue

        partes_periodos = periodo.split(";")
        ultimo_periodo = partes_periodos[-1].strip()

        partes = ultimo_periodo.split(" a ")

        if len(partes) != 2:
            continue

        data_final = _parse_data(partes[1].strip(), "%d.%m.%Y")

        if data_final is None:
            continue

        # data_final >= 1º dia do próximo mês  <=>  (ano, mês) > (ano, mês) atual
        if (data_final.year, data_final.month) > mes_atual:
            celula_e.value = "Incompleto"
            incompletos += 1
            log_detalhe(f"Linha {celula_d.row}: Incompleto")

    log(f"Verificação de prazos concluída ({incompletos} incompletos).")


def aplicar_status_menu(sessao):
    """Acrescenta à coluna E os status diferentes de 'Ativo' lidos no Menu."""
    for numero_linha, status in sessao.status.items():
        descricao = status.split("-", 1)[-1].strip()

        if descricao.casefold() == "ativo":
            continue

        celula = sessao.planilha.cell(numero_linha, 5)
        atual = str(celula.value or "").strip()

        celula.value = f"{atual} / {status}" if atual else status

        log(f"Status '{status}' registrado na coluna E, linha {numero_linha}.")


def coletar_pessoas(planilha):
    """Lista (linha, nome) a partir da coluna B, numa única passagem."""
    pessoas = []

    for (celula,) in planilha.iter_rows(min_col=2, max_col=2):
        valor = celula.value

        if valor is not None:
            nome = str(valor).strip()

            if nome:
                pessoas.append((celula.row, nome))

    return pessoas


# ============================================================
# ORQUESTRAÇÃO
# ============================================================

INTERRUPCOES = (KeyboardInterrupt, pyautogui.FailSafeException)


def fase_consulta(sessao, pessoas):
    """Percorre os funcionários. Retorna True se foi interrompida."""
    planilha = sessao.planilha

    for numero_linha, nome in pessoas:
        try:
            resultado = processar_pessoa(nome, numero_linha, sessao)

            if resultado is not None:
                planilha.cell(numero_linha, 3).value = resultado
                sessao.salvador.marcar()
                log(f"Resultado registrado na coluna C, linha {numero_linha}.")
            else:
                log(f"Linha {numero_linha} sem resultado. Seguindo adiante.")

        except INTERRUPCOES:
            log("\nProcessamento interrompido pelo usuário.")
            return True

        except Exception as erro:
            log(f"Erro na linha {numero_linha}: {erro}")

        finally:
            sessao.salvador.salvar_se_preciso()

    return False


def fase_recuperacao_matriculas(sessao, pessoas):
    """Reprocessa só quem ficou sem matrícula. Retorna True se interrompida."""
    for numero_linha, nome in pessoas:
        try:
            lerMatricula(nome, numero_linha, sessao)

        except INTERRUPCOES:
            log("\nProcessamento interrompido pelo usuário.")
            return True

        except Exception as erro:
            log(f"Erro na linha {numero_linha}: {erro}")

        finally:
            sessao.salvador.salvar_se_preciso()

    return False


def finalizar(sessao, pessoas, interrompido, data_limite=None):
    """Interpreta datas, completa matrículas, aplica status e salva."""
    salvador = sessao.salvador
    planilha = sessao.planilha

    try:
        interpretar_excel_final(planilha, data_limite)
        verificar_prazo(planilha)
        salvador.marcar()

        # Quem não teve a matrícula lida no 1º passe (erro no meio do fluxo
        # ou --fluxo-legado) é reprocessado. Não ocorre após interrupção.
        faltantes = [
            (linha, nome)
            for linha, nome in pessoas
            if linha not in sessao.matriculas
        ]

        if faltantes and not interrompido:
            log(
                f"\nLendo matrícula/status de {len(faltantes)} "
                "servidor(es) pendente(s)..."
            )
            salvador.salvar()  # checkpoint antes de voltar à interface
            fase_recuperacao_matriculas(sessao, faltantes)

        aplicar_status_menu(sessao)
        salvador.marcar()

        if salvador.salvar(forcar=True):
            log("Planilha salva com sucesso.")
        else:
            try:
                copia = salvador.salvar_copia_de_seguranca()
                log(
                    "\nERRO: Não foi possível salvar a planilha original.\n"
                    "Feche o arquivo no Excel e execute novamente.\n"
                    f"Seus dados foram preservados em: {copia}"
                )
            except Exception as erro:
                log(f"\nERRO: Falha também ao gravar a cópia de segurança: {erro}")

    except Exception as erro:
        log(f"Erro durante a finalização: {erro}")


def executar(args):
    if not args.planilha.exists():
        log(f"Planilha não encontrada: {args.planilha}")
        return

    pyautogui.PAUSE = PAUSA_PYAUTOGUI

    with medir("carregar_workbook"):
        workbook = load_workbook(args.planilha)

    planilha = workbook.active
    salvador = Salvador(workbook, args.planilha)
    sessao = Sessao(
        planilha=planilha,
        salvador=salvador,
        ler_menu_junto=not args.fluxo_legado,
    )

    pessoas = []
    interrompido = False

    try:
        pessoas = coletar_pessoas(planilha)

        if not pessoas:
            log("Nenhum nome encontrado na coluna B.")
        else:
            log(f"Total de pessoas: {len(pessoas)}")
            interrompido = fase_consulta(sessao, pessoas)

    except INTERRUPCOES:
        log("\nProcessamento interrompido pelo usuário.")
        interrompido = True

    finally:
        try:
            finalizar(sessao, pessoas, interrompido, args.sec)
        finally:
            workbook.close()

    log("\nProcessamento finalizado.")


def main():
    global PERF_ATIVO, VERBOSO

    parser = argparse.ArgumentParser(
        description="Consulta funcionários e registra resultados no Excel."
    )

    parser.add_argument(
        "planilha",
        type=Path,
        help="Caminho da planilha Excel.",
    )
    parser.add_argument(
        "--sec",
        type=lambda valor: datetime.strptime(valor, "%d/%m/%Y").date(),
        metavar="DD/MM/AAAA",
        help="Gera os períodos completos até a data informada.",
    )
    parser.add_argument(
        "--perf",
        action="store_true",
        help="Exibe o relatório de desempenho ao final.",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Exibe logs detalhados (janelas, linhas interpretadas).",
    )
    parser.add_argument(
        "--fluxo-legado",
        action="store_true",
        help=(
            "Mantém as duas fases originais: consulta de todos e, depois, "
            "nova pesquisa de todos para ler matrícula/status."
        ),
    )

    args = parser.parse_args()

    PERF_ATIVO = args.perf
    VERBOSO = args.verbose

    # Medido aqui (e não por decorator) porque --perf só é conhecido agora.
    with medir("main"):
        executar(args)


def _ponto_de_entrada():
    inicio_execucao = time.perf_counter()

    try:
        main()
    except KeyboardInterrupt:
        print("\nProcessamento interrompido pelo usuário.")
    finally:
        if PERF_ATIVO:
            tempo_total = time.perf_counter() - inicio_execucao

            print("\n" + "=" * 60)
            print(f" TEMPO TOTAL DE EXECUÇÃO: {tempo_total:.2f} segundos")
            print("=" * 60)

            imprimir_relatorio_tempos(tempo_total)


if __name__ == "__main__":
    _ponto_de_entrada()