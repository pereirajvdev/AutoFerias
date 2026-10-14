import argparse
import time
from pathlib import Path

import pyautogui
import pyperclip
import win32gui
from pywinauto import Desktop
from openpyxl import load_workbook


# ============================================================
# CONFIGURAÇÕES
# ============================================================

TITULO_MENU = "Menu: Arquivos / Cadastro de Funcionários"
TITULO_CADASTRO = "Cadastro de Férias"
TITULO_PESQUISA = "Pesquisando dados dos funcionários"

X_INICIAL = 583
Y_INICIAL = 841

X_CONSULTA = 594
Y_CONSULTA = 727

X_DATA = 983
Y_DATA = 494

INTERVALO = 0.3
INTERVALO_VERIFICACAO = 0.5


# ============================================================
# LISTAGEM E IDENTIFICAÇÃO DE JANELAS
# ============================================================

def listarJanelas():
    ignorar = {
        "lidDialog",
        "DataSetValidDialog",
        "Chart editor",
        "Editor de Cross-tab",
        "Dictionary editor",
        "Rich Edit",
        "Task Host Window",
        "CDA Server",
        "RadeonSoftware",
        "WUIconWindow",
        "OneDrive",
        "ad_msg_wnd",
        "Window",
        "Program Manager",
        "Default IME",
        "M",
    }

    titulos = []

    for janela in Desktop(backend="win32").windows():
        try:
            titulo = janela.window_text().strip()

            if not titulo:
                continue

            if titulo in ignorar:
                continue

            if "GDI+ Window" in titulo:
                continue

            if "cmd.exe" in titulo.lower():
                continue

            if "C:\\Windows" in titulo:
                continue

            if any(
                termo.lower() in titulo.lower()
                for termo in (
                    "AMD",
                    "Microsoft",
                    "Realtek",
                    "Bluetooth",
                    "HiddenFax",
                    ".NET-BroadcastEventWindow",
                )
            ):
                continue

            titulos.append(titulo)

        except Exception:
            pass

    return titulos


def encontrar_janela_por_titulo(titulo_procurado):
    """
    Retorna o identificador da primeira janela visível
    cujo título contenha o texto procurado.
    """
    hwnd_encontrado = None
    titulo_procurado = titulo_procurado.casefold()

    def callback(hwnd, _):
        nonlocal hwnd_encontrado

        try:
            titulo = win32gui.GetWindowText(hwnd).strip()

            if (
                win32gui.IsWindowVisible(hwnd)
                and titulo_procurado in titulo.casefold()
            ):
                hwnd_encontrado = hwnd
                return False

        except Exception:
            pass

        return True

    win32gui.EnumWindows(callback, None)
    return hwnd_encontrado


def menu_aberto():
    return encontrar_janela_por_titulo(TITULO_MENU) is not None


def cadastro_aberto():
    return encontrar_janela_por_titulo(TITULO_CADASTRO) is not None


def aguardar_estado_menu(timeout=30):
    """
    Aguarda até que o Menu esteja aberto e o Cadastro de Férias
    esteja fechado.
    """
    limite = time.time() + timeout

    while time.time() < limite:
        if menu_aberto() and not cadastro_aberto():
            return True

        time.sleep(INTERVALO_VERIFICACAO)

    return False


def aguardar_estado_cadastro(timeout=30):
    """
    Aguarda até que a janela Cadastro de Férias esteja aberta.
    """
    limite = time.time() + timeout

    while time.time() < limite:
        if cadastro_aberto():
            return True

        time.sleep(INTERVALO_VERIFICACAO)

    return False


# ============================================================
# ÁREA DE TRANSFERÊNCIA
# ============================================================

def ler_conteudo_copiado(timeout=3):
    limite = time.time() + timeout

    while time.time() < limite:
        conteudo = pyperclip.paste()

        if conteudo.strip():
            return conteudo.strip()

        time.sleep(0.1)

    return ""


# ============================================================
# PROCESSAMENTO
# ============================================================

def processar_pessoa(nome, numero_linha):
    print(f"\n{'=' * 50}")
    print(f"Linha {numero_linha}: {nome}")

    # ETAPA 1: estado inicial.
    print("Esperando Menu aberto e Cadastro de Férias fechado...")

    if not aguardar_estado_menu():
        print("Estado inicial não encontrado.")
        return None

    print(f"Janela encontrada: {TITULO_MENU}")

    if not ativar_janela(TITULO_MENU):
        return None

    # ETAPA 2: abrir pesquisa.
    print("Abrindo pesquisa...")
    pyautogui.click(X_INICIAL, Y_INICIAL)
    time.sleep(INTERVALO)

    # ETAPA 3: pesquisar o servidor.
    print(f"Pesquisando servidor: {nome}")

    pyperclip.copy(nome)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(INTERVALO)

    print("Pressionando primeiro Enter...")
    pyautogui.press("enter")
    time.sleep(INTERVALO)

    print("Pressionando segundo Enter...")
    pyautogui.press("enter")

    # ETAPA 4: confirmar o retorno ao Menu.
    print("Aguardando retorno ao Menu...")

    if not aguardar_estado_menu(timeout=15):
        print("O sistema não retornou ao Menu.")
        return None

    print(f"Janela encontrada: {TITULO_MENU}")

    # ETAPA 5: abrir Cadastro de Férias.
    if not ativar_janela(TITULO_MENU):
        return None

    print("Abrindo Cadastro de Férias...")
    pyautogui.click(X_CONSULTA, Y_CONSULTA)

    # Aguarda a janela correta aparecer.
    if not ativar_janela(TITULO_CADASTRO, timeout=30):
        print("Não foi possível abrir Cadastro de Férias.")
        return None

    # ETAPA 6: consultar e tentar copiar o resultado duas vezes.
    conteudo = consultar_resultado()

    if not conteudo:
        print(
            f"Nenhum conteúdo copiado após 2 tentativas para: {nome}. "
            "Prosseguindo para a próxima pessoa."
        )

        # Fecha o cadastro somente se ainda estiver aberto.
        if cadastro_aberto():
            print("Fechando Cadastro de Férias com Esc...")
            pyautogui.press("esc")

            aguardar_fechamento_janela(
                TITULO_CADASTRO,
                timeout=15,
            )

        return None

    # ETAPA 7: resultado copiado; fecha Cadastro de Férias.
    if not cadastro_aberto():
        print("Cadastro de Férias já está fechado.")
        return conteudo

    print("Fechando Cadastro de Férias com Esc...")
    pyautogui.press("esc")

    aguardar_fechamento_janela(
        TITULO_CADASTRO,
        timeout=15,
    )

    print(f"Processamento concluído: {nome}")
    return conteudo


def aguardar_janela_pesquisa(timeout=15):
    limite = time.time() + timeout

    while time.time() < limite:
        hwnd = encontrar_janela_por_titulo(TITULO_PESQUISA)

        if hwnd:
            print("Janela de pesquisa detectada.")
            return hwnd

        time.sleep(0.2)

    print("A janela de pesquisa não foi encontrada.")
    return None


def aguardar_janela(titulo, timeout=15):
    print(f"Esperando janela: {titulo}")

    limite = time.time() + timeout

    while time.time() < limite:
        hwnd = encontrar_janela_por_titulo(titulo)

        if hwnd:
            print(f"Janela encontrada: {win32gui.GetWindowText(hwnd)}")
            return hwnd

        time.sleep(0.2)

    print(f"Tempo esgotado. Janela não encontrada: {titulo}")
    return None


def aguardar_fechamento_janela(titulo, timeout=15):
    print(f"Esperando fechar a janela: {titulo}")

    limite = time.time() + timeout

    while time.time() < limite:
        hwnd = encontrar_janela_por_titulo(titulo)

        if not hwnd:
            print(f"Janela fechada: {titulo}")
            return True

        time.sleep(0.2)

    print(f"A janela continua aberta: {titulo}")
    return False


def ativar_janela(titulo, timeout=15):
    hwnd = aguardar_janela(titulo, timeout)

    if not hwnd:
        print(f"Janela não encontrada: {titulo}")
        return False

    try:
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(0.3)
        print(f"Janela ativa: {win32gui.GetWindowText(hwnd)}")
        return True
    except Exception as erro:
        print(f"Não foi possível ativar '{titulo}': {erro}")
        return False


def copiar_resultado():
    """
    Copia o conteúdo selecionado.
    Retorna uma string vazia se não conseguir copiar.
    """
    pyperclip.copy("")
    pyautogui.hotkey("ctrl", "c")

    conteudo = ler_conteudo_copiado(timeout=2)

    if conteudo:
        print(f"Conteúdo copiado: {conteudo}")
    else:
        print("Nenhum conteúdo foi copiado.")

    return conteudo


def consultar_resultado():
    """
    Executa a consulta e tenta copiar o resultado.
    Se a primeira tentativa falhar, repete o clique e a cópia
    uma única vez, desde que Cadastro de Férias continue aberto.
    """
    for tentativa in range(1, 3):
        if not cadastro_aberto():
            print("Consulta cancelada: Cadastro de Férias não está aberto.")
            return ""

        print(f"Tentativa de consulta {tentativa}/2...")

        pyautogui.click(X_DATA, Y_DATA)
        time.sleep(INTERVALO)

        if not cadastro_aberto():
            print("Cadastro de Férias foi fechado durante a consulta.")
            return ""

        conteudo = copiar_resultado()

        if conteudo:
            return conteudo

    return ""


# ============================================================
# PROGRAMA PRINCIPAL
# ============================================================

def main():
    parser = argparse.ArgumentParser(
        description="Consulta funcionários e registra resultados no Excel."
    )

    parser.add_argument(
        "planilha",
        type=Path,
        help="Caminho da planilha Excel.",
    )

    args = parser.parse_args()

    if not args.planilha.exists():
        print(f"Planilha não encontrada: {args.planilha}")
        return

    workbook = load_workbook(args.planilha)
    planilha = workbook.active

    try:
        pessoas = []

        for numero_linha in range(1, planilha.max_row + 1):
            valor = planilha.cell(numero_linha, 1).value

            if valor is not None and str(valor).strip():
                pessoas.append((numero_linha, str(valor).strip()))

        if not pessoas:
            print("Nenhum nome encontrado na coluna A.")
            return

        print(f"Total de pessoas: {len(pessoas)}")

        print("\nJanelas identificadas:")
        for titulo in listarJanelas():
            print(f"- {titulo}")

        for numero_linha, nome in pessoas:
            try:
                resultado = processar_pessoa(nome, numero_linha)

                if resultado is not None:
                    planilha.cell(numero_linha, 2).value = resultado
                    workbook.save(args.planilha)
                    print(f"Resultado salvo na coluna B, linha {numero_linha}.")
                else:
                    print(f"Linha {numero_linha} sem resultado. Seguindo adiante.")

            except KeyboardInterrupt:
                print("\nProcessamento interrompido pelo usuário.")
                break

            except Exception as erro:
                print(f"Erro na linha {numero_linha}: {erro}")

    finally:
        workbook.close()

    print("\nProcessamento finalizado.")


if __name__ == "__main__":
    main()