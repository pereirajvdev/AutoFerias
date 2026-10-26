import argparse
import time
from pathlib import Path

import re

from openpyxl.styles import PatternFill

import pyautogui
import pyperclip
import win32gui
from pywinauto import Desktop
from openpyxl import load_workbook

from datetime import date, datetime, timedelta


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

ROSA_CLARO = PatternFill(
    fill_type="solid",
    start_color="FFFFC0CB",
    end_color="FFFFC0CB",
)

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

def processar_pessoa(nome, numero_linha, planilha):
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
    conteudo = consultar_resultado(planilha, numero_linha)

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


def consultar_resultado(planilha, numero_linha):
    """
    Tenta copiar o resultado.
    Se falhar, executa a recuperação e tenta copiar novamente.
    """

    if not cadastro_aberto():
        print("Consulta cancelada: Cadastro de Férias não está aberto.")
        return ""

    print("Tentando copiar o resultado...")

    pyautogui.click(X_DATA, Y_DATA)
    time.sleep(INTERVALO)

    if cadastro_aberto():
        conteudo = copiar_resultado()

        if conteudo:
            return conteudo

    print("Falha na primeira cópia. Iniciando recuperação...")

    recuperar_copia(planilha, numero_linha)

    # A recuperação já retornou a tela ao Menu.
    # Não pressionar Esc novamente.
    print("Tentando Ctrl+C novamente...")
    conteudo = copiar_resultado()

    if conteudo:
        # Conteúdo retirado da data de admissão: 
        # deve ser retirado 01 dia e entregar o dia anterior
        data = datetime.strptime(conteudo.strip(), "%d/%m/%Y")
        data_anterior = data - timedelta(days=1)
        return data_anterior.strftime("%d/%m/%Y")

    print("Não foi possível copiar o resultado após a recuperação.")
    return ""


def recuperar_copia(planilha, numero_linha):
    """
    Executa a recuperação e marca o resultado na coluna B em rosa-claro.
    """

    print("Executando recuperação da cópia...")

    # Marca o resultado da pessoa na coluna B.
    planilha.cell(numero_linha, 3).fill = ROSA_CLARO

    pyautogui.press("esc")
    time.sleep(INTERVALO)

    pyautogui.click(895, 360)
    time.sleep(INTERVALO)

    pyautogui.click(1300, 399)
    time.sleep(INTERVALO)

    print("Recuperação concluída. Continuando com Ctrl+C.")


def extrair_periodo(valor):
    """
    Aceita uma data real do Excel ou texto com uma data.

    Retorna:
        (data_inicial, data_final)

    A data inicial é o dia seguinte à data extraída.
    A data final é a mesma data no ano seguinte.
    """

    if valor is None or not str(valor).strip():
        return None

    # Caso seja uma data real do Excel.
    if isinstance(valor, (datetime, date)):
        data_extraida = valor

    else:
        texto = str(valor).strip()

        # Identifica o formato pela presença de "/".
        if "/" in texto:
            formato = "%d/%m/%Y"
            padrao = r"\b\d{2}/\d{2}/\d{4}\b"
        else:
            formato = "%d.%m.%Y"
            padrao = r"\b\d{2}\.\d{2}\.\d{4}\b"

        correspondencia = re.search(padrao, texto)

        if not correspondencia:
            print(f"Nenhuma data reconhecida: {texto}")
            return None

        try:
            data_extraida = datetime.strptime(
                correspondencia.group(0),
                formato,
            )
        except ValueError:
            print(f"Data inválida: {texto}")
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
        data_final = data_extraida.replace(
            year=data_extraida.year + 1,
        )
    except ValueError:
        # Trata 29 de fevereiro em ano não bissexto.
        data_final = data_extraida.replace(
            year=data_extraida.year + 1,
            day=28,
        )

    return data_inicial, data_final


def interpretar_excel_final(planilha):
    print("\nInterpretando resultados da coluna C...")

    for linha in range(1, planilha.max_row + 1):
        resultado = planilha.cell(linha, 3).value

        if resultado is None or not str(resultado).strip():
            continue

        periodo = extrair_periodo(resultado)

        if periodo:
            data_inicial, data_final = periodo

            planilha.cell(linha, 4).value = (
                f"{data_inicial:%d.%m.%Y} a {data_final:%d.%m.%Y}"
            )

            # Coluna E permanece vazia.
            planilha.cell(linha, 5).value = None

            print(
                f"Linha {linha}: "
                f"{planilha.cell(linha, 4).value}"
            )

    print("Interpretação do Excel concluída.")


def verificar_prazo(planilha):
    """
    Coluna D: período aquisitivo interpretado.
    Coluna F: indicação de período incompleto.
    """

    hoje = datetime.today()

    if hoje.month == 12:
        inicio_proximo_mes = datetime(hoje.year + 1, 1, 1)
    else:
        inicio_proximo_mes = datetime(hoje.year, hoje.month + 1, 1)

    for linha in range(1, planilha.max_row + 1):
        periodo = planilha.cell(linha, 4).value
        celula_resultado = planilha.cell(linha, 5)

        celula_resultado.value = None

        if not isinstance(periodo, str):
            continue

        partes = periodo.split(" a ")

        if len(partes) != 2:
            continue

        try:
            data_final = datetime.strptime(
                partes[1].strip(),
                "%d.%m.%Y",
            )
        except ValueError:
            continue

        if data_final >= inicio_proximo_mes:
            celula_resultado.value = "Incompleto"
            print(f"Linha {linha}: Incompleto")

    print("Verificação de prazos concluída.")


def ler_matricula_menu():
    """
    Localiza a matrícula usando o mesmo mecanismo de lerControles().
    """
    try:
        janela = Desktop(backend="win32").window(
            title=TITULO_MENU
        )

        if not janela.exists(timeout=3):
            print("Janela do menu não encontrada.")
            return None

        controles = janela.descendants()

        for controle in controles:
            try:
                classe = controle.class_name()
                texto = controle.window_text().strip()
                rect = controle.rectangle()

                if (
                    classe == "TelDBKeyEdit"
                    and 270 <= rect.top <= 305
                    and texto.isdigit()
                ):
                    matricula = int(texto)
                    print(f"Matrícula identificada: {matricula}")
                    return matricula

            except Exception:
                continue

        print("Campo da matrícula não identificado.")

        # Diagnóstico: exibe os controles candidatos encontrados.
        print("Controles TelDBKeyEdit encontrados:")

        for controle in controles:
            try:
                if controle.class_name() == "TelDBKeyEdit":
                    rect = controle.rectangle()
                    print(
                        f"Texto={controle.window_text()!r} | "
                        f"Posição=({rect.left}, {rect.top}) | "
                        f"Tamanho=({rect.width()}x{rect.height()})"
                    )
            except Exception:
                continue

        return None

    except Exception as erro:
        print(f"Erro ao obter matrícula: {erro}")
        return None


def obter_janela_menu():
    """
    Obtém a janela do menu e verifica se ela continua válida.
    """
    try:
        hwnd = encontrar_janela_por_titulo(TITULO_MENU)

        if not hwnd or not win32gui.IsWindow(hwnd):
            return None

        janela = Desktop(backend="win32").window(handle=hwnd)

        if not janela.exists(timeout=1):
            return None

        return janela

    except Exception as erro:
        print(f"Erro ao localizar a janela do menu: {erro}")
        return None


def lerMatricula(nome, numero_linha, planilha):
    print(f"\nLendo matrícula de: {nome}")

    if not aguardar_estado_menu():
        print("Menu não está disponível.")
        return False

    if not ativar_janela(TITULO_MENU):
        return False

    print("Abrindo pesquisa...")
    pyautogui.click(X_INICIAL, Y_INICIAL)
    time.sleep(INTERVALO)

    print(f"Pesquisando servidor: {nome}")
    pyperclip.copy(nome)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(INTERVALO)

    pyautogui.press("enter")
    time.sleep(INTERVALO)
    pyautogui.press("enter")

    if not aguardar_estado_menu(timeout=15):
        print("O sistema não retornou ao menu.")
        return False

    # Lê a matrícula do servidor selecionado.
    matricula = ler_matricula_menu()

    if matricula is not None:
        planilha.cell(numero_linha, 1).value = matricula

        print(
            f"Matrícula {matricula} preenchida na coluna A, "
            f"linha {numero_linha}."
        )
    else:
        print(f"Não foi possível obter a matrícula de {nome}.")

        # Lê o status do servidor selecionado.
    status = ler_status_menu()

    if status:
        # Considera ativo somente o status cujo texto termina em "Ativo".
        descricao_status = status.split("-", 1)[-1].strip()

        if descricao_status.casefold() != "ativo":
            celula_status = planilha.cell(numero_linha, 5)
            status_atual = str(celula_status.value or "").strip()

            if status_atual:
                celula_status.value = f"{status_atual} / {status}"
            else:
                celula_status.value = status

            print(
                f"Status '{status}' registrado na coluna E, "
                f"linha {numero_linha}."
            )

    return matricula is not None


def ler_status_menu():
    """
    Lê o status do servidor no controle TComboBox.
    Retorna o texto completo, por exemplo: '00 - Ativo'
    ou '01 - Licença'.
    """
    try:
        janela = Desktop(backend="win32").window(
            title=TITULO_MENU
        )

        if not janela.exists(timeout=3):
            print("Janela do menu não encontrada.")
            return None

        for controle in janela.descendants():
            try:
                if controle.class_name() != "TComboBox":
                    continue

                rect = controle.rectangle()

                # Identifica o campo de status pela posição.
                if (
                    854 <= rect.left <= 860
                    and 525 <= rect.top <= 535
                ):
                    status = controle.window_text().strip()

                    if status:
                        print(f"Status identificado: {status}")
                        return status

            except Exception:
                continue

        print("Campo de status não identificado.")
        return None

    except Exception as erro:
        print(f"Erro ao obter status: {erro}")
        return None


def localizar_janela(titulo_esperado):
    """
    Localiza uma janela pelo título ou por um trecho dele.
    Retorna o objeto da janela ou None se não encontrar.
    """

    for janela in Desktop(backend="win32").windows():
        try:
            titulo = janela.window_text().strip()

            if titulo_esperado.lower() in titulo.lower():
                print(f"[OK] Janela identificada: {titulo}")
                return janela

        except Exception:
            continue

    return None


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
            valor = planilha.cell(numero_linha, 2).value

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
                resultado = processar_pessoa(nome, numero_linha, planilha)

                if resultado is not None:
                    planilha.cell(numero_linha, 3).value = resultado
                    workbook.save(args.planilha)
                    print(f"Resultado salvo na coluna C, linha {numero_linha}.")
                else:
                    print(f"Linha {numero_linha} sem resultado. Seguindo adiante.")

            except KeyboardInterrupt:
                print("\nProcessamento interrompido pelo usuário.")
                break

            except Exception as erro:
                print(f"Erro na linha {numero_linha}: {erro}")
        
    finally:
        try:
            interpretar_excel_final(planilha)
            verificar_prazo(planilha)
            
            workbook.save(args.planilha)
            
            for numero_linha, nome in pessoas:
                try:
                    sucesso = lerMatricula(nome, numero_linha, planilha)

                    if sucesso:
                        workbook.save(args.planilha)

                    time.sleep(.5)

                except KeyboardInterrupt:
                    print("\nProcessamento interrompido pelo usuário.")
                    break

                except Exception as erro:
                    print(f"Erro na linha {numero_linha}: {erro}")
            
            print("Planilha salva com sucesso.")

        except PermissionError:
            print(
                "\nERRO: Não foi possível salvar a planilha.\n"
                "Feche o arquivo no Excel e execute novamente."
            )

        finally:
            workbook.close()

    print("\nProcessamento finalizado.")


if __name__ == "__main__":
    main()