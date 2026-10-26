import argparse
import time
from pathlib import Path

import re

from openpyxl.styles import PatternFill

import pyautogui
import pyperclip
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


TEMPO_LIMITE = 60
INTERVALO_VERIFICACAO = 1


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

    # ETAPA 1: aguardar o Menu.
    print("Esperando Menu aberto e Cadastro de Férias fechado...")

    janela = aguardar_janela(TITULO_MENU)
    janela.set_focus()

    # ETAPA 2: abrir a pesquisa.
    print("Abrindo pesquisa...")
    pyautogui.click(X_INICIAL, Y_INICIAL)

    janela = aguardar_janela(TITULO_PESQUISA)
    janela.set_focus()

    # ETAPA 3: pesquisar o servidor.
    print(f"Pesquisando servidor: {nome}")

    pyperclip.copy(nome)
    pyautogui.hotkey("ctrl", "v")

    print("Retornando ao menu...")
    pyautogui.press("enter")
    pyautogui.press("enter")

    # ETAPA 4: voltar ao Menu.
    janela = aguardar_janela(TITULO_MENU)
    janela.set_focus()

    # ETAPA 5: abrir Cadastro de Férias.
    print("Abrindo Cadastro de Férias...")
    pyautogui.click(X_CONSULTA, Y_CONSULTA)

    janela = aguardar_janela(TITULO_CADASTRO)
    janela.set_focus()

    # ETAPA 6: consultar o resultado.
    conteudo = consultar_resultado(planilha, numero_linha)

    if not conteudo:
        print(
            f"Nenhum conteúdo copiado após as tentativas para: {nome}. "
            "Prosseguindo para a próxima pessoa."
        )

        print("Fechando Cadastro de Férias com Esc...")
        pyautogui.press("esc")

        return None

    if janela is not None and janela.window_text().strip() == TITULO_CADASTRO:
        # ETAPA 7: fechar o cadastro.
        print("Fechando Cadastro de Férias com Esc...")
        pyautogui.press("esc")

    janela = aguardar_janela(TITULO_MENU)
    janela.set_focus()

    print(f"Processamento concluído: {nome}")
    return conteudo


# ============================================================
# IDENTIFICAÇÃO E ESPERA DE JANELAS
# ============================================================

def localizar_janela(titulo_esperado):
    """
    Localiza uma janela pelo título ou por um trecho dele.
    Retorna um objeto pywinauto ou None.
    """
    for janela in Desktop(backend="win32").windows():
        try:
            titulo = janela.window_text().strip()

            if titulo_esperado.casefold() in titulo.casefold():
                print(f"[OK] Janela identificada: {titulo}")
                return janela

        except Exception:
            continue

    return None


def aguardar_janela(titulo_esperado, timeout=TEMPO_LIMITE):
    """
    Aguarda até a janela aparecer.
    Retorna o objeto da janela ou None se atingir o limite.
    """
    print(f"[INFO] Aguardando a janela: {titulo_esperado}")

    limite = time.monotonic() + timeout

    while time.monotonic() < limite:
        janela = localizar_janela(titulo_esperado)

        if janela is not None:
            return janela

        time.sleep(INTERVALO_VERIFICACAO)

    print(
        f"[ERRO] A janela '{titulo_esperado}' "
        f"não apareceu em {timeout} segundos."
    )
    return None


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
    Se falhar, executa a recuperação e tenta novamente.
    A data copiada é convertida para o dia anterior.
    """

    print("Tentando copiar o resultado...")

    pyautogui.click(X_DATA, Y_DATA)
    pyautogui.click(X_DATA, Y_DATA)
    time.sleep(INTERVALO)

    conteudo = copiar_resultado()

    if not conteudo:
        print("Falha na primeira cópia. Iniciando recuperação...")

        recuperar_copia(planilha, numero_linha)

        # Recuperar a tela do Menu e colocar o foco nela.
        janela = aguardar_janela(TITULO_MENU)
        janela.set_focus()

        print("Tentando Ctrl+C novamente...")
        conteudo = copiar_resultado()

    if not conteudo:
        print("Não foi possível copiar o resultado após a recuperação.")
        return ""

    try:
        data = datetime.strptime(conteudo.strip(), "%d/%m/%Y")
    except ValueError:
        print(f"O conteúdo copiado não é uma data válida: {conteudo!r}")
        return ""

    data_anterior = data - timedelta(days=1)

    resultado = data_anterior.strftime("%d/%m/%Y")

    print(f"Data ajustada: {resultado}")

    return resultado


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


def lerMatricula(nome, numero_linha, planilha):
    print(f"\nLendo matrícula de: {nome}")

    # Garantir que o Menu esteja disponível.
    janela = aguardar_janela(TITULO_MENU)

    if janela is None:
        print("Não foi possível localizar o Menu.")
        return False

    janela.set_focus()

    # Abrir a pesquisa.
    print("Abrindo pesquisa...")
    pyautogui.click(X_INICIAL, Y_INICIAL)

    janela = aguardar_janela(TITULO_PESQUISA, timeout=15)

    if janela is None:
        print("A janela de pesquisa não foi aberta.")
        return False

    janela.set_focus()

    # Pesquisar o servidor.
    print(f"Pesquisando servidor: {nome}")

    pyperclip.copy(nome)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(INTERVALO)

    pyautogui.press("enter")
    time.sleep(INTERVALO)
    pyautogui.press("enter")

    janela = aguardar_janela(TITULO_MENU)

    if janela is None:
        print("O Menu não foi localizado após a pesquisa.")
        return False

    janela.set_focus()

    # Ler e registrar a matrícula.
    matricula = ler_matricula_menu()

    if matricula is not None:
        planilha.cell(numero_linha, 1).value = matricula

        print(
            f"Matrícula {matricula} preenchida na coluna A, "
            f"linha {numero_linha}."
        )
    else:
        print(f"Não foi possível obter a matrícula de {nome}.")

    # Ler o status do servidor.
    status = ler_status_menu()

    if status:
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


# ============================================================
# IDENTIFICAÇÃO E ESPERA DE JANELAS
# ============================================================

def localizar_janela(titulo_esperado):
    """
    Localiza uma janela pelo título ou por um trecho dele.
    Retorna o objeto da janela ou None se não encontrar.
    """
    for janela in Desktop(backend="win32").windows():
        try:
            titulo = janela.window_text().strip()

            if titulo_esperado.casefold() in titulo.casefold():
                print(f"[OK] Janela identificada: {titulo}")
                return janela

        except Exception:
            continue

    return None


def aguardar_janela(
    titulo_esperado,
    timeout=TEMPO_LIMITE,
):
    """
    Aguarda a janela aparecer.
    Retorna o objeto da janela ou None se atingir o limite.
    """
    print(f"[INFO] Aguardando a janela: {titulo_esperado}")

    limite = time.monotonic() + timeout

    while time.monotonic() < limite:
        janela = localizar_janela(titulo_esperado)

        if janela is not None:
            return janela

        time.sleep(INTERVALO_VERIFICACAO)

    print(
        f"[ERRO] A janela '{titulo_esperado}' "
        f"não apareceu em {timeout} segundos."
    )
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