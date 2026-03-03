import json
import subprocess
from datetime import datetime
from tkinter import Tk, filedialog

import xmltodict
from openpyxl import load_workbook

# ─────────────────────────────────────────────
# Carrega configurações externas (caminhos de rede)
# ─────────────────────────────────────────────
try:
    with open("config.json", encoding="utf-8") as config_file:
        CONFIG = json.load(config_file)
    CAMINHO_EXCEL = CONFIG["caminho_excel"]
    CAMINHO_XML   = CONFIG["caminho_xml"]
except FileNotFoundError:
    raise FileNotFoundError(
        "\nArquivo config.json não encontrado. "
        "Crie o arquivo com os campos 'caminho_excel' e 'caminho_xml'."
    )
except KeyError as e:
    raise KeyError(f"\nChave ausente no config.json: {e}")


# ─────────────────────────────────────────────
# Validação básica: verifica se o XML é uma NF-e
# ─────────────────────────────────────────────
def valida_estrutura_nfe(doc):
    """Retorna True se o documento possuir a estrutura mínima esperada de uma NF-e."""
    try:
        _ = doc["nfeProc"]["NFe"]["infNFe"]["emit"]["xNome"]
        _ = doc["nfeProc"]["NFe"]["infNFe"]["ide"]["nNF"]
        return True
    except (KeyError, TypeError):
        return False


# ─────────────────────────────────────────────
# Formatações
# ─────────────────────────────────────────────
def formata_valor(valor_string):
    """Formata valor numérico string para o padrão brasileiro (ex: 1.250,00)."""
    array_valor = list(valor_string)
    if len(array_valor) <= 6:
        return f"{''.join(array_valor[:-3])},{''.join(array_valor[-2:])}"
    else:
        return f"{''.join(array_valor[:-6])}.{''.join(array_valor[-6:-3])},{''.join(array_valor[-2:])}"


def fomata_data_padraoBR(string):
    """Converte data do formato ISO (AAAA-MM-DD) para o padrão brasileiro (DD/MM/AAAA)."""
    return datetime.strptime(string, "%Y-%m-%d").date().strftime("%d/%m/%Y")


# ─────────────────────────────────────────────
# Importação dos dados do XML
# ─────────────────────────────────────────────
def importa_dados_xml():
    arquivos_xml = buscaNotaExplorer()
    if not arquivos_xml:
        return None

    lista_pronta = []

    for arquivo_xml in arquivos_xml:
        # Leitura do arquivo XML
        try:
            with open(arquivo_xml, encoding="utf-8") as file:
                doc = xmltodict.parse(file.read())
        except (OSError, IOError) as e:
            print(f"\nErro ao abrir o arquivo {arquivo_xml}: {e}")
            return None
        except xmltodict.expat.ExpatError as e:
            print(f"\nArquivo XML inválido ou corrompido ({arquivo_xml}): {e}")
            return None

        # Validação da estrutura NF-e
        if not valida_estrutura_nfe(doc):
            print(f"\nO arquivo '{arquivo_xml}' não parece ser uma NF-e válida. Operação cancelada.")
            return None

        # Extração dos dados
        try:
            nome_fornecedor = doc["nfeProc"]["NFe"]["infNFe"]["emit"]["xNome"]
            numero_nfe      = int(doc["nfeProc"]["NFe"]["infNFe"]["ide"]["nNF"])
            data_de_emissao = fomata_data_padraoBR(
                doc["nfeProc"]["NFe"]["infNFe"]["ide"]["dhEmi"].split("T")[0]
            )
            valor = formata_valor(
                doc["nfeProc"]["NFe"]["infNFe"]["total"]["ICMSTot"]["vNF"]
            )
            frete = (
                "emitente"
                if doc["nfeProc"]["NFe"]["infNFe"]["transp"]["modFrete"] == "0"
                else "destinatario"
            )
        except (KeyError, TypeError, ValueError) as e:
            print(f"\nErro ao extrair dados do arquivo '{arquivo_xml}': {e}")
            return None

        # Processamento dos boletos/duplicatas
        boletos = []
        boletos_exibicao = ""

        try:
            for index, itens in enumerate(
                doc["nfeProc"]["NFe"]["infNFe"]["cobr"]["dup"]
            ):
                if isinstance(itens, str):
                    dup = doc["nfeProc"]["NFe"]["infNFe"]["cobr"]["dup"]
                    linha = f'{fomata_data_padraoBR(dup["dVenc"])} -- {formata_valor(dup["vDup"])}'
                    boletos.append(linha)
                    boletos_exibicao = linha
                    break
                else:
                    linha = f'{fomata_data_padraoBR(itens["dVenc"])} -- R$ {formata_valor(itens["vDup"])}'
                    boletos.append(linha)
                    boletos_exibicao += f'Boleto -- {index + 1}º: {linha} \n'
        except (KeyError, TypeError, ValueError) as e:
            print(f"\nAviso: não foi possível ler as duplicatas do arquivo '{arquivo_xml}': {e}")
            boletos.append("verificar as duplicatas na nfe ou com o fornecedor")
            boletos_exibicao = "verificar as duplicatas na nfe ou com o fornecedor"

        lista_pronta.append(
            [
                nome_fornecedor,
                numero_nfe,
                data_de_emissao[0:10],
                valor,
                frete,
                boletos,
                boletos_exibicao,
            ]
        )
    return lista_pronta


# ─────────────────────────────────────────────
# Gravação no Excel
# ─────────────────────────────────────────────
def cadastra_no_excel(lista_dados_xml):
    try:
        tabela_excel = load_workbook(CAMINHO_EXCEL)
    except (OSError, IOError) as e:
        print(f"\nErro ao abrir a planilha Excel: {e}")
        return

    aba_ativa = tabela_excel.active

    for dados_xml in lista_dados_xml:
        valores_a_ser_inserido = [
            f"{dados_xml[0]} - {datetime.today().strftime('%d/%m/%Y')}",
            dados_xml[1],
            dados_xml[2],
            float(dados_xml[3].replace(".", "").replace(",", ".")),
            dados_xml[4],
            dados_xml[5],
        ]
        if isinstance(dados_xml[5], list):
            valores_a_ser_inserido.pop()
            for boleto in dados_xml[5]:
                valores_a_ser_inserido.append(boleto)

        aba_ativa.append(valores_a_ser_inserido)

    try:
        tabela_excel.save(CAMINHO_EXCEL)
        print("\nDados gravados com sucesso")
    except (OSError, IOError) as e:
        print(f"\nErro ao salvar a planilha Excel: {e}")


# ─────────────────────────────────────────────
# Confirmação dos dados com o usuário
# ─────────────────────────────────────────────
def confere_os_dados_fornecidos(lista_dados_xml):
    if not lista_dados_xml:
        return

    tudo_ok = False
    for dados_xml in lista_dados_xml:
        valores_xml = (
            f"\nFornecedor: {dados_xml[0]}"
            f"\nNumero nfe: {dados_xml[1]}"
            f"\nData emissão: {dados_xml[2]}"
            f"\nValor: R$ {dados_xml[3]}"
            f"\nFrete: {dados_xml[4]}"
            f"\n{dados_xml[6]}"
        )

        resposta = input(f"{valores_xml}\n\nOs dados da nota fiscal estao corretos? (s ou n) \n")

        if resposta.lower() == "s":
            tudo_ok = True
        elif resposta.lower() == "n":
            tudo_ok = False
            print("\nFavor verificar o arquivo selecionado. Caso esteja correto, entre em contato com o suporte tecnico.")
        else:
            tudo_ok = False
            raise ValueError("\nFavor inserir um valor valido (s ou n)")

    if tudo_ok:
        cadastra_no_excel(lista_dados_xml)


# ─────────────────────────────────────────────
# Abertura do Excel para conferência
# ─────────────────────────────────────────────
def inicia_excel():
    try:
        subprocess.run(
            [
                "powershell",
                "-Command",
                f"Start-Process excel.exe '\"{CAMINHO_EXCEL}\"'",
            ]
        )
    except OSError as e:
        print(f"\nNão foi possível abrir o Excel automaticamente: {e}")


# ─────────────────────────────────────────────
# Loop para importação de múltiplas notas
# ─────────────────────────────────────────────
def mais_notas():
    while True:
        tem_mais_notas = input("\nDeseja importar mais notas? (s ou n)\n")

        if tem_mais_notas.lower() == "s":
            confere_os_dados_fornecidos(importa_dados_xml())
        elif tem_mais_notas.lower() == "n":
            print("\nMuito obrigado por usar nossos serviços\n")
            inicia_excel()
            break
        else:
            raise ValueError("\nFavor inserir um valor valido (s ou n)")


# ─────────────────────────────────────────────
# Ponto de entrada
# ─────────────────────────────────────────────
def buscaNotaExplorer():
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    arquivo_xml = filedialog.askopenfilenames(
        title="Selecione o arquivo XML da nota fiscal",
        initialdir=CAMINHO_XML,
        filetypes=[("Arquivos XML", "*.xml"), ("Todos os arquivos", "*.*")],
    )

    root.destroy()

    if not arquivo_xml:
        print("\nNenhum arquivo selecionado. Operação cancelada.")
        return None
    return list(arquivo_xml)


confere_os_dados_fornecidos(importa_dados_xml())
mais_notas()
