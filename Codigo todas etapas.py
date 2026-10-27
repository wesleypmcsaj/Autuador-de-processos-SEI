import struct
import subprocess
from datetime import datetime
import os
import time
from pathlib import Path
import shutil
import re
import fitz  # PyMuPDF
from PyPDF2 import PdfReader, PdfWriter
import PyPDF2
import pdfplumber
import json
import threading
import win32print
import win32api

import unicodedata

import imapclient
import pyzmail
import pyperclip
from tqdm import tqdm

from collections import Counter

from reportlab.pdfgen import canvas
from io import BytesIO
import gspread  # Para acessar o Google Planilhas
from oauth2client.service_account import ServiceAccountCredentials  # Para autenticação com conta de serviço
from openpyxl import load_workbook
import openpyxl  # Para manipular o Excel sem alterar a formatação
from PyPDF2 import PdfMerger
from send2trash import send2trash  # Biblioteca para mover arquivos para a lixeira
import pandas as pd
from docxtpl import DocxTemplate
from docx import Document


from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    NoSuchFrameException, StaleElementReferenceException,
)
from selenium.common.exceptions import NoSuchFrameException

from selenium.common.exceptions import StaleElementReferenceException


from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import  WebDriverException

import tempfile

from reportlab.lib.colors import red

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

import logging


# ----------------------------------------------------------------
# CONFIGURAÇÃO DOS LOGS
# ----------------------------------------------------------------

class FormatoRoxo(logging.Formatter):
    def format(self, record):
        mensagem = super().format(record)
        return f"\033[35m{mensagem}\033[0m"

#Inseri essa classe para ter a barra de progresso do TDQM
class HandlerTerminalTqdm(logging.StreamHandler):
    def emit(self, record):
        try:
            mensagem = self.format(record)
            tqdm.write(mensagem, file=self.stream)
        except Exception:
            self.handleError(record)


pasta_logs = Path("logs")
pasta_logs.mkdir(exist_ok=True)

data_hora = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
arquivo_log = pasta_logs / f"execucao_{data_hora}.log"

formato_log = (
    "%(asctime)s | %(levelname)s | "
    "%(filename)s:%(lineno)d | %(message)s"
)

# Arquivo de log com texto normal
handler_arquivo = logging.FileHandler(
    arquivo_log,
    encoding="utf-8"
)
handler_arquivo.setFormatter(logging.Formatter(formato_log))

# Terminal com texto roxo
#Substituirei essa linha abaixo funcional, para ter o controle da quantidade de processos que ainda precisa
#Inserir documentos usando o tdqm

#handler_terminal = logging.StreamHandler()
handler_terminal = HandlerTerminalTqdm()

handler_terminal.setFormatter(FormatoRoxo(formato_log))

logging.basicConfig(
    level=logging.INFO,
    handlers=[
        handler_arquivo,
        handler_terminal
    ]
)

logger = logging.getLogger(__name__)

# ----------------------------------------------------------------
# Código -1 [EXTRA]
# Verificação e limpeza de planilha para caso de execução do código pós erro
# ----------------------------------------------------------------

# Implementado em 04/07/2025

logger.info("============================================================")
logger.info("CÓDIGO -1 - Verificação de execução anterior")
logger.info("============================================================")


try:

    # Caminhos dos arquivos
    planilha_path = "dados_para_autuar_processos.xlsx"
    bloco_notas_path = "livro_de_registros_pa_procuradores_judiciais.txt"

    logger.info(
        "Código -1 - Abrindo planilha: %s",
        planilha_path
    )

    # Abrir a planilha mantendo a formatação
    wb = load_workbook(planilha_path)
    ws = wb.active

    logger.info(
        "Código -1 - Planilha aberta com sucesso"
    )

    # Identificar índice da coluna "Livro saj"
    header = [cell.value for cell in ws[1]]

    coluna_livro_saj_idx = header.index("Livro saj") + 1

    logger.info(
        "Código -1 - Coluna 'Livro saj' encontrada na posição %s",
        coluna_livro_saj_idx
    )

    # Coletar valores da coluna "Livro saj"
    valores_coluna = []

    for row in ws.iter_rows(
        min_row=2,
        min_col=coluna_livro_saj_idx,
        max_col=coluna_livro_saj_idx
    ):

        cell_value = row[0].value

        if (
            cell_value
            and isinstance(cell_value, str)
            and cell_value.strip()
        ):
            valores_coluna.append(
                cell_value.strip()
            )

    logger.info(
        "Código -1 - Foram encontrados %s registros na coluna 'Livro saj'",
        len(valores_coluna)
    )

    # Se houver valores para processar
    if valores_coluna:

        logger.warning(
            "Código -1 - Foram encontrados registros de uma execução anterior"
        )

        logger.info(
            "Código -1 - Iniciando recuperação dos registros"
        )

        # Contar quantas vezes cada valor aparece
        contador_planilha = Counter(valores_coluna)

        logger.info(
            "Código -1 - %s valores distintos serão analisados",
            len(contador_planilha)
        )

        logger.info(
            "Código -1 - Abrindo arquivo de registros: %s",
            bloco_notas_path
        )

        # Ler o conteúdo do bloco de notas
        with open(
            bloco_notas_path,
            "r",
            encoding="latin-1"
        ) as f:
            linhas_txt = f.readlines()

        logger.info(
            "Código -1 - Arquivo de registros lido com sucesso. Total de linhas: %s",
            len(linhas_txt)
        )

        # Preparar nova lista
        linhas_txt_atualizadas = linhas_txt.copy()

        total_registros_restaurados = 0

        for valor, qtd_planilha in contador_planilha.items():

            logger.info(
                "Código -1 - Procurando '%s'. Quantidade necessária: %s",
                valor,
                qtd_planilha
            )

            # Encontrar índices das linhas
            indices_alvo = [
                i
                for i, linha in enumerate(linhas_txt)
                if linha.strip().startswith(
                    f"{valor},(já utilizado)"
                )
            ]

            logger.info(
                "Código -1 - '%s': encontrados %s registros marcados como já utilizados",
                valor,
                len(indices_alvo)
            )

            # Selecionar os últimos índices
            indices_selecionados = indices_alvo[-qtd_planilha:]

            for idx in reversed(indices_selecionados):

                linhas_txt_atualizadas[idx] = (
                    linhas_txt_atualizadas[idx].replace(
                        ",(já utilizado)",
                        ",(utilizável)",
                        1
                    )
                )

                total_registros_restaurados += 1

        logger.info(
            "Código -1 - Total de registros restaurados: %s",
            total_registros_restaurados
        )

        # Escrever de volta o bloco de notas
        with open(
            bloco_notas_path,
            "w",
            encoding="latin-1"
        ) as f:
            f.writelines(linhas_txt_atualizadas)

        logger.info(
            "Código -1 - Arquivo de registros atualizado com sucesso"
        )

        logger.info(
            "Código -1 - Limpando dados das colunas A até I da planilha"
        )

        # Apagar colunas de A até L
        for row in ws.iter_rows(
            min_row=2,
            max_col=12
        ):

            for cell in row:
                cell.value = None

        logger.info(
            "Código -1 - Salvando planilha após limpeza"
        )

        wb.save(planilha_path)

        logger.info(
            "Código -1 - Procedimento de recuperação concluído com sucesso"
        )

        print(
            "Erro detectado em execução anterior do programa, "
            "procedimentos necessários tomados para nova execução da automação."
        )

    else:

        logger.info(
            "Código -1 - Nenhum registro de execução anterior encontrado"
        )

        print(
            "Não foram encontrados erros de execução anteriores, "
            "código pronto para copiar os dados registrados do Google Planilhas."
        )


except Exception:

    logger.exception(
        "Código -1 - ERRO durante a verificação/recuperação de execução anterior"
    )

    raise


# ----------------------------------------------------------------
# LIMPEZA DAS PASTAS DA AUTOMAÇÃO
# ----------------------------------------------------------------

logger.info(
    "Código -1 - Iniciando limpeza das pastas da automação"
)

# Pastas a serem verificadas
pastas = [
    "docs numerados",
    "docs PJs",
    "docs processados"
]

total_arquivos_lixeira = 0


for pasta in pastas:

    if os.path.exists(pasta):

        logger.info(
            "Código -1 - Verificando pasta: %s",
            pasta
        )

        arquivos_encontrados = os.listdir(pasta)

        logger.info(
            "Código -1 - Pasta '%s' possui %s itens",
            pasta,
            len(arquivos_encontrados)
        )

        for arquivo in arquivos_encontrados:

            caminho_completo = os.path.join(
                pasta,
                arquivo
            )

            # Só envia arquivos, não pastas
            if os.path.isfile(caminho_completo):

                try:

                    send2trash(caminho_completo)

                    total_arquivos_lixeira += 1

                    logger.info(
                        "Código -1 - Arquivo enviado para a lixeira: %s",
                        caminho_completo
                    )

                except Exception:

                    logger.exception(
                        "Código -1 - Erro ao mover arquivo para a lixeira: %s",
                        caminho_completo
                    )

    else:

        logger.warning(
            "Código -1 - Pasta não encontrada: %s",
            pasta
        )


logger.info(
    "Código -1 - Limpeza concluída. %s arquivos enviados para a lixeira",
    total_arquivos_lixeira
)

logger.info(
    "Código -1 - Finalizado com sucesso"
)


# ----------------------------------------------------------------
#Código 0, copiar dados do Google Planilhas
# ----------------------------------------------------------------




logger.info("Código 0 - Iniciando cópia dos dados do Google Planilhas para o Excel.")

print("Copiando os dados disponíveis da planilha Google, aguarde.")

# Configuração da autenticação
def autenticar_google():
    logger.info("Código 0 - Iniciando autenticação no Google.")

    # Define o escopo para acessar o Google Drive e Google Sheets
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets",
             "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]

    # Carrega as credenciais do arquivo JSON da conta de serviço
    creds = ServiceAccountCredentials.from_json_keyfile_name('chave.json', scope)
    cliente = gspread.authorize(creds)

    logger.info("Código 0 - Autenticação no Google concluída com sucesso.")

    return cliente

# Função principal para copiar dados do Google Planilhas para o Excel
def copiar_dados_para_excel():
    logger.info("Código 0 - Iniciando função copiar_dados_para_excel.")

    # Autentica e conecta ao Google Planilhas
    logger.info("Código 0 - Autenticando e conectando ao Google Planilhas.")
    cliente = autenticar_google()

    # Abre a planilha pelo ID e seleciona a aba "CAPTACOES"
    logger.info("Código 0 - Abrindo a planilha Google e selecionando a aba CAPTACOES.")
    planilha = cliente.open_by_key("16_zlC5bRdyGTqFcVFvRIBCYzP-fjoPN9i64tD5DGe5c")
    aba = planilha.worksheet("CAPTACOES")

    # Define os cabeçalhos esperados, caso algum esteja duplicado ou em branco
    expected_headers = ["ID processo", "Procurador", "Orgão"]

    logger.info("Código 0 - Lendo registros da aba CAPTACOES.")
    dados = aba.get_all_records(expected_headers=expected_headers)

    logger.info(
        "Código 0 - %s registros encontrados no Google Planilhas.",
        len(dados)
    )

    # Abre o arquivo Excel existente sem alterar a formatação
    arquivo_excel = "dados_para_autuar_processos.xlsx"

    logger.info(
        "Código 0 - Abrindo o arquivo Excel local: %s",
        arquivo_excel
    )

    workbook = openpyxl.load_workbook(arquivo_excel)
    sheet = workbook["CAPTACOES"]  # Nome da aba onde os dados serão adicionados

    # Limpa as linhas de dados antigas (sem apagar o cabeçalho)
    logger.info("Código 0 - Limpando dados antigos da aba CAPTACOES do Excel.")
    sheet.delete_rows(2, sheet.max_row)

    # Insere os dados a partir da segunda linha
    logger.info("Código 0 - Inserindo os novos registros no Excel.")

    for i, linha in enumerate(dados, start=2):  # Começa na linha 2 para manter o cabeçalho original
        sheet.cell(row=i, column=1, value=linha["ID processo"])
        sheet.cell(row=i, column=2, value=linha["Procurador"])
        sheet.cell(row=i, column=3, value=linha["Orgão"])

    # Salva o arquivo Excel com os dados atualizados
    logger.info("Código 0 - Salvando o arquivo Excel atualizado.")
    workbook.save(arquivo_excel)

    logger.info(
        "Código 0 - Dados copiados com sucesso do Google Planilhas para o Excel."
    )

    print("Dados copiados com sucesso do Google Planilhas para o Excel.")

# Executa a função para copiar os dados
try:
    copiar_dados_para_excel()

except Exception:
    logger.exception(
        "Código 0 - ERRO durante a cópia dos dados do Google Planilhas para o Excel."
    )
    raise


#Abaixo terei todos os 3 códigos para que eu baixe o arquivo com número de linhas,
#atualize a planilha e faça o upload do arquivo para a pasta correta do Gdrive novamente.
#


#---------------------------------------------------
# Código 0.1, Download do arquivo com o número de linhas atualizado
#---------------------------------------------------

logger.info("Código 0.1 - Iniciando download do arquivo com o número de linhas atualizado.")

print("Baixando arquivo de texto com o número de linhas atualizado do Google Drive.")
print("Por favor, aguarde um momento...")

SERVICE_ACCOUNT_FILE = "atualizar-numero-de-linhas-dd22a7445647.json"  # coloque o nome do seu arquivo aqui
SCOPES = ["https://www.googleapis.com/auth/drive"]
FOLDER_ID = "1BY492VgzAMM2FyZWMrKsOKCDx_sKZWzx"
ARQUIVO = "backup_linha.txt"


def encontrar_arquivo(service, nome, folder_id):
    logger.info(
        "Código 0.1 - Procurando o arquivo '%s' na pasta especificada do Google Drive.",
        nome
    )

    query = f"name='{nome}' and '{folder_id}' in parents and trashed=false"
    resp = service.files().list(q=query, fields="files(id, name)").execute()
    files = resp.get("files", [])

    if not files:
        logger.warning(
            "Código 0.1 - O arquivo '%s' não foi encontrado no Google Drive.",
            nome
        )
        return None

    logger.info(
        "Código 0.1 - Arquivo '%s' encontrado com sucesso no Google Drive.",
        nome
    )

    return files[0]["id"]


def baixar_arquivo():
    logger.info(
        "Código 0.1 - Carregando credenciais da conta de serviço."
    )

    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES
    )

    logger.info(
        "Código 0.1 - Credenciais carregadas com sucesso."
    )

    logger.info(
        "Código 0.1 - Criando conexão com a API do Google Drive."
    )

    service = build("drive", "v3", credentials=creds)

    logger.info(
        "Código 0.1 - Conexão com a API do Google Drive criada com sucesso."
    )

    file_id = encontrar_arquivo(service, ARQUIVO, FOLDER_ID)

    if not file_id:
        logger.warning(
            "Código 0.1 - Download não realizado porque o arquivo '%s' não foi encontrado.",
            ARQUIVO
        )

        print("❌ Arquivo não encontrado na pasta Cadernos Digitais.")
        return

    logger.info(
        "Código 0.1 - Iniciando download do arquivo '%s'.",
        ARQUIVO
    )

    print("⬇️ Baixando arquivo...")

    request = service.files().get_media(fileId=file_id)
    conteudo = request.execute()

    logger.info(
        "Código 0.1 - Conteúdo do arquivo recebido do Google Drive."
    )

    logger.info(
        "Código 0.1 - Gravando arquivo recebido em '%s'.",
        ARQUIVO
    )

    with open(ARQUIVO, "wb") as f:
        f.write(conteudo)

    logger.info(
        "Código 0.1 - Download e gravação do arquivo '%s' concluídos com sucesso.",
        ARQUIVO
    )

    print(f"✔️ Download concluído! ({ARQUIVO})")


if __name__ == "__main__":
    try:
        baixar_arquivo()

        logger.info(
            "Código 0.1 - Finalizado."
        )

    except Exception:
        logger.exception(
            "Código 0.1 - ERRO durante o download do arquivo do Google Drive."
        )
        raise


#----------------------------------------------
# Fim Código 0.1
#----------------------------------------------

# ----------------------------------------------------------------
#Código 1 inserir livro SAJ nas células correspondentes
# ----------------------------------------------------------------

logger.info("Código 1 - Iniciando inserção dos livros SAJ nas células correspondentes.")


def gerar_registros():
    logger.info("Código 1 - Iniciando geração do arquivo de registros.")

    # Cria o arquivo de registros com o padrão necessário
    with open("livro_de_registros_pa_procuradores_judiciais.txt", "w") as arquivo:
        for livro in range(19, 51):  # Começa no Livro 19 e vai até o Livro 50
            for pagina in range(1, 101):  # Vai da página 1 até a 100 (verso incluído)
                # Formata cada linha como "L <livro> P <pagina>, (utilizável)" e duplica a entrada
                arquivo.write(f"L {livro} P {pagina},(utilizável)\n")
                arquivo.write(f"L {livro} P {pagina},(utilizável)\n")  # Duplicação da página com "(utilizável)"
                arquivo.write(f"L {livro} P {pagina} (v),(utilizável)\n")
                arquivo.write(f"L {livro} P {pagina} (v),(utilizável)\n")  # Duplicação da página com "(v),(utilizável)"

    logger.info("Código 1 - Arquivo de registros criado com sucesso.")

    print("Arquivo de registros criado com sucesso.")


def verificar_e_preencher_excel():
    logger.info("Código 1 - Iniciando verificação e preenchimento da planilha.")

    # Verifica se o arquivo "livro_de_registros_pa_procuradores_judiciais.txt" existe; se não, gera o arquivo com o padrão especificado
    if not os.path.exists("livro_de_registros_pa_procuradores_judiciais.txt"):
        logger.warning(
            "Código 1 - Arquivo 'livro_de_registros_pa_procuradores_judiciais.txt' não encontrado."
        )

        print("Arquivo 'livro_de_registros_pa_procuradores_judiciais.txt' não encontrado. Gerando arquivo...")

        gerar_registros()

    logger.info(
        "Código 1 - Arquivo de registros localizado e disponível para uso."
    )

    # Abre o arquivo de Excel e a aba relevante
    arquivo_excel = "dados_para_autuar_processos.xlsx"

    logger.info(
        "Código 1 - Abrindo o arquivo Excel: %s",
        arquivo_excel
    )

    workbook = openpyxl.load_workbook(arquivo_excel)
    sheet = workbook["CAPTACOES"]

    logger.info(
        "Código 1 - Arquivo Excel aberto e aba CAPTACOES selecionada."
    )

    # Lê o arquivo de registros
    logger.info(
        "Código 1 - Lendo arquivo de registros."
    )

    with open("livro_de_registros_pa_procuradores_judiciais.txt", "r") as arquivo:
        registros = arquivo.readlines()

    logger.info(
        "Código 1 - Arquivo de registros carregado. Total de registros: %s",
        len(registros)
    )

    registros_inseridos = 0

    # Itera sobre as linhas da planilha do Excel para preencher "Livro saj" com registros disponíveis
    for row in range(2, sheet.max_row + 1):
        numero_processo = sheet.cell(row=row, column=1).value
        procurador = sheet.cell(row=row, column=2).value

        # Verifica se "numero processo" e "Procurador" estão preenchidos
        if numero_processo and procurador:
            # Encontra o primeiro registro com "(utilizável)"
            registro_disponivel = None
            for i, linha in enumerate(registros):
                if "(utilizável)" in linha:
                    registro_disponivel = linha.split(",")[0]
                    registros[i] = linha.replace("(utilizável)", "(já utilizado)")
                    break

            # Se não houver registros disponíveis, gera mais 50 livros e reinicia o preenchimento
            if not registro_disponivel:
                logger.warning(
                    "Código 1 - Todos os registros disponíveis foram utilizados. Gerando mais 50 livros."
                )

                print("Todos os registros foram utilizados. Gerando mais 50 livros.")

                with open("livro_de_registros_pa_procuradores_judiciais.txt", "a") as arquivo:
                    for livro in range(51, 101):
                        for pagina in range(1, 101):
                            arquivo.write(f"L {livro} P {pagina},(utilizável)\n")
                            arquivo.write(f"L {livro} P {pagina},(utilizável)\n")
                            arquivo.write(f"L {livro} P {pagina} (v),(utilizável)\n")
                            arquivo.write(f"L {livro} P {pagina} (v),(utilizável)\n")

                logger.info(
                    "Código 1 - Novos registros adicionados ao arquivo. Reiniciando preenchimento."
                )

                return verificar_e_preencher_excel()

            # Insere o registro disponível na coluna "Livro saj" (sexta coluna)
            sheet.cell(row=row, column=6, value=registro_disponivel)

            registros_inseridos += 1

            logger.info(
                "Código 1 - Registro '%s' inserido na linha %s da planilha.",
                registro_disponivel,
                row
            )

    logger.info(
        "Código 1 - Preenchimento concluído. Total de registros inseridos: %s",
        registros_inseridos
    )

    # Atualiza o arquivo de registros
    logger.info(
        "Código 1 - Atualizando o arquivo de registros."
    )

    with open("livro_de_registros_pa_procuradores_judiciais.txt", "w") as arquivo:
        arquivo.writelines(registros)

    logger.info(
        "Código 1 - Arquivo de registros atualizado com sucesso."
    )

    # Salva o arquivo Excel com as alterações
    logger.info(
        "Código 1 - Salvando o arquivo Excel com as alterações."
    )

    workbook.save(arquivo_excel)

    logger.info(
        "Código 1 - Livros de registros dos Pa's da procuradoria inseridos com sucesso."
    )

    print("Livros de registros dos Pa's da procuradoria inseridos com sucesso!")


# Gera o arquivo de registros na primeira execução, se necessário
try:
    verificar_e_preencher_excel()

    logger.info(
        "Código 1 - Finalizado."
    )

except Exception:
    logger.exception(
        "Código 1 - ERRO durante a inserção dos livros SAJ na planilha."
    )
    raise


####################
#FIM CÓDIGO 1
####################



# ----------------------------------------------------------------
# Código 2 baixar arquivos do ESAJ e inicio de completar informações necessárias
# ----------------------------------------------------------------

logger.info("Código 2 - Iniciando etapa de acesso ao ESAJ e complementação de informações.")


def update_webdriver():
    logger.info("Código 2 - Atualizando/inicializando o WebDriver.")

    print("Atualizando o WebDriver e completando os dados da tabela. Por favor, aguarde.")
    service = Service(ChromeDriverManager().install())

    logger.info("Código 2 - WebDriver preparado com sucesso.")

    return service


def realizar_login(driver):
    logger.info("Código 2 - Iniciando login inicial no ESAJ.")

    with open("login_esaj_TJSP.txt", "r") as file:
        login = file.readline().strip()
        senha = file.readline().strip()

    logger.info("Código 2 - Credenciais locais carregadas.")

    driver.get(
        "https://esaj.tjsp.jus.br/sajcas/login?service=https%3A%2F%2Fesaj.tjsp.jus.br%2Fesaj%2Fj_spring_cas_security_check"
    )

    logger.info("Código 2 - Página de login do ESAJ carregada.")

    driver.find_element(By.XPATH, '//*[@id="usernameForm"]').send_keys(login)
    driver.find_element(By.XPATH, '//*[@id="passwordForm"]').send_keys(senha)
    driver.find_element(By.XPATH, '//*[@id="pbEntrar"]').click()

    logger.info("Código 2 - Login inicial enviado. Aguardando tela do token.")

    print("Login inicial feito, aguardando tela do token...")
    time.sleep(5)


def buscar_token(driver):
    logger.info("Código 2 - Iniciando busca do token por e-mail.")

    email_usuario = 'protocolosexpedientesajpmc@gmail.com'
    senha_app = 'otby gbko eryn fgmz'

    imap = imapclient.IMAPClient('imap.gmail.com', ssl=True)
    imap.login(email_usuario, senha_app)
    imap.select_folder('INBOX')

    logger.info("Código 2 - Conectado à caixa de e-mail para busca do token.")

    criterio_busca = ['FROM', 'esaj@tjsp.jus.br']

    print("Aguardando o e-mail com o token...")

    logger.info("Código 2 - Aguardando e-mail enviado pelo ESAJ.")

    tentativas_sem_token = 0

    while True:
        imap.noop()

        mensagens = imap.search(criterio_busca)

        if mensagens:
            logger.info(
                "Código 2 - %s mensagem(ns) encontrada(s) para análise do token.",
                len(mensagens)
            )

            mensagens.sort(reverse=True)

            for uid in mensagens:
                mensagem = imap.fetch([uid], ['BODY[]'])
                email = pyzmail.PyzMessage.factory(mensagem[uid][b'BODY[]'])

                if email.text_part:
                    conteudo = email.text_part.get_payload().decode(email.text_part.charset)
                    resultado = re.search(r'(\d{6})', conteudo)

                    if resultado:
                        token = resultado.group(1)

                        logger.info("Código 2 - Token encontrado no e-mail.")

                        print(f"Token encontrado: {token}")

                        pyperclip.copy(token)
                        imap.move([uid], '[Gmail]/Trash')

                        logger.info("Código 2 - E-mail do token movido para a lixeira.")

                        print("E-mail movido para a lixeira.")

                        return token

        tentativas_sem_token += 1

        logger.info(
            "Código 2 - Token ainda não encontrado. Tentativa %s.",
            tentativas_sem_token
        )

        print(f"Token não encontrado - tentativa {tentativas_sem_token}")

        if tentativas_sem_token >= 61:
            logger.warning(
                "Código 2 - Token não encontrado após aproximadamente 120 segundos. Tentando reenviar."
            )

            print("Token não encontrado após 120s, clicando para reenviar o token...")

            try:
                driver.find_element(By.XPATH, '//*[@id="btnReceberToken"]').click()

                logger.info("Código 2 - Botão para reenvio do token acionado.")

                print("Cliquei para reenviar o token!")

            except:
                logger.warning(
                    "Código 2 - Não foi possível localizar o botão de reenviar token."
                )

                print("Não encontrei o botão de reenviar token.")

            tentativas_sem_token = 0

        time.sleep(2)


def inserir_token(driver, token):
    logger.info("Código 2 - Inserindo token recebido no ESAJ.")

    campo_token = driver.find_element(By.XPATH, '//*[@id="tokenInformado"]')
    campo_token.clear()
    campo_token.send_keys(token)
    driver.find_element(By.XPATH, '//*[@id="btnEnviarToken"]').click()

    print("Token inserido, aguardando validação...")

    logger.info("Código 2 - Token enviado. Aguardando validação.")

    try:
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="avisoTokenInvalido"]/span'))
        )

        logger.warning("Código 2 - Token inválido detectado.")

        print("Token inválido detectado, buscando novo token...")

        return False

    except:
        logger.info("Código 2 - Token aceito pelo ESAJ.")

        print("Token aceito!")

        return True


def login_esaj(driver):
    try:
        logger.info("Código 2 - Iniciando fluxo completo de login no ESAJ.")

        realizar_login(driver)

        while True:
            token = buscar_token(driver)

            if inserir_token(driver, token):
                break

            else:
                logger.warning("Código 2 - Novo token será solicitado.")

                print("Repetindo processo de busca de novo token...")

        logger.info("Código 2 - Login realizado com sucesso no sistema E-SAJ.")

        print("Login realizado no sistema E-SAJ")
        time.sleep(5)

    except Exception as e:
        logger.exception("Código 2 - ERRO durante o login no ESAJ.")

        print(f"Erro durante o login: {e}")
        raise


def acessar_site(driver):
    logger.info("Código 2 - Acessando página de consulta do ESAJ.")

    driver.get("https://esaj.tjsp.jus.br/cpopg/open.do")
    time.sleep(1)

    logger.info("Código 2 - Página de consulta acessada.")


def fechar_janelas_extras(driver, janela_principal):
    logger.info("Código 2 - Verificando janelas extras do navegador.")

    for handle in driver.window_handles:
        if handle != janela_principal:
            driver.switch_to.window(handle)
            driver.close()

    driver.switch_to.window(janela_principal)

    logger.info("Código 2 - Janelas extras fechadas.")


downloads_path = Path.home() / "Downloads"


def mover_arquivo_downloads(destino_pasta):
    logger.info(
        "Código 2 - Procurando PDF baixado para mover para a pasta '%s'.",
        destino_pasta
    )

    docs_pjs_path = Path(destino_pasta)
    docs_pjs_path.mkdir(parents=True, exist_ok=True)

    today = datetime.now().date()

    pdf_files_today = [
        f for f in downloads_path.glob('*.pdf')
        if datetime.fromtimestamp(f.stat().st_mtime).date() == today
    ]

    if pdf_files_today:
        latest_file = max(pdf_files_today, key=os.path.getmtime)

        logger.info(
            "Código 2 - PDF mais recente localizado: %s",
            latest_file.name
        )

        shutil.move(str(latest_file), str(docs_pjs_path / latest_file.name))

        logger.info(
            "Código 2 - PDF movido para '%s'.",
            destino_pasta
        )

    else:
        logger.warning(
            "Código 2 - Nenhum PDF baixado hoje foi encontrado na pasta Downloads."
        )

    with open("diretorio_docs_pjs.txt", "w") as file:
        file.write(str(docs_pjs_path))

    logger.info(
        "Código 2 - Caminho da pasta de documentos registrado em diretorio_docs_pjs.txt."
    )


def consultar_processo(driver, id_processo, row_index, sheet):
    logger.info(
        "Código 2 - Iniciando consulta do processo '%s' na linha %s.",
        id_processo,
        row_index
    )

    if id_processo:

        match = re.match(r"(\d{7})-(\d{2})\.(\d{4})\.8\.26\.(\d{4})", id_processo)

        if match:

            numero_digito_ano = f"{match.group(1)}-{match.group(2)}.{match.group(3)}"
            foro_numero_unificado = match.group(4)

            logger.info(
                "Código 2 - Processo '%s' reconhecido no formato esperado.",
                id_processo
            )

            driver.get("https://esaj.tjsp.jus.br/cpopg/open.do?gateway=true")
            time.sleep(1)

            select_element = driver.find_element(By.XPATH, '//*[@id="cbPesquisa"]')
            select = Select(select_element)
            select.select_by_visible_text("Número do Processo")

            driver.find_element(By.XPATH, '//*[@id="numeroDigitoAnoUnificado"]').send_keys(numero_digito_ano)
            driver.find_element(By.XPATH, '//*[@id="foroNumeroUnificado"]').send_keys(foro_numero_unificado)

            driver.find_element(By.XPATH, '//*[@id="botaoConsultarProcessos"]').click()
            time.sleep(10)

            logger.info(
                "Código 2 - Consulta enviada para o processo '%s'.",
                id_processo
            )

            try:
                classe_element = WebDriverWait(driver, 30).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="classeProcesso"]'))
                )

                classe_texto = classe_element.text
                sheet.cell(row=row_index, column=4, value=classe_texto)

                logger.info(
                    "Código 2 - Classe do processo '%s' encontrada.",
                    id_processo
                )

            except TimeoutException:
                logger.warning(
                    "Código 2 - Campo 'classeProcesso' não encontrado no processo '%s'.",
                    id_processo
                )

                print(f"Não encontrei o campo 'classeProcesso' no processo {id_processo}")
                sheet.cell(row=row_index, column=4, value="NÃO ENCONTRADO")

            assunto_texto = driver.find_element(By.XPATH, '//*[@id="assuntoProcesso"]').text
            sheet.cell(row=row_index, column=5, value=assunto_texto)

            logger.info(
                "Código 2 - Assunto do processo '%s' coletado.",
                id_processo
            )

            nome_parte_completo = driver.find_element(By.XPATH,
                                                      '//*[@id="tablePartesPrincipais"]/tbody/tr[1]/td[2]').text

            if "Advogada:" in nome_parte_completo:
                nome_parte = nome_parte_completo.split("Advogada:")[0].strip()

            elif "Advogado:" in nome_parte_completo:
                nome_parte = nome_parte_completo.split("Advogado:")[0].strip()

            else:
                nome_parte = nome_parte_completo

            sheet.cell(row=row_index, column=7, value=nome_parte)

            logger.info(
                "Código 2 - Nome da parte do processo '%s' coletado.",
                id_processo
            )

            sheet.parent.save("dados_para_autuar_processos.xlsx")

            logger.info(
                "Código 2 - Dados do processo '%s' salvos na planilha.",
                id_processo
            )

            try:

                logger.info(
                    "Código 2 - Iniciando download dos documentos do processo '%s'.",
                    id_processo
                )

                driver.find_element(By.XPATH, '//*[@id="linkPasta"]').click()

                WebDriverWait(driver, 10).until(lambda d: len(d.window_handles) > 1)

                janela_principal = driver.current_window_handle

                for handle in driver.window_handles:
                    if handle != driver.current_window_handle:
                        driver.switch_to.window(handle)
                        break

                driver.find_element(By.XPATH, '//*[@id="selecionarButton"]').click()
                time.sleep(3)

                driver.find_element(By.XPATH, '//*[@id="salvarButton"]').click()
                time.sleep(3)

                continuar_button = WebDriverWait(driver, 20).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="botaoContinuar"]'))
                )

                continuar_button.click()
                time.sleep(10)

                logger.info(
                    "Código 2 - Aguardando disponibilização do download do processo '%s'.",
                    id_processo
                )

                download_button = WebDriverWait(driver, 3600).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="btnDownloadDocumento"]'))
                )

                download_button.click()

                logger.info(
                    "Código 2 - Botão de download acionado para o processo '%s'.",
                    id_processo
                )

                time.sleep(10)

                driver.close()
                driver.switch_to.window(janela_principal)

                while any(downloads_path.glob("*.crdownload")):
                    time.sleep(1)

                logger.info(
                    "Código 2 - Download finalizado para o processo '%s'.",
                    id_processo
                )

                mover_arquivo_downloads("docs PJs")

            except Exception as e:
                logger.exception(
                    "Código 2 - ERRO durante o processo de download do processo '%s'.",
                    id_processo
                )

                print(f"Erro durante o processo de download: {e}")

        else:
            logger.warning(
                "Código 2 - ID do processo '%s' não está no formato esperado.",
                id_processo
            )

            print(f"Aviso: ID do processo '{id_processo}' não está no formato esperado.")

    else:
        logger.warning(
            "Código 2 - ID do processo vazio. Linha %s ignorada.",
            row_index
        )

        print("Aviso: O ID do processo está vazio. Linha ignorada.")


# ----------------------------------------------------------------
# NOVA FUNÇÃO DE VERIFICAÇÃO
# ----------------------------------------------------------------

def precisa_consultar_esaj(sheet):
    logger.info("Código 2 - Verificando se existem processos do TJSP para consulta no ESAJ.")

    for row in sheet.iter_rows(min_row=2, max_col=3, values_only=True):

        id_processo, procurador, orgao = row

        if id_processo and not id_processo.startswith("///") and orgao == "TJSP":
            logger.info("Código 2 - Existe ao menos um processo TJSP que exige consulta ao ESAJ.")
            return True

    logger.info("Código 2 - Nenhum processo exige consulta ao ESAJ.")

    return False


# ----------------------------------------------------------------
# FUNÇÃO PRINCIPAL
# ----------------------------------------------------------------

def processar_dados():
    logger.info("Código 2 - Iniciando função principal processar_dados.")

    wb = load_workbook("dados_para_autuar_processos.xlsx")
    sheet = wb["CAPTACOES"]

    logger.info("Código 2 - Planilha carregada para processamento.")

    # Verifica se precisa abrir o Selenium
    precisa_esaj = precisa_consultar_esaj(sheet)

    if precisa_esaj:

        logger.info("Código 2 - Processos TJSP encontrados. Iniciando Selenium.")

        chrome_options = Options()
        service = update_webdriver()
        driver = webdriver.Chrome(service=service, options=chrome_options)

        logger.info("Código 2 - Navegador Chrome iniciado.")

        login_esaj(driver)

    else:

        logger.info(
            "Código 2 - Nenhum processo do TJSP precisa ser consultado. Selenium não será aberto."
        )

        print("Nenhum processo do TJSP precisa ser consultado. Selenium não será aberto.")
        driver = None

    for i, row in enumerate(sheet.iter_rows(min_row=2, max_col=8, values_only=True), start=2):

        id_processo, procurador, orgao = row[0], row[1], row[2]

        if id_processo and id_processo.startswith("///"):

            logger.info(
                "Código 2 - Linha %s iniciada com '///'. Processamento local sem Selenium.",
                i
            )

            partes = id_processo[3:].strip().split(" / ")

            if len(partes) >= 4:

                numero_processo, classe, assunto, nome_parte = partes[:4]

                sheet.cell(row=i, column=4, value=classe)
                sheet.cell(row=i, column=5, value=assunto)
                sheet.cell(row=i, column=7, value=nome_parte)

                if procurador:

                    informacoes_completas = (
                        f"Ação envolvendo {nome_parte} e PREFEITURA MUNICIPAL DE CARAPICUIBA {classe} {assunto} {numero_processo}"
                    )

                else:

                    informacoes_completas = (
                        f"Solicitação de providências/ solicitação de informações envolvendo {nome_parte} e PREFEITURA MUNICIPAL DE CARAPICUIBA {classe} {assunto} {numero_processo}"
                    )

                sheet.cell(row=i, column=8, value=informacoes_completas)

            logger.info(
                "Código 2 - Linha %s processada sem uso do Selenium.",
                i
            )

            print(f"Linha {i} iniciada com '///' foi processada sem usar o Selenium.")
            continue

        if precisa_esaj and orgao == "TJSP" and id_processo:

            logger.info(
                "Código 2 - Linha %s será consultada no ESAJ. Processo: %s",
                i,
                id_processo
            )

            consultar_processo(driver, id_processo, i, sheet)

    logger.info("Código 2 - Salvando planilha após processamento.")

    wb.save("dados_para_autuar_processos.xlsx")

    logger.info("Código 2 - Planilha salva com sucesso.")

    if precisa_esaj:

        logger.info("Código 2 - Encerrando navegador.")

        driver.quit()

        logger.info("Código 2 - Navegador encerrado.")


if __name__ == "__main__":
    try:
        processar_dados()

        logger.info("Código 2 - Finalizado.")

    except Exception:
        logger.exception(
            "Código 2 - ERRO não tratado durante o processamento do ESAJ."
        )
        raise

##############
# FIM CÓDIGO 2
##############

# Continuar a partir do código 3



# ----------------------------------------------------------------
#Código 3 - Finalizar as informações nescessarias para autuar os processos no GIAP
# ----------------------------------------------------------------



logger.info("Código 3 - Iniciando finalização das informações necessárias para autuar os processos no GIAP.")





# Mensagem inicial
print("As informações necessárias para autuar os processos administrativos Judiciais estão sendo completadas, por favor aguarde.")

logger.info("Código 3 - Carregando arquivo Excel com os dados dos processos.")

# Carregar o arquivo Excel
arquivo_excel = "dados_para_autuar_processos.xlsx"
workbook = openpyxl.load_workbook(arquivo_excel)
planilha = workbook.active

logger.info("Código 3 - Arquivo Excel carregado com sucesso.")

# Identificar os índices das colunas (usando os nomes fornecidos)
coluna_id_processo = 1  # "ID processo"
coluna_procurador = 2  # "Procurador"
coluna_orgao = 3  # "Orgão"
coluna_classe = 4  # "Classe"
coluna_assunto = 5  # "Assunto"
coluna_livro_saj = 6  # "Livro saj"
coluna_nome_parte = 7  # "Nome da parte"
coluna_info_completa = 8  # "Informacoes completas para autuar o PA"

# Função para extrair o número do processo judicial
def extrair_numero_processo(id_processo):
    if id_processo is None:
        return ""
    id_processo = str(id_processo)
    if "///" in id_processo:
        partes = id_processo.split("/")
        for parte in partes:
            if parte and len(parte) >= 25 and "-" in parte and "." in parte:
                return parte
    elif len(id_processo) >= 25 and "-" in id_processo and "." in id_processo:
        return id_processo
    return ""

try:
    logger.info(
        "Código 3 - Iniciando processamento de %s linhas da planilha.",
        max(planilha.max_row - 1, 0)
    )

    linhas_processadas = 0

    # Iterar pelas linhas (ignorar a primeira linha, que é o cabeçalho)
    for linha in range(2, planilha.max_row + 1):
        procurador = planilha.cell(row=linha, column=coluna_procurador).value
        orgao = planilha.cell(row=linha, column=coluna_orgao).value
        nome_parte = planilha.cell(row=linha, column=coluna_nome_parte).value
        classe = planilha.cell(row=linha, column=coluna_classe).value
        assunto = planilha.cell(row=linha, column=coluna_assunto).value
        id_processo = planilha.cell(row=linha, column=coluna_id_processo).value
        livro_saj = planilha.cell(row=linha, column=coluna_livro_saj).value

        logger.info(
            "Código 3 - Processando linha %s. ID processo: %s",
            linha,
            id_processo
        )

        numero_processo = extrair_numero_processo(id_processo)  # Extrair número do processo judicial

        # Adicionar espaço simples e valor da coluna "Livro saj" (ou apenas espaço se vazio)
        complemento_livro_saj = f" {livro_saj}" if livro_saj else " "

        # Condicional para quando o "Orgão" for "TJSP"
        if orgao == "TJSP":
            if procurador:  # Se o "Procurador" estiver preenchido
                texto = (f"Ação Judicial envolvendo {nome_parte} e PREFEITURA MUNICIPAL DE CARAPICUIBA. "
                         f"{classe}   {assunto}  - número do processo: {numero_processo} {complemento_livro_saj} {procurador}")

                logger.info(
                    "Código 3 - Linha %s tratada como processo TJSP com procurador.",
                    linha
                )

            else:  # Se o "Procurador" estiver vazio
                texto = (f"Solicitação de providências/ Solicitação de Informações envolvendo {nome_parte} "
                         f"e PREFEITURA MUNICIPAL DE CARAPICUIBA   {classe}   {assunto} - número do processo: {numero_processo}")

                logger.info(
                    "Código 3 - Linha %s tratada como processo TJSP sem procurador.",
                    linha
                )

        else:  # Quando o "Orgão" for diferente de "TJSP"
            texto = f"Solicitação de providências/ Solicitação de Informações envolvendo {id_processo} {complemento_livro_saj} {procurador}"

            logger.info(
                "Código 3 - Linha %s tratada como processo de órgão diferente do TJSP.",
                linha
            )

        # Inserir o texto na 8ª coluna ("Informacoes completas para autuar o PA")
        planilha.cell(row=linha, column=coluna_info_completa).value = texto

        linhas_processadas += 1

    logger.info(
        "Código 3 - Processamento das linhas concluído. Total de linhas processadas: %s",
        linhas_processadas
    )

    # Salvar o arquivo Excel atualizado
    logger.info("Código 3 - Salvando arquivo Excel atualizado.")

    workbook.save(arquivo_excel)

    logger.info("Código 3 - Arquivo Excel salvo com sucesso.")

    # Mensagem final
    print("Todas as informações necessárias para autuar os processos foram inseridas.")

    logger.info("Código 3 - Todas as informações necessárias para autuação foram inseridas.")

except Exception:
    logger.exception(
        "Código 3 - ERRO durante a finalização das informações para autuar os processos no GIAP."
    )
    raise


print("Os documentos estão sendo divididos em suas respectivas petições e decisões. Esse processo tende a demorar.")

logger.info(
    "Código 3 - Finalizado. Próxima etapa: divisão dos documentos em petições e decisões."
)


################
# FIM Código 3
################




###############
#             #
# Código 3.1  #
#             #
###############

#Fazer a classificação dos processos em segredo de justiça ou não.

#Isso é util para definir assim um processo no SEI, aquele radio button lá e qual texto usar
# para justificar o uso de segredo de justiça.

# ----------------------------------------------------------------
# Classificar processos em segredo de justiça para autuação no SEI
# ----------------------------------------------------------------

logger = logging.getLogger("classificacao_segredo_justica")
logger.setLevel(logging.INFO)
logger.propagate = False

# Evitar duplicação dos logs se o código for executado novamente
if not logger.handlers:
    formato = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(filename)s:%(lineno)d | %(message)s"
    )

    arquivo_log = logging.FileHandler(
        "classificacao_segredo_justica.log",
        encoding="utf-8"
    )
    arquivo_log.setFormatter(formato)

    terminal_log = logging.StreamHandler()
    terminal_log.setFormatter(formato)

    logger.addHandler(arquivo_log)
    logger.addHandler(terminal_log)

arquivo_excel = "dados_para_autuar_processos.xlsx"
workbook = None

try:
    logger.info("Iniciando classificação dos processos em segredo de justiça.")
    logger.info("Carregando arquivo Excel: %s", arquivo_excel)

    workbook = openpyxl.load_workbook(arquivo_excel)
    planilha = workbook.active

    logger.info("Arquivo Excel carregado com sucesso.")

    linhas_processadas = 0

    # Ignorar a primeira linha, que contém os cabeçalhos
    for linha in range(2, planilha.max_row + 1):
        id_processo = planilha.cell(row=linha, column=1).value

        # Parar na primeira linha com o ID processo vazio
        if id_processo is None or str(id_processo).strip() == "":
            logger.info(
                "Coluna A vazia na linha %s. Encerrando processamento.",
                linha
            )
            break

        # Verificar se o ID processo começa exatamente com "///"
        classificacao = "SIM" if str(id_processo).startswith("///") else "NAO"

        # Preencher a coluna J da mesma linha
        planilha.cell(row=linha, column=10).value = classificacao

        linhas_processadas += 1

        logger.info(
            "Linha %s - Segredo de justiça: %s",
            linha,
            classificacao
        )

    logger.info(
        "Classificação concluída. Total de linhas processadas: %s",
        linhas_processadas
    )

    logger.info("Salvando arquivo Excel.")
    workbook.save(arquivo_excel)
    logger.info("Arquivo Excel salvo com sucesso.")

    print(
        "Processos classificados em segredo de justiça ou não "
        "para autuação no sistema SEI"
    )

    logger.info("Finalização da classificação realizada com sucesso.")

except Exception:
    logger.exception(
        "ERRO durante a classificação dos processos em segredo de justiça."
    )
    raise

finally:
    if workbook is not None:
        workbook.close()



###################
#                 #
# Fim Código 3.1  #
#                 #
###################




"""Preparação da planilha antes de abrir o SEI.

Dependência: python -m pip install openpyxl
Execução: python preparar_identificacao_processos.py
Coloque a planilha na mesma pasta deste script e feche-a no Excel.

Regras do DOCX:
  A: número do processo, sozinho ou precedido de /// (com ou sem espaço).
  B: Procurador. G: Nome da parte. K: Curta identificacao do processo.
  K recebe: NUMERO - PROCURADOR - Parte: NOME, limitado a 99 caracteres.
  O processamento começa na linha 2 e para na primeira A vazia.

Integração com outro código:
  from preparar_identificacao_processos import executar_preparacao_identificacoes
  executar_preparacao_identificacoes(caminho_planilha, logger=logger)
  # Somente depois, iniciar as próximas etapas e abrir o SEI.
"""


#Código 3.2 adicionado em 21/10/2026

###################################################
#                                                 #
# Criar texto para fácil identificação do processo#
#                                                 #
###################################################




# ----------------------------------------------------------------
# CONFIGURAÇÕES — ALTERE AQUI QUANDO NECESSÁRIO
# ----------------------------------------------------------------

PASTA_BASE = Path(__file__).resolve().parent
CAMINHO_PLANILHA = PASTA_BASE / "dados_para_autuar_processos.xlsx"
NOME_ABA = None  # None utiliza a aba ativa; ou informe, por exemplo, "Planilha1".
LINHA_INICIAL = 2  # Linha 1 contém os cabeçalhos.
LIMITE_CARACTERES = 99

PADRAO_PROCESSO = re.compile(
    r"(?:///\s*)?([0-9]{7}-[0-9]{2}\.[0-9]{4}\.[0-9]\.[0-9]{2}\.[0-9]{4})"
)


# ----------------------------------------------------------------
# CONFIGURAÇÃO DOS LOGS — ROXO NO TERMINAL E TEXTO NORMAL NO ARQUIVO
# ----------------------------------------------------------------

class FormatoRoxo(logging.Formatter):
    def format(self, record):
        return f"\033[35m{super().format(record)}\033[0m"


def configurar_logs():
    logger = logging.getLogger("preparacao_identificacoes")
    logger.setLevel(logging.INFO)
    logger.propagate = False
    if logger.handlers:
        return logger

    pasta_logs = PASTA_BASE / "logs"
    pasta_logs.mkdir(exist_ok=True)
    data_hora = datetime.now().strftime("%Y-%m-%d_%H-%M-%S_%f")
    formato = "%(asctime)s | %(levelname)s | %(filename)s:%(lineno)d | %(message)s"

    arquivo = logging.FileHandler(
        pasta_logs / f"preparacao_identificacoes_{data_hora}.log", encoding="utf-8"
    )
    arquivo.setFormatter(logging.Formatter(formato))
    terminal = logging.StreamHandler()
    terminal.setFormatter(FormatoRoxo(formato))
    logger.addHandler(arquivo)
    logger.addHandler(terminal)
    return logger


# ----------------------------------------------------------------
# 1 - EXTRAIR O NÚMERO NAS DUAS CONDIÇÕES DO DOCUMENTO
# ----------------------------------------------------------------

def extrair_numero_processo(valor):
    texto = str(valor).strip()

    if texto.startswith("///"):
        # Remove as três barras e possíveis espaços seguintes.
        texto = texto[3:].lstrip()

        # Captura apenas o número no início, ignorando a descrição.
        correspondencia = re.match(
            r"([0-9]{7}-[0-9]{2}\.[0-9]{4}\.[0-9]\.[0-9]{2}\.[0-9]{4})(?=$|\s|/)",
            texto
        )

    else:
        # Sem as barras, mantém a regra de conter somente o número.
        correspondencia = re.fullmatch(
            r"([0-9]{7}-[0-9]{2}\.[0-9]{4}\.[0-9]\.[0-9]{2}\.[0-9]{4})",
            texto
        )

    if correspondencia is None:
        raise ValueError(
            "Não foi possível identificar o número do processo, "
            "sozinho ou após ///."
        )

    return correspondencia.group(1)

# ----------------------------------------------------------------
# 2 - MONTAR A IDENTIFICAÇÃO E CORTAR O EXCESSO À DIREITA
# ----------------------------------------------------------------

def montar_identificacao(numero, procurador, nome_parte):
    procurador = "" if procurador is None else str(procurador)
    nome_parte = "" if nome_parte is None else str(nome_parte)
    texto_completo = f"{numero} - {procurador} - Parte: {nome_parte}"
    return texto_completo[:LIMITE_CARACTERES], len(texto_completo)


# ----------------------------------------------------------------
# 3 - PROCESSAR A PLANILHA E SALVAR ANTES DE CONTINUAR O FLUXO
# ----------------------------------------------------------------

def executar_preparacao_identificacoes(
    caminho_planilha=CAMINHO_PLANILHA, nome_aba=NOME_ABA, logger=None
):
    """Atualiza K na mesma planilha e retorna a quantidade de linhas processadas.

    Cria uma cópia de segurança antes da gravação. Um ID inválido ou fórmula
    em A/B/G interrompe a etapa sem salvar alterações, para permitir correção.
    B/G vazias geram aviso e são tratadas como texto vazio.
    """
    if logger is None:
        logger = configurar_logs()
    caminho = Path(caminho_planilha).resolve()
    livro = None
    temporario = None
    logger.info("============================================================")
    logger.info("PREPARAÇÃO - GERAR CURTA IDENTIFICAÇÃO DOS PROCESSOS")
    logger.info("============================================================")

    try:
        logger.info("Preparação - Abrindo planilha: %s", caminho)
        livro = load_workbook(caminho)
        aba = livro[nome_aba] if nome_aba is not None else livro.active
        logger.info("Preparação - Aba selecionada: %s", aba.title)
        quantidade = 0
        truncadas = 0

        for linha in range(LINHA_INICIAL, aba.max_row + 2):
            id_processo = aba.cell(linha, 1).value
            if id_processo is None or str(id_processo).strip() == "":
                logger.info("Preparação - A%s vazia. Fim do processamento.", linha)
                break

            logger.info("Preparação - Processando linha %s.", linha)
            for coluna in (1, 2, 7):
                celula = aba.cell(linha, coluna)
                if celula.data_type == "f":
                    raise ValueError(
                        f"{celula.coordinate} contém fórmula. Esta etapa requer valores "
                        "em A, B e G; openpyxl não calcula fórmulas."
                    )
            try:
                numero = extrair_numero_processo(id_processo)
            except ValueError as erro:
                raise ValueError(f"Linha {linha}, coluna A: {erro}") from erro

            condicao = "com ///" if str(id_processo).strip().startswith("///") else "sem ///"
            logger.info("Preparação - Linha %s: número identificado %s.", linha, condicao)
            procurador = aba.cell(linha, 2).value
            nome_parte = aba.cell(linha, 7).value
            for coluna, valor in (("B", procurador), ("G", nome_parte)):
                if valor is None or str(valor).strip() == "":
                    logger.warning("Preparação - %s%s vazia; usando texto vazio.", coluna, linha)

            identificacao, tamanho_original = montar_identificacao(numero, procurador, nome_parte)
            if tamanho_original > LIMITE_CARACTERES:
                truncadas += 1
                logger.info(
                    "Preparação - Linha %s: removidos %s caracteres do final.",
                    linha, tamanho_original - LIMITE_CARACTERES,
                )
            aba.cell(linha, 11).value = identificacao
            quantidade += 1
            logger.info("Preparação - K%s preparada com %s caracteres.", linha, len(identificacao))

        if quantidade == 0:
            logger.info("Preparação - Nenhuma linha para processar. Arquivo não alterado.")
            return 0

        # Salvar primeiro em arquivo temporário na mesma pasta.
        # Assim, falhas durante a geração não sobrescrevem a planilha original.
        with tempfile.NamedTemporaryFile(
            prefix=".identificacoes_", suffix=".xlsx", dir=caminho.parent, delete=False
        ) as arquivo:
            temporario = Path(arquivo.name)
        livro.save(temporario)
        livro.close()
        livro = None

        pasta_backup = caminho.parent / "backups"
        pasta_backup.mkdir(exist_ok=True)
        momento = datetime.now().strftime("%Y-%m-%d_%H-%M-%S_%f")
        backup = pasta_backup / f"{caminho.stem}_{momento}{caminho.suffix}"
        shutil.copy2(caminho, backup)
        logger.info("Preparação - Cópia de segurança criada: %s", backup)
        os.replace(temporario, caminho)
        temporario = None
        logger.info(
            "Preparação - Planilha salva. Linhas processadas: %s. "
            "Identificações cortadas para 99 caracteres: %s.", quantidade, truncadas,
        )
        return quantidade

    except Exception:
        logger.exception("Preparação - ERRO. Etapa interrompida; verificar os detalhes acima.")
        raise
    finally:
        if livro is not None:
            livro.close()
        if temporario is not None:
            temporario.unlink(missing_ok=True)


# ----------------------------------------------------------------
# EXECUTAR — ESTA ETAPA NÃO ABRE O SEI
# ----------------------------------------------------------------

if __name__ == "__main__":
    logger = configurar_logs()
    executar_preparacao_identificacoes(logger=logger)





#Fim código 3.2

###################################################
#                                                 #
# CÓDIGO 3.3 Só deixar o PJ em uma célula sozinha #
#                                                 #
###################################################





# ----------------------------------------------------------------
# CONFIGURAÇÕES
# ----------------------------------------------------------------

PASTA_DO_SCRIPT = Path(__file__).resolve().parent

CAMINHO_PLANILHA = PASTA_DO_SCRIPT / "dados_para_autuar_processos.xlsx"

# None utiliza a aba ativa.
# Para escolher uma aba específica, use, por exemplo: "Planilha1".
NOME_ABA = None

# Use 2 se a primeira linha contiver cabeçalhos.
# Use 1 se os dados começarem na primeira linha.
LINHA_INICIAL = 2


# ----------------------------------------------------------------
# CONFIGURAÇÃO DO LOGGER
# ----------------------------------------------------------------

logger = logging.getLogger("pre_sei")


def configurar_logs():
    logger.setLevel(logging.INFO)
    logger.propagate = False

    if logger.handlers:
        return

    # ------------------------------------------------------------
    # FORMATO DOS LOGS NO ARQUIVO
    # ------------------------------------------------------------

    formato = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(message)s",
        datefmt="%d/%m/%Y %H:%M:%S",
    )

    # ------------------------------------------------------------
    # FORMATO ROXO DOS LOGS NO TERMINAL
    # ------------------------------------------------------------

    formato_roxo = logging.Formatter(
        "\033[35m"
        "%(asctime)s | %(levelname)s | %(message)s"
        "\033[0m",
        datefmt="%d/%m/%Y %H:%M:%S",
    )

    console = logging.StreamHandler()
    console.setFormatter(formato_roxo)

    # ------------------------------------------------------------
    # SALVAR ARQUIVO NA PASTA LOGS
    # ------------------------------------------------------------

    arquivo_log = logging.FileHandler(
        PASTA_DO_SCRIPT / "logs" / "pre_sei.log",
        encoding="utf-8",
    )
    arquivo_log.setFormatter(formato)

    logger.addHandler(console)
    logger.addHandler(arquivo_log)


# ----------------------------------------------------------------
# PRÉ-SEI - COPIAR NÚMEROS DA COLUNA A PARA A COLUNA L
# ----------------------------------------------------------------

def preparar_planilha(
    caminho_planilha=CAMINHO_PLANILHA,
    nome_aba=NOME_ABA,
    linha_inicial=LINHA_INICIAL,
):
    """
    Lê a coluna A e copia o número do processo para a coluna L.

    Aceita:
        1002746-20.2026.8.26.0127
        /// 1006021-74.2026.8.26.0127 / outras informações

    Ignora as informações posteriores ao número.
    Mantém a coluna L quando A está vazia ou fora do padrão.
    """

    if linha_inicial < 1:
        raise ValueError("A linha inicial deve ser maior ou igual a 1.")

    padrao = re.compile(
        r"^\s*(?:///\s*)?"
        r"([0-9]{7}-[0-9]{2}\.[0-9]{4}\.[0-9]\.[0-9]{2}\.[0-9]{4})"
        r"(?![0-9])"
    )

    preenchidas = 0
    vazias = 0
    sem_padrao = 0

    logger.info(
        "PRÉ-SEI - Abrindo planilha: %s",
        caminho_planilha,
    )

    workbook = load_workbook(caminho_planilha)

    try:
        planilha = (
            workbook.active
            if nome_aba is None
            else workbook[nome_aba]
        )

        logger.info(
            "PRÉ-SEI - Processando aba: %s.",
            planilha.title,
        )

        for linha in range(linha_inicial, planilha.max_row + 1):
            valor = planilha.cell(row=linha, column=1).value

            if valor is None or not str(valor).strip():
                vazias += 1
                continue

            resultado = padrao.match(str(valor))

            if resultado is None:
                sem_padrao += 1

                logger.warning(
                    "PRÉ-SEI - Linha %s: padrão não encontrado "
                    "na coluna A. Coluna L mantida.",
                    linha,
                )
                continue

            numero_processo = resultado.group(1)

            destino = planilha.cell(row=linha, column=12)
            destino.value = numero_processo
            destino.number_format = "@"

            preenchidas += 1

            logger.info(
                "PRÉ-SEI - Linha %s: %s copiado para a coluna L.",
                linha,
                numero_processo,
            )

        logger.info("PRÉ-SEI - Salvando planilha.")

        workbook.save(caminho_planilha)

        logger.info(
            "PRÉ-SEI - Concluído. "
            "Preenchidas: %s | Vazias: %s | Sem padrão: %s.",
            preenchidas,
            vazias,
            sem_padrao,
        )

    finally:
        workbook.close()


# ----------------------------------------------------------------
# EXECUÇÃO INDEPENDENTE
# ----------------------------------------------------------------

def main():
    configurar_logs()

    logger.info("PRÉ-SEI - Programa iniciado.")

    try:
        preparar_planilha()

    except PermissionError:
        logger.exception(
            "PRÉ-SEI - Acesso negado à planilha. "
            "Verifique se ela está aberta no Excel "
            "ou se a pasta permite gravação."
        )
        return 1

    except Exception:
        logger.exception(
            "PRÉ-SEI - Falha no processamento da planilha."
        )
        return 1

    logger.info("PRÉ-SEI - Programa finalizado.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())



# FIM CÓDIGO 3.3



# ----------------------------------------------------------------
#Código 4 dividir os PDF's em suas respectivas petições e Decisões alterado em 24-06-2025
# ----------------------------------------------------------------


#Novo código, ele possui agora uma variação  utilizando "pdfplumber" para caso o PyPDF não consiga processar algum arquivo em especifico,
#como se fosse uma contramedida.


logger.info("Código 4 - Iniciando divisão dos PDFs em petições e decisões.")


# Configuração de pastas
PASTA_ORIGEM = "docs PJs"
PASTA_SAIDA = "docs processados"
os.makedirs(PASTA_SAIDA, exist_ok=True)

logger.info(
    "Código 4 - Pastas configuradas. Origem: '%s' | Saída: '%s'.",
    PASTA_ORIGEM,
    PASTA_SAIDA
)


# Termos de busca aprimorados (combinando ambos códigos)
TERMINOS_PETICAO = [
    "Dá-se", "Dá à",
    "Nestes termos p. deferimento", "Termos em que, Pede Deferimento.", "Pede Deferimento",
    "Nestes termos, pede deferimento", "Nesses termos pede e aguarda deferimento",
    "Termos em que, Pede e Espera Deferimento.", "Termosemque, PedeeEsperaDeferimento.",
    "Termosemque, PedeeEsperaDeferimento", "Termos em que, Pede e Espera Deferimento",
    "Termos em que,", "p. deferimento", "Pede e Espera Deferimento.", "Termosemque,",
    "PedeeEsperaDeferimento.", "PedeeEsperaDeferimento", "Nestes termos, pede e espera deferimento",
    "Nestes termos, espera deferimento", "Nesses termos pede deferimento.", "Nesses termos, pede deferimento.",
    "Nesses termos pede deferimento", "Nesses termos, pede deferimento", "Nesses termos pede e aguarda deferimento",
    "Nesses termos, pede e aguarda deferimento", "Nesses termos pede e aguarda deferimento."
]

TERMO_DECISAO = "DECISÃO"
TERMO_DESPACHO = "DESPACHO"
TERMO_OFICIO = "OFÍCIO"
TERMO_TRIBUNAL = "TRIBUNAL DE JUSTIÇA DO ESTADO DE SÃO PAULO"

TERMO_JUIZ = [
    "Juiz(a) de Direito: Dr(a)", "Juiz de Direito:", "Juíza de Direito:",
    "Juiz(a) de Direito", "Juiz de Direito", "Juíza de Direito"
]


def localizar_paginas(doc, termos, termos_adicionais=None, ultima_ocorrencia=False):
    """
    Localiza páginas em um documento PDF com base em termos principais e adicionais.

    Args:
        doc: Documento PDF carregado com fitz.
        termos: Lista de termos principais a serem buscados.
        termos_adicionais: Lista de termos adicionais a serem buscados (opcional).
        ultima_ocorrencia: Se True, retorna apenas a última ocorrência encontrada.

    Returns:
        Lista de índices das páginas encontradas ou a última página se ultima_ocorrencia for True.
    """

    paginas_encontradas = []

    for i, pagina in enumerate(doc):
        texto = pagina.get_text()

        if any(termo in texto for termo in termos):
            if termos_adicionais is None or any(termo in texto for termo in termos_adicionais):
                paginas_encontradas.append(i)

    logger.info(
        "Código 4 - Busca de páginas concluída. %s página(s) correspondente(s) encontrada(s).",
        len(paginas_encontradas)
    )

    return [paginas_encontradas[-1]] if ultima_ocorrencia and paginas_encontradas else paginas_encontradas


def processar_com_pymupdf(caminho_arquivo, nome_base):
    """Processa o PDF usando PyMuPDF (fitz) como método principal."""

    logger.info(
        "Código 4 - Iniciando processamento com PyMuPDF: %s",
        nome_base
    )

    doc = fitz.open(caminho_arquivo)

    logger.info(
        "Código 4 - PDF '%s' aberto com PyMuPDF. Total de páginas: %s.",
        nome_base,
        len(doc)
    )

    try:
        # Localizar páginas da petição
        paginas_peticao = []

        for i, pagina in enumerate(doc):
            texto = pagina.get_text()

            if any(termo in texto for termo in TERMINOS_PETICAO):
                paginas_peticao = list(range(0, i + 1))

                logger.info(
                    "Código 4 - '%s': término da petição identificado na página %s.",
                    nome_base,
                    i + 1
                )

                break

        if not paginas_peticao:
            paginas_peticao = list(range(0, min(15, len(doc))))

            logger.warning(
                "Código 4 - '%s': nenhum termo de petição encontrado. Utilizando as primeiras %s páginas.",
                nome_base,
                len(paginas_peticao)
            )

        # Localizar decisão/despacho/ofício
        paginas_decisao = localizar_paginas(doc, [TERMO_DECISAO], TERMO_JUIZ, ultima_ocorrencia=True)

        if paginas_decisao:
            logger.info(
                "Código 4 - '%s': decisão localizada na página %s.",
                nome_base,
                paginas_decisao[-1] + 1
            )

        if not paginas_decisao:
            paginas_decisao = localizar_paginas(doc, [TERMO_DESPACHO], TERMO_JUIZ, ultima_ocorrencia=True)

            if paginas_decisao:
                logger.info(
                    "Código 4 - '%s': despacho localizado na página %s.",
                    nome_base,
                    paginas_decisao[-1] + 1
                )

        if not paginas_decisao:
            paginas_decisao = localizar_paginas(doc, [TERMO_OFICIO], ultima_ocorrencia=True)

            if paginas_decisao:
                logger.info(
                    "Código 4 - '%s': ofício localizado na página %s.",
                    nome_base,
                    paginas_decisao[-1] + 1
                )

        if not paginas_decisao:
            logger.warning(
                "Código 4 - '%s': nenhuma decisão, despacho ou ofício foi localizado.",
                nome_base
            )

        # Criar documentos separados
        doc_peticao = fitz.open()
        doc_decisao = fitz.open()

        for p in paginas_peticao:
            doc_peticao.insert_pdf(doc, from_page=p, to_page=p)

        for p in paginas_decisao:
            doc_decisao.insert_pdf(doc, from_page=p, to_page=p)

        logger.info(
            "Código 4 - '%s': documento de petição possui %s página(s) e documento de decisão possui %s página(s).",
            nome_base,
            len(doc_peticao),
            len(doc_decisao)
        )

        # Combinar resultados
        doc_final = fitz.open()

        if len(doc_peticao) > 0:
            doc_final.insert_pdf(doc_peticao)

        if len(doc_decisao) > 0:
            doc_final.insert_pdf(doc_decisao)

        if len(doc_final) > 0:
            caminho_saida = os.path.join(PASTA_SAIDA, f"{nome_base}.pdf")
            doc_final.save(caminho_saida)

            logger.info(
                "Código 4 - '%s' processado com sucesso pelo PyMuPDF. Arquivo salvo em: %s",
                nome_base,
                caminho_saida
            )

            print(f"Processado com PyMuPDF: {caminho_saida}")

            return True

        else:
            raise Exception("Nenhuma página válida encontrada.")

    except Exception as e:
        logger.exception(
            "Código 4 - Erro no processamento principal com PyMuPDF para '%s'.",
            nome_base
        )

        print(f"Erro no processamento principal: {e}")

        return False

    finally:
        doc.close()

        if 'doc_peticao' in locals():
            doc_peticao.close()

        if 'doc_decisao' in locals():
            doc_decisao.close()

        if 'doc_final' in locals():
            doc_final.close()

        logger.info(
            "Código 4 - Recursos do PyMuPDF fechados para '%s'.",
            nome_base
        )


def processar_com_metodo_alternativo(caminho_arquivo, nome_base):
    """Método alternativo usando PyPDF2 + pdfplumber quando o principal falha."""

    logger.warning(
        "Código 4 - Iniciando método alternativo com PyPDF2 + pdfplumber para '%s'.",
        nome_base
    )

    print(f"⚠️ Tentando método alternativo para: {nome_base}")

    try:
        paginas_peticao = []
        paginas_decisao = []

        with pdfplumber.open(caminho_arquivo) as pdf:

            logger.info(
                "Código 4 - '%s' aberto com pdfplumber. Total de páginas: %s.",
                nome_base,
                len(pdf.pages)
            )

            for i, page in enumerate(pdf.pages):
                texto = page.extract_text() or ""

                if not paginas_peticao and any(t in texto for t in TERMINOS_PETICAO):
                    paginas_peticao = list(range(0, i + 1))

                    logger.info(
                        "Código 4 - '%s': término da petição encontrado pelo método alternativo na página %s.",
                        nome_base,
                        i + 1
                    )

                if any(t in texto for t in [TERMO_DECISAO, TERMO_DESPACHO, TERMO_OFICIO]):
                    if any(j in texto for j in TERMO_JUIZ):
                        paginas_decisao = [i]  # Mantém apenas a última

                        logger.info(
                            "Código 4 - '%s': possível decisão/despacho/ofício encontrada pelo método alternativo na página %s.",
                            nome_base,
                            i + 1
                        )

        if not paginas_peticao:
            paginas_peticao = list(range(0, min(15, len(pdf.pages))))

            logger.warning(
                "Código 4 - '%s': nenhum termo de petição encontrado pelo método alternativo. Utilizando as primeiras %s páginas.",
                nome_base,
                len(paginas_peticao)
            )

        reader = PyPDF2.PdfReader(caminho_arquivo)
        writer = PyPDF2.PdfWriter()

        for i in paginas_peticao:
            writer.add_page(reader.pages[i])

        for i in paginas_decisao:
            writer.add_page(reader.pages[i])

        if writer.pages:
            caminho_saida = os.path.join(PASTA_SAIDA, f"{nome_base}.pdf")

            with open(caminho_saida, "wb") as f_out:
                writer.write(f_out)

            logger.info(
                "Código 4 - '%s' processado com sucesso pelo método alternativo. Arquivo salvo em: %s",
                nome_base,
                caminho_saida
            )

            print(f"✅ Processado com método alternativo: {caminho_saida}")

            return True

        else:
            raise Exception("Nenhuma página válida encontrada no método alternativo.")

    except Exception as e:
        logger.exception(
            "Código 4 - Erro no método alternativo para '%s'.",
            nome_base
        )

        print(f"Erro no método alternativo: {e}")

        return False


def mover_para_lixeira(caminho_arquivo):
    """Tenta mover para a lixeira, se falhar remove permanentemente."""

    logger.info(
        "Código 4 - Removendo arquivo original após processamento: %s",
        caminho_arquivo
    )

    try:
        # Tentativa de mover para lixeira (Windows)
        if os.name == 'nt':
            import send2trash
            time.sleep(3)
            send2trash.send2trash(caminho_arquivo)

            logger.info(
                "Código 4 - Arquivo enviado para lixeira: %s",
                caminho_arquivo
            )

            print(f"Enviado para lixeira: {caminho_arquivo}")

        else:
            # Em outros sistemas, remove permanentemente
            os.remove(caminho_arquivo)

            logger.info(
                "Código 4 - Arquivo removido permanentemente: %s",
                caminho_arquivo
            )

            print(f"Removido permanentemente: {caminho_arquivo}")

    except Exception as e:
        logger.exception(
            "Código 4 - Falha ao mover '%s' para a lixeira. Será realizada tentativa de remoção permanente.",
            caminho_arquivo
        )

        print(f"Falha ao mover para lixeira, removendo permanentemente: {e}")

        os.remove(caminho_arquivo)

        logger.warning(
            "Código 4 - Arquivo removido permanentemente após falha ao mover para a lixeira: %s",
            caminho_arquivo
        )


def main():
    logger.info("Código 4 - Iniciando função principal de processamento dos PDFs.")

    falhas = []
    sucessos = 0

    arquivos_pasta = os.listdir(PASTA_ORIGEM)

    logger.info(
        "Código 4 - %s item(ns) encontrado(s) na pasta de origem.",
        len(arquivos_pasta)
    )

    for arquivo in arquivos_pasta:
        if not arquivo.endswith(".pdf"):
            continue

        caminho_arquivo = os.path.join(PASTA_ORIGEM, arquivo)
        nome_base = Path(arquivo).stem

        logger.info(
            "Código 4 - Iniciando processamento do arquivo: %s",
            arquivo
        )

        try:
            # Tentar primeiro com PyMuPDF
            sucesso = processar_com_pymupdf(caminho_arquivo, nome_base)

            if not sucesso:

                logger.warning(
                    "Código 4 - Método principal falhou para '%s'. Iniciando método alternativo.",
                    arquivo
                )

                # Se falhar, tentar método alternativo
                sucesso = processar_com_metodo_alternativo(caminho_arquivo, nome_base)

                if sucesso:
                    falhas.append((arquivo, "Método alternativo"))

                    logger.warning(
                        "Código 4 - '%s' foi processado utilizando o método alternativo.",
                        arquivo
                    )

                else:
                    falhas.append((arquivo, "Falha total"))

                    logger.error(
                        "Código 4 - Falha total ao processar '%s'.",
                        arquivo
                    )

            else:
                sucessos += 1

                logger.info(
                    "Código 4 - '%s' processado com sucesso pelo método principal.",
                    arquivo
                )

            # Mover arquivo original para lixeira/apagar
            mover_para_lixeira(caminho_arquivo)

        except Exception as e:
            logger.exception(
                "Código 4 - Erro inesperado ao processar '%s'.",
                arquivo
            )

            print(f"Erro inesperado ao processar {arquivo}: {e}")

            falhas.append((arquivo, "Erro inesperado"))

    # Relatório final
    print("\n=== RELATÓRIO FINAL ===")
    print(f"✅ {sucessos} arquivos processados com sucesso pelo método principal")

    logger.info(
        "Código 4 - Relatório final: %s arquivo(s) processado(s) com sucesso pelo método principal.",
        sucessos
    )

    if falhas:
        print("\n⚠️ Arquivos com problemas:")

        logger.warning(
            "Código 4 - %s arquivo(s) apresentaram problemas ou utilizaram método alternativo.",
            len(falhas)
        )

        for arquivo, motivo in falhas:
            print(f" - {arquivo}: {motivo}")

            logger.warning(
                "Código 4 - Arquivo: %s | Resultado: %s",
                arquivo,
                motivo
            )

    else:
        print("\n🎉 Todos os arquivos foram processados com sucesso!")

        logger.info(
            "Código 4 - Todos os arquivos foram processados com sucesso."
        )


if __name__ == "__main__":
    try:
        main()

        logger.info(
            "Código 4 - Finalizado."
        )

    except Exception:
        logger.exception(
            "Código 4 - ERRO não tratado durante a divisão dos PDFs."
        )
        raise


###

#FIM CÓDIGO 4

###


###_________________________________________________________________________________________
##Código 4.1, injetar os arquivos dos segredos de justiça dentro da pasta "docs processados"

#Implementado em 16/07/2026
###_________________________________________________________________________________________

logger.info("Código 4.1 - Iniciando cópia dos processos da pasta de segredo de justiça.")

print("Copiando processos da pasta de segredo de justiça (caso existam). Aguarde...")

pasta_origem = "segredo de justica"
pasta_destino = "docs processados"

# Cria as pastas caso ainda não existam
os.makedirs(pasta_origem, exist_ok=True)
os.makedirs(pasta_destino, exist_ok=True)

logger.info(
    "Código 4.1 - Pastas verificadas/criadas. Origem: '%s' | Destino: '%s'.",
    pasta_origem,
    pasta_destino
)

arquivos_copiados = 0

try:
    for nome_arquivo in os.listdir(pasta_origem):
        caminho_origem = os.path.join(pasta_origem, nome_arquivo)
        caminho_destino = os.path.join(pasta_destino, nome_arquivo)

        # Copia somente arquivos PDF
        if os.path.isfile(caminho_origem) and nome_arquivo.lower().endswith(".pdf"):
            logger.info(
                "Código 4.1 - Copiando documento de segredo de justiça: %s",
                nome_arquivo
            )

            shutil.copy2(caminho_origem, caminho_destino)

            logger.info(
                "Código 4.1 - Documento copiado com sucesso: %s",
                nome_arquivo
            )

            print(f"Documento de segredo de justiça copiado: {nome_arquivo}")
            arquivos_copiados += 1

    if arquivos_copiados:
        logger.info(
            "Código 4.1 - %s documento(s) de segredo de justiça copiado(s) para a pasta 'docs processados'.",
            arquivos_copiados
        )

        print(
            f"{arquivos_copiados} documento(s) de segredo de justiça "
            f"copiado(s) para a pasta 'docs processados'."
        )
    else:
        logger.info(
            "Código 4.1 - Nenhum documento de segredo de justiça foi encontrado."
        )

        print("Nenhum documento de segredo de justiça foi encontrado.")

    print("Segredos de justiça movidos!")

    logger.info("Código 4.1 - Finalizado.")

except Exception:
    logger.exception(
        "Código 4.1 - ERRO durante a cópia dos processos de segredo de justiça."
    )
    raise


###

#FIM CÓDIGO 4.1

###

########################################
#Código 5 autuar processos sistema SEI #
########################################

#Autuador de processos do sistema SEI COMPLETO:
    #Loga, autua, preenche os processos com documentos e depois os envia para o setor correto.
    #Só falta testar o envio de processos ao setor de procuradoria, pois o botão de enviar está comentado.

# ----------------------------------------------------------------
# CONFIGURAÇÃO DOS LOGS
# ----------------------------------------------------------------

class FormatoRoxo(logging.Formatter):
    def format(self, record):
        mensagem = super().format(record)
        return f"\033[35m{mensagem}\033[0m"

#Inseri essa classe para ter a barra de progresso do TDQM
class HandlerTerminalTqdm(logging.StreamHandler):
    def emit(self, record):
        try:
            mensagem = self.format(record)
            tqdm.write(mensagem, file=self.stream)
        except Exception:
            self.handleError(record)


pasta_logs = Path("logs")
pasta_logs.mkdir(exist_ok=True)

data_hora = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
arquivo_log = pasta_logs / f"execucao_{data_hora}.log"

formato_log = (
    "%(asctime)s | %(levelname)s | "
    "%(filename)s:%(lineno)d | %(message)s"
)

# Arquivo de log com texto normal
handler_arquivo = logging.FileHandler(
    arquivo_log,
    encoding="utf-8"
)
handler_arquivo.setFormatter(logging.Formatter(formato_log))

# Terminal com texto roxo
#Substituirei essa linha abaixo funcional, para ter o controle da quantidade de processos que ainda precisa
#Inserir documentos usando o tdqm

#handler_terminal = logging.StreamHandler()
handler_terminal = HandlerTerminalTqdm()

handler_terminal.setFormatter(FormatoRoxo(formato_log))

logging.basicConfig(
    level=logging.INFO,
    handlers=[
        handler_arquivo,
        handler_terminal
    ]
)

logger = logging.getLogger(__name__)

logger.info("============================================================")
logger.info("Etapa 0, login sendo realizado.")
logger.info("============================================================")

#
# Código de login SEI
#

# ----------------------------------------------------------------
# CONFIGURAÇÕES DO SELENIUM
# ----------------------------------------------------------------

TEMPO_ESPERA = 5

# https://cidades.sei.sp.gov.br/rasaopaulo/sip/login.php?
LINK_DIRETO_SEI = (
    "https://cidades.sei.sp.gov.br/rasaopaulo/sip/login.php?"
    "sigla_orgao_sistema=RASAOPAULO&"
    "sigla_sistema=SEI&"
    "infra_url=L3Jhc2FvcGF1bG8vc2VpLw=="
)

LINK_ALTERNATIVO_SEI = (
    "https://cidades.sei.sp.gov.br/rasaopaulo/sip/login.php?sigla_orgao_sistema=RASAOPAULO&sigla_sistema=SEI&infra_url=L3Jhc2FvcGF1bG8vc2VpLw=="
)


# ----------------------------------------------------------------
# LER DADOS DE LOGIN
# ----------------------------------------------------------------

def carregar_login_sei():

    logger.info(
        "Login SEI - Carregando dados do arquivo login_SEI_cidades.txt."
    )

    with open(
        "login_SEI_cidades.txt",
        "r",
        encoding="utf-8"
    ) as arquivo:

        linhas = arquivo.readlines()

    usuario = linhas[0].strip()
    senha = linhas[1].strip()
    orgao = linhas[2].strip()

    logger.info(
        "Login SEI - Dados de login carregados com sucesso."
    )

    return usuario, senha, orgao


# ----------------------------------------------------------------
# PREENCHER LOGIN DO SEI
# ----------------------------------------------------------------

def preencher_login_sei(driver):

    logger.info(
        "Login SEI - Iniciando preenchimento da tela de login."
    )

    usuario, senha, orgao = carregar_login_sei()

    espera = WebDriverWait(driver, TEMPO_ESPERA)

    # Campo usuário
    logger.info(
        "Login SEI - Aguardando campo de usuário."
    )

    campo_usuario = espera.until(
        EC.visibility_of_element_located(
            (By.XPATH, '//*[@id="txtUsuario"]')
        )
    )

    campo_usuario.clear()
    campo_usuario.send_keys(usuario)

    logger.info(
        "Login SEI - Campo de usuário preenchido."
    )


    # Campo senha
    logger.info(
        "Login SEI - Aguardando campo de senha."
    )

    campo_senha = espera.until(
        EC.visibility_of_element_located(
            (By.XPATH, '//*[@id="pwdSenha"]')
        )
    )

    campo_senha.clear()
    campo_senha.send_keys(senha)

    logger.info(
        "Login SEI - Campo de senha preenchido."
    )


    # Seleção do órgão
    logger.info(
        "Login SEI - Aguardando lista de órgãos."
    )

    campo_orgao = espera.until(
        EC.presence_of_element_located(
            (By.XPATH, '//*[@id="selOrgao"]')
        )
    )

    Select(campo_orgao).select_by_visible_text(orgao)

    logger.info(
        "Login SEI - Órgão selecionado."
    )


    # Botão acessar
    logger.info(
        "Login SEI - Aguardando botão de acesso."
    )

    botao_acessar = espera.until(
        EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="sbmAcessar"]')
        )
    )

    botao_acessar.click()

    logger.info(
        "Login SEI - Botão de acesso clicado."
    )

    print("Dados de login inseridos. Acessando o sistema SEI.")


# ----------------------------------------------------------------
# CASO 1 - TENTAR LINK DIRETO
# ----------------------------------------------------------------

def tentar_link_direto(driver):

    logger.info(
        "Login SEI - Tentando acessar o sistema pelo link direto."
    )

    try:

        driver.get(LINK_DIRETO_SEI)

        logger.info(
            "Login SEI - Link direto carregado. "
            "Verificando se a tela de login está disponível."
        )

        WebDriverWait(driver, TEMPO_ESPERA).until(
            EC.visibility_of_element_located(
                (By.XPATH, '//*[@id="txtUsuario"]')
            )
        )

        print(
            "Link PRINCIPAL do sistema SEI funcionando, "
            "iniciando o processo de login no sistema"
        )

        logger.info(
            "Login SEI - Link direto funcionando."
        )

        return True

    except (TimeoutException, WebDriverException):

        logger.warning(
            "Login SEI - Link direto indisponível ou "
            "tela de login não encontrada."
        )

        print(
            "Link direto do SEI cidades indisponível no momento, "
            "tentando acessar por outro método"
        )

        return False


# ----------------------------------------------------------------
# CONDICIONAL DO CASO 1 - MÉTODO ALTERNATIVO
# ----------------------------------------------------------------

def tentar_link_alternativo(driver):

    logger.info(
        "Login SEI - Tentando acessar o sistema pelo link alternativo."
    )

    try:

        driver.get(LINK_ALTERNATIVO_SEI)

        logger.info(
            "Login SEI - Link alternativo carregado. "
            "Verificando se a tela de login está disponível."
        )

        WebDriverWait(driver, TEMPO_ESPERA).until(
            EC.visibility_of_element_located(
                (By.XPATH, '//*[@id="txtUsuario"]')
            )
        )

        print(
            "Link alternativo do sistema SEI funcionando, "
            "iniciando o processo de login no sistema"
        )

        logger.info(
            "Login SEI - Link alternativo funcionando."
        )

        return True

    except (TimeoutException, WebDriverException):

        logger.warning(
            "Login SEI - Link alternativo indisponível ou "
            "tela de login não encontrada."
        )

        print(
            "\nLink alternativo do SEI cidades indisponível no momento."
        )

        return False


# ----------------------------------------------------------------
# FLUXO PRINCIPAL DE LOGIN
# ----------------------------------------------------------------

def acessar_e_fazer_login_sei(driver):

    logger.info(
        "Login SEI - Iniciando fluxo de acesso ao sistema."
    )

    # Primeiro tenta o acesso direto
    acesso_direto = tentar_link_direto(driver)

    if acesso_direto:

        preencher_login_sei(driver)

        logger.info(
            "Login SEI - Fluxo de login pelo método direto concluído."
        )

        return


    # Caso o método direto tenha falhado
    logger.info(
        "Login SEI - Iniciando segunda alternativa de acesso."
    )

    acesso_alternativo = tentar_link_alternativo(driver)

    if acesso_alternativo:

        logger.info(
            "Login SEI - Retornando ao fluxo normal de login "
            "após acesso alternativo."
        )

        preencher_login_sei(driver)

        logger.info(
            "Login SEI - Fluxo de login pelo método alternativo concluído."
        )

        return


    # Ambos os métodos falharam
    logger.error(
        "Login SEI - Método direto e método alternativo falharam."
    )

    print(
        "\nMétodo de login PRINCIPAL e SECUNDÁRIO falharam, "
        "verificar se o sistema SEI está online ou com problemas de conexão "
        "com a internet.\n"
    )

    print("Tente novamente executar esse programa mais tarde.")

    raise RuntimeError(
        "Não foi possível acessar a tela de login do SEI."
    )


#ATÉ AQUI EM CIMA, TUDO OKAY


# ----------------------------------------------------------------
# ETAPA 1 - PREPARAR TELA APÓS LOGIN
# ----------------------------------------------------------------

def executar_etapa_1(driver):

    logger.info(
        "============================================================"
    )

    logger.info(
        "ETAPA 1 - CHEGAR ATÉ TELA DE INSERIR INFORMAÇÕES"
    )

    logger.info(
        "============================================================"
    )

    # ----------------------------------------------------------------
    # 1 - FECHAR POP-UP, CASO EXISTA
    # ----------------------------------------------------------------

    logger.info(
        "Etapa 1 - Verificando se existe POP-UP aberto."
    )

    try:

        botao_fechar_popup = WebDriverWait(
            driver,
            2
        ).until(
            EC.element_to_be_clickable(
                (
                    By.XPATH,
                    '//div[starts-with(@id, "divInfraSparklingModalClose")]/img'
                )
            )
        )

        logger.info(
            "Etapa 1 - POP-UP localizado."
        )

        botao_fechar_popup.click()

        logger.info(
            "Etapa 1 - POP-UP fechado com sucesso."
        )

    except TimeoutException:

        logger.info(
            "Etapa 1 - Nenhum POP-UP foi localizado."
        )


# ----------------------------------------------------------------
# ETAPA 2 - AUTUAÇÃO DE PROCESSO
# ----------------------------------------------------------------

def etapa_2_autuacao_de_processo(driver):

    espera = WebDriverWait(driver, TEMPO_ESPERA)

    logger.info("ETAPA 2 - CHEGAR ATÉ A TELA DE AUTUAÇÃO DE PROCESSO.")

    # ----------------------------------------------------------------
    # 2 - VERIFICAR MENU LATERAL E BOTÃO "INICIAR PROCESSO"
    # ----------------------------------------------------------------

    logger.info(
        'Etapa 2 - Verificando botão "Iniciar Processo".'
    )

    xpath_iniciar_processo = (
        '//*[@id="linkMenu17"]/span'
    )

    xpath_menu_lateral = (
        '//*[@id="lnkInfraMenuSistema"]/span'
    )

    try:

        botao_iniciar_processo = WebDriverWait(
            driver,
            3
        ).until(
            EC.element_to_be_clickable(
                (
                    By.XPATH,
                    xpath_iniciar_processo
                )
            )
        )

        logger.info(
            'Etapa 2 - Botão "Iniciar Processo" localizado.'
        )

    except TimeoutException:

        logger.warning(
            'Etapa 2 - Botão "Iniciar Processo" não apareceu. '
            "Tentando abrir o menu lateral."
        )

        try:

            botao_menu = WebDriverWait(
                driver,
                3
            ).until(
                EC.element_to_be_clickable(
                    (
                        By.XPATH,
                        xpath_menu_lateral
                    )
                )
            )

            logger.info(
                "Etapa 2 - Botão de menu localizado."
            )

            botao_menu.click()

            logger.info(
                "Etapa 2 - Botão de menu clicado."
            )

        except TimeoutException:

            logger.error(
                'Etapa 2 - Nem o botão "Menu" foi localizado.'
            )

            print(
                'Tanto o botão de "menu" como o de '
                '"Iniciar Processo" não foram localizados.'
            )

            raise RuntimeError(
                'Botões "Menu" e "Iniciar Processo" não localizados.'
            )

        try:

            botao_iniciar_processo = WebDriverWait(
                driver,
                TEMPO_ESPERA
            ).until(
                EC.element_to_be_clickable(
                    (
                        By.XPATH,
                        xpath_iniciar_processo
                    )
                )
            )

            logger.info(
                'Etapa 2 - Botão "Iniciar Processo" localizado '
                "após abrir o menu lateral."
            )

        except TimeoutException:

            logger.error(
                'Etapa 2 - Botão "Iniciar Processo" não foi localizado '
                "mesmo após abrir o menu."
            )

            print(
                'Tanto o botão de "menu" como o de '
                '"Iniciar Processo" não foram localizados.'
            )

            raise RuntimeError(
                'Botão "Iniciar Processo" não localizado.'
            )








    # ----------------------------------------------------------------
    # 3 - CLICAR EM "INICIAR PROCESSO"
    # ----------------------------------------------------------------

    logger.info(
        'Etapa 2 - Clicando em "Iniciar Processo".'
    )

    botao_iniciar_processo.click()

    logger.info(
        'Etapa 2 - Botão "Iniciar Processo" clicado.'
    )


    # ----------------------------------------------------------------
    # 4 - LOCALIZAR A BARRA DE PESQUISA
    # ----------------------------------------------------------------

    logger.info(
        "Etapa 2 - Procurando barra de pesquisa de tipos de processo."
    )

    campo_filtro = espera.until(
        EC.visibility_of_element_located(
            (
                By.XPATH,
                '//*[@id="txtFiltro"]'
            )
        )
    )

    logger.info(
        "Etapa 2 - Barra de pesquisa localizada."
    )


    # ----------------------------------------------------------------
    # 5 - TIPOS DE PROCESSO ACEITOS
    # ----------------------------------------------------------------

    tipos_processo = [
        (
            "ADMINISTRAÇÃO: Expediente de acompanhamento "
            "de ação judicial"
        ),
        (
            "ADMINISTRAÇÃO: Expediente de atendimento de "
            "requisições, ordens judiciais ou administrativas"
        )
    ]


    # ----------------------------------------------------------------
    # 6 - FUNÇÃO INTERNA PARA PROCURAR OS DOIS TIPOS
    # ----------------------------------------------------------------

    def procurar_tipo_processo():

        for tipo_processo in tipos_processo:

            logger.info(
                "Etapa 2 - Pesquisando tipo de processo: %s",
                tipo_processo
            )

            campo_filtro.clear()

            campo_filtro.send_keys(
                tipo_processo
            )

            try:

                resultado = WebDriverWait(
                    driver,
                    3
                ).until(
                    EC.element_to_be_clickable(
                        (
                            By.XPATH,
                            f'//*[normalize-space()="{tipo_processo}"]'
                        )
                    )
                )

                logger.info(
                    "Etapa 2 - Tipo de processo localizado: %s",
                    tipo_processo
                )

                return resultado, tipo_processo

            except TimeoutException:

                logger.info(
                    "Etapa 2 - Tipo de processo não localizado: %s",
                    tipo_processo
                )

        return None, None


    # ----------------------------------------------------------------
    # 7 - PRIMEIRA TENTATIVA DE LOCALIZAR OS TIPOS
    # ----------------------------------------------------------------

    logger.info(
        "Etapa 2 - Realizando primeira tentativa de localizar "
        "um dos tipos de processo."
    )

    resultado, tipo_encontrado = procurar_tipo_processo()

    if resultado is not None:

        logger.info(
            "Etapa 2 - Tipo de processo encontrado na primeira tentativa."
        )

        resultado.click()

        logger.info(
            "Etapa 2 - Tipo de processo selecionado: %s",
            tipo_encontrado
        )

        print(
            "Tipo de processo localizado e selecionado."
        )

        return


    # ----------------------------------------------------------------
    # 8 - SE NÃO ENCONTROU NENHUM DOS DOIS, CLICAR NO BOTÃO DE EXIBIÇÃO
    # ----------------------------------------------------------------

    logger.warning(
        "Etapa 2 - Nenhum dos dois tipos de processo foi localizado "
        "na primeira tentativa."
    )

    logger.info(
        "Etapa 2 - Clicando no botão de exibição "
        "dos tipos de processo."
    )

    botao_exibir_tipos = espera.until(
        EC.element_to_be_clickable(
            (
                By.XPATH,
                '//*[@id="ancExibirTiposProcedimento"]/img'
            )
        )
    )

    botao_exibir_tipos.click()

    logger.info(
        "Etapa 2 - Botão de exibição dos tipos de processo clicado."
    )


    # ----------------------------------------------------------------
    # 9 - LOCALIZAR NOVAMENTE A BARRA DE PESQUISA
    # ----------------------------------------------------------------

    campo_filtro = espera.until(
        EC.visibility_of_element_located(
            (
                By.XPATH,
                '//*[@id="txtFiltro"]'
            )
        )
    )


    # ----------------------------------------------------------------
    # 10 - SEGUNDA TENTATIVA
    # ----------------------------------------------------------------

    logger.info(
        "Etapa 2 - Realizando segunda tentativa de localizar "
        "um dos tipos de processo."
    )

    resultado, tipo_encontrado = procurar_tipo_processo()

    if resultado is not None:

        logger.info(
            "Etapa 2 - Tipo de processo localizado após clicar "
            "no botão de exibição."
        )

        resultado.click()

        logger.info(
            "Etapa 2 - Tipo de processo selecionado: %s",
            tipo_encontrado
        )

        print(
            "Tipo de processo localizado e selecionado."
        )

        return


    # ----------------------------------------------------------------
    # 11 - NENHUM DOS DOIS TIPOS FOI LOCALIZADO
    # ----------------------------------------------------------------

    logger.error(
        "Etapa 2 - Nenhum dos dois tipos de processo foi localizado "
        "mesmo após clicar no botão de exibição."
    )

    print(
        "Não foi possível localizar nenhum dos dois tipos de processo\n"
        "e por isso o programa será encerrado."
    )

    raise RuntimeError(
        "Nenhum dos tipos de processo esperados foi localizado."
    )



"""Etapa 3 — preencher, salvar, registrar o PA e repetir para as demais linhas.

"""

# ----------------------------------------------------------------
# CONFIGURAÇÕES E DICIONÁRIOS — ALTERAÇÕES FUTURAS FICAM AQUI
# ----------------------------------------------------------------

CAMINHO_PLANILHA = Path(__file__).resolve().parent / "dados_para_autuar_processos.xlsx"
NOME_ABA = "CAPTACOES"
TEMPO_ESPERA_ETAPA_3 = 15
TEMPO_APOS_SALVAR = 60

ORGAOS = {
    "TJSP": "Tribunal de Justiça do Estado de São Paulo",
    "TRT 2": "JUSTIÇA DO TRABALHO TRIBUNAL REGIONAL DO TRABALHO 2° REGIÃO",
    "DEFENS. PÚBLICA de SP": "DEFENSORIA PÚBLICA DO ESTADO DE SÃO PAULO",
}

# Ordem de preferência definida no documento.
HIPOTESES = (
    ("Segredo de justiça no processo civil (Artigo 189 da Lei Federal nº 13.105/2015)", False),
    ("Segredo de justiça no processo penal (Art. 201,§ 6º, do Decreto-Lei Federal nº 3.689/1941)", False),
    ("Segredo de justiça no processo penal (Art. 201", True),
    ("Informação pessoal (Artigo 31 da Lei Federal nº 12527/2011)", False),
    ("Materiais e sistemas sigiloso (Artigo 45 do Decreto Federal nº 7845/2012)", False),
)

XPATH_ESPECIFICACAO = '//*[@id="txtDescricao"]'
XPATH_INTERESSADO = '//*[@id="txtInteressadoProcedimento"]'
XPATH_OBSERVACOES = '//*[@id="txaObservacoes"]'
XPATH_PUBLICO = '//*[@id="divOptPublico"]/div/label'
XPATH_RESTRITO = '//*[@id="divOptRestrito"]/div/label'
XPATH_SALVAR = '//*[@id="btnSalvar"]'
XPATH_CONTROLE = '//*[@id="lnkControleProcessos"]/img'
# O ID span2648362 do exemplo pode mudar a cada novo processo.
XPATH_CANDIDATOS_PA = '//a/span | //span[starts-with(@id, "span")]'
PADRAO_PA = re.compile(r"[0-9]{7}\.[0-9]{3}\.[0-9]{8}/[0-9]{4}-[0-9]{2}")


def normalizar(valor):
    texto = "" if valor is None else str(valor)
    texto = unicodedata.normalize("NFKD", texto)
    return " ".join("".join(c for c in texto if not unicodedata.combining(c)).upper().split())


# ----------------------------------------------------------------
# 3.1 - LER E VALIDAR AS LINHAS ANTES DE CRIAR PROCESSOS
# ----------------------------------------------------------------

def ler_linhas_pendentes(aba, logger):
    registros = []
    mapa = {normalizar(sigla): nome for sigla, nome in ORGAOS.items()}
    for linha in range(2, aba.max_row + 2):
        valor_a = aba.cell(linha, 1).value
        if valor_a is None or str(valor_a).strip() == "":
            logger.info("Etapa 3 - A%s vazia: fim dos registros.", linha)
            break
        # Evita recriar processos já registrados em execuções anteriores.
        pa_existente = aba.cell(linha, 9).value
        if pa_existente is not None and str(pa_existente).strip():
            if not PADRAO_PA.fullmatch(str(pa_existente).strip()):
                raise ValueError(f"I{linha} contém um PA fora do padrão esperado; confira antes de continuar.")
            logger.info("Etapa 3 - Linha %s já possui PA em I; não será autuada novamente.", linha)
            continue
        for coluna in (1, 3, 8, 9, 10, 11):
            if aba.cell(linha, coluna).data_type == "f":
                raise ValueError(f"{aba.cell(linha, coluna).coordinate}: use valor, não fórmula.")
        orgao = mapa.get(normalizar(aba.cell(linha, 3).value))
        if orgao is None:
            raise ValueError(f"C{linha}: órgão não cadastrado no dicionário ORGAOS.")
        sigilo = normalizar(aba.cell(linha, 10).value)
        if sigilo not in ("SIM", "NAO"):
            raise ValueError(f"J{linha}: informe SIM ou NAO.")
        especificacao = aba.cell(linha, 11).value
        if especificacao is None or not str(especificacao).strip():
            raise ValueError(f"K{linha}: curta identificação vazia. Execute a preparação primeiro.")
        especificacao = str(especificacao)
        if len(especificacao) > 99:
            raise ValueError(f"K{linha}: identificação excede 99 caracteres. Execute a preparação primeiro.")
        observacoes = aba.cell(linha, 8).value
        registros.append({
            "linha": linha, "orgao": orgao, "restrito": sigilo == "SIM",
            "especificacao": especificacao,
            "observacoes": "" if observacoes is None else str(observacoes),
        })
    return registros


# ----------------------------------------------------------------
# 3.2 - LOCALIZAR ELEMENTOS NA PÁGINA OU NOS FRAMES DO SEI
# ----------------------------------------------------------------
def buscar_nos_frames(driver, procurar, profundidade=0):
    resultado = procurar(driver)

    if resultado:
        return resultado

    if profundidade >= 5:
        return False

    frames = driver.find_elements(
        By.CSS_SELECTOR,
        "iframe, frame",
    )

    for frame in frames:
        driver.switch_to.frame(frame)

        resultado = buscar_nos_frames(
            driver,
            procurar,
            profundidade + 1,
        )

        if resultado:
            # Mantém o contexto do frame onde encontrou o resultado.
            return resultado

        driver.switch_to.parent_frame()

    return False


# ----------------------------------------------------------------
# AGUARDAR A BUSCA E REPETIR EM CASO DE MUDANÇA DA PÁGINA
# ----------------------------------------------------------------

def esperar_nos_frames(
    driver,
    procurar,
    tempo=TEMPO_ESPERA_ETAPA_3,
):
    log_busca = logging.getLogger(__name__)
    aviso_registrado = False

    def registrar_aviso():
        nonlocal aviso_registrado

        # Evita repetir o mesmo aviso a cada tentativa.
        if not aviso_registrado:
            log_busca.warning(
                "Etapa 3 - Página ou frame alterado durante a busca. "
                "Tentando novamente com os elementos atualizados, "
                "dentro do limite de %s segundos.",
                tempo,
            )
            aviso_registrado = True

    def tentar(navegador):
        try:
            # Também fica dentro do try, pois a mudança de contexto
            # pode coincidir com uma atualização da página.
            navegador.switch_to.default_content()

            return buscar_nos_frames(
                navegador,
                procurar,
            )

        except (
            StaleElementReferenceException,
            NoSuchFrameException,
        ):
            registrar_aviso()
            return False

        except WebDriverException as erro:
            mensagem = str(erro).casefold()

            # Trata somente o erro observado no seu log.
            # Outros erros do navegador continuam sendo informados.
            if (
                "node with given id does not belong to the document"
                in mensagem
            ):
                registrar_aviso()
                return False

            raise

    return WebDriverWait(
        driver,
        tempo,
        poll_frequency=0.5,
    ).until(tentar)



def localizar(driver, xpath):
    def procurar(navegador):
        for elemento in navegador.find_elements(By.XPATH, xpath):
            if elemento.is_displayed() and elemento.is_enabled():
                return elemento
        return False
    return esperar_nos_frames(driver, procurar)


def preencher_campo(driver, xpath, texto):
    campo = localizar(driver, xpath)
    campo.clear()
    if texto:
        campo.send_keys(texto)
    if campo.get_attribute("value") != texto:
        raise RuntimeError(f"O campo {xpath} não manteve o texto completo informado.")
    return campo


# ----------------------------------------------------------------
# 3.3 - PREENCHER INTERESSADO: LISTAGEM E ALTERNATIVA COM ENTER
# ----------------------------------------------------------------

def preencher_interessado(driver, orgao, logger):
    preencher_campo(driver, XPATH_INTERESSADO, orgao)
    # Os nomes cadastrados não contêm aspas duplas.
    xpath_opcao = (
        f'//*[self::li or self::a or self::span or self::div or self::td]'
        f'[normalize-space(.)="{orgao}" and not(.//*[normalize-space(.)="{orgao}"])]'
    )
    try:
        opcao = WebDriverWait(driver, 4).until(
            EC.element_to_be_clickable((By.XPATH, xpath_opcao))
        )
    except TimeoutException:
        logger.warning("Etapa 3 - Interessado não encontrado na lista; usando Enter.")
        campo = preencher_campo(driver, XPATH_INTERESSADO, orgao)
        campo.send_keys(Keys.ENTER)
        logger.info("Etapa 3 - Enter enviado para adicionar interessado: %s", orgao)
    else:
        opcao.click()
        logger.info("Etapa 3 - Interessado selecionado na lista: %s", orgao)
    # O preenchimento seguinte só prossegue se o formulário continuar disponível.
    localizar(driver, XPATH_OBSERVACOES)


# ----------------------------------------------------------------
# 3.4 - DEFINIR NÍVEL DE ACESSO E HIPÓTESE LEGAL
# ----------------------------------------------------------------

def selecionar_hipotese_legal(driver, logger):
    def procurar(navegador):
        listas = [e for e in navegador.find_elements(By.TAG_NAME, "select")
                  if e.is_displayed() and e.is_enabled()]
        # Percorrer primeiro as preferências preserva a ordem do documento.
        for termo, parcial in HIPOTESES:
            for elemento in listas:
                seletor = Select(elemento)
                for indice, opcao in enumerate(seletor.options):
                    atual = normalizar(opcao.text)
                    desejado = normalizar(termo)
                    combina = atual.startswith(desejado) if parcial else atual == desejado
                    if combina and opcao.is_enabled():
                        return elemento, indice, opcao.text
        return False
    try:
        elemento, indice, texto = esperar_nos_frames(driver, procurar)
    except TimeoutException as erro:
        raise RuntimeError(
            "Nenhuma hipótese prevista foi encontrada em uma lista SELECT visível. "
            "Informe o XPath/HTML da lista se o SEI usar outro componente. "
            "O processo não será salvo sem selecionar a hipótese."
        ) from erro
    seletor = Select(elemento)
    seletor.select_by_index(indice)
    if normalizar(seletor.first_selected_option.text) != normalizar(texto):
        raise RuntimeError("A hipótese legal não permaneceu selecionada.")
    logger.info("Etapa 3 - Hipótese legal utilizada: %s", texto)


def definir_nivel_acesso(driver, restrito, logger):
    xpath = XPATH_RESTRITO if restrito else XPATH_PUBLICO
    localizar(driver, xpath).click()
    # Confere o input associado ao label, pelo atributo for ou contêiner.
    def selecionado(navegador):
        rotulos = navegador.find_elements(By.XPATH, xpath)
        if not rotulos:
            return False
        alvo = rotulos[0].get_attribute("for")
        entradas = (navegador.find_elements(By.ID, alvo) if alvo else
                    navegador.find_elements(By.XPATH, xpath + '/ancestor::div[2]//input[@type="radio"]'))
        return any(e.is_selected() for e in entradas)
    esperar_nos_frames(driver, selecionado)
    logger.info("Etapa 3 - Nível selecionado: %s", "Restrito" if restrito else "Público")
    if restrito:
        selecionar_hipotese_legal(driver, logger)


# ----------------------------------------------------------------
# 3.5 - SALVAR UMA ÚNICA VEZ E CAPTURAR O NÚMERO DO PA
# ----------------------------------------------------------------

# ----------------------------------------------------------------
# CAPTURAR O NÚMERO DO PA APÓS SALVAR
# ----------------------------------------------------------------

def capturar_numero_pa(driver):

    def procurar(navegador):
        encontrados = set()

        elementos = navegador.find_elements(
            By.XPATH,
            XPATH_CANDIDATOS_PA,
        )

        for elemento in elementos:
            if elemento.is_displayed():
                texto = elemento.text.strip()

                if PADRAO_PA.fullmatch(texto):
                    encontrados.add(texto)

        if len(encontrados) > 1:
            raise RuntimeError(
                "Mais de um número de PA visível. "
                "Confira o seletor da árvore do processo."
            )

        return next(iter(encontrados), False)

    return esperar_nos_frames(
        driver,
        procurar,
        tempo=TEMPO_APOS_SALVAR,
    )

def salvar_planilha(livro, caminho):
    temporario = None
    try:
        with tempfile.NamedTemporaryFile(
            dir=caminho.parent, prefix=".etapa3_", suffix=".xlsx", delete=False
        ) as arquivo:
            temporario = Path(arquivo.name)
        livro.save(temporario)
        os.replace(temporario, caminho)
        temporario = None
    finally:
        if temporario is not None:
            temporario.unlink(missing_ok=True)


# ----------------------------------------------------------------
# 3.6 - FLUXO PRINCIPAL: LINHA ATUAL, GRAVAÇÃO E PRÓXIMO PROCESSO
# ----------------------------------------------------------------

def etapa_3_registrar_informacoes(
    driver, executar_etapa_1, etapa_2_autuacao_de_processo,
    caminho_planilha=CAMINHO_PLANILHA, nome_aba=NOME_ABA, logger=None,
):
    """Entrar já na tela de autuação, depois da etapa 2.

    As funções das etapas 1 e 2 são recebidas como argumentos para reutilizar
    exatamente o fluxo existente. Exceções interrompem o lote, sem repetir
    automaticamente o clique em Salvar.
    """
    logger = logger or logging.getLogger(__name__)
    caminho = Path(caminho_planilha).resolve()
    livro = None
    linha_atual = None
    salvar_acionado = False
    pa = None
    quantidade = 0
    #Adicionarei 1 linha abaixo para usar o tdqm na autuação dos processos
    barra = None

    try:
        logger.info("============================================================")
        logger.info("ETAPA 3 - PREENCHIMENTO E PREPARAÇÃO DOS PROCESSOS")
        logger.info("============================================================")
        livro = load_workbook(caminho)
        aba = livro[nome_aba] if nome_aba is not None else livro.active
        registros = ler_linhas_pendentes(aba, logger)
        if not registros:
            logger.info("Etapa 3 - Nenhuma linha pendente.")
            return 0
        backups = caminho.parent / "backups"
        backups.mkdir(exist_ok=True)
        copia = backups / f"{caminho.stem}_antes_etapa3_{datetime.now():%Y-%m-%d_%H-%M-%S_%f}.xlsx"
        shutil.copy2(caminho, copia)
        logger.info("Etapa 3 - Backup criado: %s", copia)



        # Verifica a possibilidade de gravar ANTES da primeira autuação.

        # Vou inserir outro trecho aqui no MEIO para poder usar o TDQM

        # Verifica a possibilidade de gravar ANTES da primeira autuação.
        salvar_planilha(livro, caminho)

        barra = tqdm(
            total=len(registros),
            desc="ETAPA 3 - Contagem de processos autuados",
            unit="processo",
            dynamic_ncols=True,
            mininterval=0,
            leave=True,
            postfix={"faltam": len(registros)},
            bar_format=(
                "\033[31m"
                "{desc}: {percentage:3.0f}% |"
                "\033[32m{bar}\033[31m"
                "| {n_fmt}/{total_fmt} "
                "[decorrido: {elapsed} | estimativa: {remaining}]"
                "{postfix}"
                "\033[0m"
            ),
        )



        #Fim desse novo trecho para usar o TDQM na autuação de processos

        for indice, dados in enumerate(registros):




            linha_atual = dados["linha"]
            salvar_acionado = False
            pa = None
            logger.info("Etapa 3 - Iniciando preenchimento da linha %s.", linha_atual)
            preencher_campo(driver, XPATH_ESPECIFICACAO, dados["especificacao"])
            logger.info("Etapa 3 - Especificação preenchida a partir de K%s.", linha_atual)
            preencher_interessado(driver, dados["orgao"], logger)
            preencher_campo(driver, XPATH_OBSERVACOES, dados["observacoes"])
            logger.info("Etapa 3 - Observações preenchidas a partir de H%s.", linha_atual)
            definir_nivel_acesso(driver, dados["restrito"], logger)
            botao = localizar(driver, XPATH_SALVAR)
            logger.info("Etapa 3 - Acionando Salvar uma única vez, linha %s.", linha_atual)

            #Mesmo sendo esse arquivo uma cópia, estou com medinho de apagar o trecho abaixo
            #Então vou deixar comentado mesmo

            #salvar_acionado = True
            #botao.click()
            #pa = capturar_numero_pa(driver)

            #Já sabe né? pra usar o TDQM para autuação de processo,
            # preciso colocar um outro código logo antes dessas duas linhas
            # Posiciona o botão longe da barra superior da página.

            driver.execute_script(
                "arguments[0].scrollIntoView("
                "{block: 'center', inline: 'nearest', behavior: 'instant'});",
                botao,
            )

            logger.info(
                "Etapa 3 - Botão Salvar reposicionado para o clique."
            )



            #Fim desse outro código adicionado.
            salvar_acionado = True
            botao.click()

            # Atualiza a barra após o clique em Salvar.
            # Chat mandou eu fazer esse incremento mais para frente.
            #barra.set_postfix(
            #    faltam=len(registros) - (barra.n + 1),
            #    refresh=False,
            #)
            #barra.update(1)

            # Continua normalmente a captura do número do processo.
            pa = capturar_numero_pa(driver)


            #Fim desse novo trecho acima para usar o TDQM para registrar a atuação dos processos.

            logger.info("Etapa 3 - PA capturado para linha %s: %s", linha_atual, pa)
            celula = aba.cell(linha_atual, 9)
            celula.number_format = "@"
            celula.value = pa  # String literal: preserva pontos, barra e hífen.

            salvar_planilha(livro, caminho)
            salvar_acionado = False
            quantidade += 1
            logger.info("Etapa 3 - I%s salva em disco com PA %s.", linha_atual, pa)

            #Incremento da barra de autuação de processos

            barra.set_postfix(
                faltam=len(registros) - (barra.n + 1),
                refresh=False,
            )
            barra.update(1)

            #Fim do incremento da barra de autuação

            if indice + 1 < len(registros):
                logger.info("Etapa 3 - Voltando ao controle para preparar a próxima linha.")
                localizar(driver, XPATH_CONTROLE).click()
                driver.switch_to.default_content()
                executar_etapa_1(driver)
                etapa_2_autuacao_de_processo(driver)
        logger.info("Etapa 3 - Concluída. Processos criados e registrados: %s.", quantidade)
        return quantidade

    except Exception:
        logger.exception("Etapa 3 - ERRO na linha %s. Lote interrompido.", linha_atual)
        if salvar_acionado:
            logger.critical(
                "Etapa 3 - Salvar foi acionado, mas o registro local não foi confirmado. "
                "Antes de executar novamente, confira o SEI e preencha I%s com o PA "
                "se ele foi criado, para evitar duplicação. PA capturado: %s",
                linha_atual, pa or "não capturado",
            )
        raise
    #Só falta substituir agora esse Finally para usar o TDQM na parte de autuação de processos.

    finally:
        if barra is not None:
            barra.close()

        if livro is not None:
            livro.close()

    #finally:
    #    if livro is not None:
    #        livro.close()
    #Deixei o código funcional anterior comentado (como sempre)

# ----------------------------------------------------------------
# INTEGRAÇÃO NO ARQUIVO PRINCIPAL (NÃO EXECUTADA NESTE MÓDULO)
# ----------------------------------------------------------------
# 1. Adicione junto aos imports:
# from etapa_3_autuacao import etapa_3_registrar_informacoes
#
# 2. Dentro do try principal, logo depois de:
# etapa_2_autuacao_de_processo(driver)
#
# acrescente:
# etapa_3_registrar_informacoes(
#     driver,
#     executar_etapa_1=executar_etapa_1,
#     etapa_2_autuacao_de_processo=etapa_2_autuacao_de_processo,
#     caminho_planilha=Path(__file__).resolve().parent / "dados_para_autuar_processos.xlsx",
#     logger=logger,
# )
#
# 3. No except principal, ajuste a mensagem para mencionar a etapa 3 também.




"""ETAPA 4 - PREENCHIMENTO DE DOCUMENTOS NOS PROCESSOS.

Cole este conteúdo antes do bloco EXECUTAR do script principal, ou importe
esta função de um arquivo colocado na mesma pasta do script principal.
Chamada após a etapa 3:
    etapa_4_preenchimento_documentos_nos_processos(driver, logger=logger)

Dependências: selenium e openpyxl.
Usa o logger existente (incluindo sua cor roxa e arquivo de log).
Não abre outro navegador, não altera a planilha e não configura outro logger.
Planilha: aba ativa por padrão; passe nome_aba='...' para escolher outra.
Arquivos: pasta 'docs processados', ao lado deste .py, sem subpastas.
PJ: nome do arquivo sem extensão ou número CNJ presente no nome.
Não escolhe arbitrariamente entre vários arquivos correspondentes.

ATENÇÃO: não há controle persistente de documentos já inseridos. Ao retomar,
confira o último PA no SEI e use linha_inicial=... para evitar duplicação.
Se falhar depois de clicar em Salvar, não execute novamente essa linha sem
conferir o SEI. A confirmação usa um novo item na árvore contendo o nome
'Petição inicial do PJ'; esse critério depende do HTML real do seu SEI.
"""



def _e4_texto(valor):
    return '' if valor is None else str(valor).strip()


def _e4_normalizar(valor):
    texto = unicodedata.normalize('NFKD', _e4_texto(valor))
    return ' '.join(''.join(c for c in texto if not unicodedata.combining(c)).casefold().split())


def _e4_xpath_literal(texto):
    if "'" not in texto:
        return "'" + texto + "'"
    if '"' not in texto:
        return '"' + texto + '"'
    return 'concat(' + ', "\'", '.join("'" + p + "'" for p in texto.split("'")) + ')'


def _e4_arquivo(pasta, pj):
    arquivos = [p for p in pasta.iterdir() if p.is_file() and not p.name.startswith('~$')]
    exatos = [p for p in arquivos if _e4_normalizar(p.stem) == _e4_normalizar(pj)
              or _e4_normalizar(p.name) == _e4_normalizar(pj)]
    if exatos:
        encontrados = exatos
    else:
        # CNJ completo, com ou sem pontuação. Evita confundir números parciais.
        digitos = re.sub(r'\D', '', pj)
        if len(digitos) != 20:
            raise ValueError(f'PJ {pj!r}: use o nome exato do arquivo ou número CNJ completo.')
        padrao = r'(?<!\d)' + r'[.\-\s]*'.join(digitos) + r'(?!\d)'
        encontrados = [p for p in arquivos if re.search(padrao, p.stem)]
    if len(encontrados) != 1:
        raise ValueError(f'PJ {pj!r}: encontrados {len(encontrados)} arquivos: '
                         f'{[p.name for p in encontrados]}. Necessário exatamente um.')
    arquivo = encontrados[0].resolve()
    if arquivo.stat().st_size == 0:
        raise ValueError(f'Arquivo vazio: {arquivo.name}')
    return arquivo


class _E4Tela:
    """Busca também dentro de frames, comuns na árvore e formulários do SEI."""
    def __init__(self, driver, logger, timeout):
        self.driver, self.log, self.timeout = driver, logger, timeout

    def percorrer(self, verificar, profundidade=0):
        resultado = verificar()
        if resultado:
            return resultado
        if profundidade < 5:
            for frame in self.driver.find_elements(By.CSS_SELECTOR, 'iframe, frame'):
                self.driver.switch_to.frame(frame)
                resultado = self.percorrer(verificar, profundidade + 1)
                if resultado:
                    return resultado
                self.driver.switch_to.parent_frame()
        return False

    def esperar(self, verificar, timeout=None):
        def tentativa(_):
            self.driver.switch_to.default_content()
            try:
                return self.percorrer(verificar)
            except (StaleElementReferenceException, NoSuchFrameException):
                return False
        return WebDriverWait(self.driver, self.timeout if timeout is None else timeout).until(tentativa)

    def elemento(self, *xpaths, timeout=None, visivel=True):
        def verificar():
            for xpath in xpaths:
                for el in self.driver.find_elements(By.XPATH, xpath):
                    if el.is_enabled() and (not visivel or el.is_displayed()):
                        return el
            return False
        return self.esperar(verificar, timeout)

    def clicar(self, descricao, *xpaths):
        self.log.info('ETAPA 4 - %s.', descricao)
        self.elemento(*xpaths).click()

    def preencher(self, identificador, valor, descricao):
        self.log.info('ETAPA 4 - Preenchendo %s.', descricao)
        el = self.elemento(f'//*[@id={_e4_xpath_literal(identificador)}]')
        el.clear()
        el.send_keys(valor)
        el.send_keys(Keys.TAB)
        return el

    def radio(self, identificador, texto):
        xp = f'//*[@id={_e4_xpath_literal(identificador)}]'
        self.clicar(f'Selecionando {texto}', xp + '/div/label',
                    '//label[normalize-space(.)=' + _e4_xpath_literal(texto) + ']')
        # Quando o label aponta para input, conferir a seleção real.
        def selecionado():
            labels = self.driver.find_elements(By.XPATH, xp + '//label')
            labels += self.driver.find_elements(By.XPATH, '//label[normalize-space(.)=' + _e4_xpath_literal(texto) + ']')
            inputs = self.driver.find_elements(By.XPATH, xp + '//input[@type="radio"]')
            for label in labels:
                alvo = label.get_attribute('for')
                if alvo:
                    inputs += self.driver.find_elements(By.ID, alvo)
            return any(e.is_selected() for e in inputs)
        self.esperar(selecionado)

    def selecionar(self, xpath, prioridades, descricao, parcial=False):
        def selecionar_disponivel():
            for el in self.driver.find_elements(By.XPATH, xpath):
                if not el.is_displayed() or not el.is_enabled():
                    continue
                select = Select(el)
                for termo in prioridades:
                    for i, opcao in enumerate(select.options):
                        alvo, texto = _e4_normalizar(termo), _e4_normalizar(opcao.text)
                        if opcao.is_enabled() and (alvo == texto or (parcial and alvo in texto)):
                            select.select_by_index(i)
                            if not opcao.is_selected():
                                raise RuntimeError(f'Não foi possível selecionar {opcao.text}')
                            return opcao.text
            return False
        escolhido = self.esperar(selecionar_disponivel)
        self.log.info('ETAPA 4 - %s selecionado: %s.', descricao, escolhido)
        return escolhido
    #Criei esse novo def itens arvores em 23/10/2026 a pedido do Chat GPT
    def itens_arvore(self, pa):
        self.log.info(
            "ETAPA 4 - Localizando árvore de documentos do PA %s.",
            pa,
        )

        def verificar():
            # Procura o contêiner da árvore.
            for arvore in self.driver.find_elements(
                By.ID,
                "divArvore",
            ):
                if arvore.is_displayed():
                    return {"elemento": arvore}

            # Alternativa: procura o conteúdo do iframe da árvore.
            frame_atual = self.driver.execute_script(
                """
                return window.frameElement
                    ? {
                        id: window.frameElement.id,
                        nome: window.frameElement.name
                      }
                    : null;
                """
            )

            if frame_atual:
                identificadores = (
                    str(frame_atual.get("id", "")),
                    str(frame_atual.get("nome", "")),
                )

                if "ifrArvore" in identificadores:
                    for corpo in self.driver.find_elements(
                        By.TAG_NAME,
                        "body",
                    ):
                        if corpo.is_displayed():
                            return {"elemento": corpo}

            return False

        try:
            resultado = self.esperar(verificar)

        except TimeoutException as erro:
            raise TimeoutException(
                "ETAPA 4 - Árvore não localizada. "
                "Precisamos conferir o HTML da árvore neste SEI."
            ) from erro

        arvore = resultado["elemento"]

        itens = {
            (
                elemento.get_attribute("id"),
                elemento.get_attribute("href"),
                elemento.text,
            )
            for elemento in arvore.find_elements(
                By.CSS_SELECTOR,
                "a",
            )
        }

        self.log.info(
            "ETAPA 4 - Árvore localizada: %s links encontrados.",
            len(itens),
        )

        return itens

    # Chat em 22/10/2026 mandou eu excluir esse trecho, pois era uma verificação que ele fez sem me consultar que gerou
    #Problemas ao código
    #def itens_arvore(self, pa):
        # Deve encontrar primeiro a árvore do PA atual, inclusive se estiver em iframe.
     #   def verificar():
     #       for arvore in self.driver.find_elements(By.ID, 'divArvore'):
     #           if pa in arvore.text:
     #               return arvore
     #      return False
     #   arvore = self.esperar(verificar)
     #   return {(e.get_attribute('id'), e.get_attribute('href'), e.text)
     #           for e in arvore.find_elements(By.CSS_SELECTOR, 'a')}


def etapa_4_preenchimento_documentos_nos_processos(
    driver, *, logger=None, pasta_base=None, nome_aba=None,
    linha_inicial=2, timeout=60, timeout_salvar=360,
):
    """Retorna a quantidade de documentos confirmados; interrompe no primeiro erro.

    pasta_base pode ser Path(__file__).resolve().parent do script principal.
    Os timeouts são limites de espera, não pausas fixas.
    """
    log = logger if logger is not None else logging.getLogger(__name__)
    base = Path(pasta_base).resolve() if pasta_base else Path(__file__).resolve().parent
    planilha = base / 'dados_para_autuar_processos.xlsx'
    pasta_docs = base / 'docs processados'
    if linha_inicial < 2:
        raise ValueError('linha_inicial deve ser pelo menos 2.')
    tela = _E4Tela(driver, log, timeout)
    quantidade, linha, numero_pa, clicou_salvar = 0, linha_inicial, '', False
    log.info('ETAPA 4 - INÍCIO - PREENCHIMENTO DE DOCUMENTOS NOS PROCESSOS.')

    log.info('ETAPA 4 - Planilha: %s. Documentos: %s.', planilha, pasta_docs)
    wb = None
    #Adicionei essa linha abaixo "barra = None" para usar o tdqm
    barra = None
    try:
        wb = load_workbook(planilha, read_only=True, data_only=True)
        ws = wb[nome_aba] if nome_aba else wb.active
        log.info('ETAPA 4 - Aba utilizada: %s.', ws.title)
        if _e4_normalizar(ws['I1'].value) != 'numero pa gerado' or _e4_normalizar(ws['L1'].value) != 'somente pj':
            raise ValueError('Cabeçalhos inesperados: esperado I1=Numero PA gerado e L1=Somente PJ.')
        if not pasta_docs.is_dir():
            raise FileNotFoundError(f'Pasta não encontrada: {pasta_docs}')
        #Vou colocar mais um trecho aqui para usar o tdqm

        # Conta os PAs até a primeira célula vazia na coluna I.
        total_processos = 0

        for (valor_pa,) in ws.iter_rows(
            min_row=linha_inicial,
            max_row=max(ws.max_row, linha_inicial),
            min_col=9,
            max_col=9,
            values_only=True,
        ):
            if not _e4_texto(valor_pa):
                break

            total_processos += 1

        # Barra de progresso da juntada de documentos.
        barra = tqdm(
            total=total_processos,
            desc="ETAPA 4 - Processos preenchidos com documentos",
            unit="processo",
            colour="green",
            dynamic_ncols=True,
            mininterval=0,
            leave=True,
            postfix={"faltam": total_processos},
            bar_format=(
                "{desc}: {percentage:3.0f}% |{bar}| "
                "{n_fmt}/{total_fmt} "
                "[decorrido: {elapsed} | estimativa: {remaining}]"
                "{postfix}"
            ),
        )




        #Fim desse novo trecho para usar o tdqm
        def controle():
            tela.clicar('Abrindo Controle de Processos', '//*[@id="lnkControleProcessos"]/img',
                        '/html/body/div[1]/nav/div/div[3]/div[2]/div[4]/a/img')
        controle()
        # max_row+1 permite observar a primeira linha vazia após os dados.
        for linha, valores in enumerate(ws.iter_rows(min_row=linha_inicial,
                max_row=max(ws.max_row + 1, linha_inicial), max_col=12, values_only=True), linha_inicial):
            numero_pa = _e4_texto(valores[8])


            #para usar o tdqm, mandou substituir esse trecho, coisa que não vou fazer, vou deixar o trecho salvo mesmo.
            #Só comentado
            #if not numero_pa:
            #    log.info('ETAPA 4 - Coluna I vazia na linha %s. CONCLUÍDA: %s documentos.', linha, quantidade)
            #    return quantidade

            if not numero_pa:
                # Não há próximo processo: contabiliza o último.
                barra.set_postfix(
                    faltam=total_processos - quantidade,
                    refresh=False,
                )
                barra.update(quantidade - barra.n)

                log.info(
                    "ETAPA 4 - Coluna I vazia na linha %s. "
                    "CONCLUÍDA: %s documentos.",
                    linha,
                    quantidade,
                )
                return quantidade

            #Fim do trecho adicional para colocar o TDQM nessa etapa

            clicou_salvar = False
            pj, observacoes = _e4_texto(valores[11]), _e4_texto(valores[7])
            sigilo = _e4_normalizar(valores[9])
            log.info('ETAPA 4 - Linha %s | PA %s | PJ %s.', linha, numero_pa, pj)
            if not pj or sigilo not in ('sim', 'nao'):
                raise ValueError(f'Linha {linha}: coluna L obrigatória; J deve conter SIM ou NAO.')
            arquivo = _e4_arquivo(pasta_docs, pj)
            log.info('ETAPA 4 - Arquivo identificado: %s.', arquivo.name)

            # 4.1 - Abrir o PA exato e incluir documento externo.
            tela.clicar(f'Abrindo PA {numero_pa}', '//a[normalize-space(.)=' + _e4_xpath_literal(numero_pa) + ']')
            antes = tela.itens_arvore(numero_pa)

            #Linha que realiza o incremento na barra de atualização para o tdqm

            # Ao entrar neste processo, contabiliza o anterior
            # cujo salvamento já foi confirmado.
            barra.set_postfix(
                faltam=total_processos - quantidade,
                refresh=False,
            )
            barra.update(quantidade - barra.n)


            #Fim das linhas que realizam o incremento em si.

            #Mudei TODOS OS TERMOS QUE ERAM "Petição inicial do PJ" APENAS PARA "inicial do PJ"
            if any('inicial do PJ' in item[2] for item in antes):
                raise RuntimeError('Já existe Petição inicial do PJ neste PA. Confira antes de prosseguir.')
            tela.clicar('Incluindo documento', '//*[@id="divArvoreAcoes"]/a[1]/img',
                        '/html/body/div[1]/div/div/div[2]/div/a[1]/img')
            tela.preencher('txtFiltro', 'Externo', 'filtro de documentos')
            externo = '//a[normalize-space(.)="Externo"]'
            try:
                el = tela.elemento(externo, timeout=4)
            except TimeoutException:
                tela.clicar('Expandindo tipos de documento', '//*[@id="ancExibirSeries"]/img',
                            '/html/body/div[1]/div/div/form/div[2]/a/img')
                tela.preencher('txtFiltro', 'Externo', 'filtro ampliado de documentos')
                el = tela.elemento(externo)
            el.click()

            # 4.2 - Dados do documento e interessado.
            tela.selecionar('//*[@id="selSerie"]', ('Petição', 'Anexo', 'Protocolo', 'Modelo'), 'Tipo de documento')
            # Usa a data local do computador que executa o script.
            data = datetime.now().strftime('%d%m%Y')
            el = tela.preencher('txtDataElaboracao', data, 'data do documento')
            if re.sub(r'\D', '', el.get_attribute('value') or '') != data:
                raise RuntimeError('Data do documento não foi preenchida corretamente.')
            tela.preencher('txtNomeArvore', 'inicial do PJ', 'nome da árvore')
            tela.radio('divOptNato', 'Nato-digital')

            #Vou substituir essa procura por interessado identico ao termo abaixo por uma procura aproximada
            #Mudança feita em 23/10/2026
            #interessado = 'PREFEITURA MUNICIPAL DE CARAPICUÍBA'
            # Não usa TAB aqui para não fechar a lista de sugestões.
            #campo = tela.elemento('//*[@id="txtInteressado"]')
            #campo.clear()
            #campo.send_keys(interessado)
            #literal = _e4_xpath_literal(interessado)
            #tela.clicar('Selecionando interessado na lista de sugestões',
            #            f'//*[self::a or self::li or self::td or self::div or self::span][normalize-space(.)={literal} and not(.//*[normalize-space(.)={literal}])]')

            # Digita o termo de pesquisa no campo Interessados.
            interessado = "PREFEITURA MUNICIPAL DE CARAPICUÍBA"

            campo = tela.elemento('//*[@id="txtInteressado"]')
            campo.clear()
            campo.send_keys(interessado)

            log.info(
                "ETAPA 4 - Pesquisando interessado: %s.",
                interessado,
            )

            # Localiza a sugestão pelo nome, ignorando acentos,
            # maiúsculas/minúsculas e complementos como (CARAP).
            def localizar_prefeitura():
                sugestoes = driver.find_elements(
                    By.CSS_SELECTOR,
                    "#divInfraAjaxtxtInteressado li a",
                )

                correspondencias = [
                    sugestao
                    for sugestao in sugestoes
                    if sugestao.is_displayed()
                    and sugestao.is_enabled()
                    and _e4_normalizar(interessado)
                    in _e4_normalizar(sugestao.text)
                ]

                if len(correspondencias) > 1:
                    raise RuntimeError(
                        "Mais de uma sugestão corresponde à Prefeitura. "
                        "Confira a lista para evitar selecionar a errada."
                    )

                return correspondencias[0] if correspondencias else False

            sugestao = tela.esperar(localizar_prefeitura)
            texto_selecionado = sugestao.text
            sugestao.click()

            log.info(
                "ETAPA 4 - Clique realizado no interessado: %s.",
                texto_selecionado,
            )
            #Esse trecho acima é o novo pesquisar nome da prefeitura

            tela.preencher('txaObservacoes', observacoes, 'observações desta unidade')

            # 4.3 - Nível de acesso e hipótese legal, por ordem de preferência.
            if sigilo == 'nao':
                tela.radio('divOptPublico', 'Público')
            else:
                tela.radio('divOptRestrito', 'Restrito')
                # O DOCX não informa o ID desse select: identificar pelas opções.
                tela.selecionar('//select', (
                    'Segredo de justiça no processo civil (Artigo 189 da Lei Federal nº 13.105/2015)',
                    'Segredo de justiça no processo penal (Art. 201',
                    'Informação pessoal (Artigo 31 da Lei Federal nº 12527/2011)',
                    'Materiais e sistemas sigiloso (Artigo 45 do Decreto Federal nº 7845/2012)',
                ), 'Hipótese legal', parcial=True)

            # 4.4 - Anexar diretamente no input type=file (sem janela do Windows).
            log.info('ETAPA 4 - Anexando %s.', arquivo.name)
            def campo_arquivo():
                labels = driver.find_elements(By.ID, 'lblArquivo')
                for label in labels:
                    alvo = label.get_attribute('for')
                    if alvo:
                        for el in driver.find_elements(By.ID, alvo):
                            if el.get_attribute('type') == 'file' and el.is_enabled():
                                return el
                inputs = [el for el in driver.find_elements(By.CSS_SELECTOR, 'input[type="file"]') if el.is_enabled()]
                return inputs[0] if len(inputs) == 1 else False
            entrada = tela.esperar(campo_arquivo)
            entrada.send_keys(str(arquivo))
            # Aceita input que mantém FileList ou upload assíncrono que mostra o nome.
            def arquivo_pronto():
                for el in driver.find_elements(By.CSS_SELECTOR, 'input[type="file"]'):
                    nome = driver.execute_script('return arguments[0].files.length ? arguments[0].files[0].name : "";', el)
                    if nome == arquivo.name:
                        return True
                return any(e.is_displayed() for e in driver.find_elements(By.XPATH,
                    '//*[self::a or self::td or self::span][normalize-space(.)=' + _e4_xpath_literal(arquivo.name) + ']'))
            tela.esperar(arquivo_pronto)


            #Abaixo vou testar um código novo bem simples:
            #Se o arquivo .pdf tiver mais de 3 mega, que o programa espere 30 segundos adicionais
            # só para ter certeza que o arquivo realmente foi carregado


            # Pausa adicional para arquivos maiores que 3 MB.
            tamanho_mb = arquivo.stat().st_size / (1024 * 1024)

            if tamanho_mb > 3:
                log.info(
                    "ETAPA 4 - Arquivo com %.2f MB. "
                    "Aguardando 30 segundos adicionais antes de salvar.",
                    tamanho_mb,
                )

                time.sleep(30)

                log.info(
                    "ETAPA 4 - Pausa adicional concluída."
                )

            # Fim desse limite de tamanho de arquivo

            # 4.5 - Salvar UMA vez e verificar um novo documento na árvore do PA.
            salvar = tela.elemento('//*[@id="btnSalvar"]',
                '//button[normalize-space(.)="Salvar"]', '//input[@type="submit" and @value="Salvar"]')
            log.info('ETAPA 4 - Clicando em Salvar para PA %s.', numero_pa)
            clicou_salvar = True
            salvar.click()

            #Acho que esse é o último problema para resolver esse TDQM na autuação de processos
            #def documento_confirmado(_):
            #    try:
            #        novos = tela.itens_arvore(numero_pa) - antes
            #        return any('inicial do PJ' in item[2] for item in novos)
            #    except TimeoutException:
            #        return False

            def documento_confirmado(_):
                try:
                    novos = tela.itens_arvore(numero_pa) - antes

                    return any(
                        'inicial do PJ' in item[2]
                        for item in novos
                    )

                except StaleElementReferenceException:
                    log.info(
                        "ETAPA 4 - A árvore foi atualizada pelo SEI. "
                        "Aguardando nova leitura para confirmar o documento."
                    )
                    return False

                except TimeoutException:
                    return False

            #Fim desse novo trecho final para adicionar o TDQM nas autuações

            WebDriverWait(driver, timeout_salvar).until(documento_confirmado)
            quantidade += 1
            log.info('ETAPA 4 - SALVAMENTO CONFIRMADO | linha %s | PA %s | arquivo %s.', linha, numero_pa, arquivo.name)
            controle()
    except Exception:
        log.exception('ETAPA 4 - INTERROMPIDA | linha %s | PA %s | confirmados %s.', linha, numero_pa, quantidade)
        if clicou_salvar:
            log.error('ETAPA 4 - Salvar foi acionado. Confira este PA no SEI antes de retomar para evitar duplicação.')
        raise

    #Por fim (sem trocadilho), mandou eu substituir esse finalLy abaixo por esse outro para usar o TDQM

    finally:
        if barra is not None:
            barra.close()

        if wb is not None:
            wb.close()

        driver.switch_to.default_content()

    #finally:
    #    if wb is not None:
    #        wb.close()
    #    driver.switch_to.default_content()



# ----------------------------------------------------------------
# ETAPA 5 - PREPARAR ENVIO DOS PROCESSOS PARA OUTRO SETOR
# ----------------------------------------------------------------

#Etapa 5 - Envio de todos os processos autuados tanto por seleção geral (caso encontre o xpath do item) como caso não encontre
#Essa etapa irá selecionar todos os processos que foram autuados e os selecionar.


def etapa_5_preparar_envio_processos(driver, logger=None, timeout=60):
    """
    Prepara o envio dos processos para CARAP-SAJ-PAGM.

    Seleção principal:
        Usa o botão Selecionar todos.

    Seleção alternativa:
        Se o botão não for localizado, seleciona somente os PAs
        registrados na coluna I da planilha.

    O clique final em Enviar permanece comentado.
    """

    log = logger if logger is not None else logging.getLogger(__name__)
    tela = _E4Tela(driver, log, timeout)

    # ------------------------------------------------------------
    # CONFIGURAÇÕES
    # ------------------------------------------------------------

    caminho_planilha = (
        Path(__file__).resolve().parent
        / "dados_para_autuar_processos.xlsx"
    )

    # None utiliza a aba ativa.
    # Se necessário, substitua por um nome, como "CAPTACOES".
    nome_aba_planilha = None

    orgao_destino = "CARAP"
    unidade_destino = "CARAP-SAJ-PAGM"

    # ------------------------------------------------------------
    # LOCALIZADORES - CONTROLE E SELEÇÃO DOS PROCESSOS
    # ------------------------------------------------------------

    # Controle de Processos: abre a tabela de processos.
    xpath_controle = '//*[@id="lnkControleProcessos"]/img'
    xpath_controle_completo = (
        '/html/body/div[1]/nav/div/div[3]/div[2]/div[4]/a/img'
    )

    # Selecionar todos: marca os processos abrangidos pela tabela.
    xpath_selecionar_todos = '//*[@id="imgGeradosCheck"]'
    xpath_selecionar_todos_completo = (
        '/html/body/div[1]/div/div[2]/form/div/div[5]/div[3]/div[2]/table/tbody/tr[1]/th[1]/a/img'
    )

    # Checkboxes dos processos gerados.
    # Exemplo do HTML: name="chkGeradosItem0".
    xpath_checkboxes_gerados = (
        './/input[@type="checkbox" '
        'and starts-with(@name, "chkGeradosItem")]'
    )

    # Tabela que contém os checkboxes dos processos gerados.
    xpath_tabela_gerados = (
        '//table[.//input[@type="checkbox" '
        'and starts-with(@name, "chkGeradosItem")]]'
    )

    # ------------------------------------------------------------
    # LOCALIZADORES - FORMULÁRIO DE ENVIO
    # ------------------------------------------------------------

    # Enviar processos: abre o formulário de encaminhamento.
    xpath_abrir_envio = (
        '//*[@id="divBotoesControleProcessos"]/a[1]/img'
    )
    xpath_abrir_envio_completo = (
        '/html/body/div[1]/div/div[2]/form/div/div[3]/div/a[1]/img'
    )

    # Órgão das Unidades: lista suspensa para selecionar CARAP.
    xpath_orgao = '//*[@id="selOrgao"]'
    xpath_orgao_completo = (
        '/html/body/div[1]/div/div[2]/form/div[3]/select'
    )
    xpath_orgao_por_texto = (
        '//label[contains(normalize-space(.), "Órgão das Unidades")]'
        '/following::select[1]'
    )

    # Unidades: campo de pesquisa da unidade de destino.
    xpath_unidade = '//*[@id="txtUnidade"]'
    xpath_unidade_completo = (
        '/html/body/div[1]/div/div[2]/form/div[4]/input[1]'
    )
    xpath_unidade_por_texto = (
        '//label[normalize-space(.)="Unidades:" '
        'or normalize-space(.)="Unidades"]'
        '/following::input[not(@type="hidden")][1]'
    )

    # Lista de sugestões do campo Unidades.
    xpath_sugestoes = (
        '//*[@id="divInfraAjaxtxtUnidade"]//li/a'
        ' | '
        '//div[contains(concat(" ", normalize-space(@class), " "), '
        '" infraAjaxAutoCompletar ")]//li/a'
    )

    # Enviar: botão FINAL que efetiva o encaminhamento.
    xpath_enviar = '//*[@id="sbmEnviar"]'
    xpath_enviar_completo = (
        '/html/body/div[1]/div/div[2]/form/div[1]/button[1]'
    )
    xpath_enviar_por_texto = (
        '//button[normalize-space(.)="Enviar"]'
        ' | //input[(@type="submit" or @type="button") '
        'and @value="Enviar"]'
    )

    # ------------------------------------------------------------
    # FUNÇÃO AUXILIAR - SELEÇÃO ALTERNATIVA PELA PLANILHA
    # ------------------------------------------------------------

    def selecionar_processos_pela_planilha():
        log.info(
            "Etapa 5 - Lendo os PAs da coluna I para seleção alternativa."
        )

        pas_planilha = set()

        livro = load_workbook(
            caminho_planilha,
            read_only=True,
            data_only=True,
        )

        try:
            aba = (
                livro[nome_aba_planilha]
                if nome_aba_planilha is not None
                else livro.active
            )

            log.info(
                "Etapa 5 - Aba utilizada na seleção alternativa: %s.",
                aba.title,
            )

            for (valor_pa,) in aba.iter_rows(
                min_row=2,
                min_col=9,
                max_col=9,
                values_only=True,
            ):
                numero_pa = _e4_texto(valor_pa)

                # Para na primeira célula vazia da coluna I.
                if not numero_pa:
                    break

                pas_planilha.add(numero_pa)

        finally:
            livro.close()

        if not pas_planilha:
            raise RuntimeError(
                "Etapa 5 - Nenhum PA encontrado na coluna I."
            )

        log.info(
            "Etapa 5 - Encontrados %s PAs distintos na planilha.",
            len(pas_planilha),
        )

        # Localiza a tabela de Gerados, inclusive dentro de frames.
        def localizar_tabela():
            tabelas = driver.find_elements(
                By.XPATH,
                xpath_tabela_gerados,
            )

            tabelas_diretas = []

            for tabela in tabelas:
                if not tabela.is_displayed():
                    continue

                # Evita escolher uma tabela externa que contenha
                # outra tabela de processos dentro dela.
                tabelas_internas = tabela.find_elements(
                    By.XPATH,
                    './/table[.//input[@type="checkbox" '
                    'and starts-with(@name, "chkGeradosItem")]]',
                )

                if not tabelas_internas:
                    tabelas_diretas.append(tabela)

            if len(tabelas_diretas) > 1:
                raise RuntimeError(
                    "Etapa 5 - Mais de uma tabela de Gerados encontrada."
                )

            return tabelas_diretas[0] if tabelas_diretas else False

        tabela = tela.esperar(localizar_tabela)

        # Identifica todos os PAs antes de alterar as seleções.
        checkboxes_por_pa = {}

        for checkbox in tabela.find_elements(
            By.XPATH,
            xpath_checkboxes_gerados,
        ):
            numero_pa = _e4_texto(checkbox.get_attribute("title"))

            if not numero_pa:
                raise RuntimeError(
                    "Etapa 5 - Checkbox de processo sem PA no atributo title."
                )

            linha = checkbox.find_element(
                By.XPATH,
                "./ancestor::tr[1]",
            )

            links = linha.find_elements(
                By.XPATH,
                './/a[@href]',
            )

            # Confere se o título do checkbox corresponde ao texto
            # do link do processo na mesma linha da tabela.
            corresponde_ao_link = any(
                _e4_texto(link.get_attribute("textContent")) == numero_pa
                for link in links
            )

            if not corresponde_ao_link:
                raise RuntimeError(
                    "Etapa 5 - PA do checkbox não corresponde "
                    f"ao link da linha: {numero_pa}."
                )

            if numero_pa in checkboxes_por_pa:
                raise RuntimeError(
                    f"Etapa 5 - PA repetido na tabela: {numero_pa}."
                )

            checkboxes_por_pa[numero_pa] = checkbox

        faltantes = pas_planilha - set(checkboxes_por_pa)

        if faltantes:
            raise RuntimeError(
                "Etapa 5 - PAs não encontrados na tabela atual: "
                + ", ".join(sorted(faltantes))
                + ". Confira os filtros e a paginação."
            )

        # Marca os PAs da planilha e desmarca os demais
        # que estiverem selecionados nesta tabela.
        for numero_pa, checkbox in checkboxes_por_pa.items():
            deve_marcar = numero_pa in pas_planilha

            if checkbox.is_selected() == deve_marcar:
                continue

            if not checkbox.is_enabled():
                raise RuntimeError(
                    f"Etapa 5 - Checkbox desabilitado para PA {numero_pa}."
                )

            identificador = checkbox.get_attribute("id")

            linha = checkbox.find_element(
                By.XPATH,
                "./ancestor::tr[1]",
            )

            # O label está associado ao checkbox pelo atributo for.
            labels = linha.find_elements(
                By.XPATH,
                ".//label[@for="
                + _e4_xpath_literal(identificador)
                + "]",
            )

            if len(labels) != 1:
                raise RuntimeError(
                    "Etapa 5 - Label de seleção não identificado "
                    f"de forma única para PA {numero_pa}."
                )

            driver.execute_script(
                "arguments[0].scrollIntoView("
                "{block: 'center', behavior: 'instant'});",
                labels[0],
            )

            labels[0].click()

            WebDriverWait(driver, timeout).until(
                lambda _, campo=checkbox, esperado=deve_marcar:
                    campo.is_selected() == esperado
            )

            log.info(
                "Etapa 5 - PA %s: %s.",
                numero_pa,
                "selecionado" if deve_marcar else "desmarcado",
            )

        # Confere a seleção final na tabela atual.
        pas_selecionados = {
            _e4_texto(checkbox.get_attribute("title"))
            for checkbox in tabela.find_elements(
                By.XPATH,
                xpath_checkboxes_gerados,
            )
            if checkbox.is_selected()
        }

        if pas_selecionados != pas_planilha:
            raise RuntimeError(
                "Etapa 5 - A seleção final não corresponde "
                "aos PAs da planilha."
            )

        log.info(
            "Etapa 5 - Seleção alternativa conferida: %s processos.",
            len(pas_selecionados),
        )

    # ------------------------------------------------------------
    # EXECUÇÃO DA ETAPA 5
    # ------------------------------------------------------------

    try:
        log.info("============================================================")
        log.info("ETAPA 5 - PREPARAR ENVIO PARA OUTRO SETOR")
        log.info("============================================================")

        # --------------------------------------------------------
        # 5.1 - ABRIR CONTROLE E SELECIONAR OS PROCESSOS
        # --------------------------------------------------------

        log.info("Etapa 5 - Abrindo Controle de Processos.")

        tela.elemento(
            xpath_controle,
            xpath_controle_completo,
        ).click()
        #EM 27/10/2026 13:59 Esse trecho todo abaixo funciona, mas eu estou testando selecionar os processos de acordo com a planilha primeiro

        # Primeiro tenta localizar o botão Selecionar todos.
        #Esse é o método funcional antigo, vou substituir por um método com 3 alternativas diferentes:
        #Primeiro usando somente os processos da planilha, segundo procurando por um texto especifico e terceiro procurando por xpath

        #try:
        #    botao_selecionar_todos = tela.elemento(
        #        xpath_selecionar_todos,
        #        xpath_selecionar_todos_completo,
        #        timeout=10,
        #    )

        #except TimeoutException:
        #    log.warning(
        #        "Etapa 5 - Botão Selecionar todos não localizado. "
        #        "Usando os PAs da coluna I como alternativa."
        #    )

        #    selecionar_processos_pela_planilha()

        #else:
        #    log.info(
        #        "Etapa 5 - Botão Selecionar todos localizado. "
        #        "Acionando seleção da tabela."
        #    )

            # Se o clique falhar, interrompe.
            # A alternativa é usada apenas se o botão não for localizado.
         #   botao_selecionar_todos.click()

        # TESTE - Forçar seleção individual pelos PAs da planilha.
        ##log.info(
        ##    "Etapa 5 - TESTE: selecionando processos pelos PAs da coluna I."
        ##)

        ##selecionar_processos_pela_planilha()

        #Vou adicionar todos os 3 métodos de uma só vez logo abaixo

        # --------------------------------------------------------
        # AUXILIAR - LOCALIZAR SELECIONAR TODOS PELO TÍTULO "Gerados"
        # --------------------------------------------------------

        def localizar_selecionar_todos_por_gerados():

            def procurar():
                # O título "Gerados" e o botão ficam em células
                # diferentes, mas dentro da mesma linha <tr>.
                xpath = (
                    '//tr[th[normalize-space(.)="Gerados"]]'
                    '/th/a['
                    '@onclick="infraSelecaoMultipla(\'Gerados\');"'
                    ']/img[@title="Selecionar Tudo"]'
                )

                candidatos = driver.find_elements(
                    By.XPATH,
                    xpath,
                )

                controles = [
                    elemento
                    for elemento in candidatos
                    if elemento.is_displayed()
                       and elemento.is_enabled()
                ]

                if len(controles) > 1:
                    raise RuntimeError(
                        'Etapa 5 - Mais de um botão Selecionar Tudo '
                        'encontrado na linha do título "Gerados". '
                        'Seleção interrompida.'
                    )

                return controles[0] if controles else False

            return tela.esperar(procurar)

        # --------------------------------------------------------
        # AUXILIAR - CONFERIR A SELEÇÃO DOS PROCESSOS GERADOS
        # --------------------------------------------------------

        def todos_os_gerados_marcados():
            # Consulta no mesmo contexto de página/frame do botão.
            checkboxes = driver.find_elements(
                By.XPATH,
                '//input[@type="checkbox" '
                'and starts-with(@name, "chkGeradosItem")]',
            )

            if not checkboxes:
                return False

            return all(
                checkbox.is_selected()
                for checkbox in checkboxes
            )

        # --------------------------------------------------------
        # 1º MÉTODO - SELECIONAR SOMENTE OS PAs DA PLANILHA
        # --------------------------------------------------------

        try:
            log.info(
                "Etapa 5 - Método principal: "
                "selecionando os PAs registrados na coluna I."
            )

            #AQUI é o método funcional, estou testando os outros métodos agora
            selecionar_processos_pela_planilha()



            log.info(
                "Etapa 5 - Seleção pela planilha concluída "
                "e conferida."
            )

        except (TimeoutException, RuntimeError):
            log.exception(
                "Etapa 5 - Não foi possível concluir "
                "a seleção dos PAs da planilha."
            )

            # Os métodos alternativos selecionam TODOS os Gerados,
            # podendo incluir processos que não estão na planilha.


            # ----------------------------------------------------
            # 2º MÉTODO - LOCALIZAR PELO TEXTO EXATO "Gerados"
            # ----------------------------------------------------

            try:
                log.info(
                    'Etapa 5 - Método secundário: procurando '
                    'Selecionar Tudo na linha do título "Gerados".'
                )

                botao_selecionar_todos = (
                    localizar_selecionar_todos_por_gerados()
                )



                metodo_utilizado = (
                    'botão associado ao título exato "Gerados"'
                )

            except TimeoutException:
                # Somente a ausência do botão leva ao terceiro método.
                # Se houver mais de um candidato, interrompe a etapa.
                log.warning(
                    'Etapa 5 - Botão associado ao título "Gerados" '
                    "não localizado. Tentando os XPaths configurados."
                )

                # ------------------------------------------------
                # 3º MÉTODO - XPATHS CONFIGURADOS
                # ------------------------------------------------

                botao_selecionar_todos = tela.elemento(
                    xpath_selecionar_todos,
                    xpath_selecionar_todos_completo,
                    timeout=10,
                )

                metodo_utilizado = "XPaths configurados"

            # ----------------------------------------------------
            # VALIDAR QUE O BOTÃO ENCONTRADO PERTENCE A GERADOS
            # ----------------------------------------------------

            links_gerados = botao_selecionar_todos.find_elements(
                By.XPATH,
                './ancestor-or-self::a['
                '@onclick="infraSelecaoMultipla(\'Gerados\');"'
                ']',
            )

            if len(links_gerados) != 1:
                raise RuntimeError(
                    "Etapa 5 - O botão localizado não foi confirmado "
                    "como controle de seleção dos Gerados. "
                    "Confira os XPaths antes de continuar."
                )

            # ----------------------------------------------------
            # ACIONAR SELECIONAR TODOS E CONFERIR O RESULTADO
            # ----------------------------------------------------

            if todos_os_gerados_marcados():
                log.info(
                    "Etapa 5 - Todos os Gerados já estão marcados. "
                    "Não será necessário clicar em Selecionar Tudo."
                )

            else:
                driver.execute_script(
                    "arguments[0].scrollIntoView("
                    "{block: 'center', behavior: 'instant'});",
                    botao_selecionar_todos,
                )

                log.info(
                    "Etapa 5 - Acionando Selecionar Tudo. "
                    "Método utilizado: %s.",
                    metodo_utilizado,
                )

                botao_selecionar_todos.click()

                # Apenas verifica o resultado.
                # Não repete o clique automaticamente.
                def conferir_selecao(_):
                    try:
                        return todos_os_gerados_marcados()
                    except StaleElementReferenceException:
                        return False

                WebDriverWait(
                    driver,
                    timeout,
                ).until(conferir_selecao)

                log.info(
                    "Etapa 5 - Seleção de todos os Gerados confirmada."
                )



        # --------------------------------------------------------
        # 5.2 - ABRIR O FORMULÁRIO DE ENVIO
        # --------------------------------------------------------

        log.info("Etapa 5 - Abrindo formulário de envio.")

        tela.elemento(
            xpath_abrir_envio,
            xpath_abrir_envio_completo,
        ).click()

        # --------------------------------------------------------
        # 5.3 - SELECIONAR O ÓRGÃO CARAP
        # --------------------------------------------------------

        log.info(
            "Etapa 5 - Aguardando o órgão %s na lista.",
            orgao_destino,
        )

        def selecionar_orgao(_):
            try:
                campo_orgao = tela.elemento(
                    xpath_orgao,
                    xpath_orgao_completo,
                    xpath_orgao_por_texto,
                    timeout=2,
                )

                lista = Select(campo_orgao)

                for indice, opcao in enumerate(lista.options):
                    if (
                        opcao.is_enabled()
                        and _e4_normalizar(opcao.text)
                        == _e4_normalizar(orgao_destino)
                    ):
                        lista.select_by_index(indice)
                        return True

                return False

            except (TimeoutException, StaleElementReferenceException):
                return False

        WebDriverWait(driver, timeout).until(selecionar_orgao)

        log.info(
            "Etapa 5 - Órgão selecionado: %s.",
            orgao_destino,
        )

        # --------------------------------------------------------
        # 5.4 - PESQUISAR E CLICAR NA UNIDADE DE DESTINO
        # --------------------------------------------------------

        log.info(
            "Etapa 5 - Pesquisando unidade: %s.",
            unidade_destino,
        )

        campo_unidade = tela.elemento(
            xpath_unidade,
            xpath_unidade_completo,
            xpath_unidade_por_texto,
        )

        campo_unidade.clear()
        campo_unidade.send_keys(unidade_destino)

        def localizar_sugestao():
            correspondencias = []

            for sugestao in driver.find_elements(
                By.XPATH,
                xpath_sugestoes,
            ):
                if not sugestao.is_displayed() or not sugestao.is_enabled():
                    continue

                texto = _e4_normalizar(sugestao.text)
                termo = _e4_normalizar(unidade_destino)

                # Aceita a sigla acompanhada de uma descrição.
                # Evita confundir PAGM com uma sigla maior, como PAGM2.
                if re.search(
                    r"(?<![\w-])" + re.escape(termo) + r"(?![\w-])",
                    texto,
                ):
                    correspondencias.append(sugestao)

            if len(correspondencias) > 1:
                raise RuntimeError(
                    "Mais de uma sugestão corresponde à unidade "
                    f"{unidade_destino}. Confira a lista."
                )

            return correspondencias[0] if correspondencias else False

        sugestao = tela.esperar(localizar_sugestao)
        texto_sugestao = sugestao.text
        sugestao.click()

        log.info(
            "Etapa 5 - Clique realizado na sugestão: %s.",
            texto_sugestao,
        )

        # --------------------------------------------------------
        # 5.5 - LOCALIZAR ENVIAR, SEM EFETIVAR O ENVIO
        # --------------------------------------------------------

        botao_enviar = tela.elemento(
            xpath_enviar,
            xpath_enviar_completo,
            xpath_enviar_por_texto,
        )

        log.info(
            "Etapa 5 - Botão Enviar localizado. "
            "Confira os processos e a unidade selecionada na tela."
        )

        # ENVIO FINAL DESABILITADO PARA TESTES.
        # Remova o # abaixo somente quando quiser efetivar o envio.
        # botao_enviar.click()

        log.info(
            "Etapa 5 - Preparação encerrada. "
            "O clique final em Enviar está desabilitado."
        )

    except Exception:
        log.exception(
            "Etapa 5 - Falha durante a preparação do envio."
        )
        raise

    finally:
        driver.switch_to.default_content()


# ----------------------------------------------------------------
# EXECUTAR
# ----------------------------------------------------------------

try:

    logger.info(
        "Programa - Iniciando navegador Chrome."
    )

    chrome_options = Options()

    # Linha adicionada para conter o problema de entregar o navegador
    # ao selenium tendo que estar tudo carregado
    chrome_options.page_load_strategy = "eager"

    driver = webdriver.Chrome(
        options=chrome_options
    )

    logger.info(
        "Programa - Navegador Chrome iniciado."
    )


    # ----------------------------------------------------------------
    # LOGIN NO SISTEMA SEI
    # ----------------------------------------------------------------

    acessar_e_fazer_login_sei(
        driver
    )


    # ----------------------------------------------------------------
    # ETAPA 1 - PREPARAR TELA APÓS LOGIN
    # ----------------------------------------------------------------

    executar_etapa_1(
        driver
    )

    # ----------------------------------------------------------------
    # ETAPA 2 - AUTUAÇÃO DE PROCESSO
    # ----------------------------------------------------------------

    #etapa_2_autuacao_de_processo(driver)

    # ------------------------------------------------------------------------------
    # ETAPA 3 - PREENCHIMENTO DE INFORMAÇÕES PARA CONCLUIR A AUTUAÇÃO NO SISTEMA SEI
    # ------------------------------------------------------------------------------

    #etapa_3_registrar_informacoes(
    #    driver,
    #    executar_etapa_1=executar_etapa_1,
    #    etapa_2_autuacao_de_processo=etapa_2_autuacao_de_processo,
    #    caminho_planilha=(
    #            Path(__file__).resolve().parent
    #            / "dados_para_autuar_processos.xlsx"
    #    ),
    #    logger=logger,
    #)

    # ----------------------------------------------------------------
    # ETAPA 4 - PREENCHIMENTO DE DOCUMENTOS NOS PROCESSOS
    # ----------------------------------------------------------------

    #etapa_4_preenchimento_documentos_nos_processos(
    #    driver,
    #    logger=logger,
    #)

    # ----------------------------------------------------------------
    # ETAPA 5 - PREPARAR ENVIO PARA OUTRO SETOR
    # ----------------------------------------------------------------

    etapa_5_preparar_envio_processos(
        driver,
        logger=logger,
    )


except Exception:

    logger.exception(
        "Programa - ERRO durante o acesso, preparação ou autuação no SEI."
    )

    raise


time.sleep(3)








# ----------------------------------------------------------------
#Código 6 atualizar as planilhas de judicial e do expediente
# ----------------------------------------------------------------
#15/08/2025 Esse código está funcional, vou o substituir por que o novo código utiliza menos a API do Google Drive,
# Dessa Maneira, eu consigo inserir ainda mais registros na planilha sem bugs.



#Esse código serve para consumir a API do google Drive de forma otimizada.
#Pois eu tenho a atualização da planilha de procuradores, a atualização da planilha do expediente
# E depois eu terei a atualização da nova planilha, daquele de caderno virtual.

#Só estou pensando em evitar esse tipo de problema, uma vez que se consumir a API do G Drive e atingir o seu limite sem eu estar por perto,
# pode fazer o programa crashar sem que saibam como resolver.

#15/08/2025 Testei esse código com aquela planilha do excell com 27 processos, inclusive os processos do expediente deram certos
#Ou seja, mesmo autuando processos da procuradoria e logo DEPOIS os processos do expediente,
# Foram todos inseridos em suas respectivas planilhas sem problemas.
#Só preciso agora fazer o backup do Código antigo e deixar salvo e colocar esse código aqui no lugar.

logger.info("Código 6 - Iniciando atualização das planilhas de judicial e do expediente.")

# --------------------------
# Autenticação Google Sheets
# --------------------------
def autenticar_google_sheets():
    logger.info("Código 6 - Iniciando autenticação no Google Sheets.")

    credenciais_json = 'chave.json'
    escopos = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    credenciais = ServiceAccountCredentials.from_json_keyfile_name(credenciais_json, escopos)
    cliente = gspread.authorize(credenciais)

    logger.info("Código 6 - Autenticação no Google Sheets concluída com sucesso.")

    return cliente

# --------------------------
# Caso 1 - Processos Judiciais
# --------------------------
def processar_excel_para_google_planilhas():
    logger.info("Código 6 - Caso 1 - Iniciando processamento dos Processos Judiciais.")

    caminho_excel = 'dados_para_autuar_processos.xlsx'

    logger.info(
        "Código 6 - Caso 1 - Abrindo arquivo Excel: %s",
        caminho_excel
    )

    workbook = openpyxl.load_workbook(caminho_excel)
    sheet = workbook.active

    logger.info("Código 6 - Caso 1 - Arquivo Excel carregado com sucesso.")

    cliente = autenticar_google_sheets()

    logger.info(
        "Código 6 - Caso 1 - Abrindo planilha Google 'Processos Judiciais'."
    )

    planilha_caso_1 = cliente.open_by_key('1B-zoQMhTcQM4mSSKeQdHJj6uidH_q_NfkvkJYrrTMoU').worksheet('Processos Judiciais')

    logger.info(
        "Código 6 - Caso 1 - Planilha Google 'Processos Judiciais' aberta com sucesso."
    )

    linhas_para_inserir = []

    for row in sheet.iter_rows(min_row=2, values_only=True):
        if all(cell is None for cell in row):
            break

        id_processo = row[0]
        procurador = row[1]
        livro_saj = row[5]
        info_autuar_pa = row[7]
        numero_pa_gerado = row[8]

        if isinstance(id_processo, str) and id_processo.startswith('///'):
            id_processo = id_processo.split('/')[3]

        if procurador:
            data_atual = datetime.now().strftime('%d/%m/%Y')
            linhas_para_inserir.append([
                numero_pa_gerado,
                id_processo,
                info_autuar_pa,
                f"{procurador} {data_atual}",
                procurador,
                livro_saj
            ])

    logger.info(
        "Código 6 - Caso 1 - %s processo(s) judicial(is) preparado(s) para inserção.",
        len(linhas_para_inserir)
    )

    if linhas_para_inserir:
        logger.info(
            "Código 6 - Caso 1 - Enviando registros para a planilha de Processos Judiciais."
        )

        planilha_caso_1.append_rows(linhas_para_inserir, value_input_option="RAW")

        logger.info(
            "Código 6 - Caso 1 - %s registro(s) inserido(s) com sucesso na planilha da procuradoria.",
            len(linhas_para_inserir)
        )

    else:
        logger.info(
            "Código 6 - Caso 1 - Nenhum processo judicial encontrado para inserção."
        )

    print("Processos inseridos na planilha da procuradoria com sucesso!")

# --------------------------
# Caso 2 - Controle de Ofícios
# --------------------------
def processar_caso_2():
    logger.info("Código 6 - Caso 2 - Iniciando processamento do Controle de Ofícios.")

    caminho_excel = 'dados_para_autuar_processos.xlsx'

    logger.info(
        "Código 6 - Caso 2 - Abrindo arquivo Excel: %s",
        caminho_excel
    )

    workbook = openpyxl.load_workbook(caminho_excel)
    sheet = workbook.active

    logger.info("Código 6 - Caso 2 - Arquivo Excel carregado com sucesso.")

    cliente = autenticar_google_sheets()

    logger.info(
        "Código 6 - Caso 2 - Abrindo planilha Google de Controle de Ofícios."
    )

    planilha_caso_2 = cliente.open_by_key('16_zlC5bRdyGTqFcVFvRIBCYzP-fjoPN9i64tD5DGe5c').worksheet('Atual')

    logger.info(
        "Código 6 - Caso 2 - Planilha Google de Controle de Ofícios aberta com sucesso."
    )

    logger.info(
        "Código 6 - Caso 2 - Obtendo valores existentes para identificar a próxima linha."
    )

    valores_existentes = planilha_caso_2.get_all_values()
    ultima_linha = len(valores_existentes) + 1

    logger.info(
        "Código 6 - Caso 2 - Próxima linha disponível para inserção: %s",
        ultima_linha
    )

    updates = []

    for row in sheet.iter_rows(min_row=2, values_only=True):
        if all(cell is None for cell in row):
            break



        id_processo = row[0]
        procurador = row[1]
        numero_pa_gerado = row[8]
        orgao = row[2]




        if isinstance(id_processo, str) and id_processo.startswith('///'):
            id_processo = id_processo.split('/')[3]

        if procurador:
            continue

        #Essas duas linhas abaixo são funcionais, só comentei para ver o que preciso alterar para a execução do código da maneira que quero

        #data_msg = f"PA autuado em: {datetime.now().strftime('%d/%m/%Y')}"
        #updates.append([numero_pa_gerado, id_processo, orgao, "", "", "", "", "", "", "", "", "", "", "Aguardando providências", data_msg])

        data_msg = f"PA autuado em: {datetime.now().strftime('%d/%m/%Y')}"
        updates.append(
            ["", numero_pa_gerado, id_processo, orgao, "", "", "", "", "", "", "", "", "", "Aguardando providências",
             data_msg])

    logger.info(
        "Código 6 - Caso 2 - %s processo(s) preparado(s) para inserção na planilha de Controle de Ofícios.",
        len(updates)
    )

    if updates:
        logger.info(
            "Código 6 - Caso 2 - Inserindo registros na planilha de Controle de Ofícios."
        )

        planilha_caso_2.insert_rows(updates, row=ultima_linha, value_input_option="RAW")

        logger.info(
            "Código 6 - Caso 2 - %s registro(s) inserido(s) com sucesso na planilha de Controle de Ofícios.",
            len(updates)
        )

    else:
        logger.info(
            "Código 6 - Caso 2 - Nenhum processo do expediente encontrado para inserção."
        )

    print("Processos Inseridos com sucesso na planilha de controle de Ofícios (Expediente)!")

# --------------------------
# Executar
# --------------------------
try:
    processar_excel_para_google_planilhas()
    processar_caso_2()

    logger.info("Código 6 - Finalizado.")

except Exception:
    logger.exception(
        "Código 6 - ERRO durante a atualização das planilhas de judicial e do expediente."
    )
    raise




# -------------------------------------------------------------------------
# Código 6.1 - Atualizar a nova planilha que substitui os cadernos fisicos.
# -------------------------------------------------------------------------
#Atualizado em 14/08/2025, atualiza agora por lote, assim não atingindo o limite de edição por minuto do google planilhas.

logger.info("Código 6.1 - Iniciando atualização da nova planilha que substitui os cadernos físicos.")

# ====== CONFIGURAÇÕES ======
EXCEL_FILE = "dados_para_autuar_processos.xlsx"
GOOGLE_SHEET_ID = "1P2LJkCqfNhVIU5klskWdcD4Uf8RaPC12EsybR7tEH0k"
CREDENTIALS_FILE = "chave.json"
BACKUP_FILE = "backup_linha.txt"

INTERVALO_LINHAS = 50   # salto entre registros
TAMANHO_LOTE = 20       # quantos registros enviar de uma vez

logger.info(
    "Código 6.1 - Configurações carregadas. Intervalo entre registros: %s | Tamanho do lote: %s.",
    INTERVALO_LINHAS,
    TAMANHO_LOTE
)

# ====== AUTENTICAÇÃO ======
logger.info("Código 6.1 - Iniciando autenticação no Google Sheets.")

scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
client = gspread.authorize(creds)

logger.info("Código 6.1 - Autenticação no Google Sheets concluída com sucesso.")

# Abra a planilha e a primeira aba
logger.info("Código 6.1 - Abrindo a planilha do caderno virtual.")

spreadsheet = client.open_by_key(GOOGLE_SHEET_ID)
sheet = spreadsheet.sheet1

logger.info(
    "Código 6.1 - Planilha aberta com sucesso. Aba selecionada: %s",
    sheet.title
)

# Ler cabeçalho do Google Sheets (linha 1)
logger.info("Código 6.1 - Lendo cabeçalho da planilha Google.")

gs_headers = sheet.row_values(1)

if not gs_headers:
    logger.error(
        "Código 6.1 - Cabeçalho da planilha Google está vazio."
    )
    raise Exception("Cabeçalho (linha 1) da planilha está vazio. Preencha os nomes das colunas antes de rodar o script.")

logger.info(
    "Código 6.1 - Cabeçalho carregado com sucesso. Total de colunas: %s",
    len(gs_headers)
)


def col_index(col_name):
    try:
        return gs_headers.index(col_name) + 1

    except ValueError:
        logger.exception(
            "Código 6.1 - Coluna '%s' não encontrada no Google Sheets.",
            col_name
        )
        raise Exception(f"Coluna '{col_name}' não encontrada no Google Sheets.")


def col_letter(n):
    """Converte índice de coluna (1-based) para letra (A, B, ..., AA, AB, ...)."""

    letters = ""

    while n > 0:
        n, r = divmod(n - 1, 26)
        letters = chr(65 + r) + letters

    return letters


def ler_backup():
    logger.info(
        "Código 6.1 - Tentando ler arquivo de backup: %s",
        BACKUP_FILE
    )

    if not os.path.exists(BACKUP_FILE):
        logger.warning(
            "Código 6.1 - Arquivo de backup não encontrado."
        )
        return None

    try:
        with open(BACKUP_FILE, "r") as f:
            linha_backup = int(f.read().strip())

            logger.info(
                "Código 6.1 - Backup carregado. Próxima linha registrada: %s",
                linha_backup
            )

            return linha_backup

    except:
        logger.warning(
            "Código 6.1 - Não foi possível interpretar o conteúdo do arquivo de backup."
        )
        return None


def salvar_backup(linha):
    logger.info(
        "Código 6.1 - Atualizando arquivo de backup com a linha: %s",
        linha
    )

    with open(BACKUP_FILE, "w") as f:
        f.write(str(linha))

    logger.info(
        "Código 6.1 - Arquivo de backup atualizado com sucesso."
    )


# ====== ENTRADA ======
logger.info(
    "Código 6.1 - Carregando arquivo Excel: %s",
    EXCEL_FILE
)

df = pd.read_excel(EXCEL_FILE)

logger.info(
    "Código 6.1 - Arquivo Excel carregado. Total de linhas encontradas: %s",
    len(df)
)

data_atual = datetime.now().strftime("%d/%m/%Y")

linha_google = ler_backup()

if linha_google is None:
    linha_google = 2  # logo abaixo do cabeçalho

    logger.warning(
        "Código 6.1 - Nenhum backup válido encontrado. Iniciando na linha 2."
    )

else:
    logger.info(
        "Código 6.1 - Processamento será iniciado na linha %s da planilha Google.",
        linha_google
    )

# Preparos
buffer_linhas = []     # lista de listas (valores na ordem do cabeçalho)
posicoes_linhas = []   # linhas exatas onde cada registro será escrito
ultima_coluna_letra = col_letter(len(gs_headers))
titulo_aba = sheet.title


def envia_lote():
    """Envia o conteúdo de buffer_linhas para as linhas especificadas em posicoes_linhas, em uma única chamada."""

    if not buffer_linhas:
        logger.info(
            "Código 6.1 - Nenhum registro no buffer. Nenhum lote será enviado."
        )
        return

    logger.info(
        "Código 6.1 - Preparando envio de lote com %s registro(s).",
        len(buffer_linhas)
    )

    # Garante que a aba tenha linhas suficientes
    max_destino = max(posicoes_linhas)

    if sheet.row_count < max_destino:
        logger.info(
            "Código 6.1 - Planilha possui %s linhas e precisa alcançar a linha %s. Adicionando linhas.",
            sheet.row_count,
            max_destino
        )

        sheet.add_rows(max_destino - sheet.row_count)

        logger.info(
            "Código 6.1 - Linhas adicionais criadas com sucesso."
        )

    # Monta as faixas A1 e os valores
    data_payload = []

    for dados_linha, pos in zip(buffer_linhas, posicoes_linhas):
        a1 = f"{titulo_aba}!A{pos}:{ultima_coluna_letra}{pos}"
        data_payload.append({"range": a1, "values": [dados_linha]})

    logger.info(
        "Código 6.1 - Payload do lote preparado. Primeira linha de destino: %s | Última linha de destino: %s.",
        min(posicoes_linhas),
        max(posicoes_linhas)
    )

    # Um único request de escrita para todo o lote
    logger.info(
        "Código 6.1 - Enviando lote para o Google Sheets."
    )

    spreadsheet.values_batch_update({
        "valueInputOption": "USER_ENTERED",
        "data": data_payload
    })

    logger.info(
        "Código 6.1 - Lote com %s registro(s) enviado com sucesso.",
        len(buffer_linhas)
    )


try:
    registros_preparados = 0
    lotes_enviados = 0

    # ====== LOOP PRINCIPAL ======
    for _, row in df.iterrows():
        procurador = str(row.get("Procurador", "")).strip()

        if not procurador or procurador.lower() == "nan":
            continue

        numero_pa = row.get("Numero PA gerado", "")
        info_completa = row.get("Informacoes completas para autuar o PA", "")
        livro_saj = row.get("Livro saj", "")

        # Cria linha alinhada ao cabeçalho
        dados = [""] * len(gs_headers)
        dados[col_index("Data") - 1] = data_atual
        dados[col_index("numero do Processo") - 1] = numero_pa
        dados[col_index("Informações do Processo") - 1] = info_completa
        dados[col_index("Saída (COLUNA 1)") - 1] = f"{procurador} {data_atual}"
        dados[col_index("ID Livro de Registros") - 1] = livro_saj

        buffer_linhas.append(dados)
        posicoes_linhas.append(linha_google)

        registros_preparados += 1

        logger.info(
            "Código 6.1 - Registro preparado para a linha %s da planilha Google.",
            linha_google
        )

        # Próxima posição, mantendo o salto de 50
        linha_google += INTERVALO_LINHAS

        # Envia quando o lote completa
        if len(buffer_linhas) >= TAMANHO_LOTE:
            envia_lote()

            lotes_enviados += 1

            salvar_backup(linha_google)
            buffer_linhas.clear()
            posicoes_linhas.clear()

    # Envia o restante
    if buffer_linhas:
        logger.info(
            "Código 6.1 - Enviando lote final com %s registro(s).",
            len(buffer_linhas)
        )

        envia_lote()

        lotes_enviados += 1

        salvar_backup(linha_google)
        buffer_linhas.clear()
        posicoes_linhas.clear()

    logger.info(
        "Código 6.1 - Processamento concluído. Registros preparados: %s | Lotes enviados: %s | Próxima linha: %s.",
        registros_preparados,
        lotes_enviados,
        linha_google
    )

    print("✅ Dados enviados em lotes para o caderno virtual do expediente, mantendo salto de 50 linhas e evitando limite de escrita!")

    logger.info(
        "Código 6.1 - Finalizado."
    )

except Exception:
    logger.exception(
        "Código 6.1 - ERRO durante a atualização da planilha que substitui os cadernos físicos."
    )
    raise


#---------------
#Fim código 6.1
#--------------


# ----------------------------------------------------------------
#Código 7 - Limpar os itens da planilha excel e diretórios para a próxima execução
# ----------------------------------------------------------------


logger.info("Código 7 - Iniciando rotina de limpeza da planilha para a próxima execução.")



def limpar_planilha_captacoes():
    def contagem_regressiva():
        logger.info(
            "Código 7 - Contagem regressiva de 30 segundos iniciada antes da limpeza da aba CAPTACOES."
        )

        for t in range(30, 0, -1):
            print(f"\rAs informações da aba 'CAPTACOES' do arquivo Excel serão apagadas em {t} segundos...", end="")
            time.sleep(1)

        print("\nTempo esgotado. Limpando a aba automaticamente para futuros usos.")

        logger.info(
            "Código 7 - Contagem regressiva encerrada. Iniciando limpeza automática da aba CAPTACOES."
        )

        realizar_limpeza()

    def realizar_limpeza():
        logger.info(
            "Código 7 - Iniciando limpeza da aba CAPTACOES."
        )

        # Caminho do arquivo Excel
        caminho_arquivo = "dados_para_autuar_processos.xlsx"

        logger.info(
            "Código 7 - Abrindo arquivo Excel: %s",
            caminho_arquivo
        )

        # Abrindo o arquivo e selecionando a aba "CAPTACOES"
        workbook = load_workbook(caminho_arquivo)
        planilha = workbook["CAPTACOES"]

        logger.info(
            "Código 7 - Arquivo Excel aberto e aba CAPTACOES selecionada."
        )

        # Iterando sobre as linhas, preservando o cabeçalho
        for row in planilha.iter_rows(min_row=2, min_col=1, max_col=12):  # Colunas A (1) até L (12)
            for cell in row:
                cell.value = None  # Apaga o conteúdo da célula

        logger.info(
            "Código 7 - Conteúdo das colunas A até I apagado, preservando o cabeçalho."
        )

        # Salvando as modificações
        workbook.save(caminho_arquivo)

        logger.info(
            "Código 7 - Planilha salva após a limpeza."
        )

        print("As informações da aba 'CAPTACOES' foram limpas para a próxima execução.")

        logger.info(
            "Código 7 - Limpeza da aba CAPTACOES concluída com sucesso."
        )

    # Inicia a contagem regressiva em uma nova thread
    logger.info(
        "Código 7 - Criando thread responsável pela contagem regressiva e limpeza."
    )

    thread_timer = threading.Thread(target=contagem_regressiva)
    thread_timer.start()

    logger.info(
        "Código 7 - Thread de limpeza iniciada."
    )


# Chama a função para execução
try:
    limpar_planilha_captacoes()

except Exception:
    logger.exception(
        "Código 7 - ERRO ao iniciar a rotina de limpeza da planilha."
    )
    raise




# ----------------------------------------------------------------
#Código 7.1 - mover  o conteudo dos diretórios para a lixeira
# ----------------------------------------------------------------


logger.info(
    "Código 7.1 - Iniciando rotina para mover o conteúdo dos diretórios para a lixeira."
)


def mover_para_lixeira():
    # Lista das pastas que terão seus conteúdos movidos para a lixeira
    pastas = ["docs PJs", "docs numerados", "docs processados"]

    logger.info(
        "Código 7.1 - Pastas que serão verificadas: %s",
        pastas
    )

    for pasta in pastas:
        logger.info(
            "Código 7.1 - Verificando pasta: %s",
            pasta
        )

        # Verifica se a pasta existe
        if not os.path.exists(pasta):
            logger.warning(
                "Código 7.1 - A pasta '%s' não foi encontrada.",
                pasta
            )

            print(f"A pasta '{pasta}' não foi encontrada.")
            continue

        logger.info(
            "Código 7.1 - Pasta '%s' localizada.",
            pasta
        )

        # Itera sobre os arquivos e subdiretórios na pasta
        for item in os.listdir(pasta):
            caminho_item = os.path.join(pasta, item)

            # Move para a lixeira apenas arquivos ou diretórios válidos
            if os.path.isfile(caminho_item) or os.path.isdir(caminho_item):
                tentativas = 0
                max_tentativas = 20  # Número máximo de tentativas

                logger.info(
                    "Código 7.1 - Preparando para mover para a lixeira: %s",
                    caminho_item
                )

                while tentativas < max_tentativas:
                    try:
                        # Verifica se o arquivo está em uso tentando abrir em modo exclusivo
                        with open(caminho_item, 'rb+'):
                            pass

                        # Se o arquivo não estiver em uso, move para a lixeira
                        send2trash(caminho_item)

                        logger.info(
                            "Código 7.1 - Item movido para a lixeira: %s",
                            caminho_item
                        )

                        print(f"Movido para a lixeira: {caminho_item}")
                        break

                    except PermissionError:
                        tentativas += 1

                        logger.warning(
                            "Código 7.1 - O item '%s' está em uso. Tentativa %s/%s.",
                            caminho_item,
                            tentativas,
                            max_tentativas
                        )

                        print(f"O arquivo '{caminho_item}' está em uso. Tentando novamente ({tentativas}/{max_tentativas})...")
                        time.sleep(5)  # Aguarda 5 segundos antes de tentar novamente

                    except Exception as e:
                        logger.exception(
                            "Código 7.1 - Erro ao tentar mover '%s' para a lixeira.",
                            caminho_item
                        )

                        print(f"Erro ao tentar mover '{caminho_item}' para a lixeira: {e}")
                        break

                else:
                    logger.error(
                        "Código 7.1 - O item '%s' não pôde ser movido após %s tentativas.",
                        caminho_item,
                        max_tentativas
                    )

                    print(f"O arquivo '{caminho_item}' não pôde ser movido após {max_tentativas} tentativas.")

    print("Processo de mover arquivos para a lixeira concluído.")

    logger.info(
        "Código 7.1 - Processo de mover arquivos para a lixeira concluído."
    )


# Chamar a função
try:
    mover_para_lixeira()

    logger.info(
        "Código 7.1 - Finalizado."
    )

except Exception:
    logger.exception(
        "Código 7.1 - ERRO não tratado durante a limpeza dos diretórios."
    )
    raise


#----------------
# FIM CÓDIGO 7.1
#----------------



#-----------------------------------------------------------
# Código 7.2 APAGAR o conteudo da pasta "segredo de justica"
#-----------------------------------------------------------

logger.info(
    "Código 7.2 - Iniciando limpeza da pasta 'segredo de justica'."
)

pasta_segredo_justica = "segredo de justica"

if os.path.exists(pasta_segredo_justica):

    logger.info(
        "Código 7.2 - Pasta 'segredo de justica' localizada."
    )

    arquivos_removidos = 0

    for nome_item in os.listdir(pasta_segredo_justica):
        caminho_item = os.path.join(
            pasta_segredo_justica,
            nome_item
        )

        logger.info(
            "Código 7.2 - Tentando mover para a lixeira: %s",
            nome_item
        )

        try:
            send2trash(caminho_item)

            logger.info(
                "Código 7.2 - Item movido para a lixeira com sucesso: %s",
                nome_item
            )

            print(f"Movido para a lixeira: {nome_item}")
            arquivos_removidos += 1

        except Exception as erro:
            logger.exception(
                "Código 7.2 - Erro ao mover '%s' para a lixeira.",
                nome_item
            )

            print(
                f"Erro ao mover '{nome_item}' para a lixeira: "
                f"{erro}"
            )

    if arquivos_removidos == 0:
        logger.info(
            "Código 7.2 - A pasta 'segredo de justica' já está vazia."
        )

        print("A pasta 'segredo de justica' já está vazia.")

    else:
        logger.info(
            "Código 7.2 - Limpeza concluída. %s item(ns) movido(s) para a lixeira.",
            arquivos_removidos
        )

        print(
            f"Limpeza concluída: {arquivos_removidos} "
            f"item(ns) movido(s) para a lixeira."
        )

else:
    logger.warning(
        "Código 7.2 - A pasta 'segredo de justica' não foi encontrada."
    )

    print("A pasta 'segredo de justica' não foi encontrada.")

print("Conteudo da pasta de 'segredo de justica' movido para a lixeira.")

logger.info(
    "Código 7.2 - Finalizado."
)

#--------------------------------------------------------------
#FIM DO CÓDIGO 7.2
#--------------------------------------------------------------


#----------------------------------------------------
# Código 8 - Realizar o upload do arquivo de linhas atualizado.
#----------------------------------------------------

logger.info(
    "Código 8 - Iniciando upload do arquivo de linhas atualizado para o Google Drive."
)

print("Enviando o arquivo de texto com o número da linha atual para o Google Drive.")
print("Por favor, aguarde.")


SERVICE_ACCOUNT_FILE = "atualizar-numero-de-linhas-dd22a7445647.json"  # seu arquivo
SCOPES = ["https://www.googleapis.com/auth/drive"]
FOLDER_ID = "1BY492VgzAMM2FyZWMrKsOKCDx_sKZWzx"
ARQUIVO = "backup_linha.txt"


def encontrar_arquivo(service, nome, folder_id):
    logger.info(
        "Código 8 - Procurando o arquivo '%s' na pasta do Google Drive.",
        nome
    )

    query = f"name='{nome}' and '{folder_id}' in parents and trashed=false"
    resp = service.files().list(q=query, fields="files(id, name)").execute()
    files = resp.get("files", [])

    if not files:
        logger.info(
            "Código 8 - Arquivo '%s' não encontrado no Google Drive.",
            nome
        )

        return None

    logger.info(
        "Código 8 - Arquivo '%s' encontrado no Google Drive.",
        nome
    )

    return files[0]["id"]


def atualizar_drive():
    logger.info(
        "Código 8 - Carregando credenciais da conta de serviço."
    )

    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES
    )

    logger.info(
        "Código 8 - Credenciais carregadas com sucesso."
    )

    logger.info(
        "Código 8 - Criando conexão com a API do Google Drive."
    )

    service = build("drive", "v3", credentials=creds)

    logger.info(
        "Código 8 - Conexão com a API do Google Drive criada com sucesso."
    )

    file_id = encontrar_arquivo(service, ARQUIVO, FOLDER_ID)

    logger.info(
        "Código 8 - Preparando arquivo local '%s' para upload.",
        ARQUIVO
    )

    media = MediaFileUpload(ARQUIVO, mimetype="text/plain")

    if file_id:
        logger.info(
            "Código 8 - Arquivo já existe no Google Drive. Iniciando atualização."
        )

        print("🔄 Atualizando arquivo com a numeração de linhas no Google Drive (upload)...")

        service.files().update(fileId=file_id, media_body=media).execute()

        logger.info(
            "Código 8 - Arquivo existente atualizado com sucesso no Google Drive."
        )

    else:
        logger.info(
            "Código 8 - Arquivo não existe no Google Drive. Iniciando criação."
        )

        print("⬆️ Enviando arquivo para o Drive...")

        service.files().create(
            body={"name": ARQUIVO, "parents": [FOLDER_ID]},
            media_body=media
        ).execute()

        logger.info(
            "Código 8 - Novo arquivo criado com sucesso no Google Drive."
        )

    print("✔️ Operação concluída! Arquivo enviado com sucesso.")

    logger.info(
        "Código 8 - Operação de upload concluída com sucesso."
    )


if __name__ == "__main__":
    try:
        atualizar_drive()

        logger.info(
            "Código 8 - Finalizado."
        )

    except Exception:
        logger.exception(
            "Código 8 - ERRO durante o upload do arquivo de linhas atualizado para o Google Drive."
        )
        raise


#FIM Código 8


#----------------------------------------------------
# Código 9 - refinar os arquivos a serem excluidos
#----------------------------------------------------

logger.info(
    "Código 9 - Iniciando refino dos arquivos presentes na lixeira."
)

time.sleep(10)

FRAGMENTOS_MANTER = [
    "capa_ref_pa_",
    "_completo"
]

MODO_TESTE = False  # depois mude para False para apagar

logger.info(
    "Código 9 - Modo de teste: %s",
    MODO_TESTE
)

logger.info(
    "Código 9 - Fragmentos protegidos contra exclusão: %s",
    FRAGMENTOS_MANTER
)

print("\n")
print("\nCódigo 9 - Listando itens a serem excluidos:")
print("\n")

logger.info(
    "Código 9 - Iniciando localização da lixeira do usuário atual."
)

# Descobre a Lixeira do usuário atual
saida = subprocess.check_output(
    ["whoami", "/user"],
    text=True,
    encoding="utf-8",
    errors="ignore"
)

sid = next(x for x in saida.split() if x.startswith("S-1-5-"))
lixeira = rf"C:\$Recycle.Bin\{sid}"

logger.info(
    "Código 9 - Lixeira do usuário localizada."
)


def nome_original(arquivo):
    with open(arquivo, "rb") as f:
        dados = f.read()

    if len(dados) < 28:
        return None

    versao = struct.unpack_from("<Q", dados, 0)[0]

    if versao == 2:
        tamanho = struct.unpack_from("<I", dados, 24)[0]
        return dados[28:28 + tamanho * 2].decode(
            "utf-16-le",
            errors="ignore"
        ).rstrip("\x00")

    return dados[24:].decode(
        "utf-16-le",
        errors="ignore"
    ).split("\x00")[0]


logger.info(
    "Código 9 - Iniciando análise dos itens presentes na lixeira."
)

itens_mantidos = 0
itens_excluidos = 0
itens_que_seriam_excluidos = 0
erros_refino = 0


for arquivo in os.listdir(lixeira):

    if not arquivo.startswith("$I"):
        continue

    arquivo_i = os.path.join(lixeira, arquivo)
    arquivo_r = os.path.join(lixeira, "$R" + arquivo[2:])

    # Se o arquivo real não existe, ignora
    if not os.path.exists(arquivo_r):
        logger.warning(
            "Código 9 - Arquivo de metadados encontrado, mas o item real correspondente não existe: %s",
            arquivo
        )
        continue

    try:
        original = nome_original(arquivo_i)

        if not original:
            logger.warning(
                "Código 9 - Não foi possível determinar o nome original do item: %s",
                arquivo
            )
            continue

        nome = os.path.basename(original)

        if any(fragmento.lower() in nome.lower() for fragmento in FRAGMENTOS_MANTER):
            logger.info(
                "Código 9 - Item protegido e mantido na lixeira: %s",
                nome
            )

            itens_mantidos += 1

            print(f"[MANTER]  {nome}")
            continue

        if MODO_TESTE:
            logger.info(
                "Código 9 - MODO TESTE - Item seria excluído: %s",
                nome
            )

            itens_que_seriam_excluidos += 1

            print(f"[APAGARIA] {nome}")

        else:
            logger.info(
                "Código 9 - Excluindo definitivamente item da lixeira: %s",
                nome
            )

            if os.path.isdir(arquivo_r):
                shutil.rmtree(arquivo_r)
            else:
                os.remove(arquivo_r)

            os.remove(arquivo_i)

            itens_excluidos += 1

            logger.info(
                "Código 9 - Item excluído definitivamente: %s",
                nome
            )

            print(f"[EXCLUÍDO] {nome}")

    except Exception as erro:
        erros_refino += 1

        logger.exception(
            "Código 9 - ERRO durante o processamento do item '%s'.",
            arquivo
        )

        print(f"[ERRO] {arquivo}: {erro}")


print("\nProcesso de refino de arquivos na lixeira finalizado.")

logger.info(
    "Código 9 - Processo de refino da lixeira finalizado."
)

logger.info(
    "Código 9 - Resumo: mantidos=%s | excluídos=%s | seriam excluídos no modo teste=%s | erros=%s",
    itens_mantidos,
    itens_excluidos,
    itens_que_seriam_excluidos,
    erros_refino
)

if MODO_TESTE:
    print("Nenhum arquivo foi apagado.")

    logger.info(
        "Código 9 - MODO TESTE ativo. Nenhum arquivo foi apagado."
    )

logger.info(
    "Código 9 - Finalizado."
)




#----------------------------------------------------
# Fim código 9
#----------------------------------------------------



