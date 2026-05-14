# -*- Coding: UTF-8 -*-
# encoding: utf-8
# encoding: iso-8859-1
# encoding: win-1252

"""
Ferramenta gerencia os processos no navegador
author: Teonas Gonçalves Dourado Netto
e-mail: teonasnetto@gmail.com
version: 1.0.0
date: 2022-02-09
"""

import os
import traceback
from time import sleep
import shutil
from dotenv import load_dotenv
from typing import Tuple

import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.common.exceptions import NoSuchElementException, TimeoutException, WebDriverException

from directories import Directories
from access import Access
from files import Files
from gdrive import GDrive
from bot import Bot
from gsheets import Psheets

# Carrega as variáveis de ambiente do arquivo .env
load_dotenv()

ALERT_MESSAGE_ERROR = 'Timed out waiting for PA creation confirmation popup to appear'


class WebCrawler:  # pylint: disable=line-too-long
    """
    Funções para trabalhar diretamente no navegador
    """

    def __init__(self, _dir: Directories, access: Access, files: Files, telegram: Bot, gsheets: Psheets, gdrive: GDrive) -> None:
        self._dir = _dir
        self.access = access
        self.files = files
        self.telegram = telegram
        self.gsheets = gsheets
        self.gdrive = gdrive

    def start_chorme_driver(self):  # type: ignore
        """
        Inicialização da configuração do selenium usando o chromedriver
        """
        self.close_old_instances()
        chrome_cache, chrome_binaries, chromedriver = self.create_variables()
        chrome_options = self.create_chrome_options(
            chrome_cache, chrome_binaries)

        if not os.path.isfile(chromedriver) or not os.path.isfile(chrome_binaries):
            print("Chromedriver/Chromium não encontrado. Baixando...")
            shutil.rmtree(os.path.join(self._dir.dir_data, os.getenv(
                'CHROME_BINARIES_PATH').strip('\\')), ignore_errors=True)
            gdrive_content = self.gdrive._list_folder_content(
                os.getenv("GDRIVE_CHROME_ID"))
            df_gdrive_bd = pd.DataFrame.from_dict(gdrive_content.items())
            file_id = df_gdrive_bd.loc[df_gdrive_bd[0] == os.getenv('ZIP_CHROME_NAME'), [
                1]].values[0][0]
            file_name = df_gdrive_bd.loc[df_gdrive_bd[0] == os.getenv('ZIP_CHROME_NAME'), [
                0]].values[0][0]
            self.gdrive._download_file(
                file_id, os.path.join(self._dir.dir_temp, file_name))
            self.files.unzip_files(os.path.join(
                self._dir.dir_data, file_name.split('.')[0]))
            os.rename(os.path.join(self._dir.dir_data, os.getenv('CHROME_BINARIES_PATH').strip(
                '\\'), 'chromiumportable.exe'), os.path.join(self._dir.dir_data, chrome_binaries))
            os.remove(os.path.join(self._dir.dir_temp, file_name))

        global driver
        try:
            # driver = webdriver.Chrome(service=Service(chromedriver), options=chrome_options)
            # ALTERAÇÃO: Adicionado log_output=os.devnull para silenciar o "DevTools listening..."
            driver = webdriver.Chrome(service=Service(chromedriver, log_output=os.devnull), options=chrome_options)
        except Exception as _e:
            traceback.print_exc()
            print(_e)

    def close_old_instances(self) -> None:
        print('#Fechando instancias antigas')
        proccess_name = os.getenv('CHROME_PROCCESS_NAME').strip('\\')
        os.system(f'wmic process where name="{proccess_name}" delete')

    def create_variables(self) -> Tuple[str, str, str]:
        chrome_cache = None
        chrome_binaries = None
        chromedriver = None

        if os.getenv("CHROME_CACHE") is not None and os.getenv("CHROME_CACHE").startswith('C:'):
            chrome_cache = os.getenv("CHROME_CACHE")
        elif os.getenv("CHROME_CACHE") is not None and os.getenv("CHROME_CACHE").startswith('\\'):
            chrome_cache = os.path.join(
                self._dir.dir_data, os.getenv("CHROME_CACHE").strip("\\"))

        if os.getenv("CHROME_BINARIES_PATH") is not None and os.getenv("CHROME_BINARIES_PATH").startswith('C:'):
            chrome_binaries = os.getenv("CHROME_BINARIES_PATH")
        elif os.getenv("CHROME_BINARIES_PATH") is not None and os.getenv("CHROME_BINARIES_PATH").startswith('\\'):
            chrome_binaries = os.path.join(self._dir.dir_data, os.getenv(
                "CHROME_BINARIES_PATH").strip("\\"), os.getenv("CHROME_PROCCESS_NAME").strip("\\"))

        if os.getenv("CHROMEDRIVER") is not None and os.getenv("CHROMEDRIVER").startswith('C:'):
            chromedriver = os.getenv("CHROMEDRIVER")
        elif os.getenv("CHROMEDRIVER") is not None and os.getenv("CHROMEDRIVER").startswith('\\'):
            chromedriver = os.path.join(
                self._dir.dir_data, os.getenv("CHROMEDRIVER").strip("\\"))

        return chrome_cache, chrome_binaries, chromedriver

    def create_chrome_options(self, chrome_cache: str, chrome_binaries: str) -> Options:
        chrome_options = Options()
        prefs = {"download.default_directory": self._dir.dir_temp,
                 "download.prompt_for_download": False,
                 "download.directory_upgrade": True}
        chrome_options.add_experimental_option("prefs",
                                               prefs)
        chrome_options.add_argument('no-sandbox')
        chrome_options.add_argument('disable-gpu')
        chrome_options.add_argument('disable-dev-shm-usage')
        # chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])        
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])# ALTERAÇÃO: Adicionado "enable-logging" para limpar erros de hardware/buffer
        chrome_options.add_experimental_option('useAutomationExtension', False)
        chrome_options.add_argument(
            "--disable-blink-features=AutomationControlled")
        chrome_options.add_argument(
            "--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.110 Safari/537.36")
        chrome_options.add_argument("--window-size=1920,1080")
        chrome_options.add_argument("--log-level=3")
        chrome_options.add_argument("--disable-logging")
        chrome_options.add_argument(f'user-data-dir={chrome_cache}')
        chrome_options.binary_location = chrome_binaries
        return chrome_options

    def quit_chrome_driver(self) -> None:
        """
        Fechar o chromedriver
        """
        print('#Fechando navegador')
        driver.quit()

    def log_into_gpm(self, operation: str) -> None:
        """
        Realizar login no GPM
        """
        try:
            print('#Logando portal GPM-BA')
            driver.get(f"https://sirtec{operation}.gpm.srv.br/includes/common/logout.php")
            driver.get(f"https://sirtec{operation}.gpm.srv.br/")

            try:
                driver.find_element(By.PARTIAL_LINK_TEXT, "Sair")

            except NoSuchElementException:
                WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.ID, "idLogin")))
                driver.find_element(
                    By.ID, "idLogin").send_keys(self.access.login_gpm)
                driver.find_element(
                    By.ID, "idSenha").send_keys(self.access.passw_gpm)
                driver.find_element(By.XPATH, "//input[contains(@value, 'ntrar')]").click()
        except Exception(NoSuchElementException, TimeoutException, WebDriverException) as _e:
            print(_e)
            driver.quit()

    def consulta_servicos_set_os(self, lote: list) :
            print('#Abrindo consulta serviços')
            driver.switch_to.default_content()
            driver.get(
                'https://sirtecba.gpm.srv.br/gpm/geral/consulta_servico.php')
            driver.switch_to.default_content()
            WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.ID, "num_os"))).clear()
            WebDriverWait(driver, 20).until(EC.element_to_be_clickable(
                (By.ID, "num_os"))).send_keys(" ".join(lote))
            WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.NAME, "submit"))).click()
            WebDriverWait(driver, 20).until(EC.element_to_be_clickable(
                (By.XPATH, "//*[contains(text(), 'CSV/Excel')]"))).click()
            filename = self.files.check_download_name_file(
                'consulta_servicos_')
            self.files.unzip_files(self._dir.dir_temp + "\\parts")
            os.remove(os.path.join(self._dir.dir_temp, filename))
            window_handles = driver.window_handles
            driver.switch_to.window(window_handles[1])
            driver.close()
            driver.switch_to.window(window_handles[0])
            driver.switch_to.default_content()


    def consulta_servicos_gpm2(self) :
        try:
            print("Listando notas para download")
            iw_59_df = pd.read_excel(os.path.join(
                self._dir.dir_temp, 'RELIGAS_EXCEL.xlsx'), sheet_name=0)
            iw_59_df['Nota'] = iw_59_df['Nota'].astype(str)
            files_to_download = iw_59_df['Nota'].tolist()
            tamanho_list = len(files_to_download) # retorna o tamanho (quantidade) da lista de serviços
            tamanho_lote = 3000 # limita o número de buscas no gpm
            # Inicia busca fracionada pelos serviços no gpm
            for inicio in range(0, tamanho_list, tamanho_lote):
                fim = min(inicio + tamanho_lote, tamanho_list)
                lote = files_to_download[inicio:fim]
                
                self.consulta_servicos_set_os(lote)
                sleep(1)
            

            # Lista todos os arquivos .csv na pasta
            arquivos_csv = [file for file in os.listdir(self._dir.dir_temp + "\\parts") if file.endswith('.csv')]

            # Inicializa uma lista para armazenar os DataFrames
            lista_dataframes = []

            # Percorre cada arquivo .csv, lê o conteúdo e adiciona à lista de DataFrames
            for arquivo in arquivos_csv:
                caminho_arquivo = os.path.join(self._dir.dir_temp + "\\parts", arquivo)
                df = pd.read_csv(caminho_arquivo, encoding='utf-8-sig', sep=';', low_memory=False, keep_default_na=False)  # Lê o CSV
                lista_dataframes.append(df)
                

            # Concatena todos os DataFrames em um só
            df_concatenado = pd.concat(lista_dataframes, ignore_index=True)

            # Salva o DataFrame concatenado em um novo arquivo CSV
            df_concatenado.to_csv(os.path.join(self._dir.dir_temp, "GPM.csv"), encoding='utf-8-sig', sep=';', index=False)


        except (NoSuchElementException, TimeoutException, WebDriverException):
            driver.quit
        
    def consulta_servicos_gpm(self):
        """
        Abre o consulta serviços do GPM para baixar o relatório das notas que foram repassadas pelo clipboard no clipboard_IW59_GPM().
        """
        try:
            print('#Abrindo consulta serviços')
            driver.get(
                'https://sirtecba.gpm.srv.br/gpm/geral/consulta_servico.php')
            driver.switch_to.default_content()
            WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.ID, "num_os"))).clear()
            iw_59_df = pd.read_excel(os.path.join(
                self._dir.dir_temp, 'RELIGAS_EXCEL.xlsx'), sheet_name=0)
            iw_59_df['Nota'] = iw_59_df['Nota'].astype(str)
            WebDriverWait(driver, 20).until(EC.element_to_be_clickable(
                (By.ID, "num_os"))).send_keys(" ".join(iw_59_df['Nota']))
            WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.NAME, "submit"))).click()
            WebDriverWait(driver, 20).until(EC.element_to_be_clickable(
                (By.XPATH, "//*[contains(text(), 'Excel')]"))).click()
            filename = self.files.check_download_name_file(
                'consulta_servicos_')
            extracted_files = self.files.unzip_files(self._dir.dir_temp)
            os.remove(os.path.join(self._dir.dir_temp, filename))
            for idx, file in enumerate(extracted_files):
                ext = file.split('.')[-1]
                new_file_name = f'GPM.{ext}'
                os.rename(os.path.join(self._dir.dir_temp, file),
                          os.path.join(self._dir.dir_temp, new_file_name))
            window_handles = driver.window_handles
            driver.switch_to.window(window_handles[1])
            driver.close()
            driver.switch_to.window(window_handles[0])
            driver.switch_to.default_content()
        except (NoSuchElementException, TimeoutException, WebDriverException):
            driver.quit

    def open_servicos_sap_2_bd(self, operation):
        """
        Importa os arquivos gerados e tratados para o GPM
        """

        try:
            driver.get(
                f'https://sirtec{operation}.gpm.srv.br/gpm/geral/importar_dados_bd.php?ori=241')
        except (NoSuchElementException, TimeoutException, WebDriverException):
            driver.quit()

    def check_import_old_files_gpm(self):   
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable(
            (By.XPATH, "/html/body/form[2]/div[2]/input"))).click()
        try:
            WebDriverWait(driver, 10).until(EC.alert_is_present(),
                                            ALERT_MESSAGE_ERROR)
            text_alert = driver.switch_to.alert.text
            print(text_alert)
            driver.switch_to.alert.accept()
        except (NoSuchElementException, TimeoutException, WebDriverException):
            WebDriverWait(driver, 20).until(EC.element_to_be_clickable(
                (By.XPATH, "/html/body/form[2]/div[4]/div/div[2]/span[2]/input"))).click()
            WebDriverWait(driver, 10).until(EC.alert_is_present(),
                                            ALERT_MESSAGE_ERROR)
            text_alert = driver.switch_to.alert.text
            driver.switch_to.alert.accept()
            WebDriverWait(driver, 10).until(EC.alert_is_present(),
                                            ALERT_MESSAGE_ERROR)
            driver.switch_to.alert.accept()

    def check_import_files(self):
        """
        Funções para manipulação de arquivos
        """
        print("#Checando arquivos e importando.")
        utds = ['ITG',
                'BJL',
                'BRU',
                'VTC',
                'JEQ',
                'IRE',
                'FSA',
                'SER',
                'ITA',
                'GUA',
                'IBO',
                'SBA',
                'LIV',
                'BAR']

        # Carrega as páginas das configurações usadas
        wks = self.gsheets.worksheet_select(
            'bd_config_robo_religacao', 'unidades_coord')

        unidades_coord_df = pd.DataFrame(wks.get_as_df(), dtype=str)

        unidades_coord_df.set_index("utd", inplace=True)

        try:
            for utd in utds:
                contrato = unidades_coord_df.loc[utd]['contrato']
                coordenador = unidades_coord_df.loc[utd]['coordenador']
                supervisor = unidades_coord_df.loc[utd]['supervisor']
                if (os.path.isfile(self._dir.dir_temp + utd + '.csv')):
                    self.import_files_gpm(
                        utd, contrato, coordenador, supervisor)
                else:

                    self.telegram.bot_message(f'Sem notas para {utd}')
                    print(f'Sem notas para {utd}')
        except RuntimeError as e:
            raise RuntimeError(e)
        except Exception as e:
            print(e)
            traceback.print_exc()

    def import_files_gpm(self, utd, contrato, coordenador, supervisor):
      try :
        sleep(1)
        # Selecionando contrato
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable(
            (By.XPATH, "/html/body/form[2]/div[1]/table/tbody/tr[1]/td[2]/div/a"))).click()
        sleep(1)
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable(
            (By.XPATH, "/html/body/form[2]/div[1]/table/tbody/tr[1]/td[2]/div/div/ul/li[contains(text(),'" + contrato + "')]"))).click()

        sleep(1)
        # Selecionando coordenador
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable(
            (By.XPATH, "/html/body/form[2]/div[1]/table/tbody/tr[2]/td[2]/div/a"))).click()
        sleep(1)
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable(
            (By.XPATH, "/html/body/form[2]/div[1]/table/tbody/tr[2]/td[2]/div/div/ul/li[contains(text(),'" + coordenador + "')]"))).click()

        sleep(1)
        # Selecionando supervisor
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable(
            (By.XPATH, "/html/body/form[2]/div[1]/table/tbody/tr[2]/td[4]/div/a"))).click()
        sleep(1)
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable(
            (By.XPATH, "/html/body/form[2]/div[1]/table/tbody/tr[2]/td[4]/div/div/ul/li[contains(text(),'" + supervisor + "')]"))).click()

        # Desmarcar checkbox de atualizar coordenadas
        if (WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "cCoords"))).is_selected()):
            WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, "cCoords"))).click()
            
        #Marcar cZona somente se nao estiver marcando por padrao
        #if (not WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, "cZona"))).is_selected()):
        #    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, "cZona"))).click()

        # Selecionar arquivo
        WebDriverWait(driver, 20).until(EC.presence_of_element_located(
            (By.ID, "dados"))).send_keys(self._dir.dir_temp + utd + '.csv')
        sleep(1)
        WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.ID, "idSubmit"))).click()
        

        # Aguardar alerta
        WebDriverWait(driver, 20).until(EC.alert_is_present(),
                                        ALERT_MESSAGE_ERROR)
        alert = driver.switch_to.alert
        if "Erro na importação" in alert.text :
            self.telegram.bot_message('Erro, verificar urgentemente: '+alert.text)
            raise RuntimeError(alert.text)
        alert.accept()

        
        WebDriverWait(driver, 20).until(EC.alert_is_present(),
                                        ALERT_MESSAGE_ERROR)
        alert = driver.switch_to.alert
        if "Erro na importação" in alert.text :
            self.telegram.bot_message('Erro, verificar urgentemente: '+alert.text)
            raise RuntimeError(alert.text)
        alert.accept()
        sleep(2)
        WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.ID, "idProcessa"))).click()
        WebDriverWait(driver, 30).until(EC.alert_is_present(),
                                        ALERT_MESSAGE_ERROR)
        alert = driver.switch_to.alert
        if "Erro na importação" in alert.text :
            self.telegram.bot_message('Erro, verificar urgentemente: '+alert.text)
            raise RuntimeError(alert.text)
        alert.accept()
        self.telegram.bot_message('Notas ' + utd + ' importadas')
        print('#Notas ' + utd + ' importadas')
      except Exception as e:
        print(e)
        raise RuntimeError("Erro na importação")