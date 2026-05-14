import time
import shutil
import pandas as pd
from datetime import datetime, timedelta
import multiprocessing
import os
from dotenv import load_dotenv
import warnings

warnings.simplefilter(action='ignore', category=FutureWarning) # ignorar os FutureWarnings do terminal

pd.options.mode.chained_assignment = None  # Desativa o aviso SettingWithCopyWarning

from gdrive import GDrive
from gsheets import Psheets
from directories import Directories
from sap_handler import SapHandler
from access import Access
from web_crawler import WebCrawler
from data_handle import DataHandling
from bot import Bot
from files import Files
from gdrive import GDrive
from image_finder import ImageFinder


# Carrega as variáveis de ambiente do arquivo .env
load_dotenv()

_dir = Directories()
telegram = Bot()
gsheets = Psheets(_dir)
files = Files(_dir)
gdrive = GDrive('service_account', _dir)
image_finder = ImageFinder(_dir)
access = Access(gsheets)
sap_handler = SapHandler(_dir, image_finder)
web_crawler = WebCrawler(_dir, access, files, telegram, gsheets, gdrive)
data_handle = DataHandling(_dir, gsheets, gdrive, telegram)

# tempo para reiniciar o robo a cada ciclo
TIME_RESTART = int(os.getenv('TIME_RESTART', 900))
# tempo para resetar o robo considerando travamentos
TIME_RESET = int(os.getenv('TIME_RESET', 1800))

print(f'Tempo de reinicio: {TIME_RESTART} segundos')
print(f'Tempo de reset: {TIME_RESET} segundos')

os.system('title Robo de religas/cortes - BA')


def inicio(q: multiprocessing.Queue):
    # Marcar hora inicial
    hora_i = datetime.now().strftime('%d-%m-%Y %H:%M:%S')
    q.put(hora_i)

    # verifica se a hora está entre 02hrs e 05hrs
    now = datetime.now().time()
    if timedelta(hours=0) <= timedelta(hours=now.hour, minutes=now.minute, seconds=now.second) < timedelta(hours=6):
        variante = 'PCP_CORTE'
        message = 'cortes/religas'
    else:
        variante = 'PCP_RELIGA'
        message = 'religações'

    start_message = f' ############# Iniciando importação de {message}: {hora_i} ############# '
    print(start_message)
    telegram.bot_message(start_message)

    return variante, hora_i


def fim(o: multiprocessing.Queue, hora_i, time_restart):
    # Hora que foi finalizado
    hora_f = datetime.now().strftime('%d-%m-%Y %H:%M:%S')
    o.put(hora_f)

    # Calculo do tempo gasto
    total_time = (datetime.strptime(hora_f, '%d-%m-%Y %H:%M:%S') -
                  datetime.strptime(hora_i, '%d-%m-%Y %H:%M:%S'))

    finish_message = f' ############# Finalizado as: {hora_f} em {total_time} ############# '
    print(finish_message)

    # tempo para reiniciar o robo
    print(
        f' ############# Reiniciando em {time_restart/60} minutos ############# ')


def close_programs():
    os.system('wmic process where name="saplogon.exe" delete')
    os.system('wmic process where name="excel.exe" delete')


def app(q: multiprocessing.Queue, o: multiprocessing.Queue):

    while True:
        try:

            shutil.rmtree(_dir.dir_temp, ignore_errors=False)

            _dir.create_dir()

            close_programs()

            variante, hora_i = inicio(q)

            # download bd_coords if not exists
            data_handle.bd_coords_download()

            # Begin, start session
            ses = sap_handler.sap_login(access.login_sap, access.passw_sap)
            sap_handler.maximize(ses)
            sap_handler.open_iw59(ses, variante)
            # Download GPM.csv
            web_crawler.start_chorme_driver()
            web_crawler.log_into_gpm("ba")
            web_crawler.consulta_servicos_gpm2()
            web_crawler.quit_chrome_driver()

            # Create nimport file
            data_handle.create_nimport_file()
            
            
            #aqui tem que copiar para o clipboard somente as notas que deverão ser importadas
            data_handle.copiar_notas_a_importar()
            
            #Puxar restante do SAP
            sap_handler.open_zcsec_texto_nota(ses)
            sap_handler.open_zsec_click_relat(ses)
            sap_handler.sap_logout(ses)
            # Close session

            close_programs()
            
            time.sleep(12)

#            # Download GPM.csv
#            web_crawler.start_chorme_driver()
#            web_crawler.log_into_gpm("ba")
#            web_crawler.consulta_servicos_gpm()
#            web_crawler.quit_chrome_driver()
#
#            # Create files
#            data_handle.create_nimport_file()
            data_handle.create_coord_file()
            data_handle.create_religas_file()
            data_handle.data_processing(variante)

            #while(True):
            #    try :
            #        valor = datetime.strftime(input("Insira a datahora desejada"), "%d/%m/%y %H:%M:%S")
            #        print(str(valor))
            #        print(str(data_handle.proximo_dia_util(valor)))
            #    except Exception as ex:
            #        print(ex)
            #        pass
            
            # Upload files
            web_crawler.start_chorme_driver()
            web_crawler.log_into_gpm("ba")
            web_crawler.open_servicos_sap_2_bd("ba")
            web_crawler.check_import_old_files_gpm()
            web_crawler.check_import_files()
            web_crawler.quit_chrome_driver()

            try:
                date_hour_att = {'Ultima_att': [
                    datetime.now().strftime('%d/%m/%Y %H:%M:%S')]}
                hora_att = pd.DataFrame(date_hour_att)
                wks = gsheets.worksheet_select(
                    'Controle acessos usuários/robôs', 'att_religa')
                gsheets._worksheet_clear(wks)
                gsheets._worksheet_update(wks, hora_att)
            except Exception as e:
                print(e)
                close_programs()

            fim(o, hora_i, TIME_RESTART)

            return True

        except Exception as e:
            print(e)
            telegram.bot_message(
                "erro no religas/cortes - BA, reiniciando")
            close_programs()


def terminate_process(p: multiprocessing.Queue):
    p.terminate()
    p.join()


def clear_queues(q: multiprocessing.Queue, o: multiprocessing.Queue):
    while not q.empty():
        q.get()
    while not o.empty():
        o.get()


def restart_process(q: multiprocessing.Queue, o: multiprocessing.Queue):
    p = multiprocessing.Process(
        target=app, name="executar_robo", args=(q, o))
    p.start()
    return p, q.get()


def calculate_wait_time(now, final, time_restart):
    wait_restart = datetime.strptime(
        now, '%d-%m-%Y %H:%M:%S') - datetime.strptime(final, '%d-%m-%Y %H:%M:%S')
    time_wait = timedelta(seconds=time_restart) - wait_restart
    return time_wait.seconds


def sleep_count(seconds_to_wait):
    i = 0
    count = round(seconds_to_wait)
    while seconds_to_wait > i:
        print(' ########## Requisição de dados finalizada, retornando em ',
              str(timedelta(seconds=count)), 'minutos ########## ', end="")
        time.sleep(1)
        count = count - 1
        i += 1
        delete_last_line()


def delete_last_line():
    """
    Use this function to delete the last line in the STDOUT
    """

    print("\r", end="")


def main():
    print("Bom dia, o bender acordou!")

    telegram.bot_message('Bom dia, o bender acordou!')

    q = multiprocessing.Queue()
    o = multiprocessing.Queue()

    p = multiprocessing.Process(
        target=app, name="executar_robo", args=(q, o))
    p.start()

    time.sleep(3)
    inicio = q.get()

    while True:
        now = datetime.now().strftime('%d-%m-%Y %H:%M:%S')
        tempo_corrido = (datetime.strptime(now, '%d-%m-%Y %H:%M:%S') -
                         datetime.strptime(inicio, '%d-%m-%Y %H:%M:%S'))

        if tempo_corrido > timedelta(seconds=TIME_RESET) and o.empty():
            if p.is_alive():
                telegram.bot_message("Tempo excedido, reiniciando")
                terminate_process(p)
                clear_queues(q, o)
                p, inicio = restart_process(q, o)
            else:
                clear_queues(q, o)
                p, inicio = restart_process(q, o)
        elif not o.empty():
            final = o.get()
            clear_queues(q, o)
            time_wait = calculate_wait_time(now, final, TIME_RESTART)
            print(f'Esperando {time_wait} segundos para reiniciar')
            sleep_count(time_wait)
            inicio = q.get() if not q.empty() else '01-01-2022 00:00:00'
            


if __name__ == "__main__":
    multiprocessing.freeze_support()
    main()
