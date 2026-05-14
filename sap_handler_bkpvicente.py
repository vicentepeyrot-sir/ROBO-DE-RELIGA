import os
import traceback
import win32com.client
import subprocess

import pyautogui as pya
import pandas as pd
import pygetwindow as gw

from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timedelta
from time import sleep

from directories import Directories
from image_finder import ImageFinder
import config
import shutil


class SapHandler:
    
    def __init__(self, _dir: Directories, image_finder: ImageFinder) -> None:
        self._dir = _dir
        self.image_finder = image_finder

    def sap_login(self, sap_user, sap_pass) -> win32com.client.CDispatch:
        try:
            subprocess.check_call(
                f'{config.SAP_EXE} -user={sap_user} -pw={sap_pass} -system={config.SAP_SID} -client={config.SAP_MANDANT}')

            # Verifica a imagem de segurança e se aparecer clicar em ok
            if self.image_finder.find_image("seg_sap_gui", 5, 0.8):
                sleep(3)
                self.image_finder.click("seg_sap_gui", 0.8)
                pya.hotkey('alt', 'p')

            # thread para verificar existencia das telas sap já logado ou multiplo logon
            self.handle_logon_image(self.image_finder)

            try:
                sleep(5) #espera um pouco, por que estava gerando erro às vezes
                sap_gui_auto = win32com.client.GetObject('SAPGUI')
                application = sap_gui_auto.GetScriptingEngine
                connection = application.Children(0)
                session = connection.Children(0)
            except Exception as ex:
                session = None
                for i in range(100):
                    try:
                        session = application.findById(f"con[{i}]/ses[0]")
                        break
                    except Exception:
                        print(
                            f"Cannot find session in child({i}) try again next child")

            if not session:
                print("Cannot find session in any child. Exit")
                self.err_screenshot()
                raise RuntimeError("Cannot find session in any child. Exit")

            if session.children.count > 1:
                self.handle_session(session)

            print("Logged to SAP.")
            return session

        except Exception as ex:
            print(f"Error. Cannot create session. Exit {ex}")
            self.err_screenshot()
            raise RuntimeError("Cannot create session. Exit")

    def handle_logon_image(self, image_finder):
        with ThreadPoolExecutor(max_workers=2) as executor:
            future1 = executor.submit(
                image_finder.find_image, "SAP_logged", 30)
            future2 = executor.submit(
                image_finder.find_image, "multiplo_logon", 30)
            futures = [future1, future2]

            for future in as_completed(futures):
                if future.result():
                    for f in futures:
                        if f != future:
                            image_finder.cancel()
                    break
            else:
                print("Cannot find SAP window. Exit")
                self.err_screenshot()
                raise RuntimeError("Cannot find SAP window. Exit")

    def handle_session(self, session):
        try:
            if (session.children(1).text) == "Informação de licença em logon múltiplo":
                self.handle_multi_logon(session)
            elif session.children(1).text == "Mensagem do sistema":
                self.handle_system_message(session)
            else:
                print("Unknown window. Exit.")
                self.err_screenshot()
                raise RuntimeError("Unknown window. Exit.")
        except Exception as ex:
            print(f"Unknown error. Exit. {ex}")
            self.err_screenshot()
            raise RuntimeError("Unknown error. Exit.")

    def handle_multi_logon(self, session):
        try:
            session.findById("wnd[1]/usr/radMULTI_LOGON_OPT1").select()
            session.findById("wnd[1]/usr/radMULTI_LOGON_OPT1").setFocus()
            session.findById("wnd[0]").sendVKey(0)
        except Exception as ex:
            print(f"Cannot perform an operation. Exit. {ex}")
            self.err_screenshot()
            raise RuntimeError("Cannot perform an operation. Exit.")

    def handle_system_message(self, session):
        try:
            print("%s | %s" % (session.findById("wnd[1]/usr/lbl[4,1]").text,
                               session.findById("wnd[1]/usr/lbl[17,1]").text))
            print("%s | %s" % (session.findById("wnd[1]/usr/lbl[4,3]").text,
                               session.findById("wnd[1]/usr/lbl[17,3]").text))
            session.findById("wnd[1]").sendVKey(0)
        except Exception as ex:
            print(f"Unknown window. Exit. {ex}")
            self.err_screenshot()
            raise RuntimeError("Unknown window. Exit.")

    def maximize(self, session):
        if config.ICONIFY:
            session.findById("wnd[0]").iconify()
        else:
            session.findById("wnd[0]").maximize()

    def sap_logout(self, session: win32com.client.CDispatch):
        session.findById("wnd[0]").close()
        session.findById("wnd[1]/usr/btn[0]").press()
        os.system('wmic process where name="saplogon.exe" delete')

    def err_screenshot(self, name=None):
        if name is None:
            name = self._dir.dir_log
        pya.screenshot(os.path.join(
            name, f'err_screen{datetime.now().strftime("%Y%m%d%H%M%S")}.jpg'))

    def export_relatorio_leitura(self, utd: str, date: datetime, region_in: str, region_in_out: str, filename: str, session: win32com.client.CDispatch):
        try:
            date_to_input = date.strftime("%d%m%Y")
            date_file_name = date.strftime("%Y%m%d")
            session.startTransaction("ZCMOB_RELATORIO")
            #session.findById("wnd[0]/tbar[0]/okcd").text = "ZCMOB_RELATORIO"
            #session.findById("wnd[0]").sendVKey(0)
            session.findById(
                "wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[0,21]").text = utd
            session.findById("wnd[0]").sendVKey(0)
            # Regiao de leitura
            session.findById("wnd[0]/usr/ctxtP_MRC-LOW").text = region_in
            session.findById("wnd[0]/usr/ctxtP_MRC-HIGH").text = region_in_out
            # Data de leitura
            session.findById("wnd[0]/usr/ctxtP_PLAN-LOW").text = date_to_input
            session.findById("wnd[0]/usr/ctxtP_PLAN-HIGH").text = date_to_input
            session.findById("wnd[0]/usr/radP_DET").select()
            session.findById("wnd[0]/tbar[1]/btn[8]").press()
            if ImageFinder().find_image("sap_error", 3):
                session.findById("wnd[1]").sendVKey(12)
                return
            session.findById(
                "wnd[0]/usr/cntlCC/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
            session.findById(
                "wnd[0]/usr/cntlCC/shellcont/shell").selectContextMenuItem("&PC")
            session.findById(
                "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").select()
            session.findById("wnd[1]/tbar[0]/btn[0]").press()

            cols_to_use = ['Instal', 'Local', 'Bairro', 'Rua',
                           'NomeCliente', 'Nº da casa', 'Latitude', 'Longitude', 'Nota leit.']

            sleep(2)

            df = pd.read_clipboard(header=0,
                                   skiprows=5,
                                   sep='|',
                                   index_col=False,
                                   dtype=str,
                                   low_memory=False,
                                   on_bad_lines='skip')
            df.columns = [col.strip() for col in df.columns]
            df = df[cols_to_use]
            df = df.dropna(how='all')
            df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
            df = df.loc[(df['Nota leit.'] != "V100") &
                        (df['Nota leit.'] != "N160")]

            def format_func(x): return float(
                '-' + str(x).replace('-', '').replace(',', '.'))
            df['Latitude'] = df['Latitude'].apply(format_func)
            df['Longitude'] = df['Longitude'].apply(format_func)
            df.to_csv(os.path.join(self._dir.dir_temp,
                                   f'{filename}_{date_file_name}.csv'), index=False, sep=';', encoding='latin-1')

            session.findById("wnd[0]").sendVKey(12)
            
            session.EndTransaction()

        except Exception as ex:
            print("Error. Cannot export relatorio de leitura (%s)" % ex)
            traceback.print_exc()
            raise RuntimeError("Cannot export relatorio de leitura")

    def open_iw59(self, session: win32com.client.CDispatch, variante: str):
        try:
            session.StartTransaction("IW59")
            session.findById("wnd[0]/tbar[1]/btn[17]").press()
            session.findById("wnd[1]/tbar[0]/btn[12]").press()
            session.findById("wnd[0]/tbar[1]/btn[17]").press()
            session.findById("wnd[1]/usr/txt[1]").text = variante
            session.findById("wnd[1]/tbar[0]/btn[8]").press()

            if variante == "PCP_CORTE":
                tomorow = (datetime.now() + timedelta(days=1)
                           ).strftime('%d.%m.%Y')
                yesterday = (datetime.now() - timedelta(days=1)
                             ).strftime('%d.%m.%Y')
                session.findById("wnd[0]/usr/ctxt[18]").text = yesterday
                session.findById("wnd[0]/usr/ctxt[19]").text = tomorow

            session.findById("wnd[0]/tbar[1]/btn[8]").press()
            session.findById("wnd[0]/tbar[1]/btn[16]").press()
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
            session.findById("wnd[1]/usr/sub/2/sub/2/1/rad[0,0]").select()
            session.findById("wnd[1]/usr/sub/2/sub/2/1/rad[0,0]").setFocus()
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
            session.findById("wnd[1]/tbar[0]/btn[0]").press()

            excel_window = gw.getWindowsWithTitle('Excel')[0]  # Assumindo que há apenas uma janela do Excel aberta

            # Maximizar a janela do Excel
            excel_window.maximize()
            sleep(2)
            pya.moveTo(59, 228)
            pya.click()
            sleep(2)
            pya.press('f12')
            self.image_finder.find_image("excel_salvar_como", 10)
            self.image_finder.click("excel_salvar_como", 0.8)
            pya.write(os.path.join(
                self._dir.dir_temp, 'RELIGAS_EXCEL'))
            pya.press('enter')
            sleep(2)
            pya.hotkey('ctrl', 'space')
            if self.image_finder.find_image("excel_coluna_A", 10):
                pya.hotkey('ctrl', 'c')
                sleep(2)
                session.findById("wnd[1]/tbar[0]/btn[0]").press()
            session.EndTransaction()
            #session.findById("wnd[0]/tbar[0]/btn[12]").press()
            #session.findById("wnd[0]/tbar[0]/btn[12]").press()
            excel_window.close() #Fecha o Excel, pois tem gerado erros repetitivamente
            
        except Exception as ex:
            print("Error. Cannot open IW59 (%s)" % ex)
            traceback.print_exc()
            raise RuntimeError("Cannot open IW59")

    def open_zcsec_texto_nota(self, session: win32com.client.CDispatch):
        try:
            session.StartTransaction("ZCSEC_TEXTO_NOTA")
            session.findById("wnd[0]/usr/ctxt[0]").text = "a"
            session.findById("wnd[0]/usr/btn[0]").press()
            session.findById("wnd[1]/tbar[0]/btn[16]").press()
            session.findById("wnd[1]/tbar[0]/btn[24]").press()
            session.findById("wnd[1]/tbar[0]/btn[8]").press()
            session.findById("wnd[0]/tbar[1]/btn[8]").press()
            if self.image_finder.find_image("sap_error", 3):
                raise RuntimeError("sem_dados_zcsec")
            aguardar = True
            limitar = 0
            while (aguardar) :
              if limitar > 10 :
                  aguardar = False
              try:
                session.findById("wnd[0]/tbar[1]/btn[21]").press()
                aguardar = False
              except:
                if self.image_finder.find_image("sap_error", 1):
                  raise RuntimeError("sem_dados_zcsec")
                limitar += 1
                if not aguardar:
                    raise RuntimeError("Botao de exportar para Excel nao encontrado")
            
            session.findById(
                "wnd[1]/usr/ctxt[0]").text = self._dir.dir_temp
            session.findById("wnd[1]/usr/ctxt[1]").text = "RELIGAS_TXT.TXT"
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
            session.EndTransaction()
            
        except Exception as ex:
            print("Error. Cannot open zcsec_text_nota (%s)" % ex)
            traceback.print_exc()
            raise RuntimeError("Cannot open zcsec_text_nota")

    def open_zsec_click_relat(self, session: win32com.client.CDispatch):
        try:
            session.StartTransaction("ZSEC_CLICK_RELAT")
            #session.findById("wnd[0]/tbar[0]/okcd").text = "ZSEC_CLICK_RELAT"
            #session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/usr/ctxt[2]").text = "a"
            session.findById("wnd[0]/usr/btn[1]").press()
            session.findById("wnd[1]/tbar[0]/btn[16]").press()
            session.findById("wnd[1]/tbar[0]/btn[24]").press()

            if self.image_finder.find_image("sap_error", 3):
                shutil.copy(self._dir.dir_temp + "\\..\\RELIGAS_COORDS.TXT", self._dir.dir_temp + "\\RELIGAS_COORDS.TXT")
                print("SEM COORDENADAS")
                return

            session.findById("wnd[1]/tbar[0]/btn[8]").press()
            session.findById("wnd[0]/tbar[1]/btn[8]").press()
            session.findById("wnd[0]/tbar[1]/btn[33]").press()
            
            layout = None
            
            if self.image_finder.find_image("sap_error", 3):
                print("SEM COORDENADAS")
                shutil.copy(self._dir.dir_temp + "\\..\\RELIGAS_COORDS.TXT", self._dir.dir_temp + "\\RELIGAS_COORDS.TXT")
                return
            
            while layout is None:
                try:
                    layout = session.findById(
                        "wnd[1]/usr/ssub/1/cntlG51_CONTAINER/shellcont/shell")
                except:
                    pass
                sleep(0.25)

            rows = -1
            while rows < 0 :
              try :
                rows = layout.rowCount
              except:
                sleep(0.25)

            for i in range(0, rows):
                layout_variant = layout.getCellValue(i, "VARIANT")
                if layout_variant == "///SIRTEC":
                    layout.currentCellRow = i
                    layout.firstVisibleRow = i
                    layout.selectedRows = i
                    layout.clickCurrentCell()
                    break
            session.findById("wnd[0]/tbar[1]/btn[45]").press()
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
            session.findById("wnd[1]/usr/ctxt[0]").text = self._dir.dir_temp
            session.findById("wnd[1]/usr/ctxt[1]").text = "RELIGAS_COORDS.TXT"
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
            session.EndTransaction()
            #session.findById("wnd[0]/tbar[0]/btn[12]").press()
            #session.findById("wnd[0]/tbar[0]/btn[12]").press()
        except Exception as ex:
            print("Error. Cannot open zsec_click_relat (%s)" % ex)
            traceback.print_exc()
            raise RuntimeError("Cannot open zsec_click_relat")
