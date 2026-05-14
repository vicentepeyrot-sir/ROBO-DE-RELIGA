import os
import re

import pandas as pd
import numpy as np
import traceback
from datetime import datetime, timedelta

from directories import Directories
from gdrive import GDrive
from gsheets import Psheets

from bot import Bot

import pyperclip

from time import sleep


# BD_ID
BD_ID = '1MBNEZHglGpOac1KvOJuIeigi3dchqpWm'
FINAL_FOLDER_ID = '1kHKg_w8Gv85oxHyEKDWGyP2k8mtvV_kU'
TRATATIVAS_FOLDER_ID = '1BmvqbncEf664PnWFwgs-1Dzoi8IRgXe_'


class DataHandling:
    

    def __init__(self, _dir: Directories, gsheets: Psheets, gdrive: GDrive, telegram: Bot) -> None:
        self._dir = _dir
        self.gsheets = gsheets
        self.gdrive = gdrive
        
        self.telegram = telegram
        
        self.feriadoes = pd.DataFrame([])

        # List of files to be downloaded if they don't exist
        self.bd_bar_bjl = os.path.join(_dir.dir_data, 'bd_BAR_BJL.csv')
        self.bd_bru_vtc_jeq_ire = os.path.join(
            _dir.dir_data, 'bd_BRU_VTC_JEQ_IRE.csv')
        self.bd_fsa = os.path.join(_dir.dir_data, 'bd_FSA.csv')
        self.bd_ser_ita_gua = os.path.join(_dir.dir_data, 'bd_SER_ITA_GUA.csv')

        # file paths
        self.file_iw59 = os.path.join(_dir.dir_temp, 'RELIGAS_EXCEL.xlsx')
        self.file_zcsec = os.path.join(_dir.dir_temp, 'RELIGAS_TXT.TXT')
        self.file_coords = os.path.join(_dir.dir_temp, 'RELIGAS_COORDS.TXT')
        self.file_viqmel = os.path.join(_dir.dir_temp, 'VIQMEL.csv')
        self.file_adrc = os.path.join(_dir.dir_temp, 'ADRC.csv')
        self.file_zcsec_csv = os.path.join(_dir.dir_temp, 'RELIGAS_TXT.csv')
        self.file_nimport_csv = os.path.join(_dir.dir_temp, 'nimport.csv')
        self.file_coords_csv = os.path.join(
            _dir.dir_temp, 'RELIGAS_COORDS.csv')
        self.file_gpm = os.path.join(_dir.dir_temp, 'GPM.csv')
        self.file_config = os.path.join(_dir.dir_data, 'bd_config.xlsx')
        self.bd_bar_bjl = os.path.join(_dir.dir_data, 'bd_BAR_BJL.csv')
        self.bd_bru_vtc_jeq_ire = os.path.join(
            _dir.dir_data, 'bd_BRU_VTC_JEQ_IRE.csv')
        self.bd_fsa = os.path.join(_dir.dir_data, 'bd_FSA.csv')
        self.bd_ser_ita_gua = os.path.join(_dir.dir_data, 'bd_SER_ITA_GUA.csv')
        
        
    def copiar_notas_a_importar(self) :
            dir_path = os.getcwd() + os.path.sep
            os.chdir(dir_path)

            # Carrega arquivo file_gpm
            print('#Carregando arquivo GPM')
            gpm_df = pd.read_csv(self.file_gpm, sep=";",
                                 on_bad_lines='skip',
                                 encoding='utf-8-sig',
                                 low_memory=False)

            gpm_df['Nota'] = gpm_df['Nota'].astype(str)

            
            # filtrar as notas abertas e concluidas do gpm
            gpm_abertas = gpm_df.query(
                'situacao_servico == "Aberto" or situacao_servico == "Concluido"')

            # Carrega arquivo file_iw59
            print('#Carregando arquivo IW59')

            iw59_df = pd.read_excel(self.file_iw59,
                                    sheet_name=0)

            iw59_df['Nota'] = iw59_df['Nota'].astype(str)
            
            if len(iw59_df.index) <= 3000 :
                pyperclip.copy("\r\n".join(iw59_df['Nota'].tolist()))
                sleep(1)
                return

            # Baixar CSV Não Importar e criar DF
            print('#Carregando arquivo Não importar')
            nimport_df = pd.read_csv(self.file_nimport_csv, sep=";",
                                     on_bad_lines='skip',
                                     encoding='utf8')

            nimport_df['Nota'] = nimport_df['Nota'].astype(str)

            # remove as notas que estao no nao import do IW59
            print('#Removendo notas da não import')
            iw59_df = iw59_df[~iw59_df['Nota'].isin(nimport_df['Nota'])]

            # Traz a as notas abertas e concluidas para o arquivo da IW59 que já esta com as observações e remove as notas abertas
            print('#Removendo notas abertas da importação')
            iw59_df = pd.merge(iw59_df,
                                  gpm_abertas,
                                  how='left',
                                  on='Nota')
            iw59_df = iw59_df.query('situacao_servico != "Aberto"')


            # Remove as notas onde a data do GPM é maior ou igual ao do SAP e remove as notas sem prazo (checando se precisa reimportar alguma nota já concluida no GPM)
            print(
                '#Removendo notas sem prazo e carregando somente notas com prazo do SAP maior que as do GPM')
            iw59_df = iw59_df.rename(
                columns={'Concl/desejada': 'concl_desejada'})
            iw59_df = iw59_df.rename(
                columns={'Concl.desejada': 'concl_desejada'})
            iw59_df = iw59_df.rename(
                columns={'Conclusão desejada': 'concl_desejada'})
            iw59_df = iw59_df.rename(
                columns={'Hora final desejada': 'H final desej.'})

            # data e hora GPM
            iw59_df['prazo_GPM'] = pd.to_datetime(
                iw59_df['prazo_servico'], dayfirst=True)
            iw59_df['prazo_SAP'] = pd.to_datetime(iw59_df['concl_desejada'].astype(str) + ' ' + iw59_df['H final desej.'].astype(str),
                                                     errors='coerce')

            iw59_df['prazo_SAP'] = iw59_df['prazo_SAP'].map(
                lambda x: x.replace(second=0))
            iw59_df = iw59_df.query(
                'prazo_SAP != "" & prazo_GPM == "" | prazo_SAP > prazo_GPM')
            if iw59_df.empty :
                self.telegram.bot_message("Sem notas para importação")
                raise RuntimeError("There are no changes to do")
            pyperclip.copy("\r\n".join(iw59_df['Nota'].tolist()))
            sleep(1)


    def bd_coords_download(self):
        """Checa se os arquivos existem, caso contrário, baixa eles do Gdrive

        Args:
            dir_data (_type_): _description_
            bd_bar_bjl (_type_): _description_
            bd_bru_vtc_jeq_ire (_type_): _description_
            bd_fsa (_type_): _description_
            bd_ser_ita_gua (_type_): _description_
        """
        files = [self.bd_bar_bjl, self.bd_bru_vtc_jeq_ire,
                 self.bd_fsa, self.bd_ser_ita_gua]
        file_dont_exist: bool = False
        for bd_file in files:
            if os.path.isfile(bd_file) != True:
                file_dont_exist = True

        if file_dont_exist:
            gdrive_content = self.gdrive._list_folder_content(BD_ID)
            data = list(gdrive_content.items())
            columns = ['title', 'id']
            df_gdrive_bd = pd.DataFrame(data, columns=columns)

        # Check existence of files
            for bd_file in files:
                if os.path.isfile(bd_file) != True:
                    print(f"Arquivo de BD não encontrado: {bd_file}")
                    real_name_file = bd_file.split(os.path.sep)[-1]
                    file_id = df_gdrive_bd.loc[df_gdrive_bd['title']
                                               == real_name_file, 'id'].values[0]
                    print(f"Baixando arquivo: {real_name_file}")
                    self.gdrive._download_file(
                        file_id, bd_file)
                    print(f"Arquivo baixado com sucesso: {bd_file}")
        else:
            print("Todos os arquivos de BD existem")

    def create_coord_file(self):
        # Tratando arquivo de coordenadas
        print('#Tratando arquivo de coordenadas')
        coords_df = pd.read_csv(os.path.join(self._dir.dir_temp, 'RELIGAS_COORDS.TXT'),
                                names=['index',
                                       'Nota',
                                       'Latitude',
                                       'Longitude'],
                                encoding='latin-1',
                                header=None,
                                skiprows=6,
                                skipinitialspace=True,
                                sep='|',
                                index_col=False,
                                dtype=str)
        coords_df = coords_df[['Nota', 'Latitude', 'Longitude']].dropna()
        coords_df['Latitude'] = "-" + coords_df['Latitude'].str.rstrip('-')
        coords_df['Longitude'] = "-" + coords_df['Longitude'].str.rstrip('-')
        coords_df.to_csv(os.path.join(self._dir.dir_temp, 'RELIGAS_COORDS.csv'),
                         sep=';',
                         encoding='utf-8-sig',
                         index=False)

    def create_religas_file(self):
        # Tratando o arquivo do ZCSEC e salva em csv
        print('#Tratando arquivo de TXT')

        txt_df = pd.read_csv(os.path.join(self._dir.dir_temp, 'RELIGAS_TXT.TXT'),
                             sep="\t",
                             names=['Nota',
                                    'Tipo da Nota',
                                    'Texto',
                                    'Vazio'],
                             on_bad_lines='skip',
                             encoding='latin-1',
                             engine='python',
                             dtype=str)

        txt_df = txt_df.drop(0)
        # Trata a Coluna NOTA do TXT retira os espaços e transforma em STRING (Não consegui transformar para valor pq algumas linhas tem textos)
        txt_df['Nota'] = txt_df['Nota'].astype(str)
        txt_df['Nota'] = txt_df['Nota'].str.strip()
        txt_df['Tipo da Nota'] = txt_df['Tipo da Nota'].astype(str)
        txt_df['Tipo da Nota'] = txt_df['Tipo da Nota'].str.strip()
        txt_df['Texto'] = txt_df['Texto'].astype(str)
        txt_df['Texto'] = txt_df['Texto'].str.strip()
        txt_df.to_csv(os.path.join(self._dir.dir_temp + 'RELIGAS_TXT.csv'),
                      sep=';',
                      encoding='utf-8-sig',
                      index=False)

    def create_nimport_file(self):
        # Baixando e criando o arquivo de nimport e salvando em csv
        print('#Tratando arquivo de nimport')
        nimport_url = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vSdYjH0obp2tqFvQS6qnA5eGiava-HJIM_Z1A5RRNqpn0O8nsyUtKPeKilemyf60IgC8RrnQFG-B5ru/pub?gid=1341271327&single=true&output=csv'

        # Lê sem cabeçalho fixo para localizar dinamicamente a linha que contém "Nota"
        nimport_raw = pd.read_csv(
            nimport_url,
            sep=",",
            encoding='utf-8',
            skip_blank_lines=True,
            header=None,
            dtype=str,
            engine='python'
        )

        # Remove linhas e colunas totalmente vazias
        nimport_raw = nimport_raw.dropna(axis=0, how='all').dropna(axis=1, how='all')
        nimport_raw = nimport_raw.fillna('')

        # Procura a linha de cabeçalho real (onde exista a coluna Nota)
        header_idx = None
        for idx, row in nimport_raw.iterrows():
            valores = [str(v).strip() for v in row.tolist()]
            if 'Nota' in valores:
                header_idx = idx
                break

        if header_idx is None:
            print(repr(nimport_raw.head(10).values.tolist()))
            raise KeyError("Coluna 'Nota' não encontrada no arquivo nimport")

        nimport_df = nimport_raw.iloc[header_idx + 1:].copy()
        nimport_df.columns = [str(col).strip() for col in nimport_raw.iloc[header_idx].tolist()]

        print(repr(nimport_df.columns.tolist()))

        # Remove colunas sem nome útil
        nimport_df = nimport_df.loc[:, [
            str(col).strip() != '' and not str(col).startswith('Unnamed')
            for col in nimport_df.columns
        ]]
        nimport_df = nimport_df.dropna(axis=1, how='all')
        nimport_df = nimport_df.dropna(axis=0, how='all')

        nimport_df.columns = [str(col).strip() for col in nimport_df.columns]

        if 'Nota' not in nimport_df.columns:
            raise KeyError("Coluna 'Nota' não encontrada após normalização do nimport")

        nimport_df['Nota'] = nimport_df['Nota'].fillna('').astype(str).str.strip()
        nimport_df = nimport_df[nimport_df['Nota'] != '']
        nimport_df = nimport_df.reset_index(drop=True)

        # Normaliza Nota para evitar .0 no final
        try:
            nimport_df['Nota'] = pd.to_numeric(nimport_df['Nota'], errors='raise').astype(np.int64).astype(str)
        except Exception:
            nimport_df['Nota'] = nimport_df['Nota'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()

        nimport_df.to_csv(os.path.join(self._dir.dir_temp + 'nimport.csv'),
                          sep=';',
                          encoding='utf-8-sig',
                          index=False)

    def data_processing(self, variante):
        """
        Baixa os dados faltantes e trata os arquivos para importação
        """
        try:
            dir_path = os.getcwd() + os.path.sep
            os.chdir(dir_path)

            # Carrega as páginas das configurações usadas
            wks = self.gsheets.worksheet_select(
                'bd_config_robo_religacao', 'municipios')

            municipios_df = pd.DataFrame(wks.get_as_df(), dtype=str)

            wks = self.gsheets.worksheet_select(
                'bd_config_robo_religacao', 'tensao')

            tensao_df = pd.DataFrame(wks.get_as_df())

            # Carrega arquivo file_gpm
            print('#Carregando arquivo GPM')
            gpm_df = pd.read_csv(self.file_gpm, sep=";",
                                 on_bad_lines='skip',
                                 encoding='utf-8-sig',
                                 low_memory=False)

            gpm_df['Nota'] = gpm_df['Nota'].astype(str)

            # filtrar as notas abertas e concluidas do gpm
            gpm_abertas = gpm_df.query(
                'situacao_servico == "Aberto" or situacao_servico == "Concluido"')

            # Carrega arquivo file_iw59
            print('#Carregando arquivo IW59')

            iw59_df = pd.read_excel(self.file_iw59,
                                    sheet_name=0)

            iw59_df['Nota'] = iw59_df['Nota'].astype(str)

            # Carrega o arquivo ZCSEC (TXT) em csv
            print('#Carregando arquivo TXT')
            txt_df = pd.read_csv(self.file_zcsec_csv, sep=";",
                                 on_bad_lines='skip',
                                 encoding='utf8')
            
            # resolve bug do tipo object x string
            txt_df['Nota'] = txt_df['Nota'].astype(str)
            txt_df['Tipo da Nota'] = txt_df['Tipo da Nota'].fillna('').astype(str).str.strip()

            # Adiciona os dados dos valores do CS e as faixas e codes
            cortes_df = txt_df[txt_df['Tipo da Nota'].str.contains('CS|CA', regex=True, na=False)].copy()
            cortes_df['Valores'] = cortes_df.apply(
                self.extrair_debitos, axis=1)
            cortes_df['code_cortes'] = cortes_df.apply(
                self.adicionar_codes_cs, axis=1)

            # Baixar CSV Não Importar e criar DF
            print('#Carregando arquivo Não importar')
            nimport_df = pd.read_csv(self.file_nimport_csv, sep=";",
                                     on_bad_lines='skip',
                                     encoding='utf8')

            nimport_df['Nota'] = nimport_df['Nota'].astype(str)

            # remove as notas que estao no nao import do IW59
            print('#Removendo notas da não import')
            iw59_df = iw59_df[~iw59_df['Nota'].isin(nimport_df['Nota'])]

            # Traz a OBSERVAÇÃO para o arquivo gerado da IW59
            print('#Adicionando observação nas notas')
            sap_concat = pd.merge(iw59_df, txt_df[['Nota',
                                                   'Texto']],
                                  how='left',
                                  on='Nota')

            # Traz a as notas abertas e concluidas para o arquivo da IW59 que já esta com as observações e remove as notas abertas
            print('#Removendo notas abertas da importação')
            sap_concat = pd.merge(sap_concat,
                                  gpm_abertas,
                                  how='left',
                                  on='Nota')
            sap_concat = sap_concat.query('situacao_servico != "Aberto"')

            # Traz as UTDs de trabalho para o arquivo IW59
            print('#Adicionando as UTDs')
            sap_concat = pd.merge(sap_concat, municipios_df[['Local',
                                                            'UTD']],
                                  how='left',
                                  on='Local')

            # Remove as notas onde a data do GPM é maior ou igual ao do SAP e remove as notas sem prazo
            # (checando se precisa reimportar alguma nota já concluida no GPM).
            # Para cortes (CS), não aplica o filtro de prazo.
            print(
                '#Removendo notas sem prazo e carregando somente notas com prazo do SAP maior que as do GPM')
            sap_concat = sap_concat.rename(
                columns={'Concl/desejada': 'concl_desejada'})
            sap_concat = sap_concat.rename(
                columns={'Concl.desejada': 'concl_desejada'})
            sap_concat = sap_concat.rename(
                columns={'Conclusão desejada': 'concl_desejada'})
            sap_concat = sap_concat.rename(
                columns={'Hora final desejada': 'H final desej.'})

            sap_concat['prazo_GPM'] = pd.to_datetime(
                sap_concat['prazo_servico'], dayfirst=True, errors='coerce')

            data_sap = sap_concat['concl_desejada'].fillna('').astype(str).str.strip()
            hora_sap = sap_concat['H final desej.'].fillna('').astype(str).str.strip()

            # Quando concl_desejada já vem com data e hora completas (ex.: 2026-03-17 10:20:53),
            # usa direto esse valor. Caso contrário, concatena com a hora final desejada.
            prazo_sap_str = np.where(
                data_sap.str.contains(r'\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2}', regex=True),
                data_sap,
                (data_sap.str.replace('/', '.', regex=False) + ' ' + hora_sap)
            )
            prazo_sap_str = pd.Series(prazo_sap_str, index=sap_concat.index).astype(str).str.strip()
            prazo_sap_str = prazo_sap_str.str.replace(r'^(\d{2}:\d{2})$', r'\1:00', regex=True)
            prazo_sap_str = prazo_sap_str.replace('', np.nan)

            sap_concat['prazo_SAP'] = pd.to_datetime(
                prazo_sap_str,
                errors='coerce',
                dayfirst=True
            )

            mask_prazo_ok = sap_concat['prazo_SAP'].notna()
            sap_concat.loc[mask_prazo_ok, 'prazo_SAP'] = sap_concat.loc[
                mask_prazo_ok, 'prazo_SAP'
            ].apply(lambda x: x.replace(second=0))

            sap_concat = sap_concat[
                (sap_concat['Tipo de nota'] == 'CS') |
                (
                    sap_concat['prazo_SAP'].notna() & (
                        sap_concat['prazo_GPM'].isna() |
                        (sap_concat['prazo_SAP'] > sap_concat['prazo_GPM'])
                    )
                )
            ].copy()

            # filtrar e remover as notas cujo retorno de campo seja diferente de 600 e 9000
            print('#Removendo notas com retorno de concluidas e finalizadas')
            try:
                sap_concat['num_grupo_retorno_campo'] = (
                    sap_concat['Grupo Retorno de Campo']
                    .fillna('')
                    .astype(str)
                    .str.extract(r'(\d+)', expand=False)
                    .fillna('')
                    .str.strip()
                )

                sap_remove = sap_concat[
                    (sap_concat['num_grupo_retorno_campo'] != '') &
                    (~sap_concat['num_grupo_retorno_campo'].isin(['600', '9000']))
                ][['Nota']].copy()

                sap_concat = sap_concat[
                    ~sap_concat['Nota'].isin(sap_remove['Nota'])
                ].copy()

            except Exception:
                print('#Sem notas com esse tipo de retorno')

            if sap_concat.empty :
                self.telegram.bot_message("Sem notas para importação")
                raise RuntimeError("There are no changes to do")
            
            # colocando nome "DESCONHECIDO" nas celulas onde o nome está em branco
            print('#Checando nome do clientes e complementando com DESCONHECIDO')
            sap_concat['Nome do parceiro'] = sap_concat['Nome do parceiro'].replace(
                np.nan, 'DESCONHECIDO')
            
            sap_concat['Nome do parceiro'] = sap_concat['Nome do parceiro'].apply(lambda par : par if len(par) >= 4 else "DESCONHECIDO")

            # Adicionando os dados de tensão de medição para gerar a descrição
            print('#Adicionando dados da tensão de medição')

            sap_concat = pd.merge(sap_concat, tensao_df[['Tensão de medição',
                                                        'Fases']],
                                  how='left',
                                  on='Tensão de medição')

            # Gerando a descrição
            print('#Criando a descrição da nota')
            sap_concat['Descrição_iw59'] = sap_concat['Descrição']
            sap_concat['Descrição'] = sap_concat[['Fases',
                                                  'Descrição CNAE',
                                                  'Ponto de referência',
                                                  'Nº do pedido',
                                                  'Texto']].astype(str).apply('|'.join,
                                                                            axis=1)
            
            #cols = ['Fases','Descrição CNAE','Ponto de referência','Nº do pedido','Texto']
            #sap_concat['Descrição'] = sap_concat[cols].apply(lambda row: '|'.join(row.values.astype(str)), axis=1)

            # Removendo os nan das celulas que estão vazias
            sap_concat['Descrição'] = sap_concat['Descrição'].astype(str).str.replace('[|]nan',
                                                                                      '',
                                                                                      regex=True)

            # Adicionando o codes e codificação ao Tipo de nota
            print('#Adicionando os Codes nas notas')
            sap_concat['Tipo de nota iw59'] = sap_concat['Tipo de nota']
            sap_concat['Tipo de nota'] = sap_concat['Grupo de codes'].str[-6:] + \
                sap_concat['Codificação']

            # Carrega os bancos de dados das instalações
            print('#Carregando arquivos do banco de dados de localizações')
            bd_bar_bjl_df = pd.read_csv(self.bd_bar_bjl, sep=";",
                                        on_bad_lines='skip',
                                        encoding='utf8')

            bd_bru_vtc_jeq_ire_df = pd.read_csv(self.bd_bru_vtc_jeq_ire,
                                                sep=";",
                                                on_bad_lines='skip',
                                                encoding='utf8')
            bd_fsa_df = pd.read_csv(self.bd_fsa, sep=";",
                                    on_bad_lines='skip',
                                    encoding='utf8')

            bd_ser_ita_gua_df = pd.read_csv(self.bd_ser_ita_gua, sep=";",
                                            on_bad_lines='skip',
                                            encoding='utf8')

            bd_coord_df = pd.concat([bd_bar_bjl_df,
                                    bd_bru_vtc_jeq_ire_df,
                                    bd_fsa_df,
                                    bd_ser_ita_gua_df])

            bd_coord_df = bd_coord_df.query(
                'Latitude != "0" | Longitude != "0"')

            # Organizando e ajustando endereço removendo dados e ponto e vírgula e numero das casa para buscar coordenadas via endereco
            print('#Ajustando os endereços')
            sap_concat['rua_sem_num'] = sap_concat['Rua'].str.split(',').str[0].str.replace('\d+',
                                                                                            '',
                                                                                            regex=True)
            sap_concat['Cidade+Bairro+Endereço'] = sap_concat['Local'] + \
                sap_concat['Bairro_x'] + sap_concat['rua_sem_num']
            sap_concat['Cidade+Bairro+Endereço'] = sap_concat['Cidade+Bairro+Endereço'].str.strip()

            # Lendo o arquivo de coordenadas em csv
            print('#Carregando arquivo de coordenadas')
            coords_df = pd.read_csv(self.file_coords_csv, sep=";",
                                    on_bad_lines='skip',
                                    encoding='utf8')
            coords_df['Nota'] = coords_df['Nota'].astype(np.int64).astype(str)
            coords_df = coords_df.query('Latitude != 0 | Longitude != 0')

            # Buscando coordenadas via endereço
            print('#Adicionando coordendas nas notas baseado nos endereços')
            sap_concat = sap_concat.rename(columns={'Instalação_x': 'Instal'})

            sap_concat_1 = pd.merge(sap_concat, coords_df[['Nota',
                                                           'Longitude',
                                                           'Latitude']],
                                    how='left',
                                    on=['Nota'])

            sap_concat_2 = pd.merge(sap_concat, bd_coord_df[['Longitude',
                                                            'Latitude',
                                                             'Instal']],
                                    how='left',
                                    on=['Instal'])

            sap_concat = sap_concat_1.fillna(sap_concat_2).join(
                sap_concat_2[sap_concat_2.columns.drop(sap_concat_1.columns)])

            sap_concat_3 = sap_concat[sap_concat['Latitude'].isna()]
            sap_concat_3 = sap_concat_3[~sap_concat_3['UTD'].isnull()]
            sap_concat_3.drop("Latitude", axis=1, inplace=True)
            sap_concat_3.drop("Longitude", axis=1, inplace=True)

            sap_concat_3 = pd.merge(sap_concat_3, bd_coord_df[['Cidade+Bairro+Endereço',
                                                               'Longitude',
                                                               'Latitude']],
                                    how='left',
                                    on=['Cidade+Bairro+Endereço'])

            sap_concat_3 = sap_concat_3.drop_duplicates(subset=['Nota'])
            sap_concat_3 = sap_concat_3[~sap_concat_3['Latitude'].isnull()]
            sap_concat = pd.concat([sap_concat,
                                    sap_concat_3])
            sap_concat = sap_concat.sort_values(by=['Latitude', 'Longitude'])

            # # Juntando VIQMEL e ADRC e trazendo para a observação da nota
            # df_adrc = pd.read_csv(file_adrc,
            #                       sep=";",
            #                       on_bad_lines='skip',
            #                       encoding='utf8')

            # df_viqmel = pd.read_csv(file_viqmel,
            #                         sep=";",
            #                         on_bad_lines='skip',
            #                         encoding='utf8')

            # df_observacao = pd.merge(
            #     df_adrc, df_viqmel, left_on='Nº ender.', right_on='numero', how='left')

            # df_observacao = df_observacao[[
            #     'nota', 'Pto Referê', 'Bloco', 'Quadra']]

            # df_observacao.dropna(how='all', subset=[
            #                      'Pto Referê', 'Bloco', 'Quadra'], inplace=True)

            # df_observacao['rua_extra'] = df_observacao[['Pto Referê',
            #                                             'Bloco', 'Quadra']].astype(str).agg(', '.join, axis=1)
            # df_observacao['rua_extra'] = df_observacao['rua_extra'].astype(str).str.replace(
            #     'nan,', '', regex=True).str.replace('nan', '', regex=True).str.rstrip().str.lstrip().str.rstrip(',')
            # df_observacao = df_observacao[['nota', 'rua_extra']]
            # df_observacao['nota'] = df_observacao['nota'].astype(str)
            # df_observacao.rename(columns={'nota': 'Nota'}, inplace=True)

            # # Adicionando dados extras do endereço na nota
            # sap_concat = pd.merge(sap_concat, df_observacao, on='Nota', how='left')

            # sap_concat['Rua'] = sap_concat[['Rua', 'rua_extra']].astype(
            #     str).agg(', '.join, axis=1)
            sap_concat['Rua'] = sap_concat['Rua'].str.replace('nan,', '', regex=True).str.replace(
                'nan', '', regex=True).str.rstrip().str.lstrip().str.rstrip(',')

            # Adicionando tratativas
            sap_concat.loc[(sap_concat['prazo_SAP'].notna()) & (sap_concat['prazo_GPM'].notna()) & (
                sap_concat['prazo_SAP'] > sap_concat['prazo_GPM']) & (~sap_concat['UTD'].isna()), 'tratativas'] = 0
            sap_concat.loc[(sap_concat['prazo_SAP'].notna()) & (
                ~sap_concat['prazo_GPM'].notna()) & (~sap_concat['UTD'].isna()), 'tratativas'] = 0
            sap_concat.loc[(~sap_concat['prazo_SAP'].notna()) & (~sap_concat['prazo_GPM'].notna()) & (
                ~sap_concat['UTD'].isna()), ['tratativas', 'Nota sem prazo']] = [1, 'Nota sem prazo']
            sap_concat.loc[(sap_concat['UTD'].isna()), [
                'tratativas', 'Cidade não cadastrada']] = [1, 'Cidade não cadastrada']
            sap_concat.loc[(sap_concat['Descrição'].str.len() > 500), [
                'tratativas', '+500 caracteres']] = [1, '+500 caracteres']
            sap_concat.loc[(sap_concat['Codificação'] == "BESB") & (sap_concat['Grupo de codes'] == "DESLDEBT"), [
                'tratativas', 'SOLAR']] = [1, 'Nota SOLAR, verificar SAP e não enviar a campo.']

            wks = self.gsheets.worksheet_select(
                'bd_config_robo_religacao', 'texto_analise')

            sap_concat_txt_analise = pd.DataFrame(wks.get_as_df(), dtype=str)

            array_txt_analise = sap_concat_txt_analise['Texto p/ analisar'].to_list(
            )

            for txt_analise in array_txt_analise:
                sap_concat.loc[(sap_concat['Descrição'].str.contains(
                    txt_analise, case=False)), 'analise'] = "| " + txt_analise + " "
                sap_concat.loc[(sap_concat['Descrição'].str.contains(
                    txt_analise, case=False)), 'tratativas'] = 1

            # Remove as religações das tratativas
            sap_concat.loc[(sap_concat['Tipo de nota iw59']
                            == "CR"), ['tratativas']] = [0]

            sap_concat.loc[(sap_concat['Nota sem prazo'].notna()), 'analise'] = sap_concat.loc[(
                sap_concat['Nota sem prazo'].notna())]['analise'].astype(str).str.replace('nan', '') + "| Nota sem prazo "
            sap_concat.loc[(sap_concat['Cidade não cadastrada'].notna()), 'analise'] = sap_concat.loc[(
                sap_concat['Cidade não cadastrada'].notna())]['analise'].astype(str).str.replace('nan', '') + "| Cidade não cadastrada "
            sap_concat.loc[(sap_concat['+500 caracteres'].notna()), 'analise'] = sap_concat.loc[(
                sap_concat['+500 caracteres'].notna())]['analise'].astype(str).str.replace('nan', '') + "| +500 caracteres "
            sap_concat.loc[(sap_concat['SOLAR'].notna()), 'analise'] = sap_concat.loc[(sap_concat['SOLAR'].notna(
            ))]['analise'].astype(str).str.replace('nan', '') + "| Nota SOLAR, verificar SAP e não enviar a campo. "

            sap_concat.loc[(sap_concat['analise'].notna()), 'analise'] = sap_concat.loc[sap_concat['analise'].notna(
            )]['analise'].astype(str).str.replace('\A[| ] ', '', regex=True)

            # Modificando a letra do tipo de nota para M onde contem no texto substituir medidor obsoleto
            print('#Gerando religas com substituição de medidor')
            sap_concat['Tipo de nota'] = np.where(sap_concat['Descrição'].str.contains('Substituir medidor obsoleto', case=False),
                                                  sap_concat['Tipo de nota'].str.replace('R', 'M', 1), sap_concat['Tipo de nota'])

            # # Adiciona a data correta baseada no calendário para as notas devido a resolução 1.000
            # municipios_df = pd.read_excel(config_df, 'calendario')
            # municipios_df['Dia'] = pd.to_datetime(
            #     municipios_df['Dia'], format='%d/%m/%Y', dayfirst=True)
            # municipios_df['5° Dia util'] = pd.to_datetime(
            #     municipios_df['5° Dia util'], format='%d/%m/%Y', dayfirst=True).dt.strftime('%d/%m/%Y')
            # # sap_concat['Data da nota'] = pd.to_datetime(sap_concat['Data da nota'], format='%d/%m/%Y', dayfirst=True)
            # sap_concat['Data da nota'] = pd.to_datetime(
            #     sap_concat['Data da nota'], format='%Y/%m/%d', dayfirst=True)
            # sap_concat = pd.merge(sap_concat, municipios_df[[
            #                       'Dia', '5° Dia util']], how='left', left_on=['Data da nota'], right_on='Dia')
            # sap_concat.loc[(pd.to_datetime(sap_concat['Data da nota'], format='%d/%m/%Y', dayfirst=True) > datetime(2019, 1, 1)) &
            #                (sap_concat['Descrição_iw59'] == 'Vistoria / Ligação') &
            #                # (pd.to_datetime(sap_concat['5° Dia util'], format='%d/%m/%Y', dayfirst=True) > pd.to_datetime(sap_concat['concl_desejada'], format='%d/%m/%Y', dayfirst=True)) &
            #                (pd.to_datetime(sap_concat['5° Dia util'], format='%d/%m/%Y', dayfirst=True) >
            #                 pd.to_datetime(sap_concat['concl_desejada'], format='%Y/%m/%d', dayfirst=True)) &
            #                (sap_concat['Status usuário'] != 'RLIB'), 'concl_desejada'] = sap_concat['5° Dia util']

            # Adiciona 100 na descrição da nota onde na descrição da nota no iw59 for enlace CM e tiver 100 ou 101
            sap_concat = sap_concat.reset_index(drop=True)
            sap_concat.loc[(sap_concat['Tipo de nota iw59'] == 'CM') & (sap_concat['Descrição_iw59'].astype(
                str).str.contains('100 | 101')), 'Descrição'] = sap_concat['Descrição'] + '| 100'

            # Adiciona o local do corte na descrição da nota que o tipo for corte CS
            pedido_str = sap_concat['Nº do pedido'].fillna('').astype(str)
            mask_cs = sap_concat['Tipo de nota iw59'] == 'CS'

            sap_concat['local_corte'] = ''

            sap_concat.loc[
                mask_cs & pedido_str.str.contains('@D', na=False),
                'local_corte'
            ] = 'Local Corte: Disjuntor | '

            sap_concat.loc[
                mask_cs
                & ~pedido_str.str.contains('@D', na=False)
                & pedido_str.str.contains('@S', na=False),
                'local_corte'
            ] = 'Local Corte: Solo | '

            sap_concat.loc[
                mask_cs
                & ~pedido_str.str.contains('@D', na=False)
                & ~pedido_str.str.contains('@S', na=False)
                & pedido_str.str.contains('@P', na=False),
                'local_corte'
            ] = 'Local Corte: Poste | '

            if 'local_corte' in sap_concat.columns:
                sap_concat['nova_desc'] = sap_concat['local_corte'].fillna('').astype(
                    str).str.cat(sap_concat['Descrição'].fillna('').astype(str), sep='')
                sap_concat['Descrição'] = sap_concat['nova_desc']

            sap_concat = sap_concat.merge(
                cortes_df[['Nota', 'Valores', 'code_cortes']], how='left', on='Nota')
            sap_concat.loc[sap_concat['code_cortes'].notna(
            ), 'Tipo de nota'] = sap_concat.loc[sap_concat['code_cortes'].notna(), 'code_cortes']

            # Muda os tipos das notas
            sap_concat.loc[(sap_concat['Nº do pedido'].astype(str).str.contains(
                '@TOP')) & (sap_concat['Tipo de nota iw59'] == 'CS'), 'Tipo de nota'] = 'TOP25-'

            sap_concat.loc[(sap_concat['Nº do pedido'].astype(str).str.contains(
                '@FT')) & (sap_concat['Tipo de nota iw59'] == 'CS'), 'Tipo de nota'] = 'FT-'

            # Carrega o dataframe feriados e gera o dataframe de vespera
            wks = self.gsheets.worksheet_select(
                'bd_config_robo_religacao', 'Feriados')

            df_feriados = pd.DataFrame(wks.get_as_df(), dtype=str)
            # Ajusta a data para formato de data
            df_feriados['data_feriado'] = pd.to_datetime(
                df_feriados['data_feriado'], format='%d/%m/%Y', dayfirst=True)
            self.feriadoes = df_feriados.copy()
            self.feriadoes['data_feriado'] = self.feriadoes['data_feriado'].dt.date
            # Adiciona se é feriado ou não
            df_feriados['feriado'] = 'S'
            df_vespera = df_feriados.copy()
            df_vespera['data_vespera'] = df_vespera['data_feriado'] - \
                timedelta(days=1)
            df_vespera['vespera'] = 'S'
            df_vespera.drop(columns=['data_feriado', 'feriado'], inplace=True)

            # Remove as linhas com data em branco
            sap_concat = sap_concat.loc[sap_concat['Data da nota'].notna()]
            # Ajusta a data para formato de data
            sap_concat['Data da nota'] = pd.to_datetime(
                sap_concat['Data da nota'], format='%d/%m/%Y', dayfirst=True)
            # Gera uma coluna com a data e hora da nota juntas
            sap_concat['Gerado'] = pd.to_datetime(sap_concat['Data da nota'].astype(
                str) + ' ' + sap_concat['Hora da nota'].astype(str), errors='coerce')
            # Mescla com o dataframe de feriado
            sap_concat = pd.merge(sap_concat, df_feriados,
                                  left_on='Data da nota', right_on='data_feriado', how='left')
            # Mescla com o dataframe de vespera
            sap_concat = pd.merge(sap_concat, df_vespera,
                                  left_on='Data da nota', right_on='data_vespera', how='left')

            # ordena df_feriados pelo 'data_feriado' e cria uma nova coluna 'next_feriado' para salvar o próximo feriado a seguir da data da nota
            df_feriados = df_feriados.sort_values(by='data_feriado')
            sap_concat['next_feriado'] = pd.NaT
            for index, row in sap_concat.iterrows():
                next_feriado = df_feriados[df_feriados['data_feriado']
                                           > row['Data da nota']]['data_feriado'].min()
                sap_concat.at[index, 'next_feriado'] = next_feriado
            sap_concat['next_feriado'] = pd.to_datetime(
                sap_concat['next_feriado'], errors='coerce')

            # Antes de aplicar a função para as datas remove as notas duplicadas existentes
            sap_concat.drop_duplicates(['Nota'], inplace=True, keep='first')

            # Aplicar a função para calcular os novos prazos
            sap_concat['Prazo'] = sap_concat.apply(
                self.calcular_prazo_religacoes, axis=1, result_type='expand')
            
            sap_concat['Prazo'] = pd.to_datetime(
                sap_concat['Prazo'], errors="coerce")

            sap_concat['Prazo'] = sap_concat[['prazo_SAP', 'Prazo']].min(axis=1)

            if 'Instal' in sap_concat.columns:
                sap_concat['Instal'] = sap_concat['Instal'].apply(
                    lambda x: str(int(float(x))) if pd.notna(x) and str(x).strip() != '' and str(x).replace('.', '', 1).isdigit() else x
                )

            # Separa os dataframes para não mexer na data dos que não tem prazo
            sap_sem_prazo = sap_concat[sap_concat['Prazo'].isna(
            ) | ~sap_concat['Tipo de nota iw59'].str.contains("CR")]
            sap_com_prazo = sap_concat[sap_concat['Prazo'].notna(
            ) & sap_concat['Tipo de nota iw59'].str.contains("CR")]
            # Dividir a coluna 'Prazo' em duas colunas: 'Data' e 'Hora' somente para as religações
            if not sap_com_prazo.empty:
                sap_com_prazo[['concl_desejada', 'H final desej.']] = sap_com_prazo['Prazo'].astype(
                    str).str.split(' ', expand=True)

            # Unir novamente os dataframes
            sap_concat = pd.concat(
                [sap_sem_prazo, sap_com_prazo], ignore_index=True, sort=False)

            # Converta a coluna 'concl_desejada' para o formato datetime
            sap_concat['concl_desejada'] = pd.to_datetime(
                sap_concat['concl_desejada'])
            
            # Se nao tiver informacao de urbana ou rural, coloca I
            sap_concat.loc[(sap_concat['Urbano/Rural'] == "") | (sap_concat['Urbano/Rural'].isna()), "Urbano/Rural"] = "I"

            # Organizando os dados
            sap_concat.rename(columns={'Nr. Série Equip.': 'Nr. Série dos Equipamentos',
                                       'Cta.contrato': 'Conta de contrato',
                                       'Denom.executor': 'Denominação',
                                       'Poste da Instalação': 'Nº do Poste da Instalação',
                                       'Ctg.tar.': 'Tipo de tarifa',
                                       'Bairro_x': 'Bairro',
                                       'Cliente_x': 'Cliente',
                                       'concl_desejada': 'Concl.desejada',
                                       'Instal': 'Instalação',
                                       'Urbano/Rural' : 'urbano-rural'},
                              inplace=True)

            df_final = sap_concat[['Nota',
                                   'Nome do parceiro',
                                   'Telefone do parceiro',
                                   'Nr. Série dos Equipamentos',
                                   'Rua',
                                   'Bairro',
                                   'Local',
                                   'Conta de contrato',
                                   'Denominação',
                                   'Cliente', 'Código postal',
                                   'Data da nota', 'Hora da nota',
                                   'Nº do Poste da Instalação',
                                   'Concl.desejada',
                                   'H final desej.',
                                   'Instalação',
                                   'Descrição',
                                   'Tipo de nota',
                                   'Unid.leitura',
                                   'Ponto de referência',
                                   'Nº do pedido',
                                   'Txt.code codif.',
                                   'Tensão de medição',
                                   'Tipo de tarifa',
                                   'Descrição CNAE',
                                   'Latitude',
                                   'Longitude',
                                   'UTD',
                                   'tratativas',
                                   'analise',
                                   'urbano-rural']].copy()

            df_tratativas = df_final.loc[df_final['tratativas'] == 1].copy()
            importar_df = df_final.loc[df_final['tratativas'] == 0].copy()
            importar_df.drop(columns=['tratativas', 'analise'], inplace=True)

            time = datetime.now().strftime('%Y-%m-%d_%H-%M')

            if df_tratativas.empty:
                print("Não há notas nas tratativas")
            else:
                df_tratativas.drop(columns=['urbano-rural'], inplace=True)
                df_tratativas.to_csv(self._dir.dir_temp + 'tratativas_' + time + '.csv',
                                     sep=';',
                                     decimal=',',
                                     encoding='utf-8-sig',
                                     index=False)
                self.gdrive._upload_file(self._dir.dir_temp + 'tratativas_' + time + '.csv', 'tratativas_' + time + '.csv',
                                         TRATATIVAS_FOLDER_ID, True)

            # Gera os arquivos de importação
            print('#Gerando arquivo de importação')
            importar_df.drop_duplicates(['Nota'], inplace=True, keep='first')
            utds = ['BAR',
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
                    'ITG']

            importar_df['Descrição'] = importar_df['Descrição'].str.replace(
                '[;]', '', regex=True)
            importar_df['Data da nota'] = pd.to_datetime(
                importar_df['Data da nota'], format='%d/%m/%Y', dayfirst=True).dt.strftime("%d/%m/%Y")
            importar_df['Concl.desejada'] = pd.to_datetime(
                importar_df['Concl.desejada'], dayfirst=True).dt.strftime('%d/%m/%Y')


            importar_df.to_csv(self._dir.dir_temp + 'final_' + time + '.csv',
                               sep=';',
                               decimal=',',
                               encoding='utf-8-sig',
                               index=False)
            if variante == "PCP_CORTE":
                self.gdrive._upload_file(self._dir.dir_temp + 'final_' + time + '.csv', 'final_' + time + '.csv',
                                         FINAL_FOLDER_ID, True)
                
            postar_aviso = importar_df.copy()
            postar_aviso['Prazo'] = postar_aviso["Concl.desejada"] + " " + postar_aviso["H final desej."]
            postar_aviso['Prazo'] = pd.to_datetime(postar_aviso['Prazo'], format='%d/%m/%Y %H:%M:%S', dayfirst=True)
            
            depoisdedoze = datetime.now() + pd.Timedelta(hours=12)
            postar_aviso = postar_aviso.loc[(postar_aviso['Prazo'] <= depoisdedoze)]
            if not postar_aviso.empty :
              for index,row in postar_aviso.iterrows() :
                self.telegram.bot_message("🚨🚨 Curto prazo em "+ str(row['UTD']) + ": " + str(row['Nota']))
                #print("🚨🚨 Curto prazo em "+ str(row['UTD']) + ": " + str(row['Nota']))

            
            for utd in utds:
                import_file = importar_df.query('UTD == "{0}"'.format(utd)).drop('UTD',
                                                                                 axis=1)
                if import_file.empty:
                    pass
                else:
                    import_file.to_csv(self._dir.dir_temp + utd+'.csv',
                                       sep=';',
                                       encoding='utf-8-sig',
                                       index=False)
        except Exception as e:
            print('Erro ao gerar o arquivo de importação')
            print(e)
            traceback.print_exc()
            raise e
    
    #fsz um loop com is_dia_util acrescentando 1 dia ate chegar num dia util
    def proximo_dia_util(self, datahora) :
        retorno = datahora
        # se for entre 18h e 8h, coloca pra próxima 8h
        if (retorno.hour >= 18) :
            retorno = (retorno + pd.Timedelta(days=1)).replace(hour=8, minute=0, second=0)
        elif (retorno.hour < 8) :
            retorno = retorno.replace(hour=8, minute=0, second=0)
        # vai adicionando um dia ate chegar num dia util
        while not self.is_dia_util(retorno) :
            retorno = retorno + pd.Timedelta(days=1)
        return retorno
    
    #se .weekday() for 5 ou 6, retorna False
    #se feriadoes conter a data, retorna False
    #de resto, retorna True
    def is_dia_util(self, datahora) :
        if datahora.weekday() >= 5 :
            return False
        if (self.feriadoes['data_feriado'] == datahora.date()).any():
            return False
        return True

    def calcular_prazo_religacoes(self, row):
        # Só acessa notas que forem de religação
        if 'Religação' in row['Txt.code codif.']:
            data_inicio = self.proximo_dia_util(row['Gerado'])
            if (row['Urbano/Rural'] == 'R') :
                return data_inicio + pd.Timedelta(days=2)
            else :
                return data_inicio + pd.Timedelta(days=1)
#            # Religas geradas em dias que não são feriado
#            if (row['feriado'] != 'S' and row['vespera'] != 'S') or (row['vespera'] == 'S' and row['Gerado'].hour < 18):
#                # Religas geradas Entre segunda a sexta
#                if 0 <= row['Gerado'].weekday() <= 4:
#                    # Religas geradas nas sextas a partir das 18:00:00 e com feriado na próxima segunda
#                    if row['Gerado'].weekday() == 4 and row['Gerado'].hour >= 18 and (row['next_feriado'] - row['Data da nota']) == pd.Timedelta(days=3):
#                        if row['Urbano/Rural'] == 'U':
#                            return (row['Gerado'] + pd.Timedelta(days=5)).replace(hour=8, minute=0, second=0)
#                        if row['Urbano/Rural'] == 'R':
#                            return (row['Gerado'] + pd.Timedelta(days=6)).replace(hour=8, minute=0, second=0)
#                    # Religas geradas nas sextas a partir das 18:00:00
#                    elif row['Gerado'].weekday() == 4 and row['Gerado'].hour >= 18:
#                        if row['Urbano/Rural'] == 'U':
#                            return (row['Gerado'] + pd.Timedelta(days=4)).replace(hour=8, minute=0, second=0)
#                        if row['Urbano/Rural'] == 'R':
#                            return (row['Gerado'] + pd.Timedelta(days=5)).replace(hour=8, minute=0, second=0)
#                    # Religa gerada entre 08:00:00 e 17:59:59
#                    elif 8 <= row['Gerado'].hour < 18:
#                        if row['Urbano/Rural'] == 'U':
#                            return row['Gerado'] + pd.Timedelta(days=1)
#                        if row['Urbano/Rural'] == 'R':
#                            return row['Gerado'] + pd.Timedelta(days=2)
#                    # Religa gerada entre 18:00:00 e 23:59:59
#                    elif 18 <= row['Gerado'].hour <= 23:
#                        if row['Urbano/Rural'] == 'U':
#                            return (row['Gerado'] + pd.Timedelta(days=2)).replace(hour=8, minute=0, second=0)
#                        if row['Urbano/Rural'] == 'R':
#                            return (row['Gerado'] + pd.Timedelta(days=3)).replace(hour=8, minute=0, second=0)
#                    # Religa gerada entre 00:00:00 e 08:00:00
#                    elif 0 <= row['Gerado'].hour < 8:
#                        if row['Urbano/Rural'] == 'U':
#                            return (row['Gerado'] + pd.Timedelta(days=1)).replace(hour=8, minute=0, second=0)
#                        if row['Urbano/Rural'] == 'R':
#                            return (row['Gerado'] + pd.Timedelta(days=2)).replace(hour=8, minute=0, second=0)
#                # Religas geradas nos sábados
#                elif row['Gerado'].weekday() == 5:
#                    # se houver feriado na segunda, enviar para quarta
#                    if (row['next_feriado'] - row['Data da nota']) == pd.Timedelta(days=2):
#                        if row['Urbano/Rural'] == 'U':
#                            return (row['Gerado'] + pd.Timedelta(days=4)).replace(hour=8, minute=0, second=0)
#                        if row['Urbano/Rural'] == 'R':
#                            return (row['Gerado'] + pd.Timedelta(days=5)).replace(hour=8, minute=0, second=0)
#                    else:
#                        if row['Urbano/Rural'] == 'U':
#                            return (row['Gerado'] + pd.Timedelta(days=3)).replace(hour=8, minute=0, second=0)
#                        if row['Urbano/Rural'] == 'R':
#                            return (row['Gerado'] + pd.Timedelta(days=4)).replace(hour=8, minute=0, second=0)
#                # Religas geradas nos domingos
#                elif row['Gerado'].weekday() == 6:
#                    # se houver feriado na segunda, enviar para quarta
#                    if (row['next_feriado'] - row['Data da nota']) == pd.Timedelta(days=1):
#                        if row['Urbano/Rural'] == 'U':
#                            return (row['Gerado'] + pd.Timedelta(days=3)).replace(hour=8, minute=0, second=0)
#                        if row['Urbano/Rural'] == 'R':
#                            return (row['Gerado'] + pd.Timedelta(days=4)).replace(hour=8, minute=0, second=0)
#                    else:
#                        if row['Urbano/Rural'] == 'U':
#                            return (row['Gerado'] + pd.Timedelta(days=2)).replace(hour=8, minute=0, second=0)
#                        if row['Urbano/Rural'] == 'R':
#                            return (row['Gerado'] + pd.Timedelta(days=3)).replace(hour=8, minute=0, second=0)
#            # Religas geradas no feriado ou vespera
#            elif row['feriado'] == 'S' or row['vespera'] == 'S':
#                if row['vespera'] == 'S' and 18 <= row['Gerado'].hour <= 23:
#                    if row['Urbano/Rural'] == 'U':
#                        return (row['Gerado'] + pd.Timedelta(days=3)).replace(hour=8, minute=0, second=0)
#                    if row['Urbano/Rural'] == 'R':
#                        return (row['Gerado'] + pd.Timedelta(days=4)).replace(hour=8, minute=0, second=0)
#                # Se o feriado for na sexta, enviar para terça
#                if row['Gerado'].weekday() == 4:
#                    if row['Urbano/Rural'] == 'U':
#                        return (row['Gerado'] + pd.Timedelta(days=4)).replace(hour=8, minute=0, second=0)
#                    if row['Urbano/Rural'] == 'R':
#                        return (row['Gerado'] + pd.Timedelta(days=5)).replace(hour=8, minute=0, second=0)
#                # Se o prazo não cair nos dias anteriores, adicionar
#                else:
#                    if row['Urbano/Rural'] == 'U':
#                        return (row['Gerado'] + pd.Timedelta(days=2)).replace(hour=8, minute=0, second=0)
#                    if row['Urbano/Rural'] == 'R':
#                        return (row['Gerado'] + pd.Timedelta(days=3)).replace(hour=8, minute=0, second=0)

    def extrair_debitos(self, row):
        try:
            if row['Texto']:
                # Usar regex para encontrar números que começam com uma das strings desejadas e terminam após a vírgula
                match_1 = re.search(
                    r'@(S|D|P|C)\$(\d+,\d+)', str(row['Texto']))
                match_2 = re.search(r'R\$\s*([\d,.]+)', str(row['Texto']))
                match_3 = re.search(r'@\$\s*([\d,.]+)', str(row['Texto']))
                match_4 = re.search(r'\$([\d,.]+)', str(row['Texto']))
                if match_1:
                    match = match_1
                    # Retornar o número encontrado ou NaN se não houver correspondência
                    return str(match.group(2))
                if match_2:
                    match = match_2
                    numero_str = match.group(1)
                    # Remover vírgulas e converter para float
                    numero = str(numero_str.replace('.', ''))
                    return numero
                if match_3:
                    match = match_3
                    numero_str = match.group(1)
                    return numero_str
                if match_4:
                    match = match_4
                    numero_str = match.group(1)
                    return numero_str
                else:
                    return None
            else:
                return None
        except:
            return None

    def adicionar_codes_cs(self, row):
        try:
            codes = ['']
            if 'CS' in row['Tipo da Nota']:
                codes = ['CORTEFX1', 'CORTEFX2', 'CORTEFX3']
            if 'CA' in row['Tipo da Nota']:
                codes = ['ACOMPFX1', 'ACOMPFX2', 'ACOMPFX3']
            if (row['Valores']):
                try:
                    valor = float(str(row['Valores']).replace(',', '.'))
                    if valor < 500:
                        # faixa 1
                        return codes[0]
                    if valor >= 500 and valor < 2000:
                        # faixa 2
                        return codes[1]
                    if valor >= 2000:
                        # faixa 3
                        return codes[2]
                except:
                    return None
            else:
                return codes[0]
        except:
            return None
