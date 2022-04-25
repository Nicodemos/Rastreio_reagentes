import pandas as pd
import numpy as np
import datetime as dt
import os
import xlrd
import openpyxl
colunas = ['MW1','MW2','AGUA','MLB','LOTE_MIX']
df_final = pd.DataFrame(columns=colunas + ['DATA'])
cont_erro = 0
cont_ok = 0

anos = [2021,2022]
for ano in anos:

    path_med, pastas, file_med = next(os.walk(r"\\ceara-fs\Ceara\Central Analítica\Biologia Molecular\Room3-record\S3. EXAMES\{}".format(ano)))
    for mes in pastas:

        path_med1, pastas_dias, file_med1 = next(os.walk(r"\\ceara-fs\Ceara\Central Analítica\Biologia Molecular\Room3-record\S3. EXAMES\{}\{}".format(ano,mes)))
        for dias in pastas_dias:

            path_med, dirs_med, file_names = next(os.walk(r"\\ceara-fs\Ceara\Central Analítica\Biologia Molecular\Room3-record\S3. EXAMES\{}\{}\{}".format(ano, mes, dias)))
            for file in file_names:

                if file != 'Thumbs.db':
                    try:

                        df = pd.read_excel(r""+path_med+"\\"+file)
                        df = df.iloc[19:20,2:11]
                        colunas_temp = ['Unnamed: 2','Unnamed: 4','Unnamed: 6','Unnamed: 8','Unnamed: 10']
                        df = df[colunas_temp]
                        df = pd.DataFrame(df.values,columns=colunas)

                        df.MW1 = str(df.MW1.str.split(':')[0][1]).strip()
                        df.MW2 = str(df.MW2.str.split(':')[0][1]).strip()
                        df.AGUA = str(df.AGUA.str.split(':')[0][1]).strip()
                        df.MLB = str(df.MLB.str.split(':')[0][1]).strip()
                        df.LOTE_MIX = str(df.LOTE_MIX.str.split(':')[0][1]).strip()
                        df['DATA'] = dias.replace('.','/').strip()

                        df_final = pd.concat([df,df_final], ignore_index=True)
                        cont_ok+=1
                    except Exception as inst:
                        cont_erro+=1

df_final.to_excel(r"\\ceara-fs\Ceara\Central Analítica\Biologia Molecular\Room3-record\S3. Supervisão\TESTE_TI\dados_placa_termociclador.xlsx", sheet_name='historico',index=False)
print('Quantidade de Erros: ',cont_erro)
print('Quantidade de planilhas lidas corretamente: ',cont_ok)