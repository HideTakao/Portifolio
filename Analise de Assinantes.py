import pandas as pd
import time as tm
import re

corte = pd.read_excel('AP1 PENDENTES.xlsx')
corte = corte.fillna('')
corte['ENDEREÇO'] = corte['END_COMPLETO'] + ' ' + corte['ID_COMPL1'] + ' ' + corte['COMPL1_DESCR'].astype(str)
lista_corte = corte['ENDEREÇO'].unique().tolist()
assinantes = pd.read_excel('Relatorio assinantes.xlsx')
assinantes = assinantes.fillna('')
assinantes['NUM_ENDERECO'] = assinantes['NUM_ENDERECO'].astype(str).apply(lambda x: x.split('.')[0])
assinantes['ENDEREÇO'] = assinantes['NOM_LOGR_COMPLETO'] + ' ' + assinantes['NUM_ENDERECO'] + ' ' + assinantes['COD_TIPO_COMPL1'] + ' ' + assinantes['TXT_TIPO_COMPL1'].astype(str) + ' ' + 
assinantes['COD_TIPO_COMPL2'] + ' ' + assinantes['TXT_TIPO_COMPL2'].astype(str) + ' ' + assinantes['COD_TIPO_COMPL'] + ' ' + assinantes['TXT_COMPL'].astype(str)
lista_assinantes = assinantes['ENDEREÇO'].unique().tolist()
instalados = {}
desconectados = []
for i in lista_corte:
    a = re.sub(r'\s+', ' ', i.strip())
    a = re.sub(r'\b(\d)\b', lambda x: x.group(1).zfill(2), a)
    encontrado = False

    for b in lista_assinantes:
        c = re.sub(r'\s+', ' ', b.strip())
        c = re.sub(r'\b(\d)\b', lambda x: x.group(1).zfill(2), c)
        if a == c:
            instalados[i] = b
            encontrado = True
            break

    if not encontrado:
        desconectados.append(i)  # ou outra ação desejada
instalados = pd.DataFrame(list(instalados.items()), columns=['ENDEREÇO CORTE', 'ENDEREÇO'])
lista_ativos = pd.merge(instalados,corte[['COD_IMOVEL','COD_NODE','ENDEREÇO','NUM_CONTRATO']], left_on='ENDEREÇO CORTE', right_on= 'ENDEREÇO', how='left')
lista_ativos = lista_ativos.rename(columns={'NUM_CONTRATO' : 'CONTRATO ANTERIOR','ENDEREÇO_x' : 'ENDEREÇO'})        
lista_ativos.drop_duplicates(subset=['ENDEREÇO'], keep='first', inplace=True)
lista_ativos = lista_ativos.reset_index(drop=True)
assinantes_ativos = pd.merge(lista_ativos,assinantes[['ENDEREÇO','NUM_CONTRATO']], on='ENDEREÇO', how='left')
assinantes_ativos.drop_duplicates(subset=['ENDEREÇO'], keep='first', inplace=True)
assinantes_ativos = assinantes_ativos.reset_index(drop=True)
colunas = ['COD_IMOVEL','COD_NODE', 'ENDEREÇO', 'CONTRATO ANTERIOR', 'NUM_CONTRATO']
assinantes_ativos = assinantes_ativos[colunas]
assinantes_ativos = assinantes_ativos.rename(columns={'NUM_CONTRATO':'CONTRATO ATUAL'})
desconectados = pd.DataFrame({'ENDEREÇO': desconectados})
desconectados = pd.merge(desconectados,corte[['COD_IMOVEL','COD_NODE','ENDEREÇO','NUM_CONTRATO']], on= 'ENDEREÇO', how='left')
desconectados['MDU/SDU'] = 'SDU'
MDUS = pd.read_excel(r'C:\Users\N5802623\OneDrive - Claro SA\Relatorios\GED.xlsx')
desconectados.loc[desconectados['COD_IMOVEL'].isin(MDUS['Código GED']), 'MDU/SDU'] = 'MDU'
desconectados.drop_duplicates(subset=['ENDEREÇO'], keep='first', inplace=True)
assinantes_ativos.to_excel('assinantes ativos.xlsx', index=False)
MDU_desc = desconectados[desconectados['MDU/SDU'] == 'MDU']
SDU_des = desconectados[desconectados['MDU/SDU'] == 'SDU']
MDU_desc.to_excel('lista para desconexão MDU.xlsx', index=False)
SDU_des.to_excel('lista de desconexão SDU.xlsx', index=False)
