import pandas as pd
import glob


def concat_allexcel(files):
    all_files = glob.glob(files + "*.xlsx")
    df = pd.concat((pd.read_excel(f) for f in all_files), ignore_index=True)
    return  df

def clean_checklist(df):
    df.rename(columns={ 'Unnamed: 21' : 'NP', 
                                    'Unnamed: 22' : 'Lot', 
                                    'Unnamed: 11' : 'Diagrama', 
                                    'Unnamed: 23' : 'Cantidad',
                                    'Unnamed: 28' : 'Modelo',
                                    'Unnamed: 29' : 'Area'}, inplace=True)
    for col in df.columns:
        if 'Unnamed' in col:
            del df[col]
    df['unique'] = df['Nº de circuito'].astype(str) + ' ' + df['NP']
    df.drop_duplicates(subset ="unique",keep = 'first', inplace = True)
    df[['Machine 1', 'Machine 2', 'Machine 3','Machine 4',
        'Machine 5', 'Machine 6', 'Machine 7', 'Machine 8']]=df['Machine'].str.split(' ', expand=True)
    return df

def checklist_info(df):
    part_num = df.NP.unique()
    qty = [ df[df['NP'] == i ]['Nº de circuito'].count()
            for i in part_num ]
    corte = [ df[(df['NP'] == i ) & (df['Machine'].str.startswith('A'))]['Nº de circuito'].count()
              for i in part_num]
    riv = [ df[(df['NP'] == i ) & (df['Machine'].str.startswith('B'))]['Nº de circuito'].count()
            for i in part_num]
    sl = [ df[(df['NP'] == i )&(df['Machine'].str.startswith('SLD'))]['Nº de circuito'].count()
           for i in part_num]
    termo_joint =[df[(df['NP'] == i )&(df['Machine'].str.contains('SJ', case = False)) &
                     (df['Machine'].str.contains('H', case  = False))]['Nº de circuito'].count()
                for i in part_num]
    joint_sld = [ df[(df['NP'] == i )&(df['Nº de circuito'].str.endswith('#')) & (df['Terminal(L)'].str.startswith('TKT'))]['Nº de circuito'].count() +
                  df[(df['NP'] == i )&(df['Nº de circuito'].str.endswith('#')) & (df['Terminal(R)'].str.startswith('TKT'))]['Nº de circuito'].count()               
                  for i in part_num]
    desmalle = [df[(df['NP'] == i )&(df['Nº de circuito'].str.endswith('#'))]['Nº de circuito'].count() * 2
                for i in part_num]
    twist_malla = desmalle
    encinte_manual = desmalle
    termo_sld = desmalle
    df = pd.DataFrame(list(zip(part_num, qty, corte, riv, sl, joint_sld, desmalle, twist_malla, encinte_manual, termo_joint, termo_sld)),
                      columns=['PN', 'QTY', 'CORTE', 'RIVIAN', 'SLD', 'JOINT SLD', 'DESMALLE',
                               'TWIST SLD', 'ENCINTE MANUAL', 'TERMO  JOINT', 'TERMO SLD'])
    return df

def clean_manual(df):
    df.rename(columns={'P/N' : 'NP'}, inplace=True)
    df['Unique'] = df['NP'] + " " + df['No.circuito'].astype(str) + " " + df['Terminal de union']
    df.drop_duplicates(subset ="Unique",keep = 'first', inplace = True)
    return df

def manual_info(df):
    part_num = df.NP.unique()
    prensa_total =[df[(df['NP'] == i ) & (df['Maquina'].str.startswith('C'))]['No.circuito'].count()
              for i in part_num]
    joint = [df[(df['NP'] == i ) & (df['Maquina'].str.startswith('SJ'))]['No.circuito'].count()
             for i in part_num]
    prensas_sld = [df[(df['NP'] == i ) & (df['Maquina'] == 'SL52')]['No.circuito'].count() +
                df[(df['NP'] == i ) & (df['Maquina'] == 'SL53')]['No.circuito'].count() +
                df[(df['NP'] == i ) & (df['Maquina'] == 'SL54')]['No.circuito'].count() 
                   for i in part_num]
    prensa_batt = [df[(df['NP'] == i ) & (df['Maquina'].str.startswith('C15'))]['No.circuito'].count()
                   for i in part_num]
    prensas = [ a - b for a, b in zip(prensa_total, prensa_batt)]
    df = pd.DataFrame(list(zip(part_num, prensas, joint, prensas_sld, prensa_batt)),
                   columns = ['PN', 'PRENSAS', 'JONT', 'PRENSAS SLD', 'PRENSA BATT'])
    return df

def clean_postp(df):
    df.rename(columns={'P/N' : 'NP'}, inplace=True)
    df['Unique'] = df['NP'] + " " +df['Nº de circuito'].astype(str) + " " + df['Ruta']
    df.drop_duplicates(subset ="Unique",keep = 'first', inplace = True)
    return df

def postp_info(df):
    part_num = df.NP.unique()
    twist = [df[(df['NP'] == i ) & (df['Maquina'].str.startswith('TW'))]['Nº de circuito'].count()
                      for i in part_num]
    sello = [df[(df['NP'] == i ) & (df['Ruta'] == 'Insert Sub-materials [Seal]')]['Nº de circuito'].count()
                      for i in part_num]
    desforre_medio = [df[(df['NP'] == i ) & (df['Ruta'] == 'Manual stripping(Middle)')]['Nº de circuito'].count()
                      for i in part_num]
    desforre_punta = [df[(df['NP'] == i ) & (df['Ruta'] == 'Shield Wire Inner Sheath stripping')]['Nº de circuito'].count() * 2
                      for i in part_num]
    encinte_auto = [df[(df['NP'] == i ) & (df['Ruta'] == 'Manual Taping')]['Nº de circuito'].count()
                      for i in part_num]
    inser_tubo = [df[(df['NP'] == i ) & (df['Ruta'] == 'lnsert Sub-materials [PVC,COT,SLEEVE]')]['Nº de circuito'].count() +
                  df[(df['NP'] == i ) & (df['Ruta'] == 'Insert Sub-materials [HMT,HSC]')]['Nº de circuito'].count()
                       for i in part_num]    
    termo_batt = [df[(df['NP'] == i ) & (df['Maquina'].str.startswith('H')) & (df['SQ'] >= 15 )]['Nº de circuito'].count()
                       for i in part_num]   
    df = pd.DataFrame(list(zip(part_num, twist, sello, desforre_medio, desforre_punta, encinte_auto, inser_tubo, termo_total, termo_batt)),
                  columns = ['PN', 'TWIST', 'INSERCION DE SELLO', 'DESFORRE MEDIO', 'DESFORRE DE PUNTA',
                             'ENCINTE AUTOMATICO',  'INSERCION DE TUBO', 'TERMO', 'TERMO BATT'])
    return df

def combinar(df1, df2, df3):
    procesos = pd.merge(df1, df2, on ='PN', how ='left')
    procesos = pd.merge(procesos, df3, on = 'PN', how = 'left')
    procesos['MATERIAL'] = procesos['PN'].str[:-3]
    procesos['REV'] = procesos['PN'].str[-2:]  
    np = pd.read_excel('numeros_de_parte.xlsx')
    procesos = pd.merge(procesos, np, on='MATERIAL', how = 'left')
    procesos = procesos.fillna(0)
    #procesos = procesos.reindex(['AREA', 'MODELO', 'ITEM', 'PN', 'MATERIAL', 'REV', 'QTY', 'CORTE', 'RIVIAN', 'SLD', 'PRENSAS', 'JONT',
     #                            'PRENSAS SLD', 'JOINT SLD', 'DESFORRE MEDIO', 'DESFORRE DE PUNTA',  'TWIST', 'ENCINTE AUTOMATICO',
     #                            'TERMO', 'INSERCION DE SELLO', 'INSERCION DE TUBO', 'TWIST SLD', 'ENCINTE MANUAL', 'DESMALLE'], axis=1)
    return procesos

print('[+] Procesando Checklist...')
print('')
checklist = checklist_info(clean_checklist(concat_allexcel('CHECKLIST/')))

print('[+] Procesando Manuales...')
print('')
manual = manual_info(clean_manual(concat_allexcel('MANUAL/')))

print('[+] Procesando Procesos Posteriores...')
print('')
post_procesos = postp_info(clean_postp(concat_allexcel('POST-PROCESOS/')))

print('[+] Combinando Procesos...')
print('')
procesos = combinar(checklist, manual, post_procesos)

print('[+] Guardando Archivo...')
print('')

procesos.to_excel('Procesos.xlsx', index = False)
print('[+] Terminbado...')

