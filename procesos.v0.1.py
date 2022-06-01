import pandas as pd
import glob


def concat_allexcel(files):
    all_files = glob.glob(files + "*.xlsx")
    df = pd.concat((pd.read_excel(f) for f in all_files), ignore_index=True)
    return  df

def clean_checklist(df):
    df.rename(columns={ 
        'Unnamed: 21' : 'NP', 
        'Unnamed: 22' : 'Lot', 
        'Unnamed: 11' : 'Diagrama', 
        'Unnamed: 23' : 'Cantidad',
        'Unnamed: 28' : 'Modelo',
        'Unnamed: 29' : 'Area'
        }, 
            inplace=True)
    for col in df.columns:
        if 'Unnamed' in col:
            del df[col]
    df['unique'] = df['Nº de circuito'] + ' ' + df['NP']
    df.drop_duplicates(subset ="unique",keep = 'first', inplace = True)
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
    joint_sld = [ df[(df['NP'] == i )&(df['Nº de circuito'].str.endswith('#')) & (df['Terminal(L)'].str.startswith('TKT'))]['Nº de circuito'].count() +
                  df[(df['NP'] == i )&(df['Nº de circuito'].str.endswith('#')) & (df['Terminal(R)'].str.startswith('TKT'))]['Nº de circuito'].count()               
                  for i in part_num]
    df = pd.DataFrame(list(zip(part_num, qty, corte, riv, sl, joint_sld)),
                      columns=['PN', 'QTY', 'Corte', 'RIVIAN', 'SLD', 'JOINT SLD'])
    return df

def clean_manual(df):
    df.rename(columns={'P/N' : 'NP'}, inplace=True)
    df['Unique'] = df['NP'] + " " + df['No.circuito'] + " " + df['Terminal de union']
    df.drop_duplicates(subset ="Unique",keep = 'first', inplace = True)
    return df

def manual_info(df):
    part_num = df.NP.unique()
    prensas =[df[(df['NP'] == i ) & (df['Maquina'].str.startswith('C'))]['No.circuito'].count()
              for i in part_num]
    joint = [df[(df['NP'] == i ) & (df['Maquina'].str.startswith('SJ'))]['No.circuito'].count()
             for i in part_num]
    prensas_sld = [df[(df['NP'] == i ) & (df['Maquina'] == 'SL04')]['No.circuito'].count() +
                df[(df['NP'] == i ) & (df['Maquina'] == 'SL05')]['No.circuito'].count() +
                df[(df['NP'] == i ) & (df['Maquina'] == 'SL06')]['No.circuito'].count() +
                df[(df['NP'] == i ) & (df['Maquina'] == 'SL07')]['No.circuito'].count()
                   for i in part_num]
    df = pd.DataFrame(list(zip(part_num, prensas, joint, prensas_sld)),
                   columns = ['PN', 'PRENSAS', 'JONT', 'PRENSAS SLD'])
    return df

def clean_postp(df):
    df.rename(columns={'P/N' : 'NP'}, inplace=True)
    df['Unique'] = df['NP'] + " " +df['Nº de circuito'] + " " + df['Ruta']
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
    inser_tuboter = [df[(df['NP'] == i ) & (df['Ruta'] == 'Insert Sub-materials [HMT,HSC]')]['Nº de circuito'].count()
                     for i in part_num]
    inser_tubo = [df[(df['NP'] == i ) & (df['Ruta'] == 'lnsert Sub-materials [PVC,COT,SLEEVE]')]['Nº de circuito'].count()
                  for i in part_num]
    termo = [df[(df['NP'] == i ) & (df['Ruta'] == 'Heat-melting')]['Nº de circuito'].count()
             for i in part_num]
    df = pd.DataFrame(list(zip(part_num, twist, sello, desforre_medio, desforre_punta, encinte_auto, inser_tuboter, inser_tubo, termo)),
                  columns = ['PN', 'TWIST', 'SELLO', 'DESFORRE MEDIO', 'DESFORRE DE PUNTA',
                             'ENCINTE AUTOMATICO', 'INSERCION DE TUBO TERMO', 'INSERCION DE TUBO', 'TERMO'])
    return df

def combinar(df1, df2, df3):
    procesos= pd.merge(df1, df2, on ='PN', how ='left')
    procesos= pd.merge(procesos, df3, on = 'PN', how = 'left')
    procesos = procesos.fillna(0)
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
procesos.to_excel('Procesos.xlsx...', index = False)
print('[+] Terminbado')

