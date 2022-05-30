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
    qty = []
    corte = []
    riv = []
    sl = []
    for i in part_num:    
        qty.append(
            df[df['NP'] == i ]['Nº de circuito'].count()
            )
    for i in part_num:    
        corte.append(
            df[(df['NP'] == i ) & (df['Machine'].str.startswith('A'))]['Nº de circuito'].count()
            )
    for i in part_num:    
        riv.append(
            df[(df['NP'] == i ) & (df['Machine'].str.startswith('B'))]['Nº de circuito'].count()
            )
    for i in part_num:    
        sl.append(
            df[(df['NP'] == i )&(df['Machine'].str.startswith('SLD'))]['Nº de circuito'].count()
            )
    return part_num, qty, corte, riv, sl

def clean_postp(df):
    df.rename(columns={'P/N' : 'NP'}, inplace=True)
    df['Unique'] = df['NP'] + " " +df['Nº de circuito'] + " " + df['Ruta']
    df.drop_duplicates(subset ="Unique",keep = 'first', inplace = True)
    return df

def postp_info(df):
    part_num = df.NP.unique()
    twist = []
    sello = []
    desforre_medio = []
    desforre_punta = []
    encinte_auto = []
    inser_tuboter = []
    inser_tubo = []
    termo = []
    for i in part_num:
            twist.append(
                df[(df['NP'] == i ) & (df['Maquina'].str.startswith('TW'))]['Nº de circuito'].count()
                )
    for i in part_num:    
            sello.append(
                df[(df['NP'] == i ) & (df['Ruta'] == 'Insert Sub-materials [Seal]')]['Nº de circuito'].count()
                )
    for i in part_num:    
            desforre_medio.append(
                df[(df['NP'] == i ) & (df['Ruta'] == 'Manual stripping(Middle)')]['Nº de circuito'].count()
                )
    for i in part_num:    
            desforre_punta.append(
                df[(df['NP'] == i ) & (df['Ruta'] == 'Shield Wire Inner Sheath stripping')]['Nº de circuito'].count() * 2
                )
    for i in part_num:    
            encinte_auto.append(
                df[(df['NP'] == i ) & (df['Ruta'] == 'Manual Taping')]['Nº de circuito'].count()
                )
    for i in part_num:    
            inser_tubo.append(
                df[(df['NP'] == i ) & (df['Ruta'] == 'lnsert Sub-materials [PVC,COT,SLEEVE]')]['Nº de circuito'].count()
                )
    for i in part_num:    
            inser_tuboter.append(
                df[(df['NP'] == i ) & (df['Ruta'] == 'Insert Sub-materials [HMT,HSC]')]['Nº de circuito'].count()
                )
    for i in part_num:    
            termo.append(
                df[(df['NP'] == i ) & (df['Ruta'] == 'Heat-melting')]['Nº de circuito'].count()
                )
    return  part_num, twist, sello, desforre_medio, desforre_punta, encinte_auto, inser_tuboter, inser_tubo, termo



    

df = concat_allexcel('CHECKLIST/')

df = clean_checklist(df)

df.to_excel('checklist.xlsx', index = False)


# part_num, qty, corte, riv, sl = checklist_info(df)

# print(part_num)

#df2.to_excel('edited_3.xlsx', index = False)
