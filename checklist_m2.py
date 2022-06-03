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
        }, inplace=True)
    for col in df.columns:
        if 'Unnamed' in col:
            del df[col]
    df['unique'] = df['Nº de circuito'].astype(str) + ' ' + df['NP']
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
    desmalle = [df[(df['NP'] == i )&(df['Nº de circuito'].str.endswith('#'))]['Nº de circuito'].count() * 2
            for i in part_num]
    termo_joint = [df[(df['NP'] == i ) & (df['Machine'].str.contains('H', case=False)) & (df['Machine'].str.contains('SJ', case = False))]['Nº de circuito'].count()
            for i in part_num]
    twist_sld = desmalle
    df = pd.DataFrame(list(zip(part_num, qty, corte, riv, sl, joint_sld, desmalle, twist_sld, termo_joint)),
                      columns=['PN', 'QTY', 'Corte', 'RIVIAN', 'SLD', 'JOINT SLD', 'DESMALLE', 'TWIST MALLA', 'TERMO JOINT'])
    return df


checklist = checklist_info(clean_checklist(concat_allexcel('CHECKLIST/')))


checklist.to_excel('checklist.xlsx')
