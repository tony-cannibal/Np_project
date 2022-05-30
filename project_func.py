import pandas as pd


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
        qty.append(df[df['NP'] == i ]['Nº de circuito'].count())
    for i in part_num:    
        corte.append(df[(df['NP'] == i )&(df['Machine'].str.startswith('A'))]['Nº de circuito'].count())
    for i in part_num:    
        riv.append(df[(df['NP'] == i )&(df['Machine'].str.startswith('B'))]['Nº de circuito'].count())
    for i in part_num:    
        sl.append(df[(df['NP'] == i )&(df['Machine'].str.startswith('SLD'))]['Nº de circuito'].count())
    return part_num, qty, corte, riv, sl


df = pd.read_excel('c 2022.05.26.xlsx')

df2 = clean_checklist(df)

part_num, qty, corte, riv, sl = checklist_info(df2)

print(part_num)

#df2.to_excel('edited_3.xlsx', index = False)
