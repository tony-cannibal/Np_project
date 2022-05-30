import pandas as pd


df = pd.read_excel('info.xlsx')

df.rename(columns={'P/N' : 'NP'}, inplace=True)

df['Unique'] = df['NP'] + " " + df['No.circuito'] + " " + df['Terminal de union']

df.drop_duplicates(subset ="Unique",keep = 'first', inplace = True)

part_num = df.NP.unique()
prensas =[]
joint = []
prensas_sld = []

for i in part_num:
        prensas.append(
            df[(df['NP'] == i ) & (df['Maquina'].str.startswith('C'))]['No.circuito'].count()
            )
for i in part_num:
        joint.append(
            df[(df['NP'] == i ) & (df['Maquina'].str.startswith('SJ'))]['No.circuito'].count()
            )
for i in part_num:
        prensas_sld.append(
            df[(df['NP'] == i ) & (df['Maquina'] == 'SL04')]['No.circuito'].count() +
            df[(df['NP'] == i ) & (df['Maquina'] == 'SL05')]['No.circuito'].count() +
            df[(df['NP'] == i ) & (df['Maquina'] == 'SL06')]['No.circuito'].count() +
            df[(df['NP'] == i ) & (df['Maquina'] == 'SL07')]['No.circuito'].count()
            )

#df.to_excel('edited.xlsx', index = False)
print(part_num)
print(prensas)
print(joint)
print(prensas_sld)
