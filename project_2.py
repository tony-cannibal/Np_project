import pandas as pd


df = pd.read_excel('/home/luis/Documents/Python/edited.xlsx')


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

data = {'np': part_num, 'qty':qty, 'corte': corte, 'riv': riv, 'sld': sl}

df2 = pd.DataFrame(data, columns=['np', 'qty', 'corte', 'riv', 'sld'])

df2.to_excel('/home/luis/Documents/Python/edited2.xlsx', index = False)

print(part_num)
print(qty)
print(corte)
print(sl)
print(riv)
#print(qty)
