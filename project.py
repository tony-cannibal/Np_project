import pandas as pd
import openpyxl as pxl

# Load a file as a pandas data frame be it a csv or excel file,
# in this case its an excel file. You need to specify the file location,
# if tha file is in the same directory as the script then you just need
# to write the name of the file.
df = pd.read_excel('c 2022.05.26.xlsx')

# This is a method to list all the columns in
# the data frame.
# cols = df.columns


# Rename unnamed columns witch are necesary, in this case 
# I had to fisicaly check the index of every column I needed.
df.rename(columns={ 'Unnamed: 21' : 'NP', 
    'Unnamed: 22' : 'Lot', 
    'Unnamed: 11' : 'Diagrama', 
    'Unnamed: 23' : 'Cantidad',
    'Unnamed: 28' : 'Modelo',
    'Unnamed: 29' : 'Area'}, 
        inplace=True)

# remove all unecesary columns, in this case its all 
# the columns with unnmed in the text witch made it easy.
for col in df.columns:
    if 'Unnamed' in col:
        del df[col]


# Concatenate columns based on name, adding a sapece as a separator.
df['unique'] = df['Nº de circuito'] + ' ' + df['NP']


df.drop_duplicates(subset ="unique",keep = 'first', inplace = True)

part_num = df.NP.unique()


# Save the data frame to excel format, the index 
# parameter removes the index that pandas adds.
df.to_excel('edited.xlsx', 'sheet a',index=False)

#print(part_num)

excel_book = pxl.load_workbook('edited.xlsx')

with pd.ExcelWriter('edited.xlsx', engine='openpyxl') as writer:
    writer.book = excel_book
    writer.sheets = {
        worksheet.title: worksheet
        for worksheet in excel_book.worksheets
    }
    secondMockData = { 'NP': part_num }
    secondMockDF = pd.DataFrame(secondMockData)
    secondMockDF.to_excel(writer, 'sheetB', index=False)
    writer.save()



