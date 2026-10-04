import pandas as pd

df = pd.read_excel('BASE HISTORICA.xlsx', sheet_name='JUL-AGO_2024')

print(df.info())
print('==============================================')
print(df.describe())
print('==============================================')

selection = df.iloc[:, 0:20]
print(selection)