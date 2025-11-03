import pandas as pd
df = pd.read_csv('faltas.csv', sep=',')
df['Data'] = df['Data'].str.split(', ', expand=True)[1]
df['Data'] = pd.to_datetime(
  
    df['Data'],
    format='%d/%m/%Y',
    errors='coerce' 
)
df.drop(columns=['Data'], inplace=True)
print(df.head())