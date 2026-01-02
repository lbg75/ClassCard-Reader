import pandas as pd
p='entries.xlsx'
try:
    df = pd.read_excel(p, engine='openpyxl')
    print('file:', p)
    print('rows:', len(df))
    if len(df)>0:
        print('\nlast 10 rows:')
        print(df.tail(10).to_string(index=False))
except Exception as e:
    print('error reading', p, e)
