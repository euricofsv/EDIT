import pandas as pd

file_path = r"C:\Users\Eurico\Desktop\EDIT\Ambiente\Projeto1\sales.parquet" 
df = pd.read_parquet(file_path)


print(df.info())

print(df.head(15))

print(df.info())

print(df.describe())

import os

print("PORT:", os.getenv("PORT"))