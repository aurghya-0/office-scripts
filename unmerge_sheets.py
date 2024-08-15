import os
import pandas as pd

directory = os.fsencode("excel")

excel_files = []

for file in os.listdir(directory):
    filename = os.fsdecode(file)
    print(filename)
    if filename.endswith(".xlsx"):
        excel_files.append(filename)

print(excel_files)

# df_cse = pd.read_csv("CSE.csv")

# df_cse = df_cse.fillna(method='ffill')

# df_cse.to_excel("CSE_Modified.xlsx", index=False)
