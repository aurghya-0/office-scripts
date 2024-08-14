import pandas as pd

df_cse = pd.read_csv("CSE.csv")
df_cse_ds = pd.read_csv("CSE DS.csv")
df_cse_aiml = pd.read_csv("CSE AIML.csv")

df_cse = df_cse.fillna(method='ffill')
df_cse_aiml = df_cse_aiml.fillna(method='ffill')
df_cse_ds = df_cse_ds.fillna(method='ffill')

df_cse.to_excel("CSE_Modified.xlsx", index=False)
df_cse_ds.to_excel("CSE_DS_Modified.xlsx", index=False)
df_cse_aiml.to_excel("CSE_AIML_Modified.xlsx", index=False)
