import pandas as pd

df = pd.read_csv("Fall 2019.csv", skiprows=3)
print(df.head())