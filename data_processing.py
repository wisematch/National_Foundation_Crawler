import pandas as pd

data_path = r'国自然17-22_筛选.xlsx'
dataset = pd.read_excel(data_path, sheet_name=0)
df = pd.DataFrame(dataset)

