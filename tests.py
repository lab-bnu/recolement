import pandas as pd

print('Calcul de la différence')
target_df = pd.read_excel(pd.ExcelFile('200_Suite de récolement.xlsx'))
master_df = pd.read_excel(pd.ExcelFile('Suite de récolement.xlsx'))

target = list(target_df['Code-barres'].astype(str))
master = list(master_df['Code-barres'].astype(str))

print('Present in output but not in target (' + str(len(set(master) - set (target))) + ')')
print(set(master) - set (target))

print('Present in target but not in output (' + str(len(set(target) - set (master))) + ')')
print(set(target) - set (master))