import pandas as pd
import glob
from tqdm import tqdm

file_list = glob.glob("./rapports_récolement/*.xls")

def find_one_file(path):
    xls = pd.ExcelFile(path)
    df1 = pd.read_excel(xls, 'Récolement')
    df2 = pd.read_excel(xls, 'Mal placé') 
    # Tidy data
    df1['Code-barres'] = df1['Code-barres'].astype(str)
    df2['Code-barres'] = df2['Code-barres'].astype(str)
    df1['Code-barres'] = df1.apply(lambda row: row['Code-barres'].replace('.0', ''), axis=1)
    df2['Code-barres'] = df2.apply(lambda row: row['Code-barres'].replace('.0', ''), axis=1)
    
   
    df1_final = df1.dropna(subset=['Etat copie', 'Précision exemplaire'], how='all') # Document avec Etat copie ou Précision exemplaire pas vide
    df1_final = pd.concat([df1_final, df1[df1['Description']=="Exemplaires absents dans le fichier"]], ignore_index=True, axis=0) # Exemplaires absents dans le fichier
    df1_final = pd.concat([df1_final, df1[(df1['Description']=="Exemplaires en prêt")]], ignore_index=True, axis=0)  # Exemplaires en prêt
    df1_final = pd.concat([df1_final, df1[(df1['Description']=="Exemplaire avec demandes ou réservations")]], ignore_index=True, axis=0) # Exemplaires avec demandes ou réservations
    df1_final = pd.concat([df1_final, df1[(df1['Type circulation']!="Prêt externe")]], ignore_index=True, axis=0) # Documents qui ne sont pas en prêt externe

    df1_final = pd.concat([df1_final, df1[(df1['Code-barres'].str.len() != 14)]], ignore_index=True, axis=0) # Codument dont l'ISBN n'a pas 14 caractères
    df2_final = df2.dropna(subset=['Etat copie', 'Précision exemplaire'], how='all') # Document avec Etat copie ou Précision exemplaire pas vide
    df2_final = pd.concat([df2_final, df2[(df2['Description']=="Exemplaire avec cote hors de l'intervalle")]], ignore_index=True, axis=0) # Exemplaires avec cote hors de l'intervalle
    df2_final = pd.concat([df2_final, df2[(df2['Description']=="Exemplaires en prêt")]], ignore_index=True, axis=0) # Exemplaires en prêt
    df2_final = pd.concat([df2_final, df2[(df2['Description']=="Exemplaire avec demandes ou réservations")]], ignore_index=True, axis=0) # Exemplaires avec demandes ou réservations
    df2_final = pd.concat([df2_final, df2[(df2['Type circulation']!="Prêt externe")]], ignore_index=True, axis=0) # Documents qui ne sont pas en prêt externe
    df_rec_malpla = pd.concat([df1_final, df2_final], ignore_index=True, axis=0)
    return df_rec_malpla

master_df = pd.DataFrame(columns=['Description', 'Section',	'Cote',	'Spécification', 'Séquence', 'Exemplaire', 'Code-barres', 'Type circulation', 'Etat copie', 'Précision exemplaire', 'Id.', 'Description isbd'])

for file in tqdm(file_list):
    master_df = pd.concat([master_df, find_one_file(file)], ignore_index=True, axis=0)

master_df = master_df.sort_values(by=['Cote'], ignore_index=True)
master_df = master_df.drop_duplicates(subset=['Code-barres'])
master_df.to_excel("Suite de récolement.xlsx")
print('Fini !')
print('Nombre de lignes :' + str(len(master_df.index)))