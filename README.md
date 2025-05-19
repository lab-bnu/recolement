# Script pour fichiers de récolement

Script utilisé par la Bibliothèque Nationale Universitaire de Strasbourg pour traiter les fichiers de récolement.

## Première installation (Windows)

Télécharger et installer [Python](https://www.python.org/downloads/)

Télécharger le [script](https://github.com/lab-bnu/recolement/blob/main/recolement.py), en cliquant sur le bouton 'Download raw file', et le placer dans le dossier de votre choix.

Dans ce même dossier, ajouter un dossier vide nommé 'rapports_récolement'.

Ouvrir un Powershell. Pour faire ça, dans l'explorateur de fichiers et dans le dossier où le script se trouve, il faut cliquer sur le chemin d'accès puis entrer `cmd`. Une fenêtre Powershell s'ouvre.

Dans cette fenêtre, vérifier que pythonn est installé avec `python --version`

Finalement, pour faire fonctionner le script, il faut télécharger les modules : `pip install pandas openpyxl`

Et voilà, tout est prêt !


## Utilisation

Placez dans le dossier 'rapports_récolement' les rapports à traiter, au format .xls

Dans le Powershell, utlisez la commande  

`python recolement.py` 

Les résultats se trouvent dans le fichier 'Suite de récolement.xlsx'
