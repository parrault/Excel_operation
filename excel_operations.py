import pandas as pd
import json
import os
import datetime
from PySide6.QtWidgets import QMessageBox
from utils import transformation_coordonnees, read_case,ajuster_pourcentage,get_resource_path,  Valide 
from styling import ajuster_largeur_colonnes_et_styles
import time
import re 


def create_dict(file):
    file_path = get_resource_path(file) 
    with open(file_path, 'r') as f:
        config_data = json.load(f)
    liste_variables = config_data['options']
    Dict = dict()
    for element in liste_variables :
        Dict[ element['label'] ] =  element['cell']
    return Dict

def lire_feuille_par_nom(fichier_excel, mot_cle, ext):
    """
    Lit une feuille d'un fichier Excel dont le nom correspond au mot clé (ignorant les espaces et majuscules) 
    et la retourne sous forme de dataframe.

    :param fichier_excel: Chemin du fichier Excel.
    :param mot_cle: Mot clé pour identifier la feuille.
    :param ext: Extension du fichier (xls ou xlsx).
    :return: DataFrame contenant la feuille correspondante ou None si aucune feuille ne correspond.
    """
    # Lire toutes les feuilles du fichier Excel
    if ext in ['.xls', '.XLS']:
        feuilles = pd.read_excel(fichier_excel, sheet_name=None, engine='xlrd')
    else:
        feuilles = pd.read_excel(fichier_excel, sheet_name=None, engine='openpyxl')
    
    # Préparer le mot clé pour la comparaison
    mot_cle_normalise = re.sub(r'\s+', '', mot_cle.strip()).lower()
    
    # Trouver la feuille dont le nom correspond au mot clé
    for feuille_nom, df in feuilles.items():
        feuille_nom_normalise = re.sub(r'\s+', '', feuille_nom.strip()).lower()
        if mot_cle_normalise == feuille_nom_normalise:
            return df
    
    # Retourner None si aucune feuille ne correspond
    return None


def excel_colonne(dossier_choisi, options):
    dossier = dossier_choisi
    colonnes = read_case(options)
    resultat = []
    Dict = create_dict('config.json')

    if not colonnes or not dossier:
        QMessageBox.warning(None, "Attention", "Veuillez remplir tous les champs.")
        return

    for fichier in os.listdir(dossier):
        try:
            ext = os.path.splitext(fichier)[1].lower()
            if ext in ['.xlsx', '.xls', '.XLS', '.XLSX'] and fichier != 'resultats_recherche_colonnes.xlsx':
                chemin_fichier = os.path.join(dossier, fichier)
                #print(f"Processing file: {chemin_fichier}")
                
                # Boucle de réessai pour ouvrir le fichier
                retry_count = 5
                while retry_count > 0:
                    try:
                        df = lire_feuille_par_nom(chemin_fichier, 'output', ext)
                        if df is None :
                            df = lire_feuille_par_nom(chemin_fichier, 'valooutput', ext)
                        if df is None:
                            df =  pd.read_excel(chemin_fichier, engine='xlrd' if ext in ['.xls', '.XLS'] else 'openpyxl')
                        break  # Si la lecture réussit, sortir de la boucle
                    except PermissionError as e:
                        retry_count -= 1
                        if retry_count == 0:
                            #print(f"Erreur de permission pour le fichier {chemin_fichier}: {e}")
                            QMessageBox.warning(None, "Attention", f"Le fichier {chemin_fichier} est ouvert ailleurs. Veuillez le fermer et réessayer.")
                            raise
                        #print(f"Le fichier {chemin_fichier} est ouvert ailleurs. Réessayer dans 3 secondes...")
                        time.sleep(3)  # Attendre 3 secondes avant de réessayer
                
                new_dict = {}
                new_dict['entreprise'] = fichier.replace('.xlsx', '').replace('.xls', '').replace('.XLS', '').replace('XLSX', '')
                new_dict['lien'] = f'=HYPERLINK("{chemin_fichier}", "lien vers {fichier}")'
                
                for colonne in colonnes:
                    try:
                        cellule = Dict[colonne]
                        coordonnee = transformation_coordonnees(cellule)
                        new_dict[colonne] = df.iat[coordonnee[0][0], coordonnee[0][1]]
                    except KeyError as e:
                        print(f"Erreur lors de l'accès à la cellule '{cellule}' pour la colonne '{colonne}': {e}")
                        QMessageBox.warning(None, "Attention", f"Problème d'accès à la cellule '{cellule}' pour la colonne '{colonne}' dans le fichier {fichier}.")
                        return
                    except IndexError as e:
                        print(f"Erreur d'index pour la cellule '{cellule}' dans le fichier {fichier}: {e}")
                        QMessageBox.warning(None, "Attention", f"Problème d'indexation pour la cellule '{cellule}' dans le fichier {fichier}.")
                        return
                    except ValueError as e:
                        QMessageBox.warning(None, "Attention", f"Erreur de type de cellule pour la colonne '{colonne}' dans le fichier {fichier}.")
                        return
                resultat.append(new_dict)
        except Exception as e:
            print(f"Erreur lors de la lecture du fichier {fichier}: {e}")
            QMessageBox.warning(None, "Attention", f"Le fichier {fichier} pose problème: {e}")

    if not resultat:
        print("Aucun résultat trouvé.")
        QMessageBox.information(None, "Terminé", "Aucun résultat trouvé.")
        return

    resultats_df = pd.DataFrame(resultat)
    resultats_df = ajuster_pourcentage(resultats_df, Valide)
    nom_fichier_resultats = 'resultats_recherche_colonnes.xlsx'
    nom_fichier_resultats = os.path.join(dossier, nom_fichier_resultats)
    try:
        resultats_df.to_excel(nom_fichier_resultats, index=False)
        ajuster_largeur_colonnes_et_styles(nom_fichier_resultats)
        QMessageBox.information(None, "Terminé", f"Recherche terminée. Résultats enregistrés dans {nom_fichier_resultats}")
        os.startfile(nom_fichier_resultats)
    except PermissionError as e:
        print(f"Erreur de permission lors de la tentative de sauvegarde du fichier {nom_fichier_resultats}: {e}")
        QMessageBox.warning(None, "Attention", f"Impossible de sauvegarder le fichier {nom_fichier_resultats} car il est ouvert ailleurs. Veuillez le fermer et réessayer.")
