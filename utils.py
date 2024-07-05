import re
import os 
from tkinter import messagebox

Dict = { 'Secteur' : 'E2', 'Sous-secteur' : 'F2', 'Cours de bourse' : 'B5', 'WACC' : 'S144', 'PGR' : 'S145', 'DCFvalue' : 'U196', 'Upside/Downside' : 'U198', 
'Organic24' : 'U50','Organic25' : 'U51', 'Organic26' : 'U52', 'Organic27' : 'U53' }

Valide = ['WACC', 'PGR', 'Upside/Downside', 'Organic24', 'Organic25', 'Organic26', 'Organic27']


def get_resource_path(relative_path):# inutile au finale 
    try:
        base_path = os.path.join(os.environ['_MEIPASS'])
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def read_case(options):
    colonne = []
    for (case) in options:
        if case['checked']:
            colonne.append(case['label'])
    if colonne == [] : 
        for (case) in options:
            colonne.append(case['label'])
    return colonne

def colonne_to_index(colonne):
    index = 0
    for c in colonne.upper():
        index = index * 26 + (ord(c) - ord('A') + 1)
    return index - 1

def str_separate(ch):
    pattern = re.compile(r'[A-Za-z]+\d*')
    ch = pattern.findall(ch)
    return ch

def transformation_coordonnees(cellule_reference):
    cellule_reference = str_separate(cellule_reference)
    coordonnees = []
    for cellule in cellule_reference:
        match = re.match(r"([A-Za-z]+)([0-9]+)", cellule)
        if not match:
            messagebox.showerror("Erreur", "La référence de cellule est invalide.")
            return
        colonne, ligne = match.groups()
        ligne = int(ligne) - 2  # Les index de pandas commencent à 0
        colonne_index = colonne_to_index(colonne)
        coordonnees.append([ligne, colonne_index]) 
    return coordonnees

def transformation_colonne(colonnes):
    for i in range(len(colonnes)):
        colonnes[i] = colonnes[i].lower().replace(' ', '')
    return colonnes

def ajuster_pourcentage(df, valide):
    colonnes =  df.columns
    for colonne in colonnes:
        if colonne in valide :
            df[colonne] = df[colonne].apply(lambda x: f"{float(x) * 100:.2f}%")
    return df