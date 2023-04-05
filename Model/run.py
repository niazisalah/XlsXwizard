import os
from Model import parseur
from Model import combine
from Model import vote






def dernier_fichier(path):
    fichiers = os.listdir(path)
    if not fichiers:
        return None
    dernier_fichier = max(fichiers, key=os.path.getctime)
    return dernier_fichier

def list_files(path):

    file_list = []
    for root, dirs, files in os.walk(path):
        for filename in files:
            file_path = os.path.join(root, filename)
            if file_path.endswith('.xlsx'):
                file_list.append(file_path)
    return file_list

def selecttemplate(list):
    for i in list:
        if i=="template.xlsx":
            list.remove(i)

    return list

def traiter(template,listfile):
    if(parseur.try_combinaison_vertical(template)):
        combine.combiner_tout(listfile,combine.combiner_dataframe_v)

    if(parseur.try_vote(template)):
        vote.calculer_all_votes(listfile)

    if(parseur.try_combinaison_horizental(template)):
        combine.combiner_tout(listfile, combine.combiner_dataframe_h)
