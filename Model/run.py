import os
from ProjetSynthese.Model import parseur, vote, combine


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
        if i.endswith("template.xlsx"):
            list.remove(i)
            return i


def traiter(template,listfile):
    if(parseur.trycombv(template)):
        combine.combiner_tout2(listfile)


    if(parseur.try_vote(template)):
        vote.calculer_all_votes(listfile)


    if(parseur.trycombh(template)):
        combine.combiner_tout(listfile)

