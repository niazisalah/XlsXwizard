import pandas
import openpyxl
from Model import dataFrame_file
from Model import analyseur_cellules

def parseur(fichier_specification):
    spec=dataFrame_file.openfile(fichier_specification)

    #on determine la dimension du fichier de specification

    max_l=spec.max_row
    max_c=spec.max_column

    #on va determiner les bloques de fichiers
    # sous forme Colomne de début et colomne de fin ,ligne début, ligne de fin

    file1_spec=[1,0,1,0]
    file2_spec=[0,0,1,0]
    file_result_spec=[0,0,1,0]


#remplir les données sur les colomnes
    i=1
    while spec.cell(row=1,column=i).value!=None:
        i=i+1

    file1_spec[1] = i-1
    i = i + 1
    #ON DECALE une colomne vide entre chaque specification

    #On décale de deux une colomn vide et on pass a la prochaine
    i=file1_spec[1]+2
    # on retrouve la premiére colomn qui concerne le fichier 2
    file2_spec[0] = i

    #on avance petit a petit
    while spec.cell(row=1,column=i+1).value!=None:
        i=i+1
    file2_spec[1]=i
    file_result_spec[0] = i+2
    i=i+1

    #on continue d avancer
    while spec.cell(row=1,column=i+1).value!=None:
        i=i+1
    file_result_spec[1] = i

    #on determine le nombre de lignes
    j = 1
    while spec.cell(row=j, column=1).value != None:
        j = j + 1

    file1_spec[3] = j-1
    j = 1


    # Maintena la spécification du fichier 2
    while spec.cell(row=j, column=file2_spec[0]).value != None:
        j = j + 1
    file2_spec[3] = j-1

    j =  1

    # on continue d avancer
    while spec.cell(row=j, column=file_result_spec[0]).value != None:
        j = j + 1
    file_result_spec[3] = j-1




#la structure suivant va permettre de definir les limites
    return [file1_spec,file2_spec,file_result_spec]

#------------------------------------------------------------
#fonction pour determiner les dimensions des specifications
#------------------------------------------------------------
def dimension(list):
    l = [] #lignes
    c=[] #colomnes
    c.append(list[0][1] - list[0][0])
    c.append(list[1][1] - list[1][0])
    c.append(list[2][1] - list[2][0])

    l.append(list[0][3] - list[0][2])
    l.append(list[1][3] - list[1][2])
    l.append(list[2][3] - list[2][2])


    return l+c
#----------------------------------------
#-------comparer deux listes-------------
#----------------------------------------

def compare_listes(liste1, liste2):
    if len(liste1) != len(liste2):
        return False
    for element in liste1:
        if element not in liste2:
            return False
    return True

#----------------------------------------
#ajouter les elements de deux listes-----
#----------------------------------------

def add_lists(list1, list2):
    # Vérifie que les deux listes ont la même longueur
    if len(list1) != len(list2):
        raise ValueError("Les deux listes doivent avoir la même longueur")

    # Initialise une nouvelle liste pour stocker le résultat
    result = []

    # Parcourt les listes et additionne les éléments correspondants
    for i in range(len(list1)):
        sum = list1[i] + list2[i]
        result.append(sum)

    return result

#----------------------------------------
#-----------les detecteurs---------------
#----------------------------------------

def determiner_fonction(fichier_specification):
    #a l interieur on va determiner s il ya une spécification et retourner

    try_combinaison_vertical(fichier_specification)

    return 0

#----------------------------------------
#----------------------------------------

def try_combinaison_vertical(fichier_specification):

    list = parseur(fichier_specification)
    dim = dimension(list)

    if dim[2] > dim[1] and dim[2] > dim[0]:
        return True
    else:
        return False

#----------------------------------------
#----------------------------------------

def try_combinaison_horizental(fichier_specification):
    list = parseur(fichier_specification)
    dim = dimension(list)
    #on test si la dimension le nombre de lignes
    if dim[5] > dim[4] and dim[5] > dim[3]:
        return True
    else:
        return False

#----------------------------------------
#----------------------------------------

def detecter_xlsx_fonctions(fichier_specification):
    list = parseur(fichier_specification)
    dataspecs = analyseur_cellules.get_column_values(list[2][0], fichier_specification)
    fonctions=[]
    for dataspec in dataspecs:
        fonctions.append(analyseur_cellules.get_fucntion_name(dataspec))



    return fonctions
#----------------------------------------
#-----Fonction pour parser le vote-------
#----------------------------------------

def try_vote(fichier_specification):


    list = parseur(fichier_specification)
    print(list)
    dataspec=analyseur_cellules.get_column_values(list[2][0],fichier_specification)
    datafile1 = analyseur_cellules.get_column_values(list[0][0], fichier_specification)
    datafile2 = analyseur_cellules.get_column_values(list[1][0], fichier_specification)
    print(datafile1)
    print(datafile2)
    print(dataspec)

    if compare_listes(pandas.unique(datafile1+datafile2),dataspec):
        dataspec = analyseur_cellules.get_column_values(list[2][1], fichier_specification)
        datafile1 = analyseur_cellules.get_column_values(list[0][1], fichier_specification)
        datafile2 = analyseur_cellules.get_column_values(list[1][1], fichier_specification)

        if compare_listes(dataspec,add_lists(datafile1,datafile2)):
            return True
        else:
            return False

    else:
        return False




def try_dublicats(fichier_specification):
    return True

def correction(fichier_specification):
    return True
