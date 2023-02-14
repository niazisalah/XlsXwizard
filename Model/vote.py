import openpyxl
import pandas
from Model import combine
from functools import reduce


#Fonction pour lire le fichier xlsx
def openfile(fichier):
    file=openpyxl.load_workbook(fichier)
    return file.active


#Fonction max pour returner le Maximum
def max(x,y):
    if x>=y:
        return x
    else:
        return y


#fonction qui récupére les ligne d un fichiers
def alllignes(file,col=1):
    list=[]
    f=openfile(file)
    max_l=f.max_row
    max_c=f.max_column
    for i in range(1,max_l+1):
        list.append(f.cell(row=i,column=col).value)
    return list

#fonction qui retourne la valeur numériqued une ligne

def ligne_value(description,fichier,col=1,col2=2):
    f=openfile(fichier)
    for i in range(1,f.max_row+1):
        if description==f.cell(row=i,column=col).value:

            return f.cell(row=i,column=col2).value

    return 0



# fonction qui calcule la somme des votes


def calculer_vote(fichier1,fichier2,fichier3="result_vote.xlsx",col=1,col2=2):

    f1=openfile(fichier1)
    f2=openfile(fichier2)

    combine.create_xlsx_file(fichier3)

    file3=openpyxl.load_workbook(fichier3)
    f3=file3.active

    lignes=pandas.unique(alllignes(fichier1)+alllignes(fichier2))
    for i in range(len(lignes)):
        #on ecrit la valeur de la premiere colonne dans la ligne i
        f3.cell(row=(i+1),column=col,value=lignes[i])

        #on rajoute la valeur
        f3.cell(row=(i+1),column=col2,value=ligne_value(lignes[i],fichier1)+ligne_value(lignes[i],fichier2))

    #------> save File 3
    file3.save(fichier3)

    return fichier3


def calculer_all_votes(fichiers):
    return reduce(lambda x,y:calculer_vote(x,y),fichiers)