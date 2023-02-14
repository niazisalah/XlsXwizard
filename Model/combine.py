import pandas
from functools import reduce
# a revoir par la suite
import openpyxl

# fonction pour désactiver les border et le bold de l'entête
def disable_bold_border(fichier):

    wb = openpyxl.load_workbook(fichier)
    sheet = wb.active


    for row in sheet.iter_rows():
        for cell in row:
            cell.font = openpyxl.styles.Font(bold=False)
            cell.border = openpyxl.styles.Border(left=openpyxl.styles.Side(border_style=None),
                                                right=openpyxl.styles.Side(border_style=None),
                                                top=openpyxl.styles.Side(border_style=None),
                                                bottom=openpyxl.styles.Side(border_style=None))


    wb.save(fichier)

#prend deux valeur en entré une fonction et une liste de fichier
def combiner_tout(liste_fichier,f):
    return reduce(lambda x,y:f(x,y),liste_fichier)

#Combinaison horizental
def combiner(fichier1, fichier2):
    # lire le premier fichier

    f1 = pandas.read_excel(fichier1)

    # Lire le deuxiéme fichier
    f2 = pandas.read_excel(fichier2)

    # combiner les deux fichier a l'aide de concat
    cobinaison = pandas.concat([f1, f2], ignore_index=True)

    #Creation de fichier de retour
    create_xlsx_file("result_comb.xlsx")

    # Enregistrer le fichier sauvegarder
    cobinaison.to_excel("result_comb.xlsx", index=False)
    disable_bold_border("result_comb.xlsx")

    return "result_comb.xlsx"

#Creation d'un fichier xlsx
def create_xlsx_file(file_name):

    wb = openpyxl.Workbook()
    ws = wb.active
    wb.save(file_name)




#Combinaison horizentale

def combiner_h(fichier1,fichier2):
    # Load the first xlsx file
    df1 = pandas.read_excel(fichier1)

    # Load the second xlsx file
    df2 = pandas.read_excel(fichier2)

    # Determiner le nombre de lignes dans chaque fichier
    rows_df1 = df1.shape[0]
    rows_df2 = df2.shape[0]

    # Pad the shorter dataframe with NaN values to make the row count equal
    if rows_df1 > rows_df2:
        padding = pandas.DataFrame(index=range(rows_df1 - rows_df2), columns=df2.columns)
        df2 = pandas.concat([df2, padding], ignore_index=True)
    elif rows_df2 > rows_df1:
        padding = pandas.DataFrame(index=range(rows_df2 - rows_df1), columns=df1.columns)
        df1 = pandas.concat([df1, padding], ignore_index=True)

    # Combiner deux  dataframes horizentalement (column-wise)
    result = pandas.concat([df1, df2], axis=1)

    # On sauvergarde dans un nouveau fichier
    result.to_excel('resultat.xlsx', index=False)
    disable_bold_border("resultat.xlsx")
    return "resultat.xlsx"

# a faire comptage de votes
# recherche de duplicata (combinaison)
#tkinter


#TClTK