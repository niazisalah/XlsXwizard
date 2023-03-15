import pandas
from functools import reduce
import openpyxl


#---------------------------------------------------------------------------------
#-----------fonction pour désactiver les border et le bold de l'entête------------
#---------------------------------------------------------------------------------

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



#---------------------------------------------------------------------------------
#--------------fonction qui convertis un fichier xlsx to DataFrame----------------
#---------------------------------------------------------------------------------

def file_toDataFrame(fichier):
    ds= pandas.read_excel(fichier)
    return ds

#---------------------------------------------------------------------------------
#------------------------------Combinaison vertical-----------------------------
#---------------------------------------------------------------------------------

def combiner_datahseet_v(df1, df2):
    # lire le premier fichier

    #f1 = pandas.read_excel(fichier1)

    # Lire le deuxiéme fichier
    #f2 = pandas.read_excel(fichier2)

    # combiner les deux fichier a l'aide de concat
    return pandas.concat([df1, df2],axis=0)

#---------------------------------------------------------------------------------
#--------------------La fonction qui convertie la feuille en fichier--------------
#---------------------------------------------------------------------------------

def DataFrame_tofile(DataFrame):
    create_xlsx_file("result_.xlsx")
    DataFrame.to_excel("result_.xlsx", index=False)
    disable_bold_border("result_.xlsx")
    return "result.xlsx"

#---------------------------------------------------------------------------------
#----------------------------Creation d'un fichier xlsx---------------------------
#---------------------------------------------------------------------------------
def create_xlsx_file(file_name):

    wb = openpyxl.Workbook()
    ws = wb.active
    wb.save(file_name)



#---------------------------------------------------------------------------------
#----------------------------Combinaison horizentale------------------------------
#---------------------------------------------------------------------------------

def combiner_dataframe_h(fichier1,fichier2):
    # Charger le premier fichier
    df1 = pandas.read_excel(fichier1)

    # Charger le deuxiéme fichier
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


    return pandas.concat([df1, df2], axis=1)

#---------------------------------------------------------------------------------
#---------Prend une liste de fichier et retourne une liste de DataFrame-----------
#---------------------------------------------------------------------------------

def files_toDataFrames(fichiers):
    return list(map(file_toDataFrame,fichiers))

#---------------------------------------------------------------------------------
#---------prend deux valeur en entré une fonction et une liste de fichier---------
#---------------------------------------------------------------------------------

def combiner_tout(fichiers,f):

    DataFrames =files_toDataFrames(fichiers)
    DataFrame=reduce(lambda x,y:f(x,y),DataFrames)

    return DataFrame_tofile(DataFrame)