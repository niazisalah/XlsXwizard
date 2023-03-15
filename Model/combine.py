import pandas
from functools import reduce
import dataFrame_file



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
    return list(map(dataFrame_file.file_toDataFrame,fichiers))

#---------------------------------------------------------------------------------
#---------prend deux valeur en entré une fonction et une liste de fichier---------
#---------------------------------------------------------------------------------

def combiner_tout(fichiers,f):

    DataFrames =files_toDataFrames(fichiers)
    DataFrame=reduce(lambda x,y:f(x,y),DataFrames)

    return dataFrame_file.dataFrame_tofile(DataFrame)