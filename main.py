
from Model import vote
import pandas



if __name__ == '__main__':

    #combinetwofiles.combiner("file1.xlsx","file2.xlsx")
    #combine.combiner("file1.xlsx","file2.xlsx")
    #combine.combiner_tout(['file1.xlsx','file2.xlsx','file3.xlsx'],combine.combiner_h)
    #combine.create_xlsx_file("file4.xlsx")
    #combine.create_xlsx_file("file5.xlsx")
    #combine.combiner_vertical("file1.xlsx", "file2.xlsx")
    #print(vote.calculer_vote("file3.xlsx","file2.xlsx"))
    #print(vote.alllignes("file4.xlsx"))
    #print(vote.alllignes("file5.xlsx"))
    #print(vote.ligne_value("prof1", "file5.xlsx"))
    #print(vote.ligne_value("prof2","file5.xlsx"))
    #print(vote.ligne_value("prof3", "file5.xlsx"))
    #vote.calculer_vote("file1.xlsx","file2.xlsx")

    vote.calculer_all_votes(["file1.xlsx","file2.xlsx","file4.xlsx"])
    #vote.calculer_vote("file1.xlsx","file2.xlsx")
    #print(duplacated_data.compare_ligne(['a',0],['a',0]))
    #duplacated_data.traiter_duplicats("file1.xlsx")


    #remove_duplicates('file1.xlsx')
    # Replace 'example.xlsx' with the path to your xlsx file
    #remove_duplicates('file1.xlsx')
    #combine.combiner_tout(["file1.xlsx","file2.xlsx"],combine.combiner_datahseet_v)
    #combine.DataFrame_tofile(combine.combiner_dataframe_h("file1.xlsx","file2.xlsx"))





