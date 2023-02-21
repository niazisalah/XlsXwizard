from Model import combine
from Model import vote
from Model import duplacated_data
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
    vote.calculer_vote("file1.xlsx","file2.xlsx")

    #print(duplacated_data.compare_ligne(['a',0],['a',0]))
    #duplacated_data.traiter_duplicats("file1.xlsx")
    #duplacated_data.clean_dup("file1.xlsx")



