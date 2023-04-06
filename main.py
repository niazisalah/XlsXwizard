

from Model import run
from ProjetSynthese.Model import parseur
from Model import combine

if __name__ == '__main__':
   #combine.combiner_tout(["file1.xlsx", "file2.xlsx","file3.xlsx"])
   #combine.combiner2("file1.xlsx", "file2.xlsx")
   #print(parseur.trycombh("template.xlsx"))
   run.traiter("Model/xlsxfiles/template.xlsx",["Model/xlsxfiles/file1.xlsx","Model/xlsxfiles/file2.xlsx"])









