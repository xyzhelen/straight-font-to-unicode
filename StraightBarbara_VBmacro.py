import csv

with open('../transliteration mappings - straight barbara.csv', encoding="utf-8") as f:
    data = [tuple(line) for line in csv.reader(f)]

with open('../word macros/to_StraightBarbara.bas', "w") as outF:
    outF.write("Sub to_StraightBarbara()"+"\n")
    outF.write("\'"+"\n")
    outF.write("\'Straight Barbara Macro"+"\n")
    outF.write("\'"+"\n")
    outF.write("\'Last updated: 5-May-2021 by Helen Zhang\n")
    outF.write("\'"+"\n")
    outF.write("\t"+"Selection.Find.ClearFormatting"+"\n")
    outF.write("\t"+"Selection.Find.Replacement.ClearFormatting"+"\n")
    #outF.write("\t"+"Selection.Find.Font.Name = \"Straight\""+"\n")
    outF.write("\t"+"Selection.Find.Replacement.Font.Name = \"Times New Roman\""+"\n")
    outF.write("\t" + "Selection.Find.Format = True" + "\n")
    outF.write("\t" + "Selection.Find.MatchCase = True" + "\n")

    # convert these in Straight font
    for row in data[1:29]:
        outF.write("\tWith Selection.Find"+"\n")
        #outF.write("\t\t'" + row[4]+"\n")
        outF.write("\t\t" + ".Text = "+row[1] + "\n")
        outF.write("\t\t" + ".Replacement.Text = "+row[3] + "\n")
        #outF.write("\t\t\'set Format=True so it only works on Straight font+\n")
        #outF.write("\t\t" + ".Format = True" + "\n")
        outF.write("\t"+"End With" + "\n")
        outF.write("\t"+"Selection.Find.Execute Replace:=wdReplaceAll" + "\n")

    #outF.write("\t"+"'change all remaining Straight font characters into BC Sans"+"\n")
    #outF.write("\t"+"Selection.Find.Font.Name = \"Straight\""+"\n")
    #outF.write("\t"+"Selection.Find.Replacement.Font.Name = \"Times New Roman\""+"\n")
    #outF.write("\t"+"With Selection.Find"+"\n")
    #outF.write("\t\t"+".Text = \"\""+"\n")
    #outF.write("\t\t"+".Replacement.Text = \"\""+"\n")
    #outF.write("\t\t"+".Format = True"+"\n")
    #outF.write("\t"+"End With"+"\n")
    #outF.write("\t"+"Selection.Find.Execute Replace:=wdReplaceAll"+"\n")
    #outF.write("\t\'clear formatting dialog for the user")
    #outF.write("\t"+"Selection.Find.ClearFormatting"+"\n")
    #outF.write("\t"+"Selection.Find.Replacement.ClearFormatting"+"\n")
    #outF.write("\t"+"Selection.Find.Format = False"+"\n")
    outF.write("End Sub")

    outF.close()

