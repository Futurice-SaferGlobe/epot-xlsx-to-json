import sys
import xlrd


def main():
    for arg in sys.argv[1:]:
        filename = arg
    book = xlrd.open_workbook(filename)
    f = open("XLSX_TO_JSON.json", "w")
    sheetcount = book.nsheets-2
    f.write("[" +"\n")
    f.write("    {" +"\n")
    f.write("    \""+ str(book.sheet_by_index(book.nsheets-2).cell(0,0).value) +"\""+":" +" \"" + str(book.sheet_by_index(book.nsheets-2).cell(0,1).value)+ "\"" + "," + "\n")
    f.write("    \""+ str(book.sheet_by_index(book.nsheets-2).cell(1,0).value) +"\""+":" +" \"" + str(book.sheet_by_index(book.nsheets-2).cell(1,1).value)+ "\"" + "," + "\n")
    f.write("    \""+ str(book.sheet_by_index(book.nsheets-2).cell(2,0).value) +"\""+":" +" \"" + str(book.sheet_by_index(book.nsheets-2).cell(2,1).value)+ "\"" + "," + "\n")
    f.write("    \"data\": [" + "\n")
    for I in range(sheetcount):
        currsheet = book.sheet_by_index(I)
        f.write("        {")
        f.write("\n")
        f.write("        \"index\""+":" +" "+  str(I+1) + ",")
        f.write("\n")
        #string = string.replace("\n",' ')
        f.write("        \"title\"" + ":" + " " + " \"" + str(currsheet.cell(0,1).value) + "\"" + ",")
        f.write("\n")
        string = "        \"content\"" + ": " + " \""+str(currsheet.cell(1,1).value) + "\" " + ","
        string = string.replace("\n",' ')
        f.write(string)
        f.write("\n")
        labels = list
        f.write("        \"labels\"" + ": " + "["+ "\"" +str(currsheet.cell(1,2).value) + "\""+ "]"+ ",")
        f.write("\n")
        f.write("        \"subheaders\"" + ": " + "[")
        f.write("\n")
        rows = currsheet.nrows
        cols = currsheet.ncols
        #print(rows)
        #print(cols)
        counter = 0
        for i in range(1,rows-1,2):
            counter = counter+1
            f.write("            {")
            f.write("\n")
            f.write("            \"index\""+":" +" "+  str(counter) + ",")
            f.write("\n")
            f.write("            \"title\"" + ":" + " " + " \"" + str(currsheet.cell(i+1,1).value) + "\"" + ",")
            f.write("\n")
            string1 = "            \"content\":" + " \"" + str(currsheet.cell(i+2,1).value) + "\"" + ","
            string1 = string1.replace("\n",' ')
            f.write(string1)
            f.write("\n")   
            if currsheet.cell(i+2,2).value == "": 
                f.write("            \"labels\"" + ": " + "[]")
            else:
                f.write("            \"labels\"" + ": " + "[" + "\"" + str(currsheet.cell(i+2,2).value) + "\""+ "]")
            f.write("\n")
            if i == rows-3:
                f.write("            }")
            else:
                f.write("            },")
            f.write("\n")
        counter = 0
        f.write("        ]")
        f.write("\n")
        if I != book.nsheets-3:
            f.write("        },")
        if I == book.nsheets-3: 
            f.write("        }")
        f.write("\n")

    f.write("        ]")
    f.write("\n")
    f.write("    }")
    f.write("\n")
    f.write("]")
    f.close()

if __name__ == "__main__":
    main()
