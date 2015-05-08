'''
Created on Aug 22, 2014


@author: Brandon_Trinh


requires python 3.4 and xlrd


Read xls file
for every column go through to get max char and column name
create table
insert each row
'''
import xlrd
import os
import io


print("Script start")


fileOutName = "createDB"
#xlsdir = "C:/Users/brandon_trinh/Desktop/data"
xlsdir = 'C:/Users/brandon_trinh/workspace/ExcelScript'


try:
os.remove(fileOutName+".sql")

except OSError:
pass
fileWriter = io.open(fileOutName+".sql", "a", encoding="utf-8-sig")


def writeSqlFile(path, file):
emptyColumn = False;
duplicateColumn = False;
existingLocaleColumn = False;

workbook = xlrd.open_workbook(path, encoding_override="utf-8-unicode")
pathSplit=path.split("/");
lenPathSplit=len(pathSplit);
tableName=pathSplit[lenPathSplit-1].split(".");
if pathSplit[lenPathSplit-2] not in ("pilot", "production"):
tableName=pathSplit[lenPathSplit-2] +"_"+ tableName[0]
else:
tableName = pathSplit[lenPathSplit-3] + "_" + pathSplit[lenPathSplit-2] +"_"+ tableName[0]

bookLength = len(workbook.sheets())

output = ""
output2 = ""
columnNames = []
columnLengths = []
notSummary = False;
for x in range(0,bookLength):
sheet = workbook.sheet_by_index(x)
if sheet.name != "Summary" and notSummary == False:
notSummary=True
for i in range(sheet.ncols):
currCell=sheet.cell(0,i)
if(currCell.value != "Test_End"):
if(currCell.value == "Locale"):
existingLocaleColumn = True;
print("ERROR: " + tableName + " has locale as a column.")
if(currCell.value == ""):
emptyColumn = True;
print("ERROR: " + tableName + " has an empty column name.")
if any(currCell.value in s for s in columnNames):
duplicateColumn = True
print("ERROR: " + tableName + " has a a duplicate column name: " + currCell.value + ".")
columnNames.append("`"+currCell.value.rstrip("\n")+"`");
maxLen = 0
for j in range(sheet.nrows-1):
currCell=sheet.cell(j+1,i)
cellLen=len(str(currCell.value))
if(cellLen>maxLen):
maxLen=cellLen
columnLengths.append(maxLen);
else:
columnNames.append("`Locale`");
columnLengths.append(2);
output += "\nCREATE TABLE " + tableName + "(\n"
for b in range(0, len(columnNames)):
if b != len(columnNames)-1:
output+= "\t" + columnNames[b] + " varchar(" + str(columnLengths[b]) + "),\n"
else:
output+= "\t" + columnNames[b] + " varchar(" + str(columnLengths[b]) + ")\n"
output += ");\n"

output2 += "\nINSERT INTO " + tableName + " VALUES "
for x in range(0,bookLength):
sheet = workbook.sheet_by_index(x)
if sheet.name != "Summary":
for i in range(sheet.nrows-1):
if(sheet.nrows>i+1 and sheet.ncols>2):
currCell=sheet.cell(i+1,2)
if(currCell.value != ""):
output2 += "("
for j in range(sheet.ncols-1):
currCell=sheet.cell(i+1,j)
value = str(currCell.value)
value = value.replace("\"", "")
output2 += "\"" + value + "\", "
'''
if(j<sheet.ncols-1):
output2 += "\"" + value + "\", "
else:
output2 += "\"" + value + "\""
'''
output2 += "\"" + sheet.name + "\""
output2 += "),\n"
output2 = output2[:-2]
output2 +=";"
if (emptyColumn is False) and (duplicateColumn is False) and (existingLocaleColumn is False):
file.write(output)
file.write(output2)
else:
print("did not write")

# search through list directory to find all xls files and write sql statements to insert data into a MySQL DB
pathList = []
for root, dirs, files in os.walk(xlsdir):
for file in files:
if file.endswith(".xls"):
pathName = os.path.join(root, file)
pathList.append(pathName.replace("\\", "/"))
for x in range(0, len(pathList)):
writeSqlFile(pathList[x], fileWriter)
fileWriter.close()
print("script finished")


