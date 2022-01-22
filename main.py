import pandas as pd
import openpyxl
import xlrd
import easygui


pd.set_option('display.max_columns', None)
currentUndecidedFileDF = pd.read_excel('UndecidedStudents.xlsx', index_col=0)

print(currentUndecidedFileDF)
newUndecidedFile = easygui.fileopenbox()
newUndecidedFileDF = pd.read_excel(newUndecidedFile, 'Undecided Units')
print(newUndecidedFileDF)

remainingUndedecided = []
for i in range(len(currentUndecidedFileDF)-1):
    for j in range(len(newUndecidedFileDF)-1):
        if currentUndecidedFileDF.loc[i, 'Employee ID'] == newUndecidedFileDF.loc[j, 'Employee ID']:
            remainingUndedecided.append(i)

print(remainingUndedecided)
droppedUndecidedStudents = newUndecidedFileDF.drop(remainingUndedecided)
print(droppedUndecidedStudents)
writer = pd.ExcelWriter('droppedUndecidedStudents.xlsx', engine='xlsxwriter')
droppedUndecidedStudents.to_excel(writer, sheet_name='Dropped')
writer.save()

currentUndecided = []
for i in range(len(newUndecidedFileDF)):
    for j in range(len(currentUndecidedFileDF)):
        if newUndecidedFileDF.loc[i, 'Employee ID'] == currentUndecidedFileDF.loc[j, 'Employee ID']:
            currentUndecided.append(i)
newUndecidedDF = newUndecidedFileDF.drop(currentUndecided)
writer2 = pd.ExcelWriter('NewUndecided.xlsx', engine='xlsxwriter')
newUndecidedDF.to_excel(writer2, sheet_name='NewUndecided')
writer2.save()
            # print(newUndecidedFileDF.loci[j-1, 'Employee ID'])
# xlsxFile = pd.ExcelFile('newUndecidedFile')
# undecidedDF = pd.read_excel(xlsxFile, 'Undecided Units')
# undecidedDF.to_excel('UndecidedStudents.xlsx')
# # print(undecidedDF)