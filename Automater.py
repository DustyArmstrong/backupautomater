import os
import subprocess
import openpyxl
print('\nBackup Automater 1.0.0\n')
print('Opening Workbook...\n')
from openpyxl import load_workbook

#Takes user input for the number of case files to be zipped

userinput = input('\nInput the number of cases to backup...')
print("\n" + userinput + " cases to be backed up...\n")
number = 1
totalinput = int(userinput) + int(number)

wb = openpyxl.load_workbook('Workbook.xlsx')
sheet = wb.active

#Get the column and row and pass it to variables

case = sheet['A'+str(number)]
password = sheet['B'+str(number)]

#Command should be 7z a -pPASSVARIABLE ZIPPEDFILENAMEVARIABLE ITEMTOZIPVARIABLE - ***needs destination****

backup = '7z a -p' + str(password.value) + " " + str(case.value) + " " + str(case.value)

#Backup loop to run on cases add 1 to number to increment Workbook

while number < int(totalinput):

    print("\nBacking up case "+str(case.value)+'\n')
    os.system(backup)
    number+=1
    case = sheet['A'+str(number)]
    password = sheet['B'+str(number)]
    backup = '7z a -p' + str(password.value) + " " + str(case.value) + " " + str(case.value)

else:
    print("\nBackup Complete    ")
