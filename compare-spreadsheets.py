'''
================================================================================
    Name:         compare-spreadsheets.py

    Author:       Barra Hart

    Repo:         https://github.com/barrahart/compare-spreadsheets

    Description:  Compared identifying columns in 2 Excel files and puts
                  all missing IDs in a third.
================================================================================
'''

from openpyxl import *

def line():
    print('\n---------------------------------------------------------------------\n')


def compare(projectsA, projectsB, missingProjectsArray, missingProjects):
    #print(f"FIRST ID: {projectsA['A1'].value}")

    for projectA in projectsA.iter_rows(values_only=True):
        flag=False

        for projectB in projectsB.iter_rows(values_only=True):
            if projectA[0]==projectB[0]:
                flag=True

        #print(f"Done with search for ID {projectA[0]}\n")

        if flag==False:
            missingProjectsArray.append(projectA[0])

    for i in range(len(missingProjectsArray)):
        missingProjects[f'A{i+1}'] = missingProjectsArray[i]



print('\n---------------------------------------------------------------------')
print("::::::::::::::::::::: Compare Spreadsheets 1.0 ::::::::::::::::::::::")
print('---------------------------------------------------------------------')
user_continue=input(":::::::::::: Press ENTER to continue, or CTRL-C to exit :::::::::::::\n")

line()
user_work1 = input("\n> Enter full path for Spreadsheet 1: ")
workbook1 = load_workbook(f'{user_work1}')
sheet1=workbook1.worksheets[0]

user_work2 = input("\n> Enter full path for Spreadsheet 2: ")
workbook2 = load_workbook(f'{user_work2}')
sheet2=workbook2.worksheets[0]

workbook3 = load_workbook('./missing.xlsx')
sheet3 = workbook3.worksheets[0]

missing=[]


line()

compare(sheet1, sheet2, missing, sheet3)

'''TEST
line()
print("MISSING PROJECTS:\n\n")
for id in missing:
    print(id)
line()
'''

workbook3.save('./missing.xlsx')

print("\n> Done. Check 'missing.xlsx' for your missing IDs!")
line()
