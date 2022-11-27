import pandas as pd
import csv
import random
from docxtpl import DocxTemplate

doc = DocxTemplate("ZAS22T.docx")

a_questions = []
b_questions = []
c_questions = []
with open('QP_R01_Py.csv') as file:
    heading = next(file)
    reader = csv.reader(file)
    for row in reader:
        for x in row:
            if x == 'A':
                a_questions.append(row)
            elif x == 'B':
                b_questions.append(row)
            elif x == 'C':
                c_questions.append(row)
            else:
                pass

a_section = []

while (len(a_section) <= 4):
    element = random.randrange((len(a_questions)))
    if element not in a_section:
        a_section.append(element)

for x in a_section:
    print(a_questions[x][1])

b_section = []

while (len(b_section) <= 1):
    element = random.randrange((len(b_questions)))
    if element not in b_section:
        b_section.append(element)

for x in b_section:
    print(b_questions[x][1])

c_section = []

while (len(c_section) <= 0):
    element = random.randrange((len(c_questions)))
    if element not in c_section:
        c_section.append(element)

for x in c_section:
    print(c_questions[x][1])

QP_Code = '01'
Institution_Name = 'Acura Software Solutions'
Examination_Name = 'Python - Base : Mid-term Examination Dec 2022'
Q1 = a_questions[a_section[0]][1]
Q2 = a_questions[a_section[1]][1]
Q3 = a_questions[a_section[2]][1]
Q4 = a_questions[a_section[3]][1]
Q5 = a_questions[a_section[4]][1]
Q6 = b_questions[b_section[0]][1]
Q7 = b_questions[b_section[1]][1]
Q8 = c_questions[c_section[0]][1]

context = {'QP_Code': QP_Code, 'Institution_Name': Institution_Name, 'Examination_Name': Examination_Name, 'Q1': Q1, 'Q2': Q2, 'Q3': Q3, 'Q4': Q4, 'Q5': Q5, 'Q6': Q6, 'Q7': Q7, 'Q8': Q8}

doc.render(context)
doc.save('demo.docx')