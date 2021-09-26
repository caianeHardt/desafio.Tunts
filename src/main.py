import gspread
from oauth2client.service_account import ServiceAccountCredentials
from math import ceil


def update_cell(sheet, id_row, status, final_grade) -> int:
    sheet.update_cell(id_row, 7, status)
    sheet.update_cell(id_row, 8, ceil(final_grade))
    return id_row+1


scope = ['https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name(
    'src/client_secret.json', scope)
client = gspread.authorize(creds)

print("Reading sheet")
sheet = client.open("Engenharia de Software - Desafio Caiane").sheet1

total_lessons = sheet.cell(2, 1).value.split(': ')[1]
total_lessons = int(total_lessons)

list = sheet.get_values('A4:H27')

print(str(len(list)) + " students found")
print("Total lessons: "+str(total_lessons))

print("Calculating students values")
id_row = 4
for row in list:

    absences = int(row[2])

    perc = (absences * 100) / total_lessons

    status = None
    if perc > 25:
        status = 'Reprovado por Falta'
        id_row = update_cell(sheet, id_row, status, 0)
        continue

    grade1 = int(row[3])
    grade2 = int(row[4])
    grade3 = int(row[5])
    average = (grade1+grade2+grade3) / 10
    average = average / 3

    final_grade = 0
    if average < 5:
        status = 'Reprovado por Nota'
    if 5 <= average < 7:
        status = 'Exame Final'
        final_grade = (average+5)/2
    if average >= 7:
        status = 'Aprovado'

    id_row = update_cell(sheet, id_row, status, final_grade)

print("Values calculated and updated on the sheet")
