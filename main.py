import openpyxl

myfile = open(input(
    "which file do you want to process?(for technical purpose put '\\\\' in place of '\\' in the file path) "))
data = myfile.read()
data = data.split('\n')
mylist = []
mylist1 = []
for dial in data:
    if ':' in dial:
        mylist.append(dial[dial.index(' : ') + 2:])

for dial in mylist:
    if '000' in dial and '_' in dial:
        mylist1.append((dial[dial.index('000'):dial.index('_')], dial[dial.index('_') + 1:]))

wb = openpyxl.Workbook()
name = input('what do you want to save the excel file as? ')
date = input('what is the date of attendance?(DD-MM-YYYY) ')
subject = input('which subject? ')
wb.create_sheet(title='attendance_' + date + '_' + subject)
sheet = wb['attendance_' + date + '_' + subject]
row = 2
column = 'A'
sheet['A1'] = 'PSID'
sheet['B1'] = 'Name'
column1 = 'B'
for a, b in mylist1:
    sheet[column + str(row)] = a
    sheet[column1 + str(row)] = b
    row += 1
del wb['Sheet']
wb.save(filename=name + '.xlsx')
print(f'File {name}.xlsx created')
#C:\\Users\\Ayush\\Documents\\Zoom\\2020-05-04 18.00.39 Abhishek Kushwaha's Personal Meeting Room 5310595525\\meeting_saved_chat.txt
