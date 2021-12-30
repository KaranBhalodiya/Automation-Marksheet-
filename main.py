import xlsxwriter

workbook = xlsxwriter.Workbook('Marksheet.xlsx')
worksheet = workbook.add_worksheet()
# The workbook object is then used to add new
# worksheet via the add_worksheet() method.
file1 = open('Names.txt', 'r')
Lines = file1.readlines()
li=[]
for line in Lines:
    li.append(line.strip())
marks=[]
#marks=[8,67,45,67,43,23,67,75,78]  #for testing without giving input
for i in range(len(li)):
    marks.append(int(input(f"Enter Marks of {li[i]}::")))
worksheet.write(0,0,"Roll No.")
worksheet.write(0,1,"Names")
worksheet.write(0,2,"Marks")
worksheet.write(0,3,"Status")
rows=1
format1 = workbook.add_format({'bg_color': '#FFC7CE',
                               'font_color': '#9C0006'})# for make red color cell
for i in range(len(li)):
    worksheet.write(rows, 0, i+1)
    worksheet.write(rows,1,li[i])
    worksheet.write(rows, 2, marks[i])
    if marks[i]>=33:                    #for pass student
        worksheet.write(rows, 3,"Pass")
    else:                               #for fail student
        worksheet.write(rows, 3, "Fail")
        worksheet.conditional_format('D{}'.format(i+2), {'type': 'cell',
                                                 'criteria': '>=',
                                                 'value': 33,
                                                 'format': format1})
    rows+=1


workbook.close()