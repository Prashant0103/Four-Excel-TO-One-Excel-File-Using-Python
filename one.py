import xlsxwriter
import pymysql

url = 'Example1.xlsx'
workbook = xlsxwriter.Workbook(url)
worksheet = workbook.add_worksheet()
print(worksheet.get_name())

 
bold = workbook.add_format({'bold': True})
money = workbook.add_format({'num_format': '$#,##0'})

worksheet.write('A1', 'ID', bold)
worksheet.write('B1', 'UserName', bold)
worksheet.write('C1', 'Email', bold)
worksheet.write('D1', 'Mobile', bold)
worksheet.write('E1', 'Php Marks', bold)
worksheet.write('F1', 'Python Marks', bold)
worksheet.write('G1', 'Cloud Computing Marks', bold)
worksheet.write('H1', 'Total Marks', bold)
worksheet.write('I1', 'OUT OF', bold)
worksheet.write('J1', 'Percentage(%)', bold)

row = 1
column = 0

conn=pymysql.connect(host='localhost',user='root' , password='root',database='student')
cur=conn.cursor()
cur1 = conn.cursor()

sql = '''SELECT 
    stud_data.id,
    stud_data.username,
    stud_data.email,
    stud_data.mobileno,
    group_concat(student_report.Marks) as subjects_marks,
    SUM(student_report.Marks) as subject_total_marks,
    SUM(stud_subject.Total_Marks) as total_marks
FROM
    student_report
    LEFT JOIN
    student_subject ON student_subject.id = student_report.Student_Subject_Id
    LEFT JOIN
    stud_subject ON stud_subject.id = student_subject.Subject_Id
    LEFT JOIN
    stud_data ON stud_data.id = student_subject.Student_Id
    group by student_subject.student_id '''
cur.execute(sql)

A = cur.fetchall()
for i in A:
    print(i)
    
row = 1
col = 0

 # Iterate over the data and write it out row by row.
for id, uname,email,mobile,pmarks,pymarks,ccmarks in (A):
    worksheet.write(row, col,     id)
    worksheet.write(row, col + 1, uname)
    worksheet.write(row, col + 2, email)
    worksheet.write(row, col + 3, mobile)
    try:
        worksheet.write(row, col + 4, pmarks.split(',')[0])
    except:
        worksheet.write(row, col + 4, '--')
        continue
    try:
        worksheet.write(row, col + 5, pmarks.split(',')[1])
    except:
        worksheet.write(row, col + 5, '--')
        continue
    try:
        worksheet.write(row, col + 6, pmarks.split(',')[2])
    except:
        worksheet.write(row, col + 6, '--')
    worksheet.write(row, col + 7, pymarks)
    worksheet.write(row, col + 8, ccmarks)
    worksheet.write(row, col + 9, '=H2/300*100',)
    row += 1
    

workbook.close()