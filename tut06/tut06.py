import pandas as pd
from datetime import datetime, date, timedelta
import os

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

start_time = datetime.now()

lecture_days = []
lecture_days_set = set()


def generate_lectures(timestamps):
    l = len(timestamps) - 1

    if l < 0:  # No attendance would make all calculated statistics meaningless
        raise Exception('No attendances recorded!')

    start = date(int(timestamps[0][6:10]), int(timestamps[0][3:5]),
                 int(timestamps[0][:2]))  # getting the first day (it is a Thursday)
    end = date(int(timestamps[l][6:10]), int(timestamps[l][3:5]), int(timestamps[l][:2]))

    d = start
    sw = False
    while d <= end:  # generating all Mondays and Thursdays
        day = d.strftime('%d-%m-%Y')
        lecture_days.append(day)
        d += timedelta(days=3 if sw else 4)
        sw = not sw

    global lecture_days_set
    lecture_days_set = set(lecture_days)


def during_class(time):  # checking if this time is valid or not
    if time[11:13] == '14' or time[11:] == '15:00':
        return True
    return False


def send_email():
    print('You will need a gmail account with 2 factor authentication enabled')
    print('After that, you have to generate an "app password" from the security settings.')
    print('You will have to enter this password later when prompted to.')
    print()
    print(
        'You can also use the Zimbra account and password. For that you will have to change the code, and comment and uncomment the respective parts')
    print()

    # getting login credentials from the user

    print('Enter your email id')
    email_from = input()
    print('Enter your password')
    password = input()

    # setting up the server

    smtp_port = 587
    smtp_server = "smtp.gmail.com"
    # smtp_server = "mail.iitp.ac.in"
    # smtp_server = "stud.iitp.ac.in"

    # completing email details

    email_to = "cs3842022@gmail.com"
    subject = "Attendance report for Tut06 from 2001EE33"

    msg = MIMEMultipart()
    msg['From'] = email_from
    msg['To'] = email_to
    msg['Subject'] = subject

    # Encoding file as Base64
    attchment_package = MIMEBase("application", "octet-stream")
    attchment_package.set_payload('output/attendance_report_consolidated.xlsx')
    encoders.encode_base64(attchment_package)
    attchment_package.add_header("Content-Disposition",
                                 'attachment; filename= ' + 'attendance_report_consolidated_from_2001EE33.xlsx')
    msg.attach(attchment_package)

    # Casting msg as string
    text = msg.as_string()

    print("Sending email...")
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(email_from, password)
    print("Connected to server!")

    print()
    print("Sending email to", email_to)
    server.sendmail(email_from, email_to, text)
    print("Email sent successfully to", email_to)
    print()

    server.quit()


def attendance_report():
    if not os.path.isdir('output'):  # Make a new directory 'output' if it does not already exist
        os.mkdir('output')

    students = pd.read_csv('input_registered_students.csv')
    attendance = pd.read_csv('input_attendance.csv')
    generate_lectures(attendance['Timestamp'])

    num_stud = len(students['Roll No'])
    num_att = len(attendance['Attendance'])

    total_lectures = len(lecture_days)
    acc_att = []
    fake_att = []
    abs_att = []
    dup_att = []
    tot_cnt = []

    tot_real = [0] * num_stud  # number of real attendances considering all days for a student
    percentage = []

    ids = {}

    for i in range(num_stud):
        tot_cnt.append({})
        acc_att.append({})
        fake_att.append({})
        dup_att.append({})
        abs_att.append({})

        for day in lecture_days:
            tot_cnt[i][day] = 0  # tot_cnt[i][day]: count of total attendances of the i-th student on the day "day"
            acc_att[i][day] = 0  # acc_att[i][day]: count of real attendances of the i-th student on the day "day"
            fake_att[i][day] = 0  # tot_cnt[i][day]: count of fake attendances of the i-th student on the day "day"
            dup_att[i][day] = 0  # tot_cnt[i][day]: count of duplicate attendances of the i-th student on the day "day"
            abs_att[i][day] = 1  # tot_cnt[i][day]: indication of absence of the i-th student on the day "day"

        ids[students['Roll No'][i]] = i  # mapping each roll number to an integer value which corresponds to its index

    for i in range(num_att):

        if type(attendance['Attendance'][i]) is not str or \
                attendance['Attendance'][i][:8] not in ids.keys():  # continue if the roll number is not proper
            continue

        day = attendance['Timestamp'][i][:10]

        if day not in lecture_days_set:  # ignoring all days which aren't Mondays or Thursdays
            continue

        roll_id = ids[attendance['Attendance'][i][:8]]
        tot_cnt[roll_id][day] += 1

        if during_class(attendance['Timestamp'][i]):  # incrementing actual attendance if timing is valid

            if acc_att[roll_id][day] == 1:  # if the attendance for a certain day and student has already been recorded
                dup_att[roll_id][day] += 1
            else:
                tot_real[roll_id] += 1

            acc_att[roll_id][day] = 1
            abs_att[roll_id][day] = 0  # Absence is equated to 0 whenever a valid attendance is recorded

        else:  # incrementing fake attendance if attendance is out of time
            fake_att[roll_id][day] += 1

    for i in range(num_stud):  # calculating percentage
        percentage.append(round(tot_real[i] * 100 / total_lectures, 2))

    for i in range(num_stud):  # arranging a file for each student

        roll = [students['Roll No'][i]] + [''] * total_lectures
        name = [students['Name'][i]] + [''] * total_lectures
        dates = ['']
        total = ['']
        real = ['']
        dup = ['']
        inv = ['']
        absent = ['']

        for j in lecture_days:  # filling up columns for each day
            dates.append(j)
            total.append(tot_cnt[i][j])
            real.append(acc_att[i][j])
            dup.append(dup_att[i][j])
            inv.append(fake_att[i][j])
            absent.append(abs_att[i][j])

        stud = pd.DataFrame()
        stud['Date'] = dates
        stud['Roll'] = roll
        stud['Name'] = name
        stud['Total Attendance Count'] = total
        stud['Real'] = real
        stud['Duplicate'] = dup
        stud['Invalid'] = inv
        stud['Absent'] = absent

        stud.to_excel('output/' + students['Roll No'][i] + '.xlsx', index=False)

    # filling up respective columns for the consolidated report

    rolls = []
    names = []
    for i in range(num_stud):
        rolls.append(students['Roll No'][i])
        names.append(students['Name'][i])

    consolidated = pd.DataFrame()

    consolidated['Roll'] = rolls
    consolidated['Name'] = names

    for day in lecture_days:
        pa = []
        for i in range(num_stud):
            pa.append('A' if abs_att[i][day] == 1 else 'P')

        consolidated[day] = pa

    consolidated['Actual Lecture Taken'] = [total_lectures] * num_stud
    consolidated['Total Real'] = tot_real
    consolidated['% Attendance'] = percentage

    consolidated.to_excel('output/attendance_report_consolidated.xlsx', index=False)

    print('Do you want to send the report to the destination email - "cs3842022@gmail.com"? (y/n)')
    res = input()
    if res == 'y':
        send_email()


attendance_report()

# This shall be the last lines of the code.
end_time = datetime.now()
print('Duration of Program Execution: {}'.format(end_time - start_time))
