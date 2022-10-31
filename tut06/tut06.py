import pandas as pd
from datetime import datetime, date, timedelta
import os

start_time = datetime.now()

lecture_days = set()


def generate_lectures(timestamps):
    start = date(2022, 7, 28)
    end = date(2022, 11, 26)

    if len(timestamps) == 0:
        raise Exception('No attendances recorded!')

    days = set()
    for t in timestamps:  # generating all days on which attendance have been marked
        days.add(t[:10])

    d = start
    sw = False
    while d < end:  # generating all Mondays and Thursdays
        day = d.strftime('%d/%m/%Y')
        if day in days:  # checking if this day is present in the attendance record
            lecture_days.add(day)
        d += timedelta(days=3 if sw else 4)
        sw = not sw


def during_class(time):  # checking if this time is valid or not
    if time[:10] in lecture_days and (time[11:13] == '14' or time[11:] == '15:00:00'):
        return True
    return False


def createDF(roll, name, lecture, actual, fake, absent, percent):  # creating a DataFrame to output to CSV
    data = pd.DataFrame()
    data['Roll'] = roll
    data['Name'] = name
    data['total_lecture_taken'] = lecture
    data['attendance_count_actual'] = actual
    data['attendance_count_fake'] = fake
    data['attendance_count_absent'] = absent
    data['Percentage (attendance_count_actual/total_lecture_taken)'] = percent
    return data


def attendance_report():
    if not os.path.isdir('output'):  # Make a new directory 'output' if it does not already exist
        os.mkdir('output')

    students = pd.read_csv('input_registered_students.csv')
    attendance = pd.read_csv('input_attendance.csv')
    generate_lectures(attendance['Timestamp'])

    num_stud = len(students['Roll No'])
    num_att = len(attendance['Attendance'])

    duplicates = pd.DataFrame()

    total_lectures = len(lecture_days)
    acc_att = [0] * num_stud
    fake_att = [0] * num_stud
    abs_att = [0] * num_stud
    percent = [0] * num_stud

    tot_cnt = []
    stamp = []

    ids = {}

    for i in range(num_stud):
        tot_cnt.append({})
        stamp.append({})
        for day in lecture_days:
            tot_cnt[i][day] = 0  # tot_cnt[i][day]: count of valid attendances of the i-th student on the day "day"
        ids[students['Roll No'][i]] = i

    for i in range(num_att):

        if type(attendance['Attendance'][i]) is not str or attendance['Attendance'][i][:8] not in ids.keys():  # continue if the roll number is not proper
            continue

        roll_id = ids[attendance['Attendance'][i][:8]]
        if during_class(attendance['Timestamp'][i]):  # incrementing actual attendance if timing is valid
            acc_att[roll_id] += 1
            tot_cnt[roll_id][attendance['Timestamp'][i][:10]] += 1

            if tot_cnt[roll_id][attendance['Timestamp'][i][:10]] == 1:  # storing the first timestamp
                stamp[roll_id][attendance['Timestamp'][i][:10]] = attendance['Timestamp'][i]

        else:  # incrementing fake attendance if attendance is out of time
            fake_att[roll_id] += 1

    for i in range(num_stud):  # calculating other attendance stats
        abs_att[i] = total_lectures - acc_att[i]
        percent[i] = '{0:.2f}'.format(acc_att[i] * 100 / total_lectures)

    time_dup = []
    roll_dup = []
    name_dup = []
    count_dup = []

    for day in lecture_days:  # calculating duplicate entries
        for i in range(num_stud):
            if tot_cnt[i][day] > 1:
                time_dup.append(stamp[i][day])
                roll_dup.append(students['Roll No'][i])
                name_dup.append(students['Name'][i])
                count_dup.append(tot_cnt[i][day])

    duplicates['Timestamp'] = time_dup
    duplicates['Roll'] = roll_dup
    duplicates['Name'] = name_dup
    duplicates['Total count of attendance on that day'] = count_dup

    duplicates.to_csv('output/attendance_report_duplicate.csv', index=False)

    consolidated = createDF(students['Roll No'], students['Name'], [total_lectures] * num_stud, acc_att, fake_att,
                            abs_att, percent)
    consolidated.to_csv('output/attendance_report_consolidated.csv', index=False)

    for i in range(num_stud):  # making separate file for each student
        individual = createDF([students['Roll No'][i]], [students['Name'][i]], [total_lectures], [acc_att[i]],
                              [fake_att[i]], [abs_att[i]], [percent[i]])
        individual.to_csv('output/' + students['Roll No'][i] + '.csv', index=False)


attendance_report()

# This shall be the last lines of the code.
end_time = datetime.now()
print('Duration of Program Execution: {}'.format(end_time - start_time))
