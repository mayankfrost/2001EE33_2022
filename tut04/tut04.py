import pandas as pd

from datetime import datetime

start_time = datetime.now()


def octant_longest_subsequence_count_with_range():
    try:  # checking if file can be opened
        data = pd.read_excel('input_octant_longest_subsequence_with_range.xlsx')
    except:
        print("Couldn't open mentioned file")

    total = data['U'].shape[0]  # total number of coordinates

    u_avg = 0
    v_avg = 0
    w_avg = 0

    try:
        for i in data['U']:
            u_avg += i
    except TypeError:
        print("Column of U has some data type besides int and float")

    try:
        for i in data['V']:
            v_avg += i
    except TypeError:
        print("Column of V has some data type besides int and float")

    try:
        for i in data['W']:
            w_avg += i
    except TypeError:
        print("Column of W has some data type besides int and float")

    u_avg /= total  # finding out the averages
    v_avg /= total
    w_avg /= total

    ua = [u_avg]
    ua.extend([''] * (total - 1))
    va = [v_avg]
    va.extend([''] * (total - 1))
    wa = [w_avg]
    wa.extend([''] * (total - 1))

    data['U Avg'] = ua
    data['V Avg'] = va
    data['W Avg'] = wa

    uu = []
    vv = []
    ww = []

    for i in data['U']:
        uu.append(i - u_avg)

    for i in data['V']:
        vv.append(i - v_avg)

    for i in data['W']:
        ww.append(i - w_avg)

    data['U\'=U - U avg'] = uu
    data['V\'=V - V avg'] = vv
    data['W\'=W - W avg'] = ww

    octant = []

    vals = [1, -1, 2, -2, 3, -3, 4, -4]  # values for different octants

    sub_len = {}
    sub_cnt = {}
    ranges = {}

    for i in vals:
        sub_cnt[i] = 0  # count of
        sub_len[i] = 0  # length of longest subsequence
        ranges[i] = []

    length = 0  # length of the current subsequence
    prev = 0  # previous octant

    octants = []  # list of octant values

    for i in range(total):

        # determining which octant it belongs to

        oc = 1
        if ww[i] < 0:
            oc = -1

        if uu[i] < 0:
            if vv[i] < 0:
                oc *= 3
            else:
                oc *= 2

        elif vv[i] < 0:
            oc *= 4

        if oc == prev:  # adjusting length of the current running subsequence
            length += 1
        else:
            length = 1

        if sub_len[oc] == length:  # using the current length to update the subsequence length and count
            sub_cnt[oc] += 1
            ranges[oc].append([data['Time'][i - length + 1], data['Time'][i]])

        elif sub_len[oc] < length:
            sub_cnt[oc] = 1
            ranges[oc] = [[data['Time'][i - length + 1], data['Time'][i]]]
            sub_len[oc] = length

        prev = oc
        octants.append(str(oc) if oc < 0 else '+' + str(oc))

    data['Octant'] = octants

    emp_col = [''] * total  # to insert a blank column
    data[''] = emp_col

    oc_col = []
    len_col = []
    cnt_col = []
    for i in vals:
        oc_col.append(str(i) if i < 0 else '+' + str(i))
        len_col.append(sub_len[i])
        cnt_col.append(sub_cnt[i])

    oc_col.extend([''] * (total - len(oc_col)))
    len_col.extend([''] * (total - len(len_col)))
    cnt_col.extend([''] * (total - len(cnt_col)))

    data['Octand ID'] = oc_col
    data['Longest Subsquence Length'] = len_col
    data['Count'] = cnt_col

    data['  '] = emp_col

    oc_col.clear()
    len_col.clear()
    cnt_col.clear()

    for i in vals:
        oc_col.append(str(i) if i < 0 else '+' + str(i))
        oc_col.append('TIme')
        oc_col.extend([''] * len(ranges[i]))

        len_col.append(sub_len[i])
        len_col.append('From')

        cnt_col.append(sub_cnt[i])
        cnt_col.append('To')

        for x, y in ranges[i]:
            len_col.append(x)
            cnt_col.append(y)

    oc_col.extend([''] * (total - len(oc_col)))
    len_col.extend([''] * (total - len(len_col)))
    cnt_col.extend([''] * (total - len(cnt_col)))

    data['Octand ID '] = oc_col
    data['Longest Subsquence Length '] = len_col
    data['Count '] = cnt_col

    data.to_excel('output_octant_longest_subsequence_with_range.xlsx', index=False)


octant_longest_subsequence_count_with_range()

end_time = datetime.now()
print('Duration of Program Execution: {}'.format(end_time - start_time))
