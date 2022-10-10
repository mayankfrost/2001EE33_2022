import pandas as pd


def octant_identification(mod=5000):
    data = pd.read_excel('input_octant_transition_identify.xlsx')

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
    cnt = {}

    mrows = (total + mod - 1) // mod  # number of rows for each range of count
    buffer = mrows + 2  # to consider the part after counts of octants

    unit = 13  # rows for one transition count

    for i in vals:  # generating the tables and filling the initial values with 1
        cnt[i] = [0] * (buffer + unit * (mrows + 1))
        cnt[i][1] = ''
        for j in range(mrows + 1):
            for k in range(4):
                cnt[i][buffer + j * unit + k] = 'To' if k == 3 and i == 1 else ''
            cnt[i][buffer + j * unit + 4] = str(i) if i < 0 else '+' + str(i)

    prev = 0

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

        octant.append(str(oc) if oc < 0 else '+' + str(oc))

        cnt[oc][0] += 1  # overall octant count
        cnt[oc][2 + i // mod] += 1  # range octant count

        if i > 0:
            cnt[oc][buffer + 4 + abs(prev) * 2 - (prev > 0)] += 1  # overall transition count

            mrange = (i + mod - 1) // mod  # determing which range this current octant belongs to
            cnt[oc][buffer + mrange * unit + 4 + abs(prev) * 2 - (prev > 0)] += 1  # overall transition count

        prev = oc

    for i in vals:
        cnt[i].extend([''] * (total - len(cnt[i])))

    data['Octant'] = octant

    emp_col = ['', 'User Input']  # creating a mostly empty column which will have "User Input" and other tidbits
    emp_col.extend([''] * (total - 2))
    for i in range(mrows + 1):
        emp_col[buffer + i * unit + 5] = 'From'
    data[''] = emp_col

    id = ['Overall Count', 'Mod ' + str(mod)]

    for i in range(mrows):
        id.append(str(i * mod) + '-' + str(min(total - 1, (i + 1) * mod - 1)))
    id.extend([''] * (total - len(id)))

    for i in range(mrows + 1):  # generating the column which will hold the first columns of the tables
        id[buffer + i * unit + 2] = ('Overall' if i == 0 else 'Mod') + ' Transition Count'
        id[buffer + i * unit + 4] = 'Count'
        c = 5
        if i > 0:
            id[buffer + i * unit + 3] = str((i - 1) * mod) + '-' + str(min(total - 1, i * mod - 1))
        for j in vals:
            id[buffer + i * unit + c] = str(j) if j < 0 else '+' + str(j)
            c += 1

    data[' '] = id

    for i in vals:
        data[str(i) if i < 0 else '+' + str(i)] = cnt[i]

    data.to_excel('output_octant_transition_identify.xlsx', index=False)


mod = 5000  # changeable value
if mod <= 0:
    raise ValueError("Mod should have a positive value")

octant_identification(mod)
