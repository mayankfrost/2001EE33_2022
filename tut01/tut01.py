import pandas as pd


def octant_identification(mod=5000):
    data = pd.read_csv('octant_input.csv', index_col=0)

    total = data['U'].shape[0]  # total number of coordinates

    u_avg = 0
    v_avg = 0
    w_avg = 0

    for i in data['U']:
        u_avg += i

    for i in data['V']:
        v_avg += i

    for i in data['W']:
        w_avg += i

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

    for i in vals:
        cnt[i] = [0] * (mrows + 2)
        cnt[i][1] = ''

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

        octant.append(oc)
        cnt[oc][0] += 1
        cnt[oc][2 + i // mod] += 1

    for i in vals:
        cnt[i].extend([''] * (total - len(cnt[i])))

    data['Octant'] = octant

    emp_col = ['', 'User Input']  # creating an empty column which will have "User Input" in 2nd row
    emp_col.extend([''] * (total - 2))
    data[''] = emp_col

    id = ['Overall Count', 'Mod ' + str(mod)]

    for i in range(mrows):
        id.append(str(i * mod) + '-' + str(min(total - 1, (i + 1) * mod - 1)))
    id.extend([''] * (total - len(id)))

    data['Octant ID'] = id

    for i in vals:
        data[str(i)] = cnt[i]

    data.to_csv('octant_output.csv')


mod = 5000  # changeable value
octant_identification(mod)
