import pandas as pd
from datetime import datetime

start_time = datetime.now()


def octant_range_names(mod=5000):
    octant_name_id_mapping = {"1": "Internal outward interaction", "-1": "External outward interaction",
                              "2": "External Ejection", "-2": "Internal Ejection", "3": "External inward interaction",
                              "-3": "Internal inward interaction", "4": "Internal sweep", "-4": "External sweep"}

    data = pd.read_excel('octant_input.xlsx')

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
    rank = {}
    scores = {}

    mrows = (total + mod - 1) // mod  # number of rows for each range of count

    for i in vals:
        cnt[i] = [0] * (mrows + 2)
        cnt[i][1] = ''

        rank[i] = [0] * (mrows + 2)
        rank[i][1] = ''

        scores[i] = 0

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
        rank[i].extend([''] * (total - len(rank[i])))

    data['Octant'] = octant

    emp_col = ['', 'User Input']  # creating an empty column which will have "User Input" in 2nd row
    emp_col.extend([''] * (total - 2))
    data[''] = emp_col

    id = ['Overall Count', 'Mod ' + str(mod)]

    for i in range(mrows):
        id.append(str(i * mod) + '-' + str(min(total - 1, (i + 1) * mod - 1)))
    id.extend([''] * (total - len(id)))

    data['Octant ID'] = id

    order = []
    for i in vals:
        order.append([cnt[i][0], i])
    order.sort(reverse=True)

    for i in range(8):
        rank[order[i][1]][0] = i + 1

    r1_id = [order[0][1], '']

    for i in range(mrows):
        order.clear()
        for j in vals:
            order.append([cnt[j][i + 2], j])
        order.sort(reverse=True)

        r1_id.append(order[0][1])
        scores[order[0][1]] += 1
        for j in range(8):
            rank[order[j][1]][i + 2] = j + 1

    r1_on = []
    for i in r1_id:
        r1_on.append(octant_name_id_mapping[str(i)] if i != '' else '')

    r1_id.extend([''] * (total - len(r1_id)))
    r1_on.extend([''] * (total - len(r1_on)))

    row = mrows + 5
    cnt[1][row] = 'Octant ID'
    cnt[-1][row] = 'Octant Name'
    cnt[2][row] = 'Count of Rank 1 Mod Values'

    for i in range(8):
        row = mrows + 6 + i
        cnt[1][row] = vals[i]
        cnt[-1][row] = octant_name_id_mapping[str(vals[i])]
        cnt[2][row] = scores[vals[i]]

    for i in vals:
        data[('+' if i > 0 else '') + str(i)] = cnt[i]
    for i in vals:
        data['Rank of ' + ('+' if i > 0 else '') + str(i)] = rank[i]

    data['Rank1 Octant ID'] = r1_id
    data['Rank1 Octant Name'] = r1_on

    data.to_excel('octant_output_ranking_excel.xlsx', index=False)


mod = 5000
octant_range_names(mod)

# This shall be the last lines of the code.
end_time = datetime.now()
print('Duration of Program Execution: {}'.format(end_time - start_time))
