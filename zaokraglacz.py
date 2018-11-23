import pandas as pd
import random
import numpy as np

OUTPATH = "E:/niedzielski/"  # save the .csv's here
nREPL = 50

df = pd.read_csv(OUTPATH + "MTX_2.csv")

sums = df.sum(axis=1).astype(int)

cum_probs = df.copy(deep=True)
cum_probs = cum_probs.div(cum_probs.sum(axis=1), axis=0)

dlugosc = len(df.columns)

for n in range(nREPL):
    realizations = list()
    i = 0
    for col in df.columns[:-1]:
        a = np.random.choice(dlugosc, sums[i], p=cum_probs.loc[i])
        unique, counts = np.unique(a, return_counts=True)
        realization = np.zeros(dlugosc)
        for idx,val in enumerate(unique):
            realization[val] = counts[idx]
        realizations.append(realization)
        i += 1
    r1 = pd.DataFrame(realizations)
    print(n)
    r1.to_csv(OUTPATH+"Rk_{}.csv".format(n))









