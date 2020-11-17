import csv
import os
import pandas as pd

# Global Variables
folder_destination = 'C:/Users/kungf/Desktop/Statement/Files/DS'


def main():

    res = pd.DataFrame({'NET_AMT': [], 'PROCESS_DATE': []})
    for i in os.listdir(folder_destination):
        if i >= '190204':
            data = pd.read_excel(folder_destination +
                                 '/' + i, usecols=[2, 13, 18])
            res = pd.concat([res, data])

    resArr = []
    curr = None
    s = 0

    for amt, d in zip(res['NET_AMT'], res['PROCESS_DATE']):
        if d != curr:
            if curr != None:
                resArr.append((curr, s))
            curr = d
            s = 0
        s += amt

    with open('Comp_Statement.csv', 'w', newline='') as f:
        writer = csv.writer(f)
        writer.writerows(resArr)


if __name__ == '__main__':
    main()
