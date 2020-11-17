import csv
import os
import pandas as pd

# Global Variables
folder_destination = 'C:/Users/kungf/Desktop/Statement/Files/DS'


def main():

    res = pd.DataFrame({'TRAN_DATE': [], 'NET_AMT': [], 'PROCESS_DATE': []})
    for i in os.listdir(folder_destination):
        if i >= '190204':
            data = pd.read_excel(folder_destination +
                                 '/' + i, usecols=[2, 13, 18])
            res = pd.concat([res, data])

    resArr = []

    for t_date, amt, d in zip(res['TRAN_DATE'], res['NET_AMT'], res['PROCESS_DATE']):
        if t_date == '20-Sep-2020':
            resArr.append((t_date, amt, d))

    with open('20_sep.csv', 'w', newline='') as f:
        writer = csv.writer(f)
        writer.writerows(resArr)


if __name__ == '__main__':
    main()
