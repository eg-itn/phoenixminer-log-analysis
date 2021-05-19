import sys
import os
import glob
import openpyxl
from openpyxl.chart import ScatterChart, Reference, BarChart, Series

log_list = glob.iglob('log/**.txt')
power_list = []
speed_list = []


def read_log():
    for log in log_list:
        print(log)
        with open(log, 'r') as f:
            while True:
                line = f.readline()
                if not line:
                    break
                elif line == '\n':
                    continue
                timestamp = line[:23]
                timestamp = timestamp[:4] + '/' + timestamp[5:7] + '/' + timestamp[8:10] + ' ' + timestamp[11:]
                line = f.readline()
                if line[:10] == 'GPUs power':
                    end = line.find(' W')
                    # print(timestamp + ' Power ' + line[12:18])
                    power_list.append(timestamp + ',' + line[12:end])
                elif line[35:50] == 'Effective speed':
                    timestamp = line[:23]
                    timestamp = timestamp[:4] + '/' + timestamp[5:7] + '/' + timestamp[8:10] + ' ' + timestamp[11:]
                    # print(timestamp + ' EffSpped ' + line[52:57])
                    speed_list.append(timestamp + ',,' + line[52:57])

    return power_list, speed_list


def write_csv(outcsv):
    with open(outcsv, 'w') as fw:
        fw.write(',power,speed\n')
        for line in power_list:
            fw.write(line)
            fw.write('\n')
        for line in speed_list:
            fw.write(line)
            fw.write('\n')


def write_excel(power_list, speed_list, outcsv, report):

    import pandas as pd

    df = pd.read_csv(outcsv)
    df.iloc[:,0] = pd.to_datetime(df.iloc[:,0])
    with pd.ExcelWriter(report) as writer:
        df.to_excel(writer, index=False)

    wb = openpyxl.load_workbook(report)
    ws = wb.active

    # for idx in range(2, ws.max_row+1):
    #     ws.cell(row=idx, column=1).number_format = 'yyyy/mm/dd hh:mm:ss.000'

    # グラフのサイズ、位置を指定
    chart = ScatterChart()
    chart.width = 40
    chart.height = 20
    pos_x = 4
    pos_y = 1

    # y軸データ
    values_power = Reference(ws, min_row=1, max_row=len(power_list), min_col=2, max_col=2)
    values_speed = Reference(ws, min_row=len(power_list)+2, max_row=ws.max_row, min_col=3, max_col=3)
    # values = Reference(ws, min_row=1, max_row=ws.max_row, min_col=2, max_col=3)
    chart.legend.legendPos = 'b'

    # x軸データ
    x_axis_power = Reference(ws, min_row=2, max_row=len(power_list), min_col=1, max_col=1)
    x_axis_speed = Reference(ws, min_row=len(power_list)+2, max_row=ws.max_row, min_col=1, max_col=1)
    # x_axis = Reference(ws, min_row=2, max_row=ws.max_row, min_col=1, max_col=1)
    # chart.set_categories(x_axis)

    series_power = Series(values_power, x_axis_power, title_from_data=True)
    chart.series.append(series_power)
    series_speed = Series(values_speed, x_axis_speed, title_from_data=True)
    chart.series.append(series_speed)
    # series = Series(values, x_axis, title_from_data=True)
    # chart.series.append(series)

    ws.add_chart(chart, ws.cell(row=pos_y, column=pos_x).coordinate)

    wb.save(report)

    return


def main():
    power_list, speed_list = read_log()
    outcsv = 'output.csv'
    report = 'output.xlsx'
    write_csv(outcsv)
    write_excel(power_list, speed_list, outcsv, report)


if __name__ == '__main__':
    main()
