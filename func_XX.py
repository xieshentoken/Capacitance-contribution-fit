import os
from collections import OrderedDict
from itertools import permutations

import numpy as np
import pandas as pd
import scipy.signal as signal


class Orz():

    def __init__(self, path, scan_rate):
        self.path = path  # 为一个包含excel位置信息的list或tup
        self.scan_rate = scan_rate

        self.data_list = []
        self.ox_peak_list = []
        self.red_peak_list = []
        self.fit_data_list = []

    # 读取多张excel中的数据
    # def read_data(self):
    #     for doc_path in self.path:
    #         df = pd.read_excel(doc_path)
    #         data = pd.concat([df['WE(1).Potential (V)'], df['WE(1).Current (A)']*1000.], axis=1)
    #         data = data.iloc[-2453:]
    #         data.columns = ['Potential(V)', 'Current(mA)']
    #         self.data_list.append(data)

    def read_data2(self):
        scan_num = len(self.scan_rate)-self.scan_rate.count(0)
        for i in range(1, scan_num+1):
            df = pd.read_excel(self.path, sheet_name = 'Sheet'+str(i))
            data = pd.concat([df['WE(1).Potential (V)'], df['WE(1).Current (A)']*1000.], axis=1)
            data = data.iloc[-2453:]
            data.columns = ['Potential(V)', 'Current(mA)']
            self.data_list.append(data)

    # 分别求氧化、还原电流的极值，取其中绝对值最大的两个为峰值电流
    def search_peak(self):
        for data in self.data_list:
            max_index = signal.argrelextrema(data['Current(mA)'].values, np.greater)[0]
            min_index = signal.argrelextrema(-1*data['Current(mA)'].values, np.greater)[0]
            self.ox_peak_list.append(data.iloc[max_index].sort_values(by=['Current(mA)'],ascending=False).iloc[:2])
            self.red_peak_list.append(data.iloc[min_index].sort_values(by=['Current(mA)'],ascending=True).iloc[:2])

    # 根据公式: i=a*v^b, 拟合出tuple(b, a)储存在self.anode_avb和self.cathode_avb中
    def avb(self):
        if len(self.scan_rate) == len(self.data_list):
            if 0 not in [len(c) for c in self.ox_peak_list]:
                if 0 not in [len(c) for c in self.red_peak_list]:
                    self.log_ox_peak1 = [np.log10(abs(c.iloc[0,1])) for c in self.ox_peak_list]
                    self.log_red_peak1 = [np.log10(abs(c.iloc[0,1])) for c in self.red_peak_list]
                    self.anode_avb = np.polyfit(np.log10(self.scan_rate), self.log_ox_peak1, 1)
                    self.cathode_avb = np.polyfit(np.log10(self.scan_rate), self.log_red_peak1, 1)

    # 求出每个扫速下电容贡献的电流大小，以dataframe形式储存在self.fit_data_list中，dataframe的三列数据分别为｜电压｜、｜总电流｜、｜电容性电流｜
    def fit(self):
        if len(self.scan_rate) == len(self.data_list):
            k_c_list = []
            i_c = pd.concat([i['Current(mA)'] for i in self.data_list], axis=1)
            for i in i_c.values:
                k_c = np.polyfit(np.sqrt(self.scan_rate), i/np.sqrt(self.scan_rate), 1)[0]
                k_c_list.append(k_c)
            for data, v in zip(self.data_list, self.scan_rate):
                data['Capacitance Current(mA)'] = pd.Series(np.array(k_c_list) * v)
                self.fit_data_list.append(data)
