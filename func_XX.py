import numpy as np
import pandas as pd
import scipy.signal as signal


class Orz():

    def __init__(self, scan_rate, data_list):
        self.scan_rate = scan_rate
        self.data_list = data_list
        self.ox_peak_list = []
        self.red_peak_list = []
        self.fit_data_list = []
        self.intergral_fit_cap_ratio = 0
        self.drop0_scan_rate = [i for i in self.scan_rate if i != 0]
        self.drop0_data_list = [j for j in self.data_list if j.empty == False]
    # def read_data(self):
    #     scan_num = len(self.scan_rate)-self.scan_rate.count(0)
    #     for i in range(1, scan_num+1):
    #         df = pd.read_excel(self.path, sheet_name = 'Sheet'+str(i))
    #         data = pd.concat([df['WE(1).Potential (V)'], df['WE(1).Current (A)']*1000.], axis=1)
    #         data = data.iloc[-2453::self.interval]
    #         data.columns = ['Potential(V)', 'Current(mA)']
    #         self.data_list.append(data)

    # 分别求氧化、还原电流的极值，取其中绝对值最大的两个为峰值电流
    def search_peak(self):
        for data in self.data_list:
            if data.empty == False:
                max_index = signal.argrelextrema(data['Current(mA)'].values, np.greater)[0]
                min_index = signal.argrelextrema(-1*data['Current(mA)'].values, np.greater)[0]
                self.ox_peak_list.append(data.iloc[max_index].sort_values(by=['Current(mA)'],ascending=False).iloc[:2])
                self.red_peak_list.append(data.iloc[min_index].sort_values(by=['Current(mA)'],ascending=True).iloc[:2])
            elif data.empty:
                self.ox_peak_list.append(pd.DataFrame(columns=('Potential(V)', 'Current(mA)')))
                self.red_peak_list.append(pd.DataFrame(columns=('Potential(V)', 'Current(mA)')))

    # 根据公式: i=a*v^b, 拟合出tuple(b, a)储存在self.anode_avb和self.cathode_avb中
    def avb(self):
        if len(self.drop0_scan_rate) == len(self.drop0_data_list):
            self.log_ox_peak1 = [np.log10(abs(c.iloc[0,1])) for c in self.ox_peak_list if c.empty == False]
            self.anode_avb = np.polyfit(np.log10(self.drop0_scan_rate), self.log_ox_peak1, 1)
            self.log_red_peak1 = [np.log10(abs(c.iloc[0,1])) for c in self.red_peak_list if c.empty == False]
            self.cathode_avb = np.polyfit(np.log10(self.drop0_scan_rate), self.log_red_peak1, 1)

    # 根据Randles-Sevcik方程拟合离子扩散系数--------------------------------------------------
    def sqrt_D(self):
        if len(self.drop0_scan_rate) == len(self.drop0_data_list):
            self.ox_peak1 = [c.iloc[0,1] for c in self.ox_peak_list if c.empty == False]
            self.anode_D_ions = np.polyfit(np.sqrt(self.drop0_scan_rate), self.ox_peak1, 1)
            self.red_peak1 = [c.iloc[0,1] for c in self.red_peak_list if c.empty == False]
            self.cathode_D_ions = np.polyfit(np.sqrt(self.drop0_scan_rate), self.red_peak1, 1)

    # 依据公式 i = k1*v + k2*v^1/2,其中k1*v为赝电容贡献项；k2*v^1/2为扩散控制项，
    # 求出每个扫速下电容贡献的电流大小，以dataframe形式储存在self.fit_data_list中，dataframe的三列数据分别为｜电压｜、｜总电流｜、｜电容性电流｜
    def fit(self):
        if len(self.drop0_scan_rate) == len(self.drop0_data_list):
            k_c_list = []
            i_c = pd.concat([i['Current(mA)'] for i in self.drop0_data_list], axis=1)
            for i in i_c.values:
                k_c = np.polyfit(np.sqrt(self.drop0_scan_rate), i/np.sqrt(self.drop0_scan_rate), 1)[0]
                # k_c = np.polyfit(np.sqrt(np.array(self.drop0_scan_rate) / 1000), i/1000/np.sqrt(np.array(self.drop0_scan_rate) / 1000), 1)[0]
                k_c_list.append(k_c)
            for data, v in zip(self.drop0_data_list, self.drop0_scan_rate):
                fit_data = pd.concat([data['Potential(V)'], pd.Series(np.array(k_c_list) * v)], axis=1)
                fit_data.columns = ('Potential(V)', 'Capacitance Current(mA)')
                # data['Capacitance Current(mA)'] = pd.Series(np.array(k_c_list) * v)
                self.fit_data_list.append(fit_data)

    # 依据公式 Qf = ∫(k1*v + k2*v^1/2)dE/v = ∫k1dE + ∫k2dE*v^(-1/2),其中∫k1dE为赝电容贡献项；∫k2dE*v^(-1/2)为扩散控制项，
    # 求出赝电容贡献及各扫速下的扩散控制贡献，并保存于self.fit2_data中，
    def intergral_fit(self):
        if len(self.drop0_scan_rate) == len(self.drop0_data_list):
            self.Qf_list = []
            # 用切片计算各个扫速下的积分容量，值储存在Qf_list中
            for v, data in zip(self.drop0_scan_rate, self.drop0_data_list):
                Qf = 0
                process_data = data.sort_values(by='Potential(V)', ascending=True).diff().dropna()
                for data_dot in process_data.values:
                    if abs(data_dot[0])<0.001:
                        Qf += abs(data_dot[0]*data_dot[1])
                self.Qf_list.append(Qf/v)
            self.coeff = np.polyfit(1/np.sqrt(self.drop0_scan_rate), np.array(self.Qf_list), 1)
            # 赝电容容量，类型为float
            self.pseudo_capacity = self.coeff[1]
            # 扩散容量，类型为ndarray, 其shape为：(所选扫速个数,)
            self.diffusion_capacity = self.coeff[0]/np.sqrt(self.drop0_scan_rate)
            self.intergral_fit_cap_ratio0 = self.pseudo_capacity/(self.pseudo_capacity+self.diffusion_capacity)*100
            self.intergral_fit_cap_ratio1 = self.pseudo_capacity/np.array(self.Qf_list)*100
