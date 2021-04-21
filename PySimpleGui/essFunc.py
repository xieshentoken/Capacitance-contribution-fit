import numpy as np
import pandas as pd
import scipy.signal as signal
import functools
import PySimpleGUI as sg

class essFunc():
    def __init__(self):
        pass
    @staticmethod     # 输入包含电流电压的原始数据，分别求氧化、还原电流的极值，取其中绝对值最大的两个为峰值电流
    def searchpeak(data, current_label='Current(mA)'):
        max_index = signal.argrelextrema(data[current_label].values, np.greater)[0]
        min_index = signal.argrelextrema(-1*data[current_label].values, np.greater)[0]
        ox_peak = data.iloc[max_index].sort_values(by=[current_label],ascending=False).iloc[:2]
        red_peak = data.iloc[min_index].sort_values(by=[current_label],ascending=True).iloc[:2]
        return ox_peak, red_peak
    
    @staticmethod    # 根据公式: i=a*v^b, 拟合出tuple(b, a)储存在avb中
    # peak_list为包含对应于sacn_rate（list类型）各数值下的峰值电流
    def avb(peak_list, scan_rate, current_label='Current(mA)'):
        lg_Ipeak = [np.log10(abs(peak_data.iloc[0,list(peak_data.columns).index(current_label)])) for peak_data in peak_list]
        avb = np.polyfit(np.log10(scan_rate), lg_Ipeak, 1)
        return avb
    
    @staticmethod    # 根据Randles-Sevcik方程拟合离子扩散系数
    def R_Sfit(peak_list, scan_rate, current_label='Current(mA)'):
        peakI = [peak_data.iloc[0,list(peak_data.columns).index(current_label)] for peak_data in peak_list]
        sqrtD_k = np.polyfit(np.sqrt(scan_rate), peakI, 1)
        return sqrtD_k, peakI

    @staticmethod   # 求出每个扫速下电容贡献的电流大小，以dataframe形式储存在fit_data_list中，dataframe的三列数据分别为｜电压｜、｜总电流｜、｜电容性电流｜
    def cdFit(data_list, scan_rate, current_label='Current(mA)', potential_label='Potential(V)'):
        k_c_list = []
        fit_data_list = []
        i_c = pd.concat([i[current_label] for i in data_list], axis=1)
        for i in i_c.values:
            k_c = np.polyfit(np.sqrt(scan_rate), i/np.sqrt(scan_rate), 1)[0]
            k_c_list.append(k_c)
        for data, v in zip(data_list, scan_rate):
            fit_data = pd.concat([data[potential_label], pd.Series(np.array(k_c_list) * v)], axis=1)
            fit_data.columns = ('Potential(V)', 'Capacitance Current(mA)')
            fit_data_list.append(fit_data)
        return fit_data_list

    @staticmethod
    # 依据公式 Qf = ∫(k1*v + k2*v^1/2)dE/v = ∫k1dE + ∫k2dE*v^(-1/2),其中∫k1dE为赝电容贡献项；∫k2dE*v^(-1/2)为扩散控制项，
    # 求出赝电容贡献及各扫速下的扩散控制贡献，并保存于fit2_data中，
    def integCdFit(data_list, scan_rate, current_label='Current(mA)', potential_label='Potential(V)'):
        Qf_list = []
        # 用矩形近似计算各个扫速下的积分容量，值储存在Qf_list中
        for v, data in zip(scan_rate, data_list):
            Qf = 0
            process_data = data.sort_values(by=potential_label, ascending=True).diff().dropna()
            for data_dot in process_data.values:
                if abs(data_dot[0])<0.001:
                    Qf += abs(data_dot[0]*data_dot[1])
            Qf_list.append(Qf/v)
        coeff = np.polyfit(1/np.sqrt(scan_rate), np.array(Qf_list), 1)
        # 赝电容容量，类型为float
        pseudo_capacity = coeff[1]
        # 扩散容量，类型为ndarray, 其shape为：(所选扫速个数,)
        diffusion_capacity = coeff[0]/np.sqrt(scan_rate)
        intergral_fit_cap_ratio0 = pseudo_capacity/(pseudo_capacity+diffusion_capacity)*100
        intergral_fit_cap_ratio1 = pseudo_capacity/np.array(Qf_list)*100
        return intergral_fit_cap_ratio0, intergral_fit_cap_ratio1, Qf_list, pseudo_capacity+diffusion_capacity, coeff

def progress_decorator(length):
    def inner_decorator(func):
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            progress_layout = [[sg.Text('Searching... ...')],
                [sg.ProgressBar(length, orientation='h', size=(20, 20), key='-PROGRESSBAR-')], [sg.Cancel()]]
            progwin = sg.Window('Custom Progress Meter', progress_layout)
            progress_bar = progwin['-PROGRESSBAR-']
            for i in range(length):
                func(*args, **kwargs)
                progwin.read(timeout=10)
                progress_bar.UpdateBar(i + 1)
            progwin.close()
        return wrapper
    return inner_decorator

# 输入十六进制颜色值string，输出一个包含(h,s,l)的list
def toHSL(rgb):
    if '#' in rgb: rgb = rgb[1:]
    r = int(rgb[:2], 16)/255
    g = int(rgb[2:4], 16)/255
    b = int(rgb[4:], 16)/255
    maxcolor = max(r, g, b)
    mincolor = min(r, g, b)

    l = (maxcolor+mincolor)/2
    
    if (l == 0)or(maxcolor == mincolor):
        s = 0
    elif (l > 0)and(l <= 0.5):
        s = (maxcolor-mincolor)/(maxcolor+mincolor)
    elif l > 0.5:
        s = (maxcolor-mincolor)/(2-maxcolor-mincolor)

    if maxcolor == mincolor:
        h = 0
    elif (maxcolor == r)and(g >= b):
        h = 60*(g-b)/(maxcolor-mincolor)
    elif (maxcolor == r)and(g < b):
        h = 60*(g-b)/(maxcolor-mincolor)+360
    elif maxcolor == g:
        h = 60*(b-r)/(maxcolor-mincolor)+120
    elif maxcolor == b:
        h = 60*(r-g)/(maxcolor-mincolor)+240
    
    return [round(h,3), round(s,3), round(l,3)]
    
# 输入一个[h,s,l]的list其值为float，输出十六进制的rgb值
def toRGB(hsl):
    h = hsl[0]/360
    s = hsl[1]
    l = hsl[2]
    r = g = b = 0
    rgbl = []
    if s == 0:
        r = g = b = round(l*255)
        rgb = '#' + str(hex(r))[2:]+ str(hex(g))[2:]+ str(hex(b))[2:]
        return rgb
    elif l < 0.5:
        q = l*(1+s)
    elif l >= 0.5:
        q = l+s-l*s
    p = 2*l-q
    t_r = h+1/3
    t_g = h
    t_b = h-1/3
    tt = []
    for t in [t_r, t_g, t_b]:
        if t < 0:
            t = t+1
        elif t > 1:
            t = t-1
        tt.append(t)
    for t, c in zip(tt,[r, g, b]):
        if t < 1/6:
            c = p+((q-p)*6*t)
        elif (t >= 1/6)and(t < 0.5):
            c = q
        elif (t >= 0.5)and(t < 2/3):
            c = p+((q-p)*6*(2/3-t))
        else:
            c = p
        rgbl.append(round(c*255))
    rgb = '#' + str(hex(rgbl[0]))[2:]+ str(hex(rgbl[1]))[2:]+ str(hex(rgbl[2]))[2:]
    return rgb

# 在色环上隔80˚取色,仅改变色相，不改变明度和饱和度
def loop_pick_color(color, i):
    hsl = toHSL(color)
    loop_h = (hsl[0] + 81*i)%360
    loop_s = hsl[1]
    picked_color = toRGB([loop_h, loop_s, hsl[2]])
    return picked_color


