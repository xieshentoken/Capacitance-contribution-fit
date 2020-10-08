import re
import sys
from collections import OrderedDict
from itertools import permutations
from tkinter import *
from tkinter import colorchooser, filedialog, messagebox, simpledialog, ttk

import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import xlrd

from func_XX import Orz


class App():
    def __init__(self, master):
        self.master = master
        self.initWidgets()

        self.init_scan_sweep()
        self.test_window = 3.
        self.interval = 1
        self.index = -1
        self.fit_data_expand = []
        self.rgb = ('#000000', 'black')
        
    def initWidgets(self):
        # 初始化菜单、工具条用到的图标
        self.init_icons()
        # 调用init_menu初始化菜单
        self.init_menu()
    #----------------------------------------------------------------------------------
    # 创建第一个容器
        fm1 = Frame(self.master)
        # 该容器放在左边排列
        fm1.pack(side=TOP, fill=BOTH, expand=NO)

        ttk.Label(fm1, text='File Path:', font=('StSong', 20, 'bold')).pack(side=LEFT, ipadx=5, ipady=5, padx= 10)
        # 创建字符串变量，用于传递文件地址
        self.excel_adr = StringVar()
        # 创建Entry组件，将其textvariable绑定到self.excel_adr变量
        ttk.Entry(fm1, textvariable=self.excel_adr,
            width=24,
            font=('StSong', 20, 'bold'),
            foreground='#8080c0').pack(side=LEFT, ipadx=5, ipady=5)#fill=BOTH, expand=YES)
        ttk.Button(fm1, text='...',
            command=self.open_filename # 绑定open_filename方法
            ).pack(side=LEFT, ipadx=1, ipady=5)
    #-----------------------------------------------------------------------------------
    # 创建第二个容器
        fm2 = Frame(self.master)
        fm2.pack(side=TOP, fill=BOTH, expand=YES)
        preview_button = Button(fm2, text = 'Preview', 
            bd=3, width = 10, height = 1, 
            command = self.preview_peak_plot, 
            activebackground='black', activeforeground='white')
        preview_button.pack(side=LEFT, ipadx=1, ipady=5, padx=55, pady=10)
        work_button = Button(fm2, text = 'Peak Rectify', 
            bd=3, width = 10, height = 1, 
            command = self.peak_rectify, 
            activebackground='black', activeforeground='white')
        work_button.pack(side=LEFT, ipadx=1, ipady=5, padx=55, pady=10)
        save_button = Button(fm2, text = 'New Path', 
            bd=3, width = 10, height = 1, 
            command = self.new_path, 
            activebackground='black', activeforeground='white')
        save_button.pack(side=LEFT, ipadx=1, ipady=5, padx=55, pady=10)
    #---------------------------------------------------------------------------------------
    # 创建Labelframe容器
        lf = ttk.Labelframe(self.master, text='Scan sweep (mV/s)',
            padding=10)
        lf.pack(side=TOP, fill=BOTH, expand=NO, padx=10, pady=10)
        # 创建lf的第一个子容器
        lf_1 = Frame(lf)
        lf_1.pack(side=TOP, fill=BOTH, expand=NO, padx=10, pady=10)
        books1 = ['v1:', 'v2:', 'v3:']
        self.v1, self.v2, self.v3 = DoubleVar(), DoubleVar(), DoubleVar()
        self.v4, self.v5, self.v6 = DoubleVar(), DoubleVar(), DoubleVar()
        self.v7, self.v8, self.v9 = DoubleVar(), DoubleVar(), DoubleVar()
        dis1 = [self.v1, self.v2, self.v3]
        self.var1, self.var2, self.var3 = IntVar(), IntVar(), IntVar()
        self.var4, self.var5, self.var6 = IntVar(), IntVar(), IntVar()
        self.var7, self.var8, self.var9 = IntVar(), IntVar(), IntVar()
        check1 = [self.var1, self.var2, self.var3]
        # 使用循环创建多个Label和Entry，并放入Labelframe中,用于输入扫描速率信息
        for book, d, vs in zip(books1, dis1, check1):
            Label(lf_1, font=('StSong', 20, 'bold'), text=book).pack(side=LEFT, padx=5, pady=10)
            ttk.Entry(lf_1, textvariable=d,
                width=3,
                font=('StSong', 20, 'bold'),
                foreground='#8080c0').pack(side=LEFT, ipadx=5, ipady=5, padx=15, pady=10)
            Checkbutton(lf_1, variable=vs, onvalue=1, offvalue=0).pack(side=LEFT, padx=1, pady=10)
        # 创建lf的第二个子容器
        lf_2 = Frame(lf)
        lf_2.pack(side=TOP, fill=BOTH, expand=NO, padx=10, pady=10)
        books2 = ['v4:', 'v5:', 'v6:']
        self.v4, self.v5, self.v6 = DoubleVar(), DoubleVar(), DoubleVar()
        dis2 = [self.v4, self.v5, self.v6]
        check2 = [self.var4, self.var5, self.var6]
        # 使用循环创建多个Label和Entry，并放入Labelframe中,用于输入扫描速率信息
        for book, d, vs in zip(books2, dis2, check2):
            Label(lf_2, font=('StSong', 20, 'bold'), text=book).pack(side=LEFT, padx=5, pady=10)
            ttk.Entry(lf_2, textvariable=d,
                width=3,
                font=('StSong', 20, 'bold'),
                foreground='#8080c0').pack(side=LEFT, ipadx=5, ipady=5, padx=15, pady=10)
            Checkbutton(lf_2, variable=vs, onvalue=1, offvalue=0).pack(side=LEFT, padx=1, pady=10)
        # 创建lf的第三个子容器
        lf_3 = Frame(lf)
        lf_3.pack(side=TOP, fill=BOTH, expand=NO, padx=10, pady=10)
        books3 = ['v7:', 'v8:', 'v9:']#, 'v4:', 'v5:', 'v6:', 'v7:', 'v8:', 'v9:']
        self.v7, self.v8, self.v9 = DoubleVar(), DoubleVar(), DoubleVar()
        dis3 = [self.v7, self.v8, self.v9]
        check3 = [self.var7, self.var8, self.var9]
        # 使用循环创建多个Label和Entry，并放入Labelframe中,用于输入扫描速率信息
        for book, d, vs in zip(books3, dis3, check3):
            Label(lf_3, font=('StSong', 20, 'bold'), text=book).pack(side=LEFT, padx=5, pady=10)
            ttk.Entry(lf_3, textvariable=d,
                width=3,
                font=('StSong', 20, 'bold'),
                foreground='#8080c0').pack(side=LEFT, ipadx=5, ipady=5, padx=15, pady=10)
            Checkbutton(lf_3, variable=vs, onvalue=1, offvalue=0).pack(side=LEFT, padx=1, pady=10)
    #--------------------------------------------------------------------------------------------------
    # 创建第四个容器
        fm3 = Frame(self.master)
        fm3.pack(side=TOP, fill=BOTH, expand=YES)
        plot_fit_result_button = Button(fm3, text = 'Capacitance-Diffusion Fit', 
            bd=3, width = 10, height = 1, 
            command = self.capac_diff_fit, 
            activebackground='black', activeforeground='white')
        plot_fit_result_button.pack(side=LEFT, ipadx=1, ipady=5, padx=55, pady=10)
        plot_fit_result_button.bind('<Double-2>', None)#self.plot_all_fit_result 
        plot_bar_button = Button(fm3, text = 'i=av^b', 
            bd=3, width = 10, height = 1, 
            command = self.plot_avb, 
            activebackground='black', activeforeground='white')
        plot_bar_button.pack(side=LEFT, ipadx=1, ipady=5, padx=55, pady=10)       
        plot_avb_button = Button(fm3, text = 'Ip-v^0.5', 
            bd=3, width = 10, height = 1, 
            command = self.plot_Dions, 
            activebackground='black', activeforeground='white')
        plot_avb_button.pack(side=LEFT, ipadx=1, ipady=5, padx=55, pady=10)       
        
    # 创建menubar
    def init_menu(self):
        # '初始化菜单的方法'
        # 定义菜单条所包含的3个菜单
        menus = ('文件', '编辑', '修正', '帮助')
        # 定义菜单数据
        items = (OrderedDict([
                # 每项对应一个菜单项，后面元组第一个元素是菜单图标，
                # 第二个元素是菜单对应的事件处理函数
                ('新建', (None, self.new_project)),
                ('打开', (None, self.open_filename)),
                ('另存为', OrderedDict([('CSV', (None, None)),
                        ('Excel',(None, None))])),
                ('-1', (None, None)),
                ('退出', (None, self.master.quit)),
                ]),
            OrderedDict([('电容/扩散拟合',(None, self.capac_diff_fit)), 
                ('电容占比图',(None, self.capac_diff_bar)),
                ('数据导出', OrderedDict([('CV导出', (None, self.save_Cfit_data)),
                        ('柱状图导出',(None, self.save_CD_bar))])),
                ('-1',(None, None)),
                ('log(i)-log(v)',(None, self.plot_avb)),
                ('数据导出 ',(None, self.save_avb)),
                ('-2',(None, None)),
                ('Ip-v^1/2',(None, self.plot_Dions)),
                ('数据导出  ',(None, self.save_Dions)),
                ('-3',(None, None)),
                # 二级菜单
                ('更多', OrderedDict([
                    ('选择颜色',(None, self.select_color))
                    ]))
                ]),
            OrderedDict([('峰值预览',(None,self.preview_peak_plot)),
                ('峰值校正',(None,self.peak_rectify)),
                ('-1',(None, None)),
                ('电化学窗口大小',(None,self.test_window_set)),
                ('-2',(None, None)),
                ('取点间隔',(None,self.interval_set))
                ]),
            OrderedDict([('帮助主题',(None, self.original_data_preparation)),
                ('-1',(None, None)),
                ('关于', (None, self.show_help))]))
        # 使用Menu创建菜单条
        menubar = Menu(self.master)
        # 为窗口配置菜单条，也就是添加菜单条
        self.master['menu'] = menubar
        # 遍历menus元组
        for i, m_title in enumerate(menus):
            # 创建菜单
            m = Menu(menubar, tearoff=0)
            # 添加菜单
            menubar.add_cascade(label=m_title, menu=m)
            # 将当前正在处理的菜单数据赋值给tm
            tm = items[i]
            # 遍历OrderedDict,默认只遍历它的key
            for label in tm:
                # print(label)
                # 如果value又是OrderedDict，说明是二级菜单
                if isinstance(tm[label], OrderedDict):
                    # 创建子菜单、并添加子菜单
                    sm = Menu(m, tearoff=0)
                    m.add_cascade(label=label, menu=sm)
                    sub_dict = tm[label]
                    # 再次遍历子菜单对应的OrderedDict，默认只遍历它的key
                    for sub_label in sub_dict:
                        if sub_label.startswith('-'):
                            # 添加分隔条
                            sm.add_separator()
                        else:
                            # 添加菜单项
                            sm.add_command(label=sub_label,image=None,
                                command=sub_dict[sub_label][1], compound=LEFT)
                elif label.startswith('-'):
                    # 添加分隔条
                    m.add_separator()
                else:
                    # 添加菜单项
                    m.add_command(label=label,image=None,
                        command=tm[label][1], compound=LEFT)
    # 生成所有需要的图标
    def init_icons(self):
        pass
        # self.master.filenew_icon = PhotoImage(file=r"E:\pydoc\Nacher\E N D\image\filenew.png")
        # self.master.fileopen_icon = PhotoImage(name='E:/pydoc/tkinter/images/fileopen.png')
        # self.master.save_icon = PhotoImage(name='E:/pydoc/tkinter/images/save.png')
        # self.master.saveas_icon = PhotoImage(name='E:/pydoc/tkinter/images/saveas.png')
        # self.master.signout_icon = PhotoImage(name='E:/pydoc/tkinter/images/signout.png')
    # 新建项目
    def new_project(self):
        self.new_path()
        self.init_scan_sweep()
        self.test_window = 3.
        self.interval = 1

    # 新建路径
    def new_path(self):
        self.index = -1
        self.excel_adr.set('')
        self.data_list = []
        self.scan_rate = []
        self.example = ''
        self.kk = ''
        self.fit_data_expand = []
        self.bar_data = pd.DataFrame()
        self.avb_data = pd.DataFrame()
        self.Dions_data = pd.DataFrame()

    def open_filename(self):
        self.excel_path = ''
        # 调用askopenfile方法获取打开的文件名
        self.excel_path = filedialog.askopenfilename(title='打开文件',
            filetypes=[('Excel文件', '*.xlsx'), ('Excel 文件', '*.xls')], # 只处理的文件类型
            initialdir='G:/测试结果\CV-EIS/') # 初始目录
        self.excel_adr.set(self.excel_path)
        self.index = 0

    def init_scan_sweep(self):
        self.v1.set(0.1)
        self.v2.set(0.2)
        self.v3.set(0.4)
        self.v4.set(0.8)
        self.v5.set(1.6)
        self.v6.set(2.0)
        self.v7.set(4.0)
        self.v8.set(8.0)
        self.v9.set(10.)

        self.var1.set(1)
        self.var2.set(1)
        self.var3.set(1)
        self.var4.set(1)
        self.var5.set(1)
        self.var6.set(1)
        self.var7.set(1)
        self.var8.set(0)
        self.var9.set(0)
    
    def get_scan_rate(self):
        self.scan_sweep = [self.v1.get(), self.v2.get(), self.v3.get(), 
            self.v4.get(), self.v5.get(), self.v6.get(), 
            self.v7.get(), self.v8.get(), self.v9.get()]
        self.selected_sweep = [self.var1.get(), self.var2.get(), self.var3.get(), 
            self.var4.get(), self.var5.get(), self.var6.get(), 
            self.var7.get(), self.var8.get(), self.var9.get()]

    def read_data(self):
        self.get_scan_rate()
        self.data_list = []
        self.scan_rate = []
        file = pd.ExcelFile(self.excel_path)
        scan_names = file.sheet_names
        dots_num = int(816.*self.test_window)
        i = 1
        for na, x, yn in zip(scan_names, self.scan_sweep, self.selected_sweep):
            if (x != 0)and(yn != 0):
                self.scan_rate.append(x)
                try:
                    df = pd.read_excel(self.excel_path, sheet_name = na)
                    data = pd.concat([df['WE(1).Potential (V)'], df['WE(1).Current (A)']*1000.], axis=1)
                    data = data.iloc[-dots_num::self.interval]
                    data.index = range(0, int(dots_num/self.interval))
                    data.columns = ['Potential(V)', 'Current(mA)']
                    self.data_list.append(data)
                except KeyError:
                    messagebox.showinfo(title='警告',message='文件读取错误:检查是否包含空Sheet。')
                    continue
                # except xlrd.biffh.XLRDError:
                #     messagebox.showinfo(title='警告',message='文件读取错误:工作表需以“Sheet”加数字命名。')
                except ValueError:
                    messagebox.showinfo(title='警告',message='文件读取错误:请确认电化学窗口。')
                    break
                finally:
                    pass
                i += 1
            elif (x == 0)and(yn != 0):
                messagebox.showinfo(title='警告',message='文件读取错误:检查扫速输入值是否为零。')
                break
            else:
                # 如果该扫速没有被选中，则在list中补充一个0或空的DataFrame
                self.scan_rate.append(0)
                self.data_list.append(pd.DataFrame(columns=('Potential(V)', 'Current(mA)')))
                i += 1
                continue
        
    def processData(self):
        if '.xls' in self.excel_path:
            self.read_data()
            # self.drop0_scan_rate = [i for i in self.scan_rate if i != 0]
            # self.drop0_data_list = [j for j in self.data_list if j.empty == False]
            self.example = Orz(self.scan_rate, self.data_list)
            try:
                self.example.search_peak()
            except:
                messagebox.showinfo(title='警告',message='寻峰出错')
        else:
            messagebox.showinfo(title='警告',message='请选择有效文件！')

    def peak_rectify(self):
        if self.index == 0:
            self.processData()
            self.index += 1
        elif self.index == -1:
            messagebox.showinfo(title='警告',message='请选择源文件！')
        else:
            self.index += 1
        self.kk = RectifyPeak(self.master, self.scan_sweep, self.selected_sweep, self.example.ox_peak_list, self.example.red_peak_list)
        self.example.ox_peak_list = self.kk.corrected_ox_peak
        self.example.red_peak_list = self.kk.corrected_red_peak

    def preview_peak_plot(self):
        if self.index == 0:
            self.processData()
        elif self.index == -1:
            messagebox.showinfo(title='警告',message='请选择源文件！')
        else:
            pass
        color = ['k','mediumslateblue','dimgrey','blueviolet','forestgreen','orchid','dodgerblue','yellowgreen','teal', 
            'r','mediumseagreen','royalblue','gold','tomato','lightgreen','lightsteelblue','hotpink','darkorchid']
        label = ['v1', 'v2', 'v3', 'v4', 'v5', 'v6','v7' ,'v8', 'v9']
        fig,((ax1,ax2,ax3),(ax4,ax5,ax6),(ax7,ax8,ax9)) = plt.subplots(3,3)
        ax = [ax1,ax2,ax3,ax4,ax5,ax6,ax7,ax8,ax9]
        for i, x, yn in zip(range(0, 9),self.scan_sweep, self.selected_sweep):
            if yn != 0:
                try :
                    ax[i].plot(self.data_list[i]['Potential(V)'], self.data_list[i]['Current(mA)'],
                    color = color[i],
                    linestyle = '-',
                    label = 'pristine '+ label[i],
                    linewidth = 1.5)
                    # ax[i].legend(loc='best')
                except IndexError:
                    messagebox.showinfo(title='警告',message='请检查每个扫速下对应文件是否存在！')
                    break
                try:
                    ax[i].annotate('ox peak1', xy=(self.example.ox_peak_list[i].iloc[0,0], self.example.ox_peak_list[i].iloc[0,1]), 
                        xytext=(self.example.ox_peak_list[i].iloc[0,0]-0.5, self.example.ox_peak_list[i].iloc[0,1]),
                        color=color[i+9],size=7,
                        arrowprops=dict(facecolor='k', shrink=0.05, width=1))
                except IndexError:
                    pass
                try:
                    ax[i].annotate('red peak1', xy=(self.example.red_peak_list[i].iloc[0,0], self.example.red_peak_list[i].iloc[0,1]), 
                        xytext=(self.example.red_peak_list[i].iloc[0,0]+0.5, self.example.red_peak_list[i].iloc[0,1]),
                        color=color[i+9],size=7,
                        arrowprops=dict(facecolor='k', shrink=0.05, width=1))
                except IndexError:
                    pass
                ax[i].set_title('scan sweep = '+str(x)+'mV/s')
                ax[i].set_xlabel('Potential (V)')
                ax[i].set_ylabel('Current (mA)')
            else:
                continue
        plt.show()

    def capac_diff_fit(self):
        if self.index == 0:
            self.processData()
        else:
            pass
        self.example.fit()
        self.fit_data_expand = []
        # 将example.fit_data_list扩展，没被选中的扫速下填入空DataFrame
        j = 0
        for yn in self.selected_sweep:
            if yn != 0:
                self.fit_data_expand.append(self.example.fit_data_list[j])
                j += 1
            elif yn == 0:
                self.fit_data_expand.append(pd.DataFrame(columns=('Potential(V)', 'Current(mA)')))
        # 作图------------------------------------------------------------------------------------------------------------
        color = ['k','mediumslateblue','dimgrey','blueviolet','forestgreen','orchid','dodgerblue','yellowgreen','teal', 
            'r','mediumseagreen','royalblue','gold','tomato','lightgreen','lightsteelblue','hotpink','darkorchid']
        label = ['v1', 'v2', 'v3', 'v4', 'v5', 'v6','v7' ,'v8', 'v9']
        fig,((ax1,ax2,ax3),(ax4,ax5,ax6),(ax7,ax8,ax9)) = plt.subplots(3,3)
        ax = [ax1,ax2,ax3,ax4,ax5,ax6,ax7,ax8,ax9]
        for i, x, yn in zip(range(0, 9),self.scan_sweep, self.selected_sweep):
            if yn != 0:
                try :
                    ax[i].plot(self.data_list[i]['Potential(V)'], self.data_list[i]['Current(mA)'],
                    color = color[i],
                    linestyle = '-',
                    label = 'pristine '+ label[i],
                    linewidth = 1.5)
                    ax[i].plot(self.fit_data_expand[i]['Potential(V)'], self.fit_data_expand[i]['Capacitance Current(mA)'],
                    color = color[i+9],
                    linestyle = '-',
                    label = 'capacitance'+ label[i],
                    linewidth = 1.5)
                except IndexError:
                    messagebox.showinfo(title='警告',message='请检查每个扫速下对应文件是否存在！')
                ax[i].legend(loc='best')
                ax[i].set_title('scan sweep = '+str(x)+'mV/s')
                ax[i].set_xlabel('Potential (V)')
                ax[i].set_ylabel('Current (mA)')
            else:
                pass
        plt.show()

    def save_Cfit_data(self):
        if self.fit_data_expand:
            save_path = filedialog.asksaveasfilename(title='保存文件', 
            filetypes=[("office Excel", "*.xls")], # 只处理的文件类型
            initialdir='/Users/hsh/Desktop/')
            # writer = pd.ExcelWriter(save_path) 
            with pd.ExcelWriter(save_path+'.xls') as writer:
                for fit_data, orig_data, v in zip(self.fit_data_expand, self.data_list, self.scan_rate):
                    if v != 0:
                        pd.concat([fit_data, orig_data['Current(mA)']], axis=1).to_excel(writer, sheet_name=str(v))
            # writer.save()
            # writer.close()
        else:
            yon = messagebox.askquestion(title='提示',message='结果为空，是否先进行数据拟合？')
            if yon:
                self.capac_diff_fit()
            else:
                pass

    def capac_diff_bar(self):
        capacitance_list = []
        total_capacity_list = []
        if self.fit_data_expand:
            for fit_data, orig_data, v in zip(self.fit_data_expand, self.data_list, self.scan_rate):
                c_bar, total = 0., 0.
                if v != 0:
                    rsl = pd.concat([fit_data, orig_data['Current(mA)']], axis=1)
                    rsl.sort_index(by='Potential(V)', ascending=True, inplace=True)

                    for i in range(0, len(fit_data)-1):
                            if (rsl.iloc[i,0]-rsl.iloc[i+1,0]) < 0.0004:
                                c_bar += abs((rsl.iloc[i,0]-rsl.iloc[i+1,0])*(rsl.iloc[i,1]-rsl.iloc[i+1,1]))
                                total += abs((rsl.iloc[i,0]-rsl.iloc[i+1,0])*(rsl.iloc[i,2]-rsl.iloc[i+1,2]))
                            elif (rsl.iloc[i,0]-rsl.iloc[i+1,0]) > 0.0004:
                                pass
                            else:
                                messagebox.showinfo(title='提示',message='读取数据出错，尝试设置取点间隔为1')
                                break
                    capacitance_list.append(c_bar)
                    total_capacity_list.append(total)
                elif v == 0:
                    capacitance_list.append(c_bar)
                    total_capacity_list.append(total)
            # ====================================================================================
            sv = [v for v in self.scan_rate if v != 0]
            yy = len(sv)
            vv = np.linspace(0,yy-1,yy)
            capacitance = np.array([c for c in capacitance_list if c != 0])
            total_capacity = np.array([t for t in total_capacity_list if t != 0])
            self.c_ratio = capacitance / total_capacity * 100
            self.d_ratio = 100 - self.c_ratio
            self.bar_data = pd.concat([pd.Series(vv), pd.Series(self.c_ratio), pd.Series(self.d_ratio), pd.Series(np.array(sv))], axis=1)
            self.bar_data.columns = ('sweep when plot', 'Capacitance ratio', 'Diffusion ratio', 'real scan sweep')

            fig,ax = plt.subplots()
            plt.bar(vv, self.c_ratio, color='#6495ED', label='Capacitance')
            plt.bar(vv, self.d_ratio, bottom=self.c_ratio, color='#A9A9A9', label='Diffusion')
            plt.xticks([i for i in vv], sv)
            ax.set_ylabel('Contribution ratio (%)')
            ax.set_xlabel('Sweep rate (mV/s)')
            ax.set_ylim(0, 125)
            ax.legend(loc='best')
            # box = {
            # 'facecolor':'.99',
            # 'edgecolor':'w',
            # 'boxstyle':'round'
            # }
            for i in range(0, yy):
                try:
                    ax.text(vv[i] - 0.5, 102, str(int(100 * self.c_ratio[i]) / 100))#, bbox = box)
                except ValueError:
                    messagebox.askquestion(title='提示',message='拟合结果不可靠，请选择合适数据或扫速。')
                    break
            plt.show()
        else:
            yon = messagebox.askquestion(title='提示',message='还未进行数据拟合，是否拟合？')
            if yon:
                self.capac_diff_fit()
            else:
                pass

    def save_CD_bar(self):
        if self.bar_data.empty == False:
            save_bar_path = filedialog.asksaveasfilename(title='保存文件', 
                filetypes=[("逗号分隔符文件", "*.csv")], # 只处理的文件类型
                initialdir='/Users/hsh/Desktop/')
            self.bar_data.to_csv(save_bar_path + '.csv')
        else:
            messagebox.showinfo(title='警告',message='结果为空！')

    def test_window_set(self):
        self.test_window = simpledialog.askfloat('设置电化学测试窗口大小', '电压上下限差(V): \n(408 points/voltage)',
            initialvalue=self.test_window, minvalue=1, maxvalue=200)

    def interval_set(self):
        # 调用askinteger函数生成一个让用户输入整数的对话框
        self.interval = simpledialog.askinteger('设置取点间隔', '即每n个点取一个,n:',
            initialvalue=self.interval, minvalue=1, maxvalue=200)

    def plot_avb(self):
        if self.index == 0:
            self.processData()
        elif self.index == -1:
            messagebox.showinfo(title='警告',message='请选择源文件！')
        else:
            pass
        try:
            self.example.avb()
        except TypeError:
            messagebox.showinfo(title='警告',message='自动寻峰出错，请手动校正峰值。')
        sv = [v for v in self.scan_rate if v != 0]
        anode_I = [ia.iloc[0, 1] for ia in self.example.ox_peak_list if ia.empty == False]
        cathode_I = [ic.iloc[0, 1] for ic in self.example.red_peak_list if ic.empty == False]
        y_a = self.example.anode_avb[0] * np.log10(sv) + self.example.anode_avb[1]
        y_c = self.example.cathode_avb[0] * np.log10(sv) + self.example.cathode_avb[1]
        # 将数据保存在一个DataFrame中=================================================================================================
        self.avb_data = pd.concat([pd.Series(np.array(sv)), pd.Series(np.log10(sv)), 
            pd.Series(np.log10(np.abs(anode_I))), pd.Series(np.log10(np.abs(cathode_I))), 
            pd.Series(y_a), pd.Series(y_c)], axis=1)
        self.avb_data.columns = ('scan sweep(mV/s)', 'lg(scan sweep)', 'lg(Ox peak)', 'lg(Red peak)', 'lg(fit Ox)', 'lg(fit Red)')
        # ==========================================================================================================================
        fig,ax = plt.subplots()
        plt.plot(np.log10(sv), y_a, color='r', label='anode')
        plt.plot(np.log10(sv), y_c, color='g', label='cathode')
        plt.plot(np.log10(sv), np.log10(np.abs(anode_I)), 'o')
        plt.plot(np.log10(sv), np.log10(np.abs(cathode_I)), '+')

        ax.set_ylabel('log(i)/(mA)')
        ax.set_xlabel('log(v)/(mV/s)')
        ax.legend(loc='best')
        text_x = ax.get_xlim()[1] - ax.get_xlim()[0]
        text_y = ax.get_ylim()[1] - ax.get_ylim()[0]
        ax.text(ax.get_xlim()[0] + text_x  *0.25, ax.get_ylim()[0] + text_y * 0.75, 
            'anode_slop=' + str(int(self.example.anode_avb[0] * 100) / 100))
        ax.text(ax.get_xlim()[0] + text_x * 0.25, ax.get_ylim()[0] + text_y * 0.65, 
            'cathode_slop=' + str(int(self.example.cathode_avb[0] * 100) / 100))
        plt.show()

    def save_avb(self):
        if self.avb_data.empty == False:
            save_avb_path = filedialog.asksaveasfilename(title='保存文件', 
                filetypes=[("逗号分隔符文件", "*.csv")], # 只处理的文件类型
                initialdir='/Users/hsh/Desktop/')
            self.avb_data.to_csv(save_avb_path + '.csv')
        else:
            messagebox.showinfo(title='警告',message='结果为空！')

    def plot_Dions(self):
        if self.index == 0:
            self.processData()
        elif self.index == -1:
            messagebox.showinfo(title='警告',message='请选择源文件！')
        else:
            pass
        try:
            self.example.sqrt_D()
        except TypeError:
            messagebox.showinfo(title='警告',message='自动寻峰出错，请手动校正峰值。')
        sv = [v for v in self.scan_rate if v != 0]
        anode_I = self.example.ox_peak1
        cathode_I = self.example.red_peak1
        y_a = self.example.anode_D_ions[0] * np.sqrt(sv) + self.example.anode_D_ions[1]
        y_c = self.example.cathode_D_ions[0] * np.sqrt(sv) + self.example.cathode_D_ions[1]
        # 将数据保存在一个DataFrame中=================================================================================================
        self.Dions_data = pd.concat([pd.Series(np.array(sv)), pd.Series(np.sqrt(sv)), 
            pd.Series(np.array(anode_I)), pd.Series(np.array(cathode_I)), 
            pd.Series(y_a), pd.Series(y_c)], axis=1)
        self.Dions_data.columns = ('scan sweep(mV/s)', 'sqrt(scan sweep)', 'Ox peak', 'Red peak', 'k*(vD)^0.5 Ox', 'k*(vD)^0.5 Red')
        # ==========================================================================================================================
        fig,ax = plt.subplots()
        plt.plot(np.sqrt(sv), y_a, color='#e4145e', label='anode')
        plt.plot(np.sqrt(sv), y_c, color='#35ca8e', label='cathode')
        plt.plot(np.sqrt(sv), anode_I, 'o')
        plt.plot(np.sqrt(sv), cathode_I, '+')

        ax.set_ylabel('Ip/(mA)')
        ax.set_xlabel('sqrt(v)/(mV/s)^0.5')
        ax.legend(loc='best')
        text_x = ax.get_xlim()[1] - ax.get_xlim()[0]
        text_y = ax.get_ylim()[1] - ax.get_ylim()[0]
        ax.text(ax.get_xlim()[0] + text_x  *0.25, ax.get_ylim()[0] + text_y * 0.75, 
            'anode_k*D^0.5=' + str(int(self.example.anode_D_ions[0] * 100) / 100))
        ax.text(ax.get_xlim()[0] + text_x * 0.25, ax.get_ylim()[0] + text_y * 0.65, 
            'cathode_k*D^0.5=' + str(int(self.example.cathode_D_ions[0] * 100) / 100))
        plt.show()

    def save_Dions(self):
        if self.Dions_data.empty == False:
            save_Dions_path = filedialog.asksaveasfilename(title='保存文件', 
                filetypes=[("逗号分隔符文件", "*.csv")], # 只处理的文件类型
                initialdir='/Users/hsh/Desktop/')
            self.Dions_data.to_csv(save_Dions_path + '.csv')
        else:
            messagebox.showinfo(title='警告',message='结果为空！')

    def original_data_preparation(self):
        messagebox.showinfo(title='原始数据准备',message='根据扫速大小，依次将不同扫速下的数据（最多9个）分别保存在同一个Excel的不同Sheet中即可。')

    def show_help(self):
        messagebox.showinfo(title='关于',message='离子导率由Randles-Sevcik方程给出：\n' +
            'Ip = 0.4463*nFA*(nF/RT)^0.5 *Δc0*(vDions)^0.5\n' + 'n：反应过程参与电子数\n' + 
            'A：电化学活性面积\n' + 'Δc0：反应前后离子浓度的变化量\n' + 'v：扫描速率')

    def select_color(self):
        self.rgb = colorchooser.askcolor(parent=self.master, title='选择线条颜色',
            color = 'black')

    # 自定义对话框类，继承Toplevel------------------------------------------------------------------------------------------
    # 创建弹窗，用于手动校正各扫速下的峰值坐标
class RectifyPeak(Toplevel):
    # 定义构造方法
    def __init__(self, parent, scan_sweep, selected_sweep, ox_peak_list, red_peak_list, title = '峰值校正', modal=False):
        Toplevel.__init__(self, parent)
        self.transient(parent)
        # 设置标题
        if title: self.title(title)
        self.parent = parent
        self.result = None
        self.scan_sweep = scan_sweep
        self.selected_sweep = selected_sweep
        self.ox_peak_list = ox_peak_list
        self.red_peak_list = red_peak_list
        # 创建对话框的主体内容
        frame = Frame(self)
        # 调用init_widgets方法来初始化对话框界面
        self.initial_focus = self.init_widgets(frame)
        frame.pack(padx=5, pady=5)
        # 调用init_buttons方法初始化对话框下方的按钮
        self.init_buttons()
        # 根据modal选项设置是否为模式对话框
        if modal: self.grab_set()
        if not self.initial_focus:
            self.initial_focus = self
        # 为"WM_DELETE_WINDOW"协议使用self.cancel_click事件处理方法
        self.protocol("WM_DELETE_WINDOW", self.cancel_click)
        # 根据父窗口来设置对话框的位置
        self.geometry("+%d+%d" % (parent.winfo_rootx()+50,
            parent.winfo_rooty()+50))
        print( self.initial_focus)
        # 让对话框获取焦点
        self.initial_focus.focus_set()
        self.wait_window(self)
    def init_var(self):
        self.o_Up1, self.o_Up2, self.o_Up3 = DoubleVar(), DoubleVar(), DoubleVar()
        self.o_Up4, self.o_Up5, self.o_Up6 = DoubleVar(), DoubleVar(), DoubleVar()
        self.o_Up7, self.o_Up8, self.o_Up9 = DoubleVar(), DoubleVar(), DoubleVar()
        self.o_Ip1, self.o_Ip2, self.o_Ip3 = DoubleVar(), DoubleVar(), DoubleVar()
        self.o_Ip4, self.o_Ip5, self.o_Ip6 = DoubleVar(), DoubleVar(), DoubleVar()
        self.o_Ip7, self.o_Ip8, self.o_Ip9 = DoubleVar(), DoubleVar(), DoubleVar()
        self.r_Up1, self.r_Up2, self.r_Up3 = DoubleVar(), DoubleVar(), DoubleVar()
        self.r_Up4, self.r_Up5, self.r_Up6 = DoubleVar(), DoubleVar(), DoubleVar()
        self.r_Up7, self.r_Up8, self.r_Up9 = DoubleVar(), DoubleVar(), DoubleVar()
        self.r_Ip1, self.r_Ip2, self.r_Ip3 = DoubleVar(), DoubleVar(), DoubleVar()
        self.r_Ip4, self.r_Ip5, self.r_Ip6 = DoubleVar(), DoubleVar(), DoubleVar()
        self.r_Ip7, self.r_Ip8, self.r_Ip9 = DoubleVar(), DoubleVar(), DoubleVar()
        self.ox_Up = [self.o_Up1, self.o_Up2, self.o_Up3, self.o_Up4, self.o_Up5, self.o_Up6, self.o_Up7, self.o_Up8, self.o_Up9]
        self.ox_Ip = [self.o_Ip1, self.o_Ip2, self.o_Ip3, self.o_Ip4, self.o_Ip5, self.o_Ip6, self.o_Ip7, self.o_Ip8, self.o_Ip9]
        self.red_Up = [self.r_Up1, self.r_Up2, self.r_Up3, self.r_Up4, self.r_Up5, self.r_Up6, self.r_Up7, self.r_Up8, self.r_Up9]
        self.red_Ip = [self.r_Ip1, self.r_Ip2, self.r_Ip3, self.r_Ip4, self.r_Ip5, self.r_Ip6, self.r_Ip7, self.r_Ip8, self.r_Ip9]
        for yn, up, ip, oxPeak in zip(self.selected_sweep, self.ox_Up, self.ox_Ip, self.ox_peak_list):
            if yn != 0:
                try:
                    up.set(float(int(oxPeak.iloc[0,0]*1000)/1000))  # 取三位小数点
                    ip.set(float(int(oxPeak.iloc[0,1]*1000)/1000))
                except IndexError:
                    continue
        for yn, up, ip, redPeak in zip(self.selected_sweep, self.red_Up, self.red_Ip, self.red_peak_list):
            if yn != 0:
                try:
                    up.set(float(int(redPeak.iloc[0,0]*1000)/1000))
                    ip.set(float(int(redPeak.iloc[0,1]*1000)/1000))
                except IndexError:
                    continue
    # 通过该方法来创建自定义对话框的内容
    def init_widgets(self, master):
        # 创建第一个容器
        fm1 = Frame(master)
        # 该容器放在左边排列
        fm1.pack(side=TOP, fill=BOTH, expand=NO)
        Label(fm1, font=('StSong', 10, 'bold'), 
            text='                        ').pack(side=LEFT, ipadx=5, ipady=5, padx=15, pady=10)
        for v, yn in zip(self.scan_sweep, self.selected_sweep):
            if yn != 0:
                Label(fm1, font=('StSong', 10, 'bold'), text=str(v)+' mV/s').pack(side=LEFT, padx=15, pady=10)

        self.init_var()

        fm2 = Frame(master)
        fm2.pack(side=TOP, fill=BOTH, expand=NO)
        Label(fm2, font=('StSong', 10, 'bold'), text=' Ox Peak(Up,Ip)').pack(side=LEFT, ipadx=5, ipady=5, padx=15, pady=10)
        # print(self.ox_peak_list[0].iloc[0,0],type(self.ox_peak_list[0].iloc[0,0]))
        # for yn, up, ip, oxPeak in zip(self.selected_sweep, self.ox_Up, self.ox_Ip, self.ox_peak_list):
        for yn, up, ip in zip(self.selected_sweep, self.ox_Up, self.ox_Ip):
            if yn != 0:
                try:
                    ttk.Entry(fm2, textvariable=up,
                        width=3,
                        font=('StSong', 10, 'bold'),
                        foreground='#8080c0').pack(side=LEFT, ipadx=5, ipady=5, padx=5, pady=10)
                    ttk.Entry(fm2, textvariable=ip,
                        width=3,
                        font=('StSong', 10, 'bold'),
                        foreground='#8080c0').pack(side=LEFT, ipadx=5, ipady=5, padx=5, pady=10)
                    # up.set(oxPeak.iloc[0,0])
                    # ip.set(oxPeak.iloc[0,1])
                except IndexError:
                    continue

        fm3 = Frame(master)
        fm3.pack(side=TOP, fill=BOTH, expand=NO)
        Label(fm3, font=('StSong', 10, 'bold'), text='Red Peak(Up,Ip)').pack(side=LEFT, ipadx=5, ipady=5, padx=15, pady=10)
        # for yn, up, ip, redPeak in zip(self.selected_sweep, self.red_Up, self.red_Ip, self.red_peak_list):
        for yn, up, ip in zip(self.selected_sweep, self.red_Up, self.red_Ip):
            if yn != 0:
                try:
                    ttk.Entry(fm3, textvariable=up,
                        width=3,
                        font=('StSong', 10, 'bold'),
                        foreground='#8080c0').pack(side=LEFT, ipadx=5, ipady=5, padx=5, pady=10)
                    ttk.Entry(fm3, textvariable=ip,
                        width=3,
                        font=('StSong', 10, 'bold'),
                        foreground='#8080c0').pack(side=LEFT, ipadx=5, ipady=5, padx=5, pady=10)
                    # up.set(redPeak.iloc[0,0])
                    # ip.set(redPeak.iloc[0,1])
                except IndexError:
                    continue

    # 通过该方法来创建对话框下方的按钮框
    def init_buttons(self):
        f = Frame(self)
        # 创建"确定"按钮,位置绑定self.ok_click处理方法
        w = Button(f, text="确定", width=10, command=self.ok_click, default=ACTIVE)
        w.pack(side=LEFT, padx=5, pady=5)
        # 创建"确定"按钮,位置绑定self.cancel_click处理方法
        w = Button(f, text="取消", width=10, command=self.cancel_click)
        w.pack(side=LEFT, padx=5, pady=5)
        self.bind("<Return>", self.ok_click)
        self.bind("<Escape>", self.cancel_click)
        f.pack()
    # 该方法可对用户输入的数据进行校验
    def validate(self):
        # 可重写该方法
        return True
    # 该方法可处理用户输入的数据
    def process_input(self):
        self.corrected_ox_peak = []
        self.corrected_red_peak = []
        for yn, up, ip in zip(self.selected_sweep, self.ox_Up, self.ox_Ip):
            if yn != 0:
                corrected = pd.DataFrame([[up.get(), ip.get()]],columns=('Potential(V)', 'Current(mA)'))
                self.corrected_ox_peak.append(corrected)
            else:
                corrected = pd.DataFrame(columns=('Potential(V)', 'Current(mA)'))
                self.corrected_ox_peak.append(corrected)
        for yn, up, ip in zip(self.selected_sweep, self.red_Up, self.red_Ip):
            if yn != 0:
                corrected = pd.DataFrame([[up.get(), ip.get()]],columns=('Potential(V)', 'Current(mA)'))
                self.corrected_red_peak.append(corrected)
            else:
                corrected = pd.DataFrame(columns=('Potential(V)', 'Current(mA)'))
                self.corrected_red_peak.append(corrected)
        # print(self.corrected_ox_peak)
    def ok_click(self, event=None):
        # print('确定')
        # 如果不能通过校验，让用户重新输入
        if not self.validate():
            self.initial_focus.focus_set()
            return
        self.withdraw()
        self.update_idletasks()
        # 获取用户输入数据
        self.process_input()
        # 将焦点返回给父窗口
        self.parent.focus_set()
        # 销毁自己
        self.destroy()
    def cancel_click(self, event=None):
        # print('取消')
        # 将焦点返回给父窗口
        self.parent.focus_set()
        # 销毁自己
        self.destroy()
