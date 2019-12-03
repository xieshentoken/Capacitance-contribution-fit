import os
import re
import sys
import time
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

        self.excel_path = tuple()
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
            command = self.preview_peak, 
            activebackground='black', activeforeground='white')
        preview_button.pack(side=LEFT, ipadx=1, ipady=5, padx=55, pady=10)
        work_button = Button(fm2, text = 'Work', 
            bd=3, width = 10, height = 1, 
            command = self.start_work, 
            activebackground='black', activeforeground='white')
        work_button.pack(side=LEFT, ipadx=1, ipady=5, padx=55, pady=10)
        save_button = Button(fm2, text = 'Save', 
            bd=3, width = 10, height = 1, 
            command = self.save_as_csv, 
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
        plot_fit_result_button = Button(fm3, text = 'Plot selected', 
            bd=3, width = 10, height = 1, 
            command = self.plot_selected, 
            activebackground='black', activeforeground='white')
        plot_fit_result_button.pack(side=LEFT, ipadx=1, ipady=5, padx=55, pady=10)
        plot_fit_result_button.bind('<Double-2>', self.plot_all_fit_result)   
        plot_bar_button = Button(fm3, text = 'Plot bar', 
            bd=3, width = 10, height = 1, 
            command = self.plot_bar, 
            activebackground='black', activeforeground='white')
        plot_bar_button.pack(side=LEFT, ipadx=1, ipady=5, padx=55, pady=10)       
        plot_avb_button = Button(fm3, text = 'i=av^b', 
            bd=3, width = 10, height = 1, 
            command = self.plot_avb, 
            activebackground='black', activeforeground='white')
        plot_avb_button.pack(side=LEFT, ipadx=1, ipady=5, padx=55, pady=10)       
        

    # 创建menubar
    def init_menu(self):
        # '初始化菜单的方法'
        # 定义菜单条所包含的3个菜单
        menus = ('文件', '编辑', '帮助')
        # 定义菜单数据
        items = (OrderedDict([
                # 每项对应一个菜单项，后面元组第一个元素是菜单图标，
                # 第二个元素是菜单对应的事件处理函数
                ('新建', (None, self.new_project)),
                ('打开', (None, self.open_filename)),
                ('另存为', OrderedDict([('CSV', (None, self.save_as_csv)),
                        ('Excel',(None, self.save_as_excel))])),
                ('-1', (None, None)),
                ('退出', (None, self.master.quit)),
                ]),
            OrderedDict([('数据预览',(None, None)), 
                ('查看峰值',(None, None)),
                ('峰位校正',(None, None)),
                ('-1',(None, None)),
                ('log(i)-log(v)',(None, None)),
                ('数据导出',(None, None)),
                ('-2',(None, None)),
                # 二级菜单
                ('更多', OrderedDict([
                    ('离子导率图',(None, self.plot_Dions)),
                    ('数据导出',(None, None)),
                    ('选择颜色',(None, self.select_color))
                    ]))
                ]),
            OrderedDict([('帮助主题',(None, None)),
                ('-1',(None, None)),
                ('关于', (None, None))]))
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
                print(label)
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
        self.master.filenew_icon = PhotoImage(name='D:/pyfold/重构/app/images/filenew.png')
        self.master.fileopen_icon = PhotoImage(name='D:/pyfold/重构/app/images/fileopen.png')
        self.master.save_icon = PhotoImage(name='D:/pyfold/重构/app/images/save.png')
        self.master.saveas_icon = PhotoImage(name='D:/pyfold/重构/app/images/saveas.png')
        self.master.signout_icon = PhotoImage(name='D:/pyfold/重构/app/images/signout.png')
    # 新建项目
    def new_project(self):
        pass
    # 新建路径
    def new_path(self):
        pass

    def open_filename(self):
        # 调用askopenfile方法获取打开的文件名
        self.excel_path = filedialog.askopenfilename(title='打开多个文件',
            filetypes=[('Excel文件', '*.xlsx'), ('Excel 文件', '*.xls'), ("逗号分隔符文件", "*.csv")], # 只处理的文件类型
            initialdir='C:/Users/wo/Desktop/SeTCoMIPPy600')#'d:/') # 初始目录
        self.excel_adr.set(self.excel_path)

        
    def init_scan_sweep(self):
        self.v1.set(0.1)
        # self.v2.set(0.2)
        # self.v3.set(0.5)
        # self.v4.set(1.0)
        # self.v5.set(2.0)
        # self.v6.set(5.0)
        # self.v7.set(10.0)
        # self.v8.set(0.0)
        # self.v9.set(0.0)

    def processData(self):
        if self.excel_path:
            io = self.excel_path
            self.scan_sweep = [self.v1.get(), self.v2.get(), self.v3.get(), 
                self.v4.get(), self.v5.get(), self.v6.get(), 
                self.v7.get(), self.v8.get(), self.v9.get()]
            example = Orz(io, self.scan_sweep)
            try:
                example.read_data2()
                example.search_peak()
            except xlrd.biffh.XLRDError:
                messagebox.showinfo(title='警告',message='文件读取错误！')
            return example.data_list, example.ox_peak_list, example.red_peak_list
        else:
            messagebox.showinfo(title='警告',message='请选择有效文件！')

    def preview_peak(self):
        prda = self.processData()
        color = ['k','mediumslateblue','dimgrey','blueviolet','forestgreen','orchid','dodgerblue','yellowgreen','teal', 
            'r','mediumseagreen','royalblue','gold','tomato','lightgreen','lightsteelblue','hotpink','darkorchid']
        label = ['v1', 'v2', 'v3', 'v4', 'v5', 'v6','v7' ,'v8', 'v9']
        fig,((ax1,ax2,ax3),(ax4,ax5,ax6),(ax7,ax8,ax9)) = plt.subplots(3,3)
        ax = [ax1,ax2,ax3,ax4,ax5,ax6,ax7,ax8,ax9]
        for i, a in zip(range(0, 9-self.scan_sweep.count(0)), self.scan_sweep):
            if a == 0:
                messagebox.showinfo(title='警告',message='请输入扫描速率。')
            elif ((a != 0)and(prda[1][i].empty))or(a != 0)and(prda[2][i].empty):
                messagebox.showinfo(title='警告',message='寻峰出错，请手动输入。')
            elif (a != 0)and(prda[1][i].empty == False)and(prda[2][i].empty == False):
                ax[i].plot(prda[0][i]['Potential(V)'], prda[0][i]['Current(mA)'],
                color = color[i],
                linestyle = '-',
                label = 'pristine '+ label[i],
                linewidth = 1.5)
                # ax[i].legend(loc='best')
                ax[i].annotate('ox peak1', xy=(prda[1][i].iloc[0,0], prda[1][i].iloc[0,1]), 
                    xytext=(prda[1][i].iloc[0,0]-0.5, prda[1][i].iloc[0,1]),
                    color=color[i+9],size=7,
                    arrowprops=dict(facecolor='k', shrink=0.05, width=1))
                ax[i].annotate('red peak1', xy=(prda[2][i].iloc[0,0], prda[2][i].iloc[0,1]), 
                    xytext=(prda[2][i].iloc[0,0]+0.5, prda[2][i].iloc[0,1]),
                    color=color[i+9],size=7,
                    arrowprops=dict(facecolor='k', shrink=0.05, width=1))
                ax[i].set_title('scan sweep = '+str(a)+'mV/s')
                ax[i].set_xlabel('Potential (V)')
                ax[i].set_ylabel('Current (mA)')
        plt.show()

    def start_work(self):
        pass

    def save_as_csv(self):
        if True:
            save_path = filedialog.asksaveasfilename(title='保存文件', 
            filetypes=[("逗号分隔符文件", "*.csv")], # 只处理的文件类型
            initialdir='d:/')
            rsl.to_csv(save_path+'c-fit.csv')
        else:
            messagebox.showinfo(title='警告',message='结果为空，请重新选择源文件！')

    def save_as_excel(self):
        if len(self.cal_rsl) > 0:
            messagebox.showinfo(title='提醒',message='该功能未完善')
        else:
            messagebox.showinfo(title='警告',message='结果为空，请重新选择源文件！')

    def select_color(self):

        self.rgb = colorchooser.askcolor(parent=self.master, title='选择线条颜色',
            color = 'black')

    def preview(self):
        pass

    def plot_selected(self):
        pass

    def plot_all_fit_result(self):
        pass


    def plot_bar(self):
        pass

    def plot_avb(self):
        pass

    def plot_Dions(self):
        pass
