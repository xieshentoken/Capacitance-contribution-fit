import numpy as np
import pandas as pd
# MacOS下使用这两行，Win下可以注释掉————————————————————
import matplotlib
matplotlib.use("TkAgg")
#——————————————————————————————————
import matplotlib.pyplot as plt
import PySimpleGUI as sg
import webbrowser

import essFunc as ef
import config as cf

class makeGUI():
    init_color = cf.init_color
    current_label = cf.current_label
    potential_label = cf.potential_label
    def __init__(self, app_title):
        self.app_title = app_title
        self.workspace = 'C:/Users/Administrator/Desktop'
        self.win1_active = False
        self.pvidx = [False, False]
        self.dotnum = 2448
        self.interval = 1
        self.scan_rate = [0.1,0.2,0.4,0.8,1.6,2.,4,]
        self.data_list = []
        self.idx = False
        self.fit_data = []
        self.integ_denom = 0
# ------------------- Create the window -------------------
    def make_window(self, theme=cf.theme):
        if theme:
            sg.theme(theme)
        # ------ Menu Definition ------ #
        menu_def = [['&File', ['New','Open file','Save all', '---', 'Set workspace', '---', '&Exit']],
                    ['&Work', ['C-D fit', 'C-D contriBar', 'Save as', ['CV curve', 'Bar'], '---', 'integral fit', 'Save as ', '---', 'logi-logv', 'Save as  ', '---', 'Randles-Sevcik fit', 'Save as   '], ], 
                    ['&Reactify', ['Peaks preview', 'Peaks rectify', 'Export peaks', 'Load peaks', '---', 'select  integQ', '---', 'Number of data dots', 'Interval']], 
                    ['&Help', ['tutorial', '---', 'Change Theme', 'Feedbacks']],]
        # ------ Column Definition ------ #
        column0 = [[sg.InputCombo(cf.scan_set, font=('StSong', 15, 'bold'), size=(30, 1), key='-VLIST-', default_value=cf.scan_set[0]), sg.OK()],
            [sg.T('v1:', font=('StSong', 25, 'bold')), sg.Input(0.1,font=('StSong', 25, 'bold'), size=(3,1), key='-V1-'), sg.Checkbox('',key='-V1CHECK-', pad=(10,1), default=True),
             sg.T('v2:', font=('StSong', 25, 'bold')), sg.Input(0.2,font=('StSong', 25, 'bold'), size=(3,1), key='-V2-'), sg.Checkbox('',key='-V2CHECK-', pad=(10,1), default=True),
             sg.T('v3:', font=('StSong', 25, 'bold')), sg.Input(0.4,font=('StSong', 25, 'bold'), size=(3,1), key='-V3-'), sg.Checkbox('',key='-V3CHECK-', pad=(10,1), default=True)], 
            [sg.T('v4:', font=('StSong', 25, 'bold')), sg.Input(0.8,font=('StSong', 25, 'bold'), size=(3,1), key='-V4-'), sg.Checkbox('',key='-V4CHECK-', pad=(10,1), default=True),
             sg.T('v5:', font=('StSong', 25, 'bold')), sg.Input(1.6,font=('StSong', 25, 'bold'), size=(3,1), key='-V5-'), sg.Checkbox('',key='-V5CHECK-', pad=(10,1), default=True),
             sg.T('v6:', font=('StSong', 25, 'bold')), sg.Input(2.0,font=('StSong', 25, 'bold'), size=(3,1), key='-V6-'), sg.Checkbox('',key='-V6CHECK-', pad=(10,1), default=True)], 
            [sg.T('v7:', font=('StSong', 25, 'bold')), sg.Input(4.0,font=('StSong', 25, 'bold'), size=(3,1), key='-V7-'), sg.Checkbox('',key='-V7CHECK-', pad=(10,1), default=True),
             sg.T('v8:', font=('StSong', 25, 'bold')), sg.Input(0,font=('StSong', 25, 'bold'), size=(3,1), key='-V8-'), sg.Checkbox('',key='-V8CHECK-', pad=(10,1), default=False),
             sg.T('v9:', font=('StSong', 25, 'bold')), sg.Input(0,font=('StSong', 25, 'bold'), size=(3,1), key='-V9-'), sg.Checkbox('',key='-V9CHECK-', pad=(10,1), default=False)], ]

        layout = [[sg.Menu(menu_def, tearoff=False)],
            [sg.Text('File Path:', pad=(10, 1), justification='center', font=("Helvetica", 20, 'bold')), 
            sg.Input(font=('StSong', 25, 'bold'), size=(20,1), key='-PATH-'), 
            sg.FileBrowse(key='-SELECTPATH-', font=('courier', 25, 'italic'), initial_folder=self.workspace, 
                        file_types=(("Microsoft Excel", "*.xlsx"), ("Microsoft Excel2003", "*.xls"),), 
                        tooltip=' Select an Excel contain your all test datas.\n right click to load a peaks data ',
                        button_color=(sg.theme_background_color(), sg.theme_background_color()))],
            [sg.Button('PeakView', key='-PEAKVIEW-', pad=(10,1), border_width=0, font=('courier', 25, 'italic')), 
            sg.Button('Peaks Rectify', key='-PEAKRECT-', pad=(10,1), border_width=0,
                tooltip=' right click to save the peaks data ', font=('courier', 25, 'italic')), 
            sg.Button('Clear', key='-CLEAR-', pad=(10,1), size=(7,1),border_width=0, font=('courier', 24, 'italic'))],
            [sg.Frame(title='Scan Rate (mV/s)', layout=column0)],
            [sg.Button('C-D Fit', key='-CDFIT-', pad=(10,1), border_width=0,size=(7,1),font=('courier', 25, 'italic'),
            tooltip=' right click to show the contribution ratio '), 
            sg.Button('avb', key='-AVB-', pad=(45,1), border_width=0,size=(7,1), font=('courier', 25, 'italic'),
            tooltip=' right click to save data as csv '), 
            sg.Button('R-S Fit', key='-DFIT-', pad=(10,1), border_width=0,size=(7,1), font=('courier', 25, 'italic'), 
            tooltip=' right click to save data as csv ')],]

        return sg.Window(self.app_title, layout, finalize=True, grab_anywhere =False)

    def event_bind(self):
        self.master.bind("<Double-Button-3>", '-CHANGETHEME-')
        self.master['-SELECTPATH-'].bind("<Button-2>", 'PEAKLOAD-')
        self.master['-SELECTPATH-'].bind("<Button-3>", 'PEAKLOAD-')
        self.master['-PEAKRECT-'].bind("<Button-2>", 'PEAKSAVE-')
        self.master['-PEAKRECT-'].bind("<Button-3>", 'PEAKSAVE-')
        self.master['-CDFIT-'].bind("<Button-2>", 'BAR-')
        self.master['-CDFIT-'].bind("<Button-3>", 'BAR-')
        self.master['-CDFIT-'].bind("<Double-Button-1>", 'CVSAVE-')
        self.master['-CDFIT-'].bind("<Double-Button-2>", 'BARSAVE-')
        self.master['-CDFIT-'].bind("<Double-Button-3>", 'BARSAVE-')
        self.master['-AVB-'].bind("<Button-2>", 'SAVE-')
        self.master['-AVB-'].bind("<Button-3>", 'SAVE-')
        self.master['-DFIT-'].bind("<Button-2>", 'SAVE-')
        self.master['-DFIT-'].bind("<Button-3>", 'SAVE-')

# ------------------- Main Program and Event Loop -------------------
    def main_loop(self):
        self.master = self.make_window()
        self.event_bind()
        while True:
            event, values = self.master.read(timeout=100)
            if event != '__TIMEOUT__':
                print(event, values)
            if event == sg.WINDOW_CLOSED:
                break
            if event in ('Change Theme', '-CHANGETHEME-'):      # Theme button clicked, so get new theme and restart window
                event, values = sg.Window('Choose Theme', [[sg.Combo(sg.theme_list(), readonly=True, k='-THEME LIST-'), sg.OK(), sg.Cancel()]]).read(close=True)
                if event == 'OK':
                    self.master.close()
                    self.master = self.make_window(values['-THEME LIST-'])
                    self.master.read(timeout=100)
                    self.event_bind()
            if not self.win1_active and (event == 'Peaks rectify' or event == '-PEAKRECT-'):
                self.win1_active = True
                self.win1 = sg.Window('Peaks rectify', self.peakrectify(event, values), finalize=True)
            if self.win1_active:
                ev1, vals1 = self.win1.read(timeout=100)
                if ev1 in (None, sg.WIN_CLOSED):
                    self.win1_active  = False
                    self.win1.close()
                if ev1 == 'OK':
                    for sv, odata, rdata in zip(self.selected_idx, self.filt_ox_peak_list, self.filt_red_peak_list):
                        v = int(sv) + 1
                        odata.iloc[0,0] = round(float(vals1[f'-UO{v}-']),3)
                        odata.iloc[0,1] = round(float(vals1[f'-IO{v}-']),3)
                        rdata.iloc[0,0] = round(float(vals1[f'-UR{v}-']),3)
                        rdata.iloc[0,1] = round(float(vals1[f'-IR{v}-']),3)
                    self.win1_active  = False
                    self.pvidx[0] = True
                    self.win1.close()
            if event in ('New', '-CLEAR-'):
                self.new_project()
            elif event == 'Set workspace':
                self.workspace = sg.popup_get_folder('Select a folder', initial_folder=self.workspace, no_window=True,)
                self.master.close()
                self.master = self.make_window()
                self.event_bind()
            elif event == 'Number of data dots':
                self.setDotsnum()
            elif event == 'Interval':
                self.setInterval()
            elif event == 'select  integQ':
                self.selectIntegQ()
            elif event == 'C-D fit' or event == '-CDFIT-':
                self.cdfit(event, values)
            elif event in ('CV curve', '-CDFIT-CVSAVE-'):
                self.save_Cfit_data(event, values)
            elif event in ('C-D contriBar', '-CDFIT-BAR-'):
                self.cdbar(event, values)
            elif event in ('Bar', '-CDFIT-BARSAVE-'):
                self.save_CD_bar(event, values)
            elif event == 'integral fit':
                self.integral_fit(event, values)
            elif event == 'Save as ':
                self.save_integfit(event, values)
            elif event == 'Peaks preview' or event == '-PEAKVIEW-':
                self.preview_peak_plot(event, values)
            elif event in ('Export peaks', '-PEAKRECT-PEAKSAVE-'):
                self.savepeaks(event, values)
            elif event in ('Load peaks', '-SELECTPATH-PEAKLOAD-'):
                self.loadpeaks(event, values)
            elif event in ('logi-logv', '-AVB-'):
                self.showAvb(event, values)
            elif event in ('Save as  ', '-AVB-SAVE-'):
                self.saveAvb(event, values)
            elif event in ('Randles-Sevcik fit', '-DFIT-'):
                self.showDions(event, values)
            elif event in ('Save as   ', '-DFIT-SAVE-'):
                self.saveDions(event, values)
            elif (event in [f'-V{i}CHECK-' for i in range(1,10)]) or (event in [f'-V{j}-' for j in range(1,10)]):
                print(event, values)
            elif event == 'OK':
                try:
                    if len(values['-VLIST-'].split(',')) <= 9:
                        for j in range(1,10):
                            self.master[f'-V{j}-'].update(0)
                            self.master[f'-V{j}CHECK-'].update(False)
                        for i, sv in enumerate(values['-VLIST-'].split(',')):
                            i = i + 1
                            self.master[f'-V{i}-'].update(float(sv))
                            self.master[f'-V{i}CHECK-'].update(True)
                    else:
                        sg.popup('设置出错！扫速不能超过9个,请重新输入。')
                        # raise Exception('Key Error')
                        pass
                except ValueError:
                    sg.popup('扫速设置最多9个，扫速间用英文输入法逗号隔开')
                except KeyError:
                    pass
            elif event == 'Feedbacks':
                sg.popup('Feedback','本项目地址：https://github.com/xieshentoken/\n欢迎提出建议共同完善。\n\t\t\tEmail：dhu-hsa@mail.dhu.edu.cn')
            elif event == 'tutorial':
                webbrowser.open_new_tab('file:///E:/pydoc/PySimpleGUI/cdFit of SG/exe/tutorial.html')

        self.master.close()

    def new_project(self):
        self.win1_active = False
        self.pvidx = [False, False]
        self.dotnum = 2448
        self.interval = 1
        self.scan_rate = [0.1,0.2,0.4,0.8,1.6,2.,4,]
        self.data_list = []
        self.idx = False
        self.fit_data = []
        self.integ_denom = 0

    def setDotsnum(self):
        self.dotnum = int(sg.popup_get_text('设置取点数', default_text=self.dotnum, modal=False,))
    def setInterval(self):
        self.interval = int(sg.popup_get_text('设置取点间隔数', default_text=self.interval, modal=False,))
    # 手动选择电压电流数据列
    def check_UIlabel(self, col):
        checklayout = [[sg.Text('Select Potential Columns', size=(25, 1), justification='center', font=('courier', 10, 'italic')), 
        sg.InputCombo(col, default_value=makeGUI.potential_label, size=(20, 1), key='-ULABEL-')],
            [sg.Text('Select Current Columns', size=(25, 1), justification='center', font=('courier', 10, 'italic')), 
            sg.InputCombo(col, default_value=makeGUI.current_label, size=(20, 1), key='-ILABEL-')],
            [sg.OK(), sg.Exit(), sg.Cancel()],
        ]
        labelWin = sg.Window('Select Your Current and Potential Data', checklayout, finalize=True)
        ev2, vl2 = labelWin.read()
        if ev2 in (None, sg.WIN_CLOSED):
            labelWin.close()
        if ev2 == 'OK':
            makeGUI.potential_label = vl2['-ULABEL-']
            makeGUI.current_label = vl2['-ILABEL-']
            labelWin.close()

    def data_read(self, event, values, I_label=current_label, U_label=potential_label):
        self.excel_path = values['-PATH-']   # 记录包含数据的excel的绝对路径
        self.scan_rate = [float(values[f'-V{i}-']) for i in range(1,10) if values[f'-V{i}CHECK-'] == True]  # 记录初次勾选的扫速值
        self.scan_rate_idx = [i-1 for i in range(1,10) if values[f'-V{i}CHECK-'] == True] # 记录初次勾选扫速的索引
        self.selected_rate = [f'v{i}' for i in range(1,10) if values[f'-V{i}CHECK-'] == True]  # 记录初次勾选扫速的名字v1、v2等
        self.data_list = []
        #——————————————————————————————————————————————————————————————————————————————————————————————————
        progress_layout = [[sg.Text('Reading data... ...')],
                [sg.ProgressBar(9, orientation='h', size=(20, 20), key='-PROGRESSBAR-')], [sg.Cancel()]]
        self.progwin = sg.Window('Custom Progress Meter', progress_layout, finalize=True)
        #——————————————————————————————————————————————————————————————————————————————————————————————————
        file = pd.ExcelFile(self.excel_path)
        scan_names = file.sheet_names
        dots_num = int(self.dotnum)
        for i, name in enumerate(scan_names):
            try:
                df = pd.read_excel(self.excel_path, sheet_name = name)
                if makeGUI.current_label not in tuple(df.columns) or (makeGUI.potential_label not in tuple(df.columns)):
                    self.check_UIlabel(tuple(df.columns))
                data = pd.concat([df[makeGUI.potential_label], df[makeGUI.current_label]*1000.], axis=1) # 此处需要加入检查电压电流标识的弹窗---------------!!!!!!!!!!!!!
                data = data.iloc[-dots_num::self.interval]
                data.index = range(0, len(data))
                data.columns = ['Potential(V)', 'Current(mA)']
                self.data_list.append(data)
                #—————————————————————————————————————
                self.progwin.read(timeout=10)
                self.progwin['-PROGRESSBAR-'].UpdateBar(i + 1)
                # pg = sg.one_line_progress_meter('Progress Meter', i+1, len(scan_names), args=None, size = (20, 20), orientation='h')
                # if pg == False:
                #     self.master.read(timeout=100)
                #—————————————————————————————————————
            except KeyError:
                sg.popup('文件读取错误:检查是否包含空Sheet。')
                continue
            # except xlrd.biffh.XLRDError:
            #     sg.popup('文件读取错误:工作表需以“Sheet”加数字命名。')
            except ValueError:
                sg.popup('文件读取错误:请确认数据点个数。')
                break
            finally:
                pass
        if len(self.scan_rate) > len(self.data_list):  # 这个判断应该每次处理数据都需要，放这里不妥，但是暂时不想改= =!
            sg.popup('选择的扫速不能多于实际数据。')
            raise Exception('选择的扫速不能多于实际数据。')
            pass
        #———————————————
        self.progwin.close()
        self.master.read(timeout=100)
        #———————————————
        self.example = ef.essFunc()
        self.idx = True

    def cdfit(self, event, values):
        try:
            self.dataReadJudge(event, values)
            self.fit_data = self.example.cdFit(self.filt_data_list, self.filt_scan_rate)
            # 作图------------------------------------------------------------------------------------------------------------
            colors = [ef.loop_pick_color(makeGUI.init_color, i*0.75) for i in range(0,9)]
            labels = ['v1', 'v2', 'v3', 'v4', 'v5', 'v6','v7' ,'v8', 'v9']
            fig,((ax1,ax2,ax3),(ax4,ax5,ax6),(ax7,ax8,ax9)) = plt.subplots(3,3)
            axs = [ax1,ax2,ax3,ax4,ax5,ax6,ax7,ax8,ax9]
            for i, vi, nvi in zip(range(0,len(self.selected_rate)), self.selected_idx, self.new_selrate):
                axs[vi].plot(self.data_list[vi]['Potential(V)'], self.data_list[vi]['Current(mA)'],
                color = colors[i],
                linestyle = '-',
                label = 'pristine '+ labels[vi],
                linewidth = 1.5)
                axs[vi].plot(self.fit_data[i]['Potential(V)'], self.fit_data[i]['Capacitance Current(mA)'],
                color = ef.loop_pick_color(colors[i], (i+1)*1.43),
                linestyle = '-',
                label = 'capacitance'+ labels[vi],
                linewidth = 1.5)
                axs[vi].legend(loc='best')
                axs[vi].set_title('scan sweep = '+str(self.scan_rate[nvi])+'mV/s')
                axs[vi].set_xlabel('Potential (V)')
                axs[vi].set_ylabel('Current (mA)')

            plt.show()
        except IndexError:
            sg.popup('请检查每个扫速下对应文件是否存在！')
        except FileNotFoundError:
            self.progwin.close()
            sg.popup('请选择数据文件。')
        except AttributeError:
            self.progwin.close()
            sg.popup('请选择数据文件。')

    def save_Cfit_data(self, event, values):
        try:
            if len(self.fit_data):
                save_path = sg.popup_get_file('Please enter a file name', title = 'Find Your Files', initial_folder=self.workspace,
                    save_as=True, no_window=True, file_types=(("Microsoft Excel2003", "*.xls"),))
                with pd.ExcelWriter(save_path+'.xls') as writer:
                    for fit_data, i, v in zip(self.fit_data, self.new_selrate, self.filt_scan_rate):
                        pd.concat([fit_data, self.data_list[i]['Current(mA)']], axis=1).to_excel(writer, sheet_name=str(v))
            else:
                yon = sg.popup('结果为空，是否先进行数据拟合？')
                if yon:
                    self.cdfit(event, values)
                else:
                    pass
        except FileNotFoundError:
            pass
        except ValueError:
            pass
        except AttributeError:
            pass

    def cdbar(self, event, values):
        capacitance_list = []
        total_capacity_list = []
        if len(self.fit_data):
            for fit_data, sv in zip(self.fit_data, self.new_selrate):
                c_bar, total = 0., 0.
                rsl = pd.concat([fit_data, self.data_list[sv]['Current(mA)']], axis=1)
                rsl.sort_values(by='Potential(V)', ascending=True, inplace=True)
                # rsl.sort_index(by='Potential(V)', ascending=True, inplace=True)
                sampleInterval = abs((rsl.iloc[0,0]-rsl.iloc[-1,0])*2/self.dotnum)
                for i in range(0, len(fit_data)-1):
                        if (rsl.iloc[i,0]-rsl.iloc[i+1,0]) <= sampleInterval*1.05:
                            c_bar += abs((rsl.iloc[i,0]-rsl.iloc[i+1,0])*(rsl.iloc[i,1]-rsl.iloc[i+1,1]))
                            total += abs((rsl.iloc[i,0]-rsl.iloc[i+1,0])*(rsl.iloc[i,2]-rsl.iloc[i+1,2]))
                        else:
                            sg.popup('拟合数据出错，请检查各扫速下数据个数是否一致，\n或在菜单栏中Rectify->Number of data dots更改取点数，使得该值不超过数据点数')
                            break
                capacitance_list.append(c_bar)
                total_capacity_list.append(total)
            # ====================================================================================
            yy = len(self.new_selrate)
            vv = np.linspace(0,yy-1,yy)
            capacitance = np.array(capacitance_list)
            total_capacity = np.array(total_capacity_list)
            self.c_ratio = capacitance / total_capacity * 100
            self.d_ratio = 100 - self.c_ratio
            self.bar_data = pd.concat([pd.Series(vv), pd.Series(self.c_ratio), pd.Series(self.d_ratio), pd.Series(np.array(self.filt_scan_rate))], axis=1)
            self.bar_data.columns = ('sweep when plot', 'Capacitance ratio(%)', 'Diffusion ratio(%)', 'real scan sweep')

            fig,ax = plt.subplots()
            plt.bar(vv, self.c_ratio, color=self.init_color, label='Capacitance')
            plt.bar(vv, self.d_ratio, bottom=self.c_ratio, color='#A9A9A9', label='Diffusion')
            plt.xticks([i for i in vv], self.filt_scan_rate)
            ax.set_ylabel('Contribution ratio (%)')
            ax.set_xlabel('Sweep rate (mV/s)')
            ax.set_ylim(0, 125)
            ax.legend(loc='best')
            try:
                for i in range(0, yy):
                    ax.text(vv[i] - 0.5, 102, str(int(100 * self.c_ratio[i]) / 100))#, bbox = box)
                plt.show()
            except ValueError:
                sg.popup('拟合结果不可靠，请选择合适数据或扫速。')
        else:
            yon = sg.popup('还未进行数据拟合，是否拟合？')
            if yon:
                self.cdfit(event, values)
            else:
                pass
    def save_CD_bar(self, event, values):
        try:
            if self.bar_data.empty == False:
                save_bar_path = sg.popup_get_file('Please enter a file name', title = 'Find Your File', initial_folder=self.workspace,
                    save_as=True, no_window=True, file_types=(("Comma-Separated Values", "*.csv"),))
                self.bar_data.to_csv(save_bar_path )
            else:
                sg.popup('结果为空！')
        except FileNotFoundError:
            pass
        except ValueError:
            pass
        except AttributeError:
            pass

    def selectIntegQ(self):
        self.integ_denom = int(sg.popup_get_text('输入0或1选择Qf', '0: 线性拟合容量\n1: 积分容量', default_text=self.integ_denom, modal=False,))

    # 依据公式 QF = ∫(k1*v + k2*v^1/2)dE/v = ∫k1dE + ∫k2dE*v^(-1/2)进行数据拟合，该方法适用于通常情形
    def integral_fit(self, event, values):
        try:
            self.dataReadJudge(event, values)
            integFitData = self.example.integCdFit(self.filt_data_list, self.filt_scan_rate, self.dotnum)
        # Plot----------------------------------------------------------------------------------
            if self.integ_denom == 0:
                integral_fit_cap_ratio = integFitData[0]
            elif self.integ_denom == 1:
                integral_fit_cap_ratio = integFitData[1]
            
            yy = len(self.new_selrate)
            vv = np.linspace(0,yy-1,yy)
            # 储存数据于一个DataFrame中
            self.integral_fit_data = pd.concat([pd.Series(self.filt_scan_rate),pd.Series(1/np.sqrt(self.filt_scan_rate)),
                pd.Series(integFitData[2]), pd.Series(integFitData[3]),
                pd.Series(integral_fit_cap_ratio), pd.Series(100-integral_fit_cap_ratio)], axis=1)
            self.integral_fit_data.columns = ('selected scan rate(v, mV/s)', '1/v^0.5', 
            'Integral Capacity(Qf)', r'$\int_0^E k1\mathrm{d}E$\+$\int_0^E k2\mathrm{d}E$*v^(-1/2)',
            'Capacitance ratio(%)', 'Diffusion ratio(%)')

            fig,(ax1, ax2) = plt.subplots(1,2)
            ax1.bar(vv, integral_fit_cap_ratio, color=self.init_color, label='Capacitance')
            ax1.bar(vv, 100-integral_fit_cap_ratio, bottom=integral_fit_cap_ratio, 
            color='#A9A9A9', label='Diffusion')
            ax1.set_xticks([i for i in vv])
            ax1.set_xticklabels(self.filt_scan_rate)
            ax1.set_ylabel('Contribution ratio (%)')
            ax1.set_xlabel('Sweep rate (mV/s)')
            ax1.set_ylim(0, 125)
            ax1.legend(loc='best')
            for i in range(0, yy):
                ax1.text(vv[i] - 0.5, 102, str(int(100 * integral_fit_cap_ratio[i]) / 100))
            y_i = integFitData[3]
            ax2.plot(1/np.sqrt(self.filt_scan_rate), y_i, color='r', label='QF = ∫k1dE + ∫k2dE*v^(-1/2)')
            ax2.plot(1/np.sqrt(self.filt_scan_rate), np.array(integFitData[2]), 'o')
            ax2.set_ylabel('Integral Capacity')
            ax2.set_xlabel('Scan rate(v^-0.5, (mV/s)^-0.5)')
            ax2.legend(loc='best')
            text_x = ax2.get_xlim()[1] - ax2.get_xlim()[0]
            text_y = ax2.get_ylim()[1] - ax2.get_ylim()[0]
            ax2.text(ax2.get_xlim()[0] + text_x  * 0.25, ax2.get_ylim()[0] + text_y * 0.75, 
                'intecept(∫k1dE, psedo)=' + str(int(integFitData[4][1] * 1000) / 1000))
            ax2.text(ax2.get_xlim()[0] + text_x * 0.25, ax2.get_ylim()[0] + text_y * 0.65, 
                'slope(∫k2dE, diffusion)=' + str(int(integFitData[4][0] * 1000) / 1000))
            plt.show()
        except ValueError:
            sg.popup('拟合结果不可靠，请选择合适数据或扫速。')
        except FileNotFoundError:
            self.progwin.close()
            sg.popup('请选择数据文件。')
        except AttributeError:
            self.progwin.close()
            sg.popup('请选择数据文件。')

    def save_integfit(self, event, values):
        try:
            if not self.integral_fit_data.empty:
                save_interfit_path = sg.popup_get_file('Please enter a file name', title = 'Find Your File', initial_folder=self.workspace,
                    save_as=True, no_window=True, file_types=(("Comma-Separated Values", "*.csv"),))
                self.integral_fit_data.to_csv(save_interfit_path + '-' + str(self.integ_denom) + '.csv')
            else:
                sg.popup('结果为空！')
        except FileNotFoundError:
            pass
        except ValueError:
            pass
        except AttributeError:
            pass

    def preview_peak_plot(self, event, values):
        try:
            self.dataReadJudge(event, values)
            self.peaksReadJudge(event, values)

            colors = [ef.loop_pick_color(makeGUI.init_color, i*0.75) for i in range(0,9)]
            labels = ['v1', 'v2', 'v3', 'v4', 'v5', 'v6','v7' ,'v8', 'v9']
            fig,((ax1,ax2,ax3),(ax4,ax5,ax6),(ax7,ax8,ax9)) = plt.subplots(3,3)
            axs = [ax1,ax2,ax3,ax4,ax5,ax6,ax7,ax8,ax9]
            for i, vi, nvi in zip(range(0,len(self.selected_rate)), self.selected_idx, self.new_selrate):
                axs[vi].plot(self.data_list[vi]['Potential(V)'], self.data_list[vi]['Current(mA)'],
                color = colors[i],
                linestyle = '-',
                label = 'pristine '+ labels[vi],
                linewidth = 1.5)
                axs[vi].annotate('ox peak1', xy=(self.filt_ox_peak_list[i].iloc[0,0], self.filt_ox_peak_list[i].iloc[0,1]), 
                    xytext=(self.filt_ox_peak_list[i].iloc[0,0]-0.5, self.filt_ox_peak_list[i].iloc[0,1]),
                    color=ef.loop_pick_color(colors[i], (i+1)*1.3),size=7,
                    arrowprops=dict(facecolor='k', shrink=0.05, width=1))

                axs[vi].annotate('red peak1', xy=(self.filt_red_peak_list[i].iloc[0,0], self.filt_red_peak_list[i].iloc[0,1]), 
                    xytext=(self.filt_red_peak_list[i].iloc[0,0]+0.5, self.filt_red_peak_list[i].iloc[0,1]),
                    color=ef.loop_pick_color(colors[i], (i+1)*1.3), size=7,
                    arrowprops=dict(facecolor='k', shrink=0.05, width=1))
                axs[vi].legend(loc='best')
                vii = vi + 1
                axs[vi].set_title('scan sweep = '+str(values[f'-V{vii}-'])+'mV/s')
                axs[vi].set_xlabel('Potential (V)')
                axs[vi].set_ylabel('Current (mA)')
            plt.show()
        except IndexError:
            sg.popup('请检查每个扫速下对应文件是否存在！')
        except FileNotFoundError:
            self.progwin.close()
            sg.popup('请选择数据文件。')
        except AttributeError:
            self.progwin.close()
            sg.popup('请选择数据文件。')

    def peakrectify(self, event, values):
        try:
            self.dataReadJudge(event, values)
            self.peaksReadJudge(event, values)
        except FileNotFoundError:
            sg.popup('请选择数据文件。')
        except AttributeError:
            sg.popup('请选择数据文件。')
        columnAccount = [[sg.Text('Scan->', size=(10, 1), justification='center', font=('courier', 10, 'italic'))],
            [sg.Text('Ox  peaks(Up,Ip)', size=(15, 1), justification='center', font=('courier', 10, 'italic'))],
            [sg.Text('Red peaks(Up,Ip)', size=(15, 1), justification='center', font=('courier', 10, 'italic'))]]
        col = [sg.Column(columnAccount)]
        for i, v, odata, rdata in zip(self.selected_idx, self.filt_scan_rate, self.filt_ox_peak_list, self.filt_red_peak_list):
            i = i+1
            iuText = sg.Column([[sg.Text(str(values[f'-V{i}-'])+'mV/s', size=(10, 1), justification='center', font=('courier', 10, 'italic'))],
            [sg.Input(round(odata.iloc[0,0],3), size=(6,1), pad=(3,10), text_color=makeGUI.init_color, font=('courier', 10, 'italic'), key=f'-UO{i}-'),
             sg.Input(round(odata.iloc[0,1],3), size=(6,1), pad=(3,10), text_color=makeGUI.init_color, font=('courier', 10, 'italic'), key=f'-IO{i}-')],
            [sg.Input(round(rdata.iloc[0,0],3), size=(6,1), pad=(3,10), text_color=makeGUI.init_color, font=('courier', 10, 'italic'), key=f'-UR{i}-'),
             sg.Input(round(rdata.iloc[0,1],3), size=(6,1), pad=(3,10), text_color=makeGUI.init_color, font=('courier', 10, 'italic'), key=f'-IR{i}-')],
            ])
            col.append(iuText)
        layout = [col,
        [sg.OK(), sg.Cancel()],
        ]
        return layout

    def savepeaks(self, event, values):
        try:
            peak_sheet = pd.DataFrame(np.zeros((9,4)), columns=['Ox Potential(V)', 'Ox Current(mA)', 'Red Potential(V)', 'Red Current(mA)'])  
            rate_get = [values[v] for v in [f'-V{i}-' for i in range(1,10)]]
            peak_sheet.index = [str(sc)+' mV/s' for sc in rate_get]
            for j,i in enumerate(self.selected_idx):
                peak_sheet.iloc[i,0] = self.filt_ox_peak_list[j].iloc[0,0]
                peak_sheet.iloc[i,1] = self.filt_ox_peak_list[j].iloc[0,1]
                peak_sheet.iloc[i,2] = self.filt_red_peak_list[j].iloc[0,0]
                peak_sheet.iloc[i,3] = self.filt_red_peak_list[j].iloc[0,1]
            save_path = sg.popup_get_file('Please enter a file name', title = 'Find Your Files', initial_folder=self.workspace,
                    save_as=True, no_window=True, file_types=(("Microsoft Excel2003", "*.xls"),))
            peak_sheet.to_excel(save_path)
        except FileNotFoundError:
            pass
        except ValueError:
            pass
        except AttributeError:
            pass

    def loadpeaks(self, event, values):
        try:
            loaded_ox_peak = []
            loaded_red_peak = []
            peak_path = sg.popup_get_file('Select cards',  initial_folder=self.workspace,
                    save_as=False, no_window=True, multiple_files=False, file_types=(("Microsoft Excel", "*.xls"),))
            self.peak_sheet = pd.read_excel(peak_path)
            if 'Unnamed: 0' in self.peak_sheet.columns:
                new_indx = self.peak_sheet['Unnamed: 0']
                self.peak_sheet = self.peak_sheet.drop(columns='Unnamed: 0')
                self.peak_sheet.index = new_indx
            else:
                pass
            v = [self.master[f'-V{sc}-'] for sc in range(1,10)]
            sel_v = [self.master[f'-V{sc}CHECK-'] for sc in range(1,10)]
            scan_num = len(self.peak_sheet)
            for i, scan_str, sc_val, sc_sel in zip(range(0, scan_num), self.peak_sheet.index, v[:scan_num], sel_v[:scan_num]):
                if all(self.peak_sheet.loc[scan_str] == np.array([0.0,0.0,0.0,0.0]))==False:
                    loaded_ox = pd.DataFrame([[self.peak_sheet.iloc[i,0], self.peak_sheet.iloc[i,1]]],columns=('Potential(V)', 'Current(mA)'))
                    loaded_red = pd.DataFrame([[self.peak_sheet.iloc[i,2], self.peak_sheet.iloc[i,3]]],columns=('Potential(V)', 'Current(mA)'))
                    loaded_ox_peak.append(loaded_ox)
                    loaded_red_peak.append(loaded_red)
                    sc_val.update(float(str(scan_str).split(' ')[0]))
                    sc_sel.update(True)
                else:
                    sc_val.update(float(str(scan_str).split(' ')[0]))
                    sc_sel.update(False)
            if not self.pvidx[0]:
                self.ox_peak_list = loaded_ox_peak
                self.red_peak_list = loaded_red_peak
                self.filt_ox_peak_list = loaded_ox_peak
                self.filt_red_peak_list = loaded_red_peak
                self.pvidx = [True, True]
            elif self.pvidx[0]:
                self.filt_ox_peak_list = loaded_ox_peak
                self.filt_red_peak_list = loaded_red_peak
                self.pvidx = [True, True]
        except FileNotFoundError:
            pass
    
    def showAvb(self, event, values):
        try:
            self.dataReadJudge(event, values)
            self.peaksReadJudge(event, values)

            sv = self.filt_scan_rate
            anode_b, anode_a = self.example.avb(self.filt_ox_peak_list, self.filt_scan_rate)
            cathode_b, cathode_a = self.example.avb(self.filt_red_peak_list, self.filt_scan_rate)
            anode_I = [ia.iloc[0, 1] for ia in self.filt_ox_peak_list]
            cathode_I = [ic.iloc[0, 1] for ic in self.filt_red_peak_list]
            y_a = anode_b * np.log10(sv) + anode_a
            y_c = cathode_b * np.log10(sv) + cathode_a
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
                'anode_slop=' + str(round(anode_b,3)))
            ax.text(ax.get_xlim()[0] + text_x * 0.25, ax.get_ylim()[0] + text_y * 0.65, 
                'cathode_slop=' + str(round(cathode_b,3)))
            plt.show()
        except FileNotFoundError:
            self.progwin.close()
            sg.popup('请选择数据文件。')
        except AttributeError:
            self.progwin.close()
            sg.popup('请选择数据文件。')

    def saveAvb(self, event, values):
        try:
            if self.avb_data.empty == False:
                save_avb_path = sg.popup_get_file('Please enter a file name', title = 'Find Your File', initial_folder=self.workspace,
                    save_as=True, no_window=True, file_types=(("Comma-Separated Values", "*.csv"),))
                self.avb_data.to_csv(save_avb_path)
            else:
                sg.popup('结果为空,请先进行拟合。')
        except FileNotFoundError:
            pass
        except ValueError:
            pass
        except AttributeError:
            pass

    def showDions(self, event, values):
        try:
            self.dataReadJudge(event, values)
            self.peaksReadJudge(event, values)

            sv = self.filt_scan_rate
            anode_sqrtDk, anode_intercept = self.example.R_Sfit(self.filt_ox_peak_list, self.filt_scan_rate)[0]
            cathode_sqrtDk, cathode_intercept = self.example.R_Sfit(self.filt_red_peak_list, self.filt_scan_rate)[0]
            anode_I = self.example.R_Sfit(self.filt_ox_peak_list, self.filt_scan_rate)[1]
            cathode_I = self.example.R_Sfit(self.filt_red_peak_list, self.filt_scan_rate)[1]
            y_a = anode_sqrtDk * np.sqrt(sv) + anode_intercept
            y_c = cathode_sqrtDk * np.sqrt(sv) + cathode_intercept
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
                'anode_k*D^0.5=' + str(round(anode_sqrtDk, 3)))
            ax.text(ax.get_xlim()[0] + text_x * 0.25, ax.get_ylim()[0] + text_y * 0.65, 
                'cathode_k*D^0.5=' + str(round(cathode_sqrtDk, 3)))
            plt.show()
        except FileNotFoundError:
            self.progwin.close()
            sg.popup('请选择数据文件。')
        except AttributeError:
            self.progwin.close()
            sg.popup('请选择数据文件。')

    def saveDions(self, event, values):
        try:
            if self.Dions_data.empty == False:
                save_Dions_path = sg.popup_get_file('Please enter a file name', title = 'Find Your File', initial_folder=self.workspace,
                    save_as=True, no_window=True, file_types=(("Comma-Separated Values", "*.csv"),))
                self.Dions_data.to_csv(save_Dions_path)
            else:
                sg.popup('结果为空,请先进行拟合。')
        except FileNotFoundError:
            pass
        except ValueError:
            pass
        except AttributeError:
            pass

    # 判断是否已读取数据
    def dataReadJudge(self, event, values):
        if len(self.data_list) == 0 or (not self.idx):
            self.data_read(event, values)
        elif len(self.data_list) > 0 or self.idx:
            pass
        self.new_selrate = [self.scan_rate.index(sv) for sv, i in zip(self.scan_rate, [j+1 for j in self.scan_rate_idx]) if values[f'-V{i}CHECK-'] == True]  # 记录当前勾选的扫速在初次勾选扫速list里的索引
        self.selected_idx = [i-1 for i in range(1,10) if values[f'-V{i}CHECK-'] == True]  # 记录当前勾选扫速的在总扫速列表里的索引 总扫速列表即为包含9个“选择框？”的list
        self.filt_data_list = [self.data_list[i] for i in self.new_selrate]  # 从初次读取的数据中选出当前勾选的数据
        self.filt_scan_rate = [self.scan_rate[i] for i in self.new_selrate]  # 记录当前勾选的扫速值
    # 判断是否读取峰值，必须先读取数据
    def peaksReadJudge(self, event, values):
        if not self.pvidx[0]:  # 判断是否读取过峰值，如果为否，则自动识别；为是，则pass
            self.ox_peak_list = []
            self.red_peak_list = []
            for data in self.filt_data_list:
                self.ox_peak_list.append(self.example.searchpeak(data)[0])
                self.red_peak_list.append(self.example.searchpeak(data)[1])
                self.filt_ox_peak_list = self.ox_peak_list
                self.filt_red_peak_list = self.red_peak_list
            self.pvidx[0] = True
        elif self.pvidx[0] and (not self.pvidx[1]):
            self.filt_ox_peak_list = [self.ox_peak_list[i] for i in self.new_selrate]
            self.filt_red_peak_list = [self.red_peak_list[i] for i in self.new_selrate]
        elif self.pvidx[0] and self.pvidx[1]:
            new_filt_ox = [self.ox_peak_list[i] for i in self.new_selrate]
            new_filt_red = [self.red_peak_list[i] for i in self.new_selrate]
            self.filt_ox_peak_list = new_filt_ox
            self.filt_red_peak_list = new_filt_red
