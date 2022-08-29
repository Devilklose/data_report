### 需要取消re_run_flag的机制，或者配置到子模块中----已经配置到子模块中
### 目前仍然只能够判断某日期文件夹下有无CSV文件，没有直接跳过，还不能做到检测全部的符合条件的数据


import os
import glob
import gc

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from datetime import datetime
from scipy.optimize import curve_fit

from pptx import Presentation
from pptx.util import Inches,Pt

import tkinter as tk
from tkinter import filedialog,dialog
from tkinter import StringVar
from tkinter import ttk
#######################################################################
def hebing(folder,signallist,sample_frequency):

    csv_list = glob.glob('*.csv')#查看同文件夹下的csv文件数
##    if (len(csv_list)>0)&('result.csv' not in csv_list):

    if ((len(csv_list)>0)):
        for csv in csv_list[:]:
            if ('Systemdata' not in csv) & ('160P' not in csv) & ('639F' not in csv):
                csv_list.remove(csv)
        print(csv_list)
        print(u'共发现%s个CSV文件'% len(csv_list))
        print(u'正在合并............')
        i = 0
        for csv in csv_list: #循环读取同文件夹下的csv文件
            #fr = open(csv,'r').read()
            try:
                fr = pd.read_csv(csv,usecols=signallist,encoding = 'gb2312',skiprows=lambda x: x > 0 and (x-1) % sample_frequency != 0)
            except:
                fr = pd.read_csv(csv,usecols=signallist,skiprows=lambda x: x > 0 and (x-1) % sample_frequency != 0)
##            fr = pd.read_csv(csv,usecols=signallist,encoding = 'gb2312')
            print(len(fr))
            if i==0:
                f = fr
            else:
                f = pd.concat([f,fr],axis =0)
            i = i +1
        f.index = range(len(f))
##        f = f.iloc[range(0,len(f),sample_frequency)]
##        f.index = range(len(f))
    else:
        f = pd.DataFrame(columns = signallist)
    return f
#######################################################################
def addcolumns(system_file,cvm_file,system_signallist,sample_frequency):## CVM data is all read into ram, so there is no signallist
    
    print ('Concatenate: '+system_file+' and '+cvm_file)
    try:
        df1 = pd.read_csv(system_file,usecols=system_signallist,encoding = 'gb2312',skiprows=lambda x: x > 0 and (x-1) % sample_frequency != 0)
    except:
        df1 = pd.read_csv(system_file,usecols=system_signallist,skiprows=lambda x: x > 0 and (x-1) % sample_frequency != 0)
    try:
        df2 = pd.read_csv(cvm_file,encoding = 'gb2312',skiprows=lambda x: x > 0 and (x-1) % sample_frequency != 0)
    except:
        df2 = pd.read_csv(cvm_file,skiprows=lambda x: x > 0 and (x-1) % sample_frequency != 0)

    df = pd.concat([df1,df2.iloc[:,1:]],axis = 1)

    return df
#######################################################################
def startup_cvm2(folder,df_weg_examination,signallist,sample_frequency):
    csv_list = glob.glob('*.csv')
    base_files = list()
    for csv in csv_list:
        if 'Cell' in csv:
            base_files.append(csv)
    print(base_files)
    i = 0
    for file in base_files:
        df_filtered = addcolumns(file.replace('Cell','Systemdata'),file,signallist,sample_frequency)

        for k in range(len(df_filtered)):
    ##        print(df_filtered.iloc[i]['AbsTime'] in list(df_weg_examination['AbsTime']))
            if (df_filtered.iloc[k]['AbsTime'] in list(df_weg_examination['AbsTime'])) &(i==0):
                df_matched = df_filtered.iloc[k:k+1]##########.iloc[i:i+1]是为了不要出错，其实仅加一行
                i = i + 1
            elif (df_filtered.iloc[k]['AbsTime'] in list(df_weg_examination['AbsTime'])):
                df_matched = pd.concat([df_matched,df_filtered.iloc[k:k+1]],axis =0)##########.iloc[i:i+1]是为了不要出错，其实仅加一行
                i = i + 1
    if 'df_matched' not in vars():
        df_matched = pd.DataFrame(columns = df_filtered.columns)
    print(len(df_matched))


    return df_matched
#######################################################################
def startup_cvm(folder,df_weg_examination,signallist,sample_frequency):
    csv_list = glob.glob('*.csv')
    base_files = list()
    for csv in csv_list:
        if 'Cell' in csv:
            base_files.append(csv)
    print(base_files)
    i = 0
    for file in base_files:
        if i == 0:
            df = addcolumns(file.replace('Cell','Systemdata'),file,signallist,sample_frequency)
            df_filtered = df

        else:
            df = addcolumns(file.replace('Cell','Systemdata'),file,signallist,sample_frequency)
            df_filtered = pd.concat([df_filtered,df],axis = 0)
        i = i + 1
    df_filtered.index = range(len(df_filtered))
    df_matched = pd.DataFrame(columns = df_filtered.columns)
    print(len(df_matched))
    for i in range(len(df_filtered)):
##        print(df_filtered.iloc[i]['AbsTime'] in list(df_weg_examination['AbsTime']))
        if df_filtered.iloc[i]['AbsTime'] in list(df_weg_examination['AbsTime']):
            df_matched = pd.concat([df_matched,df_filtered.iloc[i:i+1]],axis =0)##########.iloc[i:i+1]是为了不要出错，其实仅加一行

    return df_matched
#######################################################################
def stable_select(a_list):
    df_id = np.insert(np.diff(a_list),0,[1])
    df_id_break = [i for i,x in enumerate(df_id) if x>1]
    print(df_id_break)
    if len(df_id_break)==0:
        df_id_break_st = [a_list[0]]
        df_id_break_end = [a_list[-1]]
    else:
        df_id_break_st=[a_list[0]]
        df_id_break_end = list()
        for i in df_id_break:
            df_id_break_st.append(a_list[i])
            df_id_break_end.append(a_list[i-1])
            
##        df_id_break_st = np.insert(df_id_break,0,a_list[0])
##        df_id_break_end = [i-1 for i in df_id_break]
        df_id_break_end.append(a_list[-1])

##    print(df_id_break_st)
##    print(df_id_break_end)
    return df_id_break_st,df_id_break_end
#######################################################################
def folder_select(folder,sys_type):
    delete_list = ['txt','xlsx','png','ppt','pptx','csv','docx','rar']
    subfolder_list = glob.glob('*')
    for subfolder in subfolder_list[:]:
        if (sys_type not in subfolder) or (subfolder[(subfolder.find('.')+1):] in delete_list):
            subfolder_list.remove(subfolder)
    ## path supplement

    return subfolder_list
#######################################################################
def offtime(data,stage_signal):

    air_sdix = list()
    air_sdix2 = list()
    ocv50 = list()
    ocv30 = list()
    flag=list()
    start_temp=list()
    
    for i in range(len(data)-1):


        if ((data[stage_signal][i] == 50) & ((data[stage_signal][i+1] - data[stage_signal][i])>0)):
            air_sdix.append(i)
            

    print(len(air_sdix[1:]))

    for index in air_sdix[1:]:
        j = 1
        while (((data[stage_signal][index-j])!=110)&((data[stage_signal][index-j])!=130)&((data[stage_signal][index-j])!=90)):
##        while ((data['STM_n_MainSt'][index-j])!=90):
            j = j+1
        air_sdix2.append((index-j))
        if ((1 in list(data['SDM_n_FltGrd'][index-j-5:index-j+5]))|(2 in list(data['SDM_n_FltGrd'][index-j-5:index-j+5]))):#choose 5secs before and after the state transient because 1st and 2nd level fault will not eliminite itself until restart the 24V.
            flag.append(1)
        else:
            flag.append(0)

        k = 1
        while (((data[stage_signal][index-k])==50)|((data[stage_signal][index-k])==30)):
            k = k+1
        try:
            ocv50.append(max(data['SPM_U_StkVolt'][(index-k):index+1][data[stage_signal]==50]))
        except:
            ocv50.append(-1000)
        try:
            ocv30.append(max(data['SPM_U_StkVolt'][(index-k):index+1][data[stage_signal]==30]))
        except:
            ocv30.append(-1000)
        m = 1
        while (((data[stage_signal][index-m])==50)|((data[stage_signal][index-m])==30)):
            m = m+1        
        start_temp.append(min(data['CDC_T_TempIn'][(index-m):index+1]))
        

    print(len(air_sdix2))


        

    air_shutdownpoint = data['AbsTime'][air_sdix[1:]]
    air_shutdownpoint2 = data['AbsTime'][air_sdix2]
    print(len(air_shutdownpoint))
    print(len(air_shutdownpoint2))


    Time_offend = list(air_shutdownpoint)
    Time_offstart = list(air_shutdownpoint2)
    offtime = list()

    for i in range(len(Time_offend)):
        ta=datetime.strptime(Time_offend[i], '%Y_%m_%d %H:%M:%S.%f')
        tb=datetime.strptime(Time_offstart[i],'%Y_%m_%d %H:%M:%S.%f')

        offtime.append((ta-tb).total_seconds())


    result = pd.DataFrame({'Time_offend':Time_offend,'Time_offstart':Time_offstart,'offtime':offtime,'ocv30':ocv30,'ocv50':ocv50,'flag':flag,'start_temp':start_temp})
##    result.to_csv('./result.csv')
    

    return result
#######################################################################
def weg_examination(data,signallist,stage_signal):

    air_sdix = list()
    air_sdix2 = list()
    flag=list()
    df_weg_examination = pd.DataFrame(columns=signallist)
    
    for i in range(len(data)-1):

        if ((data[stage_signal][i] == 50) & ((data[stage_signal][i+1] - data[stage_signal][i])>0)):
            air_sdix.append(i)            

    print(len(air_sdix[1:]))

    for index in air_sdix[1:]:
        j = 1
        while (((data[stage_signal][index-j])!=110)&((data[stage_signal][index-j])!=130)&((data[stage_signal][index-j])!=90)):
##        while ((data['STM_n_MainSt'][index-j])!=90):
            j = j+1
        air_sdix2.append((index-j))
        if ((1 in list(data['SDM_n_FltGrd'][index-j-5:index-j+5]))|(2 in list(data['SDM_n_FltGrd'][index-j-5:index-j+5]))):#choose 5secs before and after the state transient because 1st and 2nd level fault will not eliminite itself until restart the 24V.
            flag.append(1)
        else:
            flag.append(0)

        m = 1
        while ((data[stage_signal][index-m]==50)|(data[stage_signal][index-m]==30)):
            m = m+1        
        df = data.loc[(index-m):index+10]
        df_weg_examination = pd.concat([df_weg_examination,(data.loc[(index-m):index+10]),pd.DataFrame(columns=signallist,index = [max(data.loc[(index-m):index+10].index)+1])],axis =0)
        

    print(len(air_sdix2))


        

    air_shutdownpoint = data['AbsTime'][air_sdix[1:]]
    air_shutdownpoint2 = data['AbsTime'][air_sdix2]
    print(len(air_shutdownpoint))
    print(len(air_shutdownpoint2))


    Time_offend = list(air_shutdownpoint)
    Time_offstart = list(air_shutdownpoint2)
    offtime = list()

    for i in range(len(Time_offend)):
        ta=datetime.strptime(Time_offend[i], '%Y_%m_%d %H:%M:%S.%f')
        tb=datetime.strptime(Time_offstart[i],'%Y_%m_%d %H:%M:%S.%f')

        offtime.append((ta-tb).total_seconds())


    result = pd.DataFrame({'Time_offend':Time_offend,'Time_offstart':Time_offstart,'offtime':offtime})
##    result.to_csv('./result.csv')
    

    return result,df_weg_examination
#######################################################################
def sigmoid(x, L ,x0, k, b):
    y = L / (1 + np.exp(-k*(x-x0)))+b
    return (y)


def SCurveFit(df,cell_num):
    xdata = df['offtime']/3600
    ydata = df['ocv50']/cell_num
    p0 = [max(ydata), np.median(xdata),1,min(ydata)] # this is an mandatory initial guess
    popt, pcov = curve_fit(sigmoid, xdata, ydata,p0, method='dogbox')
    x = np.linspace(0, 100, 1000)
    y = sigmoid(x, *popt)


    fig1, ax1 = plt.subplots()
    plt.plot(xdata, ydata, 'o', label='data')
    plt.plot(x,y, label='fit')
    ax1.set_ylim(0, 1.3)
    ax1.set_xscale('log')
    ax1.set_xlabel("Offtime(hrs)")
    ax1.set_ylabel("Max OCV without air at start up(V)")
    ax1.set_title('S-curve of 211P')
    plt.legend(loc='best')
    fig1.savefig('./S-curve.jpg')
    return 0
#######################################################################
def average(df,interval,name):
    df.index = range(len(df))
    df_id = np.insert(np.diff(df['RunTime']),0,[1])
    df_id_break_i = [i for i,x in enumerate(df_id) if x>interval]
    df_id_break = df.index[df_id_break_i]

    if len(df_id_break)==0:
        df_id_break_st = df.index[[0]]
        df_id_break_end = df.index[[-1]]
    else:
        df_id_break_st = np.insert(df_id_break,0,df.index[0])
        df_id_break_end_i = [i-1 for i in df_id_break_i]
        df_id_break_end = df.index[df_id_break_end_i]
        df_id_break_end = df_id_break_end.append(df.index[[-1]])        
    
##    abstime = list()
    abstime = list()
    time = list()
    volt_max = list()
    volt_mean = list()
    coolant_in_temp = list()
    air_rh_in = list()
    for i in range(len(df_id_break_st)):
        if (df_id_break_end[i]-df_id_break_st[i])>interval:
##            abstime.append(df['上报时间'][df_id_break_end[i]])
            abstime.append(df[Signal_abstime][df_id_break_end[i]])
            time.append(df['RunTime'][df_id_break_end[i]]/3600)
            volt_max.append(np.max(df['SPM_U_StkVolt'][df_id_break_st[i]+interval:df_id_break_end[i]]))
##            print(np.max(df['SPM_U_StkVolt'][df_id_break_st[i]+interval:df_id_break_end[i]]))
            volt_mean.append(np.mean(df['SPM_U_StkVolt'][df_id_break_st[i]+interval:df_id_break_end[i]]))
            coolant_in_temp.append(np.mean(df['CDC_T_TempIn'][df_id_break_st[i]+interval:df_id_break_end[i]]))
            air_rh_in.append(np.mean(df['TB_pct_SpclRHStkAirIn'][df_id_break_st[i]+interval:df_id_break_end[i]]))
        else:
            continue
    result = pd.DataFrame({'time':time,'volt_max':volt_max,'volt_mean':volt_mean,'coolant_in_temp':coolant_in_temp,'air_rh_in':air_rh_in})

    return result

def DTCgrd_gen(df,level):
    global DTC_name,i
    time = list()
    DTC_code = list()
    
    if level ==1:
        DTC_name = 'SDM_n_Grd1DTC' 
    elif level ==2:
        DTC_name = 'SDM_n_Grd2DTC'
    elif level ==3:
        DTC_name = 'SDM_n_Grd3DTC' 
    elif level ==4:
        DTC_name = 'SDM_n_Grd4DTC' 
    for i in range(len(df)-1):
        if (df.loc[i+1,DTC_name]>0) and ((df.loc[i+1,DTC_name]-df.loc[i,DTC_name])!=0):
            time.append(df.loc[i+1,Signal_abstime])
            DTC_code.append(df.loc[i+1,DTC_name])
    DTC_stat = pd.DataFrame({'time':time,'DTC_code':DTC_code})

    return DTC_stat
#######################################################################
def General():
    print('And you also want to re_run historical data(Yes = 1, No = 0):',General_re_run_flag)
## 1. Operating hours---done
## 2. Operating days---done
## 3. Duration of soaks (i.e., time between SD and SU)
## 4. Duration of operating time (i.e., time between SU and SD)
## 5. start-ups---done
## 6. FSUs---done
## 7. Current - Max, duration in certain ranges or time below a certain current---done
## 8. Coolant Temperature - Max, duration in certain ranges---done
## 9. Stack & system power - Max, duration in certain ranges---done
## 10. Vehicle speed - Max, duration in certain ranges
## 11. Stack Voltage - Max, Min, duration in certain range or time above certain voltage---done

    delete_list = ['txt','xlsx','png','ppt','pptx','csv','docx','rar','pdf','vsdx','xls','xlsm']

    for folder in seleted_folders:
        General_file = folder + 'General.xlsx'
        sysfolder = root_path + '\\' + folder + '\\'
        os.chdir(sysfolder)
        if ((os.path.exists(General_file)==0) or (General_re_run_flag==1) or (Allrun_re_run_flag==1)):
            print('Now generate <General> report for ',folder)
            datefolder_list = glob.glob('*')
            
            for datefolder in datefolder_list[:]:
                if (datefolder[-5:][(datefolder[-5:].find('.')+1):] in delete_list):
                    datefolder_list.remove(datefolder)
            print(datefolder_list)

            accum_hours = 0
            accum_days = 0
            start_up_times = 0
            FSU_times = 0
            max_curr = list()
            max_volt = list()
            max_coolant_temp_in = list()
            max_coolant_temp_out = list()
            max_stk_pwr = list()
            dur_curr = list()
            dur_volt = list()
            dur_coolant_temp_in = list()
            dur_coolant_temp_out = list()
            dur_stk_pwr = list()
            date = list()

            for datefolder in datefolder_list:
                ## confirm there is a CSV folder in the datefolder, unless errors come out.
                try:
                    csvfolder = sysfolder + datefolder + '\\CSV\\'
                    os.chdir(csvfolder)
                except:
                    print('No CSV file folder found!! ',datefolder) 
                    continue

                    
                
                signallist = ['SPM_I_StkCurr','SPM_U_StkVolt','CDC_T_TempIn','CDC_T_TempOut']
                df = hebing(csvfolder,signallist,sample_frequency)
                print(len(df))

                df['SPM_P_StkPwr'] = df['SPM_I_StkCurr']*df['SPM_U_StkVolt']/1000
                
                max_curr.append(max(df['SPM_I_StkCurr']))
                max_volt.append(max(df['SPM_U_StkVolt']))
                max_coolant_temp_in.append(max(df['CDC_T_TempIn'][df['CDC_T_TempIn']<120]))##120 means temperature sensor out of range, therefore we need to exclude 120.
                max_coolant_temp_out.append(max(df['CDC_T_TempOut'][df['CDC_T_TempOut']<120]))##120 means temperature sensor out of range, therefore we need to exclude 120.
                max_stk_pwr.append(max(df['SPM_P_StkPwr']))
                dur_curr.append(len(df[(df['SPM_I_StkCurr']>300)&(df['SPM_I_StkCurr']<500)]))
                dur_volt.append(len(df[(df['SPM_U_StkVolt']>4)&(df['SPM_U_StkVolt']<500)]))
                dur_coolant_temp_in.append(len(df[(df['CDC_T_TempIn']>65)&(df['CDC_T_TempIn']<80)]))
                dur_coolant_temp_out.append(len(df[(df['CDC_T_TempOut']>80)&(df['CDC_T_TempOut']<95)]))
                dur_stk_pwr.append(len(df[(df['SPM_P_StkPwr']>20)&(df['SPM_P_StkPwr']<80)]))
                date.append(datefolder)
                running_hrs = len(df[df['SPM_I_StkCurr']>5]) ## 1
                
                accum_hours = accum_hours + running_hrs
                if running_hrs>0:
                    accum_days = accum_days + 1 ## 2
                print(accum_days,'days')
                if len(df[df['SPM_I_StkCurr']>5])>0:
                    df_id_break_st,df_id_break_end = stable_select(list(df[df['SPM_I_StkCurr']>5].index)) #5,6
                    start_up_times = start_up_times + len(df_id_break_st)
                    FSU_times = FSU_times + len(df.loc[df_id_break_st][df.loc[df_id_break_st]['CDC_T_TempIn']<5])
    ##------------------------------------------------------------------------
            os.chdir(sysfolder)
            max_and_duration = pd.DataFrame({'date':date,'max_curr':max_curr,'max_voltage':max_volt,'max_coolant_temp_in':max_coolant_temp_in,
                                             'max_coolant_temp_out':max_coolant_temp_out,'max_stk_pwr':max_stk_pwr,
                                             'dur_curr':dur_curr,'max_volt':max_volt,'dur_coolant_temp_in':dur_coolant_temp_in,
                                             'dur_coolant_temp_out':dur_coolant_temp_out,'dur_stk_pwr':dur_stk_pwr})
            running_stat = pd.DataFrame({'accum_hours':accum_hours,'accum_days':accum_days,'start_up_times':start_up_times,'FSU_times':FSU_times},index=[0])
            
            with pd.ExcelWriter(General_file) as writer:
                max_and_duration.to_excel(writer,sheet_name='max_and dur')## 7,8,9,11
                running_stat.to_excel(writer,sheet_name='running statistics ') ## 1,2
            
            print(accum_days,'days')
            print(accum_hours/3600*processdata_s,'hours')
        else:
            print('Results folder<',General_file,'> exists, do not have to run it again.')

    print('All analysis <General> completed, please choose again or Quit GUI!')
    return 0
#######################################################################
def EStress():

    print('And you also want to re_run historical data(Yes = 1, No = 0):',Estress_re_run_flag)
## 1. Air/air start ups (or stack voltage when fuel flow starts) - #, temperature---done
## 2. Deviations in start-ups "off the S-curve"---done
## 3. OCV - # of times, duration, temperature---done
## 4. Idle - # of times, duration---done
## 5. High loads - # of times > certain load---done

    signallist = ['SDM_n_FltGrd','SPM_I_StkCurr','SPM_U_StkVolt','CDC_T_TempIn']
    signallist.append(Signal_subst)
    signallist.extend(Signal_CVM_avg)
    signallist.extend(Signal_CVM_std)
    signallist.append(Signal_abstime)
    
    
    delete_list = ['txt','xlsx','png','ppt','pptx','csv','docx','rar','pdf','vsdx','xls','xlsm']
    print(seleted_folders)
    for folder in seleted_folders:
        EStress_file = folder + 'EStress.xlsx'
        sysfolder = root_path + '\\' + folder + '\\'
        os.chdir(sysfolder)
        
        if (os.path.exists(EStress_file)==0) or (Estress_re_run_flag==1) or (Allrun_re_run_flag==1):
            print('Now generate <EStress> report for ',folder)
            df_airair = pd.DataFrame(columns=signallist)
            datefolder_list = glob.glob('*')
            
            for datefolder in datefolder_list[:]:
                if (datefolder[-5:][(datefolder[-5:].find('.')+1):] in delete_list):
                    datefolder_list.remove(datefolder)
            print(datefolder_list)


            upperstack_ocv_duration = list()
            upperstack_ocv_temperature = list()
            upperstack_ocv_timepoint = list()
            lowerstack_ocv_duration = list()
            lowerstack_ocv_temperature = list()
            lowerstack_ocv_timepoint = list()
            idle_duration = list()
            idle_temperature = list()
            idle_timepoint = list()
            rated_duration = list()
            rated_temperature = list()
            rated_timepoint = list()
            for datefolder in datefolder_list:
                try:
                    csvfolder = sysfolder + datefolder + '\\CSV\\'
                    os.chdir(csvfolder)
                except:
                    print('No CSV file folder found!! ',datefolder) 
                    continue
                df = hebing(csvfolder,signallist,sample_frequency)
                
                if ('Average_1' in df.columns) and (len(df[df['Average_1']>825])>0):
                    df_id_break_st,df_id_break_end = stable_select(list(df[df['Average_1']>825].index))
                    for i in range(len(df_id_break_st)):
                        upperstack_ocv_duration.append((df_id_break_end[i]-df_id_break_st[i]+1)*processdata_s)
                        upperstack_ocv_temperature.append(df.loc[df_id_break_end[i],'CDC_T_TempIn'])
                        upperstack_ocv_timepoint.append(df.loc[df_id_break_end[i],Signal_abstime])
                if ('Average_2' in df.columns) and (len(df[df['Average_2']>825])>0):
                    df_id_break_st,df_id_break_end = stable_select(list(df[df['Average_2']>825].index))
                    for i in range(len(df_id_break_st)):
                        lowerstack_ocv_duration.append((df_id_break_end[i]-df_id_break_st[i]+1)*processdata_s)
                        lowerstack_ocv_temperature.append(df.loc[df_id_break_end[i],'CDC_T_TempIn'])
                        lowerstack_ocv_timepoint.append(df.loc[df_id_break_end[i],Signal_abstime])
                if len(df[(df['SPM_I_StkCurr']<(idle_curr_set+5))&(df['SPM_I_StkCurr']>(idle_curr_set-5))])>0:                      
                    df_id_break_st,df_id_break_end = stable_select(list(df[(df['SPM_I_StkCurr']<(idle_curr_set+5))&(df['SPM_I_StkCurr']>(idle_curr_set-5))].index))
                    for i in range(len(df_id_break_st)):
                        idle_duration.append((df_id_break_end[i]-df_id_break_st[i]+1)*processdata_s)
                        idle_temperature.append(df.loc[df_id_break_end[i],'CDC_T_TempIn'])
                        idle_timepoint.append(df.loc[df_id_break_end[i],Signal_abstime])
                if len(df[df['SPM_I_StkCurr']>(rated_curr_set-5)])>0:                      
                    df_id_break_st,df_id_break_end = stable_select(list(df[df['SPM_I_StkCurr']>(rated_curr_set-5)].index))
                    for i in range(len(df_id_break_st)):
                        rated_duration.append((df_id_break_end[i]-df_id_break_st[i]+1)*processdata_s)
                        rated_temperature.append(df.loc[df_id_break_end[i],'CDC_T_TempIn'])
                        rated_timepoint.append(df.loc[df_id_break_end[i],Signal_abstime])
                
                df_airair = pd.concat([df_airair,df],axis =0)
            df_airair.index = range(len(df_airair))
            airair_result = offtime(df_airair,Signal_subst)
    ##------------------------------------------------------------------------            
            os.chdir(sysfolder)
    ##------------------------------------------------------------------------
    ## draw S-curve line when air-air points are more than 100, in case the fitting results looks strange.
            if len(airair_result)>100:
                SCurveFit(airair_result,cell_num)#2
            upperstack_ocv_stat = pd.DataFrame({'upperstack_ocv_duration':upperstack_ocv_duration,'upperstack_ocv_temperature':upperstack_ocv_temperature,
                                     'upperstack_ocv_timepoint':upperstack_ocv_timepoint})
            lowerstack_ocv_stat = pd.DataFrame({'lowerstack_ocv_duration':lowerstack_ocv_duration,'lowerstack_ocv_temperature':lowerstack_ocv_temperature,
                                     'lowerstack_ocv_timepoint':lowerstack_ocv_timepoint})
            idle_stat = pd.DataFrame({'idle_duration':idle_duration,'idle_temperature':idle_temperature,
                                     'idle_timepoint':idle_timepoint})
            rated_stat = pd.DataFrame({'rated_duration':rated_duration,'rated_temperature':rated_temperature,
                                     'rated_timepoint':rated_timepoint})
            with pd.ExcelWriter(EStress_file) as writer:
                airair_result.to_excel(writer,sheet_name='airair_result')## 1
                upperstack_ocv_stat.to_excel(writer,sheet_name='upper stack ocv stat') ## 3
                lowerstack_ocv_stat.to_excel(writer,sheet_name='lower stack ocv stat') ## 3
                idle_stat.to_excel(writer,sheet_name='idle stat') ## 4
                rated_stat.to_excel(writer,sheet_name='rated stat') ## 4
        else:
            print('Results folder<',EStress_file,'> exists, do not have to run it again.')
    print('All analysis <Estress> completed, please choose again or Quit GUI!')
    return 0
#######################################################################
def MechStress():
    print('And you also want to re_run historical data(Yes = 1, No = 0):',MechStress_re_run_flag)
## 1. Anode-Cathode cross pressure (min, max, duration > value, occurrences > value) --done
## 2. Coolant-reactant cross pressure  (min, max, duration > value, occurrences > value, coolant pressure during SD) --done partially, need to solve coolant sensor issue.
## 3. Maximum fluid pressures --done
## 4. # of EPOs -- allocated to airair starts in Estress module.
    signallist = ['SDM_n_FltGrd','ADC_p_PrsIn','HDC_p_PrsIn','CDC_p_PrsIn']
    signallist.append(Signal_subst)
##    signallist.extend(Signal_CVM_avg)
##    signallist.extend(Signal_CVM_std)
    signallist.append(Signal_abstime)

    delete_list = ['txt','xlsx','png','ppt','pptx','csv','docx','rar','pdf','vsdx','xls','xlsm']
    print(seleted_folders)
    for folder in seleted_folders:
        MechStress_file = folder + 'MechStress.xlsx'
        sysfolder = root_path + '\\' + folder + '\\'
        os.chdir(sysfolder)
        if (os.path.exists(MechStress_file)==0) or (MechStress_re_run_flag==1) or (Allrun_re_run_flag==1):
            print('Now generate <MechStress> report for ',folder)
            datefolder_list = glob.glob('*')
            
            for datefolder in datefolder_list[:]:
                if (datefolder[-5:][(datefolder[-5:].find('.')+1):] in delete_list):
                    datefolder_list.remove(datefolder)
            print(datefolder_list)

            fx_prs_max = list()
            ox_prs_max = list()
            cx_prs_max = list()
            fx_ox_dp_max = list()
            fx_ox_dp_min = list()
            fx_cx_dp_max = list()
            fx_cx_dp_min = list()
            ox_cx_dp_max = list()
            ox_cx_dp_min = list()
            fx_ox_dp_over_duration = list()
            fx_ox_dp_over_timepoint = list()

            date = list()

            for datefolder in datefolder_list:
                try:
                    csvfolder = sysfolder + datefolder + '\\CSV\\'
                    os.chdir(csvfolder)
                except:
                    print('No CSV file folder found!! ',datefolder) 
                    continue
                df = hebing(csvfolder,signallist,sample_frequency)

                fx_prs_max.append(max(df['HDC_p_PrsIn']))
                ox_prs_max.append(max(df['ADC_p_PrsIn']))
                cx_prs_max.append(max(df['CDC_p_PrsIn']))
                fx_ox_dp_max.append(max(df['HDC_p_PrsIn']-df['ADC_p_PrsIn']))
                fx_ox_dp_min.append(min(df['HDC_p_PrsIn']-df['ADC_p_PrsIn']))
                fx_cx_dp_max.append(max(df['HDC_p_PrsIn']-df['CDC_p_PrsIn']))
                fx_cx_dp_min.append(min(df['HDC_p_PrsIn']-df['CDC_p_PrsIn']))
                ox_cx_dp_max.append(max(df['ADC_p_PrsIn']-df['CDC_p_PrsIn']))
                ox_cx_dp_min.append(min(df['ADC_p_PrsIn']-df['CDC_p_PrsIn']))
                date.append(datefolder)
            
                if len(df[(df['HDC_p_PrsIn']-df['ADC_p_PrsIn'])<-10])>0:                      
                    df_id_break_st,df_id_break_end = stable_select(list(df[(df['HDC_p_PrsIn']-df['ADC_p_PrsIn'])<-10].index))
                    for i in range(len(df_id_break_st)):
                        fx_ox_dp_over_duration.append((df_id_break_end[i]-df_id_break_st[i])*processdata_s)
                        fx_ox_dp_over_timepoint.append(df.loc[df_id_break_end[i],Signal_abstime])

    ##------------------------------------------------------------------------
            os.chdir(sysfolder)
            
            prs_max_min_stat = pd.DataFrame({'date':date,'fx_prs_max':fx_prs_max,'ox_prs_max':ox_prs_max,'cx_prs_max':cx_prs_max,
                                             'fx_ox_dp_max':fx_ox_dp_max,'fx_ox_dp_min':fx_ox_dp_min,'fx_cx_dp_max':fx_cx_dp_max,
                                             'fx_cx_dp_min':fx_cx_dp_min,'ox_cx_dp_max':ox_cx_dp_max,'ox_cx_dp_min':ox_cx_dp_min})
            dp_over_duration = pd.DataFrame({'fx_ox_dp_over_duration':fx_ox_dp_over_duration,'fx_ox_dp_over_timepoint':fx_ox_dp_over_timepoint})

            with pd.ExcelWriter(MechStress_file) as writer:
                prs_max_min_stat.to_excel(writer,sheet_name='prs_max_min_stat')## 1,3
                dp_over_duration.to_excel(writer,sheet_name='dp_over_duration')## 2
        else:
            print('Results folder<',MechStress_file,'> exists, do not have to run it again.')
   
    print('All analysis <MechStress> completed, please choose again or Quit GUI!')
    return 0
#######################################################################
def Contamination():
    print('And you also want to re_run historical data(Yes = 1, No = 0):',Contamination_re_run_flag)
## 1. Low OCV before airflow & after long SD (standard deviation) - indication of WEG contamination (# of cells TBD mV before ACV)

    signallist = ['SDM_n_FltGrd','SPM_I_StkCurr','SPM_U_StkVolt','CDC_T_TempIn']
    signallist.append(Signal_subst)
    signallist.extend(Signal_CVM_avg)
    signallist.extend(Signal_CVM_std)
    signallist.append(Signal_abstime)
    print()

    
    delete_list = ['txt','xlsx','png','ppt','pptx','csv','docx','rar','pdf','vsdx','xls','xlsm']
    print(seleted_folders)
    for folder in seleted_folders:
        Contamination_file = folder + 'Contamination.xlsx'
        sysfolder = root_path + '\\' + folder + '\\'
        os.chdir(sysfolder)
        if (os.path.exists(Contamination_file)==0) or (Contamination_re_run_flag==1) or (Allrun_re_run_flag==1):
            print('Now generate <Contamination> report for ',folder)
            df_coolant_conta = pd.DataFrame(columns=signallist)
            datefolder_list = glob.glob('*')
            
            for datefolder in datefolder_list[:]:
                if (datefolder[-5:][(datefolder[-5:].find('.')+1):] in delete_list):
                    datefolder_list.remove(datefolder)
            print(datefolder_list)

            for datefolder in datefolder_list:
                try:
                    csvfolder = sysfolder + datefolder + '\\CSV\\'
                    os.chdir(csvfolder)
                except:
                    print('No CSV file folder found!! ',datefolder) 
                    continue
                df = hebing(csvfolder,signallist,sample_frequency)

                
                df_coolant_conta = pd.concat([df_coolant_conta,df],axis =0)
            df_coolant_conta.index = range(len(df_coolant_conta))
            offtime,df_weg_examination = weg_examination(df_coolant_conta,signallist,Signal_subst)


            os.chdir(sysfolder)
            with pd.ExcelWriter(Contamination_file) as writer:
                offtime.to_excel(writer,sheet_name='offtime')## 1,3
                df_weg_examination.to_excel(writer,sheet_name='df_weg_examination')## 2

            if ((Contamination_cvm_flag == 1) or (Allrun_re_run_flag == 1)) and ((sys_type != "160P-DL3-2-006") and (sys_type != "211P-PP-639")):
                count = 1
                for datefolder in datefolder_list:
                    try:
                        csvfolder = sysfolder + datefolder + '\\CSV\\'
                        os.chdir(csvfolder)
                    except:
                        print('No CSV file folder found!! ',datefolder) 
                        continue
                    if count ==1:
                        df_matched = startup_cvm2(csvfolder,df_weg_examination,signallist,sample_frequency)
                    else:
                        df_matched = pd.concat([df_matched,startup_cvm2(csvfolder,df_weg_examination,signallist,sample_frequency)],axis=0)
                    count = count+1
                os.chdir(sysfolder)
                with pd.ExcelWriter(Contamination_file) as writer:
                    offtime.to_excel(writer,sheet_name='offtime')## 1,3
                    df_weg_examination.to_excel(writer,sheet_name='df_weg_examination')## 2
                    df_matched.to_excel(writer,sheet_name='df_weg_cvm_data')

        else:
            print('Results folder<',Contamination_file,'> exists, do not have to run it again.')

    print('All analysis <Contamination> completed, please choose again or Quit GUI!')
    return 0
#######################################################################
def DryEvents():
    print('All analysis <DryEvents> completed, please choose again or Quit GUI!')
    return 0
#######################################################################
def FuelStarve():
    
    print('And you also want to re_run historical data(Yes = 1, No = 0):',FuelStarve_re_run_flag)
## 1. Low OCV before airflow & after long SD (standard deviation) - indication of WEG contamination (# of cells TBD mV before ACV)

    signallist = ['SDM_n_FltGrd','SPM_I_StkCurr','SPM_U_StkVolt','CDC_T_TempIn']
    signallist.append(Signal_subst)
    signallist.extend(Signal_CVM_avg)
    signallist.extend(Signal_CVM_std)
    signallist.append(Signal_abstime)

    
    delete_list = ['txt','xlsx','png','ppt','pptx','csv','docx','rar','pdf','vsdx','xls','xlsm']
    print(seleted_folders)
    for folder in seleted_folders:
        FuelStarve_file = folder + 'Contamination.xlsx'
        sysfolder = root_path + '\\' + folder + '\\'
        os.chdir(sysfolder)
        if (os.path.exists(FuelStarve_file)==0) or (FuelStarve_re_run_flag==1) or (Allrun_re_run_flag==1):
            print('Now generate <Contamination> report for ',folder)
            df_coolant_conta = pd.DataFrame(columns=signallist)
            datefolder_list = glob.glob('*')
            
            for datefolder in datefolder_list[:]:
                if (datefolder[-5:][(datefolder[-5:].find('.')+1):] in delete_list):
                    datefolder_list.remove(datefolder)
            print(datefolder_list)

            for datefolder in datefolder_list:
                try:
                    csvfolder = sysfolder + datefolder + '\\CSV\\'
                    os.chdir(csvfolder)
                except:
                    print('No CSV file folder found!! ',datefolder) 
                    continue
                df = hebing(csvfolder,signallist,sample_frequency)
                print(len(df))
            os.chdir(sysfolder)
##            with pd.ExcelWriter(FuelStarve_file) as writer:
##                offtime.to_excel(writer,sheet_name='offtime')## 1,3
##                df_weg_examination.to_excel(writer,sheet_name='df_weg_examination')## 2

    print('All analysis <FuelStarve> completed, please choose again or Quit GUI!')
    return 0
#######################################################################
def Polarization():
    print('And you also want to re_run historical data(Yes = 1, No = 0):',Polarization_re_run_flag)
## 1. Low OCV before airflow & after long SD (standard deviation) - indication of WEG contamination (# of cells TBD mV before ACV)
    global df_rateddata,df_idledata
    signallist = ['SDM_n_FltGrd','SPM_I_StkCurr','SPM_U_StkVolt','CDC_T_TempIn']
    signallist.append(Signal_subst)
    signallist.extend(Signal_CVM_avg)
    signallist.extend(Signal_CVM_std)
    signallist.append(Signal_abstime)
    signallist.append('TB_pct_SpclRHStkAirIn')


    
    delete_list = ['txt','xlsx','png','ppt','pptx','csv','docx','rar','pdf','vsdx','xls','xlsm']

    for folder in seleted_folders:
        Polarization_file = folder + 'Polarization.xlsx'
        sysfolder = root_path + '\\' + folder + '\\'
        os.chdir(sysfolder)
        if (os.path.exists(Polarization_file)==0) or (Polarization_re_run_flag==1) or (Allrun_re_run_flag==1):
            print('Now generate <Polarization> report for ',folder)

            datefolder_list = glob.glob('*')
            
            for datefolder in datefolder_list[:]:
                if (datefolder[-5:][(datefolder[-5:].find('.')+1):] in delete_list):
                    datefolder_list.remove(datefolder)
            print(datefolder_list)
            
            accum_time = 0
            df_rateddata = pd.DataFrame(columns=signallist)
            df_idledata = pd.DataFrame(columns=signallist)
            for datefolder in datefolder_list:
                try:
                    csvfolder = sysfolder + datefolder + '\\CSV\\'
                    os.chdir(csvfolder)
                except:
                    print('No CSV file folder found!! ',datefolder) 
                    continue
                df = hebing(csvfolder,signallist,sample_frequency)
                if len(df)>0:
                    df['RunTime'] = df['SPM_I_StkCurr']>=5
                    df['RunTime'] = df['RunTime'].cumsum(0)+accum_time
                    accum_time = max(df['RunTime'])
                else:
                    continue

                df_rateddata = pd.concat([df_rateddata,df[(df['SPM_I_StkCurr']<(rated_curr_set+10))&(df['SPM_I_StkCurr']>(rated_curr_set-10))&(df['CDC_T_TempIn']>(temp_set-2))]],axis =0)
                df_idledata = pd.concat([df_idledata,df[(df['SPM_I_StkCurr']<(idle_curr_set+10))&(df['SPM_I_StkCurr']>(idle_curr_set-10))&(df['CDC_T_TempIn']>(temp_set-2))]],axis =0)
            os.chdir(sysfolder)
            interval = 100
            print(len(df_rateddata))
            print(len(df_idledata))
            if (len(df_rateddata)>0)&(len(df_idledata)>0):
                df_rateddata_processed = average(df_rateddata,interval,'df_rateddata')
                df_idledata_processed = average(df_idledata,interval,'df_idledata')
                with pd.ExcelWriter(Polarization_file) as writer:
                    df_rateddata_processed.to_excel(writer,sheet_name='rateddata_processed')## 1,3
                    df_idledata_processed.to_excel(writer,sheet_name='idledata_processed')## 1,3

            elif (len(df_rateddata)>0)&(len(df_idledata)==0):
                print('There is no idle data!!!')
                df_rateddata_processed = average(df_rateddata,interval,'df_rateddata')
                with pd.ExcelWriter(Polarization_file) as writer:
                    df_rateddata_processed.to_excel(writer,sheet_name='rateddata_processed')## 1,3

            elif (len(df_idledata)>0)&(len(df_rateddata)==0):
                print('There is no rated data!!!')
                df_idledata_processed = average(df_idledata,interval,'df_idledata')
                with pd.ExcelWriter(Polarization_file) as writer:
                    df_idledata_processed.to_excel(writer,sheet_name='idledata_processed')## 1,3
                    
            else:
                print('There is no idle data!!!')
        else:
            print('Results folder<',Polarization_file,'> exists, do not have to run it again.')

    print('All analysis <Polarization> completed, please choose again or Quit GUI!')
    return 0
#######################################################################
def Leakage():
    print('All analysis <Leakage> completed, please choose again or Quit GUI!')
    return 0
#######################################################################
def FuelEconomy():
    print('All analysis <FuelEconomy> completed, please choose again or Quit GUI!')
    return 0
#######################################################################
def DTC():
    print('All analysis <DTC> completed, please choose again or Quit GUI!')
    global df
    signallist = ['SDM_n_Grd1DTC','SDM_n_Grd2DTC','SDM_n_Grd3DTC','SDM_n_Grd4DTC']
    signallist.append(Signal_subst)
    signallist.append(Signal_abstime)
    print()

    
    delete_list = ['txt','xlsx','png','ppt','pptx','csv','docx','rar','pdf','vsdx','xls','xlsm']
    print(seleted_folders)
    for folder in seleted_folders:
        DTC_file = folder + 'DTC.xlsx'
        sysfolder = root_path + '\\' + folder + '\\'
        os.chdir(sysfolder)
        if (os.path.exists(DTC_file)==0) or (DTC_re_run_flag==1) or (Allrun_re_run_flag==1):
            print('Now generate <DTC> report for ',folder)
            datefolder_list = glob.glob('*')


            DTC1grd_total = pd.DataFrame()
            DTC2grd_total = pd.DataFrame()
            DTC3grd_total = pd.DataFrame()
            DTC4grd_total = pd.DataFrame()
            for datefolder in datefolder_list[:]:
                if (datefolder[-5:][(datefolder[-5:].find('.')+1):] in delete_list):
                    datefolder_list.remove(datefolder)
            print(datefolder_list)

            for datefolder in datefolder_list:
                try:
                    csvfolder = sysfolder + datefolder + '\\CSV\\'
                    os.chdir(csvfolder)
                except:
                    print('No CSV file folder found!! ',datefolder) 
                    continue
                df = hebing(csvfolder,signallist,sample_frequency)
                DTC1grd = DTCgrd_gen(df,1)
                DTC2grd = DTCgrd_gen(df,2)
                DTC3grd = DTCgrd_gen(df,3)
                DTC4grd = DTCgrd_gen(df,4)

                DTC1grd_total = pd.concat([DTC1grd_total,DTC1grd],axis=0)
                DTC2grd_total = pd.concat([DTC2grd_total,DTC2grd],axis=0)
                DTC3grd_total = pd.concat([DTC3grd_total,DTC3grd],axis=0)
                DTC4grd_total = pd.concat([DTC4grd_total,DTC4grd],axis=0)
                

            os.chdir(sysfolder)
            with pd.ExcelWriter(DTC_file) as writer:
                DTC1grd_total.to_excel(writer,sheet_name='DTC1grd_total')
                DTC2grd_total.to_excel(writer,sheet_name='DTC2grd_total')
                DTC3grd_total.to_excel(writer,sheet_name='DTC3grd_total')
                DTC4grd_total.to_excel(writer,sheet_name='DTC4grd_total')
                

        else:
            print('Results folder<',DTC_file,'> exists, do not have to run it again.')
    return 0
#######################################################################
def Allrun():
    General()
    EStress()
    MechStress()
    Contamination()
    DryEvents()
    FuelStarve()
    Polarization()
    Leakage()
    FuelEconomy()
    DTC()
    print('------------------------------------------------------------------')
    print('All analysis <Allrun> completed, please choose again or Quit GUI!')
    print('------------------------------------------------------------------')
    return 0
#######################################################################
def start_generation():


    global seleted_folders,sample_frequency,processdata_s,Signal_subst,Signal_CVM_avg,Signal_CVM_std
    global cell_num,Signal_abstime,rated_curr_set,idle_curr_set,temp_set
    if sys_type == "211E-DL3-2-003":
        cell_num = 402
        rated_curr_set = 550
        idle_curr_set = 50
        temp_set = 60
        rawdata_s = 0.1 #s
        processdata_s = 1 #s
        Signal_subst = 'STM_n_SubSt'
        Signal_st = 'STM_n_Sts'
        Signal_currset = 'SPM_I_CurrSet'
        Signal_anode_prs_out = 'HDC_p_PrsOut'
        Signal_abstime = 'AbsTime'
        Signal_CVM_avg = ['Average_1','Average_2']
        Signal_CVM_std = ['StanDev_1','StanDev_2']
        opsfilepath = root_path + "\\211P Operation condition.csv"
    elif sys_type =="211P-DL3-1-001":
        cell_num = 450
        rated_curr_set = 425
        idle_curr_set = 50
        temp_set = 55
        rawdata_s = 0.1 #s
        processdata_s = 1 #s
        Signal_subst = 'STM_n_MainSt'
        Signal_st = 'STM_n_StsForVCU'
        Signal_currset = 'SPM_I_CurrSet'
        Signal_anode_prs_out = 'HDC_p_HRBPrsIn'
        Signal_abstime = 'AbsTime'
        Signal_CVM_avg = ['Average_1','Average_2']
        Signal_CVM_std = ['StanDev_1','StanDev_2']
        opsfilepath = root_path + "\\211P Operation condition.csv"
    elif sys_type == "211P-PP-001":
        cell_num = 450
        rated_curr_set = 425
        idle_curr_set = 50
        temp_set = 70
        rawdata_s = 1 #s
        processdata_s = 1 #s
        Signal_subst = 'STM_n_MainSt'
        Signal_st = 'STM_n_StsForVCU'
        Signal_currset = 'SPM_I_CurrSet'
        Signal_anode_prs_out = 'HDC_p_HRBPrsIn'
        Signal_abstime = 'AbsTime'
        Signal_CVM_avg = ['Average_1','Average_2']
        Signal_CVM_std = ['StanDev_1','StanDev_2']
        opsfilepath = root_path + "\\211P Operation condition.csv"
    elif sys_type == "160P-DL3-2-006":### 160P试验车
        cell_num = 274
        rated_curr_set = 450
        idle_curr_set = 50
        temp_set = 55
        rawdata_s = 1 #s
        processdata_s = 1 #s
        Signal_subst = 'STM_n_MainSt'
        Signal_st = 'STM_n_StsForVCU'
        Signal_currset = 'SPM_I_CurrSet'
        Signal_anode_prs_out = 'HDC_p_HRBPrsIn'
        Signal_abstime = 'AbsTime'
        Signal_CVM_avg = []
        Signal_CVM_std = []
        opsfilepath = root_path + "\\211P Operation condition.csv"
    elif sys_type == "180P-VB-003":
        cell_num = 274
        rated_curr_set = 475
        temp_set = 63.5
        rawdata_s = 0.1 #s
        processdata_s = 1 #s
        Signal_subst = 'STM_n_MainSt'
        Signal_st = 'STM_n_StsForVCU'
        Signal_currset = 'SPM_I_CurrSet'
        Signal_anode_prs_out = 'HDC_p_HRBPrsIn'
        Signal_abstime = 'AbsTime'
        Signal_CVM_avg = []
        Signal_CVM_std = []
        opsfilepath = root_path + "\\211P Operation condition.csv"
    elif sys_type == "211E-VB-003":
        cell_num = 402
        rated_curr_set = 550
        idle_curr_set = 50
        temp_set = 60
        rawdata_s = 0.1 #s
        processdata_s = 0.5 #s
        Signal_subst = 'STM_n_SubSt'
        Signal_st = 'STM_n_Sts'
        Signal_currset = 'SPM_I_CurrSet'
        Signal_anode_prs_out = 'HDC_p_PrsOut'
        Signal_abstime = 'AbsTime'
        Signal_CVM_avg = ['Average_1','Average_2']
        Signal_CVM_std = ['StanDev_1','StanDev_2']
        opsfilepath = root_path + "\\211P Operation condition.csv"
    elif sys_type == "211P-PP-639":
        cell_num = 450
        rated_curr_set = 500
        idle_curr_set = 50
        temp_set = 55
        rawdata_s = 1 #s
        processdata_s = 1 #s
        Signal_subst = 'STM_n_MainSt'
        Signal_st = 'STM_n_StsForVCU'
        Signal_currset = 'SPM_I_CurrSet'
        Signal_anode_prs_out = 'HDC_p_HRBPrsIn'
        Signal_abstime = 'AbsTime'
        Signal_CVM_avg = []
        Signal_CVM_std = []
        opsfilepath = root_path + "\\211P Operation condition.csv"
    sample_frequency = processdata_s/rawdata_s
    os.chdir(root_path)
    ## find the right system file folders
    seleted_folders = folder_select(root_path,sys_type)
#######################################################################

def ini_generation():
    global Allrun_re_run_flag,General_re_run_flag,Estress_re_run_flag,MechStress_re_run_flag,Contamination_re_run_flag,sys_type
    global DryEvents_re_run_flag,FuelStarve_re_run_flag,Polarization_re_run_flag,Leakage_re_run_flag,FuelEconomy_re_run_flag,DTC_re_run_flag
    global Contamination_cvm_flag
    Allrun_re_run_flag = CheckVar4.get()
    General_re_run_flag = CheckVar5.get()
    Estress_re_run_flag = CheckVar6.get()
    MechStress_re_run_flag = CheckVar7.get()
    Contamination_re_run_flag = CheckVar8.get()
    Contamination_cvm_flag = CheckVar81.get()
    print(Contamination_cvm_flag)
    DryEvents_re_run_flag = CheckVar9.get()
    FuelStarve_re_run_flag = CheckVar10.get()
    Polarization_re_run_flag = CheckVar11.get()
    Leakage_re_run_flag = CheckVar12.get()
    FuelEconomy_re_run_flag = CheckVar13.get()
    DTC_re_run_flag = CheckVar14.get()

    

    sys_type = comboxlist.get()
    print('Now you want to process:',sys_type)

    start_generation()

    return 0

#######################################################################
def open_file():
    '''
    打开文件
    :return:
    '''

    global root_path

    root_path = filedialog.askdirectory(title=u'选择路径', initialdir='D:\\2.DUR&EMS\\2.DURDAILYREPORT')
    root_path = root_path+'\\'
    print('选择的数据路径为：', root_path)
#######################################################################
if __name__ == '__main__':
    root_path = 'D:\\2.DUR&EMS\\2.DURDAILYREPORT\\'
    os.chdir(root_path)


    ## windows loop
    window = tk.Tk()
    window.title('DurabilityDataProcessGUI')  # 标题
    window.geometry('500x1000')  # 窗口尺寸

    comvalue=tk.StringVar()#窗体自带的文本，新建一个值
    comboxlist=ttk.Combobox(window,textvariable=comvalue) #初始化
    comboxlist["values"]=("211P-DL3-1-001","211E-DL3-2-003","211P-PP-001","160P-DL3-2-006","180P-VB-003","211E-VB-003","211P-PP-639")
    comboxlist.current(2) #选择第一个
    comboxlist.grid(row=1,column=0)
    sys_type = comboxlist.get()

##    sys_type = "211P-DL3-1-001"

##-------------------------------------------------------------------------------------------------
    button2 = tk.Button(window, text='选择路径', width=20, height=3, command=open_file)
    button2.grid(row=2,column=0)
##-------------------------------------------------------------------------------------------------
    button3 = tk.Button(window, text='确认选择,务必点', width=20, height=3, command=ini_generation)
    button3.grid(row=3,column=0)
##-------------------------------------------------------------------------------------------------
    button4 = tk.Button(window, text='Allrun', width=20, height=3, command=Allrun)
    button4.grid(row=4,column=0)

    CheckVar4 = tk.IntVar()
    checkbutton4 = tk.Checkbutton(window, text = "历史数据重跑?", variable = CheckVar4, \
                 onvalue = 1, offvalue = 0, height=3, \
                 width = 20)
    checkbutton4.grid(row=4,column=1)

##-------------------------------------------------------------------------------------------------
    button5 = tk.Button(window, text='General', width=20, height=3, command=General)
    button5.grid(row=5,column=0)

    CheckVar5 = tk.IntVar()
    checkbutton5 = tk.Checkbutton(window, text = "历史数据重跑?", variable = CheckVar5, \
                 onvalue = 1, offvalue = 0, height=3, \
                 width = 20)
    checkbutton5.grid(row=5,column=1)
##-------------------------------------------------------------------------------------------------
    button6 = tk.Button(window, text='Electrochemical Stress', width=20, height=3, command=EStress)
    button6.grid(row=6,column=0)

    CheckVar6 = tk.IntVar()
    checkbutton6 = tk.Checkbutton(window, text = "历史数据重跑?", variable = CheckVar6, \
                 onvalue = 1, offvalue = 0, height=3, \
                 width = 20)
    checkbutton6.grid(row=6,column=1)
##-------------------------------------------------------------------------------------------------
    button7 = tk.Button(window, text='Mechanical Stress', width=20, height=3, command=MechStress)
    button7.grid(row=7,column=0)

    CheckVar7 = tk.IntVar()
    checkbutton7 = tk.Checkbutton(window, text = "历史数据重跑?", variable = CheckVar7, \
                 onvalue = 1, offvalue = 0, height=3, \
                 width = 20)
    checkbutton7.grid(row=7,column=1)
##-------------------------------------------------------------------------------------------------
    button8 = tk.Button(window, text='Contamination Stress', width=20, height=3, command=Contamination)
    button8.grid(row=8,column=0)

    CheckVar8 = tk.IntVar()
    checkbutton8 = tk.Checkbutton(window, text = "历史数据重跑?", variable = CheckVar8, \
                 onvalue = 1, offvalue = 0, height=3, \
                 width = 20)
    checkbutton8.grid(row=8,column=1)
    CheckVar81 = tk.IntVar()
    checkbutton81 = tk.Checkbutton(window, text = "CVM数据输出?", variable = CheckVar81, \
                 onvalue = 1, offvalue = 0, height=3, \
                 width = 20)
    checkbutton81.grid(row=8,column=2)
##-------------------------------------------------------------------------------------------------
    button9 = tk.Button(window, text='Drying Events', width=20, height=3, command=DryEvents)
    button9.grid(row=9,column=0)

    CheckVar9 = tk.IntVar()
    checkbutton9 = tk.Checkbutton(window, text = "历史数据重跑?", variable = CheckVar9, \
                 onvalue = 1, offvalue = 0, height=3, \
                 width = 20)
    checkbutton9.grid(row=9,column=1)
##-------------------------------------------------------------------------------------------------
    button10 = tk.Button(window, text='Fuel Starvation events', width=20, height=3, command=FuelStarve)
    button10.grid(row=10,column=0)

    CheckVar10 = tk.IntVar()
    checkbutton10 = tk.Checkbutton(window, text = "历史数据重跑?", variable = CheckVar10, \
                 onvalue = 1, offvalue = 0, height=3, \
                 width = 20)
    checkbutton10.grid(row=10,column=1)
##-------------------------------------------------------------------------------------------------
    button11 = tk.Button(window, text='Polarization', width=20, height=3, command=Polarization)
    button11.grid(row=11,column=0)

    CheckVar11 = tk.IntVar()
    checkbutton11 = tk.Checkbutton(window, text = "历史数据重跑?", variable = CheckVar11, \
                 onvalue = 1, offvalue = 0, height=3, \
                 width = 20)
    checkbutton11.grid(row=11,column=1)
##-------------------------------------------------------------------------------------------------
    button12 = tk.Button(window, text='Leakage', width=20, height=3, command=Leakage)
    button12.grid(row=12,column=0)

    CheckVar12 = tk.IntVar()
    checkbutton12 = tk.Checkbutton(window, text = "历史数据重跑?", variable = CheckVar12, \
                 onvalue = 1, offvalue = 0, height=3, \
                 width = 20)
    checkbutton12.grid(row=12,column=1)
##-------------------------------------------------------------------------------------------------
    button13 = tk.Button(window, text='Fuel Economy', width=20, height=3, command=FuelEconomy)
    button13.grid(row=13,column=0)

    CheckVar13 = tk.IntVar()
    checkbutton13 = tk.Checkbutton(window, text = "历史数据重跑?", variable = CheckVar13, \
                 onvalue = 1, offvalue = 0, height=3, \
                 width = 20)
    checkbutton13.grid(row=13,column=1)
##-------------------------------------------------------------------------------------------------
    button14 = tk.Button(window, text='DTC', width=20, height=3, command=DTC)
    button14.grid(row=14,column=0)

    CheckVar14 = tk.IntVar()
    checkbutton14 = tk.Checkbutton(window, text = "历史数据重跑?", variable = CheckVar14, \
                 onvalue = 1, offvalue = 0, height=3, \
                 width = 20)
    checkbutton14.grid(row=14,column=1)
##-------------------------------------------------------------------------------------------------    
##    signallist = ['STM_n_MainSt','SPM_I_StkCurr','SPM_U_StkVolt','UDS_Id_CurrDensitySet']
##    f = hebing(root_path,signallist)# 3.contentate all the data to on file
    window.mainloop()  # 显示
    os.chdir(root_path)
