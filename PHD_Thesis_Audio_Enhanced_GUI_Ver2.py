#◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄► in the name of Allah the compassionate the merciful  ◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►
#◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄► FOREX metamodel GUI Setup to work on several strategies ◄►◄►◄►◄►◄►◄►◄►◄►◄►
#◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►   import libraries   ◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►  
import pandas as pd
import numpy as np
import tkinter as tk
import MetaTrader5 as mt5
import os
import datetime
import matplotlib.pyplot as plt
import pyttsx3
import sounddevice as sd
import wavio as wv
import speech_recognition as sr
import pytz
import winsound
import subprocess
#◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄► import from libraries ◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►  
from tkinter import ttk
from tkinter import messagebox,filedialog
from PyQt5.QtWidgets import QWidget
from seaborn import heatmap 
from scipy.io.wavfile import write
from sys import path,argv,executable,exit
from sklearn.preprocessing import MinMaxScaler , StandardScaler , RobustScaler
from statistics import stdev
from screeninfo import get_monitors
from openpyxl import load_workbook
from openpyxl.styles import Alignment
#◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄► import self-made libraries ◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►  
path.append('./Custom-libs')
import ALARM
import customized
#◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄► Visual and displayin settings ◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄► 
pd.set_option('display.max_rows', 200)
pd.set_option('display.max_columns', 80)
#◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄► initializing important Dictionaries and look-up tables ◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►  
TimeFrameDict1 = {'m1': mt5.TIMEFRAME_M1,
                 'm2': mt5.TIMEFRAME_M2,
                 'm3': mt5.TIMEFRAME_M3,
                 'm4': mt5.TIMEFRAME_M4,
                 'm5': mt5.TIMEFRAME_M5,
                 'm6': mt5.TIMEFRAME_M6,
                 'm10': mt5.TIMEFRAME_M10,
                 'm12': mt5.TIMEFRAME_M12,
                 'm15': mt5.TIMEFRAME_M15,
                 'm20': mt5.TIMEFRAME_M20,
                 'm30': mt5.TIMEFRAME_M30,                 
                 'h1': mt5.TIMEFRAME_H1,
                 'h2': mt5.TIMEFRAME_H2,
                 'h3': mt5.TIMEFRAME_H3,
                 'h4': mt5.TIMEFRAME_H4,
                 'h6': mt5.TIMEFRAME_H6,
                 'h8': mt5.TIMEFRAME_H8,
                 'h12': mt5.TIMEFRAME_H12,
                 'd1': mt5.TIMEFRAME_D1,
                 'd5' : '',
                 'w1': mt5.TIMEFRAME_W1,
                 'mn1': mt5.TIMEFRAME_MN1}

TimeFrameDict2 = {'m1': "1m",
                  'm2': "2m",
                  'm5': "5m",
                  'm15': "15m",
                  'm30': "30m",                 
                 'h1': "1h",
                 'd1': "1d",
                 'd5': "5d",
                 'w1': "1wk",
                 'mn1': "1mo"}

TradeErrorDict = {10004: 'Requote',
10006 : 'Request rejected',
10007 : 'Request canceled by trader',
10008 : 'Order placed',
10009 : 'Request completed',
10010 : 'Only part of the request was completed',
10011 : 'Request processing error',
10012 : 'Request canceled by timeout',
10013 : 'Invalid request',
10014 : 'Invalid volume in the request',
10015 : 'Invalid price in the request',
10016 : 'Invalid stops in the request',
10017 : 'Trade is disabled',
10018 : 'Market is closed',
10019 : 'There is not enough money to complete the request',
10020 : 'Prices changed',
10021 : 'There are no quotes to process the request',
10022 : 'Invalid order expiration date in the request',
10023 : 'Order state changed',
10024 : 'Too frequent requests',
10025 : 'No changes in request',
10026 : 'Autotrading disabled by server',
10027 : 'Autotrading disabled by client terminal',
10028 : 'Request locked for processing',
10029 : 'Order or position frozen',
10030 : 'Invalid order filling type',
10031 : 'No connection with the trade server',
10032 : 'Operation is allowed only for live accounts',
10033 : 'The number of pending orders has reached the limit',
10034 : 'The volume of orders and positions for the symbol has reached the limit',
10035 : 'Incorrect or prohibited order type',
10036 : 'Position with the specified POSITION_IDENTIFIER has already been closed',
10038 : 'A close volume exceeds the current position volume',
10039 : 'A close order already exists for a specified position. This may happen when working in the hedging system:',
10040 : 'The number of open positions simultaneously present on an account can be limited by the server settings.',
10041 : 'The pending order activation request is rejected, the order is canceled',
10042 : 'The request is rejected, because the "Only long positions are allowed" rule is set for the symbol (POSITION_TYPE_BUY)',
10043 : 'The request is rejected, because the "Only short positions are allowed" rule is set for the symbol (POSITION_TYPE_SELL)',
10044 : 'The request is rejected, because the "Only position closing is allowed" rule is set for the symbol',
10045 : 'The request is rejected, because "Position closing is allowed only by FIFO rule" flag is set for the trading account (ACCOUNT_FIFO_CLOSE=true)',
10046 : 'The request is rejected, because the "Opposite positions on a single symbol are disabled" rule is set for the trading account.'
} 
#◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄► initializing important general variables ◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►  
SingleCandlePredict = False
MusicPath = r'C:\Windows\Media\Windows Proximity Notification'
FONTSIZE = 10
FONTSIZE2 = 16
Symbol = 'EURUSD'
BALANCE = 1000
msg  = ''
Msg  = ''
LABL = ''
FEAT = ''
FstrCateg = ''
PresetData = ''
MDD = ''
MDDP = ''
sharpe_ratio = ''
prediction_data2 = ''
temp1 = ''
temp2 = ''
TableFile_path = ''
S1 = 'without scaling'
TimeFrame = 'd1'
TestSize = 30
Trainsize = TestSize*10
Sdate = (2018,1,1)
Edate = (2025,7,5)
Scaling_Type = 0
LOGIN = 5037819659 #52396016   #52278504
PASS = '-v1fIxSn' #"_a4xZdMl"  #"Amirolmomenin@1"
SERVER = "MetaQuotes-Demo"
C_Ensemble = False
R_Ensemble = False
AlreadyScaled = False
QuickRun = False
EvaluationFault = False
OnetimeSetting = False
LagTest = ''
ComparisonFlag = ''
comparisondf = pd.DataFrame()
labels = np.array(())
features = pd.DataFrame()
y_test = ''
y_pred_test = ''
FigureName = ''
audiointeract = True
selected_timezone = "Tehran"
ModelStorm = False
HomogeneousEnsemble = False
Strategy = 'No Strategy'
Source = ''
LoginFlag1 = False
LoginFlag2 = False
Concise_Report = False
Fcateg = ['Custom Features','Custom Window Features','OHLC Features','Just previous Closes','Just previous Labels','All Features','All features except labels','Add Binary indicators','Just Binary indicators','Selected Binary indicators','Just mean Delta','Just Mean Return Indicators'] 
CompareFeaturessoptions = Fcateg[:-5]
selecteditems = []
tempindex = []
DynamicColumns = []
Accuracy_lst = []
F1score_lst = []
precision_lst = []
recall_lst = []
Bal_Accuracy_test_lst = []
Logloss_lst = []
ROCAUC_lst = []
R2score_lst = []
MDDP_lst = []
TotalNetProfit_lst = []
Profit_ratio_lst = []
sharpe_ratio_lst = []
model_no = []
MODELSSTORM = []
Y_PRED_Storm_TEST  = np.array([],dtype=int)
Volume = 1
freq = 44100
windowsize = 5
ScreenWidth = int(get_monitors()[0].width/130)
ScreenHeight = int(get_monitors()[0].height/130)
#TrueFalsedict={True:'Yes',False:'No'}
engine = pyttsx3.init()
engine.setProperty('voice', engine.getProperty('voices')[1].id)
engine.setProperty('rate', 150)
engine = pyttsx3.init()
engine.setProperty('voice', engine.getProperty('voices')[1].id)
engine.setProperty('rate', 150)
BT_text = 'Open Results File'
DataPath = os.getcwd()+r'\Data'
File_path = os.getcwd()+r'\Simulation Results'

#◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄► Functions definition ◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►  
def VoiceAnnounce(TXT):
    if audiointeract:
        engine.say(TXT)
        engine.runAndWait()

def Total_Reset():
    global features
    global Prepared_features
    global Strategy
    #global Lottage
    global labels
    global df
    global msg
    global Msg
    global LABL
    global FEAT
    global S1
    global TimeFrame
    global TestSize
    msg  = ''
    Msg  = ''
    LABL = ''
    FEAT = ''
    df = ''
    Strategy = 'No Strategy'
    features = pd.DataFrame(np.array(()))
    Prepared_features = pd.DataFrame(np.array(()))
    labels = np.array(())
    S1 = 'without scaling'
    TimeFrame = 'd1'
    TestSize = 30
    Trainsize = TestSize*10
    Scaling_Type = 0
    Fea_MESSAGE1.configure(text = f'Not selected yet', fg='red')
    LB_MESSAGE1.configure(text = f'label: Not selected yet', fg='red')

def open_position_symbols():
    symbols = tuple(set(p.symbol for p in mt5.positions_get()))
    WIN9 = tk.Toplevel(root)
    WIN9.bind('<Escape>',EXIT_One)
    tk.Message(WIN9,text=symbols).pack()

def net_Volume(symbol):
    return sum(p.volume if p.type==mt5.POSITION_TYPE_BUY else -p.focus
        for p in mt5.positions_get(symbol=symbol))

def net_profit(symbol):
    return sum(p.profit for p in mt5.positions_get(symbol=symbol))

def verbose_status():
    positionstext =''
    for p in mt5.positions_get():
        positionstext += '{} - Net position = {}, Net profit = {}'.format(p.symbol,net_Volume(p.symbol),net_profit(p.symbol))
    print(f'positionstext:{positionstext}')
    WIN10 = tk.Toplevel(root)
    WIN10.bind('<Escape>',EXIT_One)
    Position_Message = tk.Message(WIN10,text=positionstext)
    Position_Message.pack()
    
def Process_Summarize(m=True):
    global msg
    try:
        msg =(lambda X: f'on symbol : {Symbol} \n timeframe : {TimeFrame}\nStart from : {Sdate}\nEnded in : {Edate}' if X else 'Dataloaded from Saved data for time and timeframe')(LoginFlag1)+f''' with {'Online database' if LoginFlag1 else 'Offline database'} 
Symbol & TimeFrame : {Symbol} & {TimeFrame}
Train size: {Trainsize}
Test size: {TestSize}
train_test_Split method: {"NovelSplit" if SingleCandlePredict else "NormalSplit"}
Label: {LABL}
{S1} on features: {features.columns.tolist()}
Model: {Strategy}
'''    
    except NameError as ERR:
        showmessage('Empty fields',f'Error‼ All values should be entered\n{ERR}',TIMEOUT=3000,TYP='Erro')        
    else:
        if m:
            #messagebox.showinfo(title='summary', message = msg)
            showmessage('Summary',msg,TIMEOUT=8000,TYP='info')

def HELP():
    HELPTXT='FOREX metamodel GUI\nworking on mutiple strategies\nInitialize using menues and select the choices and run to get results'
    #messagebox.showinfo(title='Help', message = HELPTXT)
    showmessage('Help',HELPTXT,TIMEOUT=6000,TYP='info')
    
def ABOUT():
    aboutTXT= '''FOREX Metamodel Graphical and Audio User Interface
A setup for a novel research on 2022-2025
created by: Maysam Yazdanpanahi
University of Shahrood'''
    showmessage('About',aboutTXT,TIMEOUT=6000,TYP='info')

def fillcurpair(_):
    global LoginFlag1
    global Symbol   
    Symbol = CURPAIR.get().upper() 
    if Symbol == '':
        Symbol = 'EURUSD'
    else:
        if LoginFlag1:
            Symbol = mt5.symbols_get(Symbol+'*')[0].name
            print(f'Symbol {Symbol} selected')
            print(mt5.symbol_info(Symbol).description)                        
    CURPAIR.delete(0,'end')   
    CURPAIR.insert(0,Symbol)    

def S_OUT(_):
    global Sdate0
    global Sdate
    Sdate0 = SDATE.get()
    if Sdate0 == '':
        SDATE.insert(0,'2018-01-01')
        Sdate0 = '2018-01-01'
#     Sdate = tuple(map(int,Sdate0.split('-')))
#     Sdate = datetime.datetime(Sdate[0],Sdate[1],Sdate[2])
#     Replaced by following 2 lines snippet!
    Sdate = Sdate0.split('-')
    Sdate = datetime.datetime(int(Sdate[0]),int(Sdate[1]),int(Sdate[2][:2]))
    if _:
        Sdatebutton.config(text="Start date",fg='gold')
    print('Start time is:',Sdate)
        
# def E_CLICK(_):
#     EDATE.delete(0,'end')
def End_Date():
    global Edate
    Edate = datetime.datetime(2025,7,5)
    EDATE.delete(0, tk.END)
    EDATE.insert(0,Edate)    
    print('End time is:',Edate)
    
def E_OUT(_):
    #global LoginFlag1
    #global LoginFlag2
    global Edate0
    global Edate
    Edate0 = EDATE.get()
    if Edate0 == '':
        if LoginFlag1 or LoginFlag2:
            EDATE.delete(0, tk.END)
            Edate = datetime.datetime.now() 
            
        else:
            Edate0 = '2025-7-5'
            EDATE.delete(0, tk.END)
            #EDATE.insert(0,Edate)     
#             Edate = tuple(map(int,(Edate0).split('-')))
#             Edate = datetime.datetime(Edate[0],Edate[1],Edate[2])
#             Replaced by following 2 lines snippet!
            Edate = Edate0.split('-')
            Edate = datetime.datetime(int(Edate[0]),int(Edate[1]),int(Edate[2][:2]))
            
    else:
        Edate0 = str(Edate0).split()[0]
#         Edate = tuple(map(int,(Edate0).split('-')))
#         Edate = datetime.datetime(Edate[0],Edate[1],Edate[2])
#             Replaced by following line snippet!
        Edate = Edate0.split('-')
        Edate = datetime.datetime(int(Edate[0]),int(Edate[1]),int(Edate[2][:2]))
        EDATE.delete(0, tk.END)
        
    EDATE.insert(0,Edate)
    print('End time is:',Edate)
           
def EXIT_One(_):  
    for widget in root.winfo_children():
        if isinstance(widget,tk.Toplevel) :#or isinstance(widget,messagebox) :
            widget.destroy()
            break
    else:
        win = tk.Tk()
        win.withdraw()
        win.option_add('*Dialog.msg.font', 'Helvetica 20')
        win.after(2000,win.destroy)
        
        response = messagebox.askyesno("Exit","Do you want to quit?",default=messagebox.NO,master=win)
        if response:
            root.destroy()
            quit()

def EXIT_All():
    for widget in root.winfo_children():
        if isinstance(widget,tk.Toplevel):
            widget.destroy()
            
def CLOSEexcellFiles():
    import win32com.client
    # Create an instance of the Excel application
    excel_app = win32com.client.Dispatch("Excel.Application")
    # Loop through all open workbooks and close them
    for workbook in excel_app.Workbooks:
        workbook.Close(SaveChanges=False)  # Set SaveChanges to True if you want to save changes
    excel_app.Quit()
    print("All Excel files have been closed.")
#◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄► All the following RESTART functions have one goal and are written for test◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►              
def RESTART():
    print("argv was:",argv)
    print("sys.executable was:", executable)
    print("Restarting now...")
    root.destroy()
    root.quit()
    os.execv(sys.executable, ['python']+sys.argv)
    exit()
def RESTART1():
    os.system('cls')
    script_name = os.path.basename(__file__)
    os.system(script_name)   
def RESTART2():
    python = sys.executable
    os.execl(python, python, *argv)
    exit()
def RESTART3():
    os.execl(sys.executable, 'python', __file__, *argv)
    exit()
                   
def DEFAULTSETTINGS():
    global TestSize
    global Symbol
    Symbol = 'EURUSD'
    TestSize = 30
    fillcurpair()
    
def SETTING1():
    def Set_rowcol(event='none'):
        nonlocal Entry1
        nonlocal Entry2
        try:
            MaxRow = int(Entry1.get())
            MaxCol = int(Entry2.get())
            pd.set_option('display.max_rows', MaxRow) #pd.set_option('display.max_rows',None)
            pd.set_option('display.max_columns', MaxCol)
            #Entry1.delete(0, tk.END)
            #Entry2.delete(0, tk.END)
            Maximumdimention = f"MaxRows: {MaxRow}\nMaxCols: {MaxCol}"
            showmessage('info',Maximumdimention,TIMEOUT=6000,TYP='info')
            print(Maximumdimention)
            WIN6.withdraw()
        except:
            showmessage("Error",f"Not valid values!",TIMEOUT=3000,TYP='Erro')

    
    WIN6 = tk.Toplevel(root)
    WIN6.bind('<Escape>',EXIT_One)
    WIN6.bind('<Return>',Set_rowcol)
    MaxRow = tk.StringVar()
    MaxCol = tk.StringVar()
    MXR = tk.Label(WIN6,text="Max. Rows: ",bg="#D7BDE2",font=("Georgia",FONTSIZE2,'bold'))
    MXR.pack()

    Entry1 = tk.Entry(WIN6,textvariable=MaxRow,width=15,borderwidth=3)
    Entry1.pack()
    Entry1.focus()
    
    MXC = tk.Label(WIN6,text="Max. Columns: ",bg="#D7BDE2",font=("Georgia",FONTSIZE2,'bold'))
    MXC.pack()
        
    Entry2 = tk.Entry(WIN6,textvariable=MaxCol,width=15,borderwidth=3)
    Entry2.pack()

    WIN6.title("Row & Columns")
    WIN6.configure(bg="#F10F0F")
    bt1 = tk.Button(WIN6,text="Set",width=10,command=Set_rowcol,bg="#D7BD00",font=("Georgia",FONTSIZE2))
    bt1.pack()
    WIN6.mainloop()

def CopySettings(): 
    try:
        SaveDict={'Symbol':Symbol, 'Lottage':Lottage, 'TimeFrame':TimeFrame, 'Strategy':Strategy, 'FEAT':FEAT, 'LABL':LABL, 'Scaling_Type':Scaling_Type, 'LoginFlag1':LoginFlag1, 'LoginFlag2':LoginFlag2, 'TestSize':TestSize, 'Trainsize':Trainsize,'SingleCandlePredict':SingleCandlePredict}
      
        if LoginFlag1 or LoginFlag2:
           SaveDict['Sdate'] = Sdate
           SaveDict['Edate'] = Edate
        if len(Strategy.split())>2:
            Short_Strategy = ''.join(x[0] for x in Strategy.split())
        copypath = DataPath+f'\\CopiedSettings\\{Strategy}'
        if not os.path.exists(copypath):
            showmessage('Finding path',f'The path {copypath} was not found!\nIt will be created!',TIMEOUT=2500,TYP='warn')
            os.makedirs(copypath)

        FileNameString = str(features.columns.values).replace(' ','_').replace("'",'').replace(')','').replace('(','').replace(']','◘').replace('[','◘')
        if len(FileNameString.split('_'))>2:
            FileNameString = FileNameString.split('_')[0]+'→'+FileNameString.split('_')[-1]

        Short_Strategy = ''.join(x[0] for x in Strategy.split())
        FileNameString = f'PresetData_NetProfit►{int(df.tail(1)["equity"].values[0]- BALANCE)}$_{Short_Strategy}♦TF_{TimeFrame}♦TestSize_{TestSize}♦Trainsize_{Trainsize}_Feature→{FEAT}_{FileNameString}♦Lable_{LABL}♦{"NovelSplit" if SingleCandlePredict else "NormalSplit"}'
        File_path = filedialog.asksaveasfilename(initialdir=copypath,defaultextension=".txt", filetypes=[("Text files", "*.txt"), ("All files", "*.*")],initialfile= FileNameString)
                 
    except PermissionError as ERR:
        showmessage('Permission Error',f'Check if file is already opened!\n{ERR}',TIMEOUT=3000,TYP='Erro')
    
    except NameError as ERR:
        showmessage('Data Shortage',f'Data is incomplete!\n{ERR}',TIMEOUT=3000,TYP='Erro')
  
    else:
        with open(File_path,'w') as PresetFile:
            print(SaveDict,file=PresetFile)
        file_name = File_path.split('.')[0]
        features.to_csv(file_name+r'_features'+'.csv')
        df.to_csv(file_name+r'_df'+'.csv')
        labels.to_csv(file_name+r'_labels'+'.csv')
        VoiceAnnounce('Preset data saved')
        showmessage('Save',f"Preset files saved successfully!\n@ {File_path}",TIMEOUT=4500,TYP='info')
        print(f"Preset files saved successfully!\n@ {File_path}")
        print(f'saved TimeFrame:{TimeFrame}')
        
def LoadSettings():
    global LABL
    global FEAT
    global Symbol
    global Lottage
    global TimeFrame
    global Strategy
    global labels
    global features
    global df
    global TestSize
    global Trainsize
    global Scaling_Type
    global LoginFlag1
    global LoginFlag2
    global Sdate
    global Edate
    global SingleCandlePredict
    
    copypath = DataPath+r'\CopiedSettings'
    try:
        file_path = filedialog.askopenfilename(initialdir=copypath,defaultextension=".txt", filetypes=[("Text files", "*.txt"),("All files", "*.*")], initialfile = "Text files")

        with open(file_path,'r') as PresetFile:
            PresetData = PresetFile.read()
        print('file_path:',file_path)
        file_name = file_path.split('.')[0]
    
        features= pd.read_csv(file_name+r'_features'+'.csv')
        features.index= features.DATE
        features.drop('DATE', axis=1, inplace = True)
        df = pd.read_csv(file_name+r'_df'+'.csv')
        df.index= df.DATE
        df.drop('DATE', axis=1, inplace = True)
        labels = pd.read_csv(file_name+r'_labels'+'.csv')
        labels.index= labels.DATE
        labels.drop('DATE', axis=1, inplace = True)

        PresetData = eval(PresetData)
        FEAT = PresetData['FEAT']
        LABL = PresetData['LABL']
        Symbol = PresetData['Symbol']
        Lottage = PresetData['Lottage']
        TimeFrame = PresetData['TimeFrame']
        Strategy = PresetData['Strategy']
        TestSize = PresetData['TestSize']
        Trainsize = PresetData['Trainsize']
        Scaling_Type = PresetData['Scaling_Type']
        LoginFlag1 = PresetData['LoginFlag1']
        LoginFlag2 = PresetData['LoginFlag2']
        SingleCandlePredict = PresetData['SingleCandlePredict']
        if LoginFlag1 or LoginFlag2:
            Sdate = PresetData['Sdate']
            Edate = PresetData['Edate']
#             except TypeError as ERR:
            ERR = 'Dataloaded is from onlinesource\nDon\'t forget to get online\nError in loading Sdate and Edate on line}'
            showmessage('onlinedata history',ERR,TIMEOUT=3000,TYP='warn')
            print(ERR)
            
    except PermissionError as ERR:
        showmessage('Permission Error',f'Check if file is already opened!\n{ERR}',TIMEOUT=3000,TYP='Erro')

    except FileNotFoundError as ERR:
        showmessage('FileNotFound Error',f'Create the presetting file first!\n{ERR}',TIMEOUT=3000,TYP='Erro')
        
    else:
        VoiceAnnounce('Preset data loaded.')
        showmessage('Preset Data Loaded',PresetData,TIMEOUT=10000,TYP='info')
        root.bind('<Return>', RUNStrategy_)
        DataBaseBTN.config(text = 'Loaded with\nPreset Data',fg='light green',font=('Times',FONTSIZE, 'bold'))
        MODLBTN.config(text = Strategy+' Selected',fg='yellow',font='normal',bg='mediumvioletred')
        ModelVAR.set(Strategy)
        RUNBTN.configure(text=f"RUN {Strategy}\n with test size:{TestSize} and train size:{Trainsize}\nTimeFrame:{TimeFrame}", fg='chartreuse',font=('Times', FONTSIZE, 'normal'))
        print('Timeframe=',TimeFrame)
        cmbCategories.set(TimeFrame)
        TF_MESSAGE.config(text = f'{TimeFrame} Timeframe is selected')
        Fea_MESSAGE1.configure(text = f'{FEAT}',fg='#F5CC21')
        cmbFeat.set(FEAT)
        LB_MESSAGE1.configure(text = f'Selected label: {LABL}',fg='#F5CC21')
        cmbLab.set(LABL)
#         chkChoice3.configure(text= ("Novel Split" if SingleCandlePredict else "Normal Split"))
        intChoice3.set(1 if SingleCandlePredict else 0)
        intChoice6.set(0 if SingleCandlePredict else 1)
        
def Instant_Trade_PYQT():
    from PyQt5.QtWidgets import QApplication
    app = QApplication(argv)
    ex = TradingBot()
    ex.show()
    exit(app.exec_())   
    
def Instant_Trade():
    global WIN1
    global Volume
    global Symbol
    global LoginFlag1
    Trade = 1
    def choice():
        nonlocal Trade
        TradeType = strChoice.get()
        lblSelect.config(text=f'Manual {TradeType} trading!')
        if TradeType == 'Buy':
            Trade = 1
        elif TradeType == 'Sell':
            Trade = 0
        else:
            Trade = -1
            
    def DoTrade(_):
        global WIN1
        global VOL
        global LoginFlag1
        global Volume
        Volume = float(VOL_entry.get())
        if Trade == 1:         
            Instant_Buy()
        elif Trade == 0:
            Instant_Sell()
        WIN1.destroy()
        WIN1.quit()
        
    WIN1 = tk.Toplevel(root)
    WIN1.bind('<Escape>',EXIT_One)
    #WIN1.configure(bg='#091A32')
    pho = tk.PhotoImage(file=r".\pictures\BUY_SELL.png")
    w = pho.width()
    h = pho.height()
    WIN1.geometry('%dx%d'%(w,h))
    tk.Label(WIN1,image=pho).place(height=h,width=w,x=0,y=0)#.grid(rowspan=8,columnspan=2)
    VOL  = tk.StringVar()
    strChoice = tk.StringVar()
    
    lblSelect = tk.Label(WIN1,relief=tk.RAISED,bg='#474747',fg='yellow',font=('times',FONTSIZE2,'bold'))
    lblSelect.grid(row=0,columnspan=2,pady=5)#.place(height=20,width=200,x=50,y=150)

    RadioBuy = tk.Radiobutton (WIN1, text='Buy',bg='#474747',fg='yellow',font=('times',FONTSIZE,'bold'), variable=strChoice,selectcolor= 'black',activeforeground='black',value='Buy', command=choice)
    RadioBuy.grid(row=1,column=0,pady=5)#.place(height=20,width=250,x=60,y=50)

    RadioSell = tk.Radiobutton(WIN1, text='Sell', variable=strChoice,bg='#474747',fg='yellow',font=('times',FONTSIZE,'bold'),selectcolor= 'black',activeforeground='black',value='Sell', command=choice)
    RadioSell.grid(row=1,column=1,pady=5)#.place(height=20,width=250,x=160,y=80)

    strChoice.set("Buy")
    choice()
    lblChoice = tk.Label(WIN1,text="select trading Type ",bg='#474747',fg='yellow',font=('times',FONTSIZE2,'bold'))
    lblChoice.grid(row=2,columnspan = 2,pady=5)   #,sticky=tk.W 
    VOLLBL = tk.Label(WIN1,text="Volume: ",bg='#474747',fg='white',font=('times',FONTSIZE2,'bold'))
    VOLLBL.grid(row=4,column=0,pady=5) #,sticky='W'

    VOL_entry = tk.Entry(WIN1,textvariable=VOL,width=10,bg='#474747',fg='yellow',font=('times',FONTSIZE2,'bold'),borderwidth=3,bd=5)
    VOL_entry.grid(row=4,column=1,pady=5)#,sticky='W'

    WIN1.title("Instant Trade")
    btTR = tk.Button(WIN1,text="Do Instant Trade!",bg='#474747',fg='yellow',font=('times',FONTSIZE2,'bold'),relief=tk.RAISED,bd=5,command=lambda:DoTrade(''))
    WIN1.bind('<Return>',DoTrade)
    btTR.grid(row=5,columnspan=2)
    WIN1.mainloop()

def Instant_Buy():
    global Volume
    global LoginFlag1
    global Symbol
    if not LoginFlag1:
        Login_To_MT5()    
    if LoginFlag1:
        symbol_info = mt5.symbol_info(Symbol)
        if symbol_info is None:
            showmessage('Symbol Error',f'Symbol Error!\n{Symbol} not recognized!',TIMEOUT=3000,TYP='Erro')
            return
        # Ensure the symbol is visible
        if not symbol_info.visible:
            if not mt5.symbol_select(Symbol, True):
                showmessage('Symbol Error', f'Failed to select symbol {Symbol}',TIMEOUT=3000,TYP='Erro')
                VoiceAnnounce(f'Failed to select symbol {Symbol}.')
                return
        # Get the price and point for the symbol       
        point = symbol_info.point
        price = mt5.symbol_info_tick(Symbol).ask       
        if price is None:
            showmessage('Price catch Error', f'Failed to catch the price for the symbol {Symbol}',TIMEOUT=3000,TYP='Erro')
            return        
#         result = mt5.order_send(symbol=Symbol, trade_operation=mt5.ORDER_TYPE_BUY, volume=Volume, slippage=3,deviation=20, type_filling=mt5.ORDER_FILLING_IOC)        
        request = {
            "action": mt5.TRADE_ACTION_DEAL,
            "symbol": Symbol,
            "volume": float(Volume),
            "type": mt5.ORDER_TYPE_BUY,
            "price": price,
            "sl": price - 100 * point,
            "tp": price + 100 * point,
            "deviation": 20,
            "magic": 234000,
            "comment": 'Buy from python script',
            "type_time": mt5.ORDER_TIME_GTC,
            "type_filling": mt5.ORDER_FILLING_FOK,  #mt5.ORDER_FILLING_IOC #Order filling type
        }
        result = mt5.order_send(request)
        print('result: ',result)
        
        if result != None:
            if result.retcode == mt5.TRADE_RETCODE_DONE:      
                SuccessfulTrade = 'Buy order Placed successfully.☺\n\n'
                VoiceAnnounce(f'{SuccessfulTrade[:-5]} with volume {Volume}.')
                SuccessfulTrade += "\n".join("{!r}: {!r}".format(k, v) for k, v in request.items())
                print(SuccessfulTrade)
                showmessage('Successful position',SuccessfulTrade,TIMEOUT=8000,TYP='info')
            
            else:
                TardeError=f"Sorry! Failed to place buy order!: {result.retcode}☹ \n{result.comment}\n\n{TradeErrorDict[result.retcode]}"
                print(TardeError)
                VoiceAnnounce('Sorry! Failed to place buy order!')
                showmessage('Trade Error',TardeError,TIMEOUT=5000,TYP='Erro')
        else:
            showmessage('Unknown Error','Unknown Error.\n mt5.order_send returned no results!',TIMEOUT=5000,TYP='Erro')            
    else:
        showmessage('Connection Error', 'Could not connect to broker\nCheck internet connection!',TIMEOUT=3000,TYP='Erro')

def Instant_Sell():
    global Volume
    global LoginFlag1
    global Symbol
    if not LoginFlag1:
        Login_To_MT5()    
    if LoginFlag1:
        symbol_info = mt5.symbol_info(Symbol)
        if symbol_info is None:
            showmessage('Symbol Error',f'Symbol Error!\n{Symbol} not recognized!',TIMEOUT=4000,TYP='Erro')
            return
        # Ensure the symbol is visible
        if not symbol_info.visible:
            if not mt5.symbol_select(Symbol, True):
                showmessage('Symbol Error', f'Failed to select symbol {Symbol}',TIMEOUT=4000,TYP='Erro')
                VoiceAnnounce('Failed to select symbol {Symbol}')
                return
        # Get the price and point for the symbol       
        point = symbol_info.point
        price = mt5.symbol_info_tick(Symbol).bid
        
        if price is None:
            showmessage('Price catch Error', f'Failed to catch the price to sell for the symbol {Symbol}',TIMEOUT=4000,TYP='Erro')
            return
        
#         result = mt5.order_send(symbol=Symbol, trade_operation=mt5.ORDER_TYPE_BUY, volume=Volume, slippage=3,deviation=20, type_filling=mt5.ORDER_FILLING_IOC)
        request = {
            "action": mt5.TRADE_ACTION_DEAL,
            "symbol": Symbol,
            "volume": float(Volume),
            "type": mt5.ORDER_TYPE_SELL,
            "price": price,
            "sl": price + 100 * point,
            "tp": price - 100 * point,
            "deviation": 20,
            "magic": 234000,
            "comment": __file__[:-3],#'Sell from python script',
            "type_time": mt5.ORDER_TIME_GTC,
            "type_filling": mt5.ORDER_FILLING_FOK,  #mt5.ORDER_FILLING_IOC #Order filling type
        }
        result = mt5.order_send(request)
        
        if result != None:
            if result.retcode == mt5.TRADE_RETCODE_DONE:
                SuccessfulTrade = 'Sell order Placed successfully.☺\n\n'
                SuccessfulTrade += "\n".join("{!r}: {!r}".format(k, v) for k, v in request.items())
                print(SuccessfulTrade)
                VoiceAnnounce(f'Sell order Placed successfully with volume {Volume}')
                showmessage('Successful position',SuccessfulTrade,TIMEOUT=6000,TYP='info')
            
            else:
                TardeError= f"Sorry! Failed to place Sell order!: {result.retcode}☹ \n{result.comment}\n{TradeErrorDict[result.retcode]}"
                print(TardeError)
                VoiceAnnounce('Sorry! Failed to place Sell order!')
                showmessage('Trade Error',TardeError,TIMEOUT=5000,TYP='Erro')
        else:
            showmessage('Unknown Error','Unknown Error.\n mt5.order_send returned no results!',TIMEOUT=3000,TYP='Erro')
            
    else:
        showmessage('Connection Error', 'Could not connect to broker\nCheck internet connection!',TIMEOUT=3000,TYP='Erro')
        VoiceAnnounce('Could not connect to broker. Check internet connection.')

def trade_option():
    print(TRADVAR.get())
    if TRADVAR.get() == 'Instant Buy':
        Instant_Buy()
    elif TRADVAR.get() == 'Instant Sell':
        Instant_Sell()
    else:
        showmessage('Trade Error',TardeError,TIMEOUT=3000,TYP='Erro')
        
def OpenIndicMenu(event):
    indicmenu = Indicopt['menu']
    x = Indicopt.winfo_rootx()
    y = Indicopt.winfo_rooty() + Indicopt.winfo_height()
    indicmenu.post(x, y)

def All_Bootstrap_indics():
    ADD_SMAs(6)
    ADD_EMAs(7)
    Add_RSI(0)
    Add_MACD()
    Add_bollingerBands(0)
    Add_ATR(0)
    Add_STOCH(0)
    Add_CMF(0)
    Add_williams_r(0)
    Add_CCI(0)
    Add_SAR()
    Add_ichimoku()
    Add_ADX(0) 

def All_indics():
    ADD_SMAs(0)
    ADD_EMAs(0)
    Add_RSI(0)
    Add_MACD()
    Add_bollingerBands(0)
    Add_ATR(0)
    Add_STOCH(0)
    Add_CMF(0)
    Add_williams_r(0)
    Add_CCI(0)
    Add_SAR()
    Add_ichimoku()
    Add_ADX(0)     
    
def Indic_option(M=True):
    global FEAT
    #print('this is M:',M)
    if M:
        mode = INDICVAR.get()
        print('mode:',mode)
        if mode == 'ADD SMA indic.':
            ADD_SMAs(6)
        elif mode == 'ADD EMA indic.':
            ADD_EMAs(7)
        elif mode == 'ADD RSI Osc.':
            Add_RSI(0)
        elif mode == 'ADD MACD indic.':
            Add_MACD()
        elif mode == 'ADD Bol Band indic.':
            Add_bollingerBands(0)
        elif mode == 'ADD ATR indic.':
            Add_ATR(0)
        elif mode == 'ADD Stochastic Osc.':
            Add_STOCH(0)
        elif mode == 'ADD CMF indic.':
            Add_CMF(0)
        elif mode == 'ADD williams_r indic.':
            Add_williams_r(0)
        elif mode == 'ADD CCI indic.':
            Add_CCI(0)
        elif mode == 'Add SAR indic.':
            Add_SAR()
        elif mode == 'Add ichimoku':
            Add_ichimoku()
        elif mode == 'Add ADX':
            Add_ADX(0)
        else:
            showmessage('Indicator selection Error','Indicator selection Error!',TIMEOUT=3000,TYP='Erro')

    else:
        mode = FEAT
        print('mode:',mode)
        if mode in 'ADD SMA indic.':
            ADD_SMAs(0)
        elif mode in 'ADD EMA indic.':
            ADD_EMAs(0)
        elif mode in 'ADD RSI Osc.':
            Add_RSI(0)
        elif mode in 'ADD MACD indic.':
            Add_MACD()
        elif mode in 'ADD Bol Band indic.':
            Add_bollingerBands(0)
        elif mode in 'ADD ATR indic.':
            Add_ATR(0)
        elif mode in 'ADD Stochastic Osc.':
            Add_STOCH(0)
        elif mode in 'ADD CMF indic.':
            Add_CMF(0)
        elif mode in 'ADD williams_r indic.':
            Add_williams_r(0)
        elif mode in 'ADD CCI indic.':
            Add_CCI(0)
        elif mode in 'Add SAR indic.':
            Add_SAR()
        elif mode in 'Add ichimoku':
            Add_ichimoku()
        elif mode in 'Add ADX':
            Add_ADX(0)
        else:
            showmessage('Indicator selection Error','Indicator selection Error!',TIMEOUT=3000,TYP='Erro')

def model_option():
    MODLBTN.config(text = f"Run {ModelVAR.get()} Analysis",fg='yellow',font='normal',bg='mediumvioletred')
    RUNBTN.configure(text=f"RUN {ModelVAR.get()}\n with test size:{TestSize} and train size:{Trainsize}\nTimeFrame:{TimeFrame}", fg='chartreuse',font=('Times', FONTSIZE, 'normal'))  

def Lagwindow_option():
    WInBTN.config(text = f"Run {CompareWINVAR.get()} lag comparison",fg='yellow',font='normal',bg='mediumvioletred')
    RUNBTN.configure(text=f"RUN {ModelVAR.get()}\nWith test size:{TestSize}/train size:{Trainsize}\nTF:{TimeFrame} {CompareWINVAR.get()}", fg='chartreuse',font=('Times', FONTSIZE, 'normal'))  

def FeaTuRes_option():
    FEAtBTN.config(text = f'Compare models for {CompareFEAtVAR.get()}',fg='yellow',font='normal',bg='mediumvioletred')
    RUNBTN.configure(text=f"RUN Models comparison\nfor {CompareFEAtVAR.get()}\nWith test size:{TestSize}/train size:{Trainsize}\nTimeFrame:{TimeFrame}", fg='chartreuse',font=('Times', FONTSIZE, 'normal'))  

def Modelsoption():
    CmpFeatBTN.config(text = f"Run {''.join(x[0] for x in CompareVAR.get().split())} Comparison",fg='yellow',font='normal',bg='mediumvioletred')
    RUNBTN.configure(text=f"RUN {''.join(x[0] for x in CompareVAR.get().split())} Features comparison\nWith test size:{TestSize}/train size:{Trainsize}TimeFrame:{TimeFrame}", fg='chartreuse',font=('Times', FONTSIZE, 'normal'))  

def ADD_SMAs(PRESETVALUE):
    global df
    global WIN13
    global DynamicColumns
    global Fcateg
    n = PRESETVALUE
    try:
        df
    except NameError as ERR:
        showmessage('Dataframe Error',f'Dataframe is not available\n{ERR}',TIMEOUT=3000,TYP='Erro')
        LoadDataBase('')
    
    def exitsma(_):
        global WIN13
        nonlocal n
        nonlocal SMASIZE
        n = int(SMASIZE.get())
        SMAWINTXT = f'SMA window size selected equal to {n}'
        Response = messagebox.askokcancel("Set window?",f"{SMAWINTXT}\nContinue?",default=messagebox.OK)
        if Response:      
            WIN13.destroy()
            WIN13.quit()
            print(SMAWINTXT)
    if not PRESETVALUE:       
        WIN13 = tk.Toplevel(root)
        WIN13.bind('<Escape>',EXIT_One)
        WIN13.config(bg='#474747')
        WIN13.attributes("-topmost", True)
        SMALABEL = tk.Label(WIN13, text="↓Specify size of SMA↓",font=('Times',FONTSIZE2),bd=2,bg='#25251E',fg='white')
        SMALABEL.pack()#.place(x=w//2+40,y=40)
        smasize = tk.StringVar(WIN13)
        SMASIZE = tk.Spinbox(WIN13, values=[v for v in range(2,41)],font=('Times', FONTSIZE2, 'bold'),justify=tk.CENTER,bd=2,activebackground='white',bg='#05051E' ,fg='orange',textvariable = smasize)#, command = MAKESMA)
        SMASIZE.pack()#.place(x=w//2+40,y=80)
        SMASIZE.delete(0,"end")
        SMASIZE.insert(0,5)
        btSMA = tk.Button(WIN13,text="Set the averaging window value!",bg='#474747',fg='yellow',font=('times',FONTSIZE,'bold'),relief=tk.RAISED,bd=5,command=lambda:exitsma(''))
        WIN13.bind('<Return>',exitsma) 
        WIN13.focus_set()
        WIN13.lift()
        WIN13.focus_force()
        SMASIZE.focus()
        btSMA.pack()
        WIN13.mainloop()
    try:
        df['SMA'+str(n)]
    except KeyError:
        showmessage(f'Adding SMA{n} to df',f'Simple Moving average{n} is added to dataframe',TIMEOUT=2500,TYP='info')
        df['SMA'+str(n)]=df['CLOSE'].rolling(window=n).mean()
        DynamicColumns += ['SMA'+str(n)]
        Update_features_List('SMA'+str(n))
    else:
        Response = messagebox.askokcancel("Overwrite?",f"Overwrite the existing {'SMA'+str(n)} indicator",default=messagebox.CANCEL)
        if Response:
            df['SMA'+str(n)]=df['CLOSE'].rolling(window=n).mean()
            DynamicColumns += ['SMA'+str(n)]
            Update_features_List('SMA'+str(n))
            showmessage('Overwrite',f"Indicator {'SMA'+str(n)} overwrited!",TIMEOUT=2500,TYP='warn')

def ADD_EMAs(PRESETVALUE):
    global df
    global WIN13
    global DynamicColumns
    global Fcateg
    n = PRESETVALUE
    try:
        df
    except NameError as ERR:
        showmessage('Dataframe Error',f'Dataframe is not available\n{ERR}',TIMEOUT=3000,TYP='Erro')
        LoadDataBase('')

    def exitema(_):
        global WIN13
        nonlocal n
        nonlocal EMASIZE
        n = int(EMASIZE.get())
        EMAWINTXT = f'EMA window size selected equal to {n}'
        Response = messagebox.askokcancel("Set window?",f"{EMAWINTXT}\nContinue?",default=messagebox.OK)
        if Response:        
            WIN13.destroy()
            WIN13.quit()
            print(EMAWINTXT)
    if not PRESETVALUE:        
        WIN13 = tk.Toplevel(root)
        WIN13.config(bg='#474747')
        WIN13.bind('<Escape>',EXIT_One)
        WIN13.attributes("-topmost", True)
        EMALABEL = tk.Label(WIN13, text="↓Specify size of EMA↓",font=('Times',FONTSIZE2),bd=2,bg='#25251E',fg='white')
        EMALABEL.pack()#.place(x=w//2+40,y=40)
        emasize = tk.StringVar(WIN13)
        EMASIZE = tk.Spinbox(WIN13, values=[v for v in range(2,20)],font=('Times', FONTSIZE2, 'bold'),justify=tk.CENTER,bd=2,activebackground='orange',bg='#05051E' ,fg='orange',textvariable = emasize)
        EMASIZE.pack()#.place(x=w//2+40,y=80)
        EMASIZE.delete(0,"end")
        EMASIZE.insert(0,5)
        btEMA = tk.Button(WIN13,text="Set the averaging window value!",bg='#474747',fg='yellow',font=('times',FONTSIZE,'bold'),relief=tk.RAISED,bd=5,command=lambda:exitema(''))
        WIN13.bind('<Return>',exitema)
        #WIN13.focus_set()
        WIN13.lift()
        WIN13.focus_force()
        btEMA.pack()
        WIN13.mainloop()
    try:
        df['EMA'+str(n)]
    except KeyError:
        showmessage(f'Adding EMA{n} to df',f'Exponential Moving average{n} is added to dataframe',TIMEOUT=2500,TYP='info')
        df['EMA'+str(n)]=df['CLOSE'].ewm(span=n, adjust=False).mean()
        DynamicColumns += ['EMA'+str(n)]
        Update_features_List('EMA'+str(n))
    else:
        Response = messagebox.askokcancel("Overwrite?",f"Overwrite the existing {'EMA'+str(n)} indicator",default=messagebox.CANCEL)
        if Response:
            df['EMA'+str(n)]=df['CLOSE'].ewm(span=n, adjust=False).mean()
            DynamicColumns += ['EMA'+str(n)]
            Update_features_List('EMA'+str(n))
            showmessage('Overwrite',f"Indicator {'EMA'+str(n)} overwrited!",TIMEOUT=2500,TYP='warn')
            DynamicColumns += ['EMA'+str(n)]

def Add_RSI(PRESETVALUE):
    global df
    global WIN13
    n = PRESETVALUE
    preferred = 14
    global DynamicColumns
    try:
        df
    except NameError as ERR:
        showmessage('Dataframe Error',f'Dataframe is not available\n{ERR}',TIMEOUT=3000,TYP='Erro')
        LoadDataBase('')
        
    def exitRSI(_):
        global WIN13
        nonlocal n
        nonlocal RSISIZE
        n = int(RSISIZE.get())
        RSIWINTXT = f'RSI window size selected equal to {n}'
        Response = messagebox.askokcancel("Set window?",f"{RSIWINTXT}\nContinue?",default=messagebox.OK)
        if Response:
            showmessage('Set window',RSIWINTXT,TIMEOUT=2500,TYP='info')        
            WIN13.destroy()
            WIN13.quit()
            print(RSIWINTXT)
    if not PRESETVALUE:        
        WIN13 = tk.Toplevel(root)
        WIN13.config(bg='#474747')
        WIN13.bind('<Escape>',EXIT_One)
        WIN13.attributes("-topmost", True)
        RSILABEL = tk.Label(WIN13, text="↓Specify size of RSI window↓",font=('Times',FONTSIZE2),bd=2,bg='#25251E',fg='white')
        RSILABEL.pack()#.place(x=w//2+40,y=40)
        RSIsize = tk.StringVar(WIN13)
        RSISIZE = tk.Spinbox(WIN13, values=[v for v in range(2,20)],font=('Times', FONTSIZE2, 'bold'),justify=tk.CENTER,bd=2,activebackground='orange',bg='#05051E' ,fg='orange',textvariable = RSIsize)
        RSISIZE.pack()#.place(x=w//2+40,y=80)
        RSISIZE.delete(0,"end")
        RSISIZE.insert(0,14)
        btRSI = tk.Button(WIN13,text="Set the averaging window value!",bg='#474747',fg='yellow',font=('times',FONTSIZE,'bold'),relief=tk.RAISED,bd=5,command=lambda:exitRSI(''))
        btRSI.pack()
        WIN13.attributes("-topmost", True)
        WIN13.bind('<Return>',exitRSI)
        WIN13.focus_set()
        WIN13.lift()
        WIN13.focus_force()
        RSISIZE.focus_set()
        WIN13.mainloop()
        
    try:
        df['RSI'+str(n)]
    except KeyError:
        showmessage(f'Adding RSI_{n} to df',f'RSI_{n} is added to dataframe',TIMEOUT=2500,TYP='info')
        delta = df['CLOSE'].diff()
        gain = (delta.where(delta > 0, 0)).rolling(window=n).mean()
        loss = (-delta.where(delta < 0, 0)).rolling(window=n).mean()
        rs = gain / loss
        df['RSI'+str(n)] = 100 - (100 / (1 + rs))
        DynamicColumns += ['RSI'+str(n)]
        Update_features_List('RSI'+str(n))
    else:
        Response = messagebox.askokcancel("Overwrite?",f"Overwrite the existing {'SMA'+str(n)} indicator",default=messagebox.CANCEL)
        if Response:
            delta = df['CLOSE'].diff()
            gain = (delta.where(delta > 0, 0)).rolling(window=n).mean()
            loss = (-delta.where(delta < 0, 0)).rolling(window=n).mean()
            rs = gain / loss
            df['RSI'+str(n)] = 100 - (100 / (1 + rs))
            DynamicColumns += ['RSI'+str(n)]
            Update_features_List('RSI'+str(n))
            showmessage('Overwrite',f"Indicator {'RSI'+str(n)} overwrited!",TIMEOUT=2500,TYP='warn')
        
def Add_MACD():
    global df
    short_window = 12
    long_window = 26
    signal_window = 9
    global DynamicColumns
    try:
        df['MACD']
    except NameError as ERR:
        showmessage('Dataframe Error',f'Dataframe is not available\n{ERR}',TIMEOUT=3000,TYP='Erro')
        LoadDataBase('') 
    except KeyError:
        showmessage(f'Adding MACD to df',f'MACD is added to dataframe',TIMEOUT=2500,TYP='info')
        df['EMA12'] = df['CLOSE'].ewm(span=short_window, adjust=False).mean()
        df['EMA26'] = df['CLOSE'].ewm(span=long_window, adjust=False).mean()
        df['MACD'] = df['EMA12'] - df['EMA26']
        df.drop(['EMA12','EMA26'],axis=1,inplace=True)
        df['Signal Line'] = df['MACD'].ewm(span=signal_window, adjust=False).mean()
        DynamicColumns += ['MACD']
        Update_features_List('MACD')
        DynamicColumns += ['Signal Line']
        Update_features_List('Signal Line')
    else:
        Response = messagebox.askokcancel("Overwrite?","Overwrite the existing MACD indicator",default=messagebox.CANCEL)
        if Response:
            df['EMA12'] = df['CLOSE'].ewm(span=short_window, adjust=False).mean()
            df['EMA26'] = df['CLOSE'].ewm(span=long_window, adjust=False).mean()
            df['MACD'] = df['EMA12'] - df['EMA26']
            df.drop(['EMA12','EMA26'],axis=1,inplace=True)
            df['Signal Line'] = df['MACD'].ewm(span=signal_window, adjust=False).mean()
            DynamicColumns += ['MACD']
            Update_features_List('MACD')
            DynamicColumns += ['Signal Line']
            Update_features_List('Signal Line')
            showmessage('Overwrite',"MACD Indicator overwrited!",TIMEOUT=2500,TYP='warn')
            
def Add_bollingerBands(PRESETVALUE):
    global df
    global WIN13
    n = PRESETVALUE
    global DynamicColumns
    global Fcateg
    try:
        df
    except NameError as ERR:
        showmessage('Dataframe Error',f'Dataframe is not available\n{ERR}',TIMEOUT=3000,TYP='Erro')
        LoadDataBase('')
    def exitBolBan(_):
        global WIN13
        nonlocal n
        nonlocal BolBanSIZE
        n = int(BolBanSIZE.get())
        BolBanWINTXT = f'Bollinger Band window size selected equal to {n}'
        Response = messagebox.askokcancel("Set window?",f"{BolBanWINTXT}\nContinue?",default=messagebox.OK)
        if Response:
            showmessage('Set window',BolBanWINTXT,TIMEOUT=2500,TYP='info')        
            WIN13.destroy()
            WIN13.quit()
            print(BolBanWINTXT)
    if not PRESETVALUE:        
        WIN13 = tk.Toplevel(root)
        WIN13.config(bg='#474747')
        WIN13.bind('<Escape>',EXIT_One)
        WIN13.attributes("-topmost", True)
        BolBanLABEL = tk.Label(WIN13, text="↓Specify size of Bollinger Band window↓",font=('Times',FONTSIZE2),bd=2,bg='#474747',fg='white')
        BolBanLABEL.pack()#.place(x=w//2+40,y=40)
        BolBansize = tk.StringVar(WIN13)
        BolBanSIZE = tk.Spinbox(WIN13, values=[v for v in range(1,60)],font=('Times', FONTSIZE2, 'bold'),justify=tk.CENTER,bd=2,activebackground='orange',bg='#05051E' ,fg='orange',textvariable = BolBansize)
        BolBanSIZE.pack()#.place(x=w//2+40,y=80)
        BolBanSIZE.delete(0,"end")
        BolBanSIZE.insert(0,20)
        btBolBan = tk.Button(WIN13,text="Set the averaging window value!",bg='#474747',fg='yellow',font=('times',FONTSIZE2,'bold'),relief=tk.RAISED,bd=5,command=lambda:exitBolBan(''))
        btBolBan.pack()
        WIN13.bind('<Return>',exitBolBan)
        WIN13.focus_set()
        WIN13.lift()
        WIN13.focus_force()
        BolBanSIZE.focus_set()
        WIN13.mainloop()  
    no_of_std = 0.1*n  
    
    df['SMA'] = df['CLOSE'].rolling(window=PRESETVALUE).mean()
    df['STD'] = df['CLOSE'].rolling(window=PRESETVALUE).std()
    df['Upper_Band'] = df['SMA'] + (df['STD'] * no_of_std)
    df['Lower_Band'] = df['SMA'] - (df['STD'] * no_of_std)
    df.drop(['SMA','STD'],axis=1,inplace=True)
    DynamicColumns += ['Upper_Band']
    Update_features_List('Upper_Band')
    DynamicColumns += ['Lower_Band']
    Update_features_List('Lower_Band')

def Add_ATR(PRESETVALUE):
    global df
    global WIN13
    n = PRESETVALUE
    global DynamicColumns
    global Fcateg
    try:
        df
    except NameError as ERR:
        showmessage('Dataframe Error',f'Dataframe is not available\n{ERR}',TIMEOUT=3000,TYP='Erro')
        LoadDataBase('')
    def exitATR(_):
        global WIN13
        nonlocal n
        nonlocal ATRSIZE
        n = int(ATRSIZE.get())
        ATRWINTXT = f'ATR window size selected equal to {n}'
        Response = messagebox.askokcancel("Set window?",f"{ATRWINTXT}\nContinue?",default=messagebox.OK)
        if Response:
            showmessage('Set window',ATRWINTXT,TIMEOUT=2500,TYP='info')        
            WIN13.destroy()
            WIN13.quit()
            print(ATRWINTXT)    
    try:
        df['ATR']
    except KeyError:
        showmessage(f'Adding ATR to df',f'ATR is supposed to add to dataframe',TIMEOUT=2500,TYP='info')
    else:
        Response = messagebox.askokcancel("Overwrite?","Overwrite the existing ATR indicator",default=messagebox.CANCEL)
        if not Response:
            return
        else:
            showmessage('Overwrite',"ATR Indicator will be overwrited!",TIMEOUT=2500,TYP='warn')
    if not PRESETVALUE:        
        WIN13 = tk.Toplevel(root)
        WIN13.config(bg='#474747')
        WIN13.bind('<Escape>',EXIT_One)
        WIN13.attributes("-topmost", True)
        ATRLABEL = tk.Label(WIN13, text="↓Specify size of ATR window↓",font=('Times',FONTSIZE2),bd=2,bg='#474747',fg='white')
        ATRLABEL.pack()#.place(x=w//2+40,y=40)
        ATRsize = tk.StringVar(WIN13)
        ATRSIZE = tk.Spinbox(WIN13, values=[v for v in range(1,60)],font=('Times', FONTSIZE2, 'bold'),justify=tk.CENTER,bd=2,activebackground='orange',bg='#05051E' ,fg='orange',textvariable = ATRsize)
        ATRSIZE.pack()#.place(x=w//2+40,y=80)
        ATRSIZE.delete(0,"end")
        ATRSIZE.insert(0,20)
        btATR = tk.Button(WIN13,text="Set the averaging window value!",bg='#474747',fg='yellow',font=('times',FONTSIZE2,'bold'),relief=tk.RAISED,bd=5,command=lambda:exitATR(''))
        btATR.pack()
        WIN13.bind('<Return>',exitATR)
        WIN13.focus_set()
        WIN13.lift()
        WIN13.focus_force()
        ATRSIZE.focus_set()
        WIN13.mainloop()  
    df['High-Low'] = df['HIGH'] - df['LOW']
    df['High-PrevClose'] = np.abs(df['HIGH'] - df['CLOSE'].shift(1))
    df['Low-PrevClose'] = np.abs(df['LOW'] - df['CLOSE'].shift(1))
    df['TrueRange'] = df[['High-Low', 'High-PrevClose', 'Low-PrevClose']].max(axis=1)
    df['ATR'] = df['TrueRange'].rolling(window = n).mean()
    df.drop(['High-Low','High-PrevClose','Low-PrevClose','TrueRange'],axis=1,inplace=True)
    DynamicColumns += ['ATR']
    Update_features_List('ATR')
    showmessage(f'Adding ATR to df',f'ATR is added to dataframe',TIMEOUT=2500,TYP='info')

def Add_STOCH(PRESETVALUE):
    global df
    global WIN13
    n = PRESETVALUE
    global DynamicColumns
    global Fcateg
    
    try:
        df
    except NameError as ERR:
        showmessage('Dataframe Error',f'Dataframe is not available\n{ERR}',TIMEOUT=3000,TYP='Erro')
        LoadDataBase('')
    def exitStochastic(_):
        global WIN13
        nonlocal n
        nonlocal StochasticSIZE
        n = int(StochasticSIZE.get())
        StochasticWINTXT = f'Stochastic window size selected equal to {n}'
        Response = messagebox.askokcancel("Set window?",f"{StochasticWINTXT}\nContinue?",default=messagebox.OK)
        if Response:
            showmessage('Set window',StochasticWINTXT,TIMEOUT=2500,TYP='info')        
            WIN13.destroy()
            WIN13.quit()
            print(StochasticWINTXT)    
    try:
        df['Stoch']
    except KeyError:
        showmessage(f'Adding Stochastic to df',f'Stochastic is supposed to add to dataframe',TIMEOUT=2500,TYP='info')
    else:
        Response = messagebox.askokcancel("Overwrite?","Overwrite the existing Stochastic Oscillator",default=messagebox.CANCEL)
        if not Response:
            return
        else:
            showmessage('Overwrite',"Stochastic Oscillator will be overwrited!",TIMEOUT=2500,TYP='warn')
    if not PRESETVALUE:        
        WIN13 = tk.Toplevel(root)
        WIN13.config(bg='#474747')
        WIN13.bind('<Escape>',EXIT_One)
        WIN13.attributes("-topmost", True)
        StochasticLABEL = tk.Label(WIN13, text="↓Specify size of Stochastic window↓",font=('Times',FONTSIZE2),bd=2,bg='#474747',fg='white')
        StochasticLABEL.pack()#.place(x=w//2+40,y=40)
        Stochasticsize = tk.StringVar(WIN13)
        StochasticSIZE = tk.Spinbox(WIN13, values=[v for v in range(1,60)],font=('Times', FONTSIZE2, 'bold'),justify=tk.CENTER,bd=2,activebackground='orange',bg='#05051E' ,fg='orange',textvariable = Stochasticsize)
        StochasticSIZE.pack()#.place(x=w//2+40,y=80)
        StochasticSIZE.delete(0,"end")
        StochasticSIZE.insert(0,20)
        btStochastic = tk.Button(WIN13,text="Set the averaging window value!",bg='#474747',fg='gold',font=('times',FONTSIZE2,'bold'),relief=tk.RAISED,bd=5,command=lambda:exitStochastic(''))
        btStochastic.pack()
        WIN13.bind('<Return>',exitStochastic)
        WIN13.focus_set()
        WIN13.lift()
        WIN13.focus_force()
        StochasticSIZE.focus_set()
        WIN13.mainloop()
        
    df['Low-Min'] = df['LOW'].rolling(window = n).min()
    df['High-Max'] = df['HIGH'].rolling(window = n).max()
    df['Stoch'] = 100 * ((df['CLOSE'] - df['Low-Min']) / (df['High-Max'] - df['Low-Min']))
    df.drop(['Low-Min','High-Max'],axis=1,inplace=True)
    DynamicColumns += ['Stoch']
    Update_features_List('Stoch')
    showmessage(f'Adding Stochastic to df',f'Stochastic Oscillator is added to dataframe',TIMEOUT=2500,TYP='info')

def Add_CMF(PRESETVALUE):
    global df
    global WIN13
    n = PRESETVALUE
    global DynamicColumns
    global Fcateg
    try:
        df
    except NameError as ERR:
        showmessage('Dataframe Error',f'Dataframe is not available\n{ERR}',TIMEOUT=3000,TYP='Erro')
        LoadDataBase('')
    def exitCMF(_):
        global WIN13
        nonlocal n
        nonlocal CMFSIZE
        n = int(CMFSIZE.get())
        CMFWINTXT = f'CMF window size selected equal to {n}'
        Response = messagebox.askokcancel("Set window?",f"{CMFWINTXT}\nContinue?",default=messagebox.OK)
        if Response:
            showmessage('Set window',CMFWINTXT,TIMEOUT=2500,TYP='info')        
            WIN13.destroy()
            WIN13.quit()
            print(CMFWINTXT)    
    try:
        df['CMF']
    except KeyError:
        showmessage(f'Adding CMF to df',f'Chaikin Money Flow is supposed to add to dataframe',TIMEOUT=2500,TYP='info')
    else:
        Response = messagebox.askokcancel("Overwrite?","Overwrite the existing Chaikin Money Flow indicator",default=messagebox.CANCEL)
        if not Response:
            return
        else:
            showmessage('Overwrite',"Chaikin Money Flow Indicator will be overwrited!",TIMEOUT=2500,TYP='warn')
    if not PRESETVALUE:        
        WIN13 = tk.Toplevel(root)
        WIN13.config(bg='#474747')
        WIN13.bind('<Escape>',EXIT_One)
        WIN13.attributes("-topmost", True)
        CMFLABEL = tk.Label(WIN13, text="↓Specify size of CMF window↓",font=('Times',FONTSIZE2),bd=2,bg='#474747',fg='white')
        CMFLABEL.pack()#.place(x=w//2+40,y=40)
        CMFsize = tk.StringVar(WIN13)
        CMFSIZE = tk.Spinbox(WIN13, values=[v for v in range(1,60)],font=('Times', FONTSIZE2, 'bold'),justify=tk.CENTER,bd=2,activebackground='orange',bg='#05051E' ,fg='orange',textvariable = CMFsize)
        CMFSIZE.pack()#.place(x=w//2+40,y=80)
        CMFSIZE.delete(0,"end")
        CMFSIZE.insert(0,20)
        btCMF = tk.Button(WIN13,text="Set the averaging window value!",bg='#474747',fg='yellow',font=('times',FONTSIZE2,'bold'),relief=tk.RAISED,bd=5,command=lambda:exitCMF(''))
        btCMF.pack()
        WIN13.bind('<Return>',exitCMF)
        WIN13.focus_set()
        WIN13.lift()
        WIN13.focus_force()
        CMFSIZE.focus_set()
        WIN13.mainloop()
    mfv = ((df['CLOSE'] - df['LOW']) - (df['HIGH'] - df['CLOSE'])) / (df['HIGH'] - df['LOW']) * df['tick_volume']
    mfv = mfv.replace([np.inf, -np.inf], 0).fillna(0)
    df['CMF'] = mfv.rolling(window = n).sum() / df['tick_volume'].rolling(window = n).sum()
    DynamicColumns += ['CMF']
    Update_features_List('CMF')
    showmessage(f'Adding CMF to df',f'Chaikin Money Flow is added to dataframe',TIMEOUT=2500,TYP='info')    

def Add_williams_r(PRESETVALUE):
    global df
    global WIN13
    n = PRESETVALUE
    global DynamicColumns
    global Fcateg
    try:
        df
    except NameError as ERR:
        showmessage('Dataframe Error',f'Dataframe is not available\n{ERR}',TIMEOUT=3000,TYP='Erro')
        LoadDataBase('')
    def exitwilliams_r(_):
        global WIN13
        nonlocal n
        nonlocal williams_rSIZE
        n = int(williams_rSIZE.get())
        williams_rWINTXT = f'williams_r window size selected equal to {n}'
        Response = messagebox.askokcancel("Set window?",f"{williams_rWINTXT}\nContinue?",default=messagebox.OK)
        if Response:
            showmessage('Set window',williams_rWINTXT,TIMEOUT=2500,TYP='info')        
            WIN13.destroy()
            WIN13.quit()
            print(williams_rWINTXT)    
    try:
        df['williams_r']
    except KeyError:
        showmessage(f'Adding williams_r to df',f'williams_r is supposed to add to dataframe',TIMEOUT=2500,TYP='info')
    else:
        Response = messagebox.askokcancel("Overwrite?","Overwrite the existing williams_r indicator",default=messagebox.CANCEL)
        if not Response:
            return
        else:
            showmessage('Overwrite',"williams_r Indicator will be overwrited!",TIMEOUT=2500,TYP='warn')
    if not PRESETVALUE:        
        WIN13 = tk.Toplevel(root)
        WIN13.config(bg='#474747')
        WIN13.bind('<Escape>',EXIT_One)
        WIN13.attributes("-topmost", True)
        williams_rLABEL = tk.Label(WIN13, text="↓Specify size of williams_r window↓",font=('Times',FONTSIZE2),bd=2,bg='#474747',fg='white')
        williams_rLABEL.pack()#.place(x=w//2+40,y=40)
        williams_rsize = tk.StringVar(WIN13)
        williams_rSIZE = tk.Spinbox(WIN13, values=[v for v in range(1,60)],font=('Times', FONTSIZE2, 'bold'),justify=tk.CENTER,bd=2,activebackground='orange',bg='#05051E' ,fg='orange',textvariable = williams_rsize)
        williams_rSIZE.pack()#.place(x=w//2+40,y=80)
        williams_rSIZE.delete(0,"end")
        williams_rSIZE.insert(0,14)
        btwilliams_r = tk.Button(WIN13,text="Set the averaging window value!",bg='#474747',fg='yellow',font=('times',FONTSIZE2,'bold'),relief=tk.RAISED,bd=5,command=lambda:exitwilliams_r(''))
        btwilliams_r.pack()
        WIN13.bind('<Return>',exitwilliams_r)
        WIN13.focus_set()
        WIN13.lift()
        WIN13.focus_force()
        williams_rSIZE.focus_set() 
        WIN13.mainloop()   
    highest_high = df['HIGH'].rolling(window = n).max()
    lowest_low = df['LOW'].rolling(window = n).min()
    df['Williams_%R'] = -100 * ((highest_high - df['CLOSE']) / (highest_high - lowest_low + 1e-9))
    DynamicColumns += ['Williams_%R']
    Update_features_List('Williams_%R')
    showmessage(f'Adding williams_r to df',f'williams_r is added to dataframe',TIMEOUT=2500,TYP='info')

def Add_CCI(PRESETVALUE):
    global df
    global WIN13
    n = PRESETVALUE
    global DynamicColumns
    global Fcateg
    
    try:
        df
    except NameError as ERR:
        showmessage('Dataframe Error',f'Dataframe is not available\n{ERR}',TIMEOUT=3000,TYP='Erro')
        LoadDataBase('')
    def exitCCI(_):
        global WIN13
        nonlocal n
        nonlocal CCISIZE
        n = int(CCISIZE.get())
        CCIWINTXT = f'CCI window size selected equal to {n}'
        Response = messagebox.askokcancel("Set window?",f"{CCIWINTXT}\nContinue?",default=messagebox.OK)
        if Response:
            showmessage('Set window',CCIWINTXT,TIMEOUT=2500,TYP='info')        
            WIN13.destroy()
            WIN13.quit()
            print(CCIWINTXT)    
    try:
        df['CCI']
    except KeyError:
        showmessage(f'Adding CCI to df',f'CCI is supposed to add to dataframe',TIMEOUT=2500,TYP='info')
    else:
        Response = messagebox.askokcancel("Overwrite?","Overwrite the existing CCI indicator",default=messagebox.CANCEL)
        if not Response:
            return
        else:
            showmessage('Overwrite',"CCI Indicator will be overwrited!",TIMEOUT=2500,TYP='warn')
    if not PRESETVALUE:        
        WIN13 = tk.Toplevel(root)
        WIN13.config(bg='#474747')
        WIN13.bind('<Escape>',EXIT_One)
        WIN13.attributes("-topmost", True)
        CCILABEL = tk.Label(WIN13, text="↓Specify size of CCI window↓",font=('Times',FONTSIZE2),bd=2,bg='#474747',fg='white')
        CCILABEL.pack()#.place(x=w//2+40,y=40)
        CCIsize = tk.StringVar(WIN13)
        CCISIZE = tk.Spinbox(WIN13, values=[v for v in range(1,60)],font=('Times', FONTSIZE2, 'bold'),justify=tk.CENTER,bd=2,activebackground='orange',bg='#05051E' ,fg='orange',textvariable = CCIsize)
        CCISIZE.pack()#.place(x=w//2+40,y=80)
        CCISIZE.delete(0,"end")
        CCISIZE.insert(0,20)
        btCCI = tk.Button(WIN13,text="Set the averaging window value!",bg='#474747',fg='yellow',font=('times',FONTSIZE2,'bold'),relief=tk.RAISED,bd=5,command=lambda:exitCCI(''))
        btCCI.pack()
        WIN13.bind('<Return>',exitCCI)
        WIN13.focus_set()
        WIN13.lift()
        WIN13.focus_force()
        WIN13.mainloop()
        
    tp = (df['HIGH'] + df['LOW'] + df['CLOSE']) / 3
    sma = tp.rolling(window = n).mean()
    mad = tp.rolling(window = n).apply(lambda x: np.mean(np.abs(x - x.mean())))
    df['CCI'] = (tp - sma) / (0.015 * mad)
    DynamicColumns += ['CCI']
    Update_features_List('CCI')
    showmessage(f'Adding CCI to df',f'CCI is added to dataframe',TIMEOUT=2500,TYP='info')
    
def Add_SAR():
    global df
    global WIN13
    global DynamicColumns
    global Fcateg
    
    try:
        df
    except NameError as ERR:
        showmessage('Dataframe Error',f'Dataframe is not available\n{ERR}',TIMEOUT=3000,TYP='Erro')
        LoadDataBase('')  
    try:
        df['SAR']
    except KeyError:
        showmessage(f'Adding Parabolic SAR to df',f'Parabolic SAR is supposed to add to dataframe',TIMEOUT=2500,TYP='info')
    else:
        Response = messagebox.askokcancel("Overwrite?","Overwrite the existing Parabolic SAR indicator",default=messagebox.CANCEL)
        if not Response:
            return
        else:
            showmessage('Overwrite',"Parabolic SAR Indicator will be overwrited!",TIMEOUT=2500,TYP='warn')    
    
    df['SAR'] = df['CLOSE'] * 0  # Initialize SAR column

    af = 0.02  # Acceleration factor
    max_af = 0.2  # Maximum acceleration factor
    trend = 1  # 1 for uptrend, -1 for downtrend
    ep = df['LOW'].iloc[0]  # Extremum price
    sar = df['HIGH'].iloc[0]  # Starting SAR value

    for i in range(1, len(df)):
        prev_sar = sar
        sar = sar + af * (ep - sar)
        if trend == 1:
            if df['LOW'].iloc[i] < sar:
                trend = -1
                sar = ep
                ep = df['LOW'].iloc[i]
                af = 0.02
            else:
                if df['HIGH'].iloc[i] > ep:
                    ep = df['HIGH'].iloc[i]
                    af = min(af + 0.02, max_af)
        else:
            if df['HIGH'].iloc[i] > sar:
                trend = 1
                sar = ep
                ep = df['HIGH'].iloc[i]
                af = 0.02
            else:
                if df['LOW'].iloc[i] < ep:
                    ep = df['LOW'].iloc[i]
                    af = min(af + 0.02, max_af)
        df.loc[df.index[i],'SAR'] = sar
    DynamicColumns += ['SAR']
    Update_features_List('SAR')
    showmessage(f'Adding SAR to df',f'Parabolic SAR is added to dataframe',TIMEOUT=2500,TYP='info')

def Add_ichimoku():
    global df
    global DynamicColumns
    global Fcateg 
    try:
        df
    except NameError as ERR:
        showmessage('Dataframe Error',f'Dataframe is not available\n{ERR}',TIMEOUT=3000,TYP='Erro')
        LoadDataBase('') 
    try:
        df['Tenkan_sen']
    except KeyError:
        showmessage(f'Adding ichimoku to df',f'ichimoku is supposed to add to dataframe',TIMEOUT=2500,TYP='info')
    else:
        Response = messagebox.askokcancel("Overwrite?","Overwrite the existing ichimoku indicator",default=messagebox.CANCEL)
        if Response:
            showmessage('Overwrite',"MACD Indicator will be overwritten!",TIMEOUT=2500,TYP='warn')
        else:
            return

    # Conversion Line (Tenkan-sen)
    period9_high = df['HIGH'].rolling(window=9).max()
    period9_low = df['LOW'].rolling(window=9).min()
    df['Tenkan_sen'] = (period9_high + period9_low) / 2

    # Base Line (Kijun-sen)
    period26_high = df['HIGH'].rolling(window=26).max()
    period26_low = df['LOW'].rolling(window=26).min()
    df['Kijun_sen'] = (period26_high + period26_low) / 2

    # Leading Span A (Senkou Span A)
    df['Senkou_span_a'] = ((df['Tenkan_sen'] + df['Kijun_sen']) / 2).shift(26)

    # Leading Span B (Senkou Span B)
    period52_high = df['HIGH'].rolling(window=52).max()
    period52_low = df['LOW'].rolling(window=52).min()
    df['Senkou_span_b'] = ((period52_high + period52_low) / 2).shift(26)

    # Lagging Span (Chikou Span)
    df['Chikou_span'] = df['CLOSE'].shift(-26)
    DynamicColumns += ['Senkou_span_a']
    Update_features_List('Senkou_span_a')
    DynamicColumns += ['Senkou_span_b']
    Update_features_List('Senkou_span_b')
    DynamicColumns += ['Chikou_span']
    Update_features_List('Chikou_span')
    DynamicColumns += ['Tenkan_sen']
    Update_features_List('Tenkan_sen')
    DynamicColumns += ['Kijun_sen']
    Update_features_List('Kijun_sen')

def Add_ADX(PRESETVALUE):
    global df
    global WIN13
    n = PRESETVALUE
    global DynamicColumns
    global Fcateg
    try:
        df
    except NameError as ERR:
        showmessage('Dataframe Error',f'Dataframe is not available\n{ERR}',TIMEOUT=3000,TYP='Erro')
        LoadDataBase('')
    def exitADX(_):
        global WIN13
        nonlocal n
        nonlocal ADXSIZE
        n = int(ADXSIZE.get())
        ADXWINTXT = f'ADX window size selected equal to {n}'
        Response = messagebox.askokcancel("Set window?",f"{ADXWINTXT}\nContinue?",default=messagebox.OK)
        if Response:
            showmessage('Set window',ADXWINTXT,TIMEOUT=2500,TYP='info')        
            WIN13.destroy()
            WIN13.quit()
            print(ADXWINTXT)    
    try:
        df['ADX']
    except KeyError:
        showmessage(f'Adding ADX to df',f'ADX is supposed to add to dataframe',TIMEOUT=2500,TYP='info')
    else:
        Response = messagebox.askokcancel("Overwrite?","Overwrite the existing ADX indicator",default=messagebox.CANCEL)
        if not Response:
            return
        else:
            showmessage('Overwrite',"ADX Indicator will be overwrited!",TIMEOUT=2500,TYP='warn')
    if not PRESETVALUE:        
        WIN13 = tk.Toplevel(root)
        WIN13.config(bg='#474747')
        WIN13.bind('<Escape>',EXIT_One)
        WIN13.attributes("-topmost", True)
        ADXLABEL = tk.Label(WIN13, text="↓Specify size of ADX window↓",font=('Times',FONTSIZE2),bd=2,bg='#474747',fg='white')
        ADXLABEL.pack()#.place(x=w//2+40,y=40)
        ADXsize = tk.StringVar(WIN13)
        ADXSIZE = tk.Spinbox(WIN13, values=[v for v in range(1,60)],font=('Times', FONTSIZE2, 'bold'),justify=tk.CENTER,bd=2,activebackground='orange',bg='#05051E' ,fg='orange',textvariable = ADXsize)
        ADXSIZE.pack()#.place(x=w//2+40,y=80)
        ADXSIZE.delete(0,"end")
        ADXSIZE.insert(0,14)
        btADX = tk.Button(WIN13,text="Set the averaging window value!",bg='#474747',fg='yellow',font=('times',FONTSIZE2,'bold'),relief=tk.RAISED,bd=5,command=lambda:exitADX(''))
        btADX.pack()
        WIN13.bind('<Return>',exitADX)
        WIN13.focus_set()
        WIN13.lift()
        WIN13.focus_force()
        ADXSIZE.focus_set()
        WIN13.mainloop()

    df['TR'] = pd.concat([df['HIGH'] - df['LOW'], abs(df['HIGH'] - df['CLOSE'].shift(1)),abs(df['LOW'] - df['CLOSE'].shift(1))], axis=1).max(axis=1)
    
    df['+DM'] = np.where((df['HIGH'] - df['HIGH'].shift(1)) > (df['LOW'].shift(1) - df['LOW']),
                         df['HIGH'] - df['HIGH'].shift(1), 0)
    df['+DM'] = df['+DM'].clip(lower=0)
    
    df['-DM'] = np.where((df['LOW'].shift(1) - df['LOW']) > (df['HIGH'] - df['HIGH'].shift(1)),
                         df['LOW'].shift(1) - df['LOW'], 0)
    df['-DM'] = df['-DM'].clip(lower=0)
    
    # Calculate the smoothed averages
    tr_smooth = df['TR'].rolling(window = n).sum()
    plus_dm_smooth = df['+DM'].rolling(window = n).sum()
    minus_dm_smooth = df['-DM'].rolling(window = n).sum()
    
    # Calculate the Directional Indicators
    df['+DI'] = 100 * (plus_dm_smooth / tr_smooth)
    df['-DI'] = 100 * (minus_dm_smooth / tr_smooth)
    df['DX'] = 100 * (abs(df['+DI'] - df['-DI']) / (df['+DI'] + df['-DI'] + 1e-9))    
    df['ADX'] = df['DX'].rolling(window = n).mean() 
    df.drop(['TR', '+DM', '-DM', 'DX'], axis=1, inplace=True)
    DynamicColumns += ['-DI']
    Update_features_List('-DI')
    DynamicColumns += ['+DI']
    Update_features_List('+DI')
    DynamicColumns += ['ADX']
    Update_features_List('ADX')


def GSE_enable():
    global ModelStorm
    global Strategy
    ModelStorm = True
    Strategy = 'GSE Strategy'
    RUNBTN.configure(text=f"RUN {Strategy}\n with test size:{TestSize} and train size:{Trainsize}\nTimeFrame:{TimeFrame}", fg='chartreuse',font=('Times', FONTSIZE, 'normal'))  
    
def Homogeneous_enable():
    global HomogeneousEnsemble
    global Strategy
    HomogeneousEnsemble = True
    Strategy = 'HomogeneousEnsemble'

def Minimize(_):
    root.iconify()
    
def navigate_options(event):  
    # Get the current selection index  
    current_index = Modeloptions.index(ModelVAR.get())     
    if event.keysym == 'Down':  
        # Move down in the list  
        current_index = (current_index + 1) % len(Modeloptions)   
    elif event.keysym == 'Up':  
        # Move up in the list  
        current_index = (current_index - 1) % len(Modeloptions)      
    # Update the selected option  
    ModelVAR.set(Modeloptions[current_index])  # Update the selected variable
    MODLBTN.config(text = f'Run {Modeloptions[current_index]}',fg='yellow',font='normal',bg='mediumvioletred')
    RUNBTN.configure(text=f"RUN {Modeloptions[current_index]}\n with test size:{TestSize} and train size:{Trainsize}\nTimeFrame:{TimeFrame}", fg='chartreuse',font=('Times', FONTSIZE, 'normal'))

def modelFocus(_):
    global Modelopt
    Modelopt.focus_set()
#     Modelopt.focus_displayof()
#     Modelopt.focus_force()
    root.bind('<Down>', navigate_options)  
    root.bind('<Up>', navigate_options)
    root.bind('<Return>', RUNStrategy_)
    
def GETTESTSIZE():
    global TestSize
    TestSize = int(TESTSIZE.get())
    TESTLBL.configure(text=f"Test size is: {TestSize}", fg='Gold')
    GETTrainSIZE()
    if not TestSize%10:
        VoiceAnnounce('Test size is'+str(TestSize)+'candles')

def GETTrainSIZE():
    global Trainsize,TrainMUL
    try:
        df
    except NameError as ERR:
        showmessage('Dataframe Error',f'Load Dataframe first!\n{ERR}',TIMEOUT=3000,TYP='Erro')
        LoadDataBase('')
    TrainMul = TrainMUL.get()
    Trainsize = int(TrainMul)*TestSize
    if Trainsize+TestSize>len(df) and len(df)!=0 and not df.empty :
        showmessage('Limited Database Size','Database is limited and multiplier will be overridden!',TIMEOUT=3000,TYP='warn')
        TrainMUL.delete(0,"end")
        TrainMul = (len(df)-TestSize)//TestSize
        TrainMUL.insert(0,TrainMul)
        Trainsize = TrainMul*TestSize
    TrainLBL.configure(text = f'Train size is: {Trainsize}', fg='Gold')
    DBSizeBTN.configure(text = f'The size of the database utilized:{TestSize+Trainsize} candles', fg='Gold') 
    
def update_time(TZ):
    global selected_timezone
    selected_timezone = TZ
    if selected_timezone == "Tehran":
        timezone = pytz.timezone('Asia/Tehran')
    elif selected_timezone == "Tokyo":
        timezone = pytz.timezone('Asia/Tokyo')
    elif selected_timezone == "London":
        timezone = pytz.timezone('Europe/London')
    elif selected_timezone == "Sydney":
        timezone = pytz.timezone('Australia/Sydney')
    elif selected_timezone == "Newyork":
        timezone = pytz.timezone('America/New_York')
    else:
        timezone = pytz.timezone('Asia/Tokyo')
    
    Time = datetime.datetime.now(timezone)#%x %X
    formatted_time = Time.strftime(" %T  %D %A\t")+ timezone.zone
    time_label.config(text=formatted_time)
    root.after(1000, lambda:update_time(selected_timezone))

def insert_element():
    global list_box
    global WIN2
    global msg
    global Msg
    WIN2 = tk.Toplevel(root)
    WIN2.bind('<Escape>',EXIT_One)
    HEIGHT = 29*len(Msg.split("\n"))
    WIDTH = 10
    scrollbar = tk.Scrollbar(WIN2, orient="vertical")
    list_box = tk.Listbox(WIN2,width=WIDTH,height=HEIGHT, yscrollcommand=scrollbar.set)
    scrollbar.config(command=list_box.yview)
    list_box.place(x=0,y=0)
    msgsplitted = Msg.split('\n')
    CLBXBT.config(fg = 'orange')
    SendBtn.config(fg = 'green')
    for i in msgsplitted:
        list_box.insert(tk.END,i)
        if len(i)> WIDTH:
            WIDTH = 19*len(i)
    list_box.configure(width = WIDTH )
    WIN2.geometry(f'{WIDTH}x{HEIGHT}+{w//2}+{h//2}')       
 
#function to clear all the data in list_box
def clear_box():
    global list_box
    global WIN2
    list_box.delete(0,tk.END)
    CLBXBT.config(fg = 'green')
    SendBtn.config(fg ='orange')       
        
def TF_SUBMIT(event='none'):
    global TimeFrame
    TimeFrame = cmbCategories.get()
    print('TimeFrame:',TimeFrame)
    TF_MESSAGE.config(text = f'{TimeFrame} Timeframe is selected')
    
def SUBMIT(event='none'):
    global Msg
    global TimeFrame
    global Symbol
    global Strategy
    global features
    global FEAT
    global LABL
    global LoginFlag1
    global LoginFlag2
    global TimeFrameDict1
    global TimeFrameDict2
    try:
        WIN2.destroy()
    except:
        pass
    TimeFrame = cmbCategories.get()
    Sdate = SDATE.get()
    Edate = EDATE.get()
    Symbol = CURPAIR.get()
    Symbol = Symbol.upper()
    if LoginFlag1:
        Symbol = mt5.symbols_get(Symbol+'*')[0].name
        INPUT_RESULTS = f'\nCurrency pair:{mt5.symbol_info(Symbol).description}'
    else:
        INPUT_RESULTS = ''
    
    if Symbol == '' or Sdate == '' or Edate == '':
        result_label = 'Error‼ All values should be entered!'
        showmessage("Empty fields",result_label,TIMEOUT=3000,TYP='Erro')

    #else:
    result_label = f"\nSymbol: {Symbol}\nfrom: {Sdate}\nto: {Edate}\nfor {Strategy}\nwith Features: {'not selected' if features.empty else FEAT}\nwith label:{LABL}\nis selected"
    Msg = f'Timeframe: {TimeFrame} '+ INPUT_RESULTS + result_label
    showmessage('Input results',Msg,TIMEOUT=5000,TYP='info')
    try:
        timeframe = TimeFrameDict1[TimeFrame]
    except KeyError as err:
        print(err,' is not valid! Enter a valid time frame')
    else:
        print(f'TimeFrame {TimeFrame} is selected →{timeframe}.')

def Login_To_MT5():
    global SERVER
    global LOGIN
    global PASS
    global LoginFlag1
    Select_connection()
    is_logged_in = mt5.initialize(login = LOGIN ,password = PASS,server = SERVER)
    if not is_logged_in:
        intChoice2.set(0)
        LoginTXT = f'Could not be initialized!☹,\nError code ={mt5.last_error()}\nAccount: {LOGIN}\nBroker: {SERVER}\n'
        ALARM.Alarm('Error')
        showmessage("Login Error",LoginTXT,TIMEOUT=3000,TYP='Erro')
        print(LoginTXT)
        VoiceAnnounce('Error! Could not be initialized.')
        response = messagebox.askretrycancel("Login error",'Retry connecting to metatrader 5')
        if  response:
            VoiceAnnounce('Try again')
            Login_To_MT5()    
        else:
            VoiceAnnounce('Cancel')          
        LoginFlag1 = False
    else:
        LoginTXT = 'Successfully initialized and Connected to MT5 Client...☺' 
        #messagebox.showinfo(title='Log in', message = LoginTXT )
        showmessage('Log in',LoginTXT,TIMEOUT=6000,TYP='info')
        intChoice2.set(1)
        lbl2.configure(text="Logged in to metatrader",fg='light green')
        account_info = mt5.account_info()
        BALANCE = account_info.balance
        account_info_Text = f'Balance : {BALANCE} {account_info.currency}\nEquity: {account_info.equity} {account_info.currency}'
        account_info_Text1 =f'Account name: {account_info.name}\nAccount: {LOGIN}\nBroker: {SERVER}\nProfit: {account_info.profit}\n'\
        +account_info_Text+f'\nmargin_free:{account_info.margin_free}\nleverage:{account_info.leverage}'
        Account_label.config(text = account_info_Text,fg='yellow')
        print(LoginTXT)
        showmessage('Account_info',account_info_Text1,TIMEOUT=6000,TYP='info')
        VoiceAnnounce('Successful log in to Metatrader 5')        
        LoginFlag1 = True
    print((lambda x: '\nlogged in...👌🏻'if x else 'is not logged in!☹')(is_logged_in))

# Function for printing the selected listbox value(s)
"""
def selected_models():
    global Msg
    global model_no
    VoiceAnnounce('Models selected')
    # Traverse the tuple returned by curselection method and print corresponding value(s) in the listbox
    for i in lstbox.curselection():
        selection = lstbox.get(i)
        VoiceAnnounce(selection)

        Msg = Msg + '\n'+selection      
        model_no.append(selection.split('.')[0])
        if model_no ==[]:
            model_no.append('1') 
        C_Ensemble = False
        R_Ensemble = False
        print(model_no)
"""        
def select_strategy():
    global Strategy
    global MODELSSTORM
    MODELSSTORM=[]
    # Traverse the tuple returned by curselection method and print corresponding value(s) in the listbox
    for i in lstbox.curselection():
        selection = lstbox.get(i)
        MODELSSTORM.append(selection.split('→')[1].strip())
    print(MODELSSTORM)
    if len(MODELSSTORM)>1:
        GSE_enable()    
    print(Strategy)
    RUNStrategy('NormalMode')  

#CHECK BUTTONS
def choice1():
    global audiointeract
    if intChoice1.get()==1:
        lbl1.configure(text="Audio UI disabled",fg='#FF0000')
        audiointeract = False
        chkChoice1.configure(text="AUI Disabled!")
        VoiceAnnounce('Audio interaction Disabled')
    else:
        lbl1.configure(text="Audio interaction enabled!",fg='Green')
        audiointeract = True
        chkChoice1.configure(text="AUI Enabled!")
        VoiceAnnounce('Audio interaction Enabled')
        
def choice2():
    global LoginFlag1
    if intChoice2.get()==1:
        Login_To_MT5()
    else:
        LoginFlag1 = False
        lbl2.configure(text="Logged out of MT5",fg='#FF0000')
        if audiointeract:
            VoiceAnnounce('Exit from metatrader')
        mt5.shutdown()
        
def Novelsplitting():
    global SingleCandlePredict
    if intChoice3.get()==1:
        intChoice6.set(0)
        SingleCandlePredict = True
        #chkChoice3.configure(text="NovelSplit")
        VoiceAnnounce('Novel Storm Split')
    else:
        SingleCandlePredict = False
        intChoice6.set(1)
        #chkChoice3.configure(text="NormalSplit")
        VoiceAnnounce('Normal Split')
        
def Normalsplitting():
    global SingleCandlePredict
    if intChoice6.get()==1:
        intChoice3.set(0)
        SingleCandlePredict = False
        #chkChoice3.configure(text="NovelSplit")
        VoiceAnnounce('Normal Split')
    else:
        SingleCandlePredict = True
        intChoice3.set(1)
        #chkChoice3.configure(text="NormalSplit")
        VoiceAnnounce('Novel Storm Split')

def Normalrunchoice():
    global QuickRun
    if intChoice4.get() == 1:
        QuickRun = False        
        VoiceAnnounce('Normal Run')
        intChoice5.set(0)       
    else:
        QuickRun = True
        VoiceAnnounce('Quick Run')
        intChoice5.set(1)
        
def Quickchoice():
    global QuickRun
    if intChoice5.get() == 1:
        QuickRun = True
        VoiceAnnounce('Quick Run')
        intChoice4.set(0)
    else:
        QuickRun = False
        VoiceAnnounce('Normal Run')
        intChoice4.set(1)

def ConciseReport():
    global Concise_Report
    if intChoice7.get():
        Concise_Report = True
        chkChoice7.config(text='Concise Report')
        VoiceAnnounce('Concise Report')
    else:
        Concise_Report = False
        chkChoice7.config(text='Complete Report')
        VoiceAnnounce('Complete Report')

def Mute(event):
    global audiointeract
    intChoice1.set(1)
    chkChoice1.configure(text="AUI Disabled!")
    audiointeract = False  

def QuickinitialSettings1(event):
    QuickinitialSettings(1)
    
def QuickinitialSettings(M):
    global df
    global LABL
    global Lcateg
    global audiointeract
    intChoice8.set(not M)
    intChoice9.set(M)
    if intChoice9.get():
        #chkChoice8.config(text='Automatic Settings')
        VoiceAnnounce('Automatic Settings')
        Respon = messagebox.askyesno("Load saved comparisons?",'Load saved comparisons?',default='yes')    
        if Respon:
            LoadCompare()
        else:        
            df = pd.read_csv(r'.\dataset\df in Timeframe d1 from 2018-02-12 to 2025-02-11.csv')
            DataBaseBTN.config(text = 'Loaded from\nSaved file',fg='light green',font=('Times',FONTSIZE, 'bold'))
            Add_Return()
            df.set_index(df.DATE, inplace=True)
            df.drop('DATE',axis=1,inplace=True)
            df['Lottage'] = 1
            print(f'loaded offline from df head:\n{df.head(1)}')
            Account_label.config(text = f'Timeframe {TimeFrame}\nfrom {df.head(1).index[0]}\nto {df.tail(1).index[0]}', fg= 'light green')   
            LABL = 'MARKET'
            cmbLab.current(0)  #cmbLab.current(Lcateg.index(LABL))
            LB_MESSAGE1.configure(text = f'Selected label: {LABL}',fg='#F5CC21')
            Set_Label(LABL)
            intChoice8.set(0)
            root.bind('<Return>', cmbFeatFocus)
            ALARM.Alarm('Beep')
        Mute('')          
        Respon = "\nWith the values from the preset file loaded!" if Respon else ""
        showmessage('Automatic Setting',f'Automatic setting {Respon}\nSelected label: {LABL}\nLoaded from Saved file\nAUI Disabled!',TIMEOUT=2500,TYP='info')
    else:     
        #chkChoice8.config(text='Manual Settings')
        intChoice8.set(1)
        intChoice9.set(0)
        intChoice8.set(1)
        lbl1.configure(text="Audio interaction enabled!",fg='Green')
        audiointeract = True
        chkChoice1.configure(text="AUI Enabled!")
        VoiceAnnounce('Manual Settings')
        
#============================== Model selection?
def showmessage(TITLE,MESSAGE,TIMEOUT=1000,TYP='info'):
    win = tk.Tk()
    win.withdraw()
    win.after(TIMEOUT,win.destroy)
    win.option_add('*Dialog.msg.font', 'Helvetica 20')
    if TYP=='info':
        messagebox.showinfo(title=TITLE, message = MESSAGE,master=win)
    elif TYP=='warn':
        messagebox.showwarning(title=TITLE, message = MESSAGE,master=win)
    elif TYP=='Erro':
        messagebox.showerror(title=TITLE, message = MESSAGE,master=win)
    
def selected_Visual(): 
    A = VISlstbox.curselection()
    M = ''
    for i in VISlstbox.curselection():
        M += str(VISlstbox.get(i))
    if M =='':
        M = 'Nothing'
    VoiceAnnounce('Save '+M)
    if 0 in A:
        CLOSE_Graph(False)
        CLOSE_LOG_Graph(False)
        Graph_Autocorr(False)
        ScaledReturn_Graph(False)
        Graph_Feat_Corr(False)
        autocorr_plot(False)
        Scatter_Plot(False)
        
    if 1 in A:
        Confusion_Matrix_Plot(False)
        Equity_Plot(False)
        DrawDown_Plot(False)
        Histogram_Plot(False)

def RECORD(duration=2):
    recording = sd.rec(int(duration*freq),samplerate=freq,channels=2)
    sd.wait()
    write("Answers.wav", freq, recording)
    wv.write("Answers1.wav", recording, freq, sampwidth=2)

def stt():
    audio_file = sr.AudioFile('Answers1.wav')
    with audio_file as source:
        r.adjust_for_ambient_noise(source)
        audio = r.record(source)
        try:
            VoiceAnswer = r.recognize_google(audio)
        except:
            print('No response!')
            return 'No response!'#'yes'
        else:
            print(VoiceAnswer)
            return VoiceAnswer

def Select_connection():
    global LoginFlag1
    global WIN8
    def ResetLogin():
        global SERVER
        global LOGIN
        global PASS
        global WIN8
        SERVER = broker_entry.get()
        LOGIN = int(account_login_entry.get())
        PASS = account_password_entry.get()
        print(SERVER,LOGIN,PASS)
        WIN8.destroy()
        WIN8.quit()
        
    Response1 = messagebox.askyesno("Connection settings?",'Hold previous connection settings?',default='yes')    
    if not Response1:
        WIN8 = tk.Toplevel(root)
        WIN8.bind('<Escape>',EXIT_One)
        BROK = tk.StringVar()
        USER = tk.StringVar()
        PASS = tk.StringVar()
        BROLBL = tk.Label(WIN8,text="Broker: ",font=('times',FONTSIZE,'bold'))
        BROLBL.grid(row=0,column=0,sticky='W',pady=5)

        broker_entry = tk.Entry(WIN8,textvariable=BROK,width=15,borderwidth=3,bd=5)
        broker_entry.grid(row=0,column=1,sticky='W',pady=5)
        
        USELBL = tk.Label(WIN8,text="Username: ",font=('times',FONTSIZE,'bold'),bd=5)
        USELBL.grid(row=1,column=0,pady=5)
        
        account_login_entry = tk.Entry(WIN8,textvariable=USER,width=15,borderwidth=3,bd=5)
        account_login_entry.grid(row=1,column=1,pady=5)
            
        PASLBL = tk.Label(WIN8,text="Password: ",font=('times',FONTSIZE,'bold'),bd=5)
        PASLBL.grid(row=2,column=0,pady=5)

        account_password_entry = tk.Entry(WIN8,textvariable=PASS,width=15,borderwidth=3,bd=5)
        account_password_entry.grid(row=2,column=1,pady=5)

        WIN8.title("Connect to MT5")
        bt1 = tk.Button(WIN8,text="Reset connection settings",font=("Georgia",FONTSIZE),relief=tk.RAISED,bd=5,command=ResetLogin)
        bt1.grid(row=3,columnspan=2)
        WIN8.mainloop()

def LoadDataBase(_):
    global df
    global LoginFlag1
    global LoginFlag2
    global Lottage
    #global comparisondf
    VoiceAnnounce('Loading database')
    Response = messagebox.askyesnocancel("Load database online",'Connect to online database?',default=messagebox.NO)    
    if Response == None:
        showmessage('Cancel loading','Cancel database loading!',TIMEOUT=2000,TYP='warn')
        DataBaseBTN.config(text = 'DataBase \nnot Loaded!',fg='red',font=('Times',FONTSIZE, 'bold'))
        VoiceAnnounce('DataBase is not Loaded!')
        ALARM.Alarm('Error')
        return
    elif Response:
        Total_Reset()
        load_online_Database()
    elif not Response:
        Total_Reset()
        load_offline_Database()
        df['MARKET'] = df['MARKET'].astype('int8')
    root.bind('<Return>', cmbLabFocus)
    cmbLab.focus()
    Lottage = 'Static'
    GETTrainSIZE()
    
def load_offline_Database():
    global df
    global WIN3
    b=''
    global Source
    Source = 'Offline'
    def select_radio_DB1():
        a = tk.StringVar(WIN3,'none')
        nonlocal b
        def RDchoice(p):
            nonlocal a , b
            b = a.get()
            SelBTN.config(text=f'Your choice is {b}')
            if p:
                WIN3.destroy()
                WIN3.quit()
                return b
            
        def navigate_RadioButtons(event):
            # Get the current selection index  
            current_index = Options.index(a.get())  
            
            if event.keysym == 'Down':  
                # Move down in the list  
                current_index = (current_index + 1) % 3 
            elif event.keysym == 'Up':  
                # Move up in the list  
                current_index = (current_index - 1) % 3 
            
            # Update the selected option  
            a.set(Options[current_index]) 
        Options = ['Directly from MT5',  'Saved file', 'Browse saved file']     
        RadioBut1 = tk.Radiobutton(WIN3,text='Load '+ Options[0],indicator = 0,background = "lightblue1",justify=tk.LEFT,variable=a,value=Options[0],command=lambda:RDchoice(0))
        RadioBut1.pack(fill = tk.X, ipady = 5)#.pack(anchor=tk.W)
        
        RadioBut2 = tk.Radiobutton(WIN3,text='Load '+ Options[1],indicator = 0,background = "lightsteelblue1",justify=tk.LEFT,variable=a,value=Options[1],command=lambda:RDchoice(0))
        RadioBut2.pack(fill = tk.X, ipady = 5)##.place(height=20,width=200,x=0,y=20)

        RadioBut3 = tk.Radiobutton(WIN3,text= Options[2],indicator = 0,background = "lightskyblue1",justify=tk.LEFT,variable=a,value=Options[2],command=lambda:RDchoice(0))
        RadioBut3.pack(fill = tk.X, ipady = 5)##.place(height=20,width=200,x=0,y=20)

        SelBTN = tk.Button(WIN3,text="↑Select one of these databases↑",bg='slateblue1',relief="raised",command=lambda:RDchoice(1))
        SelBTN.pack(fill = tk.X, ipady = 5)#.pack(anchor=tk.CENTER)
        a.set('Saved file')
        SelBTN.focus()
        WIN3.bind('<Down>', navigate_RadioButtons)  
        WIN3.bind('<Up>', navigate_RadioButtons)
        WIN3.bind('<Return>',RDchoice)
        WIN3.mainloop()

    WIN3 = tk.Toplevel(root)
    WIN3.bind('<Escape>',EXIT_One)
    #WIN3.geometry(f'+{w//2}+{h//5}')
    select_radio_DB1()
    if b == 'Directly from MT5':
        try:
            df = pd.read_csv('C:\\Users\\TOP\\AppData\\Roaming\\MetaQuotes\\Terminal\\2E8DC23981084565FA3E19C061F586B2\\MQL4\\Files\\DataSet.csv')
            ALARM.Alarm('Beep')
            DataBaseBTN.config(text = 'Loaded Directly \nfrom MT5',fg='light green',font=('Times',FONTSIZE, 'bold'))
            VoiceAnnounce('Loaded Directly from meta trader 5!')
            Add_Return()
        except FileNotFoundError as ERR:
            showmessage('Path Error',f'Path is not correct or available!\n{ERR}',TIMEOUT=3000,TYP='Erro')
            return 0
        
    elif b == 'Saved file':
        try:
            df = pd.read_csv(r'.\dataset\df in Timeframe d1 from 2018-02-12 to 2025-02-11.csv')
            ALARM.Alarm('Beep')
            DataBaseBTN.config(text = 'Loaded from\n'+b,fg='light green',font=('Times',FONTSIZE, 'bold'))
            VoiceAnnounce('Loaded from '+b)
            Add_Return()
        except FileNotFoundError as ERR:
            showmessage('Path Error',f'Path is not correct or available!\n{ERR}',TIMEOUT=3000,TYP='Erro')
            return 0
        
    elif b == 'Browse saved file': 
        try:
            df = OpenDatabase()
            ALARM.Alarm('Beep')
            DataBaseBTN.config(text = 'Loaded from\n'+b,fg='light green',font=('Times',FONTSIZE, 'bold'))
            Add_Return()
        except FileNotFoundError as ERR:
            showmessage('Path Error',f'Path is not correct or available!\n{ERR}',TIMEOUT=3000,TYP='Erro')
            return 0      
    else:        
        showmessage('Undefined Error!','Undefined Error!',TIMEOUT=3000,TYP='Erro')
        return 0
        
    df.set_index(df.DATE, inplace=True)
    df.drop('DATE',axis=1,inplace=True)
    df['Lottage'] = 1
    print(f'loaded offline from df head:\n{df.head(1)}')
    Account_label.config(text = f'Timeframe {TimeFrame}\nfrom {df.head(1).index[0]}\nto {df.tail(1).index[0]}', fg= 'light green')   
    
def load_online_Database():
    global df
    global TimeFrame
    global Sdate
    global Edate
    global Sdate0
    global Edate0
    global LoginFlag1
    global LoginFlag2
    global TimeFrameDict1
    global TimeFrameDict2
    global Source
    b=''
    def LoadingfromMT5todf(rates):
        global df
        df = pd.DataFrame(rates)
        if df.empty:
            DataBaseBTN.config(text = 'DataBase \nnot Loaded!',fg='red',font=('Times',FONTSIZE, 'bold'))
            ALARM.Alarm('Error')
            Emptydf = 'df is empty ► DataBase not Loaded!\n► Check the Entries or internet connection!'
            showmessage('Empty df',Emptydf,TIMEOUT=3000,TYP='Erro')
            VoiceAnnounce('df is empty! DataBase not Loaded!')
        else:
            DataBaseBTN.config(text = 'DataBase Loaded\nfrom MT5',fg='light Green',font=('Times',FONTSIZE, 'bold'))
            ALARM.Alarm('Beep')
            print('DataBase successfully loaded from MT5')
            Source = 'MT5'
            VoiceAnnounce('DataBase successfully loaded from Meta trader 5.')
            df.rename(columns={'open':'OPEN','high':'HIGH','low':'LOW','close':'CLOSE'},inplace=True)
            Modify_Date()
            df.set_index(df.DATE , inplace=True)
            df.drop('DATE',axis=1,inplace=True)
            df.drop('real_volume',axis=1,inplace=True)
            df['difference'] = df.CLOSE - df.OPEN
            df['MARKET'] = df['difference'].apply(lambda x: 1 if x>0.0 else 0)
            df['MARKET'] = df['MARKET'].astype('int8')
            Add_Return()
            df['Lottage'] = 1
    def select_radio_DB1():
        global LoginFlag1
        global LoginFlag2
        a = tk.StringVar(WIN3,'none')
        nonlocal b
        def RDchoice(p):
            nonlocal a , b
            b = a.get()
            SelBTN.config(text=f'Your choice is {b}')
            if p:
                WIN3.destroy()
                WIN3.quit()
                return b
        RadioBut1 = tk.Radiobutton(WIN3,text='Load online from MT5?',indicator = 0,background = "light blue",justify=tk.LEFT,variable=a,value='Online from MT5',command=lambda:RDchoice(0))
        RadioBut1.pack(fill = tk.X, ipady = 5)#.pack(anchor=tk.W)
        
        RadioBut2 = tk.Radiobutton(WIN3,text='Load online from Yfinance?',indicator = 0,background = "light blue",justify=tk.LEFT,variable=a,value='Online from Yfinance',command=lambda:RDchoice(0))
        RadioBut2.pack(fill = tk.X, ipady = 5)##.place(height=20,width=200,x=0,y=20)  
        SelBTN = tk.Button(WIN3,text="↑Select one of upper databases↑",bg='yellow',relief="raised",command=lambda:RDchoice(1))
        SelBTN.pack(fill = tk.X, ipady = 5)
        a.set('Online from MT5')
        SelBTN.focus()
        WIN3.bind('<Return>',RDchoice)
        WIN3.mainloop()
    WIN3 = tk.Toplevel(root)
    WIN3.bind('<Escape>',EXIT_One)
    select_radio_DB1()
    try:
        if b == 'Online from MT5':
            if not LoginFlag1:
                Login_To_MT5()
                print('hi')
                if not LoginFlag1:
                    print('hello')
                    print('LoginFlag1:',LoginFlag1)
                    response = messagebox.askretrycancel("Could not login to metatrader",'Retry login to metatrader?')
                    if  response:
                        VoiceAnnounce('Try again')
                        load_online_Database()
                    else:
                        VoiceAnnounce('Cancel')
                        return False
                    
                #DataBaseBTN.config(text = 'Try to Load Again!',fg='red',font=('Times',FONTSIZE, 'bold'))
                #ALARM.Alarm('Error')
                #Emptydf = 'df is empty → DataBase not Loaded!\n► Checkentries\n► Check internet connection!'
                #showmessage('Empty df',Emptydf,TIMEOUT=3000,TYP='Erro')
                #print(Emptydf)
                #return False
                else:
                    print('how are you')
                    try:
                        rates = mt5.copy_rates_range(Symbol, TimeFrameDict1[TimeFrame], Sdate, Edate)
                    except SystemError as ERRS:
                        DataBaseBTN.config(text = 'DataBase \nnot Loaded!',fg='red',font=('Times',FONTSIZE, 'bold'))
                        ALARM.Alarm('Error')
                        Emptydf = 'df is empty ► DataBase not Loaded!\n► May be out of range!'
                        global Sdate0
                        if eval(Sdate0[:4]) <1971:
                            SDATE.delete(0, 'end') 
                            SDATE.insert(0,'1971'+Sdate0[4:])
                            Sdate = SDATE.get()
                            Sdatebutton.config(text='Modified Start date',fg='light green')
                        showmessage('Empty df',Emptydf,TIMEOUT=3000,TYP='Erro')
                        VoiceAnnounce('not Loaded! May be out of range!')
                    else:
                        LoadingfromMT5todf(rates)  
            else:
                rates = mt5.copy_rates_range(Symbol, TimeFrameDict1[TimeFrame], Sdate, Edate)
                LoadingfromMT5todf(rates)

        elif b == 'Online from Yfinance':
            import yfinance
            #df = yf.download(Symbol, start = Sdate0, end = Edate)
            LoginFlag2 = True
            try:
                df = yfinance.download(f'{Symbol}=X' ,interval= TimeFrameDict2[TimeFrame],start=Sdate.strftime('%Y-%m-%d'), end=Edate.strftime('%Y-%m-%d'),auto_adjust=False)
            except AttributeError as ERR:
                showmessage("Load Error!",f'All fields should be entered\nData frame not loaded {b}\n{ERR}',TIMEOUT=3000,TYP='Erro')
                VoiceAnnounce(f'All fields should be entered. Dataframe not loaded {b}')
                return
            print(f"{Symbol}=X ,interval= {TimeFrameDict2[TimeFrame]},start={Sdate.strftime('%Y-%m-%d')}, end={Edate.strftime('%Y-%m-%d')}")
            if df.empty:
                b = 'Empty DataFrame'
                DataBaseBTN.config(text = b,fg='red',font=('Times',FONTSIZE, 'bold'))
                ALARM.Alarm('Error')
            else:
                Account_label.config(text = 'Online data\nfrom Yfinance',fg='yellow')
                ALARM.Alarm('Beep')
                DataBaseBTN.config(text = 'DataBase Loaded\nfrom Yfinance!',fg='light Green',font=('Times',FONTSIZE, 'bold'))
                VoiceAnnounce('DataBase Loaded from Yfinance!')
                print('Loaded',b)
                Source = 'Yfinance'
                df.columns = df.columns.get_level_values(0)
                df.rename(columns={'Open':'OPEN','High':'HIGH','Low':'LOW','Close':'CLOSE'},inplace=True)
                df['difference'] = df.CLOSE - df.OPEN
                df['MARKET'] = df['difference'].apply(lambda x: 1 if x>0.0 else 0)
                df['MARKET'] = df['MARKET'].astype('int8')
                df['Lottage'] = 1 
        else:
            b = 'DataBase load Error\nnot Loaded!'
            DataBaseBTN.config(text = b,fg='red',font=('Times',FONTSIZE, 'bold'))
            ALARM.Alarm('Error')
    except NameError as ERR:
        showmessage("Empty dataframe",f'All fields should be entered\nnot loaded {b}\n{ERR}',TIMEOUT=3000,TYP='Erro')
    else:
        print(b)
                       
#############################################################   
def Show_DataBase():
    global WIN4
    global df
    try:
        dftext = f'Whole Dataframe df:\n{df.to_string}'
    except NameError as ERR:
        showmessage('Dataframe Error',f'Dataframe is not available\n{ERR}',TIMEOUT=3000,TYP='Erro')
        VoiceAnnounce('Dataframe is not available!')
        LoadDataBase('')  
    else:
        winsound.PlaySound(MusicPath,winsound.SND_FILENAME)
        print(dftext)
        Show_In_Win('Show Database df',dftext,BG='slateblue4',FG='light goldenrod')

def Show_DB():
    global df
    try:
        dftext = f'Dataframe df:\n{df}'#.to_string()
    except NameError as ERR:
        showmessage('Dataframe Error',f'Dataframe is not available\n{ERR}',TIMEOUT=3000,TYP='Erro')
        VoiceAnnounce('Dataframe is not available!')
        LoadDataBase('')
    else:
        winsound.PlaySound(MusicPath,winsound.SND_FILENAME)
        print(dftext)
        Show_In_Win('Show Database df',dftext,BG='slateblue4',FG='light goldenrod')

def Show_Features():
    global WIN14
    global WIN15
    global features 
    try:
        print(f'Features:\n{features}')
        winsound.PlaySound(MusicPath,winsound.SND_FILENAME)
    except NameError as ERR:
        showmessage('Features Error',f'No feature is available to display\n{str(ERR)}',TIMEOUT=3000,TYP='Erro')
        WIN15 = tk.Toplevel(root)
        WIN15.bind('<Escape>',EXIT_One)
        WIN15.attributes('-topmost','true')
        feature_Select(WIN15,4)         
    else:
        WIN14 = tk.Toplevel(root)
        WIN14.bind('<Escape>',EXIT_One)
        #WIN5.geometry('1200x400')
        text1 = tk.Text(WIN14,wrap='word',bg='#000000',fg='light yellow',font=('times',FONTSIZE,'bold'))
        text1.grid(row=0,column=0)
        text1.insert(tk.END,features)
        scroll1 = ttk.Scrollbar(WIN14,orient=tk.VERTICAL,command=text1.yview)
        scroll1.grid(row=0,column=1,sticky='ns')
        text1.config(yscrollcommand=scroll1.set)
        
        scroll2 = ttk.Scrollbar(WIN14,orient=tk.HORIZONTAL,command=text1.xview)
        scroll2.grid(row=1,column=0,sticky='ew')
        text1.config(xscrollcommand=scroll2.set)

def Show_Label():
    global WIN15
    global labels
    global LABL
    try:
        print('labels:',labels,'\nLABL:',LABL)
        labletext = f"Label set to {LABL} and equal to:\n{df['LABELS']}"#.to_string()
        winsound.PlaySound(MusicPath,winsound.SND_FILENAME)
        Show_In_Win(4000,'lable text',labletext) 
    except (NameError,KeyError) as ERR:
        showmessage('Label Error',f'No Label is available to display!\n{ERR}',TIMEOUT=3000,TYP='Erro')
        global WIN16
        WIN16 = tk.Toplevel(root)
        WIN16.bind('<Escape>',EXIT_One)
        LstrCateg = ''
        Label_Select(WIN16)
        WIN16.lift()
        WIN16.focus_force()
        
def Show_Columns(S):
    global WIN7
    global df
    global features
    try:
        print('Columns:\n',S.columns.to_list())
        winsound.PlaySound(MusicPath,winsound.SND_FILENAME)
    except NameError as ERR:
        showmessage('Dataframe Error',f'Dataframe or column is not available!\n{ERR}',TIMEOUT=3000,TYP='Erro')
        LoadDataBase('')
    else:            
        WIN7 = tk.Toplevel(root)
        WIN7.bind('<Escape>',EXIT_One)
        DB_Message = tk.Message(WIN7,text=f'Columns:{S.columns.to_list()}',justify=tk.LEFT,width=1000,
            fg='#F5CC21',bg = '#08162f',font=('Times', FONTSIZE, 'bold'),anchor = 'w',aspect=1000)
        DB_Message.pack(fill=tk.X)

def Modify_Date():
    global df
    global LoginFlag1
    global LoginFlag2
    print('TimeFrame:',TimeFrame)
    print('TimeFrameDict1[TimeFrame]:',TimeFrameDict1[TimeFrame])   
    try:  
        df      
    except NameError as ERR:
        showmessage('Dataframe Error',f'Dataframe is not available!\n{ERR}',TIMEOUT=3000,TYP='Erro')
        LoadDataBase('')       
    if LoginFlag1: 
        def Time_Date(t):
            TIME = datetime.datetime.fromtimestamp(t)
            if TimeFrameDict1[TimeFrame] >=16408:
                return TIME.strftime('%Y-%m-%d')
            else:
                return TIME.strftime('%Y-%m-%d-%H-%M')
        df['DATE'] = df.time.apply(Time_Date)
        df.drop(['time'],axis=1,inplace =True)
        
    if LoginFlag2:
        df['DATE']=df.index

    cols = df.columns.to_list()
    Columns1 = cols[-1:]+cols[0:-1]
    df = df[Columns1]

def HoldOHLC():
    global df
    df = df[['OPEN','HIGH','LOW','CLOSE']]      

def DropNAN():
    global df
    try:
        df
    except NameError as ERR:
        showmessage('Dataframe Error',f'Dataframe is not available14!\n{ERR}',TIMEOUT=3000,TYP='Erro')
        LoadDataBase('')
    else:
        df.dropna(axis=0,inplace=True)
        
def OpenDatabase():
    global TimeFrame
    file_path = filedialog.askopenfilename(initialdir=r'.\data\Databases',defaultextension=".csv", filetypes=[("Database files", "*.csv"),("Text files", "*.txt") ,("Excel files", "*.xlsx"), ("All files", "*.*")],initialfile="Database files")
    print(file_path[-4:].lower())
    TimeFrame = file_path.lower().split('timeframe')[1].split('from')[0].strip()
    print('Timeframe=',TimeFrame)
    cmbCategories.set(TimeFrame)
    TF_MESSAGE.config(text = f'{TimeFrame} Timeframe is selected')
    if file_path[-4:].lower()=='.csv':
        Loadedfile = pd.read_csv(file_path)
        content = Loadedfile
    elif file_path[-4:].lower()=='.txt':
        Loadedfile = open(file_path,'r')
        content = Loadedfile.read()       
    return content

def save_file(s,S):
    try:
        S
    except NameError as ERR:
        showmessage(f'{s} Dataframe Error',f'{s} dataframe is not available!\n{ERR}',TIMEOUT=3000,TYP='Erro')
        LoadDataBase('')
    else:
        Savepath = f'.\\data\\Databases\\{TimeFrame}'
        if not os.path.exists(Savepath):
            showmessage('Finding path',f'The path {Savepath} was not found!\nIt will be created!',TIMEOUT= 3000,TYP='warn')
            os.makedirs(Savepath)

        file_path = filedialog.asksaveasfilename(initialdir=Savepath,defaultextension=".csv", filetypes=[("Database files", "*.csv"),("Text files", "*.txt") ,("Excel files", "*.xlsx"), ("All files", "*.*")],initialfile=f'{Source}_{s} in Timeframe {TimeFrame} from {df.head(1).index[0]} to {df.tail(1).index[0]}')
        if file_path:
            print(file_path)
            try:
                if file_path[-4:].lower()=='xlsx':
                    S.to_excel(file_path)
                else:
                    S.to_csv(file_path,sep=',')
                #messagebox.showinfo("Save", f"{s} database Saved successfully")
                showmessage('Save',f"{s} database Saved successfully",TIMEOUT=1500,TYP='info')
            except PermissionError as ERR:
                showmessage('Permission Error',f'Check if file is already opened!\n{ERR}',TIMEOUT=3000,TYP='Erro')
        else:
            showmessage("Save", f"{Source}_{s} database wasn't Saved!",TIMEOUT=2000,TYP='warn')

def indicator_features():
    global features
    global FEAT
    features = df[[FEAT]]
    features = features.shift()
    features.fillna(0,inplace=True)  
    

def Set_OHLC_Features():
    global df
    global features
    try:
        df
    except NameError as ERR:
        showmessage('Dataframe Error',f'Dataframe is not available!\n{ERR}',TIMEOUT=3000,TYP='Erro')
        LoadDataBase('')
        
    features = df[['OPEN','HIGH','LOW','CLOSE']]
    features = features.shift()
    #features.dropna(axis=0,inplace=True)
    features.fillna(0,inplace=True)    
#     Select_Scaling_Type()
    
def Just_Close_Features(COMPARING):
    global df
    global features
    columns = []
    try:
        df
    except NameError as ERR:
        showmessage('Dataframe Error',f'Dataframe is not available!\n{ERR}',TIMEOUT=3000,TYP='Erro')
        LoadDataBase('')
    try:
        df['LABELS']
    except KeyError:
        showmessage('Label not available',f'Label is not available!',TIMEOUT=3000,TYP='Erro')
        global WIN16
        WIN16 = tk.Toplevel(root)
        WIN16.bind('<Escape>',EXIT_One)
        LstrCateg = ''
        Label_Select(WIN16)
        WIN16.mainloop()
        WIN16.lift()
        WIN16.focus_force()
        
    if not COMPARING:
        n_prev_samples = Add_previous('CLOSE') 
    else:
        n_prev_samples = COMPARING
    for c in range(1,n_prev_samples+1):
        df['CLOSE('+str(-c)+')']=df['CLOSE'].shift(c)
        df.fillna(0,inplace=True)
        columns +=['CLOSE('+str(-c)+')']
    showmessage(f'No. of Closes',f'{n_prev_samples} of previous Closes are added to df database!',TIMEOUT=2000,TYP='info')
    features = df[columns]
    return n_prev_samples


def Just_Label_Features(COMPARING):
    global features
    columns = []
    try:
        df
    except NameError as ERR:
        showmessage('Dataframe Error',f'Dataframe is not available\n{ERR}',TIMEOUT=3000,TYP='Erro')
        LoadDataBase('')
    try:
        df['LABELS']
    except KeyError:
        showmessage('Label not available',f'Label is not available!',TIMEOUT=3000,TYP='Erro')
        global WIN16
        WIN16 = tk.Toplevel(root)
        WIN16.bind('<Escape>',EXIT_One)
        LstrCateg = ''
        Label_Select(WIN16)
        WIN16.mainloop()
        WIN16.lift()
        WIN16.focus_force()
        
    if not COMPARING:
        n_prev_samples = Add_previous('LABELS') 
    else:
        n_prev_samples = COMPARING
    print('Size of lag is selected:',n_prev_samples)
    for c in range(1,n_prev_samples+1):
        df['LABELS('+str(-c)+')']=df['LABELS'].shift(c)
        df.fillna(0,inplace=True)
        columns +=['LABELS('+str(-c)+')']
    showmessage(f'No. of LABELS',f'{n_prev_samples} of previous LABELS are added to df database!',TIMEOUT=2000,TYP='info')      
    features = df[columns]
    return n_prev_samples

def All_Features(m):
    global df
    global features
    try:
        df
    except NameError as ERR:
        showmessage('Dataframe Error',f'Dataframe is not available!\n{ERR}',TIMEOUT=3000,TYP='Erro')
        LoadDataBase('')  
    #df.drop('MARKET',axis=1,inplace=True)   #columns = list(df.columns.values)[:-1]
    columns = df.columns.to_list()    
    if m:
        features = df.drop(df.columns[df.columns.get_loc('Return'):].tolist(),axis=1)
    else:
        features = df.copy()
    try:
        features.drop(['tick_volume'],axis=1,inplace=True)
    except:
        pass

    try:
        features.drop(['spread'],axis=1,inplace=True)
    except:
        pass
    
    try:
        features.drop(['difference'],axis=1,inplace=True)
    except:
        pass
    
    try:
        features.drop(['Lottage'],axis=1,inplace=True)
    except:
        pass
    
    try:
        features.drop(['predictedclass'],axis=1,inplace=True)
    except:
        pass

    try:
        features.drop(['Compare'],axis=1,inplace=True)
    except:
        pass

    try:
        features.drop(['Win/Loss'],axis=1,inplace=True)
    except:
        pass

    try:
        features.drop(['profit'],axis=1,inplace=True)
    except:
        pass

    try:
        features.drop(['equity'],axis=1,inplace=True)
    except:
        pass

    try:
        features.drop(['DrawDown'],axis=1,inplace=True)
    except:
        pass    
#     except (AttributeError,NameError) as ERR:
#         showmessage('Features Error',f'No feature was available to display\nSelect features first!\n{ERR}',TIMEOUT=4000,TYP='Erro')
#     except KeyError as ERR:
#         showmessage('Column Error2',f'Column spread is not in dataframe df\n{ERR}',TIMEOUT=3000,TYP='warn')          
    features = features.shift()
    features.fillna(0,inplace=True)

def All_Features_Except_Labels():
    global df
    global features
    All_Features(False)
    try:
        features.drop(['MARKET'],axis=1,inplace=True)         
    except (AttributeError,NameError) as ERR:
        showmessage('Features Error',f'No feature was available to display\nSelect features first!\nOHLC Features\n{ERR}',TIMEOUT=4000,TYP='Erro')
    except KeyError as ERR:
        showmessage('Column Error',f'Column MARKET is not in dataframe df\n{ERR}',TIMEOUT=3000,TYP='warn')
        pass

    try:
        features.drop(['LABELS'],axis=1,inplace=True)            
    except (AttributeError,NameError) as ERR:
        showmessage('Features Error',f'No feature was available to display\nSelect features first!\nOHLC Features\n{ERR}',TIMEOUT=4000,TYP='Erro')
    except KeyError as ERR:
        #showmessage('Column Error',f'Column MARKET is not in dataframe df\n{ERR}',TIMEOUT=3000,TYP='warn')
        pass
    
    try:
        features.drop(['Rise_fall'],axis=1,inplace=True)         
    except (AttributeError,NameError) as ERR:
        showmessage('Features Error',f'No feature was available to display\nSelect features first!\nOHLC Features\n{ERR}',TIMEOUT=4000,TYP='Erro')
    except KeyError as ERR:
#        showmessage('Column Error1',f'Column MARKET is not in dataframe df\n{ERR}',TIMEOUT=3000,TYP='warn')
        pass

    try:
        features.drop(['Return'],axis=1,inplace=True)          
    except (AttributeError,NameError) as ERR:
        showmessage('Features Error',f'No feature was available to display\nSelect features first!\nOHLC Features\n{ERR}',TIMEOUT=4000,TYP='Erro')
    except KeyError as ERR:
        #showmessage('Column Error1',f'Column MARKET is not in dataframe df\n{ERR}',TIMEOUT=3000,TYP='warn')
        pass
    
def Custom_Features():
    global df
    global features
    global WIN12
    global DynamicColumns
    global OnetimeSetting
    SelectedColumns = []
    M = ''
    def toggleOnetimeSetting():
        global OnetimeSetting
        OnetimeSetting = not OnetimeSetting
        OneTimeBTN.configure(text=('Features non changeable' if OnetimeSetting else 'Features are changeable'))
        
    def selected_features(_):
        global OnetimeSetting
        global WIN12
        nonlocal SelectedColumns
        nonlocal M
        nonlocal lsbox
        global features
        if not OnetimeSetting:
            for i in lsbox.curselection():
                selection = lsbox.get(i)
                M+=selection
                SelectedColumns.append(selection.split('.')[0])
            print('SelectedColumns:',SelectedColumns)
            print('features are:',M)
            features = df[SelectedColumns]        
            features = features.shift()
            #features.dropna(axis=0,inplace=True)
            features.fillna(0,inplace=True)
    #         Select_Scaling_Type()
        Fea_MESSAGE1.configure(text = f'{FEAT}\n{features.columns.values}',fg='#F5CC21')
        WIN12.destroy()
        WIN12.quit()
    try:
        DynamicColumns = df.columns.to_list()      
    except NameError as ERR:
        showmessage('Dataframe Error',f'Dataframe is not available!\n{ERR}',TIMEOUT=3000,TYP='Erro')
        LoadDataBase('')
    else:
        try:
            DynamicColumns.remove('MARKET')
        except KeyError as ERR:
            pass
        try:
            DynamicColumns.remove('Rise_fall')
        except KeyError as ERR:
            pass
        try:
            DynamicColumns.remove('LABELS')
        except (KeyError,ValueError) as ERR:
            #showmessage('Column Error',f'Column LABELS is not in dataframe df\n{ERR}',TIMEOUT=2000,TYP='warn')
            pass
        try:
            DynamicColumns.remove('spread')
        except (KeyError,ValueError) as ERR:
            #showmessage('Column Error',f'Column spread is not in dataframe df\n{ERR}',TIMEOUT=2000,TYP='warn')
            pass
        try:
            DynamicColumns.remove('Lottage')
        except (KeyError,ValueError) as ERR:
            #showmessage('Column Error',f'Column Lottage is not in dataframe df\n{ERR}',TIMEOUT=2000,TYP='warn')
            pass
        try:
            DynamicColumns.remove('tick_volume')
        except (KeyError,ValueError) as ERR:
            #showmessage('Column Error',f'Column tick_volume is not in dataframe df\n{ERR}',TIMEOUT=2000,TYP='warn')
            pass
        
        try:
            DynamicColumns.remove('predicted')
        except (KeyError,ValueError) as ERR:
            #showmessage('Column Error',f'Column predicted is not in dataframe df\n{ERR}',TIMEOUT=2000,TYP='warn')
            pass
        
        try:
            DynamicColumns.remove('predictedclass')
        except (KeyError,ValueError) as ERR:
            #showmessage('Column Error',f'Column predictedclass is not in dataframe df\n{ERR}',TIMEOUT=2000,TYP='warn')
            pass
        
        try:
            DynamicColumns.remove('profit')
        except (KeyError,ValueError) as ERR:
            #showmessage('Column Error',f'Column profit is not in dataframe df\n{ERR}',TIMEOUT=2000,TYP='warn')
            pass
        
        try:
            DynamicColumns.remove('equity')
        except (KeyError,ValueError) as ERR:
            #showmessage('Column Error',f'Column equity is not in dataframe df\n{ERR}',TIMEOUT=2000,TYP='warn')
            pass
        
        try:
            DynamicColumns.remove('difference')
        except (KeyError,ValueError) as ERR:
            #showmessage('Column Error',f'Column difference is not in dataframe df\n{ERR}',TIMEOUT=2000,TYP='warn')
            pass
        
        try:
            DynamicColumns.remove('Compare')
        except (KeyError,ValueError) as ERR:
            #showmessage('Column Error',f'Column Compare is not in dataframe df\n{ERR}',TIMEOUT=2000,TYP='warn')
            pass
        
        try:
            DynamicColumns.remove('Win/Loss')
        except (KeyError,ValueError) as ERR:
            #showmessage('Column Error',f'Column DrawDown is not in dataframe df\n{ERR}',TIMEOUT=2000,TYP='warn')
            pass
        
        try:
            DynamicColumns.remove('DrawDown')
        except (KeyError,ValueError) as ERR:
            #showmessage('Column Error',f'Column DrawDown is not in dataframe df\n{ERR}',TIMEOUT=2000,TYP='warn')
            pass
        
        WIN12 = tk.Toplevel(root)
        WIN12.bind('<Escape>',EXIT_One)
        WIN12.bind('<Return>',selected_features)
        OneTimeBTN = tk.Button(WIN12, text='Select features one time',font=('Times', FONTSIZE, 'bold'),fg='yellow', bg='Navy blue',command = toggleOnetimeSetting)  
        OneTimeBTN.pack(fill = tk.BOTH,expand=1)
        FeatBTN = tk.Button(WIN12, text='Select the features',font=('Times', FONTSIZE, 'bold'),fg='yellow', bg='mediumvioletred',command = lambda:selected_features(''))  
        FeatBTN.pack(fill = tk.BOTH,expand=1)
        featvar = tk.Variable(value=DynamicColumns)
        lsbox = tk.Listbox(WIN12, width=40, height=len(DynamicColumns)+1,listvariable=featvar, selectmode = tk.MULTIPLE)
        lsbox.pack(fill = tk.BOTH,expand = 1)        
        WIN12.mainloop()

def Custom_Window_Features():
    global df
    global WIN17
    global DynamicColumns
    global features
    SelectedColumns = []
    M = ''
    def selected_features(_):
        global WIN17
        nonlocal SelectedColumns
        nonlocal M
        nonlocal lsbox
        global DynamicColumns
        global features
        for i in lsbox.curselection():
            selection = lsbox.get(i)
            M+=selection
            M += ','
            SelectedColumns.append(selection.split(',')[0])
            print('SelectedColumns:',SelectedColumns)
        
        print('Features are:', M)
        DynamicColumns = SelectedColumns
        no = Window_of_Features()
        STRINg = str(features.columns.values[:len(features.columns.values)//no]).replace("'","")
        Fea_MESSAGE1.configure(text = f'{FEAT}\n{STRINg} window of {no}',fg='#F5CC21')
        CompareFEAtVAR.set(FEAT)
        WIN17.destroy()
        WIN17.quit()
    try:
        DynamicColumns = df.columns.to_list() 
        
    except NameError as ERR:
        showmessage('Dataframe Error',f'Dataframe is not available!\n{ERR}',TIMEOUT=3000,TYP='Erro')
        LoadDataBase('')
    else:
        try:
            DynamicColumns.remove('MARKET')
        except KeyError as ERR:
            pass
        try:
            DynamicColumns.remove('Rise_fall')
        except KeyError as ERR:
            pass
        try:
            DynamicColumns.remove('LABELS')
        except (KeyError,ValueError) as ERR:
            #showmessage('Column Error',f'Column LABELS is not in dataframe df\n{ERR}',TIMEOUT=2000,TYP='warn')
            pass
        try:
            DynamicColumns.remove('spread')
        except (KeyError,ValueError) as ERR:
            #showmessage('Column Error',f'Column spread is not in dataframe df\n{ERR}',TIMEOUT=2000,TYP='warn')
            pass
        try:
            DynamicColumns.remove('Lottage')
        except (KeyError,ValueError) as ERR:
            #showmessage('Column Error',f'Column Lottage is not in dataframe df\n{ERR}',TIMEOUT=2000,TYP='warn')
            pass
        try:
            DynamicColumns.remove('tick_volume')
        except (KeyError,ValueError) as ERR:
            #showmessage('Column Error',f'Column tick_volume is not in dataframe df\n{ERR}',TIMEOUT=2000,TYP='warn')
            pass
        
        try:
            DynamicColumns.remove('predicted')
        except (KeyError,ValueError) as ERR:
            #showmessage('Column Error',f'Column predicted is not in dataframe df\n{ERR}',TIMEOUT=2000,TYP='warn')
            pass
        
        try:
            DynamicColumns.remove('predictedclass')
        except (KeyError,ValueError) as ERR:
            #showmessage('Column Error',f'Column predictedclass is not in dataframe df\n{ERR}',TIMEOUT=2000,TYP='warn')
            pass
        
        try:
            DynamicColumns.remove('profit')
        except (KeyError,ValueError) as ERR:
            #showmessage('Column Error',f'Column profit is not in dataframe df\n{ERR}',TIMEOUT=2000,TYP='warn')
            pass
        
        try:
            DynamicColumns.remove('equity')
        except (KeyError,ValueError) as ERR:
            #showmessage('Column Error',f'Column equity is not in dataframe df\n{ERR}',TIMEOUT=2000,TYP='warn')
            pass
        
        try:
            DynamicColumns.remove('difference')
        except (KeyError,ValueError) as ERR:
            #showmessage('Column Error',f'Column difference is not in dataframe df\n{ERR}',TIMEOUT=2000,TYP='warn')
            pass
        
        try:
            DynamicColumns.remove('Compare')
        except (KeyError,ValueError) as ERR:
            #showmessage('Column Error',f'Column Compare is not in dataframe df\n{ERR}',TIMEOUT=2000,TYP='warn')
            pass
        
        try:
            DynamicColumns.remove('Win/Loss')
        except (KeyError,ValueError) as ERR:
            #showmessage('Column Error',f'Column DrawDown is not in dataframe df\n{ERR}',TIMEOUT=2000,TYP='warn')
            pass
        
        try:
            DynamicColumns.remove('DrawDown')
        except (KeyError,ValueError) as ERR:
            #showmessage('Column Error',f'Column DrawDown is not in dataframe df\n{ERR}',TIMEOUT=2000,TYP='warn')
            pass
        
        WIN17 = tk.Toplevel(root)
        WIN17.bind('<Escape>',EXIT_One)
        WIN17.bind('<Return>',selected_features)
        FeatBTN = tk.Button(WIN17, text='Select the features',font=('Times', FONTSIZE, 'bold'),fg='yellow', bg='mediumvioletred',command = lambda:selected_features(''))  
        FeatBTN.pack(fill = tk.BOTH,expand=1)
        featvar = tk.Variable(value=DynamicColumns)
        lsbox = tk.Listbox(WIN17, width=40,height=len(DynamicColumns)+1,listvariable=featvar, selectmode = tk.MULTIPLE)
        lsbox.pack(fill = tk.BOTH,expand=1)
        WIN17.mainloop()


def Make_window_features():
    global WIN10
    n_prev_samples = 0
    WINd = tk.StringVar()
    def MakeWin(_):
        nonlocal n_prev_samples 
        try:
            n_prev_samples = int(WIND_entry.get())
            if n_prev_samples<0:
                raise RuntimeError('it\'s negative!!!')
        except Exception as ERR:
            showmessage('Size of window Error',f'it\'s not valid!\n{ERR}',TIMEOUT=3000,TYP='Erro')

        WIN10.destroy()
        WIN10.quit()
        #return n_prev_samples
    WIN10 = tk.Toplevel(root)
    WIN10.bind('<Escape>',EXIT_One)
    WIN10.title('Get Window')
    WIN10.configure(bg='#091A32')
    tk.Label(WIN10,image = pho).place(height=h,width=w,x=0,y=0)#.grid(rowspan=8,columnspan=2) 
    WindowSizeChoice = tk.Label(WIN10,text="What's the size of window: ",bg='#474747',fg='yellow',font=('times',FONTSIZE,'bold'))
    WindowSizeChoice.grid(row=0,columnspan = 4,pady=5)   #,sticky=tk.W
    WINDLBL = tk.Label(WIN10,text="Size of window: ",bg='#474747',fg='white',font=('times',FONTSIZE,'bold'))
    WINDLBL.grid(row=1,column=0,pady=5) #,sticky='W'

    WIND_entry = tk.Entry(WIN10,textvariable=WINd,width=10,bg='#474747',fg='yellow',font=('times',FONTSIZE,'bold'),borderwidth=3,bd=5)
    WIND_entry.grid(row=1,column=1,pady=5)#,sticky='W'

    WINDbtn = tk.Button(WIN10,text="Add to database",bg='#474747',fg='yellow',font=('times',FONTSIZE2,'bold'),relief=tk.RAISED,bd=5,command=lambda:MakeWin(''))
    WINDbtn.grid(row=2,columnspan=2)
    WIN10.bind('<Return>',MakeWin)
    WIND_entry.focus()
    WIN10.mainloop()
    return n_prev_samples

def Window_of_Features():
    global df
    global features
    global DynamicColumns
    try:
        df
    except NameError as ERR:
        showmessage('Dataframe Error',f'Dataframe is not available\n{ERR}',TIMEOUT=3000,TYP='Erro')
        LoadDataBase('')
        
    n_prev_samples = Make_window_features()
    try:
        DynamicColumns
    except NameError as ERR:
        showmessage('DynamicColumns Error',f'DynamicColumns are not available\n{ERR}',TIMEOUT=3000,TYP='Erro')
        LoadDataBase('')
    else:
        NewColumns = DynamicColumns.copy()
        print('n_prev_samples:',n_prev_samples)
        for c in range(1,n_prev_samples+1):
            for i in DynamicColumns:
                df[i+'('+str(-c)+')']=df[i].shift(c)
                NewColumns +=[i+'('+str(-c)+')']
                print(i+'('+str(-c)+')')
        #df.dropna(axis=0,inplace=True)
        df.fillna(0,inplace=True)
        features = df[NewColumns]
        features = features.shift()
        features.fillna(0,inplace=True)#features.dropna(axis=0,inplace=True)
    return n_prev_samples

def Build_Binary_Database():
    global DataPath
    try:
        df
    except NameError as ERR:
        showmessage('Dataframe Error',f'Dataframe is not available\n{ERR}',TIMEOUT=3000,TYP='Erro')
        LoadDataBase('')
    else:
        try:
            df['LABELS']
        except:
            global WIN16
            WIN16 = tk.Toplevel(root)
            WIN16.bind('<Escape>',EXIT_One)
            Label_Select(WIN16)
            WIN16.lift()
            WIN16.focus_force()
        else:
            df2 = df.copy()
            MaxCodeLength = MAX_Num_Bits('Build binary coded database!')
            for BC in range(MaxCodeLength):
                bitcoding = str(BC+1)
                BCDn = 'BCD'+ bitcoding
                df2[BCDn] = 0
                for i in range(BC+1):
                    df2['past(-'+str(i+1)+')'] = df2.LABELS.shift(i+1)*2**(i)
                df2.fillna(0,inplace=True)
                for i in range(BC+1):
                    df2['past(-'+str(i+1)+')'].astype(np.int64)
                    df2[BCDn]+= df2['past(-'+str(i+1)+')']
                
                df2[BCDn]=df2[BCDn].astype(np.int64)
                df2['BinaryCoded'+bitcoding] = df2[BCDn].apply(lambda x:bin(x))
                df2[bitcoding+'_BCP'] = df2[BCDn].apply(lambda x:bin(x)[2:].zfill(BC))
                filename =  r'\BinaryDatabase'+bitcoding
                #print('df2.tail(10):\n',df2.tail(10))
                df2.to_excel(DataPath+filename+'.xlsx') 
                #df = pd.concat((df,df2['Bit']),axis=1)
                #print('df.tail(5):\n',df.tail(10))
                print("df2.groupby('LABELS')[bitcoding+'_BCP'].value_counts():\n",df2.groupby('LABELS')[bitcoding+'_BCP'].value_counts())
                class_Zero = df2.groupby('LABELS').get_group(0)
                print('class_Zero:\n',class_Zero)

                class_One = df2.groupby('LABELS').get_group(1)
                print('class_One:\n',class_One)

                print("df2['LABEL'].value_counts():\n",df2['LABELS'].value_counts(),sep='')
            ############################
                print("\ndf2[bitcoding+'_BCP'].value_counts():\n",df2[bitcoding+'_BCP'].value_counts())
                # dfsorted = df.sort_values(by=BCDn)
                # print('dfsorted:\n',dfsorted)
                dfsorted_class0 = df2[df2['LABELS']==0].sort_values(by=BCDn)
                counted0 = dfsorted_class0[bitcoding+'_BCP'].value_counts()
                COUNTED0 = counted0.to_frame()
                #COUNTED0 = pd.DataFrame(counted0)
                print('COUNTED0',COUNTED0)
                dfsorted_class1 = df2[df2['LABELS']==1].sort_values(by=BCDn)
                counted1 = dfsorted_class1[bitcoding+'_BCP'].value_counts()
                COUNTED1 = counted1.to_frame()
                print('COUNTED1',COUNTED1)

                dataframen = pd.concat((COUNTED0,COUNTED1),axis=1)
                dataframen.columns=['N_Zeros_Fall','N_Ones_Rise']
                SUMofRecords = dataframen.sum().sum()
                print('dataframen:',dataframen)
                print('dataframen.sum() ',dataframen.sum())
                print('dataframen.sum().sum() ',dataframen.sum().sum())
                dataframen[bitcoding+'_BCP'] = dataframen.index
                
                print(f'sum of null numbers in dataframe{bitcoding} is {dataframen.isnull().sum().sum()} that is replaced by 0')
                dataframen.fillna(0,inplace=True)
                #dataframen.sort_index(inplace=True)
                dataframen['DeltaTrend'] = dataframen['N_Ones_Rise']-dataframen['N_Zeros_Fall']
                dataframen['Mean_Delta_'+bitcoding] = dataframen['DeltaTrend']/(dataframen['N_Ones_Rise']+dataframen['N_Zeros_Fall'])
                #dataframen['TotalMean_Delta_'+bitcoding] = dataframen['DeltaTrend']/SUMofRecords

                dataframen['Sum_Return_'+bitcoding]=0
                TR = []
                for B in dataframen[bitcoding+'_BCP']:
                    TR += [sum(df2[df2[bitcoding+'_BCP']==B]['Return'])]
                    dataframen.loc[B,'Sum_Return_'+bitcoding] = sum(df2[df2[bitcoding+'_BCP']==B]['Return'])
                    dataframen.loc[B,'Mean_Return'+bitcoding] = np.mean(df2[df2[bitcoding+'_BCP']==B]['Return'])
                    
                sumreturn = pd.DataFrame(TR,columns=['Sum_Return_'+bitcoding],index=dataframen.index)
                SumofReturns = sum(df2['Return'])

                dataframen = pd.concat((dataframen,sumreturn),axis=1)
                print(f'dataframe{bitcoding}=\n',dataframen,sep='')
                #makedictionary
                #indicatorn =dataframen['TotalDelta'+bitcoding]
                filename =  r'\BinaryDatabase'+bitcoding
                dataframen.to_excel(DataPath+filename+'.xlsx')
                dataframen.to_csv(DataPath+filename+'.csv')
                a=showmessage(f'Database Files {bitcoding} Saved',f'Database Files {bitcoding} Saved!',TIMEOUT=2000,TYP='info')
            showmessage(f'Database Built completely',f'Database is Built completely with {bitcoding} coding file!',TIMEOUT=20000,TYP='info')
A=''
tempdf = ''
def Add_Binary_indicators(temp1):
    global A
    global df
    global features
    global tempdf
    try:
        df
    except NameError as ERR:
        showmessage('Dataframe not found',f'Dataframe is not available!\n{ERR}',TIMEOUT=3000,TYP='Erro')
        LoadDataBase('')
    else:
        try:
            df['LABELS']
        except:
            global WIN16
            WIN16 = tk.Toplevel(root)
            WIN16.bind('<Escape>',EXIT_One)
            Label_Select(WIN16)
            WIN16.lift()
            WIN16.focus_force()
        else:
            #df2 = df.copy()            
            MaxCodeLength_ = MAX_Num_Bits(temp1)
            BitcodingtXT = '';BitcodingtXT1 = ''
            for BC in range(MaxCodeLength_):
                bitcoding = str(BC+1)
                filename =  r'\BinaryDatabase'+bitcoding
                try:
                    tempdf = pd.read_excel(DataPath+filename+'.xlsx')
                except FileNotFoundError as ERR:
                    showmessage('Path Error',f'Path is not correct or available!\n{ERR}',TIMEOUT=3000,TYP='Erro')
                    return 0
                tempdf.set_index(bitcoding+'_BCP', inplace=True)
                A = tempdf['Mean_Return'+bitcoding].to_dict()
                BitcodingtXT += f'A for bitcoding {bitcoding} is\n{A}\n\n'
                #Show_In_Win(f'bitcoding {bitcoding}',f'A for bitcoding {bitcoding} is\n{A}')      
                B = tempdf['Mean_Delta_'+bitcoding].to_dict()
                BitcodingtXT1 += f'B for bitcoding {bitcoding} is\n{B}\n\n'
                BCDn = 'BCD'+ bitcoding
                df[BCDn] = 0
                for i in range(BC+1):
                    df['past(-'+str(i+1)+')']= df.LABELS.shift(i)*2**(i)

                #df.dropna(inplace=True)
                df.fillna(0,inplace=True)
                for i in range(BC+1):
#                     df['past(-'+str(i+1)+')'].astype(np.int32)
                    df[BCDn] += df['past(-'+str(i+1)+')']          
                df[BCDn] = df[BCDn].astype(np.int64)
                df[bitcoding+'_BCP'] = df[BCDn].apply(lambda x:bin(x)[2:].zfill(BC))
                df['Bin_indicator_'+bitcoding] = 0
#                 try:
                df['MeanDelta'+bitcoding] = 0
                try:
                    df['MeanDelta'+bitcoding] = df[bitcoding+'_BCP'].astype(np.int64).apply(lambda x: B[x])
                except:
                    pass

                try:
                    df['Bin_indicator_'+bitcoding] = df[bitcoding+'_BCP'].astype(np.int64).apply(lambda x: A[x])
                except:
                    pass
                
            for i in range(MaxCodeLength_):
                df.drop('past(-'+str(i+1)+')',axis=1,inplace=True)        
            Show_In_Win('Dataframe df',f'df with Binary features added:\n{df}','black','#fefcf1')
            columns = df.columns.to_list()
#             M = []
#             for BC in range(MaxCodeLength_):
#                 bitcoding = str(BC+1)
#                 M += ['BCD'+bitcoding,bitcoding+'_BCP']
                         
            Show_In_Win(f'bitcoding A',BitcodingtXT,'#ffffff','blue4')
            Show_In_Win(f'bitcoding B',BitcodingtXT1,'#ffffff','green4')
            if temp1 != 'Lottage':
                All_Features(False)
            for i in range(1,MaxCodeLength_+1):
                features['BCD'+ str(i)] = features['BCD'+ str(i)].astype(np.int64)
            print('this is MaxCodeLength_:',MaxCodeLength_)
            return MaxCodeLength_

def Just_Binary_indicators():
    global df
    global features
    global Scaling_Type
    MaxCodeLength_ = Add_Binary_indicators('put just binary indicators')
    print('MaxCodeLength_:',MaxCodeLength_)
    if MaxCodeLength_ == None:
        showmessage('Try again',f'Run Just_Binary_indicators() again!',TIMEOUT=2500,TYP='warn')
        return
    M = []
    for BC in range(MaxCodeLength_):
        bitcoding = str(BC+1)
        M += ['BCD'+bitcoding]
    features = features[M]
    print(f'Features with Just Binary indicators:\n{features}')
    Show_In_Win('Features',f'Features with Just Binary indicators:\n{features}','White','Blue')

def Selected_Binary_indicators():
    global df
    global features

    Add_Binary_indicators('Select some Binary indicators later!')
    seletded_Bits = Misc_Num_Bits()
    M = []
    for BC in seletded_Bits:
        bitcoding = str(BC)
        M += ['BCD'+bitcoding]
    features = features[M]
    
    Featurestxt = f'Features of selected Binary indicators:\n{features}'
    print(Featurestxt)
    Show_In_Win('Features',Featurestxt,'White','Blue')

def Just_Mean_Return_indicators():
    global df
    global features
    MaxCodeLength_ = Add_Binary_indicators('Just Mean Return Indicators')
    if MaxCodeLength_ == None:
        showmessage('Try again',f'Run "Just_Mean_Return_indicators()" again!',TIMEOUT=2500,TYP='warn')
        return
    M = []
    for BC in range(1,MaxCodeLength_+1):
        bitcoding = str(BC)
        M += ['Bin_indicator_'+bitcoding]
    features = features[M]
            
    print(f'Features with Just Binary indicators:\n{features}')
    Show_In_Win('Features',f'Features with Just Mean Return Indicators:\n{features}','White','Blue')

def Just_Mean_Delta():
    global df
    global features
    global Scaling_Type
    MaxCodeLength_ = Add_Binary_indicators('put Just mean Delta indicator')
    if MaxCodeLength_ == None:
        showmessage('Try again',f'Run Just_Mean_Delta() again!',TIMEOUT=2500,TYP='warn')
        return
    M = []
    for BC in range(MaxCodeLength_):
        bitcoding = str(BC+1)
        M += ['MeanDelta'+bitcoding]
    features = features[M]  
    print(f'Features with Just mean Delta:\n{features}')
    Show_In_Win('Features',f'Features with Just mean Delta:\n{features}','White','purple')

def Set_Label(L):
    global df
    global labels
    try:
        df
    except NameError as ERR:
        showmessage('Dataframe not found',f'Dataframe df is not available!\n{ERR}',TIMEOUT=3000,TYP='Erro')
        LoadDataBase('')
    
    try:
        df['LABELS'] = df[L]
        labels = df['LABELS']#.values
        labletext = f"Label is {L} and equal to:\n{df['LABELS']}"#.to_string()
        Show_In_Win('lable',labletext,'black','gold',2500)
        
    except NameError as ERR:
        showmessage('Label not available',f'Label {L} is not available!\n{ERR}',TIMEOUT=3000,TYP='Erro')
        Set_Label(L)
        
    except KeyError as ERR:
        showmessage('Label not available',f'column is not available:\n{ERR}',TIMEOUT=3000,TYP='Erro')

        Add_ScaledReturn()
        Set_Label(L)
    
def Show_In_Win(Title,Sample,BG='black',FG='orange',TIMEOUT=0):
    global WIN13
    if TIMEOUT:
        win = tk.Tk()
        win.withdraw()
        win.after(TIMEOUT,win.destroy)
        win.option_add('*Dialog.msg.font', 'Helvetica 20')
        WIN13 = tk.Toplevel(win)
    else:
        WIN13 = tk.Toplevel(root)
        WIN13.bind('<Escape>',EXIT_One)
    WIN13.title(Title)
    #WIN13.geometry('%dx%d'%(ScreenWidth//3,ScreenHeight))
    WIN13.resizable(height=False,width=False)
    text1 = tk.Text(WIN13,bg=BG,fg=FG,font=('times',int(FONTSIZE*1.2),'bold'))#,wrap='word')
    text1.grid(row=0,column=0)
    text1.insert(tk.END,Sample)
    scroll1 = ttk.Scrollbar(WIN13,orient=tk.VERTICAL,command=text1.yview)
    scroll1.grid(row=0,column=1,sticky='ns')
    text1.config(yscrollcommand=scroll1.set)
    scroll2 = ttk.Scrollbar(WIN13,orient=tk.HORIZONTAL,command=text1.xview)
    scroll2.grid(row=1,column=0,sticky='we')
    text1.config(xscrollcommand=scroll2.set) 

        
    
def Select_Scaling_Type():
    global AlreadyScaled
    global df
    global WIN11
    global WIN15
    global features
    global Prepared_features
    global Scaling_Type
    global QuickRun
    global S1
    Scaling2 = 'Not Scaled features'
    def RDchoice1(p):
        nonlocal a , Scaling2
        Scaling2 = a.get()
        if p:
            WIN11.destroy()
            WIN11.quit()
            return Scaling2
    def navigate_RadioButtons(event):
        # Get the current selection index  
        current_index = Options.index(a.get())   
        if event.keysym == 'Down':
            # Move down in the list  
            current_index = (current_index + 1) % 4
        elif event.keysym == 'Up':
            # Move up in the list
            current_index = (current_index - 1) % 4
        # Update the selected option
        a.set(Options[current_index])
    try:
        df
    except NameError as ERR:
        showmessage('Dataframe Error',f'Dataframe is not available!\n{ERR}',TIMEOUT=3000,TYP='Erro')
        LoadDataBase('')
    try:
        features  
    except NameError as ERR:
        showmessage('Features Error',f'No feature is available!\n{ERR}',TIMEOUT=3000,TYP='Erro')
        WIN15 = tk.Toplevel(root)
        WIN15.bind('<Escape>',EXIT_One)
        WIN15.attributes('-topmost','true')
        feature_Select(WIN15,4) 
    else:
        if not QuickRun:
            WIN11 = tk.Toplevel(root)
            WIN11.bind('<Escape>',EXIT_One)
            WIN11.bind('<Down>', navigate_RadioButtons)  
            WIN11.bind('<Up>', navigate_RadioButtons)
            WIN11.bind('<Return>',RDchoice1)
            a = tk.StringVar(WIN11,'none')
            Options = ['MinMax Scaler','Standard Scaler','Robust Scaler','without scaling'] 
            RadiBut1 = tk.Radiobutton(WIN11,text='Use '+Options[0],indicator = 0,background = "coral1",justify=tk.LEFT,variable=a,value = Options[0],command=lambda:RDchoice1(0))
            RadiBut1.pack(fill = tk.X, ipady = 5)#.pack(anchor=tk.W)
            
            RadiBut2 = tk.Radiobutton(WIN11,text='Use '+Options[1],indicator = 0,background = "coral2",justify=tk.LEFT,variable=a,value = Options[1],command=lambda:RDchoice1(0))
            RadiBut2.pack(fill = tk.X, ipady = 5)##.place(height=20,width=200,x=0,y=20)
            
            RadiBut3 = tk.Radiobutton(WIN11,text='Use '+Options[2],indicator = 0,background = "coral3",justify=tk.LEFT,variable=a,value = Options[2],command=lambda:RDchoice1(0))
            RadiBut3.pack(fill = tk.X, ipady = 5)##.place(height=20,width=200,x=0,y=20)

            RadiBut4 = tk.Radiobutton(WIN11,text='Use '+Options[3],indicator = 0,background = "coral4",justify=tk.LEFT,variable=a,value = Options[3],command=lambda:RDchoice1(0))
            RadiBut4.pack(fill = tk.X, ipady = 5)##.place(height=20,width=200,x=0,y=20)

            scalBTN = tk.Button(WIN11,text="↑Select Scaling Method↑",bg='yellow',relief="raised",command=lambda:RDchoice1(1))
            scalBTN.pack(fill = tk.X, ipady = 5)#.pack(anchor=tk.CENTER)
            root.bind('<Return>', RDchoice1)
            scaldict={0:'MinMax Scaler',1:'Standard Scaler',2:'Robust Scaler',3:'without scaling'}
            a.set(scaldict[Scaling_Type])
            scalBTN.focus()
            WIN11.mainloop()
        else:
            Scaling2 = 'MinMax Scaler'        

        if Scaling2 == 'MinMax Scaler':
            scale = MinMaxScaler()    # MinMaxScaler(feature_range = (0, 1))
            Prepared_features = scale.fit_transform(features)
            S1 = 'Normalized features(MinMax)'
            
        elif Scaling2 == 'Standard Scaler':
            scale = StandardScaler()
            Prepared_features = scale.fit_transform(features)
            S1= 'Standardized features'
            
        elif Scaling2 == 'Robust Scaler':
            scale = RobustScaler()
            Prepared_features = scale.fit_transform(features)
            S1= 'Robust scaled features'

        else:        
            Prepared_features = features.values
            S1 = 'without scaling'

        statisticalreview =f'''max of features     :{Prepared_features.max()}
        min of features     : {Prepared_features.min()}
        mean of features    : {Prepared_features.mean()}
        Std dev of features: {Prepared_features.flatten()}
        Features describe   :\n{pd.DataFrame(Prepared_features).describe()}''' 
        if not Concise_Report:
            Show_In_Win('Statistical Review',statisticalreview,'white','black')

def Auto_PrePro():
    global df
    global features
    global Scaling_Type
    Modify_Date()
    Select_Scaling_Type()

def Add_Return():
    from sklearn.impute import SimpleImputer
    global df
    try:
        df['CLOSE']
    except NameError as ERR:
        showmessage('Dataframe Error',f'Dataframe is not available\n{ERR}',TIMEOUT=3000,TYP='Erro')

        LoadDataBase('')
    df['Return'] = df['CLOSE'] - df['CLOSE'].shift()
    raw_returns = df['Return'].values
    #============================== Data Imputation
    imp = SimpleImputer(missing_values=np.nan , strategy='mean')
    raw_returns=raw_returns.reshape(-1,1)
    imp.fit(raw_returns[:5])
    raw_returns = imp.transform(raw_returns)
    df['Return'] = pd.DataFrame(raw_returns , index=df.index)
    #============================== Data Cleansing
    df['Rise_fall'] = df['Return'].apply(lambda x: 1 if x>0.0 else 0).astype(np.int8)

def Add_ScaledReturn():
    global df
    Add_Return()
    STD = stdev(list(df['Return'].values))
    conditions = [(df['Return']>=3*STD),(df['Return']<=-3*STD),True]#df['Return'].abs()<STD]
    state = [3*STD,-3*STD,df['Return']]
    df['Return']= np.select(conditions,state)
    Cleansed_returns = df['Return'].values.reshape(-1,1)
    #============================== Data Normalizing
    # scaler1 = MinMaxScaler(feature_range=(-0.2,0.2))
    scaler1 = StandardScaler(with_mean=False)
    scaler1 = scaler1.fit(Cleansed_returns)  # scaler1.fit(raw_returns)
    Cleansed_returns = scaler1.transform(Cleansed_returns)
    df['ScaledReturn'] = pd.DataFrame(Cleansed_returns , index=df.index)

def Add_previous(M):
    global WIN10
    n_prev_samples = 0
    Clo  = tk.StringVar()
    def Addprev(_):
        nonlocal n_prev_samples
        try:
            n_prev_samples = int(PREVe_entry.get())
            if n_prev_samples<0:
                raise RuntimeError('it\'s negative!!!')       
        except Exception as ERR:
            showmessage(f'No. of {M} Error',f'it\'s not valid!\n{ERR}',TIMEOUT=3000,TYP='Erro')
        else:
            try:
                for c in range(1,n_prev_samples+1):
                    df[M+'('+str(-c)+')']=df[M].shift(c)
                #df.dropna(axis=0,inplace=True)
                df.fillna(0,inplace=True)
            except NameError as ERR:
                showmessage('df Error',ERR,TIMEOUT=3000,TYP='Erro')
            else:
                #messagebox.showinfo(title='No. of closes', message = f'{n_prev_samples} of previous  closes are added to df database!' )
                showmessage(f'No. of {M}s',f'{n_prev_samples} of previous {M} are added to df database!',TIMEOUT=2000,TYP='info')
        WIN10.destroy()
        WIN10.quit()
        return n_prev_samples
    
    WIN10 = tk.Toplevel(root)
    WIN10.bind('<Escape>',EXIT_One)
    WIN10.title(M)
    WIN10.configure(bg='#091A32')

    tk.Label(WIN10,image = pho).place(height=h,width=w,x=0,y=0)#.grid(rowspan=8,columnspan=2)

    #showmessage('test1',f'you are here1',TIMEOUT=2500,TYP='warn')
    PREVeChoice = tk.Label(WIN10,text=f"How many previous {M} records regarded as input features: ",bg='#474747',fg='yellow',font=('times',FONTSIZE,'bold'))
    PREVeChoice.grid(row=0,columnspan = 4,pady=5)   #,sticky=tk.W 
    
    PREVLBL = tk.Label(WIN10,text=f"No. of Previous {M}: ",bg='#474747',fg='white',font=('times',FONTSIZE,'bold'))
    PREVLBL.grid(row=1,column=0,pady=5) #,sticky='W'

    PREVe_entry = tk.Entry(WIN10,textvariable=Clo,width=10,bg='#474747',fg='yellow',font=('times',FONTSIZE,'bold'),borderwidth=3,bd=5)
    PREVe_entry.grid(row=1,column=1,pady=5)#,sticky='W'

    btPREV = tk.Button(WIN10,text="Add to df database",bg='#474747',fg='yellow',font=('times',FONTSIZE2,'bold'),relief=tk.RAISED,bd=5,command=lambda:Addprev(''))
    WIN10.bind('<Return>',Addprev)

    btPREV.grid(row=2,columnspan=2)
    PREVe_entry.focus()
    WIN10.mainloop()
    return n_prev_samples

def MAX_Num_Bits(M):
    global WIN10
    MaxCodeLength = 0
    Clo = tk.StringVar()
    def AddBit(_):
        nonlocal MaxCodeLength
        global LABL
        try:
            MaxCodeLength = int(Bit_entry.get())
            if MaxCodeLength<0:
                raise RuntimeError('it\'s negative!!!')
        except Exception as ERR:
            showmessage('No. of bits Error',f'theres\'s an invalid number for bit coding!\n{ERR}',TIMEOUT=3000,TYP='Erro')
        else:
            MaxCodeMessage = f'{MaxCodeLength} of previous {LABL}s are considered! to {M} '   
            showmessage('No. of bits',MaxCodeMessage,TIMEOUT=3000,TYP='info')
        WIN10.destroy()
        WIN10.quit()
        return MaxCodeLength
    
    WIN10 = tk.Toplevel(root)
    WIN10.bind('<Escape>',EXIT_One)
    WIN10.title('Bits')
    WIN10.configure(bg=('#A01A32' if M else '#071A32'))

    BitChoice = tk.Label(WIN10,text=f'How many previous {LABL} records regarded to {M}',bg='#474747',fg='yellow',font=('times',FONTSIZE,'bold'))
    BitChoice.grid(row=0,columnspan = 4,pady=5)   #,sticky=tk.W 
    
    BitLBL = tk.Label(WIN10,text=f"No. of Previous {LABL}s: ",bg='#474747',fg='white',font=('times',FONTSIZE2,'bold'))
    BitLBL.grid(row=1,column=0,pady=5) #,sticky='W'

    Bit_entry = tk.Entry(WIN10,textvariable=Clo,width=15,bg='#474747',fg='yellow',font=('times',FONTSIZE2,'bold'),borderwidth=3,bd=5)
    Bit_entry.grid(row=1,column=1,pady=5)#,sticky='W'

    BitButton = tk.Button(WIN10,text="Consider for coding "+("as binary indicator?" if M else "as binary database?"),bg='#474747',fg='yellow',font=('times',FONTSIZE2,'bold'),relief=tk.RAISED,bd=5,command=lambda:AddBit(''))
    WIN10.bind('<Return>',AddBit)

    BitButton.grid(row=2,columnspan=2)
    Bit_entry.focus()
    WIN10.mainloop()
    return MaxCodeLength

def Misc_Num_Bits():
    global WIN10
    MaxCodeLength = 0
    Clo = tk.StringVar()
    seletded_Bits = 0
    
    def AddBit(_):
        nonlocal seletded_Bits
        global LABL
        try:
            MiscCodes = Bit_entry.get()
            seletded_Bits = list(map(int,MiscCodes.split('-')))
            for b in seletded_Bits:
                if b<0:
                    raise RuntimeError('theres\'s a negative number!!!')
        
        except Exception as ERR:
            showmessage('No. of bits Error',f'theres\'s an invalid number for bit coding!\n{ERR}',TIMEOUT=3000,TYP='Erro')
        else:
            BinCodeMessage = f'{MiscCodes} of previous {LABL}s are considered to use as features.'   
            showmessage('No. of bits',BinCodeMessage,TIMEOUT=2000,TYP='info')
        WIN10.destroy()
        WIN10.quit()
#        return seletded_Bits
    
    WIN10 = tk.Toplevel(root)
    WIN10.bind('<Escape>',EXIT_One)
    WIN10.title('Bits for coding')
    WIN10.configure(bg='#A01A32')

    tk.Label(WIN10,image = pho).place(height=h,width=w,x=0,y=0)#.grid(rowspan=8,columnspan=2)
    #showmessage('test3',f'you are here3',TIMEOUT=2500,TYP='warn')
    BitChoice = tk.Label(WIN10,text=f"How many previous {LABL} records regarded to use as binary indicator?" ,bg='#474747',fg='light green',font=('times',FONTSIZE,'bold'))
    BitChoice.grid(row=0,columnspan = 4,pady=5)   #,sticky=tk.W 
    
    BitLBL = tk.Label(WIN10,text=f"Select Previous {LABL}s seperated with -: ",bg='#474747',fg='white',font=('times',FONTSIZE2,'bold'))
    BitLBL.grid(row=1,column=0,pady=5)

    Bit_entry = tk.Entry(WIN10,textvariable=Clo,width=15,bg='#474747',fg='yellow',font=('times',FONTSIZE2,'bold'),borderwidth=3,bd=5)
    Bit_entry.grid(row=1,column=1,pady=5)

    BitButton = tk.Button(WIN10,text="Consider for coding as binary indicator" ,bg='#474747',fg='yellow',font=('times',FONTSIZE2,'bold'),relief=tk.RAISED,bd=5,command=lambda:AddBit(''))
    WIN10.bind('<Return>',AddBit)

    BitButton.grid(row=2,columnspan=2)
    Bit_entry.focus()
    WIN10.mainloop()
    return seletded_Bits

def SaveFigure(FigName,folder,MOD):
    global FigureName
    FigureName = ''
    if len(Strategy.split())>2:
        Short_Strategy = ''.join(x[0] for x in Strategy.split())
    else:
        Short_Strategy = Strategy
        
    Figures_path = os.getcwd()+f'\\Simulation Results\\{folder}'       
    if not os.path.exists(Figures_path):
        showmessage('Finding path',f'The path {Figures_path} was not found!\nIt will be created!',TIMEOUT= 3000,TYP='warn')
        try:
            os.makedirs(Figures_path)
        except Exception as ERR:
            showmessage('Path Error',f'Check the file path or name: {Figures_path}\n{ERR}',TIMEOUT=3000,TYP='Erro')
            Figures_path = Figures_path.replace(' ','_').replace("'",'').replace(')','').replace('(','').replace(']','▬').replace('[','▬')
            Figures_path = filedialog.asksaveasfilename(initialdir = Figures_path ,defaultextension=".txt", filetypes=[("text file", "*.txt"),  ("All files", "*.*")],initialfile= Figures_path[:10])
            if not os.path.exists(Figures_path):
                showmessage('Finding path',f'The path {Figures_path} was not found!\nIt will be created!',TIMEOUT= 3000,TYP='warn')
                os.makedirs(Figures_path)

    FigureName = Short_Strategy+str(features.columns.values).replace(' ','_').replace("'",'').replace(')','').replace('(','').replace(']','▬').replace('[','▬')
    if ComparisonFlag or FEAT =='Just previous Closes' or FEAT == 'Just previous Labels' or FEAT == 'Custom Window Features':
        if len(FigureName.split('_'))>2:
            FigureName = FigureName.split('_')[0]+'→'+FigureName.split('_')[-1]
    #Figures_path = FigName + Figures_path
#     FigureName += f'{"NovelSplit" if SingleCandlePredict else "NormalSplit"}♦Lable_{LABL}♦TF_{TimeFrame}♦TestSize_{TestSize}♦Trainsize_{Trainsize}'
    #STRING = '' if MOD == 'models' else Short_Strategy
        elif ComparisonFlag:
            FigureName = f'{selecteditems[0]}→{selecteditems[-1]}'
#         if MOD == 'features':
#             FigureName += f'{FigName.split("_")[0]} Comparison with multiple features♦TF_{TimeFrame}♦TestSize_{TestSize}♦Trainsize_{Trainsize}♦Lbl_{LABL}♦{"NovelSplit" if SingleCandlePredict else "NormalSplit"}'
#         elif MOD == 'SMAs':
#             FigureName += f'{FigName.split("_")[0]} Comparison with multiple SMAs♦TF_{TimeFrame}♦TestSize_{TestSize}♦Trainsize_{Trainsize}♦Lbl_{LABL}♦{"NovelSplit" if SingleCandlePredict else "NormalSplit"}'      
#         else:

        FigureName = f'{FigName.split("_")[0]} Comparison with multiple {MOD}♦TF_{TimeFrame}♦TestSize_{TestSize}♦Trainsize_{Trainsize}♦Feature_{FEAT}_{FigureName}Lable_{LABL}♦{"NovelSplit" if SingleCandlePredict else "NormalSplit"}'
        
        if MOD != 'models':
            FigureName += f'♦Model_{Short_Strategy}'
        
    else:
        try:
            df["equity"]            
        except KeyError:
            pass
        else:
            FigureName =FigName + FigureName + f'_Profit→{int(df.tail(1)["equity"].values[0]- BALANCE)}$'

    Figures_path = filedialog.asksaveasfilename(initialdir = Figures_path ,defaultextension=".jpeg", filetypes=[("jpeg", "*.jpeg"), ("PNG", "*.png"), ("Bitmap", "*.bmp"), ("GIF", "*.gif"), ("All files", "*.*")],initialfile = FigureName)
    if Figures_path:
#         print('first:',f'{Figures_path')
#         print('Figures_path:',Figures_path)
        try:
            plt.savefig(Figures_path, bbox_inches='tight')#plt.savefig(f'{Figures_path}\\{FigName[:-4]}_{FigureName}')
        except FileNotFoundError as ERR:
            showmessage('figure path Error',f"Check the file path or name: {Figures_path}\\{FigName[:-4]}_{FigureName}\n{ERR}",TIMEOUT=2000,TYP='Erro')
            print(f'Check the figure path or name: {Figures_path}\\{FigName[:-4]}_{FigureName}\n{ERR}')
            response = messagebox.askretrycancel("Save error",'Retry Saving?')
            if  response:
                VoiceAnnounce('Try again')
                SaveFigure(FigName,folder,MOD)
                showmessage(f'SaveFigure',f'figure {FigureName} saved @ {Figures_path}\\{FigName[:-4]}_{FigureName}',TIMEOUT=2000,TYP='info')
            else:
                VoiceAnnounce('Cancel')
        else:
            print(f'Figure {FigName} is saved!\n@ {Figures_path}\n',50*'♦')
            VoiceAnnounce(f'Figure saved!')
            #os.startfile(Figures_path)
    else:
        showmessage("Figure not Saved!", f'Comparison figure {FigName} wasn\'t Saved!',TIMEOUT=2000,TYP='warn')
        

def Graph_CLOS(showing):
    global df
    try:
        df
    except NameError as ERR:
        showmessage('df error','Dataframe is not available',TIMEOUT=3000,TYP='Erro')
        LoadDataBase('')   
    else:     
        fig1 = plt.figure(figsize=(ScreenWidth,ScreenHeight))
        fig1.suptitle(f'All from\n{df.head(1).index[0]} to {df.tail(1).index[0]}')
        #plt.style.use('fivethirtyeight')
        fig1.subplots_adjust(hspace=0.38,wspace=0.15)
        
        ax1 = fig1.add_subplot(221)
        plt.xticks(rotation=45, horizontalalignment="center",fontsize=10)
        ax1.set_ylabel('Scaled Return')
        try:
            df['ScaledReturn']#.plot(title='Scaled Return', fontsize=5, ax=ax1,legend=True)
        except :
            Add_ScaledReturn()
            df['ScaledReturn'].plot(title='Scaled Return', fontsize=5, ax=ax1,legend=True)

        ax2 = fig1.add_subplot(222)
        ax2.set_title('Price chart')
        plt.xticks(rotation=45, horizontalalignment="center",fontsize=10)
        try:
            df.plot(y=['CLOSE','CLOSE(-1)'] , ylabel='two last Close Prices' , fontsize=5,ax=ax2,legend=True)
        except:
            df.plot(y=['CLOSE'] , ylabel='Close Price' , fontsize=5,ax=ax2,legend=True)
        print('Close:\n',df['CLOSE'])
        ax3 = fig1.add_subplot(223)
        #ax3.set_title('1ST order difference')
        plt.xticks(rotation=45, horizontalalignment="center",fontsize=10)
        ax3.set_ylabel('1ST order Close difference')
        ax3.set_title('1ST order difference of Closes')
        ax3.plot(df['CLOSE'].diff().dropna().values)

        ax4 = fig1.add_subplot(224)
        plt.xticks(rotation=45, horizontalalignment="center",fontsize=10)
        ax4.set_title('Natural Log of Close')
        np.log(df['CLOSE']).plot(ylabel='Ln(Close)' , fontsize=5,ax=ax4)
        print('log:\n',np.log(df['CLOSE']))
        
        SaveFigure('CloseGraph','PreRun','PreRun') #plt.savefig(Figures_path+'\CloseGraph.jpg')
        if showing:
            plt.show()
            
            
def Graph_CLOS(showing):
    global df
    try:
        df
    except NameError as ERR:
        showmessage('df error','Dataframe is not available',TIMEOUT=3000,TYP='Erro')
        LoadDataBase('')   
    else:     
        fig1 = plt.figure(figsize=(ScreenWidth,ScreenHeight))
        fig1.suptitle(f'All from\n{df.head(1).index[0]} to {df.tail(1).index[0]}')
        #plt.style.use('fivethirtyeight')
        fig1.subplots_adjust(hspace=0.38,wspace=0.15)
        
        ax1 = fig1.add_subplot(221)
        plt.xticks(rotation=45, horizontalalignment="center",fontsize=10)
        ax1.set_ylabel('Scaled Return')
        try:
            df['ScaledReturn']#.plot(title='Scaled Return', fontsize=5, ax=ax1,legend=True)
        except :
            Add_ScaledReturn()
            df['ScaledReturn'].plot(title='Scaled Return', fontsize=5, ax=ax1,legend=True)

        ax2 = fig1.add_subplot(222)
        ax2.set_title('Price chart')
        plt.xticks(rotation=45, horizontalalignment="center",fontsize=10)
        try:
            df.plot(y=['CLOSE','CLOSE(-1)'] , ylabel='two last Close Prices' , fontsize=5,ax=ax2,legend=True)
        except:
            df.plot(y=['CLOSE'] , ylabel='Close Price' , fontsize=5,ax=ax2,legend=True)
        print('Close:\n',df['CLOSE'])
        ax3 = fig1.add_subplot(223)
        #ax3.set_title('1ST order difference')
        plt.xticks(rotation=45, horizontalalignment="center",fontsize=10)
        ax3.set_ylabel('1ST order Close difference')
        ax3.set_title('1ST order difference of Closes')
        ax3.plot(df['CLOSE'].diff().dropna().values)

        ax4 = fig1.add_subplot(224)
        plt.xticks(rotation=45, horizontalalignment="center",fontsize=10)
        ax4.set_title('Natural Log of Close')
        np.log(df['CLOSE']).plot(ylabel='Ln(Close)' , fontsize=5,ax=ax4)
        print('log:\n',np.log(df['CLOSE']))
        
        SaveFigure('CloseGraph','PreRun','PreRun') #plt.savefig(Figures_path+'\CloseGraph.jpg')
        if showing:
            plt.show()

            

def LogPrice_Return_Hist(showing):
    from statsmodels.tsa.stattools import acf, pacf
    from statsmodels.graphics.tsaplots import plot_acf
    global df
    global TestSize
    try:
        df
    except NameError as ERR:
        showmessage('df error','Dataframe is not available',TIMEOUT=3000,TYP='Erro')
        LoadDataBase('')
        Graph_Autocorr(showing)   
    else: 
        ACF = acf(df['CLOSE'][-TestSize*10:])#[-100:])
        ACF_df = pd.DataFrame(ACF,columns=['Auto-Correlation of Closes'])
        PACF = pacf(df['CLOSE'][-TestSize*10:])
        PACF_df = pd.DataFrame(PACF)
        PACF_df.columns=['Partial Autocorrelation Function']
        PACFHigh = np.count_nonzero((PACF>0.85)==True)
        
        price_diff = (df['CLOSE'] - df['CLOSE'].shift())[-TestSize*10:]#.dropna()  maybe necessary
        price_diff.name='1ST Difference of Closes'
        
        ACF_diff = acf(price_diff)
        ACF_diff_df = pd.DataFrame(ACF_diff,columns=['1ST order Diff Autocorr.'])
        
        PACF_diff = pacf(price_diff)
        PACF_diff_df = pd.DataFrame(PACF_diff,columns=['1ST order Diff Partial Autocorr.'])
        df['CLOSE(-1)']=df['CLOSE'].shift()       
    
        fig2 = plt.figure(figsize=(ScreenWidth,ScreenHeight))
        #fig2.set_label='Price Closes AutoCorrelation & Partial AutoCorrelation'
        ax_1 = fig2.add_subplot(221)
        price_diff.plot(ax=ax_1,ylabel='1ST order difference',legend=True)
        #df['CLOSE'].diff().dropna().plot(ax=ax1,ylabel='1ST order difference',legend=True)
        plt.xticks(rotation=25, horizontalalignment="center",fontsize=10)
        ax_2 = fig2.add_subplot(222)
        ACF_df.plot(kind='bar',ax=ax_2,ylabel='Close \nAutocorrelation',legend=True)
        ax_3 = fig2.add_subplot(223)
        #ax_3.set_ylabel('Closes 1ST Order \n Difference Autocorrelation')
        plot_acf((df['CLOSE'].diff())[-TestSize*10:].dropna(),ax=ax_3,title=None)
        ACF_diff_df.plot(ax=ax_3,ylabel='Close 1ST Order Difference\nAutocorrelation',legend=True)
        plt.xticks(list(range(len(ACF_diff_df))),rotation=45,horizontalalignment="center",fontsize=10)
        ax_4 = fig2.add_subplot(224)
        PACF_df.plot(ax=ax_4,ylabel='Closes Partial Autocorrelation',legend=True)
        PACF_diff_df.plot(ax=ax_4,ylabel='Closes 1ST Difference \nPartial Autocorrelation',legend=True)
        plt.xticks(list(range(len(ACF_diff_df))),rotation=45, horizontalalignment="center",fontsize=10)
            
        SaveFigure('ACF_PACF','PreRun','PreRun') #plt.savefig(Figures_path+'\ACF_PACF.jpg')
        if showing:
            plt.show()

def Graph_Feat_Corr(showing):
    global df
    try:
        df
    except NameError as ERR:
        showmessage('df error','Dataframe is not available!',TIMEOUT=3000,TYP='Erro')
        LoadDataBase('')   
    if features.empty:
#     except NameError as ERR:
        showmessage('Features Error',f'No feature is available to display!\n Empty',TIMEOUT=4000,TYP='Erro')
        WIN15 = tk.Toplevel(root)
        WIN15.bind('<Escape>',EXIT_One)
        WIN15.attributes('-topmost','true')
        FstrCateg = ''
        feature_Select(WIN15,4)
        WIN15.mainloop()
    else:        
        heatmap(features.corr(),annot=True,cmap='RdYlGn',linewidths=0.2) #data.corr()-->correlation matrix        
        fig=plt.gcf()
        fig.set_size_inches(18,12)
        plt.title(f'Correlation among features / target',fontsize=10,fontname='Times new roman')
        SaveFigure('Feature_Correlation_Graph','PreRun','PreRun') #plt.savefig(Figures_path+'\Feature_Correlation_Graph.jpg')
        if showing:
            plt.show()
#         plt.pause(3)
#         plt.close(fig)

def Selectable_Graph():
    global df
    global TestSize
    try:
        df
    except NameError as ERR:
        showmessage('df error','Dataframe is not available',TIMEOUT=3000,TYP='Erro')
        LoadDataBase('')
        Selectable_Graph()
    else:      
        WIN11 = tk.Toplevel(root , bd = 5)#,cursor=)
        WIN11.bind('<Escape>',EXIT_One)
        BT1 = tk.Button(WIN11,text="֎Plot Close price֎",bg='#474747',fg='yellow',font=('times',FONTSIZE2,'bold'),relief=tk.RAISED,bd=5,command=lambda:CLOSE_Graph(True))
        BT1.pack(fill = tk.BOTH,expand=1)
        
        BT2 = tk.Button(WIN11,text="֎Plot Log of Close price֎",bg='#474747',fg='yellow',font=('times',FONTSIZE2,'bold'),relief=tk.RAISED,bd=5,command=lambda:CLOSE_LOG_Graph(True))
        BT2.pack(fill = tk.BOTH,expand=1)
        
        BT3 = tk.Button(WIN11,text="֎Plot Scaled Return֎",bg='#474747',fg='yellow',font=('times',FONTSIZE2,'bold'),relief=tk.RAISED,bd=5,command=lambda:ScaledReturn_Graph(True))
        BT3.pack(fill = tk.BOTH,expand=1)
        
        BT4 = tk.Button(WIN11,text="֎Plot Autocorrelation֎",bg='#474747',fg='yellow',font=('times',FONTSIZE2,'bold'),relief=tk.RAISED,bd=5,command=lambda:autocorr_plot(True))
        BT4.pack(fill = tk.BOTH,expand=1)
        
def ScaledReturn_Graph(showing):
    global df
    global TestSize
    try:
        df.plot(y='ScaledReturn',title=f'Scaled Return from {df.head(1).index[0]} to {df.tail(1).index[0]}',figsize=(ScreenWidth,ScreenHeight) , fontsize=8)#df[-2*TestSize:].plot #,title= TITLE )    
    except KeyError:
        Add_ScaledReturn()
        df.plot(y='ScaledReturn',title=f'Scaled Return from {df.head(1).index[0]} to {df.tail(1).index[0]}',figsize=(ScreenWidth,ScreenHeight) , fontsize=8)    
    plt.xticks(rotation=20, horizontalalignment="center",fontsize=12)
    SaveFigure('Scaled_Return','PreRun','PreRun') #plt.savefig(Figures_path+ '\Scaled_Return.jpg')
    if showing:
        plt.show()

def CLOSE_Graph(showing):
    global df
    global TestSize
    try:
        df
    except NameError as ERR:
        showmessage('df error','Dataframe is not available',TIMEOUT=3000,TYP='Erro')
        LoadDataBase('')
    TITLE=f'Price chart from {df.index[0]} to {df.index[-1]}'
    df.plot(y=['CLOSE'],title= TITLE , ylabel='Close Price' , fontsize=14,figsize=(ScreenWidth,ScreenHeight),legend=True)
    plt.xticks(rotation=20, horizontalalignment="center",fontsize=8)
    
    SaveFigure('CloseGraph','PreRun','PreRun') #plt.savefig(Figures_path+'\Close Price.jpg')
    if showing:
        plt.show()
       
def CLOSE_LOG_Graph(showing):
    plt.figure(f'fig{datetime.datetime.now().microsecond%100}',figsize=(ScreenWidth,ScreenHeight))
    plt.plot(np.log(df['CLOSE'][-TestSize*2:]))
    #plt.xticks(list(range(0,len(df),10)),rotation=45, horizontalalignment="center",fontsize=10)
    plt.title(f'Natural log of Close of {TestSize*10} candles from {df[-TestSize*10:].index[0]} to {df[-TestSize*10:].index[-1]}',fontsize=10,fontname='Times new roman')
    
    SaveFigure('Natural_Log','PreRun','PreRun') #plt.savefig(Figures_path+'\CloseGraph.jpg')
    if showing:
        plt.show()
    #df[-2*TestSize:].plot(y=['CLOSE','CLOSE(-1)','MA2','MA3','MA4','MA5'],title= TITLE , ylabel='Close Price' , fontsize=14,figsize=(10,6),legend=True)

def autocorr_plot(showing):
    from statsmodels.tsa.stattools import acf, pacf
    from statsmodels.graphics.tsaplots import plot_acf
    
    ACF = acf(df['CLOSE'][-TestSize*2:])#[-100:])
    ACF_df = pd.DataFrame(ACF,columns=['Auto-Correlation of Closes'])

    pd.plotting.autocorrelation_plot(df.CLOSE[-TestSize*2:])
    ACF_df.plot(kind='bar')
    plt.plot(ACF,'g-.')
    plt.xlabel('Lagged Closes')

    SaveFigure('CLOSE_ACF','PreRun','PreRun') #plt.savefig(Figures_path+'\CloseGraph.jpg')
    if showing:
        plt.show()   

    plot_acf(df['CLOSE'].values)
    SaveFigure('CLOSEs','PreRun','PreRun')
    if showing:
        plt.show()
    #================================
    PACF = pacf(df['CLOSE'][-TestSize*2:])
    PACF_df = pd.DataFrame(PACF)
    PACF_df.columns=['Partial Autocorrelation Function']
    PACFHigh = np.count_nonzero((PACF>0.85)==True)
    
    PACF_df.plot(kind='bar')
    plt.plot(PACF,'g-.')
    plt.xlabel('Lagged Closes')
    SaveFigure('CLOSE_PACF','PreRun','PreRun') #plt.savefig(Figures_path+'\CLOSE_PACF.jpg')
    if showing:
        plt.show()
    #================================
    price_diff = (df['CLOSE'] - df['CLOSE'].shift())[-TestSize*2:]#.dropna()
    price_diff.name='1ST Difference of Closes'
    
    ACF_diff = acf(price_diff)
    ACF_diff_df = pd.DataFrame(ACF_diff,columns=['1ST order Diff Autocorr.'])
    
    ACF_diff_df.plot.bar(color='purple')
    plt.xlabel('lagged Closes')
    #======================================
    price_diff = (df['CLOSE'] - df['CLOSE'].shift())[-TestSize*2:]#.dropna()
    price_diff.name='1ST Difference of Closes'
    
    ACF_diff = acf(price_diff)
    ACF_diff_df = pd.DataFrame(ACF_diff,columns=['1ST order Diff Autocorr.'])
    
    PACF_diff = pacf(price_diff)
    plt.plot(PACF_diff,'g-.')
    plt.plot(ACF_diff,'y-')
    plt.legend(['First Difference Partial Auto Correlation Function','First Difference Auto Correlation Function'])
    plt.title('First Difference Auto-Correlation of Closes')
    SaveFigure('PACF_diff','PreRun','PreRun') #plt.savefig(Figures_path+'\PACF_diff.jpg')
    if showing:
        plt.show()

def Scatter_Plot(showing):
    #global y_test
    #global y_pred_test
    plt.scatter(x=y_test,y=y_pred_test)
    plt.plot([y_test.min(),y_test.max()],[y_pred_test.min(),y_pred_test.max()],'c--',linewidth=0.5)
    plt.title(label='Scatter plot for test part'+msg,fontsize=5)
    plt.xlabel('Real Values')
    plt.ylabel('Predictions')
    SaveFigure('Scatter_plot',Strategy,'') #plt.savefig(Figures_path+'\\Scatter_plot.jpg')
    if showing:
        plt.show()
    
def Confusion_Matrix_Plot(showing):
    from sklearn.metrics import confusion_matrix
    global y_test
    global y_pred_test
    global msg
    if Strategy == 'No Strategy':
        showmessage('No strategy',f'{Strategy} defined!\nSelect a strategy from Run menu!',TIMEOUT=2000,TYP='warn')
    else:
        CM = confusion_matrix(y_test, y_pred_test)
        print('Confusion Matrix:\n■■■■■■■■■■\n',CM,'\n■■■■■■■■■■')
        a = heatmap(CM, annot=True, fmt= 'd')
        
        if len(msg.split('features:')[1].split(','))>2:
            msg1 = msg.split('features:')[0]+msg.split('features:')[1].split(',')[0]+'→'+msg.split('features:')[1].split(',')[-1]
            
        FigureName = str(features.columns.values).replace(' ','_').replace("'",'').replace(')','').replace('(','').replace(']','◘').replace('[','◘')
        if len(FigureName.split('_'))>2:
            FigureName = FigureName.split('_')[0]+'→'+FigureName.split('_')[-1]
        
        msg1 = msg.replace('\n',' ').replace(' with test size','\nwith test size').replace('features:','features:')
        plt.title('Confusion Matrix for\n'+msg1,fontsize=10,fontname='Times new roman', color='#474747')
        plt.xlabel('Predicted Values')
        plt.ylabel('Real Lable Values')

        SaveFigure('Confusion_Matrix_Plot',Strategy,'') #plt.savefig(Figures_path+'\\Confusion_Matrix_Plot.jpg')
        if showing:
            plt.show()  

def Equity_Plot(showing):
    global msg1
    plt.close('all')
#     if Strategy == 'No Strategy':
#         showmessage('No strategy',f'{Strategy} defined!\nSelect a strategy from Run menu!',TIMEOUT=2000,TYP='warn')
    try:
        df['equity']
    except NameError as ERR:
        showmessage('df error',f'Dataframe is not available\n{ERR}',TIMEOUT=3000,TYP='Erro')
        LoadDataBase('')
    except KeyError as ERR:
        showmessage('No strategy',f'Select and run a strategy from Run menu first!',TIMEOUT=2000,TYP='warn')
    else:
        msg1 = msg.replace('\n',' ').replace(' with test size','\ntest size')#.replace('features:','\nfeatures:')
        if len(msg1.split('features:')[1].split(','))>2:
            msg1 = msg1.split('features:')[0]+msg1.split('features:')[1].split(',')[0]+'→'+msg1.split('features:')[1].split(',')[-1]
        df[-len(y_test):].plot(y='equity', title=f'Account equity for test part & {BALANCE}$ initial deposit\n'+msg1,figsize=(ScreenWidth,ScreenHeight))#, figsize=(16,7))
        plt.xlabel('Trades');
        plt.ylabel('Equity ($)');
        plt.figtext(0.6,0.8,f"Final Balance= {df['equity'].values[-1]:.2f} $",fontsize=16)
        plt.figtext(0.2,0.4,f"Net Profit for test part= {(df['equity'][-1] - df['equity'][-len(y_test)-1]):.2f} $",fontsize=16)
        for r in df[-len(y_test):].iterrows():
            if r[1]['Compare']:
                plt.axvline(x=r[0], linewidth=0.2, alpha=0.8, color='g');
            else:
                plt.axvline(x=r[0], linewidth=0.2, alpha=0.8, color='r');
        SaveFigure('Account_Balance_Plot',Strategy,'') #plt.savefig(Figures_path+'\\Account_Balance_Plot.jpg')
        if showing:
            plt.show()
        
def DrawDown_Plot(showing):
    global MDD
    global MDDP
    try:
        df['DrawDown']
    except NameError as ERR:
        showmessage('df error',f'Dataframe is not available\n{ERR}',TIMEOUT=3000,TYP='Erro')
        LoadDataBase('')
    except KeyError as ERR:
        showmessage('No strategy',f'Select and run a strategy from Run menu first!',TIMEOUT=2000,TYP='warn')
    else:
        msg1 = msg.replace('\n',' ').replace(' with test size','\ntest size').replace('features:','\nfeatures:')
        df[-2*len(y_test):].plot(y='DrawDown', ylabel='Draw Down $',figsize=(ScreenWidth,ScreenHeight), title='Draw Down for\n'+msg1) #, fontsize=8)#,title= TITLE )#, figsize=(16,8));)
        plt.figtext(0.25,0.25,f'Max DrawDown= {MDD:0.2f}$ ')
        plt.figtext(0.5,0.25, f'Max DrawDown= {MDDP:0.1f}%')
        plt.figtext(0.38,0.15,f'initial_deposit= {BALANCE} $')
        
        SaveFigure('Draw_Down',Strategy,'') #plt.savefig(Figures_path+'\\Draw_Down.jpg')
        if showing:
            plt.show()

def Histogram_Plot(showing):
    if Strategy == 'No Strategy':
        showmessage('No strategy',f'{Strategy} defined!\nSelect a strategy from Run menu!',TIMEOUT=2000,TYP='warn')
    else:
        if len(msg.split('features:')[1].split(','))>2:
            TITLE = msg.split('features:')[0]+'features:'+msg.split('features:')[1].split(',')[0]+'→'+msg.split('features:')[1].split(',')[-1]
        else:
            TITLE = msg
            
        TITLE = TITLE.replace('\n',' ').replace(' with test size','\ntest size').replace('Test','\nTest')

        plt.hist(df[-len(y_test)-2:]['profit'],color='blue',edgecolor='black', bins=TestSize)#,figsize=(12,7)
        plt.title('Histogram of trades for '+TITLE,fontsize=10,fontname='Times new roman')
        plt.ylabel('Number of Trades')
        plt.xlabel('The profit of trades')     
        SaveFigure('Histogram',Strategy,'') #plt.savefig(Figures_path+'\\Histogram.jpg')
        if showing:
            plt.show()

def SetWindowsize(STRI):
    global WIN9
    global windowsize
    global QuickRun
    winsize  = tk.StringVar()
    def Addprev(_):
        global windowsize
        try:
            windowsize = int(PREVe_entry.get())
            if windowsize<0:
                raise RuntimeError('it\'s negative!!!')       
        except Exception as ERR:
            showmessage(f'No. of win Error',f'it\'s not valid!\n{ERR}',TIMEOUT=3000,TYP='Erro')
        else:
            showmessage(f'windowsize',f'windowsize is set to {windowsize} of previous {STRI}s',TIMEOUT=2000,TYP='info')
        WIN9.destroy()
        WIN9.quit()
  
    WIN9 = tk.Toplevel(root)
    WIN9.lift()
    WIN9.bind('<Escape>',EXIT_One)
    WIN9.title('window size')
    WIN9.configure(bg='#091A32')
    WIN9.resizable(False,False)

    #tk.Label(WIN9,image = pho).place(height=h,width=w,x=0,y=0)#.grid(rowspan=8,columnspan=2)

    PREVeChoice = tk.Label(WIN9,text=f"What's the max. size of window for {STRI}?",bg='#474747',fg='yellow',font=('times',FONTSIZE,'bold'))
    PREVeChoice.grid(row=0,columnspan = 4,pady=5)   #,sticky=tk.W 
    
    PREVLBL = tk.Label(WIN9,text=f"Window size: ",bg='#474747',fg='white',font=('times',FONTSIZE,'bold'))
    PREVLBL.grid(row=1,column=0,pady=5) #,sticky='W'

    PREVe_entry = tk.Entry(WIN9,textvariable=winsize,width=10,bg='#474747',fg='yellow',font=('times',FONTSIZE,'bold'),borderwidth=3,bd=5)
    PREVe_entry.grid(row=1,column=1,pady=5)#,sticky='W'

    btPREV = tk.Button(WIN9,text="Add to database",bg='#474747',fg='yellow',font=('times',FONTSIZE2,'bold'),relief=tk.RAISED,bd=5,command=lambda:Addprev(''))
    WIN9.bind('<Return>',Addprev)
    btPREV.grid(row=2,columnspan=2)
    WIN9.lift()
    WIN9.focus_force()
    PREVe_entry.focus()
    WIN9.mainloop()
    return windowsize

def RUNCompareLagbased(event):
    global FEAT
    global Concise_Report
    FEAT = CompareWINVAR.get()
    Concise_Report = True
    intChoice7.set(1)
    chkChoice7.config(text='Concise Report')
    Fea_MESSAGE1.configure(text = f'{FEAT}',fg='#F5CC21') 
    cmbFeat.set(FEAT)
    showmessage('Concise reports',f'Concise reports activated\nSelected features forced on: {FEAT}',TIMEOUT=4000,TYP='info')
    if FEAT == 'Just previous Closes':
        Compare_Closes()
        
    elif FEAT == 'Just previous Labels':
        Compare_Labels()
        
    elif FEAT == 'SMA window size':
        Compare_SMAs()
        
    elif FEAT == 'EMA window size':
        Compare_EMAs()
    else:
        raise RuntimeError
    
def Compare_Closes():
    global windowsize
    global QuickRun
    global Accuracy_lst
    global F1score_lst
    global precision_lst
    global recall_lst
    global Bal_Accuracy_test_lst
    global Logloss_lst
    global ROCAUC_lst
    global R2score_lst
    global MDDP_lst
    global TotalNetProfit_lst
    global Profit_ratio_lst
    global sharpe_ratio_lst
    global Concise_Report
    global LagTest
    global tempindex
    global FEAT
    global temp1
    global temp2
    global ComparisonFlag
    global comparisondf
    
    comparisondf = pd.DataFrame()
    CMPTableIndex = []
    Accuracy_lst = []
    F1score_lst = []
    precision_lst = []
    recall_lst = []
    Bal_Accuracy_test_lst = []
    Logloss_lst = []
    ROCAUC_lst = []
    R2score_lst = []
    MDDP_lst = []
    TotalNetProfit_lst = []
    Profit_ratio_lst = []
    sharpe_ratio_lst = []
    intChoice5.set(1)
    intChoice4.set(0)
    QuickRun = True
    LagTest = 'CLOSE'  
    if Strategy == 'No Strategy':
        showmessage('No strategy',f'{Strategy} defined!\nSelect a strategy from Run menu!',TIMEOUT=2000,TYP='warn')
    else:
        windowsize = SetWindowsize('Close')
        msg1 = msg.replace('\n',' ').replace(' with test size','\ntest size').replace('features:','\nfeatures:')

        msg1 = msg1.replace('\n',' ').replace(' with test size','\ntest size').replace('features:','\nfeatures:')
        for C in range(1,windowsize+1):
            Just_Close_Features(C)
            ComparisonFlag = f'{C} of previous Closes'
            RUNStrategy('NormalMode')
            if len(Strategy.split())>2:
                Short_Strategy = ''.join(x[0] for x in Strategy.split())
                CMPTableIndex.append(f'{Short_Strategy} window of {C} Closes')
            else:
                CMPTableIndex.append(f'{Strategy} window of {C} Closes') 
        Generate_Comparison_Table(CMPTableIndex,'Close Comparison','closes')
        Select_Comparison_Graph(CMPTableIndex,temp1, temp2)
        temp1 = f'CloseCompare {windowsize}'
        temp2 = 'closes'
        tempindex = CMPTableIndex
        
def Compare_Labels():
    global windowsize
    global QuickRun
    global Accuracy_lst
    global F1score_lst
    global precision_lst
    global recall_lst
    global Bal_Accuracy_test_lst
    global Logloss_lst
    global ROCAUC_lst
    global R2score_lst
    global MDDP_lst
    global TotalNetProfit_lst
    global Profit_ratio_lst
    global sharpe_ratio_lst
    global CMPTableIndex
    global Concise_Report
    global LagTest
    global LABL
    global tempindex
    global FEAT
    global temp1
    global temp2
    global ComparisonFlag
    global comparisondf

    comparisondf = pd.DataFrame()
    Accuracy_lst = []
    F1score_lst = []
    precision_lst = []
    recall_lst = []
    Bal_Accuracy_test_lst = []
    Logloss_lst = []
    ROCAUC_lst = []
    R2score_lst = []
    MDDP_lst = []
    TotalNetProfit_lst = []
    Profit_ratio_lst = []
    sharpe_ratio_lst = []
    CMPTableIndex = []
    intChoice5.set(1)
    intChoice4.set(0)
    QuickRun = True
    LagTest = 'LABELS'
    
    if Strategy == 'No Strategy':
        showmessage('No strategy',f'{Strategy} defined!\nSelect a strategy from Run menu!',TIMEOUT=2000,TYP='warn')
    else:
        windowsize = SetWindowsize('Label')
        msg1 = msg.replace('\n',' ').replace(' with test size','\ntest size').replace('features:','\nfeatures:')
        for C in range(1,windowsize+1):
            Just_Label_Features(C)
            ComparisonFlag = f'{C} of previous {LABL}'
            RUNStrategy('NormalMode')
            if len(Strategy.split())>2:
                Short_Strategy = ''.join(x[0] for x in Strategy.split())
                CMPTableIndex.append(f'{Short_Strategy} window of {C} Labels')
            else:
                CMPTableIndex.append(f'{Strategy} window of {C} Labels') 
        Generate_Comparison_Table(CMPTableIndex,'Label Comparison','labels')
        temp1 = f'LableCompare{windowsize}'
        temp2 = 'labels'
        Select_Comparison_Graph(CMPTableIndex,temp1, temp2)
        tempindex = CMPTableIndex

def Compare_SMAs(): 
    global windowsize
    global QuickRun
    global Accuracy_lst
    global F1score_lst
    global precision_lst
    global recall_lst
    global Bal_Accuracy_test_lst
    global Logloss_lst
    global ROCAUC_lst
    global R2score_lst
    global MDDP_lst
    global TotalNetProfit_lst
    global Profit_ratio_lst
    global sharpe_ratio_lst
    global CMPTableIndex
    global Concise_Report
    global LagTest
    global LABL
    global tempindex
    global FEAT
    global temp1
    global temp2
    global ComparisonFlag
    global comparisondf

    comparisondf = pd.DataFrame()
    Accuracy_lst = []
    F1score_lst = []
    precision_lst = []
    recall_lst = []
    Bal_Accuracy_test_lst = []
    Logloss_lst = []
    ROCAUC_lst = []
    R2score_lst = []
    MDDP_lst = []
    TotalNetProfit_lst = []
    Profit_ratio_lst = []
    sharpe_ratio_lst = []
    CMPTableIndex = []
    intChoice5.set(1)
    intChoice4.set(0)
    QuickRun = True
    LagTest = 'SMAs'
    
    if Strategy == 'No Strategy':
        showmessage('No strategy',f'{Strategy} defined!\nSelect a strategy from Run menu!',TIMEOUT=2000,TYP='warn')
    else:
        windowsize = SetWindowsize('SMA')
        msg1 = msg.replace('\n',' ').replace(' with test size','\ntest size').replace('features:','\nfeatures:')
        for C in range(1,windowsize+1):
            Select_SMAs(C)
            ComparisonFlag = f'SMA with window of {C}'
            RUNStrategy('NormalMode')
            if len(Strategy.split())>2:
                Short_Strategy = ''.join(x[0] for x in Strategy.split())
                CMPTableIndex.append(f'{Short_Strategy} SMA{C} ')
            else:
                CMPTableIndex.append(f'{Strategy} SMA{C}') 
        Generate_Comparison_Table(CMPTableIndex,'SMA Comparison','SMAs')
        temp1 = f'SMACompare{windowsize}'
        temp2 = 'SMAs'
        Select_Comparison_Graph(CMPTableIndex,temp1, temp2)
        tempindex = CMPTableIndex
        
def Select_SMAs(c):
    global features
    ADD_SMAs(c) 
    showmessage('SMA windowsize',f'SMA of windowsize {c} is fed to the model',TIMEOUT=2000,TYP='info')      
    features = df[[f'SMA{c}']]
    features = features.shift()
    features.fillna(0,inplace=True)
        
def Compare_EMAs(): 
    global windowsize
    global QuickRun
    global Accuracy_lst
    global F1score_lst
    global precision_lst
    global recall_lst
    global Bal_Accuracy_test_lst
    global Logloss_lst
    global ROCAUC_lst
    global R2score_lst
    global MDDP_lst
    global TotalNetProfit_lst
    global Profit_ratio_lst
    global sharpe_ratio_lst
    global CMPTableIndex
    global Concise_Report
    global LagTest
    global LABL
    global tempindex
    global FEAT
    global temp1
    global temp2
    global ComparisonFlag
    global comparisondf

    comparisondf = pd.DataFrame()
    Accuracy_lst = []
    F1score_lst = []
    precision_lst = []
    recall_lst = []
    Bal_Accuracy_test_lst = []
    Logloss_lst = []
    ROCAUC_lst = []
    R2score_lst = []
    MDDP_lst = []
    TotalNetProfit_lst = []
    Profit_ratio_lst = []
    sharpe_ratio_lst = []
    CMPTableIndex = []
    intChoice5.set(1)
    intChoice4.set(0)
    QuickRun = True
    LagTest = 'EMAs'
    
    if Strategy == 'No Strategy':
        showmessage('No strategy',f'{Strategy} defined!\nSelect a strategy from Run menu!',TIMEOUT=2000,TYP='warn')
    else:
        windowsize = SetWindowsize('EMA')
        msg1 = msg.replace('\n',' ').replace(' with test size','\ntest size').replace('features:','\nfeatures:')
        for C in range(1,windowsize+1):
            Select_EMAs(C)
            ComparisonFlag = f'EMA with window of {C}'
            RUNStrategy('NormalMode')
            if len(Strategy.split())>2:
                Short_Strategy = ''.join(x[0] for x in Strategy.split())
                CMPTableIndex.append(f'{Short_Strategy} EMA{C} ')
            else:
                CMPTableIndex.append(f'{Strategy} EMA{C}') 
        Generate_Comparison_Table(CMPTableIndex,'EMA Comparison','EMAs')
        temp1 = f'EMACompare{windowsize}'
        temp2 = 'EMAs'
        Select_Comparison_Graph(CMPTableIndex,temp1, temp2)
        tempindex = CMPTableIndex
    
def Select_EMAs(c):
    global features
    ADD_EMAs(c) 
    showmessage('EMA windowsize',f'EMA of windowsize {c} is fed to the model',TIMEOUT=2000,TYP='info')      
    features = df[[f'EMA{c}']]
    features = features.shift()
    features.fillna(0,inplace=True)
    
def Select_Comparison_Graph(X,Name,MOD):
    def select_Display_items(event):
        nonlocal selectednumbers
        global selecteditems
        selecteditems = []
        selectednumbers = []
        for J in DiaplayListBOX.curselection():
            selection = DiaplayListBOX.get(J)
            selecteditems.append(selection)
            selectednumbers.append(J)   
        if not selecteditems:
            selecteditems = DisplayOptionsSet
            selectednumbers = list(range(len(X)))
        showmessage('selected items',f"selected items for drawing plots:\n{selecteditems}",TIMEOUT=3000,TYP='info')
            
    def ColumnPlotComparison(Metr,COL):
        TotalValues = eval(f'{Metr}_lst')
        valuesToDisplay = [TotalValues[j] for j in selectednumbers]
        print('selecteditems:',selecteditems)
        print('selectednumbers:',selectednumbers)
        print('TotalValues:',TotalValues)
        print('valuesToDisplay:',valuesToDisplay)

        TITLE = 'Symbol &'+ msg.split('Symbol &')[1]
        plt.figure(f'fig{datetime.datetime.now().microsecond%100}',figsize=(ScreenWidth,ScreenHeight))
        
        if MOD == 'features':
            if len(Strategy) >14:
                Short_Strategy = Strategy.split()[0]+' '+''.join(x[0].upper() for x in Strategy.split()[1:])
            else:
                Short_Strategy = Strategy
                
            plt.ylabel(f'\n{Metr} with {"NovelSplit" if SingleCandlePredict else "NormalSplit"} for {Short_Strategy}')
        elif MOD == 'models':
            STR = str(features.columns.values).replace("'","").strip()
            plt.ylabel(f'\n{Metr} with {"NovelSplit" if SingleCandlePredict else "NormalSplit"} for {STR}')
        else:          
            plt.ylabel(f'\n{Metr} with {"NovelSplit" if SingleCandlePredict else "NormalSplit"} for {MOD}')
        if not selecteditems:
            select_Display_items('')
        
        if LagTest:
            TITLE = TITLE[:TITLE.find(LagTest+'(')+len(LagTest)+4]+'→'+ TITLE[TITLE.rfind(LagTest+'('):].replace('  ','')
            TITLE = f'{Metr} for {windowsize} different sizes of {LagTest} '+TITLE.replace('\n',' ').replace(' with test size','Test size').replace('train_test','\nTrain_Test')    
            plt.xlabel(f'{len(selecteditems)} different size for lagged {LagTest} as features')
            plt.bar(selecteditems,valuesToDisplay,color= COL,edgecolor='black')
            plt.xticks(rotation=2*len(selecteditems), horizontalalignment="center",fontsize=7,ticks = range(len(selecteditems)) , labels = selecteditems)
        else:
            TITLE = f'{Metr} Evaluationfor multiple {MOD} in '+TITLE.replace('\n',' ').replace(' with test size','\ntest size').replace('Test','\nTest')
            plt.xlabel(f'Multiple {MOD}')
            plt.bar(selecteditems,valuesToDisplay,color=COL,edgecolor='black')
            plt.xticks(rotation=2*len(selecteditems), horizontalalignment="center",fontsize=7,ticks = range(len(X)) , labels = selecteditems)
        
        #plt.tight_layout()
        plt.title(TITLE,fontsize=10,fontname='Times new roman')
        plt.tight_layout()
        plt.subplots_adjust(top=0.95)
        if MOD == 'features':
            STRIng = r'Comparison\\multiple features\\'+str(Strategy)
        elif MOD == 'models':
            STRIng = str(features.columns.values).replace("'","").strip()
            STRIng = f'Comparison\\multiple models\\{STRIng}'      
        else:
            STRIng = f'Comparison\\Different Windows of {MOD}'
            
        SaveFigure(f'{Metr}_{Name}', STRIng,MOD)
        plt.show()
        
    def ContinuousPlotComparison(Metr):
        #valuesToDisplay = [TotalValues[j] for j in selectednumbers]
        TITLE = 'Symbol &'+ msg.split('Symbol &')[1]
        #plt.figure(101,figsize=(ScreenWidth,ScreenHeight))
        if not selecteditems:
            select_Display_items('')
        if LagTest:
            TITLE = TITLE[:TITLE.find(LagTest+'(')+len(LagTest)+4]+'→'+ TITLE[TITLE.rfind(LagTest+'('):].replace('  ','')
            TITLE = f'{Metr} for {windowsize} different sizes of {LagTest} '+TITLE.replace('\n',' ').replace(' with test size','\ntest size').replace('Test','\nTest')

        else:
            TITLE = f'{Metr} for multiple {MOD} in '+TITLE.replace('\n',' ').replace(' with test size','\ntest size').replace('Test','\nTest')

        if MOD == 'features':
            STRIng = r'Comparison\\multiple features\\'+str(Strategy)
        elif MOD == 'models':
            STRIng = str(features.columns.values).replace("'","").strip()
            STRIng = f'Comparison\\multiple models\\{FEAT}_{STRIng}'
        else:
            STRIng = f'Comparison\\DifferentWindow{MOD}s'
            
        if Metr == 'equity':
            comparisondf.dropna(axis=1, inplace = True)  
            comparisondf.iloc[-len(y_test):,selectednumbers].plot(title=TITLE,figsize=(ScreenWidth,ScreenHeight)) 
            plt.xlabel('Trades')
            plt.ylabel(f'Account equity for test part with initial deposit of {BALANCE}$')
            SaveFigure(f'Equityplot_{Name}', STRIng, MOD)
            plt.show()

        plt.title(TITLE,fontsize=10,fontname='Times new roman')
        plt.tight_layout()
        plt.subplots_adjust(top=0.85)
        
    try:
        df
    except NameError as ERR:
        showmessage('df error','Dataframe is not available',TIMEOUT=3000,TYP='Erro')
        LoadDataBase('')
        Select_Comparison_Graph(X,Name,MOD)
    else:   

        WIN20 = tk.Toplevel(root , bd = 5)#,cursor=)
        WIN20.title('Graph selection')
        WIN20.bind('<Escape>',EXIT_One)
        WIN20.geometry('+12+12')
        BT1 = tk.Button(WIN20,text="֎Plot Accuracy Comparison֎",bg='#474747',fg='gold',font=('times',FONTSIZE2,'bold'),relief=tk.RAISED,bd=5,command =lambda:ColumnPlotComparison('Accuracy','darkgreen'))
        BT1.pack(fill = tk.BOTH,expand=1)
        
        BT2 = tk.Button(WIN20,text="֎Plot Bal_Accuracy_test Comparison֎",bg='#474747',fg='yellow',font=('times',FONTSIZE2,'bold'),relief=tk.RAISED,bd=5,command = lambda:ColumnPlotComparison('Bal_Accuracy_test','green'))
        BT2.pack(fill = tk.BOTH,expand=1)
        
        BT3 = tk.Button(WIN20,text="֎Plot precision Comparison",bg='#474747',fg='gold',font=('times',FONTSIZE2,'bold'),relief=tk.RAISED,bd=5,command=lambda:ColumnPlotComparison('precision','saddlebrown'))
        BT3.pack(fill = tk.BOTH,expand=1)
        
        BT4 = tk.Button(WIN20,text="֎Plot recall Comparison֎",bg='#474747',fg='yellow',font=('times',FONTSIZE2,'bold'),relief=tk.RAISED,bd=5,command=lambda:ColumnPlotComparison('recall','maroon'))
        BT4.pack(fill = tk.BOTH,expand=1)
        
        BT5 = tk.Button(WIN20,text="֎Plot F1score Comparison֎",bg='#474747',fg='gold',font=('times',FONTSIZE2,'bold'),relief=tk.RAISED,bd=5,command=lambda:ColumnPlotComparison('F1score','midnightblue'))
        BT5.pack(fill = tk.BOTH,expand=1)
        
        BT6 = tk.Button(WIN20,text="֎Plot ROCAUC Comparison֎",bg='#474747',fg='yellow',font=('times',FONTSIZE2,'bold'),relief=tk.RAISED,bd=5,command=lambda:ColumnPlotComparison('ROCAUC','darkgreen'))
        BT6.pack(fill = tk.BOTH,expand=1)
        
        BT7 = tk.Button(WIN20,text="֎Plot R2score Comparison֎",bg='#474747',fg='gold',font=('times',FONTSIZE2,'bold'),relief=tk.RAISED,bd=5,command=lambda:ColumnPlotComparison('R2score','darkslateblue'))
        BT7.pack(fill = tk.BOTH,expand=1)

        BT8 = tk.Button(WIN20,text="֎Plot Logloss Comparison֎",bg='#474747',fg='yellow',font=('times',FONTSIZE2,'bold'),relief=tk.RAISED,bd=5,command=lambda:ColumnPlotComparison('Logloss','saddlebrown'))
        BT8.pack(fill = tk.BOTH,expand=1)
        
        BT9 = tk.Button(WIN20,text="֎Plot MDDP Comparison֎",bg='#474747',fg='yellow',font=('times',FONTSIZE2,'bold'),relief=tk.RAISED,bd=5,command=lambda:ColumnPlotComparison('MDDP','darkgreen'))
        BT9.pack(fill = tk.BOTH,expand=1)
        
        BT10 = tk.Button(WIN20,text="֎Plot Total Net Profit Comparison֎",bg='#474747',fg='yellow',font=('times',FONTSIZE2,'bold'),relief=tk.RAISED,bd=5,command=lambda:ColumnPlotComparison('TotalNetProfit','midnightblue'))
        BT10.pack(fill = tk.BOTH,expand=1)
        
        BT11 = tk.Button(WIN20,text="֎Plot Profit ratio Comparison֎",bg='#474747',fg='yellow',font=('times',FONTSIZE2,'bold'),relief=tk.RAISED,bd=5,command=lambda:ColumnPlotComparison('Profit_ratio','darkslategray'))
        BT11.pack(fill = tk.BOTH,expand=1)
        
        BT12 = tk.Button(WIN20,text="֎Plot Sharpe ratio Comparison֎",bg='#474747',fg='yellow',font=('times',FONTSIZE2,'bold'),relief=tk.RAISED,bd=5,command=lambda:ColumnPlotComparison('sharpe_ratio','maroon'))
        BT12.pack(fill = tk.BOTH,expand=1)
        
        BT13 = tk.Button(WIN20,text="֎Plot Equity Comparison֎",bg='#474747',fg='yellow',font=('times',FONTSIZE2,'bold'),relief=tk.RAISED,bd=5,command=lambda:ContinuousPlotComparison('equity'))
        BT13.pack(fill = tk.BOTH,expand=1)
        
        global comparisondf
        global selecteditems
        selecteditems = []
        selectednumbers = []        
        DisplayOptionsSet = X #features.columns.to_list() 
        WIN19 = tk.Toplevel(root)
        WIN19.title('Choose for drawing')
        WIN19.bind('<Escape>',EXIT_One)
        WIN19.bind('<Return>',select_Display_items)
        WIN19.geometry(f'280x345+{get_monitors()[0].width//2}+313')
        WIN19.lift()
        WIN19.focus_force()
        DiaplayListBOX = tk.Listbox(WIN19, width=40, height=30,listvariable=tk.Variable(value = DisplayOptionsSet),font=('Times', FONTSIZE, 'bold'), selectmode = tk.MULTIPLE)
        DiaplayListBOX.place(height=313, width=300, x=1, y=30) 
        scrollbar1 = tk.Scrollbar(WIN19, orient=tk.VERTICAL, command=DiaplayListBOX.yview)  
        scrollbar1.place(x=250, y=0, height=343,width=300)   
        DiaplayListBOX.config(yscrollcommand=scrollbar1.set)
    #     scrollbar2 = tk.Scrollbar(frame1, orient=tk.HORIZONTAL, command=DiaplayListBOX.xview)  
    #     DiaplayListBOX.config(xscrollcommand=scrollbar2.set)
    #     scrollbar2.place(x=0, y=380,width=313) 
        SelectDisplayBTN = tk.Button(WIN19, text='↓Select Some to draw plot↓',font=('Times', FONTSIZE, 'bold'),fg='white', bg='darkgoldenrod',command = lambda:select_Display_items(''))  
        SelectDisplayBTN.place(height=30,x=0,y=0)
        
def Generate_Comparison_Table(X,folder,MOD):
    global FEAT
    global TableFile_path
    FileNameString = str(features.columns.values).replace(' ','_').replace("'",'').replace(')','').replace('(','').replace(']','▬').replace('[','▬')
    if FEAT =='Just previous Closes' or FEAT == 'Just previous Labels' or FEAT == 'Custom Window Features':
        if len(FileNameString.split('_'))>2:
            FileNameString = FileNameString.split('_')[0]+'→'+FileNameString.split('_')[-1]
        elif ComparisonFlag:
            FileNameString = f'▬Maxwindow→{FileNameString}'
        
    Short_Strategy = ''.join(x[0] for x in Strategy.split())
    STRING = '' if MOD == 'models' else Short_Strategy
#     if MOD == 'features':
#         FileNameString = f'Compare{STRING} with multiple features♦TF_{TimeFrame}♦TestSize_{TestSize}♦Trainsize_{Trainsize}♦Lbl_{LABL}♦{"NovelSplit" if SingleCandlePredict else "NormalSplit"}'
#     elif MOD == 'SMAs':
#         FileNameString = f'Compare_{STRING} with multiple SMAs♦TF_{TimeFrame}♦TestSize_{TestSize}♦Trainsize_{Trainsize}♦Lbl_{LABL}♦{"NovelSplit" if SingleCandlePredict else "NormalSplit"}'      
#     else:
    FileNameString = f'Comparison_{STRING} with multiple {MOD}♦TF_{TimeFrame}♦TestSize_{TestSize}♦Trainsize_{Trainsize}♦Feature_{FEAT}_{FileNameString}Lable_{LABL}♦{"NovelSplit" if SingleCandlePredict else "NormalSplit"}'
    Comparison_path = os.getcwd()+f'\\Comparisons\\{folder}'
    if not os.path.exists(Comparison_path):
        showmessage('Finding path',f'The path {Comparison_path} was not found!\nIt will be created!',TIMEOUT= 3000,TYP='warn')
        try:
            os.makedirs(Comparison_path)
        except OSError as ERR:
            showmessage('Path Error',f'Check the file path or name: {Comparison_path}\n{ERR}',TIMEOUT=3000,TYP='Erro')
            Comparison_path = Comparison_path.replace(' ','_').replace("'",'').replace(')','').replace('(','').replace(']','▬').replace('[','▬')
            Comparison_path = filedialog.asksaveasfilename(initialdir = Comparison_path ,defaultextension=".txt", filetypes=[("text file", "*.txt"),  ("All files", "*.*")],initialfile= 'test')
            if not os.path.exists(Comparison_path):
                showmessage('Finding path',f'The path {Comparison_path} was not found!\nIt will be created!',TIMEOUT= 3000,TYP='warn')
        
    TableFile_path = filedialog.asksaveasfilename(initialdir = Comparison_path ,defaultextension=".xlsx", filetypes=[("Excel File", "*.xlsx"), ("CSV file", "*.csv"),  ("All files", "*.*")],initialfile= FileNameString)
    ComparisonColumns = {'Accuracy':Accuracy_lst,'Balanced Accuracy':Bal_Accuracy_test_lst,'precision':precision_lst,\
                         'Recall':recall_lst, 'F1score':F1score_lst, 'R2score':R2score_lst, 'ROCAUC':ROCAUC_lst, 'Logloss':Logloss_lst,\
                         'MDDP':MDDP_lst, 'Total Net Profit':TotalNetProfit_lst, 'Profit ratio':Profit_ratio_lst, 'sharpe ratio':sharpe_ratio_lst}
    Table = pd.DataFrame(ComparisonColumns,index = X)
#         Table.to_csv(TableFile_path+'\\ComparisonData\\'+FileNameString+r'.csv')
#         Table.to_excel(TableFile_path+'\\ComparisonData\\'+FileNameString+r'.xlsx')
    if TableFile_path:
        try:
            if TableFile_path[-4:].lower()=='xlsx':
                Table.to_excel(TableFile_path)
                workbook = load_workbook(TableFile_path)
                sheet = workbook.active
                sheet['A1'] = 'Experiment Metrics→\n↓Model_SplitMethod_Features↓'
                for row in sheet.iter_rows():
                    for cell in row:
                        cell.alignment = Alignment(wrap_text= True,horizontal='center', vertical='center')
                for cell in sheet[1]:
                    cell.alignment = Alignment(textRotation=90)
                sheet['A1'].alignment = Alignment(textRotation=0)
                
                workbook.save(TableFile_path)
                Response = messagebox.askyesno("Save/open", 'Comparison file Saved successfully\nDo you want to open it?',default=messagebox.YES)
                if Response:
                    subprocess.run(['start', 'excel', TableFile_path], shell=True)
            else:
                Table.to_csv(TableFile_path,sep=',')
                Response = messagebox.askyesno("Saving", 'Comparison file Saved successfully\nDo you want to open it?',default=messagebox.YES)
                if Response:
                    subprocess.run(r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE")
            showmessage('Save',"Comparison database Saved successfully",TIMEOUT=1500,TYP='info')
            VoiceAnnounce('Comparison Table created')
            #os.startfile(TableFile_path)

        except FileNotFoundError as ERR:
            showmessage('FileNotFound Error',f'Check the file path or name: {TableFile_path}\n{ERR}',TIMEOUT=2000,TYP='Erro')
            print(f'Check the file path or name: {TableFile_path}\n{ERR}')
            response = messagebox.askretrycancel("Save error",'Retry Saving?')
            if  response:
                VoiceAnnounce('Try again')
                Generate_Comparison_Table(X,folder,MOD)   
            else:
                VoiceAnnounce('Cancel')          
        except PermissionError as ERR:
            showmessage('Permission Error',f'Check if file is already opened!\n{ERR}',TIMEOUT=3000,TYP='Erro')
            Generate_Comparison_Table(X,folder,MOD)
    else:
        showmessage("Save", 'Comparison file wasn\'t Saved!',TIMEOUT=2000,TYP='warn')
        
def OpenGenaratedTable():
    Response = messagebox.askyesno("Open table", 'Do you want to open the table?',default=messagebox.YES)
    if Response:
        subprocess.run(['start', 'excel', TableFile_path], shell=True)


def CopyCompare():
    try:
        SaveDict={'Symbol':Symbol, 'Lottage':Lottage, 'TimeFrame':TimeFrame, 'Strategy':Strategy, 'FEAT':FEAT, 'LABL':LABL, 'Scaling_Type':Scaling_Type,'LoginFlag1':LoginFlag1, 'LoginFlag2':LoginFlag2,\
                  'TestSize':TestSize, 'Trainsize':Trainsize,'SingleCandlePredict':SingleCandlePredict,'msg':msg,'TableFile_path':TableFile_path,'LagTest':LagTest,'windowsize':windowsize,'tempindex':tempindex,'temp1':temp1,'temp2':temp2, 'Accuracy_lst' : Accuracy_lst, 'F1score_lst' : F1score_lst, 'precision_lst' : precision_lst,\
                'recall_lst':recall_lst, 'Bal_Accuracy_test_lst':Bal_Accuracy_test_lst, 'Logloss_lst': Logloss_lst,'ROCAUC_lst':ROCAUC_lst,'R2score_lst':R2score_lst,\
                  'MDDP_lst':MDDP_lst, 'TotalNetProfit_lst':TotalNetProfit_lst, 'Profit_ratio_lst':Profit_ratio_lst, 'sharpe_ratio_lst':sharpe_ratio_lst }
        
        if LoginFlag1 or LoginFlag2:
           SaveDict['Sdate'] = Sdate
           SaveDict['Edate'] = Edate
        if len(Strategy.split())>2:
            Short_Strategy = ''.join(x[0] for x in Strategy.split())
        copypath = DataPath+f'\\CopiedComparison\\{Strategy}'
        if not os.path.exists(copypath):
            showmessage('Finding path',f'The path {copypath} was not found!\nIt will be created!',TIMEOUT=2500,TYP='warn')
            os.makedirs(copypath)
        FileNameString = str(features.columns.values).replace(' ','_').replace("'",'').replace(')','').replace('(','').replace(']','◘').replace('[','◘')
        if len(FileNameString.split('_'))>2:
            FileNameString = FileNameString.split('_')[0]+'→'+FileNameString.split('_')[-1]

        Short_Strategy = ''.join(x[0] for x in Strategy.split())
        FileNameString = f'ComparisonData_NetProfit►{int(df.tail(1)["equity"].values[0]- BALANCE)}$_{Short_Strategy}♦TF_{TimeFrame}♦TestSize_{TestSize}♦Trainsize_{Trainsize}_Feature→{FEAT}_{FileNameString}♦Lable_{LABL}_{"NovelSplit" if SingleCandlePredict else "NormalSplit"}'
        File_path = filedialog.asksaveasfilename(initialdir=copypath,defaultextension=".txt", filetypes=[("Text files", "*.txt"), ("All files", "*.*")],initialfile= FileNameString)
                 
    except PermissionError as ERR:
        showmessage('Permission Error',f'Check if file is already opened!\n{ERR}',TIMEOUT=3000,TYP='Erro')
        Generate_Comparison_Table(X,folder,MOD)
    
    except NameError as ERR:
        showmessage('Data Shortage',f'Data is incomplete!\n{ERR}',TIMEOUT=3000,TYP='Erro')
  
    else:
        print('File_path',File_path)
        with open(File_path,'w',encoding="utf-8") as CopyCompareFile:
            print(SaveDict,file=CopyCompareFile)
        file_name = File_path.split('.')[0]
        features.to_csv(file_name+r'_features'+'.csv')
        df.to_csv(file_name+r'_df'+'.csv')
        labels.to_csv(file_name+r'_labels'+'.csv')
        comparisondf.to_csv(file_name+r'_comparisondf'+'.csv')
        y_test.to_csv(file_name+r'_y_test'+'.csv')
        VoiceAnnounce('Comparison data saved')
        showmessage('Comparison',f"Comparison files saved successfully!\n@ {File_path}",TIMEOUT=4500,TYP='info')
        print(f"Comparison files saved successfully!\n@ {File_path}")
        
def LoadCompare():
    global LABL
    global FEAT
    global Symbol
    global Lottage
    global TimeFrame
    global Strategy
    global labels
    global features
    global df
    global y_test
    global comparisondf
    global TestSize
    global Trainsize
    global Scaling_Type
    global LoginFlag1
    global LoginFlag2
    global Sdate
    global Edate
    global msg
    global SingleCandlePredict
    global Accuracy_lst
    global F1score_lst
    global precision_lst
    global recall_lst
    global Bal_Accuracy_test_lst
    global Logloss_lst
    global ROCAUC_lst
    global R2score_lst
    global MDDP_lst
    global TotalNetProfit_lst
    global Profit_ratio_lst
    global sharpe_ratio_lst
    global CMPTableIndex
    global tempindex
    global temp1
    global temp2
    global LagTest
    global TableFile_path
    copypath = DataPath+r'\CopiedComparison'
    try:
        file_path = filedialog.askopenfilename(initialdir=copypath,defaultextension=".txt", filetypes=[("Text files", "*.txt"),("All files", "*.*")], initialfile = "Text files")

        with open(file_path,'r') as CopyCompareFile:
            ComparisonData = CopyCompareFile.read()
        print('file_path:',file_path)
        file_name = file_path.split('.')[0]
    
        features= pd.read_csv(file_name+r'_features'+'.csv')
        features.index= features.DATE
        features.drop('DATE', axis=1, inplace = True)
        df = pd.read_csv(file_name+r'_df'+'.csv')
        df.index= df.DATE
        df.drop('DATE', axis=1, inplace = True)
        
        labels = pd.read_csv(file_name+r'_labels'+'.csv')
        labels.index= labels.DATE
        labels.drop('DATE', axis=1, inplace = True)

        comparisondf = pd.read_csv(file_name+r'_comparisondf'+'.csv')
        comparisondf.index= comparisondf.DATE
        comparisondf.drop('DATE', axis=1, inplace = True)
        
        y_test = pd.read_csv(file_name+r'_y_test'+'.csv')
        y_test.index= y_test.DATE
        y_test.drop('DATE', axis=1, inplace = True)

        ComparisonData = eval(ComparisonData)
        FEAT = ComparisonData['FEAT']
        LABL = ComparisonData['LABL']
        Symbol = ComparisonData['Symbol']
        Lottage = ComparisonData['Lottage']
        TimeFrame = ComparisonData['TimeFrame']
        Strategy = ComparisonData['Strategy']
        TestSize = ComparisonData['TestSize']
        Trainsize = ComparisonData['Trainsize']
        Scaling_Type = ComparisonData['Scaling_Type']
        LoginFlag1 = ComparisonData['LoginFlag1']
        LoginFlag2 = ComparisonData['LoginFlag2']
        SingleCandlePredict = ComparisonData['SingleCandlePredict']
        msg = ComparisonData['msg']        
        windowsize = ComparisonData['windowsize']
        tempindex = ComparisonData['tempindex']
        temp1 = ComparisonData['temp1']
        temp2 = ComparisonData['temp2']
        TableFile_path = ComparisonData['TableFile_path']
        LagTest = ComparisonData['LagTest']
        Accuracy_lst = ComparisonData['Accuracy_lst']
        F1score_lst = ComparisonData['F1score_lst']
        precision_lst = ComparisonData['precision_lst']
        recall_lst = ComparisonData['recall_lst']
        Bal_Accuracy_test_lst = ComparisonData['Bal_Accuracy_test_lst']
        Logloss_lst = ComparisonData['Logloss_lst']
        ROCAUC_lst = ComparisonData['ROCAUC_lst']
        R2score_lst = ComparisonData['R2score_lst']
        MDDP_lst = ComparisonData['MDDP_lst']
        sharpe_ratio_lst = ComparisonData['sharpe_ratio_lst']
        TotalNetProfit_lst = ComparisonData['TotalNetProfit_lst']
        Profit_ratio_lst = ComparisonData['Profit_ratio_lst']
               
        if LoginFlag1 or LoginFlag2:
            Sdate = ComparisonData['Sdate']
            Edate = ComparisonData['Edate']
            warning = 'Dataloaded is from onlinesource\nDon\'t forget to get online\nError in loading Sdate and Edate on line}'
            showmessage('onlinedata history',warning,TIMEOUT=3000,TYP='warn')
            
    except PermissionError as ERR:
        showmessage('Permission Error',f'Check if file is already opened!\n{ERR}',TIMEOUT=3000,TYP='Erro')

#     except FileNotFoundError as ERR:
#         showmessage('FileNotFound Error',f'Create the Comparing file first!\n{ERR}',TIMEOUT=3000,TYP='Erro')
        
    else:
        VoiceAnnounce('Comparing data loaded.')
        showmessage('Comparing Data Loaded',ComparisonData,TIMEOUT=10000,TYP='info')
        root.bind('<Return>', RUNStrategy_)
        DataBaseBTN.config(text = 'Loaded with\nComparing Data',fg='light green',font=('Times',FONTSIZE, 'bold'))
        MODLBTN.config(text = Strategy+' Selected',fg='yellow',font='normal',bg='mediumvioletred')
        ModelVAR.set(Strategy)
        RUNBTN.configure(text=f"RUN {Strategy}\n with test size:{TestSize} and train size:{Trainsize}\nTimeFrame:{TimeFrame}", fg='chartreuse',font=('Times', FONTSIZE, 'normal'))
        print('Timeframe=',TimeFrame)
        cmbCategories.set(TimeFrame)
        TF_MESSAGE.config(text = f'{TimeFrame} Timeframe is selected')
        Fea_MESSAGE1.configure(text = f'{FEAT}',fg='#F5CC21')
        cmbFeat.set(FEAT)
        LB_MESSAGE1.configure(text = f'Selected label: {LABL}',fg='#F5CC21')
        cmbLab.set(LABL)
        #chkChoice3.configure(text= ("Novel Split" if SingleCandlePredict else "Normal Split"))
        intChoice3.set(1 if SingleCandlePredict else 0)
        intChoice6.set(0 if SingleCandlePredict else 1)
        intChoice4.set(0 if QuickRun else 1)
        intChoice5.set(1 if QuickRun else 0)

def CompareModels(event):
    global QuickRun
    global Accuracy_lst
    global F1score_lst
    global precision_lst
    global recall_lst
    global Bal_Accuracy_test_lst
    global Logloss_lst
    global ROCAUC_lst
    global R2score_lst
    global MDDP_lst
    global TotalNetProfit_lst
    global Profit_ratio_lst
    global sharpe_ratio_lst
    global FEAT
    global Strategy
    global Concise_Report
    global tempindex
    global CMPTableIndex
    global temp1
    global temp2
    global ComparisonFlag
    global comparisondf
    
    comparisondf = pd.DataFrame()
    Accuracy_lst = []
    F1score_lst = []
    precision_lst = []
    recall_lst = []
    Bal_Accuracy_test_lst = []
    Logloss_lst = []
    ROCAUC_lst = []
    R2score_lst = []
    MDDP_lst = []
    sharpe_ratio_lst = []
    TotalNetProfit_lst = []
    Profit_ratio_lst = []
    CMPTableIndex = []
    intChoice5.set(1)
    intChoice4.set(0)
    QuickRun = True
    Concise_Report = True
    intChoice7.set(1)
    chkChoice7.config(text='Concise Report')
    showmessage('Concise reports',f'Concise reports activated',TIMEOUT=2500,TYP='info')
    
    BootstrapLabelWindowDict={'Gaussian Naïve Bayes':21,'Logistic Regression':3,'Multi Layer Perceptron Classifier':17,
                              'Decision Tree Classifier':25}
    FEAT = CompareFEAtVAR.get()

    try:
        emptylabel = not labels.any()
    except Exception as ERR:
        pass
    try:
        emptylabel = not labels.any().values[0]
    except AttributeError as ERR:
        pass

    if emptylabel:
        global WIN16
        WIN16 = tk.Toplevel(root)
        WIN16.bind('<Escape>',EXIT_One)
        LstrCateg = ''
        Label_Select(WIN16)
        WIN16.lift()
        WIN16.focus_force()
        WIN16.mainloop()
    if FEAT == 'Just previous Labels':
        msg1 = msg.replace('\n',' ').replace(' with test size','\ntest size').replace('features:','\nfeatures:')
        N = Just_Label_Features(BootstrapLabelWindowDict[Strategy])
        showmessage('Bootstrap for label lag',f'Bootstrap for label lag size for\nmodel {Strategy} is {N} ',TIMEOUT=4000,TYP='info')
        X_Axis_Var = f'Just previous {N} Labels'
                
    elif FEAT == 'Just previous Closes':
        msg1 = msg.replace('\n',' ').replace(' with test size','\ntest size').replace('features:','\nfeatures:')
        N = Just_Close_Features(1)
        X_Axis_Var = f'Just previous {N} Closes'

    elif FEAT == 'OHLC Features':
        Set_OHLC_Features()
        X_Axis_Var = FEAT
      
    elif FEAT == 'All Features':
        All_Features(False)
        X_Axis_Var = FEAT
        
    elif FEAT == 'All features except labels': 
        All_Features_Except_Labels()
        X_Axis_Var = FEAT
    
    elif FEAT == 'Just Binary indicators':
        Just_Binary_indicators()
        X_Axis_Var = FEAT
                
    elif FEAT == 'Selected Binary indicators':
        N = Selected_Binary_indicators()
        X_Axis_Var = f'Just previous {N} Binary indicators'
              
    elif FEAT == 'Just mean Delta':
        N = Just_Mean_Delta()
        X_Axis_Var = f'Just previous {N} Mean_Delta'
        
    elif FEAT == 'Just Mean Return Indicators':
        N = Just_Mean_Return_indicators()
        X_Axis_Var = f'Just previous {N} Mean Return indicators'
        
    elif FEAT == 'Custom Features':
        Custom_Features() 
        X_Axis_Var = str(features.columns.values).replace("'","")
        
    elif FEAT == 'Custom Window Features':
        Custom_Window_Features()

        X_Axis_Var = str(features.columns.values).replace("'","")
        
    else:
        indicator_features()
        X_Axis_Var = FEAT

    if FEAT == 'Custom Features':
        F = str(features.columns.values).replace("'","")
        Fea_MESSAGE1.configure(text = f'{FEAT}\n{F}',fg='#F5CC21')
    elif FEAT != 'Custom Window Features':
        Fea_MESSAGE1.configure(text = f'{FEAT}',fg='#F5CC21')
    

    cmbFeat.set(FEAT)   
    for MoD in MoDeLs[:16]:
        Strategy = MoD.split('→')[1].strip()
        print(f'{Strategy} under test!')
        ModelVAR.set(Strategy)
        MODLBTN.config(text = f'{Strategy} Under Test',fg='yellow',font='normal',bg='mediumvioletred')
        FEAtBTN.config(text = f'{Strategy} Under Test',fg='greenyellow',font='normal',bg='mediumvioletred')
        showmessage('Model option',f'Selected Model option : {Strategy}',TIMEOUT=1000,TYP='info')
        ComparisonFlag = f'Model► {Strategy}'
        RUNStrategy('ModelCMPMode')
        if not EvaluationFault:
            if len(Strategy.split())>2:
                Short_Strategy = ''.join(x[0] for x in Strategy.split())
                CMPTableIndex.append(f'{Short_Strategy} {X_Axis_Var}')
            else:
                CMPTableIndex.append(f'{Strategy} {X_Axis_Var}')                
    Generate_Comparison_Table(CMPTableIndex,X_Axis_Var,'models')
    temp1 = f'Model Compare with feature'
    temp2 = 'models'
    Select_Comparison_Graph(CMPTableIndex,temp1, temp2)
    tempindex = CMPTableIndex

def CompareFeatures():
    global QuickRun
    global Accuracy_lst
    global F1score_lst
    global precision_lst
    global recall_lst
    global Bal_Accuracy_test_lst
    global Logloss_lst
    global ROCAUC_lst
    global R2score_lst
    global MDDP_lst
    global sharpe_ratio_lst
    global TotalNetProfit_lst
    global Profit_ratio_lst
    global CMPTableIndex
    global FEAT
    global Concise_Report
    global tempindex
    global temp1
    global temp2
    global ComparisonFlag
    global comparisondf
    global CompareFeaturessoptions
    
    comparisondf = pd.DataFrame()    
    Accuracy_lst = []
    F1score_lst = []
    precision_lst = []
    recall_lst = []
    Bal_Accuracy_test_lst = []
    Logloss_lst = []
    ROCAUC_lst = []
    R2score_lst = []
    MDDP_lst = []
    TotalNetProfit_lst = []
    Profit_ratio_lst = []
    sharpe_ratio_lst = []
    CMPTableIndex = []
    intChoice5.set(1)
    intChoice4.set(0)
    QuickRun = True
    Concise_Report = True
    intChoice7.set(1)
    chkChoice7.config(text='Concise Report')
    showmessage('Concise reports',f'Concise reports activated',TIMEOUT=2500,TYP='info')
    CMPTableIndex = []
    
    All_Bootstrap_indics()  
    for FEAT in CompareFeaturessoptions: #Fcateg[:-5]
        print('FEAT:',FEAT)
        cmbFeat.set(FEAT)
        #cmbFeat.current(Fcateg.index(FEAT))
        showmessage('features',f'Selected features forced on: {FEAT}',TIMEOUT=1000,TYP='info')
        if FEAT == 'Just previous Labels':
            msg1 = msg.replace('\n',' ').replace(' with test size','\ntest size').replace('features:','\nfeatures:')
            N = Just_Label_Features(0)
            X_Axis_Var = f'Just previous {N} Labels'
            
        elif FEAT == 'Just previous Closes':
            msg1 = msg.replace('\n',' ').replace(' with test size','\ntest size').replace('features:','\nfeatures:')
            N = Just_Close_Features(0)
            X_Axis_Var = f'Just previous {N} Closes'

        elif FEAT == 'OHLC Features':
            Set_OHLC_Features()
            X_Axis_Var = FEAT
       
        elif FEAT == 'All Features':
            All_Features(False)
            X_Axis_Var = FEAT
            
        elif FEAT == 'All features except labels':
            All_Features_Except_Labels()
            X_Axis_Var = FEAT
            
        elif FEAT == 'Custom Features':
            Custom_Features()
            replaceffeatures = str(features.columns.values).replace(' ','_').replace("'",'').replace(')','').replace('(','')#.replace(']','◘').replace('[','◘')
            X_Axis_Var = f'{FEAT}:{replaceffeatures}'            
        elif FEAT == 'Custom Window Features':
            Custom_Window_Features()
            replaceffeatures = str(features.columns.values).replace(' ','_').replace("'",'').replace(')','').replace('(','')#.replace(']','◘').replace('[','◘')
            X_Axis_Var = f'{FEAT}:{replaceffeatures}'   
        
        else:
            #Indic_option(False)
            indicator_features()
            X_Axis_Var = FEAT 
        
        print('FEAT:',FEAT)
        ComparisonFlag = X_Axis_Var  
        RUNStrategy('FeatureCMPMode')
        if len(X_Axis_Var)>30:
            X_Axis_Var = X_Axis_Var[:30]
        if len(Strategy.split())>2:
            Short_Strategy = ''.join(x[0] for x in Strategy.split())
            CMPTableIndex.append(f'{ X_Axis_Var}')
        else:
            CMPTableIndex.append(f'{ X_Axis_Var}') 
    
    Generate_Comparison_Table(CMPTableIndex,Strategy,'features')
    temp1 = f'feature Compare for {Strategy}'
    temp2 = 'features'
    Select_Comparison_Graph(CMPTableIndex,temp1, temp2)
    tempindex = CMPTableIndex
    
def RUNStrategy_(_):
    RUNStrategy('NormalMode')
    
def RUNStrategy(MODE):
    global Strategy
    if MODE == 'NormalMode':
        Strategy = ModelVAR.get()
        VoiceAnnounce(f'{Strategy} Selected')
    elif MODE == 'FeatureCMPMode':
        Strategy = CompareVAR.get()
        
    elif MODE == 'ModelCMPMode':
        VoiceAnnounce(f'{Strategy} Under test')
    
    else:
        showmessage('Unknown operation','Unknown operation',TIMEOUT=3000,TYP='Erro')
        
    root.bind('<Return>', RUNStrategy_)
    if Strategy == 'Gaussian Naïve Bayes':
        Gaussian_NB()
        
    elif Strategy == 'Multi Layer Perceptron Classifier':
        MLPC()
        
    elif Strategy == 'Logistic Regression':
        Log_Reg()
        
    elif Strategy == 'Decision Tree Classifier':
        DTc()

    elif Strategy == 'Random Forest Classifier':
        RFc()

    elif Strategy == 'K Nearest Neighbors Classifier':
        knn()
        
    elif Strategy == 'Support Vector Machine Classifier':
        svc()

    elif Strategy == 'Gradient Boosting Classifier':
        GBC()
        
    elif Strategy == 'X Gradient Boosting Classifier':
        XGBC()

    elif Strategy == 'Light Gradient Boosting Machine Classifier':
        LGBMC()

    elif Strategy == 'CatBoost Classifier':
        CatBC()
        
    elif Strategy == 'AdaBoost Classifier':
        AdaBC() 

    elif Strategy == 'Extra Trees Classifier':
        ExtraTreesC()
        
    elif Strategy == 'Linear Discriminant Analysis':
        LDa()
        
    elif Strategy == 'LSTM Classifier':
        LSTMC()
        
    elif Strategy == 'Deep Neural Network Classifier':
        DeepNN()

    elif Strategy == 'Linear Regressor':
        Lin_Reg()
        
    elif Strategy == 'Decision Tree Regressor':
        DT_Reg()
        
    elif Strategy == 'Random Forest Regressor':
        RF_Reg()
        
    elif Strategy == 'Support Vector Regressor':
        SV_Reg()
        
    elif Strategy == 'Simple Sum Mean Return Binary':
        SSMRB()

    elif Strategy == 'Simple Sum Mean Delta Binary':
        SSMDB()
        
    elif ModelStorm:
        GSESTRATEGY()

def RECORD(duration=2):
    recording = sd.rec(int(duration*freq),samplerate=freq,channels=2)
    sd.wait()
    write("Answers.wav", freq, recording)
    wv.write("Answers1.wav", recording, freq, sampwidth=2)

def stt():
    audio_file = sr.AudioFile('Answers1.wav')
    with audio_file as source:
        r.adjust_for_ambient_noise(source)
        audio = r.record(source)
        try:
            VoiceAnswer = r.recognize_google(audio)
        except:
            print('No response!')
            return 'No response!'#'yes'
        else:
            print(VoiceAnswer)
            return VoiceAnswer

def Toggle_Resultfileopen():
    global BT_text
    global OpenResultBTN
    global notepad_process
    if BT_text == 'Open Results File':
        #os.system(f'notepad.exe {File_path}')
        notepad_process = subprocess.Popen(['notepad.exe', File_path])
        BT_text = 'Close Results File'
    else:
        #os.system('taskkill /f /im notepad.exe')
        notepad_process.terminate()
        BT_text = 'Open Results File'
    OpenResultBTN.config(text = BT_text)
    
#♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦
def featureselect():
    global df
    global timeseries
    global feature    
    try:
        df
    except NameError as ERR:
        showmessage('df error','Dataframe is not available\n'+str(ERR),TIMEOUT=3000,TYP='Erro')
        LoadDataBase('')
    
    def choice():
        global timeseries
        global feature
        feature  = strChoice.get()
        lblSelect.config(text='Your choice is ' + feature)
        if feature == 'Close':
            timeseries = df.CLOSE
        elif feature == 'Open':
            timeseries = df.OPEN
        
    def FeatureSelection(_):
        global timeseries
        print(f'feature:\n{timeseries}')
        WIN9.destroy()
        WIN9.quit()
        
    WIN9 = tk.Toplevel(root)
    WIN11.bind('<Escape>',EXIT_One)
    #WIN9.configure(bg='#091A32')
    pho = tk.PhotoImage(file=r".\pictures\analysis-in-forex-trading.png")
    w = pho.width()
    h = pho.height()
    WIN9.geometry('%dx%d'%(w,h))
    tk.Label(WIN9,image = pho).place(height=h,width=w,x=0,y=0)#.grid(rowspan=8,columnspan=2)
    strChoice = tk.StringVar()
    lblSelect = tk.Label(WIN9,relief=tk.RAISED,font=('times',FONTSIZE2,'bold'))
    lblSelect.grid(row=0,columnspan=2,pady=5)#.place(height=20,width=200,x=50,y=150)

    RadioClose = tk.Radiobutton(WIN9, text='Close',font=('times',FONTSIZE2,'bold'), variable=strChoice, value='Close', command=choice)
    RadioClose.grid(row=1,column=0,pady=5)#.place(height=20,width=250,x=60,y=50)

    RadioOpen = tk.Radiobutton(WIN9, text='Open',font=('times',FONTSIZE2,'bold'), variable=strChoice,value='Open', command=choice)
    RadioOpen.grid(row=1,column=1,pady=5)#.place(height=20,width=250,x=160,y=80)

    strChoice.set("Close")
    choice()
    WIN9.title("Select feature to test")
    btTest = tk.Button(WIN9,text="↑↑↑Select feature to test↑↑↑",bg='#474747',fg='yellow',font=('times',FONTSIZE2,'bold'),relief=tk.RAISED,bd=5,command=lambda:FeatureSelection('a'))
    btTest.grid(row=3,columnspan=2)
    WIN9.bind('<Return>',FeatureSelection)
    WIN9.mainloop()

#♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦♦     
def adf_test():
    from statsmodels.tsa.stattools import adfuller
    global timeseries
    global feature
    featureselect()
    print (f"Results of Dickey-Fuller Test for {str(timeseries)[str(timeseries).find('Name:')+6:str(timeseries).find(',',str(timeseries).find('Name:')+6,)]}:")
    dftest = adfuller(timeseries, autolag='AIC')
    dfoutput = pd.Series(dftest[0:4], index=['Test Statistic','p-value','No. of Lags Used','No. of Observations Used'])
    for key,value in dftest[4].items():
        dfoutput['Critical Value (%s)'%key] = value
    print(dfoutput)
    ADFresult = adfuller(np.log(timeseries))
    print(f'ADF Results for {timeseries}:\n{adfuller(timeseries)}')
    print(f'p-value of {feature}                    :{adfuller(timeseries)[1]}')
    print(f'p-value of {feature} 1ST order differece:{adfuller(timeseries.diff().dropna())[1]}')
    print(f'p-value of {feature} 2ND order differece:{adfuller(timeseries.diff().diff().dropna())[1]}')
    return (dfoutput)

#============================== Kwiatkowski-Phillips-Schmidt-Shin (KPSS) Test
def kpss_test():
    from statsmodels.tsa.stattools import kpss
    global timeseries
    global feature
    featureselect()
    print (f"Results of KPSS Test for {str(timeseries)[str(timeseries).find('Name:')+6:str(timeseries).find(',',str(timeseries).find('Name:')+6,)]}:")
    kpsstest = kpss(timeseries, regression='c', nlags="auto")
    kpss_output = pd.Series(kpsstest[0:3], index=['Test Statistic','p-value','No. of Lags Used'])
    for key,value in kpsstest[3].items():
        kpss_output['Critical Value (%s)'%key] = value
    print(kpss_output)
    return (kpss_output)
#============================== Kwiatkowski-Phillips-Schmidt-Shin (KPSS) Test
def Activate_shell():
    root.quit()

def CatchVariable():
    def DisplayVariable(_):
        nonlocal WIN17
        VarValue = Var_entry.get()
        try:
            VarResults = VarValue + ':\n' + str(eval(VarValue))
        except NameError as ERR:
            showmessage('catch variable Error','Variable is not available to display\n'+str(ERR),TIMEOUT=3000,TYP='Erro')
        else:
            Show_In_Win('catch variable',VarResults,'white','darkgreen')
            print(VarResults)
            
    WIN17 = tk.Toplevel(root)
    WIN17.title("Find value")
    VAR  = tk.StringVar()
    SelVarLb = tk.Label(WIN17,text="select variable ",bg='#474747',fg='white',font=('times',FONTSIZE2,'bold'))
    SelVarLb.grid(row=0,columnspan = 2,pady=5)   #,sticky=tk.W 
    
    VarLB = tk.Label(WIN17,text="Variable: ",bg='#474747',fg='white',font=('times',FONTSIZE2,'bold'))
    VarLB.grid(row=1,column=0,pady=5) #,sticky='W'

    Var_entry = tk.Entry(WIN17,textvariable=VAR,width=20,bg='white',fg='black',font=('times',FONTSIZE2),borderwidth=3,bd=5)
    Var_entry.grid(row=1,column=1,padx=5,pady=5)#,sticky='W'

    btcat = tk.Button(WIN17,text="Catch",bg='#474747',fg='yellow',font=('times',FONTSIZE2,'bold'),relief=tk.RAISED,bd=5,command=lambda:DisplayVariable(''))
    btcat.grid(row=2,columnspan=2)
    WIN17.bind('<Return>',DisplayVariable)
    Var_entry.focus()
    WIN17.mainloop()

def CatchFunction():
    def DisplayFunction(_):
        nonlocal WIN17
        FunValue = Fun_entry.get()
        try:
            FunValue1 = eval(FunValue)
        except NameError as ERR:
            showmessage('catch function Error','Function is not available to display\n'+str(ERR),TIMEOUT=3000,TYP='Erro')
            
    WIN17 = tk.Toplevel(root)
    WIN17.title("function test")
    FUN  = tk.StringVar()
    SelFunLb = tk.Label(WIN17,text="select Function",bg='#474747',fg='white',font=('times',FONTSIZE2,'bold'))
    SelFunLb.grid(row=0,columnspan = 2,pady=5)   #,sticky=tk.W 
    
    FunLB = tk.Label(WIN17,text="Function: ",bg='#474747',fg='white',font=('times',FONTSIZE2,'bold'))
    FunLB.grid(row=1,column=0,pady=5) #,sticky='W'

    Fun_entry = tk.Entry(WIN17,textvariable=FUN,width=20,bg='white',fg='black',font=('times',FONTSIZE2),borderwidth=3,bd=5)
    Fun_entry.grid(row=1,column=1,padx=5,pady=5)#,sticky='W'

    btcat = tk.Button(WIN17,text="Catch",bg='#474747',fg='yellow',font=('times',FONTSIZE2,'bold'),relief=tk.RAISED,bd=5,command=lambda:DisplayFunction(''))
    btcat.grid(row=2,columnspan=2)
    Fun_entry.focus()
    WIN17.bind('<Return>',DisplayFunction)
    WIN17.mainloop()

def Update_features_List(newfeature):
    global Fcateg
    global cmbFeat
    global CompareFeaturessoptions
    global CompareFEAtVAR
    global CompareFEAtopt
    Fcateg.insert(0,newfeature)
    cmbFeat.config(values=Fcateg)
    cmbFeat.set(newfeature)
    CompareFeaturessoptions = Fcateg[:-5]
    CompareFEAtopt['menu'].delete(0, 'end')  
    for option in CompareFeaturessoptions:
        #CompareWInopt['menu'].add_command(label=option, command=tk._setit(CompareWINVAR, CompareWINsoptions))
        CompareFEAtopt['menu'].add_command(label=option, command=tk._setit(CompareFEAtVAR, option))
    CompareFEAtVAR.set(newfeature)        
        
def feature_Select(parent,F):
    global FstrCateg
    global cmbFeat
    global Fcateg
    def feat_Select(event='none'):
        nonlocal parent
        #global FONTSIZE
        global cmbFeat
        global FEAT,FEAT1
        try:
            FEAT = cmbFeat1.get()
        except:
            FEAT = cmbFeat.get()
            
#        root.bind('<Return>', cmbFeatFocus)
        Modelopt.focus_force()
        root.bind('<Return>', modelFocus)
        root.bind('<Up>', modelFocus)
        root.bind('<Down>', modelFocus)
            
        if FEAT == 'OHLC Features':
            Set_OHLC_Features()
        
        elif FEAT == 'Just previous Closes':
            Just_Close_Features(0)
            
        elif FEAT == 'Just previous Labels':
            Just_Label_Features(0)
            
        elif FEAT == 'All Features':
            All_Features(False)
            
        elif FEAT == 'All features except labels':
            All_Features_Except_Labels()
            
        elif FEAT == 'Custom Features':
            Custom_Features()
            
        elif FEAT == 'Custom Window Features':
            Custom_Window_Features()
            
        elif FEAT == 'Add Binary indicators':
            Add_Binary_indicators('Add simply binary indicators to dataframe!')
        
        elif FEAT == 'Just Binary indicators':
            Just_Binary_indicators()
                    
        elif FEAT == 'Selected Binary indicators':
            Selected_Binary_indicators()
            
        elif FEAT == 'Just mean Delta':
            Just_Mean_Delta()
            
        elif FEAT == 'Just Mean Return Indicators':
            Just_Mean_Return_indicators()
            
        else:            
            indicator_features()
         
        if FEAT != 'Custom Features' and FEAT != 'Custom Window Features':
            Fea_MESSAGE = tk.Message(parent,text = f'Features: {FEAT}',justify=tk.LEFT,width=200,bd=5,fg='#F5CC21',bg = '#08162f',font=('Times', int(FONTSIZE*0.8), 'bold'),anchor = 'w',aspect=200)
            Fea_MESSAGE1.configure(text = f'{FEAT}',fg='#F5CC21')
        
        if parent != root:
            Fea_MESSAGE.pack()
        
        cmbFeat.current(Fcateg.index(FEAT))
            
        showmessage('features',f'Selected features: {FEAT}',TIMEOUT=2000,TYP='info')
        root.bind('<Return>', modelFocus)
        root.bind('<Up>', modelFocus)
        root.bind('<Down>', modelFocus)
        if parent != root:
            parent.destroy()
        CompareFEAtVAR.set(FEAT) 
        
    Flbl = tk.Label(parent,text="↓Select Features↓",bd = 8,font=('Tahoma', FONTSIZE, 'bold'), bg='#F5D33B')
    FstrCateg = tk.StringVar()   
    if parent == root:
        Flbl.place(height=30,width=FONTSIZE*34,x=750,y=160)
        cmbFeat = tk.ttk.Combobox(root, textvariable = FstrCateg,values=Fcateg,font=('Tahoma', int(FONTSIZE*1.2), 'bold'))
        cmbFeat.place(height=40,width=FONTSIZE*34,x=750,y=190)
        cmbFeat.bind('<<ComboboxSelected>>', feat_Select)
        cmbFeat.current(0)
    else:
        Flbl.pack(expand=True, fill=tk.BOTH)
        cmbFeat1 = tk.ttk.Combobox(parent, textvariable = FstrCateg,values=Fcateg, font=('Tahoma', int(FONTSIZE*1.2), 'bold'))
        cmbFeat1.pack(expand=True, fill=tk.BOTH)
        cmbFeat1.bind('<<ComboboxSelected>>', feat_Select)
        cmbFeat1.current(0)
#         cmbFeat = cmbFeat1
    
def cmbLabFocus(_):
    global cmbLab
    cmbLab.focus()
    
def cmbFeatFocus(_):
    global cmbFeat
    cmbFeat.focus()
    
def Label_Select(parent):
    global LstrCateg
    global LABL
    global cmbLab,cmbLab1
    def Lab_Select(event='none'):
        global LABL
        nonlocal parent
        global cmbLab,cmbLab1
        try:
            LABL = cmbLab1.get()
            cmbLab.current(Lcateg.index(LABL)) 
        except:
            LABL = cmbLab.get()
            
        Set_Label(LABL)
        LB_MESSAGE = tk.Message(parent,text = f'Selected label: {LABL}',justify=tk.LEFT,width=200,bd=5,fg='#F5CC21',bg = '#08162f',font=('Times', int(FONTSIZE*0.8), 'bold'),anchor = 'w',aspect=200)
        LB_MESSAGE1.configure(text = f'Selected label: {LABL}',fg='#F5CC21')
        if parent != root:
            LB_MESSAGE.pack()
        root.bind('<Return>', cmbFeatFocus)
        cmbFeat.focus()
        showmessage('Label selected',f'Selected label: {LABL}',TIMEOUT=2500,TYP='info')
        root.bind('<Return>', cmbFeatFocus)
        if parent != root:
            parent.destroy()

    Lcateg = ['MARKET','Rise_fall','CLOSE','Scaledreturn']
    Llbl = tk.Label(parent,text="↓Select label↓",bd = 5,font=('Tahoma', FONTSIZE, 'bold'), bg='#F5D33B')       
        
    LstrCateg = tk.StringVar()
    
    if parent == root:
        Llbl.place(height=30,width=FONTSIZE*34,x=750,y=233)
        cmbLab = tk.ttk.Combobox(parent, textvariable = LstrCateg,font=('Tahoma', int(FONTSIZE*1.2), 'bold'))
        cmbLab['values'] = Lcateg
        cmbLab.place(height=40,width=FONTSIZE*34,x=750,y=263)
        cmbLab.bind('<<ComboboxSelected>>', Lab_Select)
        cmbLab.current(0)
    else:
        Llbl.pack(expand=True, fill=tk.BOTH)
        cmbLab1 = tk.ttk.Combobox(parent, textvariable = LstrCateg,font=('Tahoma', int(FONTSIZE*1.2), 'bold'))
        parent.lift()
        parent.focus_force()
        cmbLab1['values'] = Lcateg
        cmbLab1.pack(expand=True, fill=tk.BOTH)
        cmbLab1.bind('<<ComboboxSelected>>', Lab_Select)
        cmbLab1.current(0)
    
def train_test_data_prepar_onebyone(step):
    global labels
    global TestSize
    global Trainsize
    global y_test   
    global features
    global Prepared_features
    global Scaling_Type
    if step == 0:
        try:
            df
        except NameError as ERR:
            showmessage('Dataframe Error',f'Dataframe df is not available\n{ERR}',TIMEOUT=3000,TYP='Erro')
            LoadDataBase('')     
        try:
            if features.empty:
                raise NameError('Features not set yet!')
        except NameError as ERR:
            showmessage('Features Error',f'No feature is available to display!\n{ERR}',TIMEOUT=4000,TYP='Erro')
            WIN15 = tk.Toplevel(root)
            WIN15.bind('<Escape>',EXIT_One)
            WIN15.attributes('-topmost','true')
            FstrCateg = ''
            feature_Select(WIN15,4)
            WIN15.mainloop()

        try:
            emptylabel = not labels.any()
        except Exception as ERR:
            pass
        try:
            emptylabel = not labels.any().values[0]
        except AttributeError as ERR:
            pass

        if emptylabel:
            global WIN16
            WIN16 = tk.Toplevel(root)
            WIN16.bind('<Escape>',EXIT_One)
            LstrCateg = ''
            Label_Select(WIN16)
            WIN16.lift()
            WIN16.focus_force()
            WIN16.mainloop()
        
        try:
            if Prepared_features.empty:
                raise NameError
        except:
            Select_Scaling_Type()
            
    try:
        x_train = Prepared_features[-Trainsize-TestSize:-TestSize+step]
        #x_train = x_train.values
        
        if step+1 == TestSize:
            x_test  = Prepared_features[-1:]
        else:
            x_test  = Prepared_features[-TestSize+step:-TestSize+step+1]
        #x_test  = x_test.values
        
        y_train = labels[-Trainsize-TestSize:-TestSize+step]
        y_test  = labels[-TestSize+step:-TestSize+step+1]
        
        return x_train,y_train,x_test,y_test
    
    except NameError as ERR:
        showmessage('Features Error',f'No prepared feature is available to display\n{ERR}',TIMEOUT=3000,TYP='Erro')

def train_test_data_preparation():
    global labels
    global TestSize
    global Trainsize
    global y_test   
    global features
    global Prepared_features
    global Scaling_Type
    try:
        df
    except NameError as ERR:
        showmessage('Dataframe Error',f'Dataframe df is not available\n{ERR}',TIMEOUT=3000,TYP='Erro')
        LoadDataBase('')   
    try:
        if features.empty:
            raise NameError('Features not set yet!')
    except NameError as ERR:
        showmessage('Features Error',f'No feature is available to display!\n{ERR}',TIMEOUT=4000,TYP='Erro')
        WIN15 = tk.Toplevel(root)
        WIN15.bind('<Escape>',EXIT_One)
        WIN15.attributes('-topmost','true')
        FstrCateg = ''
        feature_Select(WIN15,4)
        WIN15.mainloop()
        #features = features.shift()
        #features.dropna(axis=0,inplace=True)
        #Select_Scaling_Type()
    try:
        emptylabel = not labels.any()
    except Exception as ERR:
        pass
    try:
        emptylabel = not labels.any().values[0]
    except AttributeError as ERR:
        pass

    if emptylabel:
        global WIN16
        WIN16 = tk.Toplevel(root)
        WIN16.bind('<Escape>',EXIT_One)
        LstrCateg = ''
        Label_Select(WIN16)
        WIN16.lift()
        WIN16.focus_force()
        WIN16.mainloop()
    
    try:
        if Prepared_features.empty:
            raise NameError
    except:
        Select_Scaling_Type()
        
    try:
        x_train = Prepared_features[-Trainsize-TestSize:-TestSize]       #x_train = x_train.values
        y_train = labels[-Trainsize-TestSize:-TestSize]
        
        x_test  = Prepared_features[-TestSize:]        #x_test  = x_test.values        
        y_test  = labels[-TestSize:]
        
        return x_train,y_train,x_test,y_test
    
    except NameError as ERR:
        showmessage('Features Error',f'No prepared feature is available to display\n{ERR}',TIMEOUT=3000,TYP='Erro')
                
def Eval_Results(y_train,y_pred_train,y_test,y_pred_test):
    from sklearn.metrics import classification_report , accuracy_score ,recall_score, precision_score,balanced_accuracy_score ,f1_score, roc_auc_score, r2_score ,log_loss 
    from sklearn.metrics import confusion_matrix, ConfusionMatrixDisplay, mean_absolute_error, mean_squared_error
    global msg
    global Strategy
    global prediction_data2
    global EvaluationFault
    EvaluationFault = False
#     global y_test
#     global y_pred_test
    prediction_data1 = f'y_train with shape {y_train.shape}     =\n{y_train}\n\ny_pred_train with shape {y_pred_train.shape}  =\n{y_pred_train}\n\ny_test with shape {y_test.shape}        =\n{y_test}\n\ny_pred_test with shape {y_pred_test.shape}    =\n{y_pred_test}'
    Process_Summarize(False)
    print(customized.color.green+msg+'\n'+prediction_data1+customized.color.normal)
    
    if not Concise_Report:
        Show_In_Win('Prediction_data1',msg+'\n'+prediction_data1,'snow','indigo')
    try:
        Acc_test = accuracy_score(y_test, y_pred_test)
        Accuracy_lst.append(round(Acc_test,2))
        
        Bal_Accuracy_test = balanced_accuracy_score(y_test, y_pred_test)
        Bal_Accuracy_test_lst.append(round(Bal_Accuracy_test,2))
        
        recall = recall_score(y_test, y_pred_test)
        recall_lst.append(round(recall,2))
        
        precision = precision_score(y_test, y_pred_test)
        precision_lst.append(round(precision,2))
        
        F1score = f1_score(y_test, y_pred_test,average='weighted')
        F1score_lst.append(round(F1score,2))
        
        R2score = r2_score(y_test, y_pred_test)
        R2score_lst.append(round(R2score,2))
        
        ROCAUC  = roc_auc_score(y_test, y_pred_test)
        ROCAUC_lst.append(round(ROCAUC,2))
        
        Logloss = log_loss(y_test, y_pred_test)
        Logloss_lst.append(round(Logloss,2))
        
        prediction_data2 = f'''
↓classification report for test part↓
{classification_report(y_test, y_pred_test)}
••••••••••••••••••••••••••••••••••••••••••••••
train samples Class 0\t\t\t:{(100*np.count_nonzero(y_train == 0))/len(y_train):.2f} %
train samples Class 1\t\t\t:{(100*np.count_nonzero(y_train == 1))/len(y_train):.2f} % 
test  samples Class 0\t\t\t:{(100*np.count_nonzero(y_test == 0))/len(y_test):.2f} %
test  samples Class 1\t\t\t:{(100*np.count_nonzero(y_test == 1))/len(y_test):.2f} % 
•••↓All for test part↓••••••••••••••••••••••••
Accuracy score\t\t\t:{Acc_test:.4f}
Bal_Accuracy score\t\t\t:{Bal_Accuracy_test:.4f}
Recall score\t\t\t:{recall:.4f}
Precision score\t\t\t:{precision:.4f}
F1_score\t\t\t:{F1score:.4f}
R2_score\t\t\t:{R2score:.4f}
ROC_AUC_score\t\t\t:{ROCAUC:.4f}
LOGLOSS Value\t\t\t:{Logloss:.4f}
'''
        print(customized.color.bold+msg+'\n'+prediction_data2+customized.color.normal)
        print(customized.color.boldNavyBlue+f'TimeFrame:{TimeFrame}'+customized.color.normal)
        #Show_In_Win('prediction_data2',msg+'\n'+prediction_data2,'blue','yellow')
    except ValueError as ERR:
        EvaluationFault = True
        showmessage('Evaluation method error',f'Classification metrics doesn\'t mean on {Strategy}\n{ERR}',TIMEOUT=3000,TYP='Erro')
    RUNBTN.configure(text=f"RUN {Strategy}\n with test size:{TestSize} and train size:{Trainsize}\nTimeFrame:{TimeFrame}", fg='chartreuse',font=('Times', FONTSIZE, 'normal'))

   
def Calculate_Equity(y_train,y_pred_train,y_test,y_pred_test):
    global MDD
    global MDDP
    global df
    global comparisondf
    global File_path
    df['predictedclass']=pd.DataFrame(y_pred_test,index=df.iloc[-len(y_pred_test):].index)
    df['Compare']  = df['LABELS'] == df['predictedclass']
    df['Win/Loss']=df['Compare'].apply(lambda x: 1 if x else 0)   
    LastPeak = BALANCE
    LastTrough = BALANCE+1000000
    commission = 0.0
    
    def sharpe_ratio( ):
        average_return = np.mean(df['profit'])#[-TestSize:])   
        std_dev = np.std(df['profit'])#[-TestSize:])   
        risk_free_rate = 0.005  
        sharpe= (average_return - risk_free_rate) / std_dev
        print(f'Sharpe Ratio: {sharpe:.2f}')
        return sharpe
    
    def calc_profit(row):
        if row['Compare']:
            return abs(row['Return'])*row['Lottage']*100000 - commission
        else:
            return -abs(row['Return'])*row['Lottage']*100000 - commission
        
    df['profit'] = df.apply(lambda row: calc_profit(row), axis=1)
    #df['profit'][:-len(y_test)] = apply(0)
    try:
        df.loc[df.index[0]:df.index[-len(y_test)],'profit'] = 0
    except:
        df.iloc[0:-len(y_test),df.columns.get_loc('profit')] = 0

    df['equity'] = df['profit'].cumsum() + BALANCE
    TestPartProfit = df['equity'].iloc[-1] - df['equity'].iloc[-len(y_test)-1] # it is possible too → {df.tail(1)["equity"].values[0]- BALANCE:.2f} $

    if ComparisonFlag:
        comparisondf[ComparisonFlag] = df['profit'].cumsum() + BALANCE
        #comparisondf = pd.concat([comparisondf,df['equity']],axis = 1)
    
    TP=0 ;   TN=0 ;    FP=0 ;    FN=0
    for m in df[-TestSize:].iterrows():
        if m[1]['LABELS']==1 and m[1]['predictedclass']==1:
            TP += 1
        elif m[1]['LABELS']==0 and m[1]['predictedclass']==0:
            TN+=1
        elif m[1]['LABELS']==0 and m[1]['predictedclass']==1:
            FP+=1
        elif m[1]['LABELS']==1 and m[1]['predictedclass']==0:
            FN+=1    
    try:
        Accuracy_test = (TP+TN)/(TP+TN+FP+FN)
    except ZeroDivisionError:
        Accuracy_test = 0
    
    try:
        Precision = TP/(TP+FP)
    except ZeroDivisionError:
        Precision = 0
    
    try:
        Specificity = TN/(TN+FP)
    except ZeroDivisionError:
        Specificity = 0
        
    try:
        Recall = TP/(TP+FN)
    except ZeroDivisionError:
        Recall = 0
    
    try:
        Bal_Accuracy_test = 0.5*(TN/(TN+FP)+TP/(TP+FN))
    except ZeroDivisionError:
        Bal_Accuracy_test = 0
       
    try:
        F_measure = 2*Precision*Recall/(Precision+Recall)
    except ZeroDivisionError:
        F_measure = 0
        
    n_win_trades  = df[-len(y_test):][df['profit']>=0.0]['profit'].count()
    n_lost_trades = df[-len(y_test):][df['profit']<0.0]['profit'].count()
    #Account_Balance   += [df.tail(1)['equity'].values]
    Account_Balance = df.tail(1)['equity'].values

    equities = df['equity'].values
    A    = []
    MDD  = []
    MDDP = []
    for j in range (len(equities)):
        A.append(equities[j] - LastPeak)
        if equities[j] > LastPeak:
            LastPeak = equities[j]
            
    df['DrawDown']=pd.DataFrame(A, index=df.index)
    MDD  = min(A)
    MDDP = 100*abs(df['DrawDown'].min())/LastPeak

    FinalEvaluationText = msg +prediction_data2+ f"""
•••••••••••••••••••••••••••••••••••••••••••••••
Total Samples\t\t\t: {TP+TN+FP+FN}
TP\t\t\t: {TP}
TN\t\t\t: {TN}
FP\t\t\t: {FP}
FN\t\t\t: {FN}
Test Accuracy\t\t\t: {Accuracy_test:.4f}
Test Balanced Accuracy\t\t\t: {Bal_Accuracy_test:.4f}
Test Precision\t\t\t: {Precision:.4f}
Test Specificity\t\t\t: {Specificity:.4f}
Test Recall\t\t\t: {Recall:.4f}
•••••••••••••••••••••••••••••••••••••••••••••••
Test F_measure\t\t\t: {F_measure:.4f}
Account Balance\t\t\t: {Account_Balance[0]:.2f} $
initial deposit\t\t\t: {BALANCE}
Test win count\t\t\t: {np.count_nonzero(df['Compare'][-len(y_test):]== True)}
Number of won Trades\t\t\t: {n_win_trades}
Test loss count\t\t\t: {np.count_nonzero(df['Compare'][-len(y_test):]== False)}
Number of Lost Trades\t\t\t: {n_lost_trades}
Test win percent\t\t\t: {100*np.count_nonzero(df['Compare'][-len(y_test):])/len(y_test):.2f} %
Hit Ratio\t\t\t: {100*n_win_trades/(n_win_trades + n_lost_trades):.2f} %
Test part Net Profit\t\t\t: {TestPartProfit:.2f}$
Accumulated Return\t\t\t: {(Account_Balance[0]-BALANCE)/BALANCE:.2f} %
Won Trades Average\t\t\t: {df[df['profit']>0.0]['profit'].mean():.2f}$
Lost Trades Average\t\t\t: {df[df['profit']<0.0]['profit'].mean():.2f} $
Largest Won Trade\t\t\t: {df[df['profit']>0.0]['profit'].max():.2f} $
Largest Lost Trade\t\t\t: {df[df['profit']<0.0]['profit'].min():.2f}$
Profit ratio\t\t\t: {abs(df[df['profit']>=0.0]['profit'].sum()/df[df['profit']<0.0]['profit'].sum()):.4f}%
•••••••••••••••••••••••••••••••••••••••••••••••
Max. DrawDown\t\t\t: {MDD:.2f} $
Max. DrawDown percent\t\t\t: {MDDP:.1f} %
sharpe ratio\t\t\t: {sharpe_ratio()}
"""
    if not EvaluationFault:
        TotalNetProfit_lst.append(round(df.tail(1)["equity"].values[0]- BALANCE,2))
        Profit_ratio_lst.append(-round(df[df['profit']>0.0]['profit'].sum()/df[df['profit']<0.0]['profit'].sum(),2))
        MDDP_lst.append(round(MDDP,2))
        sharpe_ratio_lst.append(round(sharpe_ratio(),2))
    
    df['DrawDown']= df['DrawDown'].apply(lambda x: 0 if x>0 else x)
    STRINg = str(features.columns.values).replace("'","")
#     if len(Strategy.split())>2:
#         Short_Strategy = ''.join(x[0] for x in Strategy.split())
#         Show_In_Win(f'Final Evaluation_{Short_Strategy}_{FEAT}►{STRINg}',FinalEvaluationText,'white','black')
#     else:
    Show_In_Win(f'Final Evaluation_{Strategy}_{FEAT}►{STRINg}',FinalEvaluationText,'white','black')
    
    FileName = ''
    if len(Strategy.split())>2:
        Short_Strategy = ''.join(x[0] for x in Strategy.split())
    else:
        Short_Strategy = Strategy
        
    File_path = os.getcwd()+f'\\Simulation Results\\{Strategy}'       
    if not os.path.exists(File_path):
        showmessage('Finding path',f'The path {File_path} was not found!\nIt will be created!',TIMEOUT= 3000,TYP='warn')
        try:
            os.makedirs(File_path)
        except Exception as ERR:
            showmessage('Path Error',f'Check the file path or name: {File_path}\n{ERR}',TIMEOUT=3000,TYP='Erro')
            #File_path = File_path.replace(' ','_').replace("'",'').replace(')','').replace('(','').replace(']','▬').replace('[','▬')
            File_path = filedialog.asksaveasfilename(initialdir = File_path ,defaultextension=".txt", filetypes=[("text file", "*.txt"),  ("All files", "*.*")],initialfile= File_path[:10])
#             if not os.path.exists(File_path):
#                 showmessage('Finding path',f'The path {File_path} was not found!\nIt will be created!',TIMEOUT= 3000,TYP='warn')
#                 os.makedirs(File_path)

    FileName = 'Final Results '+Short_Strategy+str(features.columns.values).replace(' ','_').replace("'",'').replace(')','').replace('(','').replace(']','▬').replace('[','▬')
    if ComparisonFlag or FEAT =='Just previous Closes' or FEAT == 'Just previous Labels' or FEAT == 'Custom Window Features':
        if len(FileName.split('_'))>2:
            FileName = FileName.split('_')[0]+'→'+FileName.split('_')[-1]
    File_path += '\\'+FileName+'.txt'
    with open(File_path,'w', encoding = 'utf-8') as ResultFile:
        try:
            print(FinalEvaluationText,file=ResultFile)
        except UnicodeEncodeError as ERR:
            print(f'UnicodeEncodeError: {ERR}')
    global OpenResultBTN
    OpenResultBTN = tk.Button(root,text=BT_text,font=('Tahoma', FONTSIZE, 'bold'),bd=8,fg='#F5D33B', bg='#05051E',activebackground='gold',activeforeground='navy blue',command = Toggle_Resultfileopen)
    OpenResultBTN.place(height=FONTSIZE*4,x=740,y=505)
                 
    print(customized.color.GhighlitedW+FinalEvaluationText+customized.color.normal)
    root.bind('<Return>', modelFocus)
    
def Gaussian_NB():
    from sklearn.naive_bayes import GaussianNB
    global Strategy
    global y_test
    global y_pred_test
    global Y_PRED_Storm_TEST
    global Scaling_Type
    Y_PRED_Storm_TEST  = np.array([],dtype=int)
    Strategy = 'Gaussian Naïve Bayes'
    MODL = GaussianNB()
    if SingleCandlePredict:
        for step in range(TestSize):
            x_train,y_train,x_test,y_test = train_test_data_prepar_onebyone(step)   
            train_test_data = f'x_train.class  ={x_train.__class__}\nx_train.shape  ={x_train.shape}\nx_train  =\n{x_train}\n\ny_train.class  ={y_train.__class__}\ny_train.shape  ={y_train.shape}\ny_train  =\n{y_train}\n\nx_test.class   ={x_test.__class__}\nx_test.shape   ={x_test.shape}\nx_test   =\n{x_test}\n\ny_test.shape   ={y_test.shape}\ny_test   =\n{y_test}'   
            MODL.fit(x_train, y_train)    
            ALARM.Alarm('Beep')
            y_pred_train = MODL.predict(x_train)
            y_pred_test  = MODL.predict(x_test)
            if step == 0:
                Y_PRED_TRAIN = y_pred_train
            Y_PRED_Storm_TEST = np.concatenate((Y_PRED_Storm_TEST,y_pred_test),axis=0)
            print('Y_PRED_Storm_TEST',Y_PRED_Storm_TEST)
        if not Concise_Report:
            Show_In_Win('train_test_data',train_test_data,'wheat','darkolivegreen')     
        pred=list(Y_PRED_TRAIN) + list(Y_PRED_Storm_TEST)
        df['predicted']=pd.DataFrame(pred,index=df.iloc[-len(pred):].index)
              
        y_pred_test = Y_PRED_Storm_TEST
        y_test  = labels[-TestSize:]  
    else:
        x_train,y_train,x_test,y_test = train_test_data_preparation()    
        train_test_data = f'x_train.shape  ={x_train.shape}\nx_train  =\n{x_train}\n\ny_train.shape  ={y_train.shape}\ny_train  =\n{y_train}\n\nx_test.shape   ={x_test.shape}\nx_test   =\n{x_test}\n\ny_test.shape   ={y_test.shape}\ny_test   =\n{y_test}'
        print(train_test_data)
        if not Concise_Report:
            Show_In_Win('train_test_data',train_test_data,'seashell','saddlebrown') 
        MODL.fit(x_train, y_train)
            
        ALARM.Alarm('Beep')
        y_pred_train = MODL.predict(x_train)
        y_pred_test  = MODL.predict(x_test)

    predTXT = f'''
y_pred_train:\n{y_pred_train}\n
y_pred_test:\n{y_pred_test}'''
    #print(predTXT)
    if not Concise_Report:
        Show_In_Win('train_test prediction',predTXT,'aliceblue','steelblue')
    Eval_Results(y_train,y_pred_train,y_test,y_pred_test)
    Calculate_Equity(y_train,y_pred_train,y_test,y_pred_test)
   
def MLPC():
    from sklearn.neural_network import MLPClassifier
    global Strategy
    global y_test
    global y_pred_test
    global Y_PRED_Storm_TEST; Y_PRED_Storm_TEST  = np.array([],dtype=int)
    global Scaling_Type
    Strategy = 'Multi Layer Perceptron Classifier'
    MODL = MLPClassifier(activation='relu',max_iter=10000,hidden_layer_sizes=(13,),solver='adam', random_state=2)
    if SingleCandlePredict:
        for step in range(TestSize):
            x_train,y_train,x_test,y_test = train_test_data_prepar_onebyone(step)     
            train_test_data = f'x_train.class  ={x_train.__class__}\nx_train.shape  ={x_train.shape}\nx_train  =\n{x_train}\n\ny_train.class  ={y_train.__class__}\ny_train.shape  ={y_train.shape}\ny_train  =\n{y_train}\n\nx_test.class   ={x_test.__class__}\nx_test.shape   ={x_test.shape}\nx_test   =\n{x_test}\n\ny_test.shape   ={y_test.shape}\ny_test   =\n{y_test}'
            print(train_test_data)
            
            MODL.fit(x_train, y_train)    
            ALARM.Alarm('Beep')
            y_pred_train = MODL.predict(x_train)
            y_pred_test  = MODL.predict(x_test)
            if step == 0:
                Y_PRED_TRAIN = y_pred_train
            Y_PRED_Storm_TEST      = np.concatenate((Y_PRED_Storm_TEST,y_pred_test),axis=0)
            print('Y_PRED_Storm_TEST',Y_PRED_Storm_TEST)
        if not Concise_Report:
            Show_In_Win('train_test_data',train_test_data)     
        pred=list(Y_PRED_TRAIN) + list(Y_PRED_Storm_TEST)
        df['predicted']=pd.DataFrame(pred,index=df.iloc[-len(pred):].index)
              
        y_pred_test = Y_PRED_Storm_TEST
        y_test  = labels[-TestSize:]  
    else:
        x_train,y_train,x_test,y_test = train_test_data_preparation()    
        train_test_data = f'x_train.shape  ={x_train.shape}\nx_train  =\n{x_train}\n\ny_train.shape  ={y_train.shape}\ny_train  =\n{y_train}\n\nx_test.shape   ={x_test.shape}\nx_test   =\n{x_test}\n\ny_test.shape   ={y_test.shape}\ny_test   =\n{y_test}'
        print(train_test_data)
        if not Concise_Report:
            Show_In_Win('train_test_data',train_test_data) 
        MODL.fit(x_train, y_train)
            
        ALARM.Alarm('Beep')
        y_pred_train = MODL.predict(x_train)
        y_pred_test  = MODL.predict(x_test)

    predTXT = f'''
y_pred_train:\n{y_pred_train}\n
y_pred_test:\n{y_pred_test}'''
    #print(predTXT)
    if not Concise_Report:
        Show_In_Win('train_test prediction',predTXT)
    Eval_Results(y_train,y_pred_train,y_test,y_pred_test)
    Calculate_Equity(y_train,y_pred_train,y_test,y_pred_test)

def Log_Reg():
    from sklearn.linear_model import LogisticRegression
    global Strategy
    global y_test
    global y_pred_test
    global Y_PRED_Storm_TEST; Y_PRED_Storm_TEST  = np.array([],dtype=int)
    global Scaling_Type
    Strategy = 'Logistic Regression'
    MODL = LogisticRegression(warm_start=False,random_state=0)
    if SingleCandlePredict:
        for step in range(TestSize):
            x_train,y_train,x_test,y_test = train_test_data_prepar_onebyone(step)     
            train_test_data = f'x_train.class  ={x_train.__class__}\nx_train.shape  ={x_train.shape}\nx_train  =\n{x_train}\n\ny_train.class  ={y_train.__class__}\ny_train.shape  ={y_train.shape}\ny_train  =\n{y_train}\n\nx_test.class   ={x_test.__class__}\nx_test.shape   ={x_test.shape}\nx_test   =\n{x_test}\n\ny_test.shape   ={y_test.shape}\ny_test   =\n{y_test}'
            print(train_test_data)
            
            MODL.fit(x_train, y_train)    
            ALARM.Alarm('Beep')
            y_pred_train = MODL.predict(x_train)
            y_pred_test  = MODL.predict(x_test)
            if step == 0:
                Y_PRED_TRAIN = y_pred_train
            Y_PRED_Storm_TEST      = np.concatenate((Y_PRED_Storm_TEST,y_pred_test),axis=0)
            print('Y_PRED_Storm_TEST',Y_PRED_Storm_TEST)
        if not Concise_Report:
            Show_In_Win('train_test_data',train_test_data)     
        pred=list(Y_PRED_TRAIN) + list(Y_PRED_Storm_TEST)
        print('len(Y_PRED_TRAIN):\n',len(Y_PRED_TRAIN))
        print('Y_PRED_TRAIN:\n',Y_PRED_TRAIN)
        print('len(Y_PRED_Storm_TEST):\n',len(Y_PRED_Storm_TEST))
        print('Y_PRED_Storm_TEST:\n',Y_PRED_Storm_TEST)
        df['predicted']=pd.DataFrame(pred,index=df.iloc[-len(pred):].index)
              
        y_pred_test = Y_PRED_Storm_TEST
        y_test  = labels[-TestSize:]   
    else:
        x_train,y_train,x_test,y_test = train_test_data_preparation()    
        train_test_data = f'x_train.shape  ={x_train.shape}\nx_train  =\n{x_train}\n\ny_train.shape  ={y_train.shape}\ny_train  =\n{y_train}\n\nx_test.shape   ={x_test.shape}\nx_test   =\n{x_test}\n\ny_test.shape   ={y_test.shape}\ny_test   =\n{y_test}'
        print(train_test_data)
        if not Concise_Report:
            Show_In_Win('train_test_data',train_test_data) 
        MODL.fit(x_train, y_train)
            
        ALARM.Alarm('Beep')
        y_pred_train = MODL.predict(x_train)
        y_pred_test  = MODL.predict(x_test)

    predTXT = f'''
y_pred_train:\n{y_pred_train}\n
y_pred_test:\n{y_pred_test}'''
    #print(predTXT)
    if not Concise_Report:
        Show_In_Win('train_test prediction',predTXT)
    Eval_Results(y_train,y_pred_train,y_test,y_pred_test)
    Calculate_Equity(y_train,y_pred_train,y_test,y_pred_test)

def DTc():
    from sklearn.tree import DecisionTreeClassifier
    global Strategy
    global y_test
    global y_pred_test
    global Y_PRED_Storm_TEST; Y_PRED_Storm_TEST  = np.array([],dtype=int)
    global Scaling_Type
    Strategy = 'Decision Tree Classifier'
    MODL = DecisionTreeClassifier()
    if SingleCandlePredict:
        for step in range(TestSize):
            x_train,y_train,x_test,y_test = train_test_data_prepar_onebyone(step)     
            train_test_data = f'x_train.class  ={x_train.__class__}\nx_train.shape  ={x_train.shape}\nx_train  =\n{x_train}\n\ny_train.class  ={y_train.__class__}\ny_train.shape  ={y_train.shape}\ny_train  =\n{y_train}\n\nx_test.class   ={x_test.__class__}\nx_test.shape   ={x_test.shape}\nx_test   =\n{x_test}\n\ny_test.shape   ={y_test.shape}\ny_test   =\n{y_test}'
            print(train_test_data)
            
            MODL.fit(x_train, y_train)    
            ALARM.Alarm('Beep')
            y_pred_train = MODL.predict(x_train)
            y_pred_test  = MODL.predict(x_test)
            if step == 0:
                Y_PRED_TRAIN = y_pred_train
            Y_PRED_Storm_TEST      = np.concatenate((Y_PRED_Storm_TEST,y_pred_test),axis=0)
            print('Y_PRED_Storm_TEST',Y_PRED_Storm_TEST)
        if not Concise_Report:
            Show_In_Win('train_test_data',train_test_data)     
        pred=list(Y_PRED_TRAIN) + list(Y_PRED_Storm_TEST)
        df['predicted']=pd.DataFrame(pred,index=df.iloc[-len(pred):].index)
              
        y_pred_test = Y_PRED_Storm_TEST
        y_test  = labels[-TestSize:]   
    else:
        x_train,y_train,x_test,y_test = train_test_data_preparation()    
        train_test_data = f'x_train.shape  ={x_train.shape}\nx_train  =\n{x_train}\n\ny_train.shape  ={y_train.shape}\ny_train  =\n{y_train}\n\nx_test.shape   ={x_test.shape}\nx_test   =\n{x_test}\n\ny_test.shape   ={y_test.shape}\ny_test   =\n{y_test}'
        print(train_test_data)
        if not Concise_Report:
            Show_In_Win('train_test_data',train_test_data) 
        MODL.fit(x_train, y_train)
            
        ALARM.Alarm('Beep')
        y_pred_train = MODL.predict(x_train)
        y_pred_test  = MODL.predict(x_test)

    predTXT = f'''
y_pred_train:\n{y_pred_train}\n
y_pred_test:\n{y_pred_test}'''
    #print(predTXT)
    if not Concise_Report:
        Show_In_Win('train_test prediction',predTXT)
    Eval_Results(y_train,y_pred_train,y_test,y_pred_test)
    Calculate_Equity(y_train,y_pred_train,y_test,y_pred_test)

def RFc():
    from sklearn.ensemble import RandomForestClassifier   
    global Strategy
    global y_test
    global y_pred_test
    global Y_PRED_Storm_TEST; Y_PRED_Storm_TEST  = np.array([],dtype=int)
    global Scaling_Type
    Strategy = 'Random Forest Classifier'
    MODL = RandomForestClassifier(n_estimators=200,max_features=9,random_state = 0)
    if SingleCandlePredict:
        for step in range(TestSize):
            x_train,y_train,x_test,y_test = train_test_data_prepar_onebyone(step)     
            train_test_data = f'x_train.class  ={x_train.__class__}\nx_train.shape  ={x_train.shape}\nx_train  =\n{x_train}\n\ny_train.class  ={y_train.__class__}\ny_train.shape  ={y_train.shape}\ny_train  =\n{y_train}\n\nx_test.class   ={x_test.__class__}\nx_test.shape   ={x_test.shape}\nx_test   =\n{x_test}\n\ny_test.shape   ={y_test.shape}\ny_test   =\n{y_test}'
            print(train_test_data)
            
            MODL.fit(x_train, y_train)    
            ALARM.Alarm('Beep')
            y_pred_train = MODL.predict(x_train)
            y_pred_test  = MODL.predict(x_test)
            if step == 0:
                Y_PRED_TRAIN = y_pred_train
            Y_PRED_Storm_TEST      = np.concatenate((Y_PRED_Storm_TEST,y_pred_test),axis=0)
            print('Y_PRED_Storm_TEST',Y_PRED_Storm_TEST)
        if not Concise_Report:
            Show_In_Win('train_test_data',train_test_data)     
        pred=list(Y_PRED_TRAIN) + list(Y_PRED_Storm_TEST)
        df['predicted']=pd.DataFrame(pred,index=df.iloc[-len(pred):].index)
              
        y_pred_test = Y_PRED_Storm_TEST
        y_test  = labels[-TestSize:]      
    else:
        x_train,y_train,x_test,y_test = train_test_data_preparation()    
        train_test_data = f'x_train.shape  ={x_train.shape}\nx_train  =\n{x_train}\n\ny_train.shape  ={y_train.shape}\ny_train  =\n{y_train}\n\nx_test.shape   ={x_test.shape}\nx_test   =\n{x_test}\n\ny_test.shape   ={y_test.shape}\ny_test   =\n{y_test}'
        print(train_test_data)
        if not Concise_Report:
            Show_In_Win('train_test_data',train_test_data) 
        MODL.fit(x_train, y_train)
            
        ALARM.Alarm('Beep')
        y_pred_train = MODL.predict(x_train)
        y_pred_test  = MODL.predict(x_test)

    predTXT = f'''
y_pred_train:\n{y_pred_train}\n
y_pred_test:\n{y_pred_test}'''
    #print(predTXT)
    if not Concise_Report:
        Show_In_Win('train_test prediction',predTXT)
    Eval_Results(y_train,y_pred_train,y_test,y_pred_test)
    Calculate_Equity(y_train,y_pred_train,y_test,y_pred_test)

def knn():
    from sklearn.neighbors import KNeighborsClassifier     
    global Strategy
    global y_test
    global y_pred_test
    global Y_PRED_Storm_TEST; Y_PRED_Storm_TEST  = np.array([],dtype=int)
    global Scaling_Type
    Strategy = 'K Nearest Neighbors Classifier'
    MODL = KNeighborsClassifier()
    if SingleCandlePredict:
        for step in range(TestSize):
            x_train,y_train,x_test,y_test = train_test_data_prepar_onebyone(step)     
            train_test_data = f'x_train.class  ={x_train.__class__}\nx_train.shape  ={x_train.shape}\nx_train  =\n{x_train}\n\ny_train.class  ={y_train.__class__}\ny_train.shape  ={y_train.shape}\ny_train  =\n{y_train}\n\nx_test.class   ={x_test.__class__}\nx_test.shape   ={x_test.shape}\nx_test   =\n{x_test}\n\ny_test.shape   ={y_test.shape}\ny_test   =\n{y_test}'
            print(train_test_data)
            
            MODL.fit(x_train, y_train)    
            ALARM.Alarm('Beep')
            y_pred_train = MODL.predict(x_train)
            y_pred_test  = MODL.predict(x_test)
            if step == 0:
                Y_PRED_TRAIN = y_pred_train
            Y_PRED_Storm_TEST      = np.concatenate((Y_PRED_Storm_TEST,y_pred_test),axis=0)
            print('Y_PRED_Storm_TEST',Y_PRED_Storm_TEST)
        if not Concise_Report:
            Show_In_Win('train_test_data',train_test_data)     
        pred=list(Y_PRED_TRAIN) + list(Y_PRED_Storm_TEST)
        df['predicted']=pd.DataFrame(pred,index=df.iloc[-len(pred):].index)
              
        y_pred_test = Y_PRED_Storm_TEST
        y_test  = labels[-TestSize:]    
    else:
        x_train,y_train,x_test,y_test = train_test_data_preparation()    
        train_test_data = f'x_train.shape  ={x_train.shape}\nx_train  =\n{x_train}\n\ny_train.shape  ={y_train.shape}\ny_train  =\n{y_train}\n\nx_test.shape   ={x_test.shape}\nx_test   =\n{x_test}\n\ny_test.shape   ={y_test.shape}\ny_test   =\n{y_test}'
        print(train_test_data)
        if not Concise_Report:
            Show_In_Win('train_test_data',train_test_data) 
        MODL.fit(x_train, y_train)
            
        ALARM.Alarm('Beep')
        y_pred_train = MODL.predict(x_train)
        y_pred_test  = MODL.predict(x_test)

    predTXT = f'''
y_pred_train:\n{y_pred_train}\n
y_pred_test:\n{y_pred_test}'''
    #print(predTXT)
    if not Concise_Report:
        Show_In_Win('train_test prediction',predTXT)
    Eval_Results(y_train,y_pred_train,y_test,y_pred_test)
    Calculate_Equity(y_train,y_pred_train,y_test,y_pred_test)

def svc():
    from sklearn.svm import SVC 
    global Strategy
    global y_test
    global y_pred_test
    global Y_PRED_Storm_TEST; Y_PRED_Storm_TEST  = np.array([],dtype=int)
    global Scaling_Type
    Strategy = 'Support Vector Machine Classifier'
    MODL = SVC(kernel='rbf', gamma=10, C = 1)
    if SingleCandlePredict:
        for step in range(TestSize):
            x_train,y_train,x_test,y_test = train_test_data_prepar_onebyone(step)     
            train_test_data = f'x_train.class  ={x_train.__class__}\nx_train.shape  ={x_train.shape}\nx_train  =\n{x_train}\n\ny_train.class  ={y_train.__class__}\ny_train.shape  ={y_train.shape}\ny_train  =\n{y_train}\n\nx_test.class   ={x_test.__class__}\nx_test.shape   ={x_test.shape}\nx_test   =\n{x_test}\n\ny_test.shape   ={y_test.shape}\ny_test   =\n{y_test}'
            print(train_test_data)
            
            MODL.fit(x_train, y_train)    
            ALARM.Alarm('Beep')
            y_pred_train = MODL.predict(x_train)
            y_pred_test  = MODL.predict(x_test)
            if step == 0:
                Y_PRED_TRAIN = y_pred_train
            Y_PRED_Storm_TEST      = np.concatenate((Y_PRED_Storm_TEST,y_pred_test),axis=0)
            print('Y_PRED_Storm_TEST',Y_PRED_Storm_TEST)
        if not Concise_Report:
            Show_In_Win('train_test_data',train_test_data)     
        pred=list(Y_PRED_TRAIN) + list(Y_PRED_Storm_TEST)
        df['predicted']=pd.DataFrame(pred,index=df.iloc[-len(pred):].index)
              
        y_pred_test = Y_PRED_Storm_TEST
        y_test  = labels[-TestSize:]     
    else:
        x_train,y_train,x_test,y_test = train_test_data_preparation()    
        train_test_data = f'x_train.shape  ={x_train.shape}\nx_train  =\n{x_train}\n\ny_train.shape  ={y_train.shape}\ny_train  =\n{y_train}\n\nx_test.shape   ={x_test.shape}\nx_test   =\n{x_test}\n\ny_test.shape   ={y_test.shape}\ny_test   =\n{y_test}'
        print(train_test_data)
        if not Concise_Report:
            Show_In_Win('train_test_data',train_test_data) 
        MODL.fit(x_train, y_train)
            
        ALARM.Alarm('Beep')
        y_pred_train = MODL.predict(x_train)
        y_pred_test  = MODL.predict(x_test)

    predTXT = f'''
y_pred_train:\n{y_pred_train}\n
y_pred_test:\n{y_pred_test}'''
    #print(predTXT)
    if not Concise_Report:
        Show_In_Win('train_test prediction',predTXT)
    Eval_Results(y_train,y_pred_train,y_test,y_pred_test)
    Calculate_Equity(y_train,y_pred_train,y_test,y_pred_test)

def GBC():
    from sklearn.ensemble import GradientBoostingClassifier
    global Strategy
    global y_test
    global y_pred_test
    global Y_PRED_Storm_TEST; Y_PRED_Storm_TEST  = np.array([],dtype=int)
    global Scaling_Type
    Strategy = 'Gradient Boosting Classifier'
    MODL = GradientBoostingClassifier(random_state=0, learning_rate=0.001, n_estimators=10000)
    if SingleCandlePredict:
        for step in range(TestSize):
            x_train,y_train,x_test,y_test = train_test_data_prepar_onebyone(step)     
            train_test_data = f'x_train.class  ={x_train.__class__}\nx_train.shape  ={x_train.shape}\nx_train  =\n{x_train}\n\ny_train.class  ={y_train.__class__}\ny_train.shape  ={y_train.shape}\ny_train  =\n{y_train}\n\nx_test.class   ={x_test.__class__}\nx_test.shape   ={x_test.shape}\nx_test   =\n{x_test}\n\ny_test.shape   ={y_test.shape}\ny_test   =\n{y_test}'
            print(train_test_data)
            
            MODL.fit(x_train, y_train)    
            ALARM.Alarm('Beep')
            y_pred_train = MODL.predict(x_train)
            y_pred_test  = MODL.predict(x_test)
            if step == 0:
                Y_PRED_TRAIN = y_pred_train
            Y_PRED_Storm_TEST      = np.concatenate((Y_PRED_Storm_TEST,y_pred_test),axis=0)
            print('Y_PRED_Storm_TEST',Y_PRED_Storm_TEST)
        if not Concise_Report:
            Show_In_Win('train_test_data',train_test_data)     
        pred=list(Y_PRED_TRAIN) + list(Y_PRED_Storm_TEST)
        df['predicted']=pd.DataFrame(pred,index=df.iloc[-len(pred):].index)
              
        y_pred_test = Y_PRED_Storm_TEST
        y_test  = labels[-TestSize:]     
    else:
        x_train,y_train,x_test,y_test = train_test_data_preparation()    
        train_test_data = f'x_train.shape  ={x_train.shape}\nx_train  =\n{x_train}\n\ny_train.shape  ={y_train.shape}\ny_train  =\n{y_train}\n\nx_test.shape   ={x_test.shape}\nx_test   =\n{x_test}\n\ny_test.shape   ={y_test.shape}\ny_test   =\n{y_test}'
        print(train_test_data)
        if not Concise_Report:
            Show_In_Win('train_test_data',train_test_data) 
        MODL.fit(x_train, y_train)
            
        ALARM.Alarm('Beep')
        y_pred_train = MODL.predict(x_train)
        y_pred_test  = MODL.predict(x_test)

    predTXT = f'''
y_pred_train:\n{y_pred_train}\n
y_pred_test:\n{y_pred_test}'''
    #print(predTXT)
    if not Concise_Report:
        Show_In_Win('train_test prediction',predTXT)
    Eval_Results(y_train,y_pred_train,y_test,y_pred_test)
    Calculate_Equity(y_train,y_pred_train,y_test,y_pred_test)

def XGBC():
    from xgboost import XGBClassifier
    global Strategy
    global y_test
    global y_pred_test
    global Y_PRED_Storm_TEST; Y_PRED_Storm_TEST  = np.array([],dtype=int)
    global Scaling_Type
    Strategy = 'X Gradient Boosting Classifier'
    MODL = XGBClassifier()
    if SingleCandlePredict:
        for step in range(TestSize):
            x_train,y_train,x_test,y_test = train_test_data_prepar_onebyone(step)     
            train_test_data = f'x_train.class  ={x_train.__class__}\nx_train.shape  ={x_train.shape}\nx_train  =\n{x_train}\n\ny_train.class  ={y_train.__class__}\ny_train.shape  ={y_train.shape}\ny_train  =\n{y_train}\n\nx_test.class   ={x_test.__class__}\nx_test.shape   ={x_test.shape}\nx_test   =\n{x_test}\n\ny_test.shape   ={y_test.shape}\ny_test   =\n{y_test}'
            print(train_test_data)
            
            MODL.fit(x_train, y_train)    
            ALARM.Alarm('Beep')
            y_pred_train = MODL.predict(x_train)
            y_pred_test  = MODL.predict(x_test)
            if step == 0:
                Y_PRED_TRAIN = y_pred_train
            Y_PRED_Storm_TEST      = np.concatenate((Y_PRED_Storm_TEST,y_pred_test),axis=0)
            print('Y_PRED_Storm_TEST',Y_PRED_Storm_TEST)
        if not Concise_Report:
            Show_In_Win('train_test_data',train_test_data)     
        pred=list(Y_PRED_TRAIN) + list(Y_PRED_Storm_TEST)
        df['predicted']=pd.DataFrame(pred,index=df.iloc[-len(pred):].index)
              
        y_pred_test = Y_PRED_Storm_TEST
        y_test  = labels[-TestSize:]      
    else:
        x_train,y_train,x_test,y_test = train_test_data_preparation()    
        train_test_data = f'x_train.shape  ={x_train.shape}\nx_train  =\n{x_train}\n\ny_train.shape  ={y_train.shape}\ny_train  =\n{y_train}\n\nx_test.shape   ={x_test.shape}\nx_test   =\n{x_test}\n\ny_test.shape   ={y_test.shape}\ny_test   =\n{y_test}'
        print(train_test_data)
        if not Concise_Report:
            Show_In_Win('train_test_data',train_test_data) 
        MODL.fit(x_train, y_train)
            
        ALARM.Alarm('Beep')
        y_pred_train = MODL.predict(x_train)
        y_pred_test  = MODL.predict(x_test)

    predTXT = f'''
y_pred_train:\n{y_pred_train}\n
y_pred_test:\n{y_pred_test}'''
    #print(predTXT)
    if not Concise_Report:
        Show_In_Win('train_test prediction',predTXT)
    Eval_Results(y_train,y_pred_train,y_test,y_pred_test)
    Calculate_Equity(y_train,y_pred_train,y_test,y_pred_test)

def LGBMC():
    from lightgbm import LGBMClassifier
    global Strategy
    global y_test
    global y_pred_test
    global Y_PRED_Storm_TEST; Y_PRED_Storm_TEST  = np.array([],dtype=int)
    global Scaling_Type
    Strategy = 'Light Gradient Boosting Classifier'
    MODL = LGBMClassifier(learning_rate=0.09,max_depth=-5,random_state=42)
    if SingleCandlePredict:
        for step in range(TestSize):
            x_train,y_train,x_test,y_test = train_test_data_prepar_onebyone(step)     
            train_test_data = f'x_train.class  ={x_train.__class__}\nx_train.shape  ={x_train.shape}\nx_train  =\n{x_train}\n\ny_train.class  ={y_train.__class__}\ny_train.shape  ={y_train.shape}\ny_train  =\n{y_train}\n\nx_test.class   ={x_test.__class__}\nx_test.shape   ={x_test.shape}\nx_test   =\n{x_test}\n\ny_test.shape   ={y_test.shape}\ny_test   =\n{y_test}'
            print(train_test_data)
            MODL.fit(x_train, y_train)    
            ALARM.Alarm('Beep')
            y_pred_train = MODL.predict(x_train)
            y_pred_test  = MODL.predict(x_test)
            if step == 0:
                Y_PRED_TRAIN = y_pred_train
            Y_PRED_Storm_TEST      = np.concatenate((Y_PRED_Storm_TEST,y_pred_test),axis=0)
            print('Y_PRED_Storm_TEST',Y_PRED_Storm_TEST)
        if not Concise_Report:
            Show_In_Win('train_test_data',train_test_data)     
        pred=list(Y_PRED_TRAIN) + list(Y_PRED_Storm_TEST)
        df['predicted']=pd.DataFrame(pred,index=df.iloc[-len(pred):].index)
              
        y_pred_test = Y_PRED_Storm_TEST
        y_test  = labels[-TestSize:]     
    else:
        x_train,y_train,x_test,y_test = train_test_data_preparation()    
        train_test_data = f'x_train.shape  ={x_train.shape}\nx_train  =\n{x_train}\n\ny_train.shape  ={y_train.shape}\ny_train  =\n{y_train}\n\nx_test.shape   ={x_test.shape}\nx_test   =\n{x_test}\n\ny_test.shape   ={y_test.shape}\ny_test   =\n{y_test}'
        print(train_test_data)
        if not Concise_Report:
            Show_In_Win('train_test_data',train_test_data) 
        MODL.fit(x_train, y_train)
            
        ALARM.Alarm('Beep')
        y_pred_train = MODL.predict(x_train)
        y_pred_test  = MODL.predict(x_test)

    predTXT = f'''
y_pred_train:\n{y_pred_train}\n
y_pred_test:\n{y_pred_test}'''
    #print(predTXT)
    if not Concise_Report:
        Show_In_Win('train_test prediction',predTXT)
    Eval_Results(y_train,y_pred_train,y_test,y_pred_test)
    Calculate_Equity(y_train,y_pred_train,y_test,y_pred_test)

def CatBC():
    from catboost import CatBoostClassifier
    global Strategy
    global y_test
    global y_pred_test
    global Y_PRED_Storm_TEST; Y_PRED_Storm_TEST  = np.array([],dtype=int)
    global Scaling_Type
    Strategy = 'CatBoost Classifier'
    MODL = CatBoostClassifier(iterations=10)
    if SingleCandlePredict:
        for step in range(TestSize):
            x_train,y_train,x_test,y_test = train_test_data_prepar_onebyone(step)     
            train_test_data = f'x_train.class  ={x_train.__class__}\nx_train.shape  ={x_train.shape}\nx_train  =\n{x_train}\n\ny_train.class  ={y_train.__class__}\ny_train.shape  ={y_train.shape}\ny_train  =\n{y_train}\n\nx_test.class   ={x_test.__class__}\nx_test.shape   ={x_test.shape}\nx_test   =\n{x_test}\n\ny_test.shape   ={y_test.shape}\ny_test   =\n{y_test}'
            print(train_test_data)
            MODL.fit(x_train, y_train)    
            ALARM.Alarm('Beep')
            y_pred_train = MODL.predict(x_train)
            y_pred_test  = MODL.predict(x_test)
            if step == 0:
                Y_PRED_TRAIN = y_pred_train
            Y_PRED_Storm_TEST      = np.concatenate((Y_PRED_Storm_TEST,y_pred_test),axis=0)
            print('Y_PRED_Storm_TEST',Y_PRED_Storm_TEST)
        if not Concise_Report:
            Show_In_Win('train_test_data',train_test_data)     
        pred=list(Y_PRED_TRAIN) + list(Y_PRED_Storm_TEST)
        df['predicted']=pd.DataFrame(pred,index=df.iloc[-len(pred):].index)
              
        y_pred_test = Y_PRED_Storm_TEST
        y_test  = labels[-TestSize:]   
    else:
        x_train,y_train,x_test,y_test = train_test_data_preparation()    
        train_test_data = f'x_train.shape  ={x_train.shape}\nx_train  =\n{x_train}\n\ny_train.shape  ={y_train.shape}\ny_train  =\n{y_train}\n\nx_test.shape   ={x_test.shape}\nx_test   =\n{x_test}\n\ny_test.shape   ={y_test.shape}\ny_test   =\n{y_test}'
        print(train_test_data)
        if not Concise_Report:
            Show_In_Win('train_test_data',train_test_data) 
        MODL.fit(x_train, y_train)
            
        ALARM.Alarm('Beep')
        y_pred_train = MODL.predict(x_train)
        y_pred_test  = MODL.predict(x_test)

    predTXT = f'''
y_pred_train:\n{y_pred_train}\n
y_pred_test:\n{y_pred_test}'''
    #print(predTXT)
    if not Concise_Report:
        Show_In_Win('train_test prediction',predTXT)
    Eval_Results(y_train,y_pred_train,y_test,y_pred_test)
    Calculate_Equity(y_train,y_pred_train,y_test,y_pred_test)

def AdaBC():
    from sklearn.ensemble import AdaBoostClassifier
    global Strategy
    global y_test
    global y_pred_test
    global Y_PRED_Storm_TEST; Y_PRED_Storm_TEST  = np.array([],dtype=int)
    global Scaling_Type
    Strategy = 'AdaBoost Classifier'
    MODL = AdaBoostClassifier(random_state=0, n_estimators=100)
    if SingleCandlePredict:
        for step in range(TestSize):
            x_train,y_train,x_test,y_test = train_test_data_prepar_onebyone(step)     
            train_test_data = f'x_train.class  ={x_train.__class__}\nx_train.shape  ={x_train.shape}\nx_train  =\n{x_train}\n\ny_train.class  ={y_train.__class__}\ny_train.shape  ={y_train.shape}\ny_train  =\n{y_train}\n\nx_test.class   ={x_test.__class__}\nx_test.shape   ={x_test.shape}\nx_test   =\n{x_test}\n\ny_test.shape   ={y_test.shape}\ny_test   =\n{y_test}'
            print(train_test_data)
            MODL.fit(x_train, y_train)    
            ALARM.Alarm('Beep')
            y_pred_train = MODL.predict(x_train)
            y_pred_test  = MODL.predict(x_test)
            if step == 0:
                Y_PRED_TRAIN = y_pred_train
            Y_PRED_Storm_TEST      = np.concatenate((Y_PRED_Storm_TEST,y_pred_test),axis=0)
            print('Y_PRED_Storm_TEST',Y_PRED_Storm_TEST)
        if not Concise_Report:
            Show_In_Win('train_test_data',train_test_data)     
        pred=list(Y_PRED_TRAIN) + list(Y_PRED_Storm_TEST)
        df['predicted']=pd.DataFrame(pred,index=df.iloc[-len(pred):].index)
              
        y_pred_test = Y_PRED_Storm_TEST
        y_test  = labels[-TestSize:]   
    else:
        x_train,y_train,x_test,y_test = train_test_data_preparation()    
        train_test_data = f'x_train.shape  ={x_train.shape}\nx_train  =\n{x_train}\n\ny_train.shape  ={y_train.shape}\ny_train  =\n{y_train}\n\nx_test.shape   ={x_test.shape}\nx_test   =\n{x_test}\n\ny_test.shape   ={y_test.shape}\ny_test   =\n{y_test}'
        print(train_test_data)
        if not Concise_Report:
            Show_In_Win('train_test_data',train_test_data) 
        MODL.fit(x_train, y_train)
            
        ALARM.Alarm('Beep')
        y_pred_train = MODL.predict(x_train)
        y_pred_test  = MODL.predict(x_test)

    predTXT = f'''
y_pred_train:\n{y_pred_train}\n
y_pred_test:\n{y_pred_test}'''
    #print(predTXT)
    if not Concise_Report:
        Show_In_Win('train_test prediction',predTXT)
    Eval_Results(y_train,y_pred_train,y_test,y_pred_test)
    Calculate_Equity(y_train,y_pred_train,y_test,y_pred_test)

def ExtraTreesC(): 
    from sklearn.ensemble import ExtraTreesClassifier
    global Strategy
    global y_test
    global y_pred_test
    global Y_PRED_Storm_TEST; Y_PRED_Storm_TEST  = np.array([],dtype=int)
    global Scaling_Type
    Strategy = 'Extra Trees Classifier'
    MODL = ExtraTreesClassifier()
    if SingleCandlePredict:
        for step in range(TestSize):
            x_train,y_train,x_test,y_test = train_test_data_prepar_onebyone(step)     
            train_test_data = f'x_train.class  ={x_train.__class__}\nx_train.shape  ={x_train.shape}\nx_train  =\n{x_train}\n\ny_train.class  ={y_train.__class__}\ny_train.shape  ={y_train.shape}\ny_train  =\n{y_train}\n\nx_test.class   ={x_test.__class__}\nx_test.shape   ={x_test.shape}\nx_test   =\n{x_test}\n\ny_test.shape   ={y_test.shape}\ny_test   =\n{y_test}'
            print(train_test_data)
            MODL.fit(x_train, y_train)    
            ALARM.Alarm('Beep')
            y_pred_train = MODL.predict(x_train)
            y_pred_test  = MODL.predict(x_test)
            if step == 0:
                Y_PRED_TRAIN = y_pred_train
            Y_PRED_Storm_TEST      = np.concatenate((Y_PRED_Storm_TEST,y_pred_test),axis=0)
            print('Y_PRED_Storm_TEST',Y_PRED_Storm_TEST)
        if not Concise_Report:
            Show_In_Win('train_test_data',train_test_data)     
        pred=list(Y_PRED_TRAIN) + list(Y_PRED_Storm_TEST)
        df['predicted']=pd.DataFrame(pred,index=df.iloc[-len(pred):].index)
              
        y_pred_test = Y_PRED_Storm_TEST
        y_test  = labels[-TestSize:]  
    else:
        x_train,y_train,x_test,y_test = train_test_data_preparation()    
        train_test_data = f'x_train.shape  ={x_train.shape}\nx_train  =\n{x_train}\n\ny_train.shape  ={y_train.shape}\ny_train  =\n{y_train}\n\nx_test.shape   ={x_test.shape}\nx_test   =\n{x_test}\n\ny_test.shape   ={y_test.shape}\ny_test   =\n{y_test}'
        print(train_test_data)
        if not Concise_Report:
            Show_In_Win('train_test_data',train_test_data) 
        MODL.fit(x_train, y_train)
            
        ALARM.Alarm('Beep')
        y_pred_train = MODL.predict(x_train)
        y_pred_test  = MODL.predict(x_test)

    predTXT = f'''
y_pred_train:\n{y_pred_train}\n
y_pred_test:\n{y_pred_test}'''
    #print(predTXT)
    if not Concise_Report:
        Show_In_Win('train_test prediction',predTXT)
    Eval_Results(y_train,y_pred_train,y_test,y_pred_test)
    Calculate_Equity(y_train,y_pred_train,y_test,y_pred_test)

def LDa():
    from sklearn.discriminant_analysis import LinearDiscriminantAnalysis
    global Strategy
    global y_test
    global y_pred_test
    global Y_PRED_Storm_TEST; Y_PRED_Storm_TEST  = np.array([],dtype=int)
    global Scaling_Type
    Strategy = 'Linear Discriminant Analysis'
    MODL = LinearDiscriminantAnalysis()
    if SingleCandlePredict:
        for step in range(TestSize):
            x_train,y_train,x_test,y_test = train_test_data_prepar_onebyone(step)     
            train_test_data = f'x_train.class  ={x_train.__class__}\nx_train.shape  ={x_train.shape}\nx_train  =\n{x_train}\n\ny_train.class  ={y_train.__class__}\ny_train.shape  ={y_train.shape}\ny_train  =\n{y_train}\n\nx_test.class   ={x_test.__class__}\nx_test.shape   ={x_test.shape}\nx_test   =\n{x_test}\n\ny_test.shape   ={y_test.shape}\ny_test   =\n{y_test}'
            print(train_test_data)
            MODL.fit(x_train, y_train)    
            ALARM.Alarm('Beep')
            y_pred_train = MODL.predict(x_train)
            y_pred_test  = MODL.predict(x_test)
            if step == 0:
                Y_PRED_TRAIN = y_pred_train
            Y_PRED_Storm_TEST      = np.concatenate((Y_PRED_Storm_TEST,y_pred_test),axis=0)
            print('Y_PRED_Storm_TEST',Y_PRED_Storm_TEST)
        if not Concise_Report:
            Show_In_Win('train_test_data',train_test_data)     
        pred=list(Y_PRED_TRAIN) + list(Y_PRED_Storm_TEST)
        df['predicted']=pd.DataFrame(pred,index=df.iloc[-len(pred):].index)
              
        y_pred_test = Y_PRED_Storm_TEST
        y_test  = labels[-TestSize:]  
    else:
        x_train,y_train,x_test,y_test = train_test_data_preparation()    
        train_test_data = f'x_train.shape  ={x_train.shape}\nx_train  =\n{x_train}\n\ny_train.shape  ={y_train.shape}\ny_train  =\n{y_train}\n\nx_test.shape   ={x_test.shape}\nx_test   =\n{x_test}\n\ny_test.shape   ={y_test.shape}\ny_test   =\n{y_test}'
        print(train_test_data)
        if not Concise_Report:
            Show_In_Win('train_test_data',train_test_data) 
        MODL.fit(x_train, y_train)
            
        ALARM.Alarm('Beep')
        y_pred_train = MODL.predict(x_train)
        y_pred_test  = MODL.predict(x_test)

    predTXT = f'''
y_pred_train:\n{y_pred_train}\n
y_pred_test:\n{y_pred_test}'''
    #print(predTXT)
    if not Concise_Report:
        Show_In_Win('train_test prediction',predTXT)
    Eval_Results(y_train,y_pred_train,y_test,y_pred_test)
    Calculate_Equity(y_train,y_pred_train,y_test,y_pred_test)

def LSTMC():
    from keras import Sequential
    from keras.layers import Dense, LSTM#, Dropout
    global Strategy
    global y_test
    global y_pred_test
    global Y_PRED_Storm_TEST
    global Scaling_Type
    global df
    global Y_PRED_TRAIN
    Y_PRED_Storm_TEST  = np.array([],dtype=int)
    Strategy = 'LSTM Classifier'
    MODL = Sequential()
    XSIZE = features.shape[1]
    MODL.add(LSTM(units = XSIZE*4, return_sequences=True, input_shape=(XSIZE,1)))
    MODL.add(LSTM(units = XSIZE*6,activation = 'relu', return_sequences= True))
    MODL.add(LSTM(XSIZE*9,activation = 'relu', return_sequences= True))
    MODL.add(LSTM(XSIZE*12,activation = 'relu', return_sequences= False))
    MODL.add(Dense(XSIZE*12,activation = 'relu'))
    MODL.add(Dense(1 , activation = 'softmax'))
    MODL.compile(optimizer='sgd', loss='mean_squared_error', metrics=['accuracy']) #optimizer='adam',loss='mean_squared_error',metrics=['accuracy'] loss='binary_crossentropy' was tested and results wasn't quite good

    if SingleCandlePredict:
        for step in range(TestSize):
            x_train,y_train,x_test,y_test = train_test_data_prepar_onebyone(step)     
            train_test_data = f'x_train.class  ={x_train.__class__}\nx_train.shape  ={x_train.shape}\nx_train  =\n{x_train}\n\ny_train.class  ={y_train.__class__}\ny_train.shape  ={y_train.shape}\ny_train  =\n{y_train}\n\nx_test.class   ={x_test.__class__}\nx_test.shape   ={x_test.shape}\nx_test   =\n{x_test}\n\ny_test.shape   ={y_test.shape}\ny_test   =\n{y_test}'
            print(train_test_data)
            MODL.fit(x_train, y_train)    
            ALARM.Alarm('Beep')
            y_pred_train = MODL.predict(x_train)
            y_pred_test  = MODL.predict(x_test)
            if step == 0:
                Y_PRED_TRAIN = y_pred_train
            Y_PRED_Storm_TEST      = np.concatenate((Y_PRED_Storm_TEST,y_pred_test[0]),axis=0)
        print('Y_PRED_Storm_TEST:',Y_PRED_Storm_TEST)
        if not Concise_Report:
            Show_In_Win('train_test_data',train_test_data)
            
        print('Y_PRED_TRAIN:',Y_PRED_TRAIN)
        print('Y_PRED_TRAIN.__class__:',Y_PRED_TRAIN.__class__)
        
        print('Y_PRED_Storm_TEST:',Y_PRED_Storm_TEST)
        print('Y_PRED_Storm_TEST.__class__:',Y_PRED_TRAIN.__class__)
        
        print('Y_PRED_TRAIN shape:',Y_PRED_TRAIN.shape)
        
        print('Y_PRED_Storm_TEST shape:',Y_PRED_Storm_TEST.shape)
        
        
        print('list(Y_PRED_TRAIN):',list(Y_PRED_TRAIN))
        
        print('list(Y_PRED_Storm_TEST):',list(Y_PRED_Storm_TEST))
        Y_PRED_TRAIN= Y_PRED_TRAIN.ravel()
        pred = list(Y_PRED_TRAIN) + list(Y_PRED_Storm_TEST)
        #Y_PRED_TRAIN = Y_PRED_TRAIN.flatten()
        
        print('Y_PRED_TRAIN:',Y_PRED_TRAIN)
              
        df['predicted'] =pd.DataFrame(pred,index=df.iloc[-len(pred):].index)
              
        y_pred_test = Y_PRED_Storm_TEST
        y_test  = labels[-TestSize:]
        
    else:
        x_train,y_train,x_test,y_test = train_test_data_preparation()    
        train_test_data = f'x_train.shape  ={x_train.shape}\nx_train  =\n{x_train}\n\ny_train.shape  ={y_train.shape}\ny_train  =\n{y_train}\n\nx_test.shape   ={x_test.shape}\nx_test   =\n{x_test}\n\ny_test.shape   ={y_test.shape}\ny_test   =\n{y_test}'
        print(train_test_data)
        if not Concise_Report:
            Show_In_Win('train_test_data',train_test_data) 
        MODL.fit(x_train, y_train)
            
        ALARM.Alarm('Beep')
        y_pred_train = MODL.predict(x_train)
        y_pred_test  = MODL.predict(x_test)

    predTXT = f'''
y_pred_train:\n{y_pred_train}\n
y_pred_test:\n{y_pred_test}'''
    #print(predTXT)
    if not Concise_Report:
        Show_In_Win('train_test prediction',predTXT)
    Eval_Results(y_train,y_pred_train,y_test,y_pred_test)
    Calculate_Equity(y_train,y_pred_train,y_test,y_pred_test)

def DeepNN():
    from keras import Sequential
    from keras.layers import Dense#, LSTM, Dropout
    global Strategy
    global y_test
    global y_pred_test
    global Y_PRED_Storm_TEST; Y_PRED_Storm_TEST  = np.array([],dtype=int)
    global Scaling_Type
    Strategy = 'Deep Neural Network Classifier'
    x_train,y_train,x_test,y_test = train_test_data_prepar_onebyone(0)
    Deepclf = Sequential()
    Deepclf.add(Dense(x_train.shape[1]*4,activation='relu',input_shape=(x_train.shape[1:])))#       Deepclf.add(Dense(50,activation='relu',input_shape=(x_train.shape[1:])))#input_shape=(x_train.shape[1],1)))
    Deepclf.add(Dense(x_train.shape[1]*8,activation='relu'))
    Deepclf.add(Dense(x_train.shape[1]*16,activation='relu')) #,activation='tanh'
    Deepclf.add(Dense(x_train.shape[1]*8,activation='relu'))
    Deepclf.add(Dense(x_train.shape[1]*4,activation='relu'))
    Deepclf.add(Dense(1,activation='softmax')) #activation='sigmoid'
    Deepclf.compile(optimizer='sgd', loss='mean_squared_error', metrics=['accuracy']) #metrics='MAE'), loss='mean_squared_error', loss='binary_crossentropy',     #optimizer='sgd' 'adam' was tested and results was not acceptable
    #Deepclf.compile(loss="sparse_categorical_crossentropy",optimizer='sgd',metrics='accuracy' )
    if SingleCandlePredict:
        for step in range(TestSize):
            x_train,y_train,x_test,y_test = train_test_data_prepar_onebyone(step)     
            train_test_data = f'x_train.class  ={x_train.__class__}\nx_train.shape  ={x_train.shape}\nx_train  =\n{x_train}\n\ny_train.class  ={y_train.__class__}\ny_train.shape  ={y_train.shape}\ny_train  =\n{y_train}\n\nx_test.class   ={x_test.__class__}\nx_test.shape   ={x_test.shape}\nx_test   =\n{x_test}\n\ny_test.shape   ={y_test.shape}\ny_test   =\n{y_test}'
            print(train_test_data)
            Deepclf.fit(x_train, y_train)    
            ALARM.Alarm('Beep')
            y_pred_train = Deepclf.predict(x_train)
            y_pred_test  = Deepclf.predict(x_test)
            if step == 0:
                Y_PRED_TRAIN = y_pred_train
            Y_PRED_Storm_TEST      = np.concatenate((Y_PRED_Storm_TEST,y_pred_test[0]),axis=0)
            print('Y_PRED_Storm_TEST',Y_PRED_Storm_TEST)
        if not Concise_Report:
            Show_In_Win('train_test_data',train_test_data)
        Y_PRED_TRAIN= Y_PRED_TRAIN.ravel()
        pred=list(Y_PRED_TRAIN) + list(Y_PRED_Storm_TEST)
        df['predicted']=pd.DataFrame(pred,index=df.iloc[-len(pred):].index)
              
        y_pred_test = Y_PRED_Storm_TEST
        y_test  = labels[-TestSize:]
        
    else:
        x_train,y_train,x_test,y_test = train_test_data_preparation()    
        train_test_data = f'x_train.shape  ={x_train.shape}\nx_train  =\n{x_train}\n\ny_train.shape  ={y_train.shape}\ny_train  =\n{y_train}\n\nx_test.shape   ={x_test.shape}\nx_test   =\n{x_test}\n\ny_test.shape   ={y_test.shape}\ny_test   =\n{y_test}'
        print(train_test_data)
        if not Concise_Report:
            Show_In_Win('train_test_data',train_test_data) 
        Deepclf.fit(x_train, y_train)
            
        ALARM.Alarm('Beep')
        y_pred_train = Deepclf.predict(x_train)
        y_pred_test  = Deepclf.predict(x_test)

    predTXT = f'''
y_pred_train:\n{y_pred_train}\n
y_pred_test:\n{y_pred_test}'''
    #print(predTXT)
    if not Concise_Report:
        Show_In_Win('train_test prediction',predTXT)
    Eval_Results(y_train,y_pred_train,y_test,y_pred_test)
    Calculate_Equity(y_train,y_pred_train,y_test,y_pred_test)

def Lin_Reg():
    from sklearn.linear_model import LinearRegression
    global Strategy
    global y_test
    global y_pred_test
    global Y_PRED_Storm_TEST; Y_PRED_Storm_TEST  = np.array([],dtype=int)
    global Scaling_Type
    Strategy = 'Linear Regressor'
    MODL = LinearRegression()
    if SingleCandlePredict:
        for step in range(TestSize):
            x_train,y_train,x_test,y_test = train_test_data_prepar_onebyone(step)     
            train_test_data = f'x_train.class  ={x_train.__class__}\nx_train.shape  ={x_train.shape}\nx_train  =\n{x_train}\n\ny_train.class  ={y_train.__class__}\ny_train.shape  ={y_train.shape}\ny_train  =\n{y_train}\n\nx_test.class   ={x_test.__class__}\nx_test.shape   ={x_test.shape}\nx_test   =\n{x_test}\n\ny_test.shape   ={y_test.shape}\ny_test   =\n{y_test}'
            print(train_test_data)
            MODL.fit(x_train, y_train)    
            ALARM.Alarm('Beep')
            y_pred_train = MODL.predict(x_train)
            y_pred_test  = MODL.predict(x_test)
            if step == 0:
                Y_PRED_TRAIN = y_pred_train
            Y_PRED_Storm_TEST      = np.concatenate((Y_PRED_Storm_TEST,y_pred_test),axis=0)
            print('Y_PRED_Storm_TEST',Y_PRED_Storm_TEST)
        if not Concise_Report:
            Show_In_Win('train_test_data',train_test_data)     
        pred=list(Y_PRED_TRAIN) + list(Y_PRED_Storm_TEST)
        df['predicted']=pd.DataFrame(pred,index=df.iloc[-len(pred):].index)
              
        y_pred_test = Y_PRED_Storm_TEST
        y_test  = labels[-TestSize:]
    else:
        x_train,y_train,x_test,y_test = train_test_data_preparation()    
        train_test_data = f'x_train.shape  ={x_train.shape}\nx_train  =\n{x_train}\n\ny_train.shape  ={y_train.shape}\ny_train  =\n{y_train}\n\nx_test.shape   ={x_test.shape}\nx_test   =\n{x_test}\n\ny_test.shape   ={y_test.shape}\ny_test   =\n{y_test}'
        print(train_test_data)
        if not Concise_Report:
            Show_In_Win('train_test_data',train_test_data) 
        MODL.fit(x_train, y_train)
            
        ALARM.Alarm('Beep')
        y_pred_train = MODL.predict(x_train)
        y_pred_test  = MODL.predict(x_test)

    predTXT = f'''
y_pred_train:\n{y_pred_train}\n
y_pred_test:\n{y_pred_test}'''
    #print(predTXT)
    if not Concise_Report:
        Show_In_Win('train_test prediction',predTXT)
    Eval_Results(y_train,y_pred_train,y_test,y_pred_test)
    Calculate_Equity(y_train,y_pred_train,y_test,y_pred_test)

def DT_Reg():
    from sklearn.tree import DecisionTreeRegressor
    global Strategy
    global y_test
    global y_pred_test
    global Y_PRED_Storm_TEST; Y_PRED_Storm_TEST  = np.array([],dtype=int)
    global Scaling_Type
    Strategy = 'Decision Tree Regressor'
    MODL = DecisionTreeRegressor(random_state= 4)
    if SingleCandlePredict:
        for step in range(TestSize):
            x_train,y_train,x_test,y_test = train_test_data_prepar_onebyone(step)     
            train_test_data = f'x_train.class  ={x_train.__class__}\nx_train.shape  ={x_train.shape}\nx_train  =\n{x_train}\n\ny_train.class  ={y_train.__class__}\ny_train.shape  ={y_train.shape}\ny_train  =\n{y_train}\n\nx_test.class   ={x_test.__class__}\nx_test.shape   ={x_test.shape}\nx_test   =\n{x_test}\n\ny_test.shape   ={y_test.shape}\ny_test   =\n{y_test}'
            print(train_test_data)
            MODL.fit(x_train, y_train)    
            ALARM.Alarm('Beep')
            y_pred_train = MODL.predict(x_train)
            y_pred_test  = MODL.predict(x_test)
            if step == 0:
                Y_PRED_TRAIN = y_pred_train
            Y_PRED_Storm_TEST      = np.concatenate((Y_PRED_Storm_TEST,y_pred_test),axis=0)
            print('Y_PRED_Storm_TEST',Y_PRED_Storm_TEST)
        if not Concise_Report:
            Show_In_Win('train_test_data',train_test_data)     
        pred=list(Y_PRED_TRAIN) + list(Y_PRED_Storm_TEST)
        df['predicted']=pd.DataFrame(pred,index=df.iloc[-len(pred):].index)
              
        y_pred_test = Y_PRED_Storm_TEST
        y_test  = labels[-TestSize:]
        
    else:
        x_train,y_train,x_test,y_test = train_test_data_preparation()    
        train_test_data = f'x_train.shape  ={x_train.shape}\nx_train  =\n{x_train}\n\ny_train.shape  ={y_train.shape}\ny_train  =\n{y_train}\n\nx_test.shape   ={x_test.shape}\nx_test   =\n{x_test}\n\ny_test.shape   ={y_test.shape}\ny_test   =\n{y_test}'
        print(train_test_data)
        if not Concise_Report:
            Show_In_Win('train_test_data',train_test_data) 
        MODL.fit(x_train, y_train)
            
        ALARM.Alarm('Beep')
        y_pred_train = MODL.predict(x_train)
        y_pred_test  = MODL.predict(x_test)

    predTXT = f'''
y_pred_train:\n{y_pred_train}\n
y_pred_test:\n{y_pred_test}'''
    #print(predTXT)
    if not Concise_Report:
        Show_In_Win('train_test prediction',predTXT)
    Eval_Results(y_train,y_pred_train,y_test,y_pred_test)
    Calculate_Equity(y_train,y_pred_train,y_test,y_pred_test)

def RF_Reg():
    from sklearn.ensemble import RandomForestRegressor
    global Strategy
    global y_test
    global y_pred_test
    global Y_PRED_Storm_TEST; Y_PRED_Storm_TEST  = np.array([],dtype=int)
    global Scaling_Type
    Strategy = 'Random Forest Regressor'
    MODL = RandomForestRegressor(max_depth=349, min_samples_leaf=2, min_samples_split=2, n_estimators=1750, random_state=42)
    if SingleCandlePredict:
        for step in range(TestSize):
            x_train,y_train,x_test,y_test = train_test_data_prepar_onebyone(step)     
            train_test_data = f'x_train.class  ={x_train.__class__}\nx_train.shape  ={x_train.shape}\nx_train  =\n{x_train}\n\ny_train.class  ={y_train.__class__}\ny_train.shape  ={y_train.shape}\ny_train  =\n{y_train}\n\nx_test.class   ={x_test.__class__}\nx_test.shape   ={x_test.shape}\nx_test   =\n{x_test}\n\ny_test.shape   ={y_test.shape}\ny_test   =\n{y_test}'
            print(train_test_data)
            MODL.fit(x_train, y_train)    
            ALARM.Alarm('Beep')
            y_pred_train = MODL.predict(x_train)
            y_pred_test  = MODL.predict(x_test)
            if step == 0:
                Y_PRED_TRAIN = y_pred_train
            Y_PRED_Storm_TEST      = np.concatenate((Y_PRED_Storm_TEST,y_pred_test),axis=0)
            print('Y_PRED_Storm_TEST',Y_PRED_Storm_TEST)
        if not Concise_Report:
            Show_In_Win('train_test_data',train_test_data)     
        pred=list(Y_PRED_TRAIN) + list(Y_PRED_Storm_TEST)
        df['predicted']=pd.DataFrame(pred,index=df.iloc[-len(pred):].index)
              
        y_pred_test = Y_PRED_Storm_TEST
        y_test  = labels[-TestSize:]
        
    else:
        x_train,y_train,x_test,y_test = train_test_data_preparation()    
        train_test_data = f'x_train.shape  ={x_train.shape}\nx_train  =\n{x_train}\n\ny_train.shape  ={y_train.shape}\ny_train  =\n{y_train}\n\nx_test.shape   ={x_test.shape}\nx_test   =\n{x_test}\n\ny_test.shape   ={y_test.shape}\ny_test   =\n{y_test}'
        print(train_test_data)
        if not Concise_Report:
            Show_In_Win('train_test_data',train_test_data) 
        MODL.fit(x_train, y_train)
            
        ALARM.Alarm('Beep')
        y_pred_train = MODL.predict(x_train)
        y_pred_test  = MODL.predict(x_test)

    predTXT = f'''
y_pred_train:\n{y_pred_train}\n
y_pred_test:\n{y_pred_test}'''
    #print(predTXT)
    if not Concise_Report:
        Show_In_Win('train_test prediction',predTXT)
    Eval_Results(y_train,y_pred_train,y_test,y_pred_test)
    Calculate_Equity(y_train,y_pred_train,y_test,y_pred_test)

def SV_Reg():
    from sklearn.svm import SVR
    global Strategy
    global y_test
    global y_pred_test
    global Y_PRED_Storm_TEST; Y_PRED_Storm_TEST  = np.array([],dtype=int)
    global Scaling_Type
    Strategy = 'Support Vector Regressor'
    MODL = SVR(kernel='rbf', gamma=10, C = 1)
    if SingleCandlePredict:
        for step in range(TestSize):
            x_train,y_train,x_test,y_test = train_test_data_prepar_onebyone(step)     
            train_test_data = f'x_train.class  ={x_train.__class__}\nx_train.shape  ={x_train.shape}\nx_train  =\n{x_train}\n\ny_train.class  ={y_train.__class__}\ny_train.shape  ={y_train.shape}\ny_train  =\n{y_train}\n\nx_test.class   ={x_test.__class__}\nx_test.shape   ={x_test.shape}\nx_test   =\n{x_test}\n\ny_test.shape   ={y_test.shape}\ny_test   =\n{y_test}'
            print(train_test_data)
            MODL.fit(x_train, y_train)    
            ALARM.Alarm('Beep')
            y_pred_train = MODL.predict(x_train)
            y_pred_test  = MODL.predict(x_test)
            if step == 0:
                Y_PRED_TRAIN = y_pred_train
            Y_PRED_Storm_TEST      = np.concatenate((Y_PRED_Storm_TEST,y_pred_test),axis=0)
            print('Y_PRED_Storm_TEST',Y_PRED_Storm_TEST)
        if not Concise_Report:
            Show_In_Win('train_test_data',train_test_data)     
        pred=list(Y_PRED_TRAIN) + list(Y_PRED_Storm_TEST)
        df['predicted']=pd.DataFrame(pred,index=df.iloc[-len(pred):].index)
              
        y_pred_test = Y_PRED_Storm_TEST
        y_test  = labels[-TestSize:]
        
    else:
        x_train,y_train,x_test,y_test = train_test_data_preparation()    
        train_test_data = f'x_train.shape  ={x_train.shape}\nx_train  =\n{x_train}\n\ny_train.shape  ={y_train.shape}\ny_train  =\n{y_train}\n\nx_test.shape   ={x_test.shape}\nx_test   =\n{x_test}\n\ny_test.shape   ={y_test.shape}\ny_test   =\n{y_test}'
        print(train_test_data)
        if not Concise_Report:
            Show_In_Win('train_test_data',train_test_data) 
        MODL.fit(x_train, y_train)
            
        ALARM.Alarm('Beep')
        y_pred_train = MODL.predict(x_train)
        y_pred_test  = MODL.predict(x_test)

    predTXT = f'''
y_pred_train:\n{y_pred_train}\n
y_pred_test:\n{y_pred_test}'''
    #print(predTXT)
    if not Concise_Report:
        Show_In_Win('train_test prediction',predTXT)
    Eval_Results(y_train,y_pred_train,y_test,y_pred_test)
    Calculate_Equity(y_train,y_pred_train,y_test,y_pred_test)

def GSESTRATEGY():
    from sklearn.svm import SVR
    global Strategy
    global y_test
    global y_pred_test
    global Y_PRED_Storm_TEST; Y_PRED_Storm_TEST  = np.array([],dtype=int)
    global Scaling_Type
    Strategy = 'Support Vector Regressor'
    MODL = SVR(kernel='rbf', gamma=10, C = 1)
    if SingleCandlePredict:
        for step in range(TestSize):
            x_train,y_train,x_test,y_test = train_test_data_prepar_onebyone(step)     
            train_test_data = f'x_train.class  ={x_train.__class__}\nx_train.shape  ={x_train.shape}\nx_train  =\n{x_train}\n\ny_train.class  ={y_train.__class__}\ny_train.shape  ={y_train.shape}\ny_train  =\n{y_train}\n\nx_test.class   ={x_test.__class__}\nx_test.shape   ={x_test.shape}\nx_test   =\n{x_test}\n\ny_test.shape   ={y_test.shape}\ny_test   =\n{y_test}'
            print(train_test_data)
            MODL.fit(x_train, y_train)    
            ALARM.Alarm('Beep')
            y_pred_train = MODL.predict(x_train)
            y_pred_test  = MODL.predict(x_test)
            if step == 0:
                Y_PRED_TRAIN = y_pred_train
            Y_PRED_Storm_TEST      = np.concatenate((Y_PRED_Storm_TEST,y_pred_test),axis=0)
            print('Y_PRED_Storm_TEST',Y_PRED_Storm_TEST)
        if not Concise_Report:
            Show_In_Win('train_test_data',train_test_data)     
        pred=list(Y_PRED_TRAIN) + list(Y_PRED_Storm_TEST)
        df['predicted']=pd.DataFrame(pred,index=df.iloc[-len(pred):].index)
              
        y_pred_test = Y_PRED_Storm_TEST
        y_test  = labels[-TestSize:]
        
    else:
        x_train,y_train,x_test,y_test = train_test_data_preparation()    
        train_test_data = f'x_train.shape  ={x_train.shape}\nx_train  =\n{x_train}\n\ny_train.shape  ={y_train.shape}\ny_train  =\n{y_train}\n\nx_test.shape   ={x_test.shape}\nx_test   =\n{x_test}\n\ny_test.shape   ={y_test.shape}\ny_test   =\n{y_test}'
        print(train_test_data)
        if not Concise_Report:
            Show_In_Win('train_test_data',train_test_data) 
        MODL.fit(x_train, y_train)
            
        ALARM.Alarm('Beep')
        y_pred_train = MODL.predict(x_train)
        y_pred_test  = MODL.predict(x_test)

    predTXT = f'''
y_pred_train:\n{y_pred_train}\n
y_pred_test:\n{y_pred_test}'''
    #print(predTXT)
    if not Concise_Report:
        Show_In_Win('train_test prediction',predTXT)
    Eval_Results(y_train,y_pred_train,y_test,y_pred_test)
    Calculate_Equity(y_train,y_pred_train,y_test,y_pred_test)

def SSMRB():
    global Strategy
    global x_train,y_train,x_test,y_test
    global y_pred_train,y_pred_test
    global labels 
    global features
    global TestSize
    global Trainsize
    global FEAT
    Strategy = 'Simple Sum Mean Return Binary'
    try:
        df
    except NameError as ERR:
        showmessage('Dataframe Error',f'Dataframe df is not available\n{ERR}',TIMEOUT=3000,TYP='Erro')
        LoadDataBase('')
    try:
        emptylabel = not labels.any()
    except:
        pass
    try:
        emptylabel = not labels.any().values[0]
    except AttributeError as ERR:
        pass
    
    if emptylabel:
        global WIN16
        global cmbFeat
        WIN16 = tk.Toplevel(root)
        WIN16.bind('<Escape>',EXIT_One)
        LstrCateg = ''
        Label_Select(WIN16)
        WIN16.lift()
        WIN16.focus_force()
        
    showmessage('Forced Features',f'Features are set to "Just Mean Return Indicators"',TIMEOUT=4000,TYP='info')
    FEAT = 'Just Mean Return Indicators'    
    Fea_MESSAGE1.configure(text = f'{FEAT}',fg='#F5CC21')
    
    cmbFeat.set(FEAT)
    Just_Mean_Return_indicators()

    x_train = features[-Trainsize-TestSize:-TestSize]#Prepared_features[-Trainsize-TestSize:-TestSize]
    x_test  = features[-TestSize:]#Prepared_features[-TestSize:]
    y_train = labels[-Trainsize-TestSize:-TestSize]
    y_test  = labels[-TestSize:]
        
    train_test_data = f'x_train with shape {x_train.shape}:\n{x_train}\n\ny_train with shape {y_train.shape}\n:\n{y_train}\n\nx_test with shape {x_test.shape}:\n{x_test}\n\ny_test with shape {y_test.shape}:\n{y_test}'
    print(train_test_data)
    if not Concise_Report:
        Show_In_Win('train_test_data',train_test_data)
    ALARM.Alarm('Beep')
    
    y_pred_train = x_train.sum(axis=1).apply(lambda x: int(1) if (x>0) else int(0))
    y_pred_test  = x_test.sum(axis=1).apply(lambda x: int(1) if (x>0) else int(0))

    predTXT = f'''
y_pred_train:\n{y_pred_train}\n
y_pred_test:\n{y_pred_test}'''
    #print(predTXT)
    if not Concise_Report:
        Show_In_Win('train_test prediction',predTXT)
    Eval_Results(y_train,y_pred_train,y_test,y_pred_test)
    Calculate_Equity(y_train,y_pred_train,y_test,y_pred_test)

def SSMDB():
    global Strategy
    global x_train,y_train,x_test,y_test
    global y_pred_train,y_pred_test
    global labels 
    global features
    global FEAT
    global TestSize
    global Trainsize
    Strategy = 'Simple Sum Mean Delta Binary'
    
    try:
        df
    except NameError as ERR:
        showmessage('Dataframe Error',f'Dataframe df is not available\n{ERR}',TIMEOUT=3000,TYP='Erro')
        LoadDataBase('')
        
    showmessage('Forced Features',f'Features are set to "Just mean Delta"',TIMEOUT=4000,TYP='info')
    FEAT = 'Just mean Delta'    
    Fea_MESSAGE1.configure(text = f'{FEAT}',fg='#F5CC21')
    
    cmbFeat.set(FEAT)
    Just_Mean_Delta()
  
    try:
        emptylabel = not labels.any()
    except ValueError as ERR:
        showmessage('empty label Error',f'No Label is available to display\n{ERR}',TIMEOUT=5000,TYP='Erro')
    try:
        emptylabel = not labels.any().values[0]
    except AttributeError as ERR:
        pass    
    if emptylabel:
        global WIN16
        WIN16 = tk.Toplevel(root)
        WIN16.bind('<Escape>',EXIT_One)
        LstrCateg = ''
        Label_Select(WIN16)
        WIN16.lift()
        WIN16.focus_force()

    x_train = features[-Trainsize-TestSize:-TestSize] #Prepared_features[-Trainsize-TestSize:-TestSize]    
    y_train =   labels[-Trainsize-TestSize:-TestSize]
    x_test  = features[-TestSize:] #Prepared_features[-TestSize:]
    y_test  =   labels[-TestSize:]
    
    train_test_data = f'x_train with shape {x_train.shape}:\n{x_train}\n\ny_train with shape {y_train.shape}\n:\n{y_train}\n\nx_test with shape {x_test.shape}:\n{x_test}\n\ny_test with shape {y_test.shape}:\n{y_test}'
    print(train_test_data)
    if not Concise_Report:
        Show_In_Win('train_test_data',train_test_data)
    ALARM.Alarm('Beep')
    y_pred_train = x_train.sum(axis=1).apply(lambda x: int(1) if (x>0) else int(0))
    y_pred_test  = x_test.sum(axis=1).apply(lambda x: int(1) if (x>0) else int(0))
    
    predTXT = f'''
y_pred_train:\n{y_pred_train}\n
y_pred_test:\n{y_pred_test}'''
    #print(predTXT)
    if not Concise_Report:
        Show_In_Win('train_test prediction',predTXT)
    Eval_Results(y_train,y_pred_train,y_test,y_pred_test)
    Calculate_Equity(y_train,y_pred_train,y_test,y_pred_test)

#◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄► Add Dynamic Lottage  ◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►
def Add_Dynamic_Lottage():
    global df
    global Lottage
    m=[]
    MaxCodeLength_ = Add_Binary_indicators('Lottage')
    if MaxCodeLength_ == None:
        showmessage('Try again',f'Run Add_Dynamic_Lottage() function again!',TIMEOUT=2500,TYP='warn')
        return
    for BC in range(1,MaxCodeLength_+1):
        bitcoding = str(BC)
        m += ['Bin_indicator_'+bitcoding]
    
    df['Lottage']= df[m].sum(axis=1)
    Lottage = df['Lottage'].values.reshape(-1,1)
    scale = MinMaxScaler(-1,1)    # MinMaxScaler(feature_range = (0, 1))
    Scaledlot = scale.fit_transform(Lottage)
    df['Lottage']  = pd.DataFrame(Scaledlot , index=df.index)
    df['absLottage'] = df['Lottage'].apply(abs)
    Lottage = 'Dynamic'
    Lot_label.config(text='Dynamic Lottage Selected')
       
def LotSel():
    if Lotvar.get() == 1:
        Add_Dynamic_Lottage()
    elif Lotvar.get() == 0:
        Set_Static_Lottage()
 
def Set_Static_Lottage():
    global df
    global Lottage
    df['Lottage'] = 1
    Lottage = 'Static'
    Lot_label.config(text='Static Lottage Selected')
#◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄► Start of Trading app ◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►
from PyQt5.QtWidgets import QVBoxLayout, QTabWidget, QLabel, QLineEdit, QPushButton, QMessageBox    
class TradingBot(QWidget):
    def __init__(self):
        super().__init__()
        # Initialize the user interface
        self.initUI()

    def initUI(self):
        # Set the window title
        self.setWindowTitle('Trading Bot')

        # Create a vertical box layout
        layout = QVBoxLayout()
        
        # Create a tab widget
        self.tabs = QTabWidget()
        layout.addWidget(self.tabs)

        # Create tabs for login and trading
        self.login_tab = QWidget()
        self.trade_tab = QWidget()

        # Add tabs to the tab widget
        self.tabs.addTab(self.login_tab, 'Login')
        self.tabs.addTab(self.trade_tab, 'Trade')

        # Initialize the content for the login and trade tabs
        self.initLoginTab()
        self.initTradeTab()

        # Set the layout for the main widget
        self.setLayout(layout)

    def initLoginTab(self):
        # Create a vertical box layout for the login tab
        layout = QVBoxLayout()

        # Create labels and input fields for login
        self.login_label = QLabel('Login to MetaTrader 5')
        self.account_label = QLabel('Account:')
        self.password_label = QLabel('Password:')
        self.server_label = QLabel('Server:')

        self.account_input = QLineEdit()
        self.password_input = QLineEdit()
        self.server_input = QLineEdit()
        self.password_input.setEchoMode(QLineEdit.Password)  # Set password input to be masked

        # Create login button and connect it to the login function
        self.login_button = QPushButton('Login')
        self.login_button.clicked.connect(self.login)

        # Add widgets to the layout
        layout.addWidget(self.login_label)
        layout.addWidget(self.account_label)
        layout.addWidget(self.account_input)
        layout.addWidget(self.password_label)
        layout.addWidget(self.password_input)
        layout.addWidget(self.server_label)
        layout.addWidget(self.server_input)
        layout.addWidget(self.login_button)

        # Set the layout for the login tab
        self.login_tab.setLayout(layout)

    def initTradeTab(self):
        # Create a vertical box layout for the trade tab
        layout = QVBoxLayout()

        # Create label and buttons for trading
        self.trade_label = QLabel('Execute Trades')
        self.buy_button = QPushButton('Buy')
        self.sell_button = QPushButton('Sell')

        # Connect buttons to their respective functions
        self.buy_button.clicked.connect(self.buy)
        self.sell_button.clicked.connect(self.sell)

        # Add widgets to the layout
        layout.addWidget(self.trade_label)
        layout.addWidget(self.buy_button)
        layout.addWidget(self.sell_button)

        # Set the layout for the trade tab
        self.trade_tab.setLayout(layout)

    def login(self):
        # Retrieve login information
        account = self.account_input.text()
        password = self.password_input.text()
        server = self.server_input.text()

        # Attempt to initialize the connection to MetaTrader 5
        if not mt5.initialize(login=int(account), password=password, server=server):
            QMessageBox.critical(self, 'Login Failed', 'Failed to connect to MetaTrader 5')
        else:
            QMessageBox.information(self, 'Login Successful', 'Welcome to MetaTrader 5')
            self.tabs.setCurrentIndex(1)  # Switch to the trade tab upon successful login

    def buy(self):
        symbol = "EURUSD"
        lot = 0.1

        # Check symbol information
        symbol_info = mt5.symbol_info(symbol)
        if symbol_info is None:
            QMessageBox.critical(self, 'Symbol Error', f'Symbol {symbol} not found')
            return

        # Ensure the symbol is visible
        if not symbol_info.visible:
            if not mt5.symbol_select(symbol, True):
                QMessageBox.critical(self, 'Symbol Error', f'Failed to select symbol {symbol}')
                return

        # Get the price and point for the symbol
        point = symbol_info.point
        price = mt5.symbol_info_tick(symbol).ask

        # Check if the price retrieval was successful
        if price is None:
            QMessageBox.critical(self, 'Price Error', 'Failed to get the price for the symbol')
            return

        deviation = 20
        request = {
            "action": mt5.TRADE_ACTION_DEAL,
            "symbol": symbol,
            "volume": lot,
            "type": mt5.ORDER_TYPE_BUY,
            "price": price,
            "sl": price - 100 * point,
            "tp": price + 100 * point,
            "deviation": deviation,
            "magic": 234000,
            "comment": 'Buy Order sent from Python',
            "type_time": mt5.ORDER_TIME_GTC,
            "type_filling": mt5.ORDER_FILLING_FOK,  # Order filling type
        }
        result = mt5.order_send(request)

        # Check if the order was successful
        if result.retcode != mt5.TRADE_RETCODE_DONE:
            error_msg = (
                f"Failed to execute buy order:\n"
                f"Retcode: {result.retcode}\n"
                f"Comment: {result.comment}\n"
                f"Request: {request}\n"
                f"Result: {result}"
            )
            QMessageBox.critical(self, 'Trade Failed', error_msg)
        else:
            QMessageBox.information(self, 'Trade Successful', 'Buy order executed successfully')

    def sell(self):
        symbol = "EURUSD"
        lot = 0.1

        # Check symbol information
        symbol_info = mt5.symbol_info(symbol)
        if symbol_info is None:
            QMessageBox.critical(self, 'Symbol Error', f'Symbol {symbol} not found')
            return

        # Ensure the symbol is visible
        if not symbol_info.visible:
            if not mt5.symbol_select(symbol, True):
                QMessageBox.critical(self, 'Symbol Error', f'Failed to select symbol {symbol}')
                return

        # Get the price and point for the symbol
        point = symbol_info.point
        price = mt5.symbol_info_tick(symbol).bid

        # Check if the price retrieval was successful
        if price is None:
            QMessageBox.critical(self, 'Price Error', 'Failed to get the price for the symbol')
            return
        deviation = 20
        request = {
            "action": mt5.TRADE_ACTION_DEAL,
            "symbol": symbol,
            "volume": lot,
            "type": mt5.ORDER_TYPE_SELL,
            "price": price,
            "sl": price + 100 * point,
            "tp": price - 100 * point,
            "deviation": deviation,
            "magic": 234000,
            "comment": "python script open",
            "type_time": mt5.ORDER_TIME_GTC,
            "type_filling": mt5.ORDER_FILLING_FOK,  # Order filling type
        }
        result = mt5.order_send(request)

        # Check if the order was successful
        if result.retcode != mt5.TRADE_RETCODE_DONE:
            error_msg = (
                f"Failed to execute sell order:\n"
                f"Retcode: {result.retcode}\n"
                f"Comment: {result.comment}\n"
                f"Request: {request}\n"
                f"Result: {result}"
            )
            QMessageBox.critical(self, 'Trade Failed', error_msg)
        else:
            QMessageBox.information(self, 'Trade Successful', 'Sell order executed successfully')

#◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄► Start of main program  ◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►
ALARM.Alarm('Treble')
root = tk.Tk()
root.title(__file__[:-3])   # root.title("FOREX Trading Meta_model GUI"+argv[0][-6:-3])
root.iconbitmap(r'.\pictures\Forexicon.ico')
root.resizable(False, False)
pho = tk.PhotoImage(file=r".\pictures\FOREX2.png")
w = pho.width()
h = pho.height()
root.geometry('%dx%d+0+0'%(w,h))
tk.Label(root,image= pho).place(height=h,width=w,x=0,y=0)#.grid(rowspan=8,columnspan=2)
#tk.Label(root, text ="FOREX Trading Meta_model GUI",font=('Times', 20, 'bold'),bg='#091A32',fg='yellow').place(x=w//3,y=0)    
#root.wm_attributes('-transparentcolor', root['bg'])
time_label = tk.Label(root, font=("Helvetica", 15),bg='#15151E',fg='yellow')
time_label.place(x=w//4,y=h-100)

menubar = tk.Menu(root)
root.config(menu=menubar)
file_menu = tk.Menu(menubar, tearoff=0)
edit_menu = tk.Menu(menubar, tearoff=0)
Set_menu  = tk.Menu(menubar, tearoff=0)
DB_menu   = tk.Menu(menubar, tearoff=0)
Run_menu  = tk.Menu(menubar, tearoff=0)
Debug_menu  = tk.Menu(menubar, tearoff=0)
Report_menu = tk.Menu(menubar, tearoff=0)
Visu_menu = tk.Menu(menubar, tearoff=0)
help_menu = tk.Menu(menubar, tearoff=0)

menubar.add_cascade(label='File', menu=file_menu)
menubar.add_cascade(label='Edit', menu=edit_menu)
menubar.add_cascade(label='Setting', menu=Set_menu)
menubar.add_cascade(label='Database', menu=DB_menu)
menubar.add_cascade(label='Run', menu=Run_menu)
menubar.add_cascade(label='Debug', menu=Debug_menu)
menubar.add_cascade(label='Report', menu=Report_menu)
menubar.add_cascade(label='Visualize', menu=Visu_menu)
menubar.add_cascade(label='Help', menu=help_menu)

file_menu.add_command(label="Open Database",command = OpenDatabase)
file_menu.add_separator()

sub_menu4 = tk.Menu(DB_menu, tearoff=0)
sub_menu4.add_command(label='Load Offline DataBase',command=load_offline_Database )
sub_menu4.add_command(label='Load Online DataBase',command=lambda:LoadDataBase(''))
file_menu.add_cascade(label="Load DataBase",menu=sub_menu4)
file_menu.add_separator()

file_menu.add_command(label="Copy Selected Settings",command = CopySettings)
file_menu.add_separator()
file_menu.add_command(label="Load Selected Settings",command = LoadSettings)
file_menu.add_separator()

sub_menu9 = tk.Menu(file_menu, tearoff=0)
sub_menu9.add_command(label='Save df',command=lambda:save_file('df',df))
sub_menu9.add_separator()
sub_menu9.add_command(label='Save Features',command=lambda:save_file('Features',features))
sub_menu9.add_separator()
file_menu.add_cascade(label="Save",menu=sub_menu9)

file_menu.add_separator()
file_menu.add_command(label="Submit", command=SUBMIT)
file_menu.add_separator()
file_menu.add_command(label="Show in Box", command=insert_element)
file_menu.add_separator()
file_menu.add_command(label="Clear the Box", command=lambda:clear_box)
file_menu.add_separator()
file_menu.add_command(label="Exit",command = lambda:EXIT_One(''))

edit_menu.add_command(label="Cut")
edit_menu.add_command(label="Copy")
edit_menu.add_command(label="Paste")

Set_menu.add_command(label="Max. rows & columns",command = SETTING1)
Set_menu.add_separator()
Set_menu.add_command(label="Copy Selected Settings",command = CopySettings)
Set_menu.add_separator()
Set_menu.add_command(label="Load Selected Settings",command = LoadSettings)
Set_menu.add_separator()
# add a submenu
sub_menu2 = tk.Menu(Set_menu, tearoff=0)
sub_menu2.add_command(label='Load settings',command = LoadSettings)
sub_menu2.add_command(label='Default settings',command=DEFAULTSETTINGS)
# add the File menu to the menubar
Set_menu.add_cascade(label="General Settings",menu=sub_menu2)
Set_menu.add_separator()
sub_menu3 = tk.Menu(Set_menu, tearoff=0)
sub_menu3.add_command(label='Tehran',command = lambda:update_time('Tehran'))
sub_menu3.add_command(label='London',command = lambda:update_time('London'))
sub_menu3.add_command(label='Tokyo',command = lambda:update_time('Tokyo'))
sub_menu3.add_command(label='Sydney',command = lambda:update_time('Sydney'))
sub_menu3.add_command(label='Newyork',command = lambda:update_time('Newyork'))
Set_menu.add_cascade(label="Timezone",menu = sub_menu3)
#============================================= add a submenu
sub_menu4 = tk.Menu(DB_menu, tearoff=0)
sub_menu4.add_command(label='Load Offline DataBase',command=load_offline_Database )
sub_menu4.add_command(label='Load Online DataBase',command=lambda:LoadDataBase(''))
DB_menu.add_cascade(label="Load DataBase",menu=sub_menu4)
DB_menu.add_separator()
#============================================= add a submenu
sub_menu5 = tk.Menu(DB_menu, tearoff=0)
sub_menu5.add_command(label='Show DataBase df',font=('Times', FONTSIZE, 'bold'),command=Show_DB)
sub_menu5.add_separator() 
sub_menu5.add_command(label='Show whole DataBase',command = Show_DataBase)
sub_menu5.add_separator() 
sub_menu5.add_command(label='Show Features',font=('Times', FONTSIZE, 'bold'),command=Show_Features)
sub_menu5.add_separator()
sub_menu5.add_command(label='Show Label',command=Show_Label)
sub_menu5.add_separator()
sub_menu5.add_command(label='Show df Columns',command=lambda:Show_Columns(df))
sub_menu5.add_separator()
sub_menu5.add_command(label='Show features Columns',command=lambda:Show_Columns(features))
DB_menu.add_cascade(label="Show DataBase",font=('Times', FONTSIZE, 'bold'),menu=sub_menu5)
#============================================= add a submenu
sub_menu9.add_command(label='Save df',command=lambda:save_file('df',df))
sub_menu9.add_separator()
sub_menu9.add_command(label='Save Features',command=lambda:save_file('Features',features))
sub_menu9.add_separator()
DB_menu.add_cascade(label='Save DataBase',menu=sub_menu9)
#============================================= add a submenu
sub_menu1 = tk.Menu(DB_menu, tearoff=0)
sub_menu1.add_command(label='Just OHLC as features',command = Set_OHLC_Features)
sub_menu1.add_separator()
sub_menu1.add_command(label='Just previous Closes',command = lambda:Just_Close_Features(0))
sub_menu1.add_separator()
sub_menu1.add_command(label='Just previous Labels',command = lambda:Just_Label_Features(0))
sub_menu1.add_separator()
sub_menu1.add_command(label='All features',command = lambda:All_Features(False))
sub_menu1.add_separator()
sub_menu1.add_command(label='All features except labels',command = All_Features_Except_Labels)
sub_menu1.add_separator()
sub_menu1.add_command(label='Custom features',command=Custom_Features)
sub_menu1.add_separator()
sub_menu1.add_command(label='Custom window of features',command=Custom_Window_Features)
sub_menu1.add_separator()
sub_menu1.add_command(label='Add Binary indicators',command = Add_Binary_indicators)
sub_menu1.add_separator()
sub_menu1.add_command(label='Just Binary indicators',command = Just_Binary_indicators)
sub_menu1.add_separator()
sub_menu1.add_command(label='Selected Binary indicators',command=Selected_Binary_indicators)
sub_menu1.add_separator()
sub_menu1.add_command(label='Just mean Delta',command=Just_Mean_Delta)
sub_menu1.add_separator()
sub_menu1.add_command(label='Just Mean Return Indicators',command=Just_Mean_Return_indicators)
DB_menu.add_cascade(label="Select features",menu=sub_menu1)
DB_menu.add_separator()
#============================================= add a submenu
sub_menu6 = tk.Menu(DB_menu, tearoff=0)
sub_menu6.add_command(label='MARKET',command = lambda:Set_Label('MARKET'))
sub_menu6.add_command(label='Rise_fall',command = lambda:Set_Label('Rise_fall'))
sub_menu6.add_command(label='CLOSE',command = lambda:Set_Label('CLOSE'))
sub_menu6.add_command(label='ScaledReturn',command = lambda:Set_Label('ScaledReturn'))
DB_menu.add_cascade(label="Select Label",menu=sub_menu6)
DB_menu.add_separator()
# add a submenu
DB_menu.add_command(label='Automatic preprocessing',command=Auto_PrePro)
DB_menu.add_separator()    
#============================================= add a submenu
sub_menu6 = tk.Menu(DB_menu, tearoff=0)
sub_menu6.add_command(label='ModifyDate',command = Modify_Date)
sub_menu6.add_command(label='Delete Except OHLC',command = HoldOHLC)
sub_menu6.add_command(label='DropNAN',command = lambda:DropNAN)
sub_menu6.add_command(label='Scaling',command = Select_Scaling_Type)
DB_menu.add_cascade(label="Manual Preprocessing",menu = sub_menu6)
DB_menu.add_separator()
#============================================= add a submenu
sub_menu7 = tk.Menu(DB_menu, tearoff=0)
sub_menu7.add_command(label='Add custom window SMA',command=lambda:ADD_SMAs(0))
sub_menu7.add_separator()
sub_menu7.add_command(label='Add custom window EMA',command=lambda:ADD_EMAs(0))
sub_menu7.add_separator()
sub_menu7.add_command(label='Add custom window RSI',command=lambda:Add_RSI(0))
sub_menu7.add_separator()
sub_menu7.add_command(label='Add MACD',command=Add_MACD)
sub_menu7.add_separator()
sub_menu7.add_command(label='Add custom window bollinger Bands',command=lambda:Add_bollingerBands(0))
sub_menu7.add_separator()
sub_menu7.add_command(label='Add ATR',command=lambda:Add_ATR(0))
sub_menu7.add_separator()
sub_menu7.add_command(label='Add Stochastic',command=lambda:Add_STOCH(0))
sub_menu7.add_separator()
sub_menu7.add_command(label='Add Chaikin Money Flow',command=lambda:Add_CMF(0))
sub_menu7.add_separator()
sub_menu7.add_command(label='Add williams_r indic.',command=lambda:Add_williams_r(0))
sub_menu7.add_separator()
sub_menu7.add_command(label='ADD CCI indic.',command=lambda:Add_CCI(0))
sub_menu7.add_separator()
sub_menu7.add_command(label='Add SAR indic.',command=Add_SAR)
sub_menu7.add_separator()
sub_menu7.add_command(label='Add ichimoku',command=Add_ichimoku)
sub_menu7.add_separator()
sub_menu7.add_command(label='Add ADX',command=Add_ADX)
sub_menu7.add_separator()
DB_menu.add_cascade(label="Add classic indicators & oscillators",menu=sub_menu7)
DB_menu.add_separator()  
#============================================= add a submenu
sub_menu8 = tk.Menu(DB_menu, tearoff=0)
sub_menu8.add_command(label='Add ScaledReturn',command=Add_ScaledReturn)
sub_menu8.add_separator()
sub_menu8.add_command(label='Add previous closes',command=lambda:Add_previous('CLOSE'))
sub_menu8.add_separator()
sub_menu8.add_command(label='Add previous Labels',command=lambda:Add_previous('LABEL'))
DB_menu.add_cascade(label="Add Extra Features",menu=sub_menu8)
DB_menu.add_separator()
#============================================= add a submenu
sub_menu10 = tk.Menu(DB_menu, tearoff=0)
sub_menu10.add_command(label='Build Binary Database',command=Build_Binary_Database)
sub_menu10.add_separator()
sub_menu10.add_command(label='Add Binary Indicators',command=Add_Binary_indicators)
sub_menu10.add_separator()
sub_menu10.add_command(label='Just Binary Indicators',command=Just_Binary_indicators)
sub_menu10.add_separator()
sub_menu10.add_command(label='Just mean Delta',command=Just_Mean_Delta)
sub_menu10.add_separator()
sub_menu10.add_command(label='Add Selected Binary Indicators',command=Selected_Binary_indicators)
sub_menu10.add_separator()
sub_menu10.add_command(label='Just Mean Return Indicators',command=Just_Mean_Return_indicators)
DB_menu.add_cascade(label="Binary Features",font=('Times', FONTSIZE, 'bold'),menu=sub_menu10)
DB_menu.add_separator()
#============================================= add a submenu
Run_menu.add_command(label="Add Dynamic Lottage",command=Add_Dynamic_Lottage)
Run_menu.add_separator()
# add a submenu
sub_menu10 = tk.Menu(Run_menu, tearoff=0)
sub_menu10.add_command(label='Gaussian Naïve Bayes Classifier',command=Gaussian_NB)
sub_menu10.add_separator()
sub_menu10.add_command(label='ML_Perceptron Classifier',command=MLPC)
sub_menu10.add_separator()
sub_menu10.add_command(label='Logistic Regression',command=Log_Reg)
sub_menu10.add_separator()
sub_menu10.add_command(label='Decision Tree Classifier',command=DTc)
sub_menu10.add_separator()
sub_menu10.add_command(label='Random Forest Classifier',command=RFc)
sub_menu10.add_separator()
sub_menu10.add_command(label='KNN Classifier',command=knn)
sub_menu10.add_separator()
sub_menu10.add_command(label='Support Vector Machine Classifier',command=svc)
sub_menu10.add_separator()
sub_menu10.add_command(label='Gradient Boosting Classifier',command=GBC)
sub_menu10.add_separator()
sub_menu10.add_command(label='X Gradient Boosting Classifier',command=XGBC)
sub_menu10.add_separator()
sub_menu10.add_command(label='Light Gradient Boosting Machine Classifier',command=LGBMC)
sub_menu10.add_separator()
sub_menu10.add_command(label='CatBoost Classifier',command=CatBC)
sub_menu10.add_separator()
sub_menu10.add_command(label='AdaBoost Classifier',command=AdaBC)
sub_menu10.add_separator()
sub_menu10.add_command(label='Extra Trees Classifier',command=ExtraTreesC)
sub_menu10.add_separator()
sub_menu10.add_command(label='Linear Discriminant Analysis',command=LDa)
sub_menu10.add_separator()
Run_menu.add_cascade(label="Base ML Classifiers",menu=sub_menu10)
Run_menu.add_separator()
#============================================= add a submenu
sub_menu11 = tk.Menu(Run_menu, tearoff=0)
sub_menu11.add_command(label='LSTM Classifier',command=LSTMC)
sub_menu11.add_separator()
sub_menu11.add_command(label='Deep Neural Network Classifier',command=DeepNN)
sub_menu11.add_separator()
Run_menu.add_cascade(label="DEEP ML Classifiers",menu=sub_menu11)
Run_menu.add_separator()
#============================================= add a submenu
sub_menu12 = tk.Menu(Run_menu, tearoff=0)
sub_menu12.add_command(label='Linear Regressor',command=Lin_Reg)
sub_menu12.add_separator()
sub_menu12.add_command(label='Decision Tree Regressor',command=DT_Reg)
sub_menu12.add_separator()
sub_menu12.add_command(label='Random Forest Regressor',command=RF_Reg)
sub_menu12.add_separator()
sub_menu12.add_command(label='Support Vector Regressor',command=SV_Reg)
sub_menu12.add_separator()
Run_menu.add_cascade(label="Base ML Regressors",menu=sub_menu12)
Run_menu.add_separator()
#============================================= add a submenu
sub_menu13 = tk.Menu(Run_menu, tearoff=0)
sub_menu13.add_command(label='Simple Sum Mean Return Binary',command=SSMRB )
sub_menu13.add_separator()
sub_menu13.add_command(label='Simple Sum Mean Delta Binary',command=SSMDB )
# add the File menu to the menubar
Run_menu.add_cascade(label="Binary Models",menu=sub_menu13)
Run_menu.add_separator()

#============================================= add a submenu
sub_menu14 = tk.Menu(Run_menu, tearoff=0)
sub_menu14.add_command(label='Simple Sum Mean Return Binary',command=SSMRB )
sub_menu14.add_separator()
sub_menu14.add_command(label='Greedy Strom Ensemble',command=GSE_enable )
sub_menu14.add_separator()
sub_menu14.add_command(label='Homogeneous Ensemble',command=Homogeneous_enable)
sub_menu14.add_separator()
Run_menu.add_cascade(label="Novel Strategies",menu=sub_menu14)
Run_menu.add_separator()
#============================================= add a submenu
sub_menu5 = tk.Menu(Run_menu, tearoff=0)
Run_menu.add_command(label='Do Instant Trade',command = Instant_Trade)
Run_menu.add_separator()
Run_menu.add_command(label='Do Instant Trade with PYQT',command = Instant_Trade_PYQT)
Run_menu.add_separator()
sub_menu5.add_command(label='instant Buy' ,command = Instant_Buy)
sub_menu5.add_separator()
sub_menu5.add_command(label='instant Sell',command = Instant_Sell)
Run_menu.add_cascade(label="Instant Trade",menu=sub_menu5)
Run_menu.add_separator()
#============================================= add a submenu
Debug_menu.add_command(label='Activate shell',command = Activate_shell)
Debug_menu.add_separator()
Debug_menu.add_command(label='Catch variable',command = CatchVariable)
Debug_menu.add_separator()
Debug_menu.add_command(label='Catch function',command = CatchFunction)
Debug_menu.add_separator()
#============================================= add a submenu
sub_menu4 = tk.Menu(Report_menu, tearoff=0)
Report_menu.add_command(label='Open positions status',command = verbose_status)
Report_menu.add_separator()
Report_menu.add_command(label='Open positions symbols',command = open_position_symbols)
Report_menu.add_separator()
sub_menu4.add_command(label='Dickey fuller test' ,command = adf_test)
sub_menu4.add_command(label='Kwiatkowski-Phillips-Schmidt-Shin (KPSS) Test',command = kpss_test)
Report_menu.add_cascade(label="Statistic tests",menu=sub_menu4)
Report_menu.add_separator()
#============================================= add a submenu
sub_menu9 = tk.Menu(Visu_menu, tearoff=0)
sub_menu10 = tk.Menu(Visu_menu, tearoff=0)
sub_menu11 = tk.Menu(Visu_menu, tearoff=0)
sub_menu15 = tk.Menu(Visu_menu, tearoff=0)

sub_menu9.add_command(label='Close Graph',command = lambda:Graph_CLOS(True))
sub_menu9.add_separator()
sub_menu9.add_command(label='Log Price | Return | Histogram',command = lambda:LogPrice_Return_Hist(True))
sub_menu9.add_separator()
sub_menu9.add_command(label='Autocorr. Graph',command = lambda:Graph_Autocorr(True))
sub_menu9.add_separator()
sub_menu9.add_command(label='features Correlation Graph',command = lambda:Graph_Feat_Corr(True))
sub_menu9.add_separator()
sub_menu9.add_command(label='Select the Graph',command = Selectable_Graph)
Visu_menu.add_cascade(label="General Visualize",menu=sub_menu9)
Visu_menu.add_separator()
#============================================= add a submenu
sub_menu10.add_command(label='Confusion Matrix' ,command = lambda:Confusion_Matrix_Plot(True))
sub_menu10.add_separator()
sub_menu10.add_command(label='Equity and Win/Loss Plot' ,command = lambda:Equity_Plot(True))
sub_menu10.add_separator()
sub_menu10.add_command(label='DrawDown Plot' ,command = lambda:DrawDown_Plot(True))
sub_menu10.add_separator()
sub_menu10.add_command(label='Histogram Plot' ,command = lambda:Histogram_Plot(True))
Visu_menu.add_cascade(label="After train Visualization",menu=sub_menu10)
Visu_menu.add_separator()
sub_menu11.add_command(label='Scatter plot',command = lambda:Scatter_Plot(True))

Visu_menu.add_cascade(label='Regression based visualize',menu=sub_menu11)
Visu_menu.add_separator()
sub_menu15.add_command(label = 'features with Different No. of lag Closes',command = Compare_Closes)
sub_menu15.add_separator()
sub_menu15.add_command(label='features with Different No. of lag Labels',command = Compare_Labels)
Visu_menu.add_cascade(label='Comparison for features',menu=sub_menu15)
Visu_menu.add_separator()
Visu_menu.add_command(label="Copy Selected Comparison",command = CopyCompare)
Visu_menu.add_separator()
Visu_menu.add_command(label="Load Selected Comparison",command = LoadCompare)
Visu_menu.add_separator()

help_menu.add_command(label="Process Summarize",command=Process_Summarize)
help_menu.add_command(label="help",command=HELP)
help_menu.add_command(label="about",command=ABOUT)

#◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄► End of menu construction  ◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►
      
lbl1 = tk.Label(anchor=tk.CENTER,text="Select audio interaction",font=('Times', FONTSIZE),bg='#091A32',fg='white',relief="raised")
lbl1.place(width= 200,height=20,x=180,y=50)
lbl2 = tk.Button(anchor=tk.CENTER,text="Not Logged in ",font=('Times', FONTSIZE),bg='#091A32',fg='white',relief="ridge",command=Login_To_MT5)
lbl2.place(width= 200,height=20,x=180,y=80)

intChoice1 = tk.IntVar()
intChoice2 = tk.IntVar()
intChoice3 = tk.IntVar()
intChoice4 = tk.IntVar()
intChoice5 = tk.IntVar()
intChoice6 = tk.IntVar()
intChoice7 = tk.IntVar()
intChoice8 = tk.IntVar()
intChoice8.set(1)
intChoice9 = tk.IntVar()
chkChoice1 = tk.Checkbutton(root,anchor=tk.W, text="Disable AUI"  , variable=intChoice1,font=('Times', FONTSIZE),bg='#091A32',fg='gold',selectcolor= 'black',activeforeground='black', command= choice1)
chkChoice1.place(height=20,x=20,y=50)
chkChoice2 = tk.Checkbutton(root,anchor=tk.W,text="Login to MT5", variable=intChoice2,font=('Times', FONTSIZE),bg='#091A32',fg='gold',selectcolor= 'black', command= choice2)
chkChoice2.place(height=20,x=20,y=80)
chkChoice3 = tk.Checkbutton(root,anchor=tk.W, text="Novel Storm Split", variable=intChoice3,font=('Times', FONTSIZE,'bold'),bg='#091A32',fg='#F5D33B',selectcolor= 'black',activeforeground='black', command= Novelsplitting)
chkChoice3.place(x=1000,y=385)
chkChoice6 = tk.Checkbutton(root,anchor=tk.W, text="Normal Split", variable=intChoice6,font=('Times', FONTSIZE,'bold'),bg='#091A32',fg='#F5D33B',selectcolor= 'black',activeforeground='black', command= Normalsplitting)
chkChoice6.place(x=1200,y=385)
chkChoice4 = tk.Checkbutton(root,anchor=tk.W, text="Normal Run", variable=intChoice4,font=('Times', FONTSIZE,'bold'),bg='#091A32',fg='#F5D33B',selectcolor= 'black',activeforeground='black', command= Normalrunchoice)
chkChoice4.place(x=1000,y=420)
chkChoice5 = tk.Checkbutton(root,anchor=tk.W, text="Quick Run", variable=intChoice5,font=('Times', FONTSIZE,'bold'),bg='#091A32',fg='#F5D33B',selectcolor= 'black',activeforeground='black', command= Quickchoice)
chkChoice5.place(x=1200,y=420)

chkChoice7 = tk.Checkbutton(root,anchor=tk.W, text="Complete report", variable=intChoice7,font=('Times', FONTSIZE,'bold'),bg='#091A32',fg='#F5D33B',selectcolor= 'black',activeforeground='black', command= ConciseReport)
chkChoice7.place(x=780,y=480)

chkChoice8 = tk.Checkbutton(root,anchor=tk.W, text="Manual     Settings", variable=intChoice8, font=('Times', FONTSIZE,'bold'),bg='#091A32',fg='#F5D33B',selectcolor= 'black',activeforeground='black', command = lambda:QuickinitialSettings(0))
chkChoice8.place(x=740,y=313)
chkChoice9 = tk.Checkbutton(root,anchor=tk.W, text="Automatic  Settings", variable=intChoice9, font=('Times', FONTSIZE,'bold'),bg='#091A32',fg='#F5D33B',selectcolor= 'black',activeforeground='black', command = lambda:QuickinitialSettings(1))
chkChoice9.place(x=740,y=343)

CURBtn = tk.Button(root,bd =2, text="Currency Pair",justify='left',font=('Times', FONTSIZE, 'bold'),bg='#091A32',fg='gold',activeforeground='black',command = lambda:fillcurpair(''))
CURBtn.place(width=150,height=20,x=20,y=110)
Sdatebutton = tk.Button(root, text="Start date",justify='left',font=('Times', FONTSIZE, 'bold'),bg='#091A32',fg='gold',command=lambda:S_OUT(True))
Sdatebutton.place(width=150,height=20,x=20,y=140)
tk.Button(root, text="End date",font=('Times', FONTSIZE, 'bold'),bg='#091A32',fg='gold',command=End_Date).place(width=150,height=20,x=20,y=170)

Account_label = tk.Label(root, text='Not logged in yet!',font=("Helvetica", FONTSIZE) ,bg='#091A32',fg='Red',anchor = 'w',justify=tk.LEFT)
Account_label.place(x=380, y=115)

DataBaseBTN = tk.Button(root,bd =5, text="Load\nDataBase",justify='center',font=('Times', FONTSIZE),bg='#091A32',fg='white',command = lambda:LoadDataBase(''))
DataBaseBTN.place(height=70,width=200,x=410,y=50)

ShowdfBTN = tk.Button(root,bd =5, text="Show\nFeatures",justify='center',font=('Times', FONTSIZE),bg='#091A32',fg='white',command = Show_Features)
ShowdfBTN.place(height=65,width=140,x=540,y=200)


ShowdfBTN = tk.Button(root,bd =5, text="Show df",justify='center',font=('Times', FONTSIZE),bg='#091A32',fg='white',command = Show_DB)
ShowdfBTN.place(height=65,width=140,x=540,y=265)

SavedfBTN = tk.Button(root,bd =5, text="Save df",justify='center',font=('Times', FONTSIZE),bg='#091A32',fg='white',command = lambda:save_file('df',df))
SavedfBTN.place(height=65,width=140,x=540,y=330)

CopySetBTN = tk.Button(root,bd =5, text="Copy \nPreset Data",justify='center',font=('Times', FONTSIZE),bg='#091A32',fg='white',command = CopySettings)
CopySetBTN.place(height=65,width=140,x=540,y=395)

PresetBTN = tk.Button(root,bd =5, text="Load \nPreset Data",justify='center',font=('Times', FONTSIZE),bg='#091A32',fg='white',command = LoadSettings)
PresetBTN.place(height=65,width=140,x=540,y=460)

# StCMPBTN = tk.Button(root, text="window Comparison",justify='center',font=('Tahoma', FONTSIZE, 'bold'),bd=8,fg='#F5D33B', bg='#05051E',activebackground='dark green',activeforeground='gold',command = Compare_Labels)
# StCMPBTN.place(height=FONTSIZE*4,width=FONTSIZE*40,x=1000,y=605)

SHOWCMPLDBTN = tk.Button(root,bd =5, text="Visualize Comparison",justify='center',font=('Times', FONTSIZE),bg='#091A32',fg='gold',command = lambda:Select_Comparison_Graph(tempindex,temp1,temp2))
SHOWCMPLDBTN.place(height=FONTSIZE*4,width=FONTSIZE*25,x=990,y=715)

SHOWTBLBTN = tk.Button(root,bd =5, text="Show Table",justify='center',font=('Times', FONTSIZE),bg='#091A32',fg='gold',command = OpenGenaratedTable)
SHOWTBLBTN.place(height=FONTSIZE*4,width=FONTSIZE*14,x=1250,y=715)

CopyCmpBTN = tk.Button(root,bd =5, text="Copy Comparing Data",justify='center',font=('Times', FONTSIZE),bg='#091A32',fg='gold',command = CopyCompare)
CopyCmpBTN.place(height=35,width=200,x=990,y=680)

LDCMPBTN = tk.Button(root,bd =5, text="Load Comparing Data",justify='center',font=('Times', FONTSIZE),bg='#091A32',fg='gold',command = LoadCompare)
LDCMPBTN.place(height=35,width=200,x=1190,y=680)
#////////////////////////////////////
VISUALLIST = ['General Visualization','After run Visualization'] 
var1 = tk.Variable(value=VISUALLIST)
VISlstbox = tk.Listbox(root, width=20, height=5,listvariable = var1,font=('Times', FONTSIZE),selectmode = tk.MULTIPLE)
VISBTN = tk.Button(root, text='Save Visualizations',bd = 5,font=('Times', FONTSIZE,'bold'), bg='#F5D33B',command=selected_Visual)
VISBTN.place(height=FONTSIZE2*2,width=FONTSIZE*19,x=740,y=400)
VISlstbox.place(height=FONTSIZE*5,width=FONTSIZE*19,x=740,y=427)

RestartENV = tk.Button(root,bd =5, text="Restart\nVariables",justify='center',font=('Times', int(FONTSIZE*0.7)),bg='#091A32',fg='white',command = Total_Reset)
RestartENV.place(x=1100,y=25)

RestartGUI = tk.Button(root,bd =5, text="Restart\nGUI",justify='center',font=('Times', int(FONTSIZE*0.7)),bg='#091A32',fg='white',command = RESTART3)
RestartGUI.place(x=1175,y=25)

CLOSEALLBTN = tk.Button(root,bd =5, text="Close\nWindows",justify='center',font=('Times', int(FONTSIZE*0.7)),bg='#091A32',fg='white',command = EXIT_All)
CLOSEALLBTN.place(x=1240,y=25)

CLOSEfigures = tk.Button(root,bd =5, text="Close\nFigures",justify='center',font=('Times', int(FONTSIZE*0.7)),bg='#091A32',fg='white',command = lambda:plt.close('all'))
CLOSEfigures.place(x=1315,y=25)

CLOSEexcells = tk.Button(root,bd =5, text="Close\nExcels",justify='center',font=('Times', int(FONTSIZE*0.7)),bg='#091A32',fg='white',command = CLOSEexcellFiles)
CLOSEexcells.place(x=1380,y=25)

CURPAIR = tk.Entry(root)
SDATE = tk.Entry(root)
EDATE = tk.Entry(root)
CURPAIR.place(height=20,width=180,x=180,y=110)
SDATE.place(height=20,width=180,x=180,y=140)
EDATE.place(height=20,width=180,x=180,y=170)
#CURPAIR.insert(0,'EURUSD_i')
#CURPAIR.bind('<FocusIn>',on_Click)
CURPAIR.bind('<FocusOut>',fillcurpair)
#SDATE.bind('<FocusIn>',S_CLICK)
SDATE.bind('<FocusOut>',S_OUT)
#EDATE.bind('<FocusIn>',E_CLICK)
EDATE.bind('<FocusOut>',E_OUT)
#////////////////////////////////////////////////////////
IndicOptions = ('ADD SMA indic.','ADD EMA indic.','ADD RSI Osc.','ADD MACD indic.','ADD Bol Band indic.','ADD ATR indic.','ADD Stochastic Osc.',\
                'ADD CMF indic.','ADD williams_r indic.','ADD CCI indic.','Add SAR indic.','Add ichimoku','Add ADX')
INDICVAR = tk.StringVar() 
INDICVAR.set(IndicOptions[0])
INDICVAR.trace_add('write', lambda *args:Indic_option())
        
IndicBTN = tk.Button(root,bd = 5,text='Add Indicators',font=('Tahoma', FONTSIZE, 'bold'), bg='#F5D33B',command = All_indics) 
IndicBTN.place(height=FONTSIZE*3,width=FONTSIZE*18,x=350,y=210)

Indicopt = tk.OptionMenu(root,INDICVAR,*IndicOptions)
Indicopt.place(height=FONTSIZE*3,width=FONTSIZE*18,x=350,y=238)
#/////////////////////////////////////////////////////////
    
CLBXBT = tk.Button(root,bd=5,bg = '#061A2E',fg='orange',text="Clear the box",font=('Times', FONTSIZE, 'bold'),command=clear_box)
CLBXBT.place(height=50,width=150,x=350,y=300)
SendBtn = tk.Button(root,bd=5,bg = '#061A2E',fg='orange',text="Show in box",font=('Times', FONTSIZE, 'bold'),command=insert_element)
SendBtn.place(height=50,width=150,x=350,y=355)
SUBMITBTN = tk.Button(root,bd=5,bg = '#061A2E',fg='orange', text="Get Values",font=('Times', FONTSIZE, 'bold'), command=SUBMIT)
SUBMITBTN.place(height=50,width=150,x=350,y=410)
CloseBtn = tk.Button(root,bd=5 ,bg = '#061A2E',fg='orange',text ='Close',font=('Times', FONTSIZE, 'bold'),command = lambda:EXIT_One(''))
CloseBtn.place(height=50,width=150,x=350,y=465)

root.bind('<Escape>',EXIT_One)
root.bind('<Return>',LoadDataBase)
root.bind('<l>',cmbLabFocus)
root.bind('<f>', cmbFeatFocus)
RUNBTN = tk.Button(root,bd=5, text=f"Run {Strategy}",font=('Times', FONTSIZE2, 'bold'),bg='#091A32' ,fg='white',activebackground="light yellow",activeforeground='maroon4',command=lambda:RUNStrategy('NormalMode'))
RUNBTN.place(x=540,y=535)
   
Lotvar = tk.IntVar()
LotRDbut1 = tk.Radiobutton(root,text = 'Dynamic Lottage',variable=Lotvar, value=1,font=("Helvetica", FONTSIZE),bg='#091A32',fg='#F5CC21' ,command=LotSel,indicatoron=0)
LotRDbut1.place(x=540,y=640)
LotRDbut2 = tk.Radiobutton(root,text = 'Static Lottage',variable=Lotvar, value=0,font=("Helvetica", FONTSIZE),bg='#091A32',fg='#F5CC21' ,command=LotSel,indicatoron=0)
LotRDbut2.place(x=750,y=640)
Lot_label = tk.Label(root, text='Static(uniform) Lottage',font=("Helvetica", FONTSIZE) ,bg='#091A32',fg='Gold',anchor = 'c',justify=tk.CENTER)
Lot_label.place(x=590,y=680)

#◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄► Timeframe selection combobox snippet  ◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►
categories = ['m1','m2','m3','m4','m5','m6','m10','m12','m15','m20','m30',
              'h1','h3','h4','h6','h8','h12','d1','d5','w1','mn1']
lblCategory = tk.Label(text="↓select the TimeFrame↓",font=('Times',FONTSIZE, 'bold'),bd=5)
lblCategory.place(height=30,width=FONTSIZE*31,x=20,y=210)
strCategory = tk.StringVar()
cmbCategories = tk.ttk.Combobox(root, textvariable = strCategory,font=('Tahoma', int(FONTSIZE*1.5), 'bold'))
cmbCategories.place(height=40,width=FONTSIZE*31,x=20,y=240)
cmbCategories['values'] = categories
cmbCategories.bind('<<ComboboxSelected>>', TF_SUBMIT)
cmbCategories.current(17)
TF_MESSAGE = tk.Message(root,text = f'Initial Timeframe is: {TimeFrame}',justify=tk.LEFT,width=250,bd=5,fg='#F5CC21',bg = '#08162f',font=('Times', FONTSIZE, 'bold'),anchor = 'w',aspect=200)
TF_MESSAGE.place(height=50,width=250,x=20,y=280)
#◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄► End of Timeframe selection combobox snippet  ◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►◄►

feature_Select(root,4)
Label_Select(root)

# Create a listbox and Inserting the listbox items
MODELSET = {1:  '1.  →  Gaussian Naïve Bayes',
            2:  '2.  →  Multi Layer Perceptron Classifier',
            3:  '3.  →  Logistic Regression',
            4:  '4.  →  Decision Tree Classifier',
            5:  '5.  →  Random Forest Classifier',
            6:  '6.  →  K Nearest Neighbors Classifier',
            7:  '7.  →  Support Vector Machine Classifier',
            8:  '8.  →  Gradient Boosting Classifier',
            9:  '9.  →  X Gradient Boosting Classifier',
            10: '10. →  Light Gradient Boosting Machine Classifier',
            11: '11. →  CatBoost Classifier',
            12: '12. →  AdaBoost Classifier',
            13: '13. →  Extra Trees Classifier',
            14: '14. →  Linear Discriminant Analysis',
            15: '15. →  LSTM Classifier',
            16: '16. →  Deep Neural Network Classifier',
            17: '17. →  Linear Regressor',
            18: '18. →  Decision Tree Regressor',
            19: '19. →  Random Forest Regressor',
            20: '20. →  Support Vector Regressor',
            21: '21. →  Greedy Dynamic Ensemble',
            22: '22. →  Simple Sum Mean Return Binary',
            23: '23. →  Simple Sum Mean Delta Binary'
            }
   
MoDeLs = list(MODELSET.values()) #MoDeLs = list(MODELSET[i] for i in range (1,len(MODELSET)+1))
frame1 = tk.Frame(root)  
frame1.place(height=410, width=313, x=20, y=355)
lstbox = tk.Listbox(frame1, width=40, height=30,listvariable=tk.Variable(value=MoDeLs), selectmode = tk.MULTIPLE)
lstbox.place(height=380, width=287, x=0, y=10) 
# Create the Scrollbar  
scrollbar1 = tk.Scrollbar(frame1, orient=tk.VERTICAL, command=lstbox.yview)  
scrollbar1.place(x=285, y=0, height=380)   
lstbox.config(yscrollcommand=scrollbar1.set)
scrollbar2 = tk.Scrollbar(frame1, orient=tk.HORIZONTAL, command=lstbox.xview)  
lstbox.config(xscrollcommand=scrollbar2.set)
scrollbar2.place(x=0, y=380,width=313) 
MDLBTN = tk.Button(root, text='↓Select the models for GDE↓',font=('Times', int(FONTSIZE*1.2), 'bold'),fg='yellow', bg='Navy blue',command = select_strategy)#selected_models)  
MDLBTN.place(height=40,width=313,x=20,y=323) 

tradoptions=('Instant Buy','Instant Sell')
TRADVAR = tk.StringVar()
TRADVAR.set(tradoptions[0])
TRADVAR.trace_add('write', lambda *args:trade_option())
        
tradeBTN = tk.Button(root,bd = 5,text='Instant Trade',font=('Tahoma', FONTSIZE, 'bold'), bg='#F5D33B',command=Instant_Trade)
tradeBTN.place(height=FONTSIZE*3,width=FONTSIZE*19,x=1100,y=90)

tradeopt = tk.OptionMenu(root,TRADVAR,*tradoptions)
tradeopt.place(height=FONTSIZE*3,width=FONTSIZE*19,x=1100,y=118)
#/////////////////////////////////////////////////////////
Modeloptions = []  
for M1 in str(MODELSET.values()).split('→'):
    Modeloptions += [M1.split("'")[0].strip()]
Modeloptions.pop(0)

ModelVAR = tk.StringVar()
ModelVAR.set(Modeloptions[3])

MODLBTN = tk.Button(root,text='Select Single Model',font=('Tahoma', FONTSIZE, 'bold'),bd=8,fg='#F5D33B', bg='#05051E',activebackground='dark green',activeforeground='gold',command=lambda:RUNStrategy('NormalMode'))
MODLBTN.place(height=FONTSIZE*4,width=FONTSIZE*40,x=990,y=310)
ModelVAR.trace_add('write', lambda *args:model_option())
Modelopt = tk.OptionMenu(root,ModelVAR,*Modeloptions)
Modelopt.configure(bg='#F5D33B',activebackground='dark green',activeforeground='gold',bd=5)
Modelopt.place(height=FONTSIZE*4,width=FONTSIZE*40,x=990,y= 343)
#/////////////////////////////////////////////////////////

selectmodelsoptions = [] 
for M1 in str(MODELSET.values()).split('→'):
    selectmodelsoptions += [M1.split("'")[0].strip()]
selectmodelsoptions.pop(0)
CompareVAR = tk.StringVar()
CompareVAR.set(selectmodelsoptions[0])
CmpFeatBTN = tk.Button(root,text='Features Comparison',font=('Tahoma', FONTSIZE, 'bold'),bd=8,fg='#F5D33B', bg='#05051E',activebackground='dark green',activeforeground='gold',command = CompareFeatures)
CmpFeatBTN.place(height=FONTSIZE*4,width=FONTSIZE*40,x=990,y=450)
CompareVAR.trace_add('write', lambda *args:Modelsoption())
Compareopt = tk.OptionMenu(root,CompareVAR,*selectmodelsoptions)
Compareopt.configure(bg='#F5D33B',activebackground='Navy blue',activeforeground='gold',bd=5)
Compareopt.place(height=FONTSIZE*4,width=FONTSIZE*40,x=990,y= 483)
#/////////////////////////////////////////////////////////

#selectmodelsoptions.pop(0)

CompareFEAtVAR = tk.StringVar()
CompareFEAtVAR.set(CompareFeaturessoptions[0])

FEAtBTN = tk.Button(root,text='Models Comparison',font=('Tahoma', FONTSIZE, 'bold'),bd=8,fg='#F5D33B', bg='#05051E',activebackground='dark green',activeforeground='gold',command=lambda:CompareModels(''))
FEAtBTN.place(height=FONTSIZE*4,width=FONTSIZE*40,x=990,y=525)
CompareFEAtVAR.trace_add('write', lambda *args:FeaTuRes_option())
CompareFEAtopt = tk.OptionMenu(root,CompareFEAtVAR,*CompareFeaturessoptions) 
CompareFEAtopt.configure(bg='#F5D33B',activebackground='Navy blue',activeforeground='gold',bd=5)
CompareFEAtopt.place(height=FONTSIZE*4,width=FONTSIZE*40,x=990,y= 558)
#/////////////////////////////////////////////////////////

ModelSET = {1:'Just previous Closes',2:'Just previous Labels',3:'SMA window size',4:'EMA window size'} 
CompareWINsoptions = [] 
for F1 in ModelSET.values():
    CompareWINsoptions += [F1]

CompareWINVAR = tk.StringVar()
CompareWINVAR.set(CompareWINsoptions[1])

WInBTN = tk.Button(root,text='Lag Window Comparison',font=('Tahoma', FONTSIZE, 'bold'),bd=8,fg='#F5D33B', bg='#05051E',activebackground='dark green',activeforeground='gold',command = lambda:RUNCompareLagbased(''))
WInBTN.place(height=FONTSIZE*4,width=FONTSIZE*40,x=990,y=605)
CompareWINVAR.trace_add('write', lambda *args:Lagwindow_option())
CompareWInopt = tk.OptionMenu(root,CompareWINVAR,*CompareWINsoptions)
CompareWInopt.configure(bg='#F5D33B',activebackground='dark orange',activeforeground='gold',bd=5)
CompareWInopt.place(height=FONTSIZE*4,width=FONTSIZE*40,x=990,y= 638)
#//////////////////

teststr = tk.StringVar(root)
TESTSIZE = tk.Spinbox(root, values=[v for v in range(1,1001)],font=('Times', FONTSIZE, 'bold'),bd=2,width=10,bg='#05051E' ,fg='white',textvariable=teststr, command=GETTESTSIZE)
TESTSIZE.place(x=w//2+40,y=80)

TESTSIZE.delete(0,"end")
TESTSIZE.insert(0,30)
trainstr = tk.StringVar(root)
TrainMUL = tk.Spinbox(root, values=[W for W in range(2,1000)], font=('Times', FONTSIZE, 'bold'), bd=2, width=10, bg='#05051E' ,fg='white',textvariable = trainstr, command = GETTrainSIZE)
TrainMUL.place(x=w//2+180,y=80)

TrainMUL.delete(0,"end")
TrainMUL.insert(0,10)
intChoice6.set(1)
intChoice4.set(1)
#/////////////////////////////////////////////////////////

TESTLBL = tk.Label(root, text='',font=('Times',FONTSIZE),bd=2,bg='#0E1734',fg='white')
TESTLBL.place(x=w//2+40,y=50)

TrainLBL = tk.Label(root, text='',font=('Times',FONTSIZE),bd=4,bg='#0E1734',fg='white')
TrainLBL.place(x=w//2+180,y=50)

DBSizeBTN = tk.Button(root, text=f'The size of the database utilized:{TestSize+Trainsize} candles',font=('Times',FONTSIZE),bd=2,bg='#0E1734',fg='white',command=GETTrainSIZE)
DBSizeBTN.place(x=w//2,y=110)

Fea_MESSAGE1 = tk.Message(root,text = f'Not selected yet',justify=tk.LEFT,width=200,bd=5,fg='red',bg = '#08162f',font=('Times', int(FONTSIZE*0.8), 'bold'),anchor = 'w',aspect=200)
Fea_MESSAGE1.place(x=1100,y=190)

Fea_MESSAGE2 = tk.Message(root,text = 'Features:',justify=tk.LEFT,width=200,bd=5,fg='#B5CC21',bg = '#08162f',font=('Times', int(FONTSIZE*0.8), 'bold'),anchor = 'w',aspect=200)
Fea_MESSAGE2.place(x=1100,y=165)

LB_MESSAGE1 = tk.Message(root,text = f'label: Not selected yet',justify=tk.LEFT,width=200,bd=5,fg='red',bg = '#08162f',font=('Times', int(FONTSIZE*0.8), 'bold'),anchor = 'w',aspect=200)
LB_MESSAGE1.place(height=48,width=200,x=1100,y=263)

TESTSIZE.focus()
update_time(selected_timezone)
root.bind('<Control-space>', Minimize)
root.bind('<Control-a>', OpenIndicMenu)
root.bind('<Control-A>', OpenIndicMenu)
root.bind('<c>', CompareModels)
root.bind('<C>', CompareModels)
root.bind('<Control-l>', RUNCompareLagbased)
root.bind('<Control-L>', RUNCompareLagbased)
root.bind('<F3>', QuickinitialSettings1)
root.bind('<F4>', QuickinitialSettings1)
root.bind('<F5>', QuickinitialSettings1)
root.bind('<F6>', QuickinitialSettings1)
root.bind('<F7>', QuickinitialSettings1)
root.bind('<F1>', Mute)

if not os.path.exists(DataPath):
    showmessage('Finding path',f'The path {DataPath} was not found!\nIt will be created!',TIMEOUT= 3000,TYP='warn')
    try:
        os.makedirs(DataPath)
    except Exception as ERR:
        showmessage('Path Error',f'Check the file path or name: {DataPath}\n{ERR}',TIMEOUT=3000,TYP='Erro')

root.mainloop()

