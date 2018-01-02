# -*- coding: utf-8 -*-
"""
Created on Sat Dec 16 19:48:20 2017

@author: PHAM DANG MANH HONG LUAN
"""

import sys, os
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.lines as mlines
import numpy as np
from scipy.misc import imread

menu_actions = {}


#USER INPUT
workbook = input('INPUT DIRECTORY AND NAME OF WORKBOOK FILE: ')
while  os.path.isfile(workbook) == False:
    print('\nFILE DOES NOT EXIST OR WAS INCORRECTLY INPUTED! PLEASE INPUT AGAIN!')
    workbook = input('INPUT DIRECTORY AND NAME OF WORKBOOK FILE: ')
#worksheet = input('INPUT NAME OF WORKSHEET TO WORK WITH: ')
numrowhdr = input('INPUT NUMBER OF HEADER ROW: ')
while numrowhdr.isdigit() == False:
    print('\nWRONG INPUT! INPUT MUS BE NUMERIC! PLEASE INPUT AGAIN!\n')
    workbook = input('INPUT NUMBER OF HEADER ROW: ')
threshold = input('INPUT MINIMUM ALLOWABLE NUMBER OF NO-NULL CELLS: ')
while threshold.isdigit() == False:
    print('\nWRONG INPUT! INPUT MUS BE NUMERIC! PLEASE INPUT AGAIN!\n')
    workbook = input('INPUT MINIMUM ALLOWABLE NUMBER OF NO-NULL CELLS: ')
needcols = input('INPUT NUMBER OF NEEDED COLUMNS: ')
while threshold.isdigit() == False:
    print('\nWRONG INPUT! INPUT MUS BE NUMERIC! PLEASE INPUT AGAIN!\n')
    workbook = input('INPUT MINIMUM ALLOWABLE NUMBER OF NO-NULL CELLS: ')
#outworkbook = input('INPUT DIRECTORY AND NAME OF FORMATTED WORKBOOK: ')
#nameoutsheet = input('INPUT NAME OF FORMATED SHEET: ')
##########

#FUNCTION GET LIST OF SHEET NAMES FROM WORKBOOK
def makelistofsheets(workbook):
    wb = pd.ExcelFile(workbook)
    return wb.sheet_names
###############################################

#FUNCTION MAKE LIST OF HEADER
def makelistofcolumn(needcols):
    listcol = ['DATE']
    for i in range(0,int(needcols)-1):
        listcol.append(str(i))
    return listcol
    
#############################

#FUNCTION READ EXCEL WORKSHEET INTO PANDAS DATAFRAME
def worksheet2df(workbook, sheetname, numhdrrow):
    hdrrows = range(0,numhdrrow)
    df = pd.read_excel(workbook, skiprows=hdrrows,sheetname = sheetname, header = None)    
    return df
####################################################

#FUNCTION REMOVE EMPTY ROWS
def removeNA(df,threshold):
    df_nonull = df.dropna(thresh = threshold)
    return df_nonull
##########################

#FUNCTION GET NECESSARY DATA FROM DATAFRAME
def getneeddataframe(df,numofcols):
    ndf = df.loc[:,range(int(numofcols))]
    return ndf
###########################################

#FUNCTION CREATE DATAFRAME
def createdf(workbook,numrowhdr,threshold,numofcols):
    df = pd.DataFrame()
    for worksheet in makelistofsheets(workbook):
        df1 = worksheet2df(workbook,str(worksheet),int(numrowhdr))        
        df2 = removeNA(df1, int(threshold))
        df3 = getneeddataframe(df2, int(numofcols))
        df3.columns = makelistofcolumn(numofcols)        
        df = df.append(df3,ignore_index = True)
    df_fin = df.set_index('DATE')     
    return df_fin
##########################

#FUNCTION CREATE DATAFRAME OF MIN, MEAN AND AVERAGE VALUES
def createVALdf(df):    
    dfVAL = pd.DataFrame()
    lst_idx = list(df.index.values)
    lst_min = []
    lst_mean = []
    lst_med = []
    lst_max = []
    lst_25quant = []
    lst_75quant = []
    for i in lst_idx:
        lst_min.append(df.loc[i,:].min())
        lst_mean.append(df.loc[i,:].mean())
        lst_med.append(df.loc[i,:].median())
        lst_max.append(df.loc[i,:].max())
        lst_25quant.append(df.loc[i,:].quantile(0.25))
        lst_75quant.append(df.loc[i,:].quantile(0.75))
    seri_idx = pd.Series(lst_idx)
    seri_min = pd.Series(lst_min)  
    seri_mean = pd.Series(lst_mean)
    seri_med = pd.Series(lst_med)
    seri_max = pd.Series(lst_max)
    seri_25quant = pd.Series(lst_25quant)
    seri_75quant = pd.Series(lst_75quant)
    dfVAL['DATE'] = seri_idx.values
    dfVAL['MIN'] = seri_min.values
    dfVAL['MEAN'] = seri_mean.values
    dfVAL['MEDIAN'] = seri_med.values    
    dfVAL['MAX'] = seri_max.values
    dfVAL['25QUANT'] = seri_25quant.values
    dfVAL['75QUANT'] = seri_75quant.values     
    dfVAL_idx = dfVAL.set_index('DATE')    
    return dfVAL_idx
##########################################################

#FUNCTION CONVERT ROWS TO COLUMN
def df_row2col(df):
    dfr2c = pd.DataFrame()
    lst_idx = df.index
    lst_idxr2c = []
    lst_val = []
    for i in lst_idx:
        for j in df.loc[i,:]:
            lst_idxr2c.append(i)
            lst_val.append(j)
    seri_idx = pd.Series(lst_idxr2c)
    seri_val = pd.Series(lst_val)
    dfr2c['DATE']= seri_idx
    dfr2c['Value']=seri_val
    dfr2c_idx = dfr2c.set_index('DATE')     
    return dfr2c
################################

#FUNCTION CREATE DATAFRAME OF FILTERED VALUES
def createthreshdf(df, thresval):    
    dffil = df[df.Value > float(thresval)]
    dffil_idx = dffil.set_index('DATE')    
    return dffil_idx
#############################################

#FUNCTION LINEAR REGRESSION OF MIN VALUES
def linregress_min(df):
    linreglst = []
    coefficients_min, residuals_min, _, _, _  = np.polyfit(range(len(df.index)), df.loc[:,'MIN'],1,full=True)
    mse_min = residuals_min[0]/(len(df.index))
    nrmse_min = np.sqrt(mse_min)/(df.loc[:,'MIN'].max()-df.loc[:,'MIN'].min())
    for x in range(len(df.loc[:,'MIN'])):
        linreglst.append(coefficients_min[0]*x + coefficients_min[1])    
    return linreglst
#########################################

#FUNCTION LINEAR REGRESSION OF MEAN VALUES
def linregress_mean(df):
    linreglst = []
    coefficients_mean, residuals_mean, _, _, _  = np.polyfit(range(len(df.index)), df.loc[:,'MEAN'],1,full=True)
    mse_mean = residuals_mean[0]/(len(df.index))
    nrmse_mean = np.sqrt(mse_mean)/(df.loc[:,'MEAN'].max()-df.loc[:,'MEAN'].min())
    for x in range(len(df.loc[:,'MEAN'])):
        linreglst.append(coefficients_mean[0]*x + coefficients_mean[1])    
    return linreglst
#########################################

#FUNCTION LINEAR REGRESSION OF MEDIAN VALUES
def linregress_med(df):
    linreglst = []
    coefficients_med, residuals_med, _, _, _  = np.polyfit(range(len(df.index)), df.loc[:,'MEDIAN'],1,full=True)
    mse_med = residuals_med[0]/(len(df.index))
    nrmse_med = np.sqrt(mse_med)/(df.loc[:,'MEDIAN'].max()-df.loc[:,'MEDIAN'].min())
    for x in range(len(df.loc[:,'MEDIAN'])):
        linreglst.append(coefficients_med[0]*x + coefficients_med[1])    
    return linreglst
############################################

#FUNCTION LINEAR REGRESSION OF MAX VALUES
def linregress_max(df):
    linreglst = []
    coefficients_max, residuals_max, _, _, _  = np.polyfit(range(len(df.index)), df.loc[:,'MAX'],1,full=True)
    mse_max = residuals_max[0]/(len(df.index))
    nrmse_max = np.sqrt(mse_max)/(df.loc[:,'MAX'].max()-df.loc[:,'MAX'].min())
    for x in range(len(df.loc[:,'MAX'])):
        linreglst.append(coefficients_max[0]*x + coefficients_max[1])    
    return linreglst
#########################################

#FUNCTION LINEAR REGRESSION OF MAX VALUES
def linregress_25quant(df):
    linreglst = []
    coefficients_25quant, residuals_25quant, _, _, _  = np.polyfit(range(len(df.index)), df.loc[:,'25QUANT'],1,full=True)
    mse_25quant = residuals_25quant[0]/(len(df.index))
    nrmse_25quant = np.sqrt(mse_25quant)/(df.loc[:,'25QUANT'].max()-df.loc[:,'25QUANT'].min())
    for x in range(len(df.loc[:,'25QUANT'])):
        linreglst.append(coefficients_25quant[0]*x + coefficients_25quant[1])    
    return linreglst
#########################################

#FUNCTION LINEAR REGRESSION OF MAX VALUES
def linregress_75quant(df):
    linreglst = []
    coefficients_75quant, residuals_75quant, _, _, _  = np.polyfit(range(len(df.index)), df.loc[:,'75QUANT'],1,full=True)
    mse_75quant = residuals_75quant[0]/(len(df.index))
    nrmse_75quant = np.sqrt(mse_75quant)/(df.loc[:,'75QUANT'].max()-df.loc[:,'75QUANT'].min())
    for x in range(len(df.loc[:,'75QUANT'])):
        linreglst.append(coefficients_75quant[0]*x + coefficients_75quant[1])    
    return linreglst
#########################################

#CREATE DATAFRAME
df = createdf(workbook,int(numrowhdr), int(threshold), int(needcols))
#################

#MAIN MENU FUNCTION
def main_menu():
    os.system('clear')
    print('WELCOME TO HYDRO-METEOROLOGICAL DATA PROCESSING AND VISUALIZING PROGRAM! \n')
    print('PLEASE CHOOSE OPERATION YOU WANT: \n')
    print('1. DRAW PLOTS OF TIME SERIES DATA')
    print('2. EXPORT VALUES ABOVE THRESHOLD AND DRAW PLOTS')
    print('9. COME BACK TO MAIN MENU')
    print('0. TERMINATE THE PROGRAM')    
    choice = input('>> ')
    exec_menu(choice)    
###################


#FUNCTION EXPORT DATAFRAME TO EXCEL FILE
def export2excel(df):   
    outfile = input('INPUT OUTPUT DIRECTORY AND NAME (WITH .XLSX OR .XLS: ')    
    writer = pd.ExcelWriter(outfile)
    outsheet = input('INPUT OUTPUT SHEET: ')
    df.to_excel(writer,outsheet)
    writer.save()
########################################

#print('Slope ' + str(coefficients_min[0]))
#print('NRMSE: ' + str(nrmse_min))

#FUNCTION DECIDE PLOTTING
def plottype(df):
    dfVAL = createVALdf(df)
    plottype = input('DO YOU WANT TO PLOT INDIVIDUAL VALUE OR ALL (EACH/ALL): ')
    if plottype.lower() == 'each':
        plotsing(dfVAL)
    elif plottype.lower() == 'all':
        plotall(dfVAL)
    print('PLEASE CHOOSE OPERATION YOU WANT: \n')
    print('1. DRAW PLOTS OF TIME SERIES DATA')
    print('2. DRAW PLOTS OF VALUES ABOVE THRESHOLD')
    print('9. COME BACK TO MAIN MENU')
    print('0. TERMINATE THE PROGRAM')
    choice = input('>> ')
    exec_menu(choice)
#########################

#FUNCTION PLOTTING MIN/MEAN/MAX TIME SERIES DATA
def plotsing(df):
    print('PLEASE CHOOSE VALUES TO PLOT INCLUDING MIN, MEAN, MEDIAN, MAX, 75TH QUANTILE AND 25TH QUANTILE!')    
    plottype = input('INPUT TYPE OF DATA TO PLOT (MIN/MEAN/MED/MAX/75Q/25Q): ')
    while plottype.lower() not in ['min', 'mean','med', 'max','25q','75q']:
        print('\nINCORRECT INPUT! PLEASE INPUT MIN OR MEAN, MEDIAN OR MAX OR 25TH QUANTILE OR 75TH QUANTILE!')
        plottype = input('INPUT TYPE OF DATA TO PLOT (MIN/MEAN/MED/MAX/75Q/25Q): ')    
    plottitle = input('INPUT TITLE OF PLOT: ')
    xlabel = input('INPUT LABEL FOR X AXIS: ')
    ylabel = input('INPUT LABEL FOR Y AXIS: ')
    xsize = input('INPUT HORIZONTAL SIZE OF PLOT: ')
    ysize = input('INPUT VERTICAL SIZE OF PLOT: ')
    print('FOR COLOR CODE PLEASE GO TO WEBSITE: http://htmlcolorcodes.com \n')
    datacolor = input('INPUT COLOR FOR LINE REPRESENTING MAIN DATA: ')
    trendcolor = input('INPUT COLOR FOR TREND LINE: ')
    exportbool = input('DO YOU WANT TO EXPORT TO PLOT TO FILE (YES/NO): ')
    while exportbool.lower() not in ['yes','no']:
        print('\nINCORRECT INPUT! PLEASE INPUT YES OR NO!')
        exportbool = input('DO YOU WANT TO EXPORT TO PLOT TO FILE (YES/NO): ') 
    dataline = mlines.Line2D([],[],color = '#'+ datacolor, label='Water level')
    trendline = mlines.Line2D([],[],color = '#'+ trendcolor, label='Trendline')
    fig = plt.figure(figsize = (int(xsize),int(ysize)))    
    plt.legend(loc=8, handles = [dataline,trendline], ncol = 2, fontsize = 'medium')
    plt.xlabel(xlabel)
    plt.ylabel(ylabel)
    if plottype.lower() == 'min':
        plt.plot(df.index,df.loc[:,'MIN'], color ='#'+ datacolor)    
        plt.plot(df.index,linregress_min(df), color='#'+ trendcolor)        
    elif plottype.lower() == 'mean':
        plt.plot(df.index,df.loc[:,'MEAN'], color = '#'+ datacolor)    
        plt.plot(df.index,linregress_mean(df), color='#'+ trendcolor)
    elif plottype.lower() == 'med':
        plt.plot(df.index,df.loc[:,'MEDIAN'], color = '#'+ datacolor)    
        plt.plot(df.index,linregress_med(df), color='#'+ trendcolor)
    elif plottype.lower == 'max':
        plt.plot(df.index,df.loc[:,'MAX'], color = '#'+ datacolor)    
        plt.plot(df.index,linregress_max(df), color='#'+ trendcolor)
    elif plottype.lower == '25q':
        plt.plot(df.index,df.loc[:,'25QUANT'], color = '#'+ datacolor)    
        plt.plot(df.index,linregress_25quant(df), color='#'+ trendcolor)
    elif plottype.lower == '75q':
        plt.plot(df.index,df.loc[:,'75QUANT'], color = '#'+ datacolor)    
        plt.plot(df.index,linregress_75quant(df), color='#'+ trendcolor)
    plt.title(plottitle)
    plt.show()
    if exportbool.lower() == 'yes':
        outplot = input('INPUT DIRECTORY AND NAME FOR OUTPUT: ')
        try:            
            fig.savefig(outplot)
        except:
            print('PLOT HAS NOT BEEN SAVED!!!')        
###############################################

#FUNCTION PLOTTING ALL
def plotall(df):
    plottitle = input('INPUT TITLE OF PLOT: ')
    xlabel = input('INPUT LABEL FOR X AXIS: ')
    ylabel = input('INPUT LABEL FOR Y AXIS: ')
    xsize = input('INPUT HORIZONTAL SIZE OF PLOT: ')
    ysize = input('INPUT VERTICAL SIZE OF PLOT: ')
    print('FOR COLOR CODE PLEASE GO TO WEBSITE: http://htmlcolorcodes.com \n')
    datacolor_min = input('INPUT COLOR FOR LINE REPRESENTING MIN VALUES: ')
    datacolor_mean = input('INPUT COLOR FOR LINE REPRESENTING MEAN VALUES: ')
    datacolor_max = input('INPUT COLOR FOR LINE REPRESENTING MAX DATA: ')
    trendcolor = input('INPUT COLOR FOR TREND LINES: ')
    exportbool = input('DO YOU WANT TO EXPORT TO PLOT TO FILE (YES/NO): ')
    while exportbool.lower() not in ['yes','no']:
        print('\nINCORRECT INPUT! PLEASE INPUT YES OR NO!')
        exportbool = input('DO YOU WANT TO EXPORT TO PLOT TO FILE (YES/NO): ')
    dataline_min = mlines.Line2D([],[],color = '#' + datacolor_min, label='Min water level')
    dataline_mean = mlines.Line2D([],[],color = '#' + datacolor_mean, label='Mean water level')
    dataline_max = mlines.Line2D([],[],color = '#' + datacolor_max, label='Max water level ')
    trendline = mlines.Line2D([],[],color = '#' + trendcolor, label='Trendline')
    #plt.xkcd()    
    fig = plt.figure(figsize = (int(xsize),int(ysize)))    
    plt.legend(loc=8, handles = [dataline_min, dataline_mean,dataline_max,trendline], ncol = 4, fontsize = 'medium')
    plt.xlabel(xlabel)
    plt.ylabel(ylabel)    
    plt.plot(df.index,df.loc[:,'MIN'], color = datacolor_min)    
    plt.plot(df.index,linregress_min(df), color=trendcolor)
    plt.plot(df.index,df.loc[:,'MEAN'], color = datacolor_mean)    
    plt.plot(df.index,linregress_mean(df), color=trendcolor)
    plt.plot(df.index,df.loc[:,'MAX'], color = datacolor_max,zorder = 1)    
    plt.plot(df.index,linregress_max(df), color=trendcolor,zorder = 1)    
    plt.title(plottitle)
    plt.show()
    if exportbool.lower() == 'yes':        
        outplot = input('INPUT DIRECTORY AND NAME FOR OUTPUT: ')
        try:            
            fig.savefig(outplot)
        except:
            print('PLOT HAS NOT BEEN SAVED!!!')                
###############################

#FUNCTION PLOTTING THRESHOLD VALUES
def plotthreshval(dfr2c,dfthresh,threshval,savebool):    
    plottitle = input('\nINPUT TITLE OF PLOT: ')
    xlabel = input('\nINPUT LABEL FOR X AXIS: ')
    ylabel = input('\nINPUT LABEL FOR Y AXIS: ')
    xsize = input('\nINPUT HORIZONTAL SIZE OF PLOT: ')
    ysize = input('\nINPUT VERTICAL SIZE OF PLOT: ')
    print('\nFOR COLOR CODE PLEASE GO TO WEBSITE: http://htmlcolorcodes.com \n')        
    threshvalcolr = input('INPUT COLOR FOR VALUE ABOVE THRESHOLD: ')        
    fig = plt.figure(figsize = (int(xsize),int(ysize)))   
    plt.xlabel(xlabel)
    plt.ylabel(ylabel)
        
    plt.scatter(dfthresh.index,dfthresh.loc[:,'Value'],color = '#' + threshvalcolr, marker = 'o', label = 'Water level above ' + str(threshval))    
    plt.legend(loc = 8, ncol=2, fontsize = 'medium')
    plt.title(plottitle)
    plt.show()
    if savebool.lower() == 'yes':
        outplot = input('INPUT DIRECTORY AND NAME FOR OUTPUT: ')
        try:            
            fig.savefig(outplot)
        except:
            print('PLOT HAS NOT BEEN SAVED!!!')     
##################################

#FUNCTION MAIN MENU PLOTTING THRESHOLD VALUES
def plotthreshmenu(df):
    threshval = input('\nINPUT THRESHOLD VALUE: ')
    while threshval.isnumeric() == False:
        print('\nWRONG INPUT! PLEASE INPUT A NUMBER!: \n')
        threshval = input('\nINPUT THRESHOLD VALUE: ')
    exportwbbool = input('\nDO YOU WANT TO EXPORT TO NEW WORKBOOK (YES/NO): ')
    while exportwbbool.lower() not in ['yes','no']:
        print('\nWRONG INPUT! PLEASE INPUT YES OR NO!\n')
        input('\nDO YOU WANT TO EXPORT TO NEW WORKBOOK (YES/NO): ')
    savebool = input('\nDO YOU WANT TO SAVE THE PLOT TO FILE (YES/NO):')
    while savebool.lower() not in ['yes','no']:
        print('\nWRONG INPUT! PLEASE INPUT YES OR NO!\n')
        input('\nDO YOU WANT TO EXPORT TO NEW WORKBOOK (YES/NO): ')
    dfr2c = df_row2col(df)
    dfthresh = createthreshdf(dfr2c, float(threshval))
    plotthreshval(dfr2c,dfthresh,threshval,savebool)
    if exportwbbool.lower() == 'yes':
        export2excel(df)
###################################

#FUNCTION EXECUTE MENU
def exec_menu(choice):
    #os.system('clear')
    ch = choice.lower()
    if ch == '':
        menu_actions['main_menu']()
    else:
        try:
            if ch == '9' or ch == '0':
                menu_actions[ch]()
            else:
                menu_actions[ch](df)
                choice = input('>> ')
                exec_menu(choice)            
        except KeyError:
            print('INVALID SECTION, PLEASE TRY AGAIN. \n')
            menu_actions['main_menu']()
######################

#FUNCTION BACK TO MAIN MENU PROGRAM
def back():
    menu_actions['main_menu']()    
######################

#FUNCTION EXIT PROGRAM
def exit():
    sys.exit()
######################
    
#MENU DEFINITION
menu_actions = {
        'main_menu': main_menu,
        '1': plottype,
        '2': plotthreshmenu,
        '9': back,
        '0': exit}
###############
#MAIN PROGRAM
if __name__ == '__main__':
    main_menu()    
#############
