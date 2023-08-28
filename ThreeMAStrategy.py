#!/usr/bin/env python
# coding: utf-8

# In[4]:


import pandas as pd

def CreateExcel(MA,num_years,filename,num_shares,margin):
    def DownloadDataset(filename):
        df = pd.read_csv(filename)
        df = df.dropna(axis=0,how='any')
        df['Date'] = pd.to_datetime(df['Date'],infer_datetime_format=True)
        df = df.reset_index(level=None,drop=False,inplace=False,col_level=0,col_fill='')
        return df
    
    def n_day_MA(n,df):
        rows = len(df.axes[0])
        MA_data = []
        for i in range(rows-1,-1,-1):
            if i>rows-n:
                MA_data.append(0)
            else:
                df1 = pd.DataFrame(df,index=range(i,i+n))
                #print(df1.tail(n))
                Mean = df1['Close'].mean()
                #print(Mean)
                MA_data.append(Mean)
        MA_data.reverse()
        df[str(n)+'MA'] = MA_data
        column = df.columns.get_loc(str(n)+'MA')
        rows_del = []
        for i in range(rows):
            row = df.iloc[i]
            value = row.iloc[column]
            if value == 0:
                rows_del.append(i)
        df = df.drop(labels=rows_del,axis=0)
        df = df.reset_index(level=None,drop=False,inplace=False,col_level=0,col_fill='')
        return df
    
    def Find_Trades(df,MA):
    
        MA_fast = df[str(MA[0])+'MA']
        MA_slow = df[str(MA[1])+'MA']
        Diff_MA = MA_fast - MA_slow
        df['Diff_MA'] = Diff_MA

        column_diff = df.columns.get_loc('Diff_MA')
        column_fast = df.columns.get_loc(str(MA[0])+'MA')
        column_slow = df.columns.get_loc(str(MA[1])+'MA')
        column_long = df.columns.get_loc(str(MA[2])+'MA')
        rows = len(df.axes[0])
        LorS = ['Neither']
        for i in range(rows-1,0,-1):
            diff_1 = df.iloc[i,column_diff] #Today's difference
            diff_2 = df.iloc[i-1,column_diff] #Tomorrow's difference
            long_MA = df.iloc[i,column_long] #long MA today
            fast_MA = df.iloc[i,column_fast] #Fast MA today
            slow_MA = df.iloc[i,column_slow] #Slow MA today
            if diff_1 > 0 and diff_2 < 0 and long_MA < fast_MA and long_MA < slow_MA:
                LorS.append('Short')
            elif diff_1 < 0 and diff_2 > 0 and long_MA > fast_MA and long_MA > slow_MA:
                LorS.append('Long')
            else:
                LorS.append('Neither')
        LorS.reverse()
        df['LorS'] = LorS

        return df
    
    # Puts information about all days traded into 1 dataframe.
    def Trades_df(df,MA,num_shares,margin):
    
        trades = df.loc[df['LorS'] != 'Neither']
        trades = trades.reset_index(level=None,drop=False,inplace=False,col_level=0,col_fill='')

        trades_df = pd.DataFrame(columns = ['Entry Date','Entry Price','Exit Date','Exit Price',str(MA[0])+'MA',
                                            str(MA[1])+'MA',str(MA[2])+'MA','Profit','Long/Short','ROI'])
        for i in range(len(trades.axes[0])-1,1,-1):
            entry_date = trades.iloc[i,trades.columns.get_loc('Date')]
            entry_price = trades.iloc[i,trades.columns.get_loc('Close')]
            exit_date = trades.iloc[i-1,trades.columns.get_loc('Date')]
            exit_price = trades.iloc[i-1,trades.columns.get_loc('Open')]
            MA_fast = trades.iloc[i,trades.columns.get_loc(str(MA[0])+'MA')]
            MA_slow = trades.iloc[i,trades.columns.get_loc(str(MA[1])+'MA')]
            MA_long = trades.iloc[i,trades.columns.get_loc(str(MA[2])+'MA')]
            long_short = trades.iloc[i,trades.columns.get_loc('LorS')]
            profit = exit_price - entry_price
            if long_short == 'Short':
                profit *= 1
            ROI = round(profit*num_shares/(margin/100),2)
            row = {'Entry Date':entry_date,'Entry Price':entry_price,'Exit Date':exit_date,'Exit Price':exit_price,
                   str(MA[0])+'MA':MA_fast,str(MA[1])+'MA':MA_slow,str(MA[2])+'MA':MA_long,'Profit':profit,
                   'Long/Short':long_short,'ROI':ROI}
            trades_df = trades_df.append(row,ignore_index=True)

        return trades_df
    
    def AddStatToDF(trades_df,Stat_name,Stat,Rows):
        Series = [Stat]
        for i in range(Rows-1):
            Series.append(None)
        
        trades_df[Stat_name] = Series
        
        return trades_df
    
    def StatsAndSpreadsheet(trades_df,MA): #Works out stats and adds them to dataframe.
        #Cum ROI
        trades_df['Cum ROI'] = trades_df['ROI'].cumsum()
        
        #Win/Loss
        win_or_loss = [] 
        total_rows = len(trades_df.axes[0])
        for i in range(total_rows):
            profit = trades_df.iloc[i,trades_df.columns.get_loc('Profit')]
            if profit > 0:
                win_or_loss.append('W')
            else:
                win_or_loss.append('L')
        trades_df['W/L'] = win_or_loss
        
        #Streak 
        # Algorithm assumes next trade is same result as previous. Steps are taken if not. 
        row = 0 #Records row you are checking
        streak = 1 #Records number of consecutive wins/losses
        streak_array = [] #Records streak on each trade day
        while row < total_rows:
            if streak == 1: #Adds 1 to the streak array. Increments streak, moves to next row
                streak_array.append(streak)
                streak += 1
                row += 1
            else:
                # Compares losses on current and previous row
                W_or_L_today = trades_df.iloc[row,trades_df.columns.get_loc('W/L')]
                W_or_L_previous = trades_df.iloc[row-1,trades_df.columns.get_loc('W/L')]
                if W_or_L_today == W_or_L_previous: #Adds the streak value to array, increments streak, moves to next row
                    streak_array.append(streak)
                    streak += 1
                    row += 1
                else: #Resets streak to 1 and begins loop again. Doesn't do anything to streak or row. Dealt with in the if
                    streak = 1                                            # Above
        
        trades_df['Streak'] = streak_array
        
        #Peak Equity
        peak_equities = []
        for i in range(total_rows):
            equity_today = trades_df.iloc[i,trades_df.columns.get_loc('Cum ROI')]
            if i == 0:
                peak_equities.append(equity_today)
            elif equity_today > peak_equities[len(peak_equities)-1]:
                peak_equities.append(equity_today)
            else:
                peak_equities.append(peak_equities[len(peak_equities)-1])
        
        trades_df['Peak Equity'] = peak_equities
        
        #Drawdown
        drawdown = trades_df['Cum ROI'] - trades_df['Peak Equity']
        trades_df['Drawdown'] = drawdown
        
        #Stats about Strategy as a whole
        
        #Wins, Losses, Win Ratio
        W_or_L = trades_df['W/L']
        num_wins = W_or_L.value_counts()['W']
        num_losses = W_or_L.value_counts()['L']
        W_ratio = round(num_wins/(num_wins+num_losses),2)
        
        #Average Win and Loss return, Reward to risk ratio, expectancy ratio
        ROI_Wins = trades_df.query("ROI > 0")
        ROI_losses = trades_df.query("ROI < 0")
        Avg_win_return = ROI_Wins['ROI'].mean()
        Avg_loss_return = ROI_losses['ROI'].mean()
        Reward_to_risk = round(-1 * Avg_win_return/ Avg_loss_return,2)
        
        expectancy_ratio = round((Reward_to_risk*W_ratio) - (1-W_ratio),2)
        
        #CALMAR ratio
        total_ROI = trades_df['ROI'].sum()
        drawdown = trades_df['Drawdown']
        Min_Drawdown = drawdown.min()
        CALMAR = round(-1*total_ROI/num_years * 1/Min_Drawdown,2)
        
        #Add all to Dataframe
        trades_df = AddStatToDF(trades_df,'Wins',num_wins,total_rows)
        trades_df = AddStatToDF(trades_df,'Losses',num_losses,total_rows)
        trades_df = AddStatToDF(trades_df,'Win Ratio',W_ratio,total_rows)
        trades_df = AddStatToDF(trades_df,'Avg Win Amount',round(Avg_win_return,2),total_rows)
        trades_df = AddStatToDF(trades_df,'Avg Loss Amount',round(Avg_loss_return,2),total_rows)
        trades_df = AddStatToDF(trades_df,'Reward to Risk Ratio',Reward_to_risk,total_rows)
        trades_df = AddStatToDF(trades_df,'Expectancy Ratio',expectancy_ratio,total_rows)
        trades_df = AddStatToDF(trades_df,'CALMAR',CALMAR,total_rows)
        
        return trades_df
        
    df = DownloadDataset(filename)
    for i in range(len(MA)):
        df = n_day_MA(MA[i],df)
    df = Find_Trades(df,MA)
    trades_df = Trades_df(df,MA,num_shares,margin)
    trades_df = StatsAndSpreadsheet(trades_df,MA)
    
    return trades_df

def CreateCombinations(stock_name):
    #Will only apply long MA to fast and slow MA pairs with high CALMAR ratio
    two_MA_doc_name = '2MA_backtest_result_'+str(stock_name)+'.xlsx'
    CALMAR_df = pd.read_excel(two_MA_doc_name,sheet_name='CALMAR Ordering')
    top_CALMAR = CALMAR_df.query('CALMAR >= 0.7') #Classifying 0.7 as a large CALMAR ratio

    rows = len(top_CALMAR.axes[0])
    MA_combos = []
    long_MAs = list(range(40,56,1))
    for i in range(rows):
        MA_combo = top_CALMAR.iloc[i,top_CALMAR.columns.get_loc('MA combination')]
        fast_MA = int(MA_combo[3])
        slow_MA = int(MA_combo[5:7])
        for j in range(len(long_MAs)):
            MA_combos.append([fast_MA,slow_MA,long_MAs[j]])
    
    return MA_combos

def CreateExcelSheets(MA_combos,filename,num_years,num_shares,margin):
    excel_sheets = []
    for i in range(len(MA_combos)):
        trades_df = CreateExcel(MA_combos[i],num_years,filename,num_shares,margin)
        excel_sheets.append(trades_df)
    
    return excel_sheets

def CreateCALMARSheet(excel_sheets,combinations):
    
    #Creates another sheet giving CALMARs in descending order. 
    CALMARs = []
    sheet_names = []
    for i in range(len(excel_sheets)):
        CALMARs.append(excel_sheets[i].iloc[0,excel_sheets[i].columns.get_loc('CALMAR')])
        sheet_names.append('MA-'+str(combinations[i][0])+','+str(combinations[i][1])+','
                          +str(combinations[i][2]))
    CALMAR_df = {'MA combination':sheet_names,'CALMAR':CALMARs}
    CALMAR_df = pd.DataFrame(CALMAR_df)
    CALMAR_df = CALMAR_df.sort_values(by='CALMAR',ascending=False)
    return CALMAR_df

def CreateExpectancySheet(excel_sheets,combinations):
    
    #Creates another sheet giving Expectancy ratio in descending order
    ERs = []
    sheet_names = []
    for i in range(len(excel_sheets)):
        ERs.append(excel_sheets[i].iloc[0,excel_sheets[i].columns.get_loc('Expectancy Ratio')])
        sheet_names.append('MA-'+str(combinations[i][0])+','+str(combinations[i][1])+
                          ','+str(combinations[i][2]))
    ER_df = {'MA combination':sheet_names,'Expectancy Ratio':ERs}
    ER_df = pd.DataFrame(ER_df)
    ER_df = ER_df.sort_values(by='Expectancy Ratio',ascending=False)
    return ER_df

def CreateExcelDoc(excel_sheets,combinations,stock_name):
    #Create excel doc
    doc_name = '3MA_backtest_result_'+str(stock_name)+'.xlsx'
    writer = pd.ExcelWriter(doc_name, engine='xlsxwriter')
    for i in range(len(excel_sheets)):
        sheet_name = 'MA-'+str(combinations[i][0])+','+str(combinations[i][1])+','+str(combinations[i][2])
        excel_sheets[i].to_excel(writer, sheet_name=sheet_name)

    CALMAR_df = CreateCALMARSheet(excel_sheets,combinations)
    ER_df = CreateExpectancySheet(excel_sheets,combinations)
    CALMAR_df.to_excel(writer,sheet_name='CALMAR Ordering')
    ER_df.to_excel(writer,sheet_name='Expectancy Ratio Ordering')
    writer.close()   
    
def ThreeMAStrategy(stock_name,filename,num_years,num_shares,margin):
    combinations = CreateCombinations(stock_name)
    excel_sheets = CreateExcelSheets(combinations,filename,num_years,num_shares,margin)
    CreateExcelDoc(excel_sheets,combinations,stock_name)
    


# In[5]:


ThreeMAStrategy('NIFTY50','NIFTY 50_data.csv',5,50,110000)
    

