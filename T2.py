import pandas as pd
import csv
from datetime import timedelta

import datetime
import os
def trade(d):
    filter_column = 'Period after CO create date'
    DateNow = datetime.datetime.now()
    day = DateNow.strftime("%d")
    month = DateNow.strftime("%b")
    year = DateNow.strftime("%Y")
    current_date = day + "_" + month + "_" + year

    files = os.listdir('files')

    for f in files:
        if f[:10] == 'FX Trades-':
            TradeData = pd.read_excel("files/"+f, sheet_name = 0, header = None, skiprows=1)

        if f[:12] == 'FX Trades HQ':
            HqTradeData = pd.read_excel("files/"+f, sheet_name = 0, header = None, skiprows=1)
            getmonth = f[18:-7]
        if f[:6] == 'FX-133':
            FX133TradeData = pd.read_excel("files/"+f, sheet_name = 0, header = None, skiprows=1)

    month_dict = {'Jan':1, 'Feb':2, 'Mar':3, 'Apr':4, 'May':5, 'Jun':6, 'Jul':7, 'Aug':8, 'Sep':9, 'Oct':10, 'Nov':11, 'Dec':12}

    #Get indices of approved records

    approved_index = []
    a_index = 0
    for data in TradeData[3]:
        if data[:9] != 'Cancelled':
            approved_index.append(a_index)
        a_index+=1
    #================Columns required from FX Trades-01-31Oct19==========

    TradeRequestID = []#trade id in FX_Trades file

    for index in approved_index:
         TradeRequestID.append(TradeData[0][index])

    #================END Columns required from FX Trades-01-31Oct19==========

    #================Columns required from FX Trades HQ-01-31Oct19==========

    #Get data from FX_Trade_HQ
    get_FXhq_Index = []

    for x in TradeRequestID:
        FX_Hq_index = 0
        for y in HqTradeData[0]:
            if y == x:
                get_FXhq_Index.append(FX_Hq_index)
            FX_Hq_index+=1


    DealNumber = []#Trans

    for x in get_FXhq_Index:

        DealNumber.append(HqTradeData[2][x])#Trans


    #================END Columns required from FX Trades HQ-01-31Oct19==========

    #================Columns required from FX-133 Rpt-01-31Oct19==========

    #=======================EUR
    EURstart_index = 0
    for x in FX133TradeData[0]:
        if isinstance(x, datetime.datetime) == False:

            if isinstance(x, float) == False and x[:12] == 'EUR ( Euro )':
                break
        EURstart_index+=1

    EURstop_index = 0

    for x in FX133TradeData[0]:
        if EURstop_index > EURstart_index:
            if isinstance(x, datetime.datetime) == False:

                if isinstance(x, float) == False and x[:14] == 'Total Currency':
                    break
        EURstop_index+=1

    EURstart_index = EURstart_index + 1

    #usd_get_dealnumbers = []
    EURIndex_get_DealNumbers = []
    count_EUR_index = 0
    for x in FX133TradeData[48]:
        if count_EUR_index > EURstart_index and count_EUR_index < EURstop_index:
            if isinstance(x, int) == True:
                #usd_get_dealnumbers.append(x)
                EURIndex_get_DealNumbers.append(count_EUR_index)
        count_EUR_index+=1



    #=======================================================

    #=======================USD
    USDstart_index = 0
    for x in FX133TradeData[0]:
        if isinstance(x, datetime.datetime) == False:

            if isinstance(x, float) == False and x[:17] == 'USD ( US Dollar )':
                break
        USDstart_index+=1

    USDstop_index = 0

    for x in FX133TradeData[0]:
        if USDstop_index > USDstart_index:
            if isinstance(x, datetime.datetime) == False:


                if isinstance(x, float) == False and x[:14] == 'Total Currency':
                    break;
        USDstop_index+=1

    USDstart_index = USDstart_index + 1

    #usd_get_dealnumbers = []
    usdIndex_get_DealNumbers = []
    count_usd_index = 0
    for x in FX133TradeData[48]:
        if count_usd_index > USDstart_index and count_usd_index < USDstop_index:
            if isinstance(x, int) == True:
                #usd_get_dealnumbers.append(x)
                usdIndex_get_DealNumbers.append(count_usd_index)
        count_usd_index+=1

    #======================GET DEALERS/TRADERS using dealnumbers
    #USD
    #USD_dealers = []
    ALL_indices = []
    getPartners = []
    with open(getmonth.upper()+year+"_Report.csv", 'w') as output:

        trade_writer = csv.writer(output, delimiter = ',')
        trade_writer.writerow(['A','B','C','D','E','F','G', 'H', 'I', 'J','K', 'L' ] )
        trade_writer.writerow(['Index','CO Request ID','Creation Date','FX Deal No','Deal Amount','Currency','Value Date HQ', 'Value Date 133', 'Business Partner', filter_column,'Add 3 days number; Average number of days not received by CO', 'Trader' ] )
        j = 0
        count_record = 1
        for y in DealNumber:

            for x in usdIndex_get_DealNumbers:
                date_133US = FX133TradeData[32][x].replace(".","/");
                date_133USobj = datetime.datetime.strptime(date_133US, '%d/%m/%Y')
                if y == FX133TradeData[48][x] and HqTradeData[7][j][:3] != 'DKK' and (date_133USobj-HqTradeData[15][j]) >= timedelta(days = d):
                    if HqTradeData[15][j].month == (month_dict[getmonth] - 1):
                        trade_writer.writerow([count_record, HqTradeData[0][j],HqTradeData[15][j],str(y),"{:,.2f}".format(HqTradeData[9][j]),HqTradeData[7][j],TradeData[9][approved_index[j]],date_133USobj,FX133TradeData[41][x],(date_133USobj-HqTradeData[15][j]),((date_133USobj-HqTradeData[15][j])),FX133TradeData[3][x] ])
                        #print(HqTradeData[0][j], TradeData[9][approved_index[j]]-date_133USobj,' : ',HqTradeData[7][j][:3])
                        ALL_indices.append(j)
                        getPartners.append( FX133TradeData[41][x] )
                        count_record+=1

                    else:
                        trade_writer.writerow([count_record, HqTradeData[0][j],HqTradeData[15][j],str(y),"{:,.2f}".format(HqTradeData[9][j]),HqTradeData[7][j],TradeData[9][approved_index[j]],date_133USobj ,FX133TradeData[41][x],(date_133USobj-HqTradeData[15][j]),((date_133USobj-HqTradeData[15][j])+timedelta(days=3)),FX133TradeData[3][x] ])
                        #print(HqTradeData[0][j], TradeData[9][approved_index[j]]-date_133USobj,' : ',HqTradeData[7][j][:3])
                        ALL_indices.append(j)
                        getPartners.append( FX133TradeData[41][x] )
                        count_record+=1

            for x in EURIndex_get_DealNumbers:
                date_133EU = FX133TradeData[32][x].replace(".","/");
                date_133EUobj = datetime.datetime.strptime(date_133EU, '%d/%m/%Y')
                if y == FX133TradeData[48][x] and HqTradeData[7][j][:3] != 'DKK' and (date_133EUobj-HqTradeData[15][j]) >= timedelta(days = d):
                    if HqTradeData[15][j].month == (month_dict[getmonth] - 1):
                        trade_writer.writerow([count_record, HqTradeData[0][j],HqTradeData[15][j],str(y),"{:,.2f}".format(HqTradeData[9][j]),HqTradeData[7][j],TradeData[9][approved_index[j]],date_133EUobj,FX133TradeData[41][x],(date_133EUobj-HqTradeData[15][j]),((date_133EUobj-HqTradeData[15][j])),FX133TradeData[3][x] ])
                        ALL_indices.append(j)
                        getPartners.append( FX133TradeData[41][x] )
                        count_record+=1
                    else:
                        trade_writer.writerow([count_record, HqTradeData[0][j],HqTradeData[15][j],str(y),"{:,.2f}".format(HqTradeData[9][j]),HqTradeData[7][j],TradeData[9][approved_index[j]],date_133EUobj,FX133TradeData[41][x],(date_133EUobj-HqTradeData[15][j]),((date_133EUobj-HqTradeData[15][j])+timedelta(days=3)),FX133TradeData[3][x] ])
                        ALL_indices.append(j)
                        getPartners.append( FX133TradeData[41][x] )
                        count_record+=1


            j+=1
    output.close()
    #for delete in files:

    #    os.remove('files/'+delete);

    missing_indices = []
    for y in range( 0, (ALL_indices[-1] + 1) ):
        if y not in ALL_indices:
            missing_indices.append(y)
    print('\n')
    print('Records missing equals: ',len(missing_indices))
    print ('See missing records below: ')
    for y in missing_indices:
        print('Missing index: ',y,'; Deal Number: ',HqTradeData[2][y], '; Currency: ',HqTradeData[7][y],'; Deal Amount: ', "{:,}".format( HqTradeData[9][y] ), '; Creation date: ',HqTradeData[15][y], '; Value date: ',HqTradeData[8][y])
    print('\n')

    x = getPartners
    dict = {}

    for p in getPartners:
        count = 0
        for q in x:
            if p == q:
                count+=1
        dict[p] = count
    total_transactions = 0
    for x in dict:
        print(x, dict[x])
        total_transactions = total_transactions+dict[x]
    print('\n', 'The total number of transactions equals:',total_transactions)






#=======================
