import pandas as pd

from datetime import timedelta
import datetime
import os
import calendar

def trade_graph(d):

    filter_column = 'Period after CO create date'
    DateNow = datetime.datetime.now()
    day = DateNow.strftime("%d")
    month = DateNow.strftime("%b")
    year = DateNow.strftime("%Y")
    current_period = month+year


    files = os.listdir('files')

    for f in files:
        if f[:10] == 'FX Trades-':
            TradeData = pd.read_excel("files/"+f, sheet_name = 0, header = None, skiprows=1)
        if f[:12] == 'FX Trades HQ':
            HqTradeData = pd.read_excel("files/"+f, sheet_name = 0, header = None, skiprows=1)
            getmonth = f[18:-7]
        if f[:6] == 'FX-133':
            FX133TradeData = pd.read_excel("files/"+f, sheet_name = 0, header = None, skiprows=1)

    month_dict = {'Jan':1, 'Feb':2, 'Mar':3, 'Apr':4, 'May':5, 'June':6, 'July':7, 'Aug':8, 'Sep':9, 'Oct':10, 'Nov':11, 'Dec':12}



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


    with open(getmonth.upper()+year+'_Graph.html', 'w') as graph:
        print('<!DOCTYPE html>', file = graph)
        print("<html lang ='en' dir='ltr'>", file = graph)
        print('<head><meta charset="utf-8">', file = graph)
        print('<script type="text/javascript" src="https://www.gstatic.com/charts/loader.js"></script>', file = graph)
        print('<link href="https://fonts.googleapis.com/css?family=Nanum+Gothic&display=swap" rel="stylesheet">', file = graph)
        print('<script type="text/javascript">', file = graph)
        print("google.charts.load('current', {'packages':['corechart']});", file = graph)
        print('google.charts.setOnLoadCallback(drawStuff);', file = graph)
        print('function drawStuff() {', file = graph)
        print('var data = new google.visualization.arrayToDataTable([', file = graph)
        print("['Days', 'Number of Transactions', {role:'style'}, {role : 'annotation'}],", file = graph)

        #======CALCULATE TOTAL NUMBER OF TRANSACTIONS======

        getPartners = []

        j = 0
        for y in DealNumber:

            for x in usdIndex_get_DealNumbers:
                if y == FX133TradeData[48][x] and HqTradeData[7][j][:3] != 'DKK':

                    getPartners.append( FX133TradeData[41][x] )


            for x in EURIndex_get_DealNumbers:
                if y == FX133TradeData[48][x] and HqTradeData[7][j][:3] != 'DKK':

                    getPartners.append( FX133TradeData[41][x] )
            j+=1

        x = getPartners
        dict = {}

        for p in getPartners:
            count = 0
            for q in x:
                if p == q:
                    count+=1
            dict[p] = count

        final_transactions = 0
        for x in dict:
            print(x, dict[x])
            final_transactions = final_transactions+dict[x]


        #=====END TOTAL NUMBER OF TRANSACTIONS====



        for day in range(-d,d+1):

            getPartners = []

            j = 0
            for y in DealNumber:

                for x in usdIndex_get_DealNumbers:

                    date_133US = FX133TradeData[32][x].replace(".","/");
                    date_133USobj = datetime.datetime.strptime(date_133US, '%d/%m/%Y')
                    periodUS = date_133USobj-TradeData[9][approved_index[j]]

                    if month_dict[getmonth] == 1:

                        if TradeData[9][approved_index[j]].month == 12:

                            if y == FX133TradeData[48][x] and HqTradeData[7][j][:3] != 'DKK' and (periodUS) == timedelta(days = day):

                                getPartners.append( FX133TradeData[41][x] )

                        else:

                            if y == FX133TradeData[48][x] and HqTradeData[7][j][:3] != 'DKK' and (periodUS) + timedelta(days = 3) == timedelta(days = day):

                                getPartners.append( FX133TradeData[41][x] )
                    else:

                        if TradeData[9][approved_index[j]].month == (month_dict[getmonth] - 1):

                            if y == FX133TradeData[48][x] and HqTradeData[7][j][:3] != 'DKK' and (periodUS) == timedelta(days = day):

                                getPartners.append( FX133TradeData[41][x] )

                        else:

                            if y == FX133TradeData[48][x] and HqTradeData[7][j][:3] != 'DKK' and (periodUS) + timedelta(days = 3) == timedelta(days = day):

                                getPartners.append( FX133TradeData[41][x] )



                for x in EURIndex_get_DealNumbers:
                    date_133EU = FX133TradeData[32][x].replace(".","/");
                    date_133EUobj = datetime.datetime.strptime(date_133EU, '%d/%m/%Y')
                    periodEU = date_133EUobj-TradeData[9][approved_index[j]]

                    if month_dict[getmonth] == 1:

                        if TradeData[9][approved_index[j]].month == 12:

                            if y == FX133TradeData[48][x] and HqTradeData[7][j][:3] != 'DKK' and (periodEU) == timedelta(days = day):

                                getPartners.append( FX133TradeData[41][x] )

                        else:

                            if y == FX133TradeData[48][x] and HqTradeData[7][j][:3] != 'DKK' and (periodEU)+timedelta(days = 3) == timedelta(days = day):

                                getPartners.append( FX133TradeData[41][x] )

                    else:

                        if TradeData[9][approved_index[j]].month == (month_dict[getmonth] - 1):

                            if y == FX133TradeData[48][x] and HqTradeData[7][j][:3] != 'DKK' and (periodEU) == timedelta(days = day):

                                getPartners.append( FX133TradeData[41][x] )

                        else:

                            if y == FX133TradeData[48][x] and HqTradeData[7][j][:3] != 'DKK' and (periodEU)+timedelta(days = 3) == timedelta(days = day):

                                getPartners.append( FX133TradeData[41][x] )

                j+=1


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

            if total_transactions != 0:

                if day < 6:

                    if day == 1 or day == -1:

                        print('["',day,' day", ',total_transactions,', "green", "',round((total_transactions/final_transactions)*100),'%"],',file = graph)

                    else:

                        print('["',day,' days", ',total_transactions,', "green", "',round((total_transactions/final_transactions)*100),'%"],',file = graph)

                else:


                    print('["',(day),' days", ',total_transactions,', "red", "',round((total_transactions/final_transactions)*100),'%"],',file = graph)


        print(']);', file = graph)
        print('var options = {', file = graph)
        print("title: 'Number of approved FX Trades',", file = graph)
        print('height: 1000,', file = graph)
        print('width: 1600,', file = graph)
        print("legend: { position: 'none' },", file = graph)
        print("bars: 'horizontal',", file = graph)
        print(' bar: { groupWidth: "90%" }', file = graph)
        print('};', file = graph)
        print("var chart = new google.visualization.BarChart(document.getElementById('top_x_div'));", file = graph)
        print("chart.draw(data, options);", file = graph)
        print('};', file = graph)
        print('</script>', file = graph)
        print('<style type="text/css">', file = graph)
        print("h2{", file = graph)
        print("font-family: 'Nanum Gothic', sans-serif;", file = graph)
        print("color:black;", file = graph)
        print("}", file = graph)
        print("h3{", file = graph)
        print("font-family: 'Nanum Gothic', sans-serif;", file = graph)
        print("color:grey;", file = graph)
        print('}', file = graph)
        print("td{ font-family: 'Nanum Gothic', sans-serif; }",file = graph)
        print("strong{ font-family: 'Nanum Gothic', sans-serif; }",file = graph)

        print("</style>", file = graph)
        print("</head>", file = graph)
        print('<body>', file = graph)
        print('<h2>FX Trades 1-',calendar.monthrange(int(year), month_dict[getmonth])[1], getmonth, year,' Days Delivery To CO</h2>', file = graph)
        #print('<h3>Days Delivery </h3>', file = graph)
        print('<table>', file = graph)

        print('<tr><td><strong>Delivery Days to CO</strong></td><td><div style = "display:inline-block;" id="top_x_div"></div></td></tr>', file = graph)
        print('</table>', file = graph)

        print('<table style="border:none">', file = graph)
        print('<tr><td><strong>Key</strong></td><td></td></tr>', file = graph)

        print('<tr><td><div style="display:inline-block; width:50px; height:50px; background-color:green"></div></td><td>FX Trades delivered to CO within 5 days</td>',file = graph)

        print('<tr><td><div style="display:inline-block; width:50px; height:50px; background-color:red"></div></td><td>FX Trades delivered to CO after 5 days onwards</td>',file = graph)
        print('</table>', file = graph)
        print('</body></html>', file = graph)

    graph.close()
    for delete in files:


        os.remove('files/'+delete)
