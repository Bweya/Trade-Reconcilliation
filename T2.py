import pandas as pd
import csv
from datetime import timedelta

TradeData = pd.read_excel("files/FX Trades-01-31Oct19.xlsx", sheet_name = 0, header = None, skiprows=1)
HqTradeData = pd.read_excel("files/FX Trades HQ-01-31Oct19.xlsx", sheet_name = 0, header = None, skiprows=1)
FX133TradeData = pd.read_excel("files/FX Deals - 133Rpt-01-31Oct19.xls.xlsx", sheet_name = 0, header = None, skiprows=1)

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
    if isinstance(x, float) == False and x[:12] == 'EUR ( Euro )':
        break
    EURstart_index+=1

EURstop_index = 0

for x in FX133TradeData[0]:
    if EURstop_index > EURstart_index:
        if isinstance(x, float) == False and x[:14] == 'Total Currency':
            break;
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
    if isinstance(x, float) == False and x[:17] == 'USD ( US Dollar )':
        break
    USDstart_index+=1

USDstop_index = 0

for x in FX133TradeData[0]:
    if USDstop_index > USDstart_index:
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

#=======================XAF
#XAFstart_index = 0
#for x in FX133TradeData[0]:
#    if isinstance(x, float) == False and x[:3] == 'XAF':
#        break
#    XAFstart_index+=1

#XAFstop_index = 0

#for x in FX133TradeData[0]:
#    if XAFstop_index > XAFstart_index:
#        if isinstance(x, float) == False and x[:14] == 'Total Currency':
#            break;
#    XAFstop_index+=1

#XAFstart_index = XAFstart_index + 1

#XAFIndex_get_dealnumbers = []
#count_XAF_index = 0
#for x in FX133TradeData[48]:
#    if count_XAF_index > XAFstart_index and count_XAF_index < XAFstop_index:
#        if isinstance(x, int) == True:
#            XAFIndex_get_dealnumbers.append(count_XAF_index)
#    count_XAF_index+=1

#=======================XOF
#XOFstart_index = 0
#for x in FX133TradeData[0]:
#    if isinstance(x, float) == False and x[:3] == 'XOF':
#        break
#    XOFstart_index+=1

#XOFstop_index = 0

#for x in FX133TradeData[0]:
#    if XOFstop_index > XOFstart_index:
#        if isinstance(x, float) == False and x[:14] == 'Total Currency':
#            break;
#    XOFstop_index+=1

#XOFstart_index = XOFstart_index + 1

#XOFIndex_get_dealnumbers = []
#count_XOF_index = 0
#for x in FX133TradeData[48]:
#    if count_XOF_index > XOFstart_index and count_XOF_index < XOFstop_index:
#        if isinstance(x, int) == True:
#            XOFIndex_get_dealnumbers.append(count_XOF_index)
#    count_XOF_index+=1

#======================GET DEALERS/TRADERS using dealnumbers
#USD
#USD_dealers = []
ALL_indices = []
getPartners = []
with open("Report.csv", 'w') as output:

    trade_writer = csv.writer(output, delimiter = ',')
    trade_writer.writerow(['Index','CO Request ID','Creation Date','CO value Date', 'FX Deal No.','Business Partner','Value Dt', 'Deal Amount', 'Currency', 'Number of Days > 3','Value Dt - CO Value Dt','Trader' ] )
    j = 0
    for y in DealNumber:

        for x in usdIndex_get_DealNumbers:
            if y == FX133TradeData[48][x]:
                #index_in_HQ.append(j+y+FX133TradeData[3][x])
                trade_writer.writerow([ str(j),HqTradeData[0][j],HqTradeData[15][j],TradeData[9][approved_index[j]],str(y),FX133TradeData[41][x],HqTradeData[8][j],"{:,.2f}".format(HqTradeData[9][j]),HqTradeData[7][j],(HqTradeData[8][j]-HqTradeData[15][j]),(HqTradeData[8][j]-TradeData[9][approved_index[j]]),FX133TradeData[3][x] ])
                ALL_indices.append(j)
                getPartners.append( FX133TradeData[41][x] )


        for x in EURIndex_get_DealNumbers:
            if y == FX133TradeData[48][x]:

            #and (HqTradeData[8][j]-HqTradeData[15][j]) > timedelta(days = 0):
                trade_writer.writerow([ str(j),HqTradeData[0][j],HqTradeData[15][j],TradeData[9][approved_index[j]],str(y),FX133TradeData[41][x],HqTradeData[8][j],"{:,.2f}".format(HqTradeData[9][j]),HqTradeData[7][j],(HqTradeData[8][j]-HqTradeData[15][j]),(HqTradeData[8][j]-TradeData[9][approved_index[j]]),FX133TradeData[3][x] ])
                ALL_indices.append(j)
                getPartners.append( FX133TradeData[41][x] )
                #print(str(j)+' USD '+' '+str(y)+' '+FX133TradeData[3][x], HqTradeData[0][j], HqTradeData[7][j], HqTradeData[8][j], HqTradeData[9][j], HqTradeData[15][j],(HqTradeData[8][j]-HqTradeData[15][j]))

        #j+=1

        #USD_dealers.append(FX133TradeData[3][x])

    #XAF
    #j = 0
    #for y in DealNumber:
    #    for x in XAFIndex_get_dealnumbers:
    #        if y == FX133TradeData[48][x] and (HqTradeData[8][j]-HqTradeData[15][j]) > timedelta(days = 0):
    #            trade_writer.writerow([ str(j),HqTradeData[0][j],HqTradeData[15][j],'XAF',str(y),FX133TradeData[42][x],HqTradeData[8][j],"{:,.2f}".format(HqTradeData[9][j]),HqTradeData[7][j],(HqTradeData[8][j]-HqTradeData[15][j]),FX133TradeData[3][x] ])
    #            ALL_indices.append(j)
                #print(str(j)+' XAF '+' '+str(y)+' '+FX133TradeData[3][x], HqTradeData[0][j], HqTradeData[7][j], HqTradeData[8][j], HqTradeData[9][j], HqTradeData[15][j],(HqTradeData[8][j]-HqTradeData[15][j]))
        #j+=1
    #XAF_dealers = []
    #for x in XAFIndex_get_dealnumbers:
    #    XAF_dealers.append(FX133TradeData[3][x])
    #XOF
    #j = 0
    #for y in DealNumber:
    #    for x in XOFIndex_get_dealnumbers:
    #        if y == FX133TradeData[48][x] and (HqTradeData[8][j]-HqTradeData[15][j]) > timedelta(days = 0):
    #            trade_writer.writerow([ str(j),HqTradeData[0][j],HqTradeData[15][j],'XOF',str(y),FX133TradeData[42][x],HqTradeData[8][j],"{:,.2f}".format(HqTradeData[9][j]),HqTradeData[7][j],(HqTradeData[8][j]-HqTradeData[15][j]),FX133TradeData[3][x] ])
    #            ALL_indices.append(j)

                #print(str(j)+' XOF '+' '+str(y)+' '+FX133TradeData[3][x], HqTradeData[0][j], HqTradeData[7][j], HqTradeData[8][j], HqTradeData[9][j], HqTradeData[15][j],(HqTradeData[8][j]-HqTradeData[15][j]))


        j+=1
output.close()

missing_indices = []
for y in range( 0, len(ALL_indices) ):
    if y not in ALL_indices:
        missing_indices.append(y)
print('\n')
print('Records missing equals: ',len(missing_indices))
print ('See missing records below: ')
for y in missing_indices:
    print(HqTradeData[2][y], HqTradeData[7][y], "{:,}".format( HqTradeData[9][y] ), ' Creation date: ',HqTradeData[15][y], ' Value date: ',HqTradeData[8][y])
print('\n')

x = getPartners
dict = {}

for p in getPartners:
    count = 0
    for q in x:
        if p == q:
            count+=1
    dict[p] = count

for x in dict:
    print(x, dict[x])





#XOF_dealers = []
#for x in XOFIndex_get_dealnumbers:
#    XOF_dealers.append(FX133TradeData[3][x])

#=======================
