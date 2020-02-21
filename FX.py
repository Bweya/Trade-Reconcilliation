import pandas as pd
import os
from datetime import timedelta
import datetime
import xlsxwriter
import calendar

def trade():

    #month = datetime.datetime.now().month
    year = datetime.datetime.now().year
    month_dict = {'Jan':1, 'Feb':2, 'Mar':3, 'Apr':4, 'May':5, 'June':6, 'July':7, 'Aug':8, 'Sep':9, 'Oct':10, 'Nov':11, 'Dec':12}

    k = list(month_dict.keys())
    v = list(month_dict.values())
    #getmonth = k[v.index(month)]
    getmonth = 'Jan'
    month = 1

    workbook = xlsxwriter.Workbook(str(getmonth).upper()+str(year)+"_Report.xlsx")
    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({'bold': True, 'align':'center', 'bg_color':'#A9A9A9', 'border': 1})
    bolds = workbook.add_format({'bold': True, 'font_size':18, 'border': 1})

    worksheet.set_column('A:H', 15)
    worksheet.set_column("I:I", 30)
    worksheet.set_column('J:K', 30)
    worksheet.set_column('L:L', 15)

    worksheet.merge_range('A1:L1', 'FX Trade Delivery Days to UNICEF Country Offices 1-'+str(calendar.monthrange(int(year), month)[1])+' '+getmonth+' '+str(year), bolds)

    worksheet.write('A2', 'A', bold)
    worksheet.write('B2', 'B', bold)
    worksheet.write('C2', 'C', bold)
    worksheet.write('D2', 'D', bold)
    worksheet.write('E2', 'E', bold)
    worksheet.write('F2', 'F', bold)
    worksheet.write('G2', 'G', bold)
    worksheet.write('H2', 'H', bold)
    worksheet.write('I2', 'I', bold)
    worksheet.write('J2', 'J', bold)
    worksheet.write('K2', 'K', bold)
    worksheet.write('L2', 'L', bold )

    filter_column = 'Period after CO value date'

    worksheet.write('A3', 'Item No.', bold)
    worksheet.write('B3', 'CO Request ID', bold)
    worksheet.write('C3', 'Creation Date', bold)
    worksheet.write('D3', 'Value Date CO', bold)
    worksheet.write('E3', 'FX Deal No', bold)
    worksheet.write('F3', 'Deal Amount', bold)
    worksheet.write('G3', 'Currency', bold)
    worksheet.write('H3', 'Value Date HQ', bold)
    worksheet.write('I3', 'Business Partner', bold)
    worksheet.write('J3', filter_column, bold)
    worksheet.write('K3', 'Add 3 average days not received by CO', bold)
    worksheet.write('L3', 'Trader', bold )



    files = os.listdir('files')

    for f in files:
        if f[:7] == 'FX -133':
            FX133 = pd.read_excel('files/'+f, sheet_name = 0, header = None, skiprows = 1)
        #if f[:9] == 'FX Trades':
        #    FXtrades = pd.read_excel('files/'+f, sheet_name = 0, header = None, skiprows = 1)
        if f[:12] == 'FX Trades HQ':
            FXhqrades = pd.read_excel('files/'+f, sheet_name = 0, header = None, skiprows = 1)

    #===========HQ File area

    COrequestID = []
    creationDate = []
    valueDateCO = []
    hqFXDealNumbers = []
    DealAmount = []
    Currency = []
    #ValueDateHQ = []

    count = 0
    for id in FXhqrades[0]:
        if isinstance(id, int) == True:

            COrequestID.append(id)
            creationDate.append(FXhqrades[15][count])
            valueDateCO.append(FXhqrades[8][count])
            hqFXDealNumbers.append(FXhqrades[2][count])
            DealAmount.append(FXhqrades[9][count])
            Currency.append(FXhqrades[7][count])

            count+=1
    #========End HQ File area

    #============Begin FX133 area

    #GET INDEX

    #=======================EUR
    EURstart_index = 0
    for x in FX133[0]:
        if isinstance(x, datetime.datetime) == False:

            if isinstance(x, float) == False and x[:12] == 'EUR ( Euro )':
                break
        EURstart_index+=1

    EURstop_index = 0

    for x in FX133[0]:
        if EURstop_index > EURstart_index:
            if isinstance(x, datetime.datetime) == False:

                if isinstance(x, float) == False and x[:14] == 'Total Currency':
                    break
        EURstop_index+=1

    EURstart_index = EURstart_index + 1

    #usd_get_dealnumbers = []
    EURIndex_get_DealNumbers = []
    count_EUR_index = 0
    for x in FX133[48]:
        if count_EUR_index > EURstart_index and count_EUR_index < EURstop_index:
            if isinstance(x, int) == True:
                #usd_get_dealnumbers.append(x)
                EURIndex_get_DealNumbers.append(count_EUR_index)
        count_EUR_index+=1


    #=======================================================

    #=======================USD
    USDstart_index = 0
    for x in FX133[0]:
        if isinstance(x, datetime.datetime) == False:

            if isinstance(x, float) == False and x[:17] == 'USD ( US Dollar )':
                break
        USDstart_index+=1

    USDstop_index = 0

    for x in FX133[0]:
        if USDstop_index > USDstart_index:
            if isinstance(x, datetime.datetime) == False:


                if isinstance(x, float) == False and x[:14] == 'Total Currency':
                    break;
        USDstop_index+=1

    USDstart_index = USDstart_index + 1


    USDIndex_get_DealNumbers = []
    count_usd_index = 0
    for x in FX133[48]:
        if count_usd_index > USDstart_index and count_usd_index < USDstop_index:
            if isinstance(x, int) == True:

                USDIndex_get_DealNumbers.append(count_usd_index)
        count_usd_index+=1

    j = 0
    c = 0
    row_record = 4
    #missingUS = []
    #missingEU = []
    ALL_indices = []

    for check in hqFXDealNumbers:
        for i in USDIndex_get_DealNumbers:

            if check == FX133[48][i] and FXhqrades[7][j] != 'DKK':
                value = FX133[32][i].replace('.','/')
                valueDateFX133 = datetime.datetime.strptime(value, '%d/%m/%Y')
                periodUS = ( valueDateFX133.date() - valueDateCO[j].date() ).days

                if (valueDateFX133.date()).weekday() >= 0 and (valueDateCO[j].date()).weekday() < 5 and (valueDateFX133.date()).isocalendar()[1] > (valueDateCO[j].date()).isocalendar()[1]:

                    periodUS = periodUS-2

                if(periodUS+3) > 5:
                    the_columns = workbook.add_format({'align':'center', 'bg_color': '#ffff00', 'border': 1})
                    date_format = workbook.add_format({'num_format': 'mm/dd/yy', 'align':'center', 'bg_color': '#ffff00', 'border': 1})

                if(periodUS+3) <= 5:
                    the_columns = workbook.add_format({'align':'center', 'border': 1})
                    date_format = workbook.add_format({'num_format': 'mm/dd/yy', 'align':'center', 'border': 1})

                worksheet.write('A'+str(row_record), c+1, the_columns)
                worksheet.write('B'+str(row_record), COrequestID[j], the_columns)
                worksheet.write('C'+str(row_record), creationDate[j].date(), date_format)
                worksheet.write('D'+str(row_record), valueDateCO[j].date(), date_format)
                worksheet.write('E'+str(row_record), check, the_columns)
                worksheet.write('F'+str(row_record), "{:,.2f}".format(DealAmount[j]), the_columns)
                worksheet.write('G'+str(row_record), Currency[j], the_columns)
                worksheet.write('H'+str(row_record), valueDateFX133.date(), date_format)
                worksheet.write('I'+str(row_record), FX133[41][i], the_columns)
                worksheet.write('J'+str(row_record), str(periodUS)+ ' days', the_columns)
                worksheet.write('K'+str(row_record), str(periodUS+3)+' days', the_columns)
                worksheet.write('L'+str(row_record), FX133[3][i], the_columns )

                row_record += 1
                c += 1
                ALL_indices.append(j)


        for i in EURIndex_get_DealNumbers:

            if check == FX133[48][i] and FXhqrades[7][j] != 'DKK':
                value = FX133[32][i].replace('.','/')
                valueDateFX133 = datetime.datetime.strptime(value, '%d/%m/%Y')
                periodEU = ( valueDateFX133.date() - valueDateCO[j].date() ).days

                if (valueDateFX133.date()).weekday() >= 0 and (valueDateCO[j].date()).weekday() < 5 and (valueDateFX133.date()).isocalendar()[1] > (valueDateCO[j].date()).isocalendar()[1]:

                    periodEU = periodEU-2

                if(periodEU+3) > 5:
                    the_columns = workbook.add_format({'align':'center', 'bg_color': '#ffff00', 'border': 1})
                    date_format = workbook.add_format({'num_format': 'mm/dd/yy', 'align':'center', 'bg_color': '#ffff00', 'border': 1})

                if(periodEU+3) <= 5:
                    the_columns = workbook.add_format({'align':'center', 'border': 1})
                    date_format = workbook.add_format({'num_format': 'mm/dd/yy', 'align':'center', 'border': 1})

                worksheet.write('A'+str(row_record), c+1, the_columns)
                worksheet.write('B'+str(row_record), COrequestID[j], the_columns)
                worksheet.write('C'+str(row_record), creationDate[j].date(), date_format)
                worksheet.write('D'+str(row_record), valueDateCO[j].date(), date_format)
                worksheet.write('E'+str(row_record), check, the_columns)
                worksheet.write('F'+str(row_record), "{:,.2f}".format(DealAmount[j]), the_columns)
                worksheet.write('G'+str(row_record), Currency[j], the_columns)
                worksheet.write('H'+str(row_record), valueDateFX133.date(), date_format)
                worksheet.write('I'+str(row_record), FX133[41][i], the_columns)
                worksheet.write('J'+str(row_record), str(periodEU)+ ' days', the_columns)
                worksheet.write('K'+str(row_record), str(periodEU+3)+' days', the_columns)
                worksheet.write('L'+str(row_record), FX133[3][i], the_columns )

                row_record += 1
                c += 1
                ALL_indices.append(j)

        j+=1
    worksheet.merge_range('J'+str(row_record+1)+':L'+str(row_record+1), "Compiled by: Louisa Tinga - Treasury Unit")


    workbook.close()

    print(' File constructed successfully')

    missing_indices = []
    for y in range( 0, ( ALL_indices[-1] + 1) ):
        if y not in ALL_indices:
            missing_indices.append(y)
    print('\n')
    print('Records missing equals: ',len(missing_indices))
    print ('See missing records below: ')
    for y in missing_indices:
        print('Missing index: ',y,'; Deal Number: ',FXhqrades[2][y], '; Currency: ',FXhqrades[7][y],'; Deal Amount: ', "{:,}".format( FXhqrades[9][y] ), '; Creation date: ',FXhqrades[15][y], '; Value date: ',FXhqrades[8][y])
    print('\n')

    #==========End FX133 area


    #print('\n')
    #print( 'Request ID array has been populated successfully. Number of items: ', len(COrequestID) )
    #print('\n')
