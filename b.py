import xlsxwriter
import pandas as pd
import os
import math

def trade():

    files = os.listdir('files')

    year = '2019'

    workbook = xlsxwriter.Workbook(year+'partners.xlsx')
    worksheet = workbook.add_worksheet()

    border = workbook.add_format({'border':1,'align':'center'})
    bolds = workbook.add_format({'bold': True, 'font_size':18, 'border': 1})

    worksheet.set_column('A:C', 45)

    worksheet.merge_range('A1:C1', 'Number of approved FX trades by business partners 01 JAN - 31 DEC '+year, bolds)

    bold = workbook.add_format({'bold':True, 'border':1})
    gtotal =  workbook.add_format({'bold': True, 'align':'center', 'bg_color':'#A9A9A9', 'border': 1})
    footer = workbook.add_format({'border':1})

    worksheet.write("A2", "Business Partners", workbook.add_format({'bold':True, 'align':'center', 'border':1, 'bg_color':'#A9A9A9'}))
    worksheet.write("B2", "Number of FX Trades", workbook.add_format({'bold':True, 'align':'center', 'border':1, 'bg_color':'#A9A9A9'}))
    worksheet.write("C2", "Total Value Add (US $)", workbook.add_format({'bold':True, 'align':'center', 'border':1, 'bg_color':'#A9A9A9'}))

    check_bus = []
    for f in files:

        if f[:3] == '.DS' or f[:1] == 't':

            print('DS File Store')

        else:

            TradeData = pd.read_excel("files/"+f, sheet_name = 0, header = None, skiprows=3)

            for x in TradeData[8]:
                if isinstance(x, str) == True:

                    if x not in check_bus:
                        check_bus.append(x.strip())


    valueAdd_sorted = []
    grand_valueAdd = 0
    grand_count = 0
    row_record = 3


    for business in check_bus:

        count = 0
        valueadd_bp = 0

        for f in files:

            if f[:3] == '.DS' or f[:1] == 't':

                print('DS File Store')

            else:

                TradeData = pd.read_excel("files/"+f, sheet_name = 0, header = None, skiprows=3)
                index = 0
                for x in TradeData[8]:

                    #if isinstance(x, str) == True:






                        #for y in TradeData[8]:
                        #if isinstance(y, str) == True:

                    if x == business:
                        count += 1
                        valueadd_bp += float( (TradeData[16][index]).replace(",","") )
                    index += 1
                        #if x not in check_businessp:

        #check_businessp.append(x)
        #print( 'Business partner: ',business,' Number of transactions: ', count, ' Total ValueAdd: ', "{:,.2f}".format(valueadd_bp) )
        #grand_valueAdd += valueadd_bp
        #grand_count += count

        #=======

        valueAdd_sorted.append(valueadd_bp)

        #======


        #worksheet.write("A"+str(row_record), business, border)
        #worksheet.write("B"+str(row_record), count, border)
        #worksheet.write("C"+str(row_record), "{:,.2f}".format(valueadd_bp), border)

        #row_record += 1


        #print('\n')
        #print(f[:-25],' Grand value add equals: ', "{:,.2f}".format(grand_valueAdd) )

        #START ====================================================================================
    valueAdd_sorted.sort(reverse = True)
    v = valueAdd_sorted
    for val in v:

        for business in check_bus:

            count = 0
            valueadd_bp = 0

            for f in files:

                if f[:3] == '.DS' or f[:1] == 't':

                    print('DS File Store')

                else:

                    TradeData = pd.read_excel("files/"+f, sheet_name = 0, header = None, skiprows=3)
                    index = 0
                    for x in TradeData[8]:

                        #if isinstance(x, str) == True:






                            #for y in TradeData[8]:
                            #if isinstance(y, str) == True:

                        if x == business:
                            count += 1
                            valueadd_bp += float( (TradeData[16][index]).replace(",","") )
                        index += 1
                            #if x not in check_businessp:

            #check_businessp.append(x)


            #=======

            if val ==  valueadd_bp:
                print( 'Business partner: ',business,' Number of transactions: ', count, ' Total ValueAdd: ', "{:,.2f}".format(valueadd_bp) )
                grand_valueAdd += valueadd_bp
                grand_count += count

                worksheet.write("A"+str(row_record), business, border)
                worksheet.write("B"+str(row_record), count, border)
                worksheet.write("C"+str(row_record), "{:,.2f}".format(valueadd_bp), border)

                row_record += 1

            #======


            #worksheet.write("A"+str(row_record), business, border)
            #worksheet.write("B"+str(row_record), count, border)
            #worksheet.write("C"+str(row_record), "{:,.2f}".format(valueadd_bp), border)

            #row_record += 1


                print('\n')
                print(f[:-25],' Grand value add equals: ', "{:,.2f}".format(grand_valueAdd) )

            #STOP ====================================================================================

    worksheet.write("A"+str(row_record + 1), "TOTAL", gtotal)
    worksheet.write("B"+str(row_record + 1), "{:,.2f}".format(grand_count), gtotal)
    worksheet.write("C"+str(row_record + 1), "{:,.2f}".format(grand_valueAdd), gtotal)
    worksheet.merge_range('A'+str(row_record+2)+':C'+str(row_record+2), "Compiled by: Louisa Tinga - Treasury Unit", footer)

    workbook.close()
    print(v)
