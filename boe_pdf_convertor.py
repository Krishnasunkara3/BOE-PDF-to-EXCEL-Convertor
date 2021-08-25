# IMPORTING NECESSARY LIBRARIES
import pandas as pd
import camelot
import re
from pdfminer.high_level import extract_text
import PyPDF2
import glob
import logging

from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from io import StringIO

# CREATING THE EXCEL FILE WITH THE HEADERS
path = r'C:\Users\hp\Desktop\Pyspark\BoePdfConvertor\PDF FILES\BOE.xlsx'

try:
    final_df1 = pd.read_excel(path, sheet_name="BOE_SUMMARY")
    final_df2 = pd.read_excel(path, sheet_name="Valuation_Details")
    final_df3 = pd.read_excel(path, sheet_name="Duties")
    final_df4 = pd.read_excel(path, sheet_name="Invoice_Details")

except:
    final_df1 = pd.DataFrame()
    final_df2 = pd.DataFrame()
    final_df3 = pd.DataFrame()
    final_df4 = pd.DataFrame()
    with pd.ExcelWriter(path) as writer:
        final_df1.to_excel(writer, sheet_name='BOE_SUMMARY', index=False)
        final_df2.to_excel(writer, sheet_name='Valuation_Details',index=False)
        final_df3.to_excel(writer, sheet_name='Duties',index=False)
        final_df4.to_excel(writer, sheet_name='Invoice_Details', index=False)

# DEFINING LOG PATH AND CREATING LOG FILE
log_path = r'C:\Users\hp\Desktop\Pyspark\BoePdfConvertor\PDF FILES\myLog.log'
logging.basicConfig(filename=log_path, filemode='w', format='%(name)s - %(levelname)s - %(message)s')
with open(log_path, 'w'):
    pass
logging.warning('--------------------------------------------------')

# GETTING DATA FROM PDF'S
mypath = r'C:\Users\hp\Desktop\Pyspark\BoePdfConvertor\PDF FILES'
for file in glob.glob(mypath + '\*.pdf'):
    text = extract_text(file, page_numbers=[0])
    # print(text)
    table = camelot.read_pdf(file, pages='1')
    list1 = text.split('\n')
    df = pd.DataFrame(table[0].df)

    sum_type = ''.join(df.iloc[8][0].split(' ')[-4:-1])
    if sum_type != 'BILLOFENTRY':
        logging.warning(f'THE {file} is not a BOE Type')
        continue

    text3 = extract_text(file, page_numbers=[3])
    # print(text)
    table3 = camelot.read_pdf(file, pages='4')
    table3
    list3 = text3.split('\n')
    sl_no = list3.index('1. S NO')
    S_NO = list3[sl_no + 1]
    S_NO
    inv_no = list3.index('2. INVOICE NO')
    INVOICE_NO = list3[inv_no + 1]
    INVOICE_NO
    inv_amt = list3.index('3. INVOICE AMOUNT')
    INVOICE_AMOUNT = list3[inv_amt + 1]
    INVOICE_AMOUNT
    cur = list3.index('4. CUR')
    CURRANCY = list3[cur + 1]
    CURRANCY
    lis5 = []
    lis5.insert(0, S_NO)
    lis5.insert(1, INVOICE_NO)
    lis5.insert(2, INVOICE_AMOUNT)
    lis5.insert(3, CURRANCY)
    if not final_df4.empty:
        inv_list = final_df4['2.INVOICE NO'].tolist()
        if INVOICE_NO in inv_list:
            logging.warning(f'This Invoice {INVOICE_NO} duplicate is already exist')
            continue

    #     ----------------------------------------------------------
    # ***********************************************************
    # FETCHING TOP DETAILS
    # ***********************************************************
    Data = df[0][0]
    Data = Data.split('\n')
    #     print(Data)
    Data1 = df[10][1]
    Data1 = Data1.split('\n')
    #     print(Data1)
    Data2 = df[10][2]
    Data2 = Data2.split('\n')
    Data2.pop()
    #     print(Data2)
    Data3 = df[10][3]
    Data3 = Data3.split('\n')
    #     print(Data3)
    Data4 = df[10][4]
    Data4 = Data4.split('\n')
    #     print(Data4)
    Data5 = df[10][6]
    Data5 = Data5.split('\n')
    #     print(Data5)
    Data6 = df[10][7]
    Data6 = Data6.split('\n')
    #     print(Data6)
    result = []
    Data1.extend(Data2)
    Data1.extend(Data3)
    Data1.extend(Data4)
    Data1.extend(Data5)
    Data1.insert(0, Data[1])
    Data1.append(Data6[0])
    Data1.append(Data6[3])
    #     print(Data1)
    # **************************************************************
    # Getting Acolumn Data
    # **************************************************************
    Status = df[1][10]
    Status1 = Status.split('\n')
    Status1
    Status2 = df[8][10]
    Status3 = Status2.split('\n')
    Status3
    Status4 = df[1][11]
    Status5 = Status4.split('\n')
    Status5
    Status6 = df[8][11]
    Status7 = Status6.split('\n')
    Status7
    Status8 = df[1][12]
    Status9 = Status8.split('\n')
    Status9
    Status10 = df[8][12]
    Status11 = Status10.split('\n')
    Status11
    Status1.extend(Status3)
    Status1.append(Status5[1])
    Status1.append(Status7[1])
    Status1.append(Status9[1])
    Status1.append(Status11[1])
    Status1
    # *********************************************************************
    # Getting Bcolumn DATA
    # *********************************************************************
    Declarant = df[1][14] + df[1][15] + df[1][16] + df[1][17] + ' ' + df[1][18]
    #     print(Declarant)
    Declarant1 = df[8][14]
    Declarant2 = Declarant1.split('\n')[1]
    #     print(Declarant2)
    Declarant3 = df[3][19]
    #     print(Declarant3)

    Fin_Declarant = []
    Fin_Declarant.append(Declarant)
    Fin_Declarant.append(Declarant2)
    Fin_Declarant.append(Declarant3)
    Fin_Declarant

    # **********************************************************
    # Getting c data
    # ***********************************************************
    c_data = df.iloc[21].tolist()
    # print(c_data)
    BCD = c_data[1]
    BCD
    ACD = c_data[3]
    ACD
    SWS = c_data[4]
    SWS
    NCCD = c_data[6]
    NCCD
    ADD = c_data[7]
    ADD
    # Getting data from pdfminer
    cvd_index = list1.index('6.CVD')
    CVD = list1[cvd_index + 1]
    IGST = c_data[8].split('\n')[0]
    G_CESS = c_data[8].split('\n')[-2]
    TOTAL_ASS_VAL = c_data[8].split('\n')[-1]

    c_data1 = df.iloc[23].tolist()
    #     print(c_data1)
    for i in c_data1:
        if len(i) > 12:
            TOT_AMOUNT = i.split('\n')[-2]
            FINE = i.split('\n')[-3]
            PNLTY = i.split('\n')[-4]
            INT = i.split('\n')[-5]
            break

    SG = c_data1[1]
    SAED = c_data1[3]
    SAED
    GSIA = c_data1[4]
    GSIA
    TTA = c_data1[6]
    TTA
    HEALTH = c_data1[7]
    HEALTH
    Amt_index = list1.index('1.SR NO 2.CHALLAN NO 3.PAID ON 4.AMOUNT(Rs.)')
    TOTAL_DUTY = list1[Amt_index + 6]
    LIS = []
    LIS.insert(0, BCD)
    LIS.insert(1, ACD)
    LIS.insert(2, SWS)
    LIS.insert(3, NCCD)
    LIS.insert(4, ADD)
    LIS.insert(5, CVD)
    LIS.insert(6, IGST)
    LIS.insert(7, G_CESS)
    LIS.insert(8, SG)
    LIS.insert(9, SAED)
    LIS.insert(10, GSIA)
    LIS.insert(11, TTA)
    LIS.insert(12, HEALTH)
    LIS.insert(13, TOTAL_DUTY)
    LIS.insert(14, INT)
    LIS.insert(15, PNLTY)
    LIS.insert(16, FINE)
    LIS.insert(17, TOTAL_ASS_VAL)
    LIS.insert(18, TOT_AMOUNT)
    #     print(LIS)

    # ****************************************************************
    # Getting D column data
    # ****************************************************************
    d_data = df.iloc[26].tolist()
    #     print(d_data)
    GW = d_data[-4].split('\n')[-1]
    #     print(GW)
    PKG = d_data[-4].split('\n')[-2]
    #     print(PKG)
    # HAWB_NO=d_data[-8:-5]
    # HAWB_NO
    dummy = d_data[-8:-5]
    for i in dummy:
        if len(i) > 5:
            dummy1 = i
    HAWB_NO = dummy1.split(' ')[-1]
    #     print(HAWB_NO)
    DATE = d_data[-8].split(' ')[0]
    #     print(DATE)
    data = d_data[8:10]
    data
    for i in data:
        if len(i) > 3:
            data1 = i
    MAWB_NO = data1.split('\n')[0]
    MAWB_NO1 = df[8][27]
    if len(MAWB_NO1) > 4:
        MAWB_NO = MAWB_NO + MAWB_NO1
    #     print(MAWB_NO)
    GIGM_NO = d_data[6]
    INW_Date = d_data[4]
    IGM_Date = d_data[3]
    IGM_NO = d_data[1]
    Lis1 = []
    Lis1.insert(0, IGM_NO)
    Lis1.insert(1, IGM_Date)
    Lis1.insert(2, INW_Date)
    Lis1.insert(3, GIGM_NO)
    Lis1.insert(4, MAWB_NO)
    Lis1.insert(5, DATE)
    Lis1.insert(6, HAWB_NO)
    Lis1.insert(7, PKG)
    Lis1.insert(8, GW)
    #     print(Lis1)

    # *************************************************************************
    # getting f details
    # *************************************************************************
    #     f_details=df.iloc[29]
    #     f_details

    #     challan_no=f_details[11]
    #     paid_on=f_details[13]
    #     Amount=f_details[16]
    sr_no = 1
    Amount = list1[Amt_index + 6]
    Paid_on = list1[Amt_index + 4]
    challan_no = list1[Amt_index + 2]
    # print(Challanno,Paidon,Amount)
    f_list = []
    f_list.insert(0, sr_no)
    f_list.insert(1, challan_no)
    f_list.insert(2, Paid_on)
    f_list.insert(3, Amount)
    #     print(f_list)

    # *************************************************************
    # GETTING i DETAILS
    # *************************************************************
    check = df[9][32]
    #     print(check)
    if check == '1':
        s_no = df[9][32]
        invoice_no = df[11][32]
        inv_amt = df[15][32]
        cur = df[17][32]
    else:
        s_no = df[9][33]
        invoice_no = df[11][33]
        inv_amt = df[15][33]
        cur = df[17][33]

    i_list = []
    i_list.insert(0, s_no)
    i_list.insert(1, invoice_no)
    i_list.insert(2, inv_amt)
    i_list.insert(3, cur)
    #     print(i_list)
    # **************************************************************************
    # getting h details
    # **************************************************************************

    sub_details = df.iloc[34]
    assess_details = df.iloc[35]

    #     print(sub_details)

    check = df[1][34]
    #     print(check)
    if check == 'Submission':
        sub_event = df[1][34]
        sub_date = df[3][34]
        sub_time = df[4][34]
        sub_exchange_rate = df[5][34].split('\n')[0]
    else:
        sub_event = df[1][35]
        sub_date = df[3][35]
        sub_time = df[4][35]
        sub_exchange_rate = df[5][35].split('\n')[0]
    check = df[1][35]
    if check == 'Assessment':
        asses_event = df[1][35]
        asses_date = df[3][35]
        asses_time = df[4][35]
        asses_exchange_rate = df[5][35].split('\n')[0]
        if len(asses_exchange_rate) < 8:
            asses_exchange_rate = ''

    else:
        asses_event = df[1][36]
        asses_date = df[3][36]
        asses_time = df[4][36]
        asses_exchange_rate = df[5][36].split('\n')[0]

    h_list = []
    h_list.insert(0, sub_event)
    h_list.insert(1, sub_date)
    h_list.insert(2, sub_time)
    h_list.insert(3, sub_exchange_rate)
    h_list.insert(4, asses_event)
    h_list.insert(5, asses_date)
    h_list.insert(6, asses_time)
    h_list.insert(7, asses_exchange_rate)
    #     print(h_list)

    # **************************************************************
    # OOC DETAILS
    # **************************************************************
    check = df[8][41]
    check1 = df[8][42]
    # print(check)
    if check == 'OOC NO.':
        occno = df[11][41]
        occdate = df[11][42]
    elif check1 == 'OOC NO.':
        occno = df[11][42]
        occdate = df[11][43]
    else:
        occno = df[11][43]
        occdate = df[11][44]

    check = df[1][37]
    if check == 'OOC':
        ooctime = df[4][37]
    else:
        ooctime = df[4][38]

    OCC_LIST = []
    OCC_LIST.insert(0, occno)
    OCC_LIST.insert(1, occdate)
    OCC_LIST.insert(2, ooctime)

    Data1.extend(Status1)
    Data1.extend(Fin_Declarant)
    Data1.extend(LIS)
    Data1.extend(Lis1)
    Data1.extend(f_list)
    Data1.extend(h_list)
    Data1.extend(i_list)
    Data1.extend(OCC_LIST)
    Data1
    # *************************************************************
    # *************************************************************
    # Getting page2 data
    text1 = extract_text(file, page_numbers=[1])
    # print(text)
    table1 = camelot.read_pdf(file, pages='2')
    table1
    list1 = text1.split('\n')
    df1 = pd.DataFrame(table1[0].df)
    #     print(file)
    # **************************************************
    # A.Invoice
    # ***************************************************
    a_data = df1.iloc[11][1]
    a_data
    sn_no = a_data.split('\n')[0]
    invoiceno = a_data.split('\n')[1]
    date = df1.iloc[12][1]
    invoicenumberanddate = invoiceno + ' ' + date
    invoicenumberanddate
    a_list = []
    a_list.insert(0, sn_no)
    a_list.insert(1, invoicenumberanddate)
    #     print(a_list)
    # ******************************************************
    # B.Transacting parties
    # *******************************************************
    buyersnameandaddress = df1.iloc[14][1] + ' ' + df1.iloc[15][1] + ' ' + df1.iloc[16][1] + df1.iloc[17][1] + ' ' + \
                           df1.iloc[18][1]
    buyersnameandaddress
    suppliernameandaddress = df1.iloc[20][1] + df1.iloc[21][1] + df1.iloc[22][1] + ' ' + df1.iloc[23][1] + ' ' + \
                             df1.iloc[24][1]
    suppliernameandaddress
    #     adcode=df1.iloc[25][6].split('\n')[1]
    #     adcode
    check = df1[6][25].split('\n')[0]
    #     print(check)
    if check == '6. AD CODE':

        adcode = df1[6][25].split('\n')[1]
    else:
        adcode = df1[7][25]
    adcode
    # df1
    b_list = []
    b_list.insert(0, buyersnameandaddress)
    b_list.insert(1, suppliernameandaddress)
    b_list.insert(1, adcode)

    # *****************************************************
    # c data
    # *****************************************************
    c_data = df1.iloc[27]
    invvalue = c_data[1]
    freight = c_data[2]
    insurance = c_data[3]
    hss = c_data[4]
    loading = c_data[5]
    commn = c_data[6]
    try:
        payterms = c_data[7].split('\n')[0]
        valuationmethod = c_data[7].split('\n')[1]
    except:
        payterms = c_data[8].split('\n')[0]
        valuationmethod = c_data[8].split('\n')[1]

    c_data2 = df1.iloc[29]
    reltd = c_data2[8]
    svbch = c_data2[9]
    svbno = c_data2[10]
    try:
        date = c_data2[11].split('\n')[0]
        loa = c_data2[11].split('\n')[1]
    except:
        date = c_data2[11]
        loa = c_data2[11]
    cur = df1[1][28].split('\n')[1]
    term = c_data2[1].split('\n')[1]

    c_list = []
    c_list.insert(0, invvalue)
    c_list.insert(1, freight)
    c_list.insert(2, insurance)
    c_list.insert(3, hss)
    c_list.insert(4, loading)
    c_list.insert(5, commn)
    c_list.insert(6, payterms)
    c_list.insert(7, valuationmethod)
    c_list.insert(8, reltd)
    c_list.insert(9, svbch)
    c_list.insert(10, svbno)
    c_list.insert(11, date)
    c_list.insert(12, loa)
    c_list.insert(13, cur)
    c_list.insert(14, term)
    #     print(c_list)

    # *****************************************************
    # dlist
    # ******************************************************
    mischarge = df1[10][33]
    assvalue = df1[11][33]
    d_list = []
    d_list.insert(0, mischarge)
    d_list.insert(1, assvalue)
    #     print(d_list)
    data = df1.iloc[35]
    data1 = data[4].split('\n')
    data2 = ''.join(data1)
    #     print(data2)

    # **********************************************************
    # E data
    # **********************************************************
    e_data = df1.iloc[35]
    snno = e_data[1].split('\n')[0]
    try:
        cth = e_data[1].split('\n')[1]
    except:
        cth = e_data[2]
    data = df1.iloc[35]
    if len(data[3]) > 10:
        data1 = data[3].split('\n')
        description = ''.join(data1)
    else:
        data1 = data[4].split('\n')
        description = ''.join(data1)
    description
    if len(e_data[6]) > 5:
        unitprice = e_data[6]
    else:
        unitprice = e_data[7]
    if len(e_data[9]) > 5:
        quantity = e_data[9].split('\n')[0]
    else:
        quantity = e_data[10]
    if len(e_data[9]) > 5:
        uqc = e_data[9].split('\n')[1]
    else:
        uqc = e_data[12]
    if len(e_data[11]) > 5:
        amount = e_data[11]
    else:
        amount = e_data[13]
    # amount

    e_list = []
    e_list.insert(0, snno)
    e_list.insert(1, cth)
    e_list.insert(2, description)
    e_list.insert(3, unitprice)
    e_list.insert(4, quantity)
    e_list.insert(5, uqc)
    e_list.insert(6, amount)
    #     print(e_list)
    e_data1 = df1.iloc[36]
    if e_data1[1].split('\n')[0] == '2':
        snno1 = e_data1[1].split('\n')[0]
        cth1 = e_data1[1].split('\n')[1]
        data1 = e_data1[3].split('\n')
        description1 = ''.join(data1)
        unitprice1 = e_data1[6]
        quantity1 = e_data1[9].split('\n')[0]
        uqc1 = e_data1[9].split('\n')[1]
        amount1 = e_data1[11]

    else:
        snno1 = ''
        cth1 = ''
        data1 = ''
        description1 = ''
        unitprice1 = ''
        quantity1 = ''
        uqc1 = ''
        amount1 = ''

    e_list1 = []
    e_list1.insert(0, snno1)
    e_list1.insert(1, cth1)
    e_list1.insert(2, description1)
    e_list1.insert(3, unitprice1)
    e_list1.insert(4, quantity1)
    e_list1.insert(5, uqc1)
    e_list1.insert(6, amount1)
    #     print(e_list1)

    final_list = []
    final_list.extend(a_list)
    final_list.extend(b_list)
    final_list.extend(c_list)
    final_list.extend(d_list)
    final_list.extend(e_list)
    final_list.extend(e_list1)
    #     print(len(final_list))

    # **************************************************
    # A.Invoice
    # ***************************************************
    a_data = df1.iloc[11][1]
    a_data
    sn_no = a_data.split('\n')[0]
    invoiceno = a_data.split('\n')[1]
    date = df1.iloc[12][1]
    invoicenumberanddate = invoiceno + ' ' + date
    invoicenumberanddate
    a_list = []
    a_list.insert(0, sn_no)
    a_list.insert(1, invoicenumberanddate)
    #     print(a_list)
    # ******************************************************
    # B.Transacting parties
    # *******************************************************
    buyersnameandaddress = df1.iloc[14][1] + ' ' + df1.iloc[15][1] + ' ' + df1.iloc[16][1] + df1.iloc[17][1] + ' ' + \
                           df1.iloc[18][1]
    buyersnameandaddress
    suppliernameandaddress = df1.iloc[20][1] + df1.iloc[21][1] + df1.iloc[22][1] + ' ' + df1.iloc[23][1] + ' ' + \
                             df1.iloc[24][1]
    suppliernameandaddress
    #     adcode=df1.iloc[25][6].split('\n')[1]
    #     adcode
    check = df1[6][25].split('\n')[0]
    #     print(check)
    if check == '6. AD CODE':

        adcode = df1[6][25].split('\n')[1]
    else:
        adcode = df1[7][25]
    adcode
    # df1
    b_list = []
    b_list.insert(0, buyersnameandaddress)
    b_list.insert(1, suppliernameandaddress)
    b_list.insert(1, adcode)

    # *****************************************************
    # c data
    # *****************************************************
    c_data = df1.iloc[27]
    invvalue = c_data[1]
    freight = c_data[2]
    insurance = c_data[3]
    hss = c_data[4]
    loading = c_data[5]
    commn = c_data[6]
    try:
        payterms = c_data[7].split('\n')[0]
        valuationmethod = c_data[7].split('\n')[1]
    except:
        payterms = c_data[8].split('\n')[0]
        valuationmethod = c_data[8].split('\n')[1]

    c_data2 = df1.iloc[29]
    reltd = c_data2[8]
    svbch = c_data2[9]
    svbno = c_data2[10]
    try:
        date = c_data2[11].split('\n')[0]
        loa = c_data2[11].split('\n')[1]
    except:
        date = c_data2[11]
        loa = c_data2[11]
    cur = df1[1][28].split('\n')[1]
    term = c_data2[1].split('\n')[1]

    c_list = []
    c_list.insert(0, invvalue)
    c_list.insert(1, freight)
    c_list.insert(2, insurance)
    c_list.insert(3, hss)
    c_list.insert(4, loading)
    c_list.insert(5, commn)
    c_list.insert(6, payterms)
    c_list.insert(7, valuationmethod)
    c_list.insert(8, reltd)
    c_list.insert(9, svbch)
    c_list.insert(10, svbno)
    c_list.insert(11, date)
    c_list.insert(12, loa)
    c_list.insert(13, cur)
    c_list.insert(14, term)
    #     print(c_list)

    # *****************************************************
    # dlist
    # ******************************************************
    mischarge = df1[10][33]
    assvalue = df1[11][33]
    d_list = []
    d_list.insert(0, mischarge)
    d_list.insert(1, assvalue)
    #     print(d_list)
    data = df1.iloc[35]
    data1 = data[4].split('\n')
    data2 = ''.join(data1)
    #     print(data2)

    # **********************************************************
    # E data
    # **********************************************************
    e_data = df1.iloc[35]
    snno = e_data[1].split('\n')[0]
    try:
        cth = e_data[1].split('\n')[1]
    except:
        cth = e_data[2]
    data = df1.iloc[35]
    if len(data[3]) > 10:
        data1 = data[3].split('\n')
        description = ''.join(data1)
    else:
        data1 = data[4].split('\n')
        description = ''.join(data1)
    description
    if len(e_data[6]) > 5:
        unitprice = e_data[6]
    else:
        unitprice = e_data[7]
    if len(e_data[9]) > 5:
        quantity = e_data[9].split('\n')[0]
    else:
        quantity = e_data[10]
    if len(e_data[9]) > 5:
        uqc = e_data[9].split('\n')[1]
    else:
        uqc = e_data[12]
    if len(e_data[11]) > 5:
        amount = e_data[11]
    else:
        amount = e_data[13]
    # amount

    e_list = []
    e_list.insert(0, snno)
    e_list.insert(1, cth)
    e_list.insert(2, description)
    e_list.insert(3, unitprice)
    e_list.insert(4, quantity)
    e_list.insert(5, uqc)
    e_list.insert(6, amount)
    #     print(e_list)
    e_data1 = df1.iloc[36]
    if e_data1[1].split('\n')[0] == '2':
        snno1 = e_data1[1].split('\n')[0]
        cth1 = e_data1[1].split('\n')[1]
        data1 = e_data1[3].split('\n')
        description1 = ''.join(data1)
        unitprice1 = e_data1[6]
        quantity1 = e_data1[9].split('\n')[0]
        uqc1 = e_data1[9].split('\n')[1]
        amount1 = e_data1[11]

    else:
        snno1 = ''
        cth1 = ''
        data1 = ''
        description1 = ''
        unitprice1 = ''
        quantity1 = ''
        uqc1 = ''
        amount1 = ''

    e_list1 = []
    e_list1.insert(0, snno1)
    e_list1.insert(1, cth1)
    e_list1.insert(2, description1)
    e_list1.insert(3, unitprice1)
    e_list1.insert(4, quantity1)
    e_list1.insert(5, uqc1)
    e_list1.insert(6, amount1)
    #     print(e_list1)

    final_list = []
    final_list.extend(a_list)
    final_list.extend(b_list)
    final_list.extend(c_list)
    final_list.extend(d_list)
    final_list.extend(e_list)
    final_list.extend(e_list1)
    #     print(final_list)

    # ************************************************************
    # ************************************************************
    # Getting page 3 data
    text2 = extract_text(file, page_numbers=[2])
    #     print(text)
    table2 = camelot.read_pdf(file, pages='3')
    table2
    list2 = text2.split('\n')
    df2 = pd.DataFrame(table2[0].df)


    #     print(file)
    #     #b.Item Duty
    def convert_pdf_to_txt(file, codec='utf-8'):
        rsrcmgr = PDFResourceManager()
        retstr = StringIO()
        laparams = LAParams()
        device = TextConverter(rsrcmgr, retstr, codec=codec, laparams=laparams)
        fp = open(file, 'rb')
        interpreter = PDFPageInterpreter(rsrcmgr, device)
        password = ""
        maxpages = 0
        caching = True
        pagenos = set()
        for page in PDFPage.get_pages(fp, pagenos, maxpages=maxpages, password=password, caching=caching,
                                      check_extractable=True):
            interpreter.process_page(page)

        text = retstr.getvalue()

        fp.close()
        device.close()
        retstr.close()
        return text


    data = convert_pdf_to_txt(file, codec='utf-8')
    data1 = data.split('\n')

    # *********************************************************************************
    # Getting A data
    # A item details
    a_details = df2.iloc[11]
    # print(a_details)
    incsno = a_details[1]
    itemsn = a_details[2]
    cth = a_details[3]
    ceth = a_details[4]
    itemdescription = a_details[5].split('\n')[0] + list2[list2.index('5.ITEM DESCRIPTION') + 3]
    fs = a_details[5].split('\n')[1]
    pq = a_details[5].split('\n')[2]
    dc = a_details[5].split('\n')[3]
    wc = a_details[5].split('\n')[4]
    aq = a_details[5].split('\n')[5]

    a_details1 = df2.iloc[13]
    # print(a_details1)
    upi = a_details1[1]
    coo = a_details1[2]
    c_qty = a_details1[3]
    c_uqc = a_details1[4]
    s_qty = a_details1[5].split('\n')[0]
    s_uqc = a_details1[5].split('\n')[1]
    sch = a_details1[6]
    stndpr = a_details1[7]
    rsp = a_details1[9].split('\n')[0]
    reimp = a_details1[9].split('\n')[1]
    prov = a_details1[10]
    enduse = a_details1[11]

    a_details2 = df2.iloc[15]
    #     print(a_details2)
    prodn = a_details2[1]
    cntrl = a_details2[2]
    qualfr = a_details2[3]
    contnt = a_details2[4]
    stmnt = a_details2[5]
    supdocs = a_details2[6]
    assessvalue = a_details2[9]
    totalduty = a_details2[11]

    list_a = []
    list_a.insert(0, incsno)
    list_a.insert(1, itemsn)
    list_a.insert(2, cth)
    list_a.insert(3, ceth)
    list_a.insert(4, itemdescription)
    list_a.insert(5, fs)
    list_a.insert(6, pq)
    list_a.insert(7, dc)
    list_a.insert(8, wc)
    list_a.insert(9, aq)
    list_a.insert(10, upi)
    list_a.insert(11, coo)
    list_a.insert(12, c_qty)
    list_a.insert(13, c_uqc)
    list_a.insert(14, s_qty)
    list_a.insert(15, s_uqc)
    list_a.insert(16, sch)
    list_a.insert(17, stndpr)
    list_a.insert(18, rsp)
    list_a.insert(19, reimp)
    list_a.insert(20, prov)
    list_a.insert(21, enduse)
    list_a.insert(22, prodn)
    list_a.insert(23, cntrl)
    list_a.insert(24, qualfr)
    list_a.insert(25, contnt)
    list_a.insert(26, stmnt)
    list_a.insert(27, supdocs)
    list_a.insert(28, assessvalue)
    list_a.insert(29, totalduty)
    #     print(list_a)

    # ********************************************************************************
    # Getting B data
    data3 = data1.index('5.IGST')
    data3
    #     igst_notnno=data1[data3+1]
    #     igst_notnsno=data1[data3+2]
    #     igst_rate=data1[data3+3]
    #     igst_amount=data1[data3+4]
    #     igst_dutyfg=data1[data3+5]

    data2 = data1.index('6.G. CESS')
    data2
    #     gcess=data1[data2+1]
    #     gcess_notnsno=data1[data2+2]
    #     gcess_rate=data1[data2+3]
    #     gcess_amount=data1[data2+4]
    #     gcess_dutyfg=data1[data2+5]

    #     lis5=[]
    #     lis5.insert(0,igst_notnno)
    #     lis5.insert(1,igst_notnsno)
    #     print(lis5)
    b_details = df2.iloc[17]
    #     b_details
    bcd = b_details[2]
    acd = b_details[3]
    sws = b_details[4]
    sad = b_details[5]
    igst = data1[data3 + 1]
    gcess = data1[data2 + 1]
    #     igst=b_details[6]
    #     gcess=b_details[7].split('\n')[0]
    add = b_details[8]
    cvd = b_details[9]
    sg = b_details[10]
    tvalue = b_details[11]
    lis = []
    lis.insert(0, bcd)
    lis.insert(1, acd)
    lis.insert(2, sws)
    lis.insert(3, sad)
    lis.insert(4, igst)
    lis.insert(5, gcess)
    lis.insert(6, add)
    lis.insert(7, cvd)
    lis.insert(8, sg)
    lis.insert(9, tvalue)
    #     print(lis)
    #     print(file)
    b_details1 = df2.iloc[18]
    b_details2 = df2.iloc[19]
    b_details3 = df2.iloc[20]
    b_details4 = df2.iloc[21]
    #     print(b_details1.tolist())
    #     print(b_details2.tolist())
    #     print(b_details3.tolist())
    #     print(b_details4.tolist())

    sno_bcd = b_details1[2]
    sno_acd = b_details1[3]
    sno_sws = b_details1[4]
    sno_sad = b_details1[5]
    sno_igst = data1[data3 + 2]
    #     sno_igst=b_details1[6]
    sno_gcess = data1[data3 + 2]
    #     sno_gcess=b_details1[7]
    sno_add = b_details1[8]
    sno_cvd = b_details1[9]
    sno_sg = b_details1[10]
    sno_tvalue = b_details1[11]

    lis1 = []
    lis1.insert(0, sno_bcd)
    lis1.insert(1, sno_acd)
    lis1.insert(2, sno_sws)
    lis1.insert(3, sno_sad)
    lis1.insert(4, sno_igst)
    lis1.insert(5, sno_gcess)
    lis1.insert(6, sno_add)
    lis1.insert(7, sno_cvd)
    lis1.insert(8, sno_sg)
    lis1.insert(9, sno_tvalue)
    #     print(lis1)

    b_details2 = df2.iloc[19].tolist()
    b_details3 = df2.iloc[20].tolist()

    rate_bcd = b_details2[2]
    rate_acd = b_details2[3]
    rate_sws = b_details2[4]
    rate_sad = b_details2[5]
    rate_igst = data1[data3 + 3]
    rate_gcess = data1[data2 + 3]
    #     rate_igst=b_details2[6]
    #     rate_gcess=b_details3[7].split('\n')[0].split(' ')[1]
    rate_add = b_details2[-5]
    if b_details2[-4] == '':
        #         print('if')
        rate_cvd = b_details3[-4][0]
    else:
        #         print('else')
        rate_cvd = b_details2[-4]
    #     rate_add=b_details2[8]
    #     rate_cvd=b_details3[10].split('\n')[0].split(' ')[1]
    rate_sg = b_details2[10]
    rate_tvalue = b_details2[11]

    lis2 = []
    lis2.insert(0, rate_bcd)
    lis2.insert(1, rate_acd)
    lis2.insert(2, rate_sws)
    lis2.insert(3, rate_sad)
    lis2.insert(4, rate_igst)
    lis2.insert(5, rate_gcess)
    lis2.insert(6, rate_add)
    lis2.insert(7, rate_cvd)
    lis2.insert(8, rate_sg)
    lis2.insert(9, rate_tvalue)
    #     print(lis2)

    amount_bcd = b_details3[2]
    amount_acd = b_details3[3]
    amount_sws = b_details3[4]
    amount_sad = b_details3[5]
    amount_igst = data1[data3 + 4]
    amount_gcess = data1[data2 + 4]
    #     amount_igst=b_details3[6]
    #     amount_gcess=b_details3[7].split('\n')[0].split(' ')[1]
    #     amount_add=b_details3[9]
    #     amount_cvd=b_details3[10].split('\n')[0].split(' ')[1]
    amount_cvd = b_details3[-4][-1]
    amount_add = b_details3[-5]
    amount_sg = b_details3[11]
    amount_tvalue = b_details3[12]

    lis3 = []
    lis3.insert(0, amount_bcd)
    lis3.insert(1, amount_acd)
    lis3.insert(2, amount_sws)
    lis3.insert(3, amount_sad)
    lis3.insert(4, amount_igst)
    lis3.insert(5, amount_gcess)
    lis3.insert(6, amount_add)
    lis3.insert(7, amount_cvd)
    lis3.insert(8, amount_sg)
    lis3.insert(9, amount_tvalue)
    #     print(lis3)

    b_details4 = df2.iloc[21].tolist()
    dutyfg_bcd = b_details4[2]
    dutyfg_acd = b_details4[3]
    dutyfg_sws = b_details4[4]
    dutyfg_sad = b_details4[5]
    dutyfg_igst = data1[data3 + 5]
    dutyfg_gcess = data1[data2 + 5]
    #     dutyfg_igst=b_details4[6]
    #     # print(dutyfg_igst)
    #     dutyfg_gcess=b_details3[7].split(' ')[1]
    #     print(dutyfg_gcess)
    dutyfg_add = b_details4[8]
    dutyfg_cvd = b_details4[9]
    dutyfg_sg = b_details4[10]
    dutyfg_tvalue = b_details4[11]

    lis4 = []
    lis4.insert(0, dutyfg_bcd)
    lis4.insert(1, dutyfg_acd)
    lis4.insert(2, dutyfg_sws)
    lis4.insert(3, dutyfg_sad)
    lis4.insert(4, dutyfg_igst)
    lis4.insert(5, dutyfg_gcess)
    lis4.insert(6, dutyfg_add)
    lis4.insert(7, dutyfg_cvd)
    lis4.insert(8, dutyfg_sg)
    lis4.insert(9, dutyfg_tvalue)
    #     print(lis4)
    list_b = []
    list_b.extend(lis)
    list_b.extend(lis1)
    list_b.extend(lis2)
    list_b.extend(lis3)
    list_b.extend(lis4)
    #     print(list_b)

    # ******************************************************************

    # getting c data
    data4 = data1.index('5.CAIDC')
    data4
    c_details = df2.iloc[23].tolist()
    c_details
    spexd = c_details[2]
    chcess = c_details[3]
    ttd = c_details[4]
    cess = c_details[5]
    caidc = data1[data4 + 1]
    #     caidc=c_details[6].split('\n')[0]
    eaidc = c_details[7]
    cusedc = c_details[8]
    cushec = c_details[9]
    ncd = c_details[10]
    aggr = c_details[11]
    lis = []
    lis.insert(0, spexd)
    lis.insert(1, chcess)
    lis.insert(2, ttd)
    lis.insert(3, cess)
    lis.insert(4, caidc)
    lis.insert(5, eaidc)
    lis.insert(6, cusedc)
    lis.insert(7, cushec)
    lis.insert(8, ncd)
    lis.insert(9, aggr)
    #     print(lis)

    c_details1 = df2.iloc[24].tolist()
    # print(c_details1)
    sno_spexd = c_details1[2]
    sno_chcess = c_details1[3]
    sno_ttd = c_details1[4]
    sno_cess = c_details1[5]
    caidc_index = list2.index('5.CAIDC')
    sno_caidc = data1[data4 + 2]
    #     sno_caidc=list2[caidc_index+2]
    sno_eaidc = c_details1[7]
    sno_cusedc = c_details1[8]
    sno_cushec = c_details1[9]
    sno_ncd = c_details1[10]
    sno_aggr = c_details1[11]
    lis1 = []
    lis1.insert(0, sno_spexd)
    lis1.insert(1, sno_chcess)
    lis1.insert(2, sno_ttd)
    lis1.insert(3, sno_cess)
    lis1.insert(4, sno_caidc)
    lis1.insert(5, sno_eaidc)
    lis1.insert(6, sno_cusedc)
    lis1.insert(7, sno_cushec)
    lis1.insert(8, sno_ncd)
    lis1.insert(9, sno_aggr)
    #     print(lis1)

    c_details2 = df2.iloc[25].tolist()
    c_details3 = df2.iloc[26].tolist()
    # print(c_details1)
    rate_spexd = c_details2[2]
    rate_chcess = c_details2[3]
    rate_ttd = c_details2[4]
    rate_cess = c_details2[5]
    # caidc_index = list1.index('5.CAIDC')
    # rate_caidc=list1[caidc_index+3]
    #     rate_caidc=c_details3[6].split(' ')[0]
    rate_caidc = data1[data4 + 3]
    rate_eaidc = c_details2[7]

    if c_details2[-5] == '':
        rate_cusedc = c_details3[-5][0]
    else:
        rate_cusedc = c_details2[-5]

    if c_details2[-4] == '':
        rate_cushec = c_details3[-5][0]
    else:
        rate_cushec = c_details2[-5]

    #     rate_cusedc=c_details3[9].split(' ')[0]
    #     rate_cushec=c_details3[10].split(' ')[0]

    rate_ncd = c_details2[-3]
    rate_aggr = c_details2[-2]
    lis2 = []
    lis2.insert(0, rate_spexd)
    lis2.insert(1, rate_chcess)
    lis2.insert(2, rate_ttd)
    lis2.insert(3, rate_cess)
    lis2.insert(4, rate_caidc)
    lis2.insert(5, rate_eaidc)
    lis2.insert(6, rate_cusedc)
    lis2.insert(7, rate_cushec)
    lis2.insert(8, rate_ncd)
    lis2.insert(9, rate_aggr)
    #     print(lis2)

    # print(c_details3)
    amount_spexd = c_details3[2]
    amount_chcess = c_details3[3]
    amount_ttd = c_details3[4]
    amount_cess = c_details3[5]
    amount_caidc = data1[data4 + 4]
    #     amount_caidc=c_details3[6].split(' ')[1].split('\n')[0]
    amount_eaidc = c_details3[8]

    amount_cusedc = c_details3[-5][-1]
    amount_cushec = c_details3[-4][-1]

    #     amount_cusedc=c_details3[9].split(' ')[1]
    #     amount_cushec=c_details3[10].split(' ')[1]
    amount_ncd = c_details3[12]
    amount_aggr = c_details3[13]
    lis3 = []
    lis3.insert(0, amount_spexd)
    lis3.insert(1, amount_chcess)
    lis3.insert(2, amount_ttd)
    lis3.insert(3, amount_cess)
    lis3.insert(4, amount_caidc)
    lis3.insert(5, amount_eaidc)
    lis3.insert(6, amount_cusedc)
    lis3.insert(7, amount_cushec)
    lis3.insert(8, amount_ncd)
    lis3.insert(9, amount_aggr)
    #     print(lis3)

    c_details4 = df2.iloc[27].tolist()
    # print(c_details4)
    dutyfg_spexd = c_details4[2]
    dutyfg_chcess = c_details4[3]
    dutyfg_ttd = c_details4[4]
    dutyfg_cess = c_details4[5]
    dutyfg_caidc = data1[data4 + 5]
    #     dutyfg_caidc=c_details4[7]
    dutyfg_eaidc = c_details4[8]
    dutyfg_cusedc = c_details4[9]
    dutyfg_cushec = c_details4[10]
    dutyfg_ncd = c_details4[12]
    dutyfg_aggr = c_details4[-1]
    lis4 = []
    lis4.insert(0, dutyfg_spexd)
    lis4.insert(1, dutyfg_chcess)
    lis4.insert(2, dutyfg_ttd)
    lis4.insert(3, dutyfg_cess)
    lis4.insert(4, dutyfg_caidc)
    lis4.insert(5, dutyfg_eaidc)
    lis4.insert(6, dutyfg_cusedc)
    lis4.insert(7, dutyfg_cushec)
    lis4.insert(8, dutyfg_ncd)
    lis4.insert(9, dutyfg_aggr)
    #     print(lis4)

    list_c = []
    list_c.extend(lis)
    list_c.extend(lis1)
    list_c.extend(lis2)
    list_c.extend(lis3)
    list_c.extend(lis4)
    # list_c
    #     print(len(list_c))
    final_list1 = []
    final_list1.extend(list_a)
    final_list1.extend(list_b)
    final_list1.extend(list_c)
    #     print(final_list1)

    # **************************************************************************
    # Creating dataframe headers
    # **************************************************************************
    dataframe1 = pd.DataFrame(columns=["Port Code", "BE No", "BE Date", "BE Type", "IEC/Br", "GSTIN/TYPE", "CB CODE",
                                       "Type-INV(Nos)", "Type-ITEM(Nos)", "Type-COUNT(Nos)", "PKG", "G.WT(KGS)",
                                       "A_1.BE STATUS", "A_2.MODE",
                                       "A_3.DEF BE", "A_4.KACHA", "A_5.SEC 48", " A_6.REIMP", " A_7.ADV BE",
                                       "A_8.ASSESS", " A_9.EXAM", "A_10.HSS ",
                                       "A_11.FIRST CHECK", "A_12. PROV/FINAL", "A_13.COUNTRY OF ORIGIN",
                                       "A_14.COUNTRY OF CONSIGNMENT", "A_15.PORT OF LOADING",
                                       "A_16.PORT OF Shipment", "B_1.IMPORTER NAME & ADDRESS", "B_2.CB NAME",
                                       "B_AD CODE",
                                       "C_1.BCD", "C_2.ACD", "C_3.SWS", "C_4.NCCD", "C_5.ADD", "C_6.CVD", "C_7.IGST",
                                       "C_8.G.CESS", "C_9.SG",
                                       "C_10.SAED", "C_11.GSIA", "C_12.TTA", "C_13.HEALTH", "C_14.TOTAL DUTY",
                                       "C_15.INT", "C_16.PNLTY", "C_17.FINE", "C_18.TOT.ASS VAL", "C_19.TOT.AMOUNT",
                                       "D_1.IGM NO", "D_2.IGM DATE", "D_3.INW DATE", "D_4.GIGMNO", "D_6.MAWB NO",
                                       "D_7.DATE", "D_8.HAWB NO",
                                       "D_10.PKG", "D_11.GW", "F_1.SR NO", "F_2.CHALLAN NO", "F_3.PAID ON",
                                       "F_4.AMOUNT(RS.)", "H_1.EVENT(Submission)",
                                       "H_2.DATE(Submission)", "H_3.TIME", "EXCHANGE RATE(Submission)",
                                       "H_1.EVENT(Assessment)", "H_2.DATE(Assessment)", "H_3.TIME(Assessment)",
                                       "H_EXCHANGE RATE(Assessment)",
                                       "I_1.SR NO", "I_2.INVOICE NUMBER", "I_3.INV.AMT", "I_4.CUR", "OOC NO",
                                       "OOC DATE", "OOC TIME"
                                       ])
    dataframe1
    dataframe2 = pd.DataFrame(
        columns=["A_1.S.NO", "A_2.INVOICE.NO AND DATE", "B_1.BUYER'S NAME AND ADDRESS", "B_3.SUPPLIER NAMR AND ADDRESS",
                 "B_6.AD CODE",
                 "C_1.INV VALUE", "C_2.FREIGHT", "C_3.INSURANCE", "C_4.HSS", "C_5.LOADING", "C_6.COMMN",
                 "C_7.PAY TERMS", "C_8.VALUATION METHOD", "C_9.RELTD", "C_10.SVB CH", "C_11.SVB NO", "C_12.DATE",
                 "C_13.LOA", "C_14.CUR", "C_15.TERM", "D_13.MISC_CHARGE",
                 "D_14.ASS. VALUE", "E_1.SNO", "E_2.CTH", "E_3.DESCRIPTION", "E_4.UNIT PRICE", "E_5.QUANTITY",
                 "E_6.UQC", "E_7.AMOUNT", "E_1.SNO", "E_2.CTH", "E_3.DESCRIPTION", "E_4.UNIT PRICE", "E_5.QUANTITY",
                 "E_6.UQC", "E_7.AMOUNT"
                 ])
    dataframe2
    dataframe3 = pd.DataFrame(
        columns=["A_1.INVSO", "A_2.ITEMSN", "A_3.CTH", "4.CETH", "A_5.ITEM DESCRIPTION", "A_6.FS", "A_7.PQ", "A_8.DC",
                 "9.WC",
                 "A_10.AQ", "A_11.UPI", "A_12.COO",
                 "A_13.C.QTY", "A_14.C.UQC", "A_15.S.QTY", "A_16.S.UQC", "A_17.SCH", "A_18.STND/PR", "A_19.RSP",
                 "A_20.REIMP", "A_21.PROV", "A_22.END USE", "A_23.PRODN", "A_24.CNTRL", "A_25.QUALFR", "A_26.CONTNT",
                 "A_27.STMNT", "A_28.SUP DOCS", "A_29.ASSESS VALUE", "A_30.TOTAL DUTY",
                 "B_NOTN NO. 1.BCD", "NOTN NO 2.ACD", "NOTN NO 3.SWS", "NOTN NO 4.SAD", "NOTN NO 5.IGST",
                 "NOTN NO 6.G.CESS",
                 "NOTN NO 7.ADD", "NOTN NO 8.CVD", "NOTN NO 9.SG", "NOTN NO 10.T.VALUE",
                 "NOTN SNO 1.BCD", "NOTN SNO 2.ACD", "NOTN SNO 3.SWS", "NOTN SNO 4.SAD", "NOTN SNO 5.IGST",
                 "NOTN SNO 6.G.CESS",
                 "NOTN SNO 7.ADD", "NOTN SNO 8.CVD", "NOTN SNO 9.SG", "NOTN SNO 10.T.VALUE",
                 "RATE 1.BCD", "RATE 2.ACD", "RATE 3.SWS", "RATE 4.SAD", "RATE 5.IGST", "RATE 6.G.CESS",
                 "RATE 7.ADD", "RATE 8.CVD", "RATE 9.SG", "RATE 10.T.VALUE",
                 "AMOUNT 1.BCD", "AMOUNT 2.ACD", "AMOUNT 3.SWS", "AMOUNT 4.SAD", "AMOUNT 5.IGST", "AMOUNT 6.G.CESS",
                 "AMOUNT 7.ADD",
                 "AMOUNT 8.CVD", "AMOUNT 9.SG", "AMOUNT 10.T.VALUE",
                 "DUTY FG 1.BCD", "DUTY FG 2.ACD", "DUTY FG 3.SWS", "DUTY FG 4.SAD", "DUTY FG 5.IGST",
                 "DUTY FG 6.G.CESS", "DUTY FG 7.ADD",
                 "DUTY FG 8.CVD", "DUTY FG 9.SG", "DUTY FG 10.T.VALUE",
                 "C_NOTN NO 1.SP EXD", "NOTN NO 2.CHCESS", "NOTN NO 3.TTA", "NOTN NO 4.CESS", "NOTN NO 5.CAIDC",
                 "NOTN NO 6.EAIDC",
                 "NOTN NO 7.CUS EDC", "NOTN NO 8.CUS HEC", "NOTN NO 9.NCD", "NOTN NO 10.AGGR",
                 "NOTN SNO 1.SP EXD", "NOTN SNO 2.CHCESS", "NOTN SNO 3.TTA", "NOTN SNO 4.CESS", "NOTN SNO 5.CAIDC",
                 "NOTN SNO 6.EAIDC",
                 "NOTN SNO 7.CUS EDC", "NOTN SNO 8.CUS HEC", "NOTN SNO 9.NCD", "NOTN SNO 10.AGGR",
                 "RATE 1.SP EXD", "RATE 2.CHCESS", "RATE 3.TTA", "RATE 4.CESS", "RATE 5.CAIDC", "RATE 6.EAIDC",
                 "RATE 7.CUS EDC", "RATE 8.CUS HEC", "RATE 9.NCD", "RATE 10.AGGR",
                 "AMOUNT 1.SP EXD", "AMOUNT 2.CHCESS", "AMOUNT 3.TTA", "AMOUNT 4.CESS", "AMOUNT 5.CAIDC",
                 "AMOUNT 6.EAIDC",
                 "AMOUNT 7.CUS EDC", "AMOUNT 8.CUS HEC", "AMOUNT 9.NCD", "AMOUNT 10.AGGR",
                 "DUTY FG 1.SP EXD", "DUTY FG 2.CHCESS", "DUTY FG 3.TTA", "DUTY FG 4.CESS", "DUTY FG 5.CAIDC",
                 "DUTY FG 6.EAIDC",
                 "DUTY FG 7.CUS EDC", "DUTY FG 8.CUS HEC", "DUTY FG 9.NCD", "DUTY FG 10.AGGR"
                 ])
    dataframe3
    dataframe4 = pd.DataFrame(columns=["1.S NO", "2.INVOICE NO", "3.INVOICE AMOUNT", "4.CUR"
                                       ])
    result_data1 = list(Data1)
    result_data2 = list(final_list)
    result_data3 = list(final_list1)
    result_data4 = list(lis5)
    dataframe1.loc[len(dataframe1)] = result_data1
    dataframe2.loc[len(dataframe2)] = result_data2
    dataframe3.loc[len(dataframe3)] = result_data3
    dataframe4.loc[len(dataframe4)] = result_data4
    final_df1 = final_df1.append(dataframe1)
    final_df2 = final_df2.append(dataframe2)
    final_df3 = final_df3.append(dataframe3)
    final_df4 = final_df4.append(dataframe4)


with pd.ExcelWriter('C:/Users/hp/Desktop/BOE1/boe.xlsx') as writer:
    final_df1.to_excel(writer, sheet_name='BOE_SUMMARY',index=False)
    final_df2.to_excel(writer, sheet_name='Valuation_Details',index=False)
    final_df3.to_excel(writer, sheet_name='Duties',index=False)
    final_df4.to_excel(writer, sheet_name='Invoice_Details',index=False)
