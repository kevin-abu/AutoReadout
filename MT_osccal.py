import openpyxl

def OSC_CAL(OC_MT1, OC_MT2, OC_MT3, OC_MT4,workbook_save):
    wb = openpyxl.load_workbook(workbook_save)
    sheet = wb['Sheet1']
    sheet['D17'] = OC_MT1#Osc cal 1
    sheet['D18'] = OC_MT2#OC2
    sheet['D19'] = OC_MT3#OC3 
    sheet['D20'] = OC_MT4#OC4
    #print (OC_MT1, OC_MT2, OC_MT3, OC_MT4)
    wb.save(workbook_save)

def check_osc_cal(workbook_save,new_flabel):
    #print(workbook_save)

    OC_hex= open(new_flabel  + "_osc_cal.hex",'r').read()
    #print (OC_hex)
    if OC_hex[2] == '2':
        OC_MT1 = ("0x" + OC_hex[9:11])
        OC_MT2 = ("0x" + OC_hex[11:13])
        OC_MT3 = "N/A"
        OC_MT4 = "N/A"
        #print (OC_hex[2])
        
    elif OC_hex[2] == '3':
        OC_MT1 = ("0x" + OC_hex[9:11])
        OC_MT2 = ("0x" + OC_hex[11:13])
        OC_MT3 = ("0x" + OC_hex[13:15])
        OC_MT4 = "N/A"
        
    elif OC_hex[2] == '4':
        OC_MT1 = ("0x" + OC_hex[9:11])
        OC_MT2 = ("0x" + OC_hex[11:13])
        OC_MT3 = ("0x" + OC_hex[13:15])
        OC_MT4 = ("0x" + OC_hex[15:17])
    else:
        OC_MT1 = ("0x" + OC_hex[9:11])
        OC_MT2 = "N/A"
        OC_MT3 = "N/A"
        OC_MT4 = "N/A"
    OSC_CAL(OC_MT1, OC_MT2, OC_MT3, OC_MT4,workbook_save)
#check_osc_cal()
