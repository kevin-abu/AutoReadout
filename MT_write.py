


    
    
def write_to_excel(Dev_Sig_hex,Fuses_MT,Lock_bits_hex,OC_hex,workbook_save,new_flabel, FA_Number, SN_number, readout_tool, dev_name, readout_interface, target_voltage, clock_frequency):
    Dev_Sig_MT = ("0x" + Dev_Sig_hex[9:15])
    Fuses_MT_Low = ("0x" + Fuses_MT[9:11])
    Fuses_MT_High = ("0x" + Fuses_MT[11:13])
    Fuses_MT_Extended = ("0x" + Fuses_MT[13:15])
    Lock_bits_MT = ("0x" + Lock_bits_hex[9:11])
    #print (Dev_Sig_MT,Fuses_MT_Low,Fuses_MT_High,Fuses_MT_Extended,Lock_bits_MT)

    import openpyxl
    wb = openpyxl.load_workbook('MEGA_TINY_READOUT_TEMPLATE.xlsx')
    print ("Generating Readout Summary")
    sheet = wb['Sheet1']
    sheet['A1']
    sheet['D3'] = FA_Number #FA number
    sheet['D4'] = SN_number.upper() #SN
    sheet['D6'] = readout_tool.upper() #Tool
    sheet['D7'] = dev_name.upper()#Device
    sheet['D8'] = readout_interface.upper() # Interface 
    sheet['D9'] = target_voltage# Target Voltage
    sheet['D10']= clock_frequency # Clock
    sheet['D12'] = Dev_Sig_MT# Dev Sig
    sheet['D13'] = Fuses_MT_Extended# Ext Fuse
    sheet['D14'] = Fuses_MT_High# High Fuse
    sheet['D15'] = Fuses_MT_Low# LOW Fuse 
    sheet['D16'] = Lock_bits_MT# Lock Bits
    wb.save(workbook_save)
    import MT_osccal
    MT_osccal.check_osc_cal(workbook_save, new_flabel)
    

def initialize_path(save_path, file_label, target_voltage, FA_Number, SN_number, readout_tool, dev_name, readout_interface, clock_frequency):
    #target_voltage input is min_tv, opt_tv, or max_tv
    #save_path = min_save / nom_save / max_save
    # file_label = file_label_min
    print ("Reading save path.")
    new_flabel = save_path + "\\\\" + file_label
    read_extracted_info(new_flabel,FA_Number, SN_number, readout_tool, dev_name, readout_interface, target_voltage, clock_frequency)
    
def read_extracted_info(new_flabel, FA_Number, SN_number, readout_tool, dev_name, readout_interface, target_voltage, clock_frequency):
    print ("Reading extracted info.")
    Dev_Sig_hex = open(new_flabel  + "_device_signature.hex",'r').read()
    Fuses_MT = open(new_flabel  + "_fuses.hex",'r').read()
    Lock_bits_hex= open(new_flabel  + "_lockbits.hex" ,'r').read()
    OC_hex= open(new_flabel  + "_osc_cal.hex",'r').read()
    workbook_save = new_flabel + "_Readout-Summary.xlsx"
    write_to_excel(Dev_Sig_hex,Fuses_MT,Lock_bits_hex,OC_hex,workbook_save,new_flabel, FA_Number, SN_number, readout_tool, dev_name, readout_interface, target_voltage, clock_frequency)


#initialize_path()


