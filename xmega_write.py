def write_to_excel(Dev_Sig_hex,Lock_bits_hex,workbook_save,new_flabel, FA_Number, SN_number, readout_tool, dev_name, readout_interface, target_voltage, clock_frequency):
    Dev_Sig_MT = ("0x" + Dev_Sig_hex[9:15])
    Lock_bits_MT = ("0x" + Lock_bits_hex[9:11])
    
    import openpyxl
    x1 = openpyxl.load_workbook("XMEGA_READOUT_TEMPLATE.xlsx")
    print ("Generating Readout Summary.")
    sheet = x1['Sheet1']
    sheet['A1']
    sheet['D3'] = FA_Number #FA number
    sheet['D4'] = SN_number.upper() #SN
    sheet['D6'] = readout_tool.upper() #Tool
    sheet['D7'] = dev_name.upper()#Device
    sheet['D8'] = readout_interface.upper() # Interface 
    sheet['D9'] = target_voltage# Target Voltage
    sheet['D10']= clock_frequency # Clock
    print ("Writing Device Name and Readout Details.")
    
    sheet['D12'] = Dev_Sig_MT# Dev Sig
    sheet['D19'] = Lock_bits_MT# Lock Bits
    #x1.save(workbook_save)
    x1.save("C:\\Users\\a50291\\Documents\\READOUT_AUTOMATION\\XMEGA_READOUT_TEMPLATE_OUT.xlsx")#
    import xmega_fuse_prodsig
    xmega_fuse_prodsig.xmega_version(dev_name,new_flabel,workbook_save)

def xmega_read_extracted_info(new_flabel,FA_Number, SN_number, readout_tool, dev_name, readout_interface, target_voltage, clock_frequency):
    print ("Reading extracted info..")
    Dev_Sig_hex = open(new_flabel  + "_device_signature.hex",'r').read()
    Lock_bits_hex= open(new_flabel  + "_lockbits.hex" ,'r').read()
    workbook_save = new_flabel + "_Readout-Summary.xlsx"
    write_to_excel(Dev_Sig_hex,Lock_bits_hex,workbook_save,new_flabel, FA_Number, SN_number, readout_tool, dev_name, readout_interface, target_voltage, clock_frequency)



def xmega_initialize_path(save_path, file_label, target_voltage, FA_Number, SN_number, readout_tool, dev_name, readout_interface, clock_frequency):
    print ("Reading Paths.")
    new_flabel = save_path + "\\\\" + file_label
    xmega_read_extracted_info(new_flabel,FA_Number, SN_number, readout_tool, dev_name, readout_interface, target_voltage, clock_frequency)




    

