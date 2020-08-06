# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'ReadoutAutomation_V3.ui'
#
# Created by: PyQt5 UI code generator 5.12.1
#
# WARNING! All changes made in this file will be lost!

#UPDATES TO BE MADE -> SCRIPT CHECKING IF THE FILE EXISTS AFTER DEVICE INFO READOUT THEN WAIT IF NOT EXISTING, CONTINUE IF FILE EXISTS.
#https://dbader.org/blog/python-check-if-file-exists

# PyQT5 Modules #
from PyQt5.QtCore import pyqtSignal, pyqtSlot
from PyQt5.QtWidgets import *
from PyQt5 import *
from PyQt5.QtWidgets import QWidget
from PyQt5.QtWidgets import QApplication, QWidget, QInputDialog, QLineEdit, QFileDialog, QMessageBox
from PyQt5.QtGui import QIcon
##
import subprocess, os, pyautogui,time,re, sys, PyQt5,time
from atprogram import atprogram
from win32api import GetKeyState
from win32con import VK_CAPITAL
import openpyxl
from pathlib import Path
import csv
import xml.etree.ElementTree as ET


def showDialog(save_pat_conv):
    error_dialog = QMessageBox()
    error_dialog.setIcon(QMessageBox.Information)
    error_dialog.setText("Readout complete. Click OK to open output folder.")
    error_dialog.setWindowTitle("Readout Completed")
    error_dialog.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)

    returnValue = error_dialog.exec()

    if returnValue == QMessageBox.Ok:
        open_folder = save_pat_conv
        open_folder = os.path.realpath(open_folder)
        os.startfile(open_folder)

def folderExistsDialog():
    error_dialog = QMessageBox()
    error_dialog.setIcon(QMessageBox.Information)
    error_dialog.setText("Folder Already Exists.")
    error_dialog.setWindowTitle("Error. Folder Already Exists.")
    error_dialog.setStandardButtons(QMessageBox.Ok)# | QMessageBox.Cancel)

    returnValue = error_dialog.exec()
    """
    if returnValue == QMessageBox.Ok:
        open_folder = save_pat_conv
        open_folder = os.path.realpath(open_folder)
        os.startfile(open_folder)
    """


def RunMin(FA_Number,SN_number,min_tv,dev_name,save_pat_conv,folder_label,readout_tool,readout_interface,clock_frequency):
    file_label_min = FA_Number + "_" + SN_number + "_" + min_tv + "_" + dev_name
    min_save =   save_pat_conv + "\\\\" + folder_label + "\\\\" + min_tv
    os.mkdir(save_pat_conv + "\\\\" + folder_label + "\\\\" + min_tv)
    print("Performing readout at Minimum Target Voltage.")
    print("Extracting device information..")
    ReadDeviceInfo(readout_tool, readout_interface, min_tv, clock_frequency, dev_name, min_save,file_label_min,FA_Number,SN_number)
    #ReadDeviceMain(readout_tool, readout_interface, min_tv, clock_frequency, dev_name, min_save,read_file_label)


    #return file_label_min,min_save
def RunNom(FA_Number,SN_number,opt_tv,dev_name,save_pat_conv,folder_label,readout_tool,readout_interface,clock_frequency):
    nom_save = save_pat_conv + "\\\\" + folder_label + "\\\\" + opt_tv
    file_label_nom = FA_Number + "_" + SN_number + "_" + opt_tv + "_" + dev_name
    os.mkdir(save_pat_conv + "\\\\" + folder_label + "\\\\" + opt_tv)
    print ("Performing readout at Nominal Target Voltage..")
    print ("Extracting device information..")
    ReadDeviceInfo(readout_tool, readout_interface, opt_tv, clock_frequency, dev_name, nom_save,file_label_nom,FA_Number,SN_number)
    #ReadDeviceMain(readout_tool, readout_interface, opt_tv, clock_frequency, dev_name, nom_save,file_label_nom)

def RunMax(FA_Number,SN_number,max_tv,dev_name,save_pat_conv,folder_label,readout_tool,readout_interface,clock_frequency):
    max_save = save_pat_conv + "\\\\" + folder_label + "\\\\" + max_tv
    file_label_max = FA_Number + "_" + SN_number + "_" + max_tv + "_" + dev_name
    os.mkdir(save_pat_conv + "\\\\" + folder_label + "\\\\" + max_tv)
    print ("Performing readout at Maximum Target Voltage..")
    print("Extracting device information..")
    ReadDeviceInfo(readout_tool, readout_interface, max_tv, clock_frequency, dev_name, max_save,file_label_max,FA_Number,SN_number)
    #ReadDeviceMain(readout_tool, readout_interface, max_tv, clock_frequency, dev_name, max_save,file_label_max)


def ReadDeviceInfo(readout_tool,readout_interface,read_target_voltage,clock_frequency,dev_name,read_save_path,read_file_label,FA_Number,SN_number):
    #read_save_path = min_save/nom_save/max save
    #read_file_label = file_label_min/file_label_nom/file_label_max
    #read_target_voltage = min_tv/opt_tv/max_tv
    time.sleep(0.5)
    subprocess.Popen("C:\\Program Files (x86)\\Atmel\\Studio\\7.0\\Extensions\\Application\\StudioCommandPrompt.exe")
    time.sleep(3)
    pyautogui.typewrite("atprogram -t " + readout_tool + " -i " + readout_interface + " -tv " + read_target_voltage + " -cg " + clock_frequency + " -d " + dev_name + " info>" + read_save_path + "\\\\" + read_file_label + "_device_information.txt" + "\"")
    pyautogui.keyDown('enter')
    print ("Reading device information.")
    time.sleep(5)
    ReadDeviceMain(readout_tool,readout_interface,read_target_voltage,clock_frequency,dev_name,read_save_path,read_file_label,FA_Number,SN_number)

def ReadDeviceMain(readout_tool,readout_interface,read_target_voltage,clock_frequency,dev_name,read_save_path,read_file_label,FA_Number,SN_number):
    print (dev_name)
    if (dev_name[:3] == "ATM" or "ATm"):
        ###ATMEGA###
        read_main = "-t " + readout_tool + " -i " + readout_interface + " -tv " + read_target_voltage + " -cg " + clock_frequency + " -d " + dev_name + " read -sg -f " + read_save_path + "\\\\" + read_file_label + "_device_signature.hex" + " read -fs -f " + read_save_path + "\\\\" + read_file_label + "_fuses.hex" + " read -lb -f " + read_save_path + "\\\\" + read_file_label + "_lockbits.hex" + " read -os -f " + read_save_path + "\\\\" + read_file_label + "_osc_cal.hex" + " read -fl -f " + read_save_path + "\\\\" + read_file_label + "_flash.hex" +  " read -ee -f " + read_save_path + "\\\\" + read_file_label + "_eeprom.hex"
        atprogram.atprogram(device_name=dev_name, tool=readout_tool, verbose=1, interface=readout_interface, make_command=None, atprogram_command=read_main,return_output=False, dry_run=False)
        time.sleep(0.5)
        #print(read_main)
        import MT_write
        MT_write.initialize_path(read_save_path, read_file_label, read_target_voltage, FA_Number, SN_number, readout_tool, dev_name, readout_interface, clock_frequency)
    elif (dev_name[:3] == "ATX" or "ATx"):
        # read_target_voltage = min_tv/opt_tv/max_tv
        # read_save_path = min_save/nom_save/max save
        # read_file_label = file_label_min/file_label_nom/file_label_max
        #print (readout_tool,readout_interface,read_target_voltage,clock_frequency,dev_name,read_save_path,read_file_label)
        read_main = "-t " + readout_tool + " -i " + readout_interface + " -tv " + read_target_voltage + " -cg " + clock_frequency + " -d " + dev_name + " read -sg -f " + read_save_path + "\\\\" + read_file_label + "_device_signature.hex" + " read -fs -f " + read_save_path + "\\\\" + read_file_label + "_fuses.hex" + " read -lb -f " + read_save_path + "\\\\" + read_file_label + "_lockbits.hex" + " read -ps -f " + read_save_path + "\\\\" + read_file_label + "_prodsig.hex" + " read -fl -f " + read_save_path + "\\\\" + read_file_label + "_flash.hex" + " read -ee -f " + read_save_path + "\\\\" + read_file_label + "_eeprom.hex" + " read -us -f " + read_save_path + "\\\\" + read_file_label + "_UserSig.hex"
        atprogram.atprogram(device_name=dev_name, tool=readout_tool, verbose=1, interface=readout_interface, make_command=None, atprogram_command=read_main,return_output=False, dry_run=False)
        time.sleep(0.5)
        import xmega_write
        xmega_write.xmega_initialize_path(read_save_path, read_file_label, read_target_voltage, FA_Number, SN_number, readout_tool, dev_name, readout_interface, clock_frequency)
        #print(read_main)




class Ui_Dialog(object):
    def setupUi(self, Dialog):
        Dialog.setObjectName("Dialog")
        Dialog.resize(251, 357)
        font = QtGui.QFont()
        font.setFamily("Nirmala UI Semilight")
        font.setPointSize(12)
        Dialog.setFont(font)
        self.label = QtWidgets.QLabel(Dialog)
        self.label.setGeometry(QtCore.QRect(81, 6, 101, 31))
        font = QtGui.QFont()
        font.setFamily("Segoe UI Semibold")
        font.setPointSize(12)
        font.setBold(False)
        font.setWeight(50)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.fa_num = QtWidgets.QLineEdit(Dialog)
        self.fa_num.setGeometry(QtCore.QRect(11, 46, 231, 20))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.fa_num.setFont(font)
        self.fa_num.setText("")
        self.fa_num.setObjectName("fa_num")
        self.sn_num = QtWidgets.QLineEdit(Dialog)
        self.sn_num.setGeometry(QtCore.QRect(11, 71, 111, 20))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.sn_num.setFont(font)
        self.sn_num.setText("")
        self.sn_num.setObjectName("sn_num")
        self.combInt = QtWidgets.QComboBox(Dialog)
        self.combInt.setGeometry(QtCore.QRect(131, 126, 111, 20))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.combInt.setFont(font)
        self.combInt.setEditable(False)
        self.combInt.setObjectName("combInt")
        self.groupBox = QtWidgets.QGroupBox(Dialog)
        self.groupBox.setGeometry(QtCore.QRect(10, 180, 231, 131))
        self.groupBox.setObjectName("groupBox")
        self.min_check = QtWidgets.QCheckBox(self.groupBox)
        self.min_check.setGeometry(QtCore.QRect(10, 20, 121, 51))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.min_check.setFont(font)
        self.min_check.setChecked(True)
        self.min_check.setTristate(False)
        self.min_check.setObjectName("min_check")
        self.minimum_voltage = QtWidgets.QDoubleSpinBox(self.groupBox)
        self.minimum_voltage.setGeometry(QtCore.QRect(150, 30, 71, 27))
        self.minimum_voltage.setDecimals(1)
        #self.minimum_voltage.setMinimum(1.6)
        self.minimum_voltage.setMaximum(6.0)
        self.minimum_voltage.setSingleStep(0.1)
        self.minimum_voltage.setProperty("value", 1.8)
        self.minimum_voltage.setObjectName("minimum_voltage")
        self.maximum_voltage = QtWidgets.QDoubleSpinBox(self.groupBox)
        self.maximum_voltage.setGeometry(QtCore.QRect(150, 90, 71, 27))
        self.maximum_voltage.setDecimals(1)
        self.maximum_voltage.setMinimum(1.6)
        #self.maximum_voltage.setMaximum(6.0)
        self.maximum_voltage.setSingleStep(0.1)
        self.maximum_voltage.setProperty("value", 5.5)
        self.maximum_voltage.setObjectName("maximum_voltage")
        self.nominal_voltage = QtWidgets.QDoubleSpinBox(self.groupBox)
        self.nominal_voltage.setGeometry(QtCore.QRect(150, 60, 71, 27))
        self.nominal_voltage.setDecimals(1)
        self.nominal_voltage.setMinimum(1.6)
        self.nominal_voltage.setMaximum(6.0)
        self.nominal_voltage.setSingleStep(0.1)
        self.nominal_voltage.setProperty("value", 3.3)
        self.nominal_voltage.setObjectName("nominal_voltage")
        self.opt_check = QtWidgets.QCheckBox(self.groupBox)
        self.opt_check.setGeometry(QtCore.QRect(10, 50, 131, 51))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.opt_check.setFont(font)
        self.opt_check.setChecked(True)
        self.opt_check.setTristate(False)
        self.opt_check.setObjectName("opt_check")
        self.max_check = QtWidgets.QCheckBox(self.groupBox)
        self.max_check.setGeometry(QtCore.QRect(10, 80, 131, 51))
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(False)
        font.setWeight(50)
        self.max_check.setFont(font)
        self.max_check.setChecked(True)
        self.max_check.setTristate(False)
        self.max_check.setObjectName("max_check")
        self.combTool = QtWidgets.QComboBox(Dialog)
        self.combTool.setGeometry(QtCore.QRect(11, 126, 111, 20))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.combTool.setFont(font)
        self.combTool.setEditable(False)
        self.combTool.setObjectName("combTool")
        self.combTool.addItem("")
        self.combTool.addItem("")
        self.combTool.addItem("")
        self.combTool.addItem("")
        self.combTool.addItem("")
        self.combTool.addItem("")
        self.combTool.addItem("")
        self.combTool.addItem("")
        self.combTool.addItem("")
        self.combTool.addItem("")
        self.combTool.addItem("")
        self.combTool.addItem("")
        self.combTool.addItem("")
        self.combTool.addItem("")
        self.combTool.addItem("")
        self.combTool.addItem("")
        self.label_2 = QtWidgets.QLabel(Dialog)
        self.label_2.setGeometry(QtCore.QRect(10, 150, 141, 31))
        font = QtGui.QFont()
        font.setFamily("Segoe UI Semibold")
        font.setPointSize(9)
        font.setBold(False)
        font.setWeight(50)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")
        self.buttonBox = QtWidgets.QDialogButtonBox(Dialog)
        self.buttonBox.setGeometry(QtCore.QRect(50, 320, 156, 23))
        font = QtGui.QFont()
        font.setFamily("Segoe UI Semilight")
        font.setPointSize(10)
        self.buttonBox.setFont(font)
        self.buttonBox.setStandardButtons(QtWidgets.QDialogButtonBox.Cancel|QtWidgets.QDialogButtonBox.Ok)
        self.buttonBox.setObjectName("buttonBox")
        self.ChusFolder = QtWidgets.QPushButton(Dialog)
        self.ChusFolder.setGeometry(QtCore.QRect(167, 99, 75, 23))
        font = QtGui.QFont()
        font.setPointSize(8)
        self.ChusFolder.setFont(font)
        self.ChusFolder.setObjectName("ChusFolder")
        self.fa_num_2 = QtWidgets.QLineEdit(Dialog)
        self.fa_num_2.setGeometry(QtCore.QRect(10, 100, 151, 20))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.fa_num_2.setFont(font)
        self.fa_num_2.setText("")
        self.fa_num_2.setObjectName("fa_num_2")
        self.device_name = QtWidgets.QComboBox(Dialog)
        self.device_name.setGeometry(QtCore.QRect(130, 71, 111, 20))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.device_name.setFont(font)
        self.device_name.setEditable(True)
        self.device_name.setCurrentText("")
        self.device_name.setObjectName("device_name")
        self.device_name.addItem("Device Name")
        self.clock_freq_2 = QtWidgets.QSpinBox(Dialog)
        self.clock_freq_2.setGeometry(QtCore.QRect(160, 160, 81, 22))
        self.clock_freq_2.setWrapping(False)
        self.clock_freq_2.setMaximum(32000000)
        self.clock_freq_2.setSingleStep(100000)
        self.clock_freq_2.setProperty("value", 8000000)
        self.clock_freq_2.setObjectName("clock_freq_2")
        interface_list = []
        speed_max = 32000000
        vcc_min = 1.6
        vcc_max = 6.0
        atdf_index = 0
        default_val = 15
        self.retranslateUi(Dialog, atdf_index, interface_list, vcc_min, vcc_max, speed_max, default_val) # -> UPDATE
        self.combInt.setCurrentIndex(0)
        self.combTool.setCurrentIndex(0)
        self.device_name.setCurrentIndex(0)
        QtCore.QMetaObject.connectSlotsByName(Dialog)
        

        ###--------####
        device_list=[]
        self.atdf_device_name = []
        for root, dirs, files in os.walk(r"C:\\Program Files (x86)\\Atmel\\Studio\\7.0\\packs"):
            for file in files:
                if file.endswith(".atdf") and (file.startswith("AT90") or file.startswith("ATmega")or file.startswith("ATtiny")or file.startswith("ATxmega")or file.startswith("AT32")):
                    #print(os.path.join(root, file))
                    RAW = (os.path.join(root, file))
                    raw1 = RAW[:len(RAW)-5]
                    device_list.append(os.path.basename(raw1))
                    #print(os.path.basename(raw1))
                    #self.device_name.addItem(os.path.basename(raw1))
                    self.atdf_device_name.append(os.path.join(root, file))
        ###--------####

        device_list = list(dict.fromkeys(device_list))
        self.device_name.addItems(device_list)
        self.buttonBox.rejected.connect(Dialog.reject)
        self.buttonBox.accepted.connect(self.Readout)
        self.ChusFolder.clicked.connect(self.savepat)
        self.device_name.currentTextChanged.connect(self.updateComboBox2)
        #print (atdf_device_name)
        
        
            
    def updateComboBox2(self,atdf_device_name):
        
        #print (self.device_name.currentIndex())
        atdf_index = self.device_name.currentIndex()
        #print (self.atdf_device_name[atdf_index-1])
        device = ET.parse(self.atdf_device_name[atdf_index-1])
        root = device.getroot()
        for variant in root.iter('variant'):
            (variant.attrib)    
        #print ("Min. VCC = " +variant.attrib.get('vccmin'))
        vcc_min = variant.attrib.get('vccmin')
        #print ("Max. VCC = " +variant.attrib.get('vccmax'))
        vcc_max = variant.attrib.get('vccmax')
        #print ("Maximum Clock Frequency = " +variant.attrib.get('speedmax'))
        speed_max = variant.attrib.get('speedmax')
        interface_list = []
        for interface in root.iter('interface'):
            interface_list.append(interface.attrib.get('name'))
        #print (interface_list)
        default_val = 0
        
        
        self.retranslateUi(Dialog, atdf_index, interface_list, vcc_min, vcc_max, speed_max, default_val)
        
        
        #minimum_voltage.setRange(variant.attrib.get('vccmin'),variant.attrib.get('vccmax'))
        #self.minimum_voltage.setProperty("value", variant.attrib.get('vccmin'))
        #self.maximum_voltage.setMaximum(variant.attrib.get('vccmax'))
        #self.maximum_voltage.setProperty("value", variant.attrib.get('vccmax'))
        #self.clock_freq_2.setMaximum(variant.attrib.get('speedmax'))
            


    def retranslateUi(self, Dialog, atdf_index, interface_list, vcc_min, vcc_max, speed_max, default_val):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "AutoReadout"))
        self.label.setText(_translate("Dialog", "Device Set-up"))
        self.fa_num.setPlaceholderText(_translate("Dialog", "FA Request Number"))
        self.sn_num.setPlaceholderText(_translate("Dialog", "Sample Number"))
        if (self.device_name.currentText() == "Device Name"): 
            self.combInt.setCurrentText("Interface")
            self.combInt.addItem("Interface")
            self.combInt.addItem("aWire")
            self.combInt.addItem("debugWIRE")
            self.combInt.addItem("HVSP")
            self.combInt.addItem("HVPP")
            self.combInt.addItem("ISP")
            self.combInt.addItem("JTAG")
            self.combInt.addItem("PDI")
            self.combInt.addItem("UPDI")
            self.combInt.addItem("TPI")
            self.combInt.addItem("SWD")
        else:
            self.combInt.clear()
            for interface in interface_list:
                self.combInt.addItem(interface)
            #self.clock_freq_2.setMaximum(speed_max)
        self.groupBox.setTitle(_translate("Dialog", "Target Voltage"))
        self.min_check.setText(_translate("Dialog", "Minimum Voltage"))
        self.opt_check.setText(_translate("Dialog", "Nominal Voltage"))
        self.max_check.setText(_translate("Dialog", "Maximum Voltage"))
        self.combTool.setCurrentText(_translate("Dialog", "Tool"))
        self.combTool.setItemText(0, _translate("Dialog", "Tool"))
        self.combTool.setItemText(1, _translate("Dialog", "avrdragon"))
        self.combTool.setItemText(2, _translate("Dialog", "avrispmk2"))
        self.combTool.setItemText(3, _translate("Dialog", "avrone"))
        self.combTool.setItemText(4, _translate("Dialog", "jtagice3"))
        self.combTool.setItemText(5, _translate("Dialog", "jtagicemkii"))
        self.combTool.setItemText(6, _translate("Dialog", "qt600"))
        self.combTool.setItemText(7, _translate("Dialog", "stk500"))
        self.combTool.setItemText(8, _translate("Dialog", "stk600"))
        self.combTool.setItemText(9, _translate("Dialog", "samice"))
        self.combTool.setItemText(10, _translate("Dialog", "edbg"))
        self.combTool.setItemText(11, _translate("Dialog", "medbg"))
        self.combTool.setItemText(12, _translate("Dialog", "atmelice"))
        self.combTool.setItemText(13, _translate("Dialog", "powerdebugger"))
        self.combTool.setItemText(14, _translate("Dialog", "megadfu"))
        self.combTool.setItemText(15, _translate("Dialog", "flip"))
        
        self.minimum_voltage.setProperty("value", vcc_min)
        self.maximum_voltage.setProperty("value", vcc_max)
        #self.minimum_voltage.setRange(vcc_min,vcc_max)
        #self.maximum_voltage.setRange(vcc_min,vcc_max)
        
        
        self.label_2.setText(_translate("Dialog", "External Clock Frequency"))
        self.ChusFolder.setText(_translate("Dialog", "Choose Folder"))
        self.fa_num_2.setPlaceholderText(_translate("Dialog", "Save Directory"))
        #print (self.min_check.isChecked(),atdf_index)
        #
        #if (self.device_name.currentText() == ("Device Name")) and (self.device_name.activated("Device Name")):
            #self.device_name.clearEditText()
        #self.fa_num.setText("2020-001100")
        #self.sn_num.setText("sn1")
        self.combTool.setCurrentText("stk600")
        self.fa_num_2.setText("")
        #self.fa_num_2.setText(r"C:\Users\a50291\Desktop\TEST")


    def savepat(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        # fileName, _ = QFileDialog.getSaveFileName(self,"QFileDialog.getSaveFileName()","","All Files (*);;Text Files (*.txt)", options=options)
        fileName = QFileDialog.getExistingDirectory()
        if fileName:
            self.fa_num_2.setText(fileName)
        
    def Readout(self):
        print("Reading inputs.")
        FA_Number = self.fa_num.text()
        dev_name = self.device_name.currentText()
        SN_number = self.sn_num.text()
        clock_frequency = str(self.clock_freq_2.value())
        readout_interface = self.combInt.currentText()
        readout_tool = self.combTool.currentText()
        min_tv = str(round(self.minimum_voltage.value(),2))
        max_tv = str(round(self.maximum_voltage.value(),2))
        opt_tv = str(round(self.nominal_voltage.value(),2))
        #print (str(min_tv),max_tv,opt_tv)
        save_pat = self.fa_num_2.text()
        save_pat_conv = re.escape(os.path.normpath(save_pat))
        error_dialog = QtWidgets.QMessageBox()
        #print (atdf_device_name)

        ####INPUT CHECKING#####
        if (self.fa_num.text()) == (""):
            error_dialog.setText('Please Enter FA Number')
            error_dialog.exec()
        elif (self.device_name.currentText() == "Device Name"):
            error_dialog.setText('Please Enter Device Name')
            error_dialog.exec()
        elif (self.sn_num.text()) == (""):
            error_dialog.setText('Please Enter Sample Number')
            error_dialog.exec()
        elif (self.fa_num_2.text()) == (""):
            error_dialog.setText('No save directory! Please Choose Folder.')
            error_dialog.exec()
        elif (self.combTool.currentText()) == ("Tool"):
            error_dialog.setText('No readout tool selected! Please select readout tool.')
            error_dialog.exec()
        elif (self.combInt.currentText()) == ("Interface"):
            error_dialog.setText('No readout interface selected! Please select readout interface.')
            error_dialog.exec()        
        #######################
        #####START OF READING INPUTS#########
        else:

            if GetKeyState(VK_CAPITAL) == 1:
                error_dialog.setText('Please turn caps lock off before proceeding to incoming readout.')
                error_dialog.exec()
            else:

                print ("FA Number: " + FA_Number)
                print ("Device Name: " + dev_name)
                print ("Sample: " + SN_number)
                print ("Clock Frequency: " + str(clock_frequency) + " Hz")

                folder_label = FA_Number + "_" + dev_name + "_" + SN_number

                try:
                    os.mkdir(save_pat_conv + "\\\\" + folder_label)

                    if self.max_check.isChecked():  # 1 -> MAX ONLY
                        if self.min_check.isChecked():  # 2 -> MIN, MAX
                            if self.opt_check.isChecked():  # 3 -> NOM, MIN, MAX
                                RunNom(FA_Number, SN_number, opt_tv, dev_name, save_pat_conv, folder_label,
                                       readout_tool, readout_interface, clock_frequency)
                            # ---------------------------------------------------------#
                            RunMin(FA_Number, SN_number, min_tv, dev_name, save_pat_conv, folder_label, readout_tool,
                                   readout_interface, clock_frequency)

                        elif self.opt_check.isChecked():  # 4 -> NOM, MAX
                            RunNom(FA_Number, SN_number, opt_tv, dev_name, save_pat_conv, folder_label, readout_tool,
                                   readout_interface, clock_frequency)
                        ##-----MAX ONLY----####
                        RunMax(FA_Number, SN_number, max_tv, dev_name, save_pat_conv, folder_label, readout_tool,
                               readout_interface, clock_frequency)


                    elif self.min_check.isChecked():  # 5 -> MIN
                        if self.opt_check.isChecked():  # 6 -> NOM, MIN
                            RunNom(FA_Number, SN_number, opt_tv, dev_name, save_pat_conv, folder_label, readout_tool,
                                   readout_interface, clock_frequency)
                        RunMin(FA_Number, SN_number, min_tv, dev_name, save_pat_conv, folder_label, readout_tool,
                               readout_interface, clock_frequency)


                    elif self.opt_check.isChecked():  # 7 -> NOM
                        RunNom(FA_Number, SN_number, opt_tv, dev_name, save_pat_conv, folder_label, readout_tool,
                               readout_interface, clock_frequency)
                    # TO BE CONTINUED## FIX READOUT BASED ON MIN NOM MAX##
                except FileExistsError:
                    folderExistsDialog()
                os.system("taskkill /f /im cmd.exe")
                print("Readout Completed.")
                showDialog(save_pat_conv)

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    app.setStyle('Fusion')
    Dialog = QtWidgets.QDialog()
    ui = Ui_Dialog()
    ui.setupUi(Dialog)
    Dialog.show()
    sys.exit(app.exec_())
