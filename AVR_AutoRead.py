import os,subprocess,sys
#os.system('pythonw MainWindowLight.pyw')
#os.system("taskkill /f /im cmd.exe")

#theproc = subprocess.Popen(['MainWindowLight.pyw'], shell=True, creationflags=subprocess.SW_HIDE)
#theproc.communicate()
#theproc = subprocess.Popen([sys.executable, "MainWindowLight.pyw"], stdout=subprocess.PIPE, stderr=subprocess.PIPE, stdin=subprocess.PIPE)
#theproc.communicate()

#theproc = subprocess.Popen("MainWindowLight.pyw", shell = True)
#theproc.communicate()  


with open(os.devnull, 'wb') as devnull:
    subprocess.check_call(['pythonw', 'ReadoutAutomation_V4.py'], stdout=devnull, stderr=subprocess.STDOUT)
