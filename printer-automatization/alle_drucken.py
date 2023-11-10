import os
from pathlib import Path
from win32 import win32api
from win32 import win32print
from win32 import win32gui
import win32com.client
import win32con
import time
from io import BytesIO



def show_printer_configuration():
    strComputer = "."
    objWMIService = win32com.client.Dispatch("WbemScripting.SWbemLocator")
    objSWbemServices = objWMIService.ConnectServer(strComputer,"root\cimv2")
    colItems = objSWbemServices.ExecQuery("Select * from Win32_PrinterConfiguration")
    for objItem in colItems:
        print ("Bits Per Pel: ", objItem.BitsPerPel)
        print ("Caption: ", objItem.Caption)
        print ("Collate: ", objItem.Collate)
        print ("Color: ", objItem.Color)
        print ("Copies: ", objItem.Copies)
        print ("Description: ", objItem.Description)
        print ("Device Name: ", objItem.DeviceName)
        print ("Display Flags: ", objItem.DisplayFlags)
        print ("Display Frequency: ", objItem.DisplayFrequency)
        print ("Dither Type: ", objItem.DitherType)
        print ("Driver Version: ", objItem.DriverVersion)
        print ("Duplex: ", objItem.Duplex)
        print ("Form Name: ", objItem.FormName)
        print ("Horizontal Resolution: ", objItem.HorizontalResolution)
        print ("ICM Intent: ", objItem.ICMIntent)
        print ("ICM Method: ", objItem.ICMMethod)
        print ("Log Pixels: ", objItem.LogPixels)
        print ("Media Type: ", objItem.MediaType)
        print ("Name: ", objItem.Name)
        print ("Orientation: ", objItem.Orientation)
        print ("Paper Length: ", objItem.PaperLength)
        print ("Paper Size: ", objItem.PaperSize)
        print ("Paper Width: ", objItem.PaperWidth)
        print ("Pels Height: ", objItem.PelsHeight)
        print ("Pels Width: ", objItem.PelsWidth)
        print ("Print Quality: ", objItem.PrintQuality)
        print ("Scale: ", objItem.Scale)
        print ("Setting ID: ", objItem.SettingID)
        print ("Specification Version: ", objItem.SpecificationVersion)
        print ("TT Option: ", objItem.TTOption)
        print ("Vertical Resolution: ", objItem.VerticalResolution)
        print ("X Resolution: ", objItem.XResolution)
        print ("Y Resolution: ", objItem.YResolution)
        print ("---------------------------------------------------")
    

# ordenar los pdf
def ordenar_lista(filepnames):
    # create an empty list to store sorted PDFs
    lista_ordenada = []
    # sort filenames in filepnames from low to high
    sorted_dir = sorted(filepnames, reverse=False)
    for file in sorted_dir:
        # if rev, Revision or  Rev is present in the path, set path to False
        if (r"\rev") in file or (r"\Revision") in file or (r"\Rev") in file:
            path = False
        # else, set path to the file's path
        else:
            path = file
            # split path into root directory and extension
            root, extension = os.path.splitext(path)
            # if the file is a PDF, add it to lista_ordenada
            if extension == '.pdf':
                lista_ordenada.append(file)
    # return sorted list
    return lista_ordenada

def find_documents(folder):     # Buscar todos los pdf de la carpeta # guardar las rutas de los pdfs
    to_print = []
    for dirpath, dirnames, filenames in os.walk(folder):
        if len(filenames)!=0:
                for file in filenames:
                    x1 = os.path.join(dirpath, file)
                    to_print.append(x1)
        # if len(dirnames)!=0:
        #     for dir in dirnames:
        #         folder_in = os.path.join(dirpath, dir)
        #         for dirppath, dirpnames, filepnames in os.walk(folder_in):
        #                 for file in filepnames:
        #                     x2 = os.path.join(dirppath, file)
        #                     to_print.append(x2)
    return ordenar_lista(to_print)


def print_pdf(folder_path):
    folder = Path(folder_path)    # Seleccionar la carpeta objetivo
    # folder = Path(r"W:\DOKUMENTATIONSBÜRO\Auto_dokumentation\Development\test")
    list_to_print = find_documents(folder)
    # print(list_to_print)
    
    current_printer = win32print.GetDefaultPrinter()
    print("Inicio: ",win32print.GetDefaultPrinter())

    PRINTER_DEFAULTS = {"DesiredAccess":win32print.PRINTER_ALL_ACCESS}  
    pHandle_n = (r'\\fs01\EPSON Drucker 2')    
    win32print.SetDefaultPrinter(pHandle_n)
    
    pHandle = win32print.OpenPrinter(pHandle_n, PRINTER_DEFAULTS)
    
    properties = win32print.GetPrinter(pHandle, 2)  
    pDevModeObj = properties["pDevMode"]
    pDevModeObj.Orientation = 1  
    pDevModeObj.Color = 2       # 1=NO_Color 2=Color
    pDevModeObj.DefaultSource = 2  
    properties["pDevMode"] = pDevModeObj
    win32print.SetPrinter(pHandle,2,properties,0)
    print("Control1: ",win32print.GetDefaultPrinter())
    
    x = win32print.DeviceCapabilities(pHandle_n, '192.168.0.29', win32con.DC_BINS)
    print(x)

    
    # time.sleep(5)
    for element in list_to_print:           
        print(element)
        win32api.ShellExecute (0, "print", element, f'/d:"{pHandle}"', ".", 0)
        time.sleep(10)
        
    win32print.ClosePrinter(pHandle)
    
    statment=False
    if statment:

        print ("formName: ", pDevModeObj.FormName)
        print ("PaperSize: ", pDevModeObj.PaperSize)
        print ("DefaultSource: ", pDevModeObj.DefaultSource)
        print ("PaperLength: ", pDevModeObj.PaperLength)
        print ("PaperWidth: ", pDevModeObj.PaperWidth)
        print ("Color: ", pDevModeObj.Color)
        
        print(pHandle)
        for i in properties:
            print(i, properties[i])
            

def print_string():
    
    ### looking for documents
    folder = Path(r"W:\DOKUMENTATIONSBÜRO\Auto_dokumentation\Development\test")    # Seleccionar la carpeta objetivo
    list_to_print = find_documents(folder)
    
    ### opening the printer
    printer_name = (r'\\fs01\EPSON Drucker 2')    
    PRINTER_DEFAULTS = {"DesiredAccess":win32print.PRINTER_ALL_ACCESS}  
    hPrinter = win32print.OpenPrinter(printer_name, PRINTER_DEFAULTS)
    printer_info = win32print.GetPrinter(hPrinter, 2)
    
    
    ### checking the driver
    drivers = win32print.EnumPrinterDrivers(None, None, 2)
    for driver in drivers:
        if driver["Name"] == printer_info["pDriverName"]:
            printer_driver = driver
    raw_type = "XPS_PASS" if printer_driver["Version"] == 4 else "RAW"
    
    
    ### setting the properties
    properties = win32print.GetPrinter(hPrinter, 2)  
    pDevModeObjIN = properties["pDevMode"]
    pDevModeObj = properties["pDevMode"]
    pDevModeObj.Orientation = 1  
    pDevModeObj.Color = 1
    pDevModeObj.DefaultSource = 2
    properties["pDevMode"] = pDevModeObj
    win32print.SetPrinter(hPrinter,2,properties,0)
    win32print.DocumentProperties(0, hPrinter, printer_name, pDevModeObj, pDevModeObjIN, 0)
    

    ### printing
    for element in list_to_print:
        try:
            hJob = win32print.StartDocPrinter(hPrinter, 1, (element, None, raw_type))
            try:
                win32print.StartPagePrinter(hPrinter)
                # win32print.WritePrinter(hPrinter, b"test")
                
                win32print.EndPagePrinter(hPrinter)
            finally:
                win32print.EndDocPrinter(hPrinter)
        finally:
            win32print.ClosePrinter(hPrinter)   



folder = Path(r"example\path")
lista = find_documents(folder)
print(lista)
# print_pdf(folder_path)
# print_string()
show_printer_configuration()

