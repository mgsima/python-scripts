import os
from pathlib import Path
from win32 import win32api
from win32 import win32print
import win32com.client
import time
from io import BytesIO
import re



# Muestra la configuración de todas las impresoras disponibles
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


def ordenar_lista(filepnames):
    """
    Sorts a list of PDF and TXT files.
    Files with numeric prefixes are sorted numerically, while others are sorted alphabetically.
    """
    ordered_list = []       # List for files without numeric prefix
    numbered_list = []      # List for files with numeric prefix

    for file in filepnames:
        path = file
        root, extension = os.path.splitext(path)

        # Check for PDF and TXT files
        if extension.lower() in ['.pdf', '.txt']:
            number = re.findall(r'\d+', root)
            # If a number is found, add to the numbered list
            if number:
                numbered_list.append((int(number[0]), file))
            else:
                # Else, add to the ordered list
                ordered_list.append(file)

    # Sort files with numbers by their numeric value
    numbered_list.sort(key=lambda x: x[0])
    # Extract only file names from the numbered list
    ordered_files_with_numbers = [item[1] for item in numbered_list]
    # Add numerically sorted files to the ordered list
    ordered_list.extend(ordered_files_with_numbers)
    # Finally, sort the combined list alphabetically
    ordered_list.sort()

    return ordered_list



# Encuentra todos los archivos PDF en la carpeta proporcionada y devuelve una lista de rutas de archivo
def find_documents(folder):  
    to_print = []
    for dirpath, dirnames, filenames in os.walk(folder):
        if len(filenames)!=0:
            for dirppath, dirpnames, filepnames in os.walk(folder):
                for file in filepnames:
                    x = dirpath + '\\' + file
                    to_print.append(x)
            return ordenar_lista(to_print)


# Imprime archivos PDF utilizando una impresora predeterminada
def print_pdf():
    folder = Path(r"Y:\D10347_BIONORICA SE\10033111_VERDAMPFERANLAGE BIONORICA\10_Dokumentation\14_Dokumentation der Zulieferanten\06_Stickstoffleitung")    # Seleccionar la carpeta objetivo
    list_to_print = find_documents(folder)
    # print(list_to_print)
    
    current_printer = win32print.GetDefaultPrinter()
    print(f"Inicio: {current_printer}")


    # Cambiar la impresora predeterminada

    PRINTER_DEFAULTS = {"DesiredAccess":win32print.PRINTER_ALL_ACCESS}  
    pHandle_n = (r'Microsoft Print to PDF')    
    win32print.SetDefaultPrinter(pHandle_n)
    
     
    # Configurar la impresora
    pHandle = win32print.OpenPrinter(pHandle_n, PRINTER_DEFAULTS)
    
    # Cambiar la configuración de la impresora (orientación, color, fuente, etc.)
    
    properties = win32print.GetPrinter(pHandle, 2)  
    pDevModeObj = properties["pDevMode"]
    pDevModeObj.Orientation = 1  
    pDevModeObj.Color = 1
    pDevModeObj.DefaultSource = 2
    properties["pDevMode"] = pDevModeObj
    win32print.SetPrinter(pHandle,2,properties,0)
    print("Control1: ",win32print.GetDefaultPrinter())
    
    # time.sleep(5)
    # Imprimir los documentos en la lista
    # for element in list_to_print:           
        # win32api.ShellExecute (0, "print", element, f'/d:"{pHandle}"', ".", 0)
        # os.startfile(element, "print")
        # time.sleep(10)
        
    win32print.ClosePrinter(pHandle)

    # Mostrar información de la configuración de la impresora    
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
    printer_name = (r'\\fs01\EPSON Drucker 1')    
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
                win32print.WritePrinter(hPrinter, b"test")
                win32print.EndPagePrinter(hPrinter)
            finally:
                win32print.EndDocPrinter(hPrinter)
        finally:
            win32print.ClosePrinter(hPrinter)   


# folder = Path(r"W:\DOKUMENTATIONSBÜRO\Auto_dokumentation\Development\test")
# lista = find_documents(folder)




print_pdf()
# print_string()
# show_printer_configuration()

