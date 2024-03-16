#! python3
import os
import win32com.client
import fitz  # PyMuPDF
from PIL import Image
from escpos.printer import Usb
from time import sleep
from typing import Tuple, List
import winreg


"""
=====================================================
        PARA OBTENER LOS ID DE LAS IMPRESORAS USB 
=====================================================

"""
def get_usb_printer_ids():
    printer_ids = []
    wmi = win32com.client.GetObject("winmgmts:")
    for printer in wmi.InstancesOf("Win32_Printer"):
        if "USB" in printer.PortName:
            printer_ids.append(printer.PnPDeviceID)
    return printer_ids

def get_vendor_product_ids(pnp_device_id):
    vendor_id, product_id = None, None
    wmi = win32com.client.GetObject("winmgmts:")
    for item in wmi.ExecQuery("SELECT * FROM Win32_PnPEntity WHERE DeviceID='{}'".format(pnp_device_id)):
        hardware_id = item.HardwareID[0]
        if "VID_" in hardware_id and "PID_" in hardware_id:
            vendor_id = hardware_id[hardware_id.index("VID_")+4:hardware_id.index("PID_")].upper()
            product_id = hardware_id[hardware_id.index("PID_")+4:].upper()
            break
    return vendor_id, product_id




def obtener_ruta_descargas():
    key_path = r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path) as key:
        downloads_directory = winreg.QueryValueEx(key, "{374DE290-123F-4565-9164-39C4925E467B}")[0]
    # Expandir la variable de entorno para obtener la ruta completa
    downloads_directory = os.path.expandvars(downloads_directory)
    return downloads_directory


def imprimir_pdf(pdf_path) -> None:
    #Obtner los ide de la impresora termica
    printer_ids = get_usb_printer_ids()
    #print(printer_ids)
    for pnp_device_id in printer_ids:
        vendor_id, product_id = get_vendor_product_ids(pnp_device_id)
        if vendor_id  and product_id:
            # Inicializa la impresora térmica
            printer = Usb(vendor_id,product_id)
            # Abre el archivo PDF
            pdf_document = fitz.open(pdf_path)
            # Itera sobre cada página del PDF
            for page_number in range(len(pdf_document)):
                # Obtiene la página
                page = pdf_document.load_page(page_number)
                # Renderiza la página como una imagen (formato PNG)
                image = page.get_pixmap()
                # Abre la imagen usando PIL
                pil_image = Image.frombytes("RGB", [image.width, image.height], image.samples)
                # Escala la imagen para que se ajuste al ancho de la impresora térmica
                width, height = pil_image.size
                new_width = 384  # Ancho de la mayoría de las impresoras térmicas
                new_height = int((new_width / width) * height)
                pil_image = pil_image.resize((new_width, new_height))

                # Convierte la imagen a escala de grises (opcional, pero común para impresoras térmicas)
                pil_image = pil_image.convert("L")

                # Imprime la imagen en la impresora térmica
                printer.image(pil_image)

            # Corta el papel después de imprimir todas las imágenes
            printer.cut()

            # Cierra el documento PDF
            pdf_document.close()

            os.remove(path=pdf_path)            
        else:
           os.remove(path=pdf_path)
           raise Exception('**No se encontro ninguna impresora**')
    
    

def main():
    descargas = obtener_ruta_descargas()

    # Obtener una lista de archivos en el directorio de descargas
    while True:
        descargas = obtener_ruta_descargas()
        pdfs = list(archivo for archivo in  os.listdir(descargas) if archivo.endswith('.pdf'))
        for file in pdfs:
            # Verificación de tickets del colegio
            list_name_file = file.split('_')
            if 'jvtk' in list_name_file or 'jvtk' in list_name_file:
                path = rf'{descargas}\{file}'
                #print(path)
                try:
                    imprimir_pdf(path)
                except Exception as e:
                    print(e)
                break    
        # Actualización del estado cada 5 segundos
        sleep(5)

if __name__ == "__main__":
    try:
        main()
    except Exception as  e:
        print(f'**Error al ejecutar el script: {e}**')

