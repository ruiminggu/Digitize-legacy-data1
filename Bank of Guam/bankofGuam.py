import re
import os
from openpyxl import Workbook
from datetime import datetime

from convert_to_txt import split_pdf, convert_folder_images, convert_folder_pdf_to_png

workbook =  Workbook()
sheet = workbook.active

def read_PIT(file_path):
    with open(file_path,'r', encoding='utf-8') as file:
        content = file.read()
        lines = content.splitlines()
        lines = [line for line in lines if line.strip()]
        name = ""
        occupation = ""
        address = ""
        city =""
        accountnumber =""
        ctrid=""
        cashin=""
        cashout =""
        idtype = ""


        for i in range(len(lines)):
            if lines[i] == "":
                pass
            else:
                if "Cash in" in lines[i]:
                    cashin += lines[i]
                if "license" in lines[i]:
                    idtype += lines[i]
                if "Cash out" in lines[i]:
                    cashout += lines[i]
                if "last" in lines[i]and i != len(lines)-1:
                    name += lines[i+1]
                if "Filing Name" in lines[i]:
                    ctrid +=lines[i];
                if "Address" in lines[i]and i != len(lines)-1:
                    address += lines[i]
                if "City" in lines[i] and i != len(lines)-1:
                    city += lines[i]
                if "business" in lines[i]and i != len(lines)-1:
                    occupation += lines[i]
                if "Account" in lines[i] and i != len(lines)-1:
                    accountnumber += lines[i]
        if ctrid!="" or name!= "":
            sheet.append([ctrid,name,address,occupation,city ,idtype, cashin,cashout,accountnumber])

def read_folder(folder_path,output_path):
    folder_path = folder_path + "/png_files/txt_files"
    sheet.append(["CTRID","lastNameOrNameOfEntity","address","occupationOrTypeOfBusiness","addressCity","idType","cashin","cashout","accountNumbers"])
    for filename in os.listdir(folder_path):
        txt_file_path = os.path.join(folder_path, filename)
        read_PIT(txt_file_path)
    workbook.save(output_path)
if __name__ == '__main__':
    # the first should enter the name of the destination folder, the second is the pdf name
    split_pdf("t1ry1", "bankofGuam1.pdf")
    # enter the first again.
    convert_folder_images(convert_folder_pdf_to_png("t1ry1"))
    # enter the first again.
    read_folder("t1ry1","tey111.xlsx")
