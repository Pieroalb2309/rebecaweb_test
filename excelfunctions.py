import openpyxl
import os

def start_webdriver():
    print("Select the browser( that you have installed in your PC) to test:Chrome,Opera,Safari,Edge,Firefox")
    input1 = input()
    if (input1 == "Chrome") or (input1 == "chrome"):
        a=os.path.abspath("msedgedriver.exe")
    elif (input1 == "Opera") or (input1 == "opera"):
        a=os.path.abspath("msedgedriver.exe")
    elif (input1 == "Safari") or (input1 == "safari"):
        a=os.path.abspath("msedgedriver.exe")
    elif (input1 == "Edge") or (input1 == "edge"):
        a=os.path.abspath("msedgedriver.exe")
    elif (input1 == "Firefox") or (input1 == "firefox"):
        a=os.path.abspath("msedgedriver.exe")
    print(input1)

def getrowcount(file,sheetname):
    workbook = openpyxl.load_workbook(file)
    sheet = workbook.get_sheet_by_name(sheetname)
    return (sheet.max_row)


def getcolumncount(file, sheetname):
    workbook = openpyxl.load_workbook(file)
    sheet = workbook.get_sheet_by_name(sheetname)
    return (sheet.max_column)


def readdata(file, sheetname, rownum, columnno):
    workbook = openpyxl.load_workbook(file)
    sheet = workbook.get_sheet_by_name(sheetname)
    return sheet.cell(row=rownum, column=columnno).value


def writedata(file, sheetname, rownum, columnno, data):
    workbook = openpyxl.load_workbook(file)
    sheet = workbook.get_sheet_by_name(sheetname)
    sheet.cell(row=rownum, column=columnno).value = data
    workbook.save(file)


def weeklist(text): #mandar los que no van
    l=['0','1','2','3','4','5','6']
    if text != "":
        a = text.split(",")
        print(a)
        for i in range(len(a)):
            c = a[i]
            if c[0] == ' ':
                a[i]=c[1:]
        #'0','1','2','3','4','5','6'
        print(a)
        for i in range(len(a)):
            l.remove(a[i])
            print(l)
            #a: arreglo q indica los valores que no estan en el arreglo original

        for j in range(len(l)):
            if l[j] == "0":
                l[j] = "Domingo"
            elif l[j] == "1":
                l[j] = "Lunes"
            elif l[j] == "2":
                l[j] = "Martes"
            elif l[j] == "3":
                l[j] = "Miércoles"
            elif l[j] == "4":
                l[j] = "Jueves"
            elif l[j] == "5":
                l[j] = "Viernes"
            else:
                l[j] = "Sábado"
        print(l)
        b=l
    else:
        b = []
    return b

def permissions(array):
    #AdminDefault, ClientViewer, Euler, ProviderViewer, ServicesEditor, TouroperatorDefault, UserViewer
    a=array.split(',')
    b=[]
    for i in range(len(a)):
        c=a[i]
        if c[0]!=' ':
            b.append(c)
        else:
            b.append(c[1:])
    #print(b)
    return b

