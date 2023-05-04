from exportCSV import *
from tkinter import *
from tkinter import filedialog
from time import *
from smallFunctions import *
import openpyxl

#from xlrd import open_workbook
import xlwt
#from xlutils.copy import copy
import os.path

def getFolderPath():
    root = Tk()
    root.withdraw()
    path = filedialog.askopenfilename(initialdir="C:\\Users\\Julien\\Desktop\\Prototype\\Thermal_Camera")
    return path[:64]


if __name__ == "__main__":

    initial_path=getFolderPath()
    print(initial_path)

    steps=['AftPremixDisp','AftPremixInc']
    #wells=[["Left up",11,46],["Left middle",42,45],["Left down",83,46],["Left-1 up",11,55],["Left-1 middle",41,55],["Left-1 down",83,56],["Middle up",10,66],["Middle middle",44,93],["Middle down",85,95],["Right up",12,123],["Right middle",43,124],["Right down",85,125]]
    wells=[]
    A1=[25,11] #read in the order of Optris
    H12=[137,82]
    for well in range(1,96+1):
        coordinates=returnXY(well,A1[0],A1[1],H12[0],H12[1])
        wells.append([getWellName(well),coordinates[1],coordinates[0]])

    temp_excel = xlwt.Workbook()
    sheets=[]
    for step in steps:
        sheets.append(temp_excel.add_sheet(step))
        sheets[steps.index(step)].write(0, 0, "Cycle")
        for well in wells:
            sheets[steps.index(step)].write(0, 1+wells.index(well), well[0]+"("+str(well[1])+";"+str(well[2])+")")

    subdirs = [x[0].split("\\")[-1] for x in os.walk(initial_path)]
    subdirs.pop(0) #we remove first element as it's the initial folder
    nb_cycles=max(int(x) for x in subdirs)

    for cycle in range(1,nb_cycles+1):
        folderPath=initial_path+"\\"+str(cycle)
        files=[x for x in os.listdir(folderPath)]
        print(files)
        for step in steps:
            for file in files:
                if file.find(step)>-1:
                    pathString=folderPath+"\\"+file
                    openFile(pathString.replace("/","\\"))
                    outputString=initial_path+"\\temp.csv"
                    exportToCsv(outputString.replace("/","\\"))
                    sleep(0.2)
                    tempMatrix=readCsv(outputString.replace("/","\\"))

                    sheets[steps.index(step)].write(cycle,0,cycle)
                    for well in wells:
                        sheets[steps.index(step)].write(cycle,1+wells.index(well),tempMatrix[well[1]][well[2]])

    print(nb_cycles)
    temp_excel.save(initial_path+"\\temp_"+ initial_path[len(initial_path)-15:] + ".xls")
