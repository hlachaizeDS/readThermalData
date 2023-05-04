from pywinauto import *
import csv
import os
import time
import array

def openFile(input_file):
    app = Application().connect(title_re="Optris PIX Connect")
    window = app.window(title_re="Optris PIX Connect")
    toolbar = window.Toolbar
    toolbar.set_focus()
    toolbar.draw_outline()
    time.sleep(0.5)
    toolbar.click(button='left', pressed='', coords=(20, 20), double=False, absolute=False)
    window = app.window(title_re="Open file")
    window.type_keys(input_file)
    window.type_keys("{ENTER}")
    time.sleep(0.5)
    try :
        window = app.window(title_re="attention")
        window.type_keys("{ENTER}")

    except:
        None


def exportToCsv(output_file):

    #If ouput_file already exists, we delete it

    if os.path.isfile(output_file) != 0:
        os.remove(output_file)

    app = Application().connect(title_re="Optris PIX Connect")
    window = app.window(title_re="Optris PIX Connect")
    window.type_keys("{F1}")
    window = app.window(title_re="Enregistrer sous")
    window=window.DirectUIHWND
    typebar=window.ComboBox2
    typebar.select("Text(Image data) (*.csv)")
    pathbar=window.Combox1
    pathbar.type_keys(output_file)
    pathbar.type_keys("{ENTER}")
    #window.print_control_identifiers()
    #pathbar.set_focus()
    #pathbar.draw_outline()

def readCsv(input_file):

    temp_matrix=[]

    csv.register_dialect('myDialect', delimiter=';')

    with open(input_file, 'r') as csvFile:
        reader = csv.reader(csvFile,dialect='myDialect')
        for row in reader:
            row.pop() #each row finishes by a ; so we need to remove one temp
            temp_matrix.append(row)

    return temp_matrix

    csvFile.close()


if __name__ == "__main__":
    # On crée la racine de notre interface
    #exportToCsv('test','test')

    #input_file="C:\\Users\\Julien\\Desktop\\Prototype\\Thermal_Camera\\20190423_173027\\1\\20190423_173028_C1_BeforeAnything.tiff"
    #output_file="C:\\Users\\Julien\\Desktop\\Prototype\\Thermal_Camera\\20190423_173027\\1\\20190423_173028_C1_BeforeAnything.csv"

    #openFile("C:\\Users\\Julien\\Desktop\\Prototype\\Thermal_Camera\\20190423_173027\\1\\20190423_173028_C1_BeforeAnything.tiff")
    #exportToCsv("C:\\Users\\Julien\\Desktop\\Prototype\\Thermal_Camera\\20190423_173027\\1\\20190423_173028_C1_BeforeAnything")
    #temp_array=readCsv(output_file)
    #print(temp_array[0][7])

    print(getWellName(1))