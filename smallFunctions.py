


def getWellName(wellNumber):
    col=((wellNumber-1)//8)+1
    row=((wellNumber-1)%8)+1
    row=chr(ord("A")+row-1)

    return row + str(col)

def returnXY(wellnumber,A1,H1,A12):

    col_well=((wellnumber-1)//8)
    row_well=((wellnumber-1)%8)

    x=A1[0]+(col_well*(A12[0]-A1[0])/11)+(row_well*(H1[0]-A1[0])/7)
    y=A1[1]+(col_well*(A12[1]-A1[1])/11)+(row_well*(H1[1]-A1[1])/7)

    return [round(x),round(y)]

def returnXYold(wellnumber,X0,Y0,X96,Y96):

    cosAlpha=(X96-X0)/114
    sinAlpha=(Y96-Y0)/71

    x=(((wellnumber-1)//8)*(114/11))*cosAlpha+X0
    y=(((wellnumber-1)%8)*(72/7))*sinAlpha+Y0

    return [round(x),round(y)]

def excelColumn(columnNumber):

    columnName=""

    if columnNumber>26:
        columnName=columnName + chr(ord("A")+((columnNumber-1)//26)-1)

    columnName= columnName + chr(ord("A")+((columnNumber-1)%26))
    return columnName

if __name__ == "__main__":

    print(returnXY(96,[31,10],[32,80],[141,9]))

    #print(excelColumn(96))