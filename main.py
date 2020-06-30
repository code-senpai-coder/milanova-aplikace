import openpyxl, time
wb = openpyxl.load_workbook("test.xlsx")
sheet = wb.get_sheet_by_name("List1")

def findLastRow(tabulka):
    main = True
    index = 1

    while main:
        if tabulka["A" + str(index)].value != None:
            index += 1
        else:
            main = False
            return index
def addNewCard(cardId, tabulka, workbook, radek):
    tabulka["A" + str(radek)].value = cardId
    tabulka["B" + str(radek)].value = 0
    workbook.save("test.xlsx")
    print("Karta uspesne pridana\n")
def showCardInfo(idCard, tabulka):
    main = True
    index = 1
    row = 0
    while main:
        if tabulka["A" + str(index)].value == idCard:
            row = index
            main = False
        else:
            index += 1
    cardBalance = tabulka["B" + str(row)].value
    return cardBalance
def widtdrawFromCard(idCard, tabulka, amount, workbook):
    main = True
    index = 1
    row = 0
    while main:
        if tabulka["A" + str(index)].value == idCard:
            row = index
            main = False
        else:
            index += 1
    if sheet["B" + str(row)].value - amount >= 0:
        sheet["B" + str(row)].value -= amount 
        print("Z karty bylo uspesne odebrano " + amount + "kc \n")
    else:
        print("Nedostatek penez\n")
    workbook.save("test.xlsx")
def addToCard(idCard, tabulka, amount, workbook):
    main = True
    index = 1
    row = 0
    while main:
        if tabulka["A" + str(index)].value == idCard:
            row = index
            main = False
        else:
            index += 1
    tabulka["B" + str(row)].value += amount
    print("Na kartu bylo supesne pridano" + amount + "kc\n")
    workbook.save("test.xlsx")
def checkUserWantedFunction():
    userInput = input("Zobrazeni informaci [1]\nPridani penez na kartu [2]\nOdebrani penez z karty [3]\nOdebrani karty[4]\nZavreni aplikace[e]\nVytvoreni nove karty[n]\n")
    return userInput
def deleteCard(cardId, tabulka, workbook):
    main = True
    index  = 1
    while main:
        if tabulka["A" + str(index)].value == cardId:
            tabulka["A" + str(index)].value = None
            tabulka["B" + str(index)].value = None
            main = False
        else:
            index +=1
    wb.save("test.xlsx")
    print("Karta uspesne odebrana")

#wdeleteCard(222, sheet, wb)
#print(showCardInfo(123, sheet))
#widtdrawFromCard(123, sheet, 20, wb)
#addToCard(123, sheet, 34, wb)

"""
functions above
app code below
"""
#addNewCard(111, sheet, wb, findLastRow(sheet))
def mainDef():
    userIn = str(checkUserWantedFunction())
    if userIn == "1":
        print("Pocet penez na ucte je " + str(showCardInfo(int(input('Zadej cislo karty\n')), sheet)) + "kc")
        time.sleep(0.3)
        mainDef()
    elif userIn == "2":
        addToCard(int(input("Zadej cislo karty\n")), sheet, float(input("Zadej mnozstvi penez na vlozeni\n")), wb)
        time.sleep(0.3)
        mainDef()
    elif userIn == "3":
        widtdrawFromCard(int(input("Zadej cislo karty\n")), sheet, float(input("Zadej mnozstvi penez na vybrani\n")), wb)
        time.sleep(0.3)
        mainDef()
    elif userIn == "n":
        addNewCard(int(input("Zadej cislo nove karty")), sheet, wb, findLastRow(sheet))
        mainDef()
    elif userIn == "4":
        deleteCard(int(input("Zadej cislo karty kterou chces odstranit\n")),sheet, wb )
        mainDef()
    elif userIn == "e":
        pass
    else:
        print("Spatny input\n")
        time.sleep(0.3)
        mainDef()
mainDef()
