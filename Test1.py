import xlwings as xw
import datetime as dt


@xw.sub
def nieuwe_verzekering():
    """
    Deze functie wordt aangeroepen als er op de knop wordt gedrukt
    """
    book = xw.Book.caller()
    mainSheet = book.sheets["Main"]
    geslacht = mainSheet.range('C4').value
    geboortedatum = mainSheet.range('C5').options(dates=dt.date).value
    Pensioenleeftijd = mainSheet.range('C6').options(numbers=int).value
    rente = mainSheet.range('C7').options(numbers=float).value
    Sterftetafel = mainSheet.range('C8').value

    
    
    leeftijdLijst = ["Leeftijd"]
    for i in range(Pensioenleeftijd, 120): leeftijdLijst += [i]


    totNuLijst = ["Aantal jaar tot nu"]
    verschilJaar = dt.datetime.now().year - geboortedatum.year
    verschilMaand = dt.datetime.now().month - geboortedatum.month
    if verschilMaand < 0:
        verschilJaar -= 1
        verschilMaand +=12
    verschilTotaal = verschilJaar + verschilMaand/12
    for i in range(Pensioenleeftijd, 120): totNuLijst += [i - verschilTotaal]

    renteLijst = ["Rente"]
    for i in range(Pensioenleeftijd, 120): renteLijst += [rente]

    leefLijst = ["Levenskans"]
    if Sterftetafel == "AEGON_2011": sterfkolom = 2
    else: sterfkolom = 5
    huidigLeef = book.sheets["sterftetafels"].range((verschilJaar + 4, sterfkolom)).options(numbers=int).value *\
                 (1 - verschilMaand / 12) +\
                 book.sheets["sterftetafels"].range((verschilJaar + 5, sterfkolom)).options(numbers=int).value *\
                 verschilMaand / 12
    for i in range(Pensioenleeftijd, 120):
        leefkans = book.sheets["sterftetafels"].range((i + 4, sterfkolom)).value
        leefLijst += [leefkans / huidigLeef]

    cwLijst = ["cw-factor"]
    cwLijst += [(1+r)**-jaar for r, jaar in zip(renteLijst[1:], totNuLijst[1:])]

    rkLijst = ["rente*kans"]
    rkLijst += [cw * leven for cw, leven in zip(cwLijst[1:], leefLijst[1:])]


    newSheet = book.sheets.add("OP1")
    newSheet.api.Move(None, After = mainSheet.api)
    mainSheet.range("B3", "D9").copy(newSheet.range("A2", "C8"))
    newSheet.range("A10").options(transpose=True).value = leeftijdLijst
    newSheet.range("B10").options(transpose=True).value = totNuLijst
    newSheet.range("C10").options(transpose=True).value = renteLijst
    newSheet.range("D10").options(transpose=True).value = leefLijst
    newSheet.range("E10").options(transpose=True).value = cwLijst
    newSheet.range("F10").options(transpose=True).value = rkLijst
    
    
    print("Hello world")

    
