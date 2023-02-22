"""
Header
Hier komen alle libraries die in het programma gebruikt worden
"""
import xlwings as xw
import datetime as dt
import matplotlib.pyplot as plt


"""
Body
Hier komen alle functies
"""

#Wachtwoord = "wachtwoord"

#import sys
#from PyQt5 import QtWidgets, uic
#Ui_MainWindow, QtBaseClass = uic.loadUiType("1AdviseurBeheerder.ui")
#class functiekeus(QtWidgets.QMainWindow):
#    def __init__(self):
#        super(functiekeus, self).__init__()
#        self.ui = Ui_MainWindow()
#        self.ui.setupUi(self)
#        self.ui.BtnAdviseur.clicked.connect(self.BtnAdviseurClicked)
#        self.ui.BtnBeheerder.clicked.connect(self.BtnBeheerderClicked)
        
        
#    def BtnAdviseurClicked(self):
#        self.close()
#        self._windowdeelnemer = deelnemerselectie()
#        self._windowdeelnemer.show()
#    def BtnBeheerderClicked(self): 
#        self.close()
#        self._windowinlog = inloggen()
#        self._windowinlog.show()
        

#Ui_MainWindow2, QtBaseClass2 = uic.loadUiType("2InlogBeheerder.ui")
#class inloggen(QtWidgets.QMainWindow):
#    def __init__(self):
#        super(inloggen, self).__init__()
#        self.ui = Ui_MainWindow2()
#        self.ui.setupUi(self)
#        self.ui.BtnTerug.clicked.connect(self.BtnTerugClicked)
#        self.ui.BtnInloggen.clicked.connect(self.BtnInloggenClicked)
        
#    def BtnInloggenClicked(self):
#        if self.ui.txtBeheerderscode.text() == Wachtwoord:
#            self.close()
#            self._windowdeelnemer = deelnemerselectie()
#            self._windowdeelnemer.show()
#        else:
#            self.ui.label.setText("Wachtwoord incorrrect")
#    def BtnTerugClicked(self):
#        self.close()
#        self._windowkeus = functiekeus()
#        self._windowkeus.show()
        
        
#Ui_MainWindow3, QtBaseClass3 = uic.loadUiType("deelnemerselectie.ui")
#class deelnemerselectie(QtWidgets.QMainWindow):
#    def __init__(self):
#        super(deelnemerselectie, self).__init__()
#        self.ui = Ui_MainWindow3()
#        self.ui.setupUi(self)    
#        self.ui.pushButton.clicked.connect(self.pushButtonClicked)
        
#    def pushButtonClicked(self):
#        self.close()
        
#Ui_MainWindow4, QtBaseClass3 = uic.loadUiType("4DeelnemerToevoegen.ui")
#class deelnemertoevoegen(QtWidgets.QMainWindow):
#    def __init__(self):
#        super(deelnemertoevoegen, self).__init__()
#        self.ui = Ui_MainWindow4()
#        self.ui.setupUi(self)
#        self.ui.BtnTerug.clicked.connect(self.BtnTerugClicked)
#        self.ui.BtnToevoegen.clicked.connect(self.BtnToevoegenClicked)
        
#    def BtnTerugClicked(self):
#        self.close()
#        self._windowdeelnemer = deelnemerselectie()
#        self._windowdeelnemer.show()
        
#    def BtnToevoegen(self):
#        self.close()
#        self._windowdeelnemer = deelnemerselectie()
#        self._windowdeelnemer.show()


#Ui_MainWindow5, QtBaseClass3 = uic.loadUiType("flexmenu.ui")
#class flexmenu(QtWidgets.QMainWindow):
#    def __init__(self):
#        super(flexmenu, self).__init__()
#        self.ui = Ui_MainWindow5()
#        self.ui.setupUi(self)
#        self.ui.pushButton.clicked.connect(self.pushButtonClicked)
#        self.ui.pushButton_2.clicked.connect(self.pushButton_2Clicked)
#        self.ui.pushButton_3.clicked.connect(self.pushButton_3Clicked)
        
#    def pushButtonClicked(self):
#        self.close()
#        self._windowdeelnemer = deelnemerselectie()
#        self._windowdeelnemer.show()
        
#    def pushButton_2Clicked(self):
#        self.close()
        
#    def pushButton_3Clicked(self):
#        self.close()
        

   
#if __name__ == "__main__":
#    app = 0
#    app = QtWidgets.QApplication(sys.argv)
#    window = functiekeus()
#    window.show()
#    sys.exit(app.exec_())


@xw.sub
def vergelijken_afbeelding_generatie():
    """
    Functie die de data leest en vevolgens een afbeelding genereerd op basis van de data
    """
    
    #sheets en book opslaan in variabelen
    book = xw.Book.caller()
    invoer = book.sheets["Tijdelijk"]
    uitvoer = book.sheets["Vergelijken"]
    
    #data lezen
    titel = (3,1)
    
    beginrij = 5
    OPbeginkolom = 2
    blokafstand = 6
    
    PPbeginkolom = 5
    PPblokafstand = 3
    
    PPnaam = 0
    PPjaarbedrag = 1
    
    OPnaam = 0
    OPbeginjaar = 1
    OPjaarbedrag = 2
    OPhooglaaggrens = 3
    OPverhouding = 4
    
    #Aantal blokken tellen
    blokaantal = blokkentellen(beginrij, OPbeginkolom, blokafstand, invoer)
    PPblokaantal = blokkentellen(beginrij, PPbeginkolom, PPblokafstand, invoer)

    #Lijst met alle voorkomende jaren van OP
    allejaren = set()
    for blok in range(blokaantal):
        allejaren.add(invoer.range((beginrij + OPbeginjaar + blok * blokafstand, OPbeginkolom)).options(numbers = float).value)
        allejaren.add(invoer.range((beginrij + OPhooglaaggrens + blok * blokafstand, OPbeginkolom)).options(numbers = float).value)
    
    #geeft de breedte aan van alle hoogtes
    randen = list(allejaren)
    randen.sort()
    randen.append(randen[-1] + 10)
    
    #een lijst met alle verzekeringsnamen
    naamlijst = list()
    for blok in range(blokaantal): naamlijst.append(invoer.range((beginrij + OPnaam + blok * blokafstand, OPbeginkolom)).value)
    
    #berekent de hoogte van elke staaf
    hoogtes = [[0 for i in range(len(randen)-1)]]
    ywaardes = set()
    ywaardes.add(0)
    
    for blok in range(blokaantal):
        startjaar = float(invoer.range((beginrij + OPbeginjaar + blok * blokafstand, OPbeginkolom)).options(numbers = float).value)
        toezegging = float(invoer.range((beginrij + OPjaarbedrag + blok * blokafstand, OPbeginkolom)).options(numbers = float).value)
        laaghoogverhouding = float(invoer.range((beginrij + OPverhouding + blok * blokafstand, OPbeginkolom)).options(numbers = float).value)
        alternatiefjaar = float(invoer.range((beginrij + OPhooglaaggrens + blok * blokafstand, OPbeginkolom)).options(numbers = float).value)
        
        hoogtes.append(list())

        for i, leeftijd in enumerate(randen[:-1]):
            if leeftijd >= alternatiefjaar:
                bedrag = float(hoogtes[blok][i] + toezegging * laaghoogverhouding)
                hoogtes[blok+1].append(bedrag)
                ywaardes.add(bedrag)
            elif leeftijd >= startjaar:
                bedrag = float(hoogtes[blok][i] + toezegging)
                hoogtes[blok+1].append(bedrag)
                ywaardes.add(bedrag)
            else: hoogtes[blok+1].append(hoogtes[blok][i])
    ywaardes = list(ywaardes)
    ywaardes.sort()

    #bereken PP
    PPtotaal = 0
    for blok in range(PPblokaantal):
        PPtotaal += float(invoer.range((beginrij + PPjaarbedrag + blok * PPblokaantal, PPbeginkolom)).options(numbers = float).value)
        


    #maak de afbeeling
    afbeelding = plt.figure()
    for i in range(len(hoogtes) - 1):
        plt.stairs(hoogtes[i+1],edges = randen,  baseline=hoogtes[i], fill=True, label = naamlijst[i])
    
    plt.xticks(randen[:-1], [getaltotijd(rand) for rand in randen[:-1]])
    plt.setp(plt.gca().get_xticklabels(), rotation=30, horizontalalignment='right')
    plt.yticks(ywaardes, [getaltogeld(ywaarde) for ywaarde in ywaardes])

    handles, labels = plt.gca().get_legend_handles_labels()
    order = range(blokaantal-1, -1, -1)
    plt.legend([handles[idx] for idx in order],[labels[idx] for idx in order]) 
    
    plt.gca().set_title("Totale partnerpensioen: €{:.2f}".format(PPtotaal).replace(".",","))
    plt.suptitle(invoer.range(titel).value, fontweight='bold')
    
    
    uitvoer.pictures.add(afbeelding, top = uitvoer.range((3,3)).top, left = uitvoer.range((3,3)).left, height = 300)
    
    

def getaltotijd(getal):
    jaar = int(getal)
    maand = round((getal - jaar) * 12)
    tijd = "{}j".format(jaar)
    if maand > 0: tijd = tijd + " {}m".format(maand)
    return tijd

def getaltogeld(getal): return "€{:.2f}".format(float(getal)).replace(".",",")

def blokkentellen(beginrij, beginkolom, blokafstand, sheet):
    blokaantal = 0
    leescell = [beginrij, beginkolom]
    while sheet.range(tuple(leescell)).value != None:
        blokaantal += 1
        leescell[0] +=blokafstand
    return blokaantal


@xw.sub
def invoer_test_klikken():
    
    #sheets en book opslaan in variabelen
    book = xw.Book.caller()
    input = book.sheets["Tijdelijk invoerscherm"]
    
    #pensioenvormen=invoer.range("B10:B18")
    kolom_t=input.range("I2")
    if input.range("B10").value != "":
        rente= input.range("D10").value
        pensioenleeftijd= input.range("E10").value
        
    i=pensioenleeftijd
    x=1
    while i<=119:
        kolom_t(x).value=i
        x=x+1
        i=i+1
        
        
    
    
    



    