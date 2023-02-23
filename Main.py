"""
Header
Hier komen alle libraries die in het programma gebruikt worden
"""
import xlwings as xw
import datetime as dt
import matplotlib.pyplot as plt
import sys
from PyQt5 import QtWidgets, uic
from string import ascii_uppercase
import functions


"""
Body
Hier komen alle functies
"""
@xw.sub
def Schermen():  
    Wachtwoord = "wachtwoord"
    
    
    Ui_MainWindow, QtBaseClass = uic.loadUiType("{}\\1AdviseurBeheerder.ui".format(sys.path[0]))
    class functiekeus(QtWidgets.QMainWindow):
        def __init__(self):
            super(functiekeus, self).__init__()
            self.ui = Ui_MainWindow()
            self.ui.setupUi(self)
            self.ui.btnAdviseur.clicked.connect(self.btnAdviseurClicked)
            self.ui.btnBeheerder.clicked.connect(self.btnBeheerderClicked)
            
            
        def btnAdviseurClicked(self):
            self.close()
            self._windowdeelnemer = deelnemerselectie()
            self._windowdeelnemer.show()
        def btnBeheerderClicked(self): 
            self.close()
            self._windowinlog = inloggen()
            self._windowinlog.show()
            
    
    Ui_MainWindow2, QtBaseClass2 = uic.loadUiType("{}\\2InlogBeheerder.ui".format(sys.path[0]))
    class inloggen(QtWidgets.QMainWindow):
        def __init__(self):
            super(inloggen, self).__init__()
            self.ui = Ui_MainWindow2()
            self.ui.setupUi(self)
            self.ui.btnTerug.clicked.connect(self.btnTerugClicked)
            self.ui.btnInloggen.clicked.connect(self.btnInloggenClicked)
            
        def btnInloggenClicked(self):
            if self.ui.txtBeheerderscode.text() == Wachtwoord:
                self.close()
                self._windowdeelnemer = deelnemerselectie()
                self._windowdeelnemer.show()
            else:
                self.ui.lblFoutmeldingInlog.setText("Wachtwoord incorrrect")
        def btnTerugClicked(self):
            self.close()
            self._windowkeus = functiekeus()
            self._windowkeus.show()
            
            
    Ui_MainWindow3, QtBaseClass3 = uic.loadUiType("{}\\deelnemerselectie.ui".format(sys.path[0]))
    class deelnemerselectie(QtWidgets.QMainWindow):
        def __init__(self):
            super(deelnemerselectie, self).__init__()
            self.ui = Ui_MainWindow3()
            self.ui.setupUi(self)    
            self.ui.btnDeelnemerToevoegen.clicked.connect(self.btnDeelnemerToevoegenClicked)
            self.ui.btnStartFlexibiliseren.clicked.connect(self.btnStartFlexibiliserenClicked)
            self.ui.btnTerug.clicked.connect(self.btnTerugClicked)
            self.ui.sbMaand.valueChanged.connect(self.maandchange)
            self.ui.sbJaar.valueChanged.connect(self.jaarchange)
            self._30maand = [4,6,9,11]
            
        def btnDeelnemerToevoegenClicked(self):
            self.close()
            self._windowtoevoeg = deelnemertoevoegen()
            self._windowtoevoeg.show()
            
        def btnStartFlexibiliserenClicked(self):
            self.close()
            self._windowflex = flexmenu()
            self._windowflex.show()
            
        def btnTerugClicked(self):
            self.close()
            self.windowstart = functiekeus()
            self.windowstart.show()
        
        def maandchange(self):
            if self.ui.sbMaand.value() in self._30maand:
                self.ui.sbDag.setMaximum(30)
            elif self.ui.sbMaand.value() == 2:
                if self.ui.sbJaar.value()%4 == 0:
                    self.ui.sbDag.setMaximum(29)
                else: 
                    self.ui.sbDag.setMaximum(28)
            else:
                self.ui.sbDag.setMaximum(31)
            
        def jaarchange(self):
            if self.ui.sbMaand.value() == 2:
                if self.ui.sbJaar.value()%4 == 0:
                    self.ui.sbDag.setMaximum(29)
                else: 
                    self.ui.sbDag.setMaximum(28)
            
            
    Ui_MainWindow4, QtBaseClass4 = uic.loadUiType("{}\\4DeelnemerToevoegen.ui".format(sys.path[0]))
    class deelnemertoevoegen(QtWidgets.QMainWindow):
        def __init__(self):
            super(deelnemertoevoegen, self).__init__()
            self.ui = Ui_MainWindow4()
            self.ui.setupUi(self)
            self.ui.btnTerug.clicked.connect(self.btnTerugClicked)
            self.ui.btnToevoegen.clicked.connect(self.btnToevoegenClicked)
            self.ui.sbMaand.valueChanged.connect(self.maandchange)
            self.ui.sbJaar.valueChanged.connect(self.jaarchange)
            self._30maand = [4,6,9,11]
            
        def btnTerugClicked(self):
            self.close()
            self._windowdeelnemer = deelnemerselectie()
            self._windowdeelnemer.show()
            
        def btnToevoegenClicked(self):
            #In de UI staat nog geen vak voor foutmeldingen, dus heb ik het nu even
            #korte komandos gegeven (Zodat hij ook geen foutmelding geeft)
            if self.ui.txtVoorletters.text() == "" or self.ui.txtAchternaam.text() == "":
                fout = "Aanwezig"
                #print("Naam gegevens incompleet")
            elif self.ui.txtFulltimeLoon.text() == "" or self.ui.txtParttimePercentage.text() == "":
                fout = "Aanwezig"
                #print("Loon informatie incompleet")
            else:
                self.close()
                self._windowdeelnemer = deelnemerselectie()
                self._windowdeelnemer.show()
            
        def maandchange(self):
            if self.ui.sbMaand.value() in self._30maand:
                self.ui.sbDag.setMaximum(30)
            elif self.ui.sbMaand.value() == 2:
                if self.ui.sbJaar.value()%4 == 0:
                    self.ui.sbDag.setMaximum(29)
                else: 
                    self.ui.sbDag.setMaximum(28)
            else:
                self.ui.sbDag.setMaximum(31)
            
        def jaarchange(self):
            if self.ui.sbMaand.value() == 2:
                if self.ui.sbJaar.value()%4 == 0:
                    self.ui.sbDag.setMaximum(29)
                else: 
                    self.ui.sbDag.setMaximum(28)
    
    
    Ui_MainWindow5, QtBaseClass5 = uic.loadUiType("{}\\flexmenu.ui".format(sys.path[0]))
    class flexmenu(QtWidgets.QMainWindow):
        def __init__(self):
            super(flexmenu, self).__init__()
            self.ui = Ui_MainWindow5()
            self.ui.setupUi(self)
            self.ui.btnAndereDeelnemer.clicked.connect(self.btnAndereDeelnemerClicked)
            self.ui.btnVergelijken.clicked.connect(self.btnVergelijkenClicked)
            self.ui.btnOpslaan.clicked.connect(self.btnOpslaanClicked)
            
        def btnAndereDeelnemerClicked(self):
            self.close()
            self._windowdeelnemer = deelnemerselectie()
            self._windowdeelnemer.show()
            
        def btnVergelijkenClicked(self):
            self.close()
            
        def btnOpslaanClicked(self):
            self.close()
            


    if __name__ == "__main__":
        app = 0
        app = QtWidgets.QApplication(sys.argv)
        window = functiekeus()
        window.show()
        sys.exit(app.exec_())



@xw.sub
def vergelijken_afbeelding_generatie():
    """
    Functie die de data leest en vevolgens een afbeelding genereerd op basis van de data
    """
    
    #sheets en book opslaan in variabelen
    book = xw.Book.caller()
    invoer = book.sheets["Tijdelijk afbeelding"]
    uitvoer = book.sheets["Vergelijken"]
    
    #data lezen
    titel = (3,1)
    
    beginrij = 5
    OPbeginkolom = 2
    blokafstand = 7
    
    PPbeginkolom = 5
    PPblokafstand = 3
    
    PPnaam = 0
    PPjaarbedrag = 1
    
    OPnaam = 0
    OPkleur = 1
    OPbeginjaar = 2
    OPjaarbedrag = 3
    OPhooglaaggrens = 4
    OPverhouding = 5
    
    #Aantal blokken tellen
    blokaantal = functions.blokkentellen(beginrij, OPbeginkolom, blokafstand, invoer)
    PPblokaantal = functions.blokkentellen(beginrij, PPbeginkolom, PPblokafstand, invoer)

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
    
    #bepaald de kleuren
    kleuren = list()
    for blok in range(blokaantal): kleuren.append(functions.kleurinvoer(invoer.range((beginrij + OPkleur + blok * blokafstand, OPbeginkolom)).value))

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
        plt.stairs(hoogtes[i+1],edges = randen,  baseline=hoogtes[i], fill=True, label = naamlijst[i], color = kleuren[i])
    
    plt.xticks(randen[:-1], [functions.getaltotijd(rand) for rand in randen[:-1]])
    plt.setp(plt.gca().get_xticklabels(), rotation=30, horizontalalignment='right')
    plt.yticks(ywaardes, [functions.getaltogeld(ywaarde) for ywaarde in ywaardes])

    handles, labels = plt.gca().get_legend_handles_labels()
    order = range(blokaantal-1, -1, -1)
    plt.legend([handles[idx] for idx in order],[labels[idx] for idx in order]) 
    
    plt.gca().set_title("Totale partnerpensioen: €{:.2f}".format(PPtotaal).replace(".",","))
    plt.suptitle(invoer.range(titel).value, fontweight='bold')
    
    
    uitvoer.pictures.add(afbeelding, top = uitvoer.range((3,3)).top, left = uitvoer.range((3,3)).left, height = 300)
    
    




@xw.sub
#Idee voor berekeningen uitvoeren: Functies schrijven
def invoer_test_klikken():
    
    #sheets en book opslaan in variabelen
    book = xw.Book.caller()
    invoer = book.sheets["Tijdelijk invoerscherm"]
    sterftetafel= book.sheets["Sterftetafels"] 
    AG2020= book.sheets["AG2020, unisex"]
    pensioencontracten= book.sheets["Gegevens pensioencontracten"]
    
    
    
    pensioenbedragen=invoer.range((10,2),(18,2))
    sterftetafel_range= invoer.range((10,3),(18,3))
    rentes= invoer.range((10,4),(18,4))
    pensioenleeftijd_range= invoer.range((10,5),(18,5))
    koopsom_range= invoer.range((10,6),(18,6))
    
    
    #testing
    Aegon2011= sterftetafel.range((5,2),(123,2))
    
    # levend_pensioenleeftijd= Aegon2011(65).value
    

    #list met alle eigenschappen en vanuit daar rekenen
    
    
    pensioenleeftijd=[]
    rente=[]

    
    letters=[]
    for p in range(3):
        if len(letters)>=26:
            for i in ascii_uppercase:
                letters.append( letters[p] + i)
        else:
            for i in ascii_uppercase:
                letters.append(i)
    
    
    
    
    counter=1
    for i in range(1,10):
        print(i)
        if pensioenbedragen(i).value != None:
            kolom_t= (counter-1)*8+8
            kolom_leeftijd=(counter-1)*8+9           
            
            rente.append(rentes(i).value)
            pensioenleeftijd.append(pensioenleeftijd_range(i).value)
            
            invoer.range((2, kolom_t)).value= 0
            invoer.range((3, kolom_t), (61, kolom_t)).formula= [['=1+' + str(letters[kolom_t-1]) + '2']]
            
            invoer.range((2, kolom_leeftijd)).value= pensioenleeftijd_range(i).value
            invoer.range((3, kolom_leeftijd), (61, kolom_leeftijd)).formula= [['=1+' + str(letters[kolom_leeftijd-1]) + '2']]
            
            counter+= 1
            

                #koopsom_range(2).formula= [['=SUMPRODUCT(J2:J61,K2:K61)']]
        else:
            rente.append(0)
            pensioenleeftijd.append(0)
            
            
    #test voor 1 pensioenvorm
    # if pensioenbedragen(2).value != None:
    #     kolom_tpx.formula= 
    
    # test=invoer.range(2,12)
    # print(str(2))
    
    
    
    print(pensioenleeftijd)   
    # print(rente)
    #kolom_tqx.formula= [['=1-kolom_tpx(1).value']]

        
    #tqx_formula= [['=1-'+ str(letters[9])+str(2)]]
    
    # kolom_tpx.formula= [['=Sterftetafels!B' + str(int(pensioenleeftijd[1]+4)) + '/Sterftetafels!$B$' + str(int(pensioenleeftijd[1]+4))]]
    # kolom_tqx.formula= [['=1-' + letters[9] + '2']]
    
    p= pensioenleeftijd[1]
    x=1
    # while p<=119:
    #     # kolom_t(x).value=x-1
    #     # kolom_leeftijd(x).value=p
    #     kolom_tpx(x).value=(Aegon2011(p).value/Aegon2011(pensioenleeftijd[1]).value)
    #     kolom_tqx(x).value= 1-kolom_tpx(x).value
    #     x=x+1
    #     p=p+1
        
    

        
        
    
    