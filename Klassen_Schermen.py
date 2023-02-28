"""
Header
Hier komen alle libraries die in het programma gebruikt worden
"""
import sys
from PyQt5 import QtWidgets, uic

from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.backends.backend_qt5agg import NavigationToolbar2QT as NavigationToolbar
import matplotlib.pyplot as plt
import functions
from datetime import datetime
from functions import pensioensdatum, isfloat, ToevoegenDeelnemer, gegevenscontrole #deze zouden ook moeten inladen met de import functions hierboven, maar dat werkt niet


"""
Body
Hier komen alle functies
"""
class Functiekeus(QtWidgets.QMainWindow):
    def __init__(self, book, logger):
        self._logger = logger
        self._logger.info("Functiekeus scherm geopend")
        Ui_MainWindow, QtBaseClass = uic.loadUiType("{}\\1AdviseurBeheerder.ui".format(sys.path[0]))
        super(Functiekeus, self).__init__()
        self.book = book
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.ui.btnAdviseur.clicked.connect(self.btnAdviseurClicked)
        self.ui.btnBeheerder.clicked.connect(self.btnBeheerderClicked)
        
        
    def btnAdviseurClicked(self):
        self.close()
        self._logger.info("Functiekeus scherm gesloten")
        self._windowdeelnemer = Deelnemerselectie(self.book, self._logger)
        self._windowdeelnemer.show()
    def btnBeheerderClicked(self): 
        self.close()
        self._logger.info("Functiekeus scherm gesloten")
        self._windowinlog = Inloggen(self.book, self._logger)
        self._windowinlog.show()
        


class Inloggen(QtWidgets.QMainWindow):
    def __init__(self, book, logger):
        self._logger = logger
        self._logger.info("Inloggen scherm geopend")
        Ui_MainWindow2, QtBaseClass2 = uic.loadUiType("{}\\2InlogBeheerder.ui".format(sys.path[0]))
        super(Inloggen, self).__init__()
        self.book = book
        self.ui = Ui_MainWindow2()
        self.ui.setupUi(self)
        self.ui.btnTerug.clicked.connect(self.btnTerugClicked)
        self.ui.btnInloggen.clicked.connect(self.btnInloggenClicked)
        self._Wachtwoord = "wachtwoord"

        
    def btnInloggenClicked(self):
        if self.ui.txtBeheerderscode.text() == self._Wachtwoord:
            self._logger.info("Inloggen scherm gesloten")
            self.close()
            self._windowdeelnemer = Deelnemerselectie(self.book, self._logger)
            self._windowdeelnemer.show()
        else:
            self.ui.lblFoutmeldingInlog.setText("Wachtwoord incorrrect")
    def btnTerugClicked(self):
        self.close()
        self._logger.info("Inloggen scherm gesloten")
        self._windowkeus = Functiekeus(self.book, self._logger)
        self._windowkeus.show()
        
        

class Deelnemerselectie(QtWidgets.QMainWindow):
    def __init__(self, book, logger):
        self._logger = logger
        self._logger.info("Deelnemerselectie scherm geopend")
        Ui_MainWindow3, QtBaseClass3 = uic.loadUiType("{}\\deelnemerselectie.ui".format(sys.path[0]))
        super(Deelnemerselectie, self).__init__()
        self.book = book
        self.ui = Ui_MainWindow3()
        self.ui.setupUi(self)
        self.deelnemerlijst = functions.getDeelnemersbestand(self.book)
        self.kleinDeelnemerlijst = list()
        self.ui.btnDeelnemerToevoegen.clicked.connect(self.btnDeelnemerToevoegenClicked)
        self.ui.btnStartFlexibiliseren.clicked.connect(self.btnStartFlexibiliserenClicked)
        self.ui.btnTerug.clicked.connect(self.btnTerugClicked)
        self.ui.sbDag.valueChanged.connect(lambda: self.onChange(False))
        self.ui.sbMaand.valueChanged.connect(lambda: self.onChange(True))
        self.ui.sbJaar.valueChanged.connect(lambda: self.onChange(True))
        self.ui.txtVoorletters.textChanged.connect(lambda: self.onChange(False))
        self.ui.txtTussenvoegsel.textChanged.connect(lambda: self.onChange(False))
        self.ui.txtAchternaam.textChanged.connect(lambda: self.onChange(False))
        self.ui.cbGeslacht.currentTextChanged.connect(lambda: self.onChange(False))
        self.ui.lwKeuzes.currentItemChanged.connect(self.clearError)
        
        
    def btnDeelnemerToevoegenClicked(self):
        self.close()
        self._logger.info("Deelnemerselectie scherm gesloten")
        self._windowtoevoeg = Deelnemertoevoegen(self.book, self._logger)
        self._windowtoevoeg.show()
        
    def btnStartFlexibiliserenClicked(self):
        if self.ui.lwKeuzes.currentRow() == -1: 
            self.ui.lblFoutmeldingKiezen.setText("Gelieve een deelnemer te slecteren voordat u gaat flexibiliseren")
            return
        deelnemer = self.kleinDeelnemerlijst[self.ui.lwKeuzes.currentRow()]
        deelnemer.actieveerFlexibilisatie()
        self.close()
        self._logger.info("Deelnemerselectie scherm gesloten")
        self._windowflex = Flexmenu(self.book, deelnemer, self._logger)
        self._windowflex.show()
        
    def btnTerugClicked(self):
        self.close()
        self._logger.info("Deelnemerselectie scherm gesloten")
        self.windowstart = Functiekeus(self.book, self._logger)
        self.windowstart.show()
    
    def clearError(self): self.ui.lblFoutmeldingKiezen.clear()
        
    def onChange(self, datumChange):
        if datumChange: functions.maanddag(self)
        kleinDeelnemerlijst = self.deelnemerlijst
        kleinDeelnemerlijst = functions.filterkolom(kleinDeelnemerlijst, self.ui.txtVoorletters.text(), "voorletters")
        kleinDeelnemerlijst = functions.filterkolom(kleinDeelnemerlijst, self.ui.txtTussenvoegsel.text(), "tussenvoegsels")
        kleinDeelnemerlijst = functions.filterkolom(kleinDeelnemerlijst, self.ui.txtAchternaam.text(), "achternaam")
        kleinDeelnemerlijst = functions.filterkolom(kleinDeelnemerlijst, datetime(self.ui.sbJaar.value(), self.ui.sbMaand.value(), self.ui.sbDag.value()), "geboortedatum")
        kleinDeelnemerlijst = functions.filterkolom(kleinDeelnemerlijst, self.ui.cbGeslacht.currentText(), "geslacht")
        self.ui.lwKeuzes.clear()
        for deelnemer in kleinDeelnemerlijst[:10]:
            weergave = "{} {}".format(getattr(deelnemer, "voorletters"), getattr(deelnemer, "achternaam"))
            if getattr(deelnemer, "tussenvoegsels") != None: weergave += ", {}".format(getattr(deelnemer, "tussenvoegsels"))
            weergave += " | {} | {}".format(getattr(deelnemer, "geboortedatum").date(), getattr(deelnemer, "geslacht"))
            self.ui.lwKeuzes.addItem(weergave)
        self.kleinDeelnemerlijst = kleinDeelnemerlijst
        self.ui.lwKeuzes.repaint()
        
        
        
        
class Deelnemertoevoegen(QtWidgets.QMainWindow):
    def __init__(self, book, logger):
        self._logger = logger
        self._logger.info("Deelnemer toevoegen scherm geopend")
        Ui_MainWindow4, QtBaseClass4 = uic.loadUiType("{}\\4DeelnemerToevoegen.ui".format(sys.path[0]))
        super(Deelnemertoevoegen, self).__init__()
        self.book = book
        self.ui = Ui_MainWindow4()
        self.ui.setupUi(self)
        self.ui.btnTerug.clicked.connect(self.btnTerugClicked)
        self.ui.btnToevoegen.clicked.connect(self.btnToevoegenClicked)
        self.ui.sbMaand.valueChanged.connect(self.onChange)
        self.ui.sbJaar.valueChanged.connect(self.onChange)
        self._30maand = [4,6,9,11]
        #voeg schaduwtekst toe aan de invoervelden
        self.ui.txtVoorletters.setPlaceholderText("A.B.")
        self.ui.txtTussenvoegsel.setPlaceholderText("van")
        self.ui.txtAchternaam.setPlaceholderText("Albert")
        self.ui.txtParttimePercentage.setPlaceholderText("70")
        for i in [self.ui.txtFulltimeLoon, self.ui.txtOPAegon65, self.ui.txtOPAegon67, self.ui.txtOPNN65, self.ui.txtOPNN67, 
                  self.ui.txtOPVLC68, self.ui.txtOPZL, self.ui.txtPPNN65, self.ui.txtPPNN67, self.ui.txtPPVLC68]:
            i.setPlaceholderText("500,00")
        
    def btnTerugClicked(self):
        self.close()
        self._logger.info("Deelnemer toevoegen scherm gesloten")
        self._windowdeelnemer = Deelnemerselectie(self.book, self._logger)
        self._windowdeelnemer.show()
        
    def btnToevoegenClicked(self):
        #lege foutmeldingen aanmaken
        foutmeldingGegevens = ""
        foutmeldingPensioen = ""
        AantalPensioenen = 0 #teller voor aantal afgeronde pensioenopbouwen
        FouteRegelingen = [] #lijst met pensioenregelingen met foute invoer
        #controleer persoonsgegevens
        
        if self.ui.txtVoorletters.text() == "" or self.ui.txtAchternaam.text() == "":
            foutmeldingGegevens = foutmeldingGegevens + "Uw naam is niet ingevuld. "
        if isfloat(self.ui.txtFulltimeLoon.text()) == False or isfloat(self.ui.txtParttimePercentage.text()) == False or float(self.ui.txtParttimePercentage.text().replace(",", ".")) > 100:
            if len(foutmeldingGegevens) > 0: #De naam is ook niet goed ingevoerd
                foutmeldingGegevens = "Uw naam en werkinformatie zijn niet (goed) ingevuld. "
            else: 
                foutmeldingGegevens = "Uw werkinformatie is niet (goed) ingevuld. "
        #controleer of de deelnemer al de pensioenleeftijd heeft behaald
        if self.ui.sbJaar.text() < str(pensioensdatum())[3:7]:
            foutmeldingGegevens = foutmeldingGegevens + "U hebt de pensioensleeftijd al bereikt."
        elif self.ui.sbJaar.text() == str(pensioensdatum())[3:7] and int(self.ui.sbMaand.text()) < int(str(pensioensdatum())[0:2]):
            foutmeldingGegevens = foutmeldingGegevens + "U hebt de pensioensleeftijd al bereikt."
        
        #controleer pensioengegevens
        #lijst met pensioensgegevens [regeling, ZL, AegonOP65, AegonOP67, NNOP65, NNPP65,NNOP67, NNPP67, PFVLCOP68, PFVLCPP68]
        Pensioensgegevens = [self.ui.cbHuidigeRegeling.currentText(), "", "", "", "", "", "", "", "", ""]
        #lijst met invoervelden van het userform
        invoerPensioenen = [[self.ui.CheckZL, self.ui.txtOPZL, "ZL"], [self.ui.CheckAegon65, self.ui.txtOPAegon65, "Aegon65"], 
                            [self.ui.CheckAegon67, self.ui.txtOPAegon67, "Aegon67"], [self.ui.CheckNN65, self.ui.txtOPNN65, self.ui.txtPPNN65, "NN65"], 
                            [self.ui.CheckNN67, self.ui.txtOPNN67, self.ui.txtPPNN67, "NN67"], [self.ui.CheckPFVLC68, self.ui.txtOPVLC68, self.ui.txtPPVLC68, "VLC68"]]
        tellerPensioenen = 1    #zorgt dat juist pensioen op juiste plek in Pensioensgegevens komt
        
        for i in invoerPensioenen:
            if i[0].isChecked() == True:    #het pensioen is aangevinkt
                AantalPensioenen += 1       #houdt het aantal aangevinkte pensioenen bij
                for x in i[1:-1]:
                    if isfloat(x.text()) == True:   #Er is een getal-waarde ingevuld
                        Pensioensgegevens[tellerPensioenen] = float(x.text().replace(".", "").replace(",", "."))
                    else:
                        FouteRegelingen.append(i[-1])   #regeling aan foutmelding toevoegen
                    tellerPensioenen += 1
            else:
                for x in i[1:-1]:
                    if isfloat(x.text()) == True:   #wel een getal-waarde ingevuld, maar pensioen niet aangevinkt
                        FouteRegelingen.append(i[-1])
                tellerPensioenen += len(i)-2        #tellerPensioenen ophogen met aantal pensioenopties OP of OP+PP
        #foutmelding pensioensgegevens genereren
        if AantalPensioenen == 0 and len(FouteRegelingen) == 0: #foutmelding als er geen regeling aangegeven is
            foutmeldingPensioen = "U heeft nog geen opgebouwd pensioen aangegeven"
        elif len(FouteRegelingen) > 0: #foutmelding als regelingen niet volledig of fout zijn ingevuld
            foutmeldingPensioen = "De volgende regelingen zijn niet (goed) ingevoerd: " + FouteRegelingen[0]
            for i in FouteRegelingen[1:]:
                foutmeldingPensioen = foutmeldingPensioen + ", " + i        
           
        #gegevens invullen of foutmelding geven
        if foutmeldingGegevens == "" and foutmeldingPensioen == "":
            geboortedatum = self.ui.sbDag.text() + "-" + self.ui.sbMaand.text() + "-" + self.ui.sbJaar.text() 
            achternaam = self.ui.txtAchternaam.text()[0].upper() + self.ui.txtAchternaam.text()[1:]
            #voorletters met hoofdletters en punten ertussen
            voorletters = ""
            for i in self.ui.txtVoorletters.text().replace(".", "").upper():
                voorletters += i + "."
            #fulltime loon en parttime percentage als float
            fulltimeLoon = float(self.ui.txtFulltimeLoon.text().replace(".", "").replace(",", "."))
            ptPercentage = float(self.ui.txtParttimePercentage.text().replace(",", "."))/10000    #delen door 100, zodat het in excel als % komt
            #lijst met deelnemersgegevens [achternaam, tussenvoegsel, voorletters, geboortedatum, geslacht, burg.staat, ftloon, pt%]
            Deelnemersgegevens = [achternaam, self.ui.txtTussenvoegsel.text(), voorletters, geboortedatum, self.ui.cbGeslacht.currentText(), 
                                  self.ui.cbBurgerlijkeStaat.currentText(), fulltimeLoon, ptPercentage]
            #lijst met alle gegevens
            gegevens = Deelnemersgegevens + Pensioensgegevens
            
            #deelnemer zijn gegevens laten controleren
            controle = gegevenscontrole(gegevens)
            if controle == "correct":
                #window sluiten
                self.close()
                self._logger.info("Deelnemer toevoegen scherm gesloten")
                #toevoegen van de gegevens van een deelnemer aan het deelnemersbestand
                ToevoegenDeelnemer(gegevens)
                #deelnemerselectie openen
                self._windowdeelnemer = Deelnemerselectie(self.book, self._logger)
                self._windowdeelnemer.show()
            elif controle == "fout":
                self.ui.lblFoutmeldingGegevens.setText("Pas uw gegevens aan en druk weer op Deelnemer toevoegen")
                #als niet op "ja" wordt geklikt, wordt de messagebox gesloten en het invoerveld weer getoont
            
        else: 
            #foutmelding tonen
            self.ui.lblFoutmeldingGegevens.setText(foutmeldingGegevens)
            self.ui.lblFoutmeldingPensioen.setText(foutmeldingPensioen)
        
        
        
        #if self.ui.txtVoorletters.text() == "" or self.ui.txtAchternaam.text() == "":
        #    print("Naam gegevens incompleet")
        #elif self.ui.txtFulltimeLoon.text() == "" or self.ui.txtParttimePercentage.text() == "":
        #    print("Loon informatie incompleet")
        #else:
         #   self.close()
          #  self._windowdeelnemer = Deelnemerselectie(self.book)
           # self._windowdeelnemer.show()
    
    
    def onChange(self): functions.maanddag(self)



class Flexmenu(QtWidgets.QMainWindow):
    
    def __init__(self, book, deelnemer, logger):
        self._logger = logger
        self._logger.info("Flexmenu scherm geopend")
        Ui_MainWindow5, QtBaseClass5 = uic.loadUiType("{}\\flexmenu.ui".format(sys.path[0]))
        super(Flexmenu, self).__init__()
        self.book = book
        
        # Setup van UI
        self.ui = Ui_MainWindow5()
        self.ui.setupUi(self)
        
        # Deelnemer
        self.deelnemerObject = deelnemer
        
        # Regeling selectie
        self.dropdownRegelingen()
    
        # Knoppen
        self.ui.btnAndereDeelnemer.clicked.connect(self.btnAndereDeelnemerClicked)
        self.ui.btnVergelijken.clicked.connect(self.btnVergelijkenClicked)
        self.ui.btnOpslaan.clicked.connect(self.btnOpslaanClicked)
        
        # Aanpassing: pensioenleeftijd
        self.ui.sbJaar.valueChanged.connect(self.invoerVerandering)
        self.ui.sbMaand.valueChanged.connect(self.invoerVerandering)
        
        # Aanpassing: OP/PP
        self.ui.cbUitruilenVan.activated.connect(self.invoerVerandering)
        self.ui.cbUMethode.activated.connect(self.invoerVerandering)
        self.ui.txtUVerhoudingOP.textEdited.connect(self.invoerVerandering)
        self.ui.txtUVerhoudingPP.textEdited.connect(self.invoerVerandering)
        self.ui.txtUPercentage.textEdited.connect(self.invoerVerandering)
           
        # Aanpassing: hoog-laag constructie
        self.ui.cbHLVolgorde.activated.connect(self.invoerVerandering)
        self.ui.cbHLMethode.activated.connect(self.invoerVerandering)
        self.ui.txtHLVerhoudingHoog.textEdited.connect(self.invoerVerandering)
        self.ui.txtHLVerhoudingLaag.textEdited.connect(self.invoerVerandering)
        self.ui.txtHLVerschil.textEdited.connect(self.invoerVerandering)
    
    def dropdownRegelingen(self):
        regelingenActief = list()
        for regeling in self.deelnemerObject._pensioenen:
            if regeling.ouderdomsPensioen != None:
                if regeling.ouderdomsPensioen > 0:
                    regelingenActief.append(regeling.pensioenVolNaam)
                    
        self.ui.cbRegeling.addItems(regelingenActief)
        
    def regelingenObject(self):
        """
        Deze functie maakt voor elke regeling een flexibilisering-object aan 
        uit flex_keuzes.py. Functie checkt ook welke regelingen actief zijn. 
        """
    
        """ 
        Onderstaande code is verouderd! De "self.deelnemer" variabel was de rij waar
        de deelnemer in het Excel deelnemersbestand stond. Deze variabel is nu verouderd 
        en er gaat nog uitgezocht worden hoe dit aangepast moet worden.
        
        if "ZL" in functions.regelingenophalen(self.deelnemer)[1]:
            self._ZL = flexibilisering("ZL",True)
        else:
            self._ZL = flexibilisering("ZL",False)
            
        if "Aegon65" in functions.regelingenophalen(self.deelnemer)[1]:
            self._A65 = flexibilisering("Aegon65",True)
        else:
            self._A65 = flexibilisering("Aegon65",True)
            
        if "Aegon67" in functions.regelingenophalen(self.deelnemer)[1]:
            self._A67 = flexibilisering("Aegon67",True)
        else:
            self._A67 = flexibilisering("Aegon67",True)
            
        if "NN65" in functions.regelingenophalen(self.deelnemer)[1]:
            self._NN65 = flexibilisering("NN65",True)
        else:
            self._NN65 = flexibilisering("NN65",True)
            
        if "NN67" in functions.regelingenophalen(self.deelnemer)[1]:
            self._NN67 = flexibilisering("NN67",True)
        else:
            self._NN67 = flexibilisering("NN67",True)
            
        if "PF_VLC68" in functions.regelingenophalen(self.deelnemer)[1]:
            self._VLC68 = flexibilisering("PF_VLC68",True)
        else:
            self._VLC68 = flexibilisering("PF_LC68",False)
        """
        
    def invoerVerandering(self):
        self.regelingCode = functions.regelingNaamCode(str(self.ui.cbRegeling.currentText()))
        
        if self.ui.CheckLeeftijdWijzigen.isChecked() == True:
            # Sla de nieuwe leeftijd op
            pass
        else:
            # Sla de oude leeftijd op
            pass
        
        if self.ui.CheckUitruilen.isChecked() == True:
            # Sla de OP/PP flexibiliseringen op
            pass
        else:
            # Sla op dat er geen flexibiliseringen zijn
            pass
        
        if self.ui.CheckHoogLaag.isChecked() ==  True:
            # Sla de hoog-laag flexibiliseringen op
            pass
        else:
            # Sla op dat er geen flexibiliseringen zijn
            pass

    def btnAndereDeelnemerClicked(self):
        self.close()
        self._logger.info("Flexmenu scherm gesloten")
        self._windowdeelnemer = Deelnemerselectie(self.book, self._logger)
        self._windowdeelnemer.show()
        
    def btnVergelijkenClicked(self):
        # Sheet van vergelijkingen openen
        self.close()
        self._logger.info("Flexmenu scherm gesloten")
        
    def btnOpslaanClicked(self): 
        # Alle huidige flexibiliserignen opslaan in een Excel sheet
        # Huidig diagram opslaan en plaats in vergelijking sheet
        self.close()
        self._logger.info("Flexmenu scherm gesloten")