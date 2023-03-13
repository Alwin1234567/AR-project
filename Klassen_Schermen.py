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
from decimal import getcontext, Decimal
from xlwings.utils import rgb_to_int

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
        #self.book.app.display_alerts = False # Dit moet de OLE melding voorkomen
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.setWindowTitle("Kies uw functie")
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
        self.setWindowTitle("Inloggen als beheerder")
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
        self.setWindowTitle("Deelnemerselectie")
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
            self.ui.lblFoutmeldingKiezen.setText("Gelieve een deelnemer te selecteren voordat u gaat flexibiliseren")
            return
        deelnemer = self.kleinDeelnemerlijst[self.ui.lwKeuzes.currentRow()]
        deelnemer.activeerFlexibilisatie()
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
        self.setWindowTitle("Deelnemer toevoegen")
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
        if self.ui.cbHuidigeRegeling.currentText() != "Inactief":
            if functions.isfloat(self.ui.txtFulltimeLoon.text()) == False or functions.isfloat(self.ui.txtParttimePercentage.text()) == False or float(self.ui.txtParttimePercentage.text().replace(",", ".")) > 100:
                if len(foutmeldingGegevens) > 0: #De naam is ook niet goed ingevoerd
                    foutmeldingGegevens = "Uw naam en werkinformatie zijn niet (goed) ingevuld. "
                else: 
                    foutmeldingGegevens = "Uw werkinformatie is niet (goed) ingevuld. "
                
        #controleer of de deelnemer al de pensioenleeftijd heeft behaald
        if self.ui.sbJaar.text() < str(functions.pensioensdatum())[3:7]:
            foutmeldingGegevens = foutmeldingGegevens + "U hebt de pensioensleeftijd al bereikt."
        elif self.ui.sbJaar.text() == str(functions.pensioensdatum())[3:7] and int(self.ui.sbMaand.text()) < int(str(functions.pensioensdatum())[0:2]):
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
                    if functions.isfloat(x.text()) == True:   #Er is een getal-waarde ingevuld
                        Pensioensgegevens[tellerPensioenen] = float(x.text().replace(".", "").replace(",", "."))
                    else:
                        FouteRegelingen.append(i[-1])   #regeling aan foutmelding toevoegen
                    tellerPensioenen += 1
            else:
                for x in i[1:-1]:
                    if functions.isfloat(x.text()) == True:   #wel een getal-waarde ingevuld, maar pensioen niet aangevinkt
                        FouteRegelingen.append(i[-1])
                tellerPensioenen += len(i)-2        #tellerPensioenen ophogen met aantal pensioenopties OP of OP+PP
        
        #controleren of OP:PP verhouding kleiner is dan 100:70
        verhoudingFout = ""
        PPRegelingen = ["NN65", "NN67", "VLC"]  #pensioenen met PP
        for i in range(4,10,2):
            #kijken of pensioen ingevuld is.
            if Pensioensgegevens[i] != "":
                verhouding = (Pensioensgegevens[i+1])/(Pensioensgegevens[i])
                if verhouding > 0.7:
                    #als verhouding groter dan 100:70, pensioen toevoegen aan foutmelding
                    verhoudingFout = verhoudingFout + ", " + PPRegelingen[int(i/2 - 2)]
        
        #foutmelding pensioensgegevens genereren
        if AantalPensioenen == 0 and len(FouteRegelingen) == 0: #foutmelding als er geen regeling aangegeven is
            foutmeldingPensioen = "U heeft nog geen opgebouwd pensioen aangegeven"
        elif len(FouteRegelingen) > 0: #foutmelding als regelingen niet volledig of fout zijn ingevuld
            foutmeldingPensioen = "De volgende regelingen zijn niet (goed) ingevoerd: " + FouteRegelingen[0]
            for i in FouteRegelingen[1:]:
                foutmeldingPensioen = foutmeldingPensioen + ", " + i        
        if len(verhoudingFout) > 0:
            foutmeldingPensioen = foutmeldingPensioen + " De verhouding OP:PP mag niet groter zijn dan 100:70:" + verhoudingFout[1:]
        
        
        #gegevens invullen of foutmelding geven
        if foutmeldingGegevens == "" and foutmeldingPensioen == "":
            geboortedatum = datetime(int(self.ui.sbJaar.text()), int(self.ui.sbMaand.text()), int(self.ui.sbDag.text()))
            achternaam = self.ui.txtAchternaam.text()[0].upper() + self.ui.txtAchternaam.text()[1:]
            #voorletters met hoofdletters en punten ertussen
            voorletters = ""
            for i in self.ui.txtVoorletters.text().replace(".", "").upper():
                voorletters += i + "."
            if self.ui.cbHuidigeRegeling.currentText() != "Inactief":
                #fulltime loon en parttime percentage als float
                fulltimeLoon = float(self.ui.txtFulltimeLoon.text().replace(".", "").replace(",", "."))
                getcontext().prec = 7
                ptPercentage = Decimal(self.ui.txtParttimePercentage.text().replace(",", "."))
            else:
                fulltimeLoon = ""
                ptPercentage = ""
            #lijst met deelnemersgegevens [achternaam, tussenvoegsel, voorletters, geboortedatum, geslacht, burg.staat, ftloon, pt%]
            Deelnemersgegevens = [achternaam, self.ui.txtTussenvoegsel.text(), voorletters, geboortedatum, self.ui.cbGeslacht.currentText(), 
                                  self.ui.cbBurgerlijkeStaat.currentText(), fulltimeLoon, ptPercentage]
                        
            #controleren of deelnemer al bestaat in deelnemersbestand
            deelnemerDubbel = functions.DeelnemerVinden(self.book, Deelnemersgegevens)
            
            #geboortedatum in goede notatie voor invoer in excel
            geboortedatum = datetime(int(self.ui.sbJaar.text()), int(self.ui.sbMaand.text()), int(self.ui.sbDag.text())).strftime("%m-%d-%Y")
            Deelnemersgegevens[3] = geboortedatum
            #lijst met alle gegevens
            gegevens = Deelnemersgegevens + Pensioensgegevens
            
            if len(deelnemerDubbel) == 0:       #deelnemer is nog niet bekend in deelnemersbestand
                #deelnemer zijn gegevens laten controleren
                self._logger.info("Ingevulde gegevens worden getoont voor controle")
                controle = functions.gegevenscontrole(gegevens)
                if controle == "correct":
                    #window sluiten
                    self.close()
                    self._logger.info("Deelnemer toevoegen scherm gesloten")
                    
                    #het parttime percentage delen door 100, zodat het in excel als % komt
                    if gegevens[7] != "": 
                        gegevens[7] = float(gegevens[7])/100
                    try: #toevoegen van de gegevens van een deelnemer aan het deelnemersbestand
                        functions.ToevoegenDeelnemer(gegevens)
                        self._logger.info("Nieuwe deelnemer is toegevoegd aan het deelnemersbestand")
                    except Exception as e:
                        self._logger.exception("Er is iets fout gegaan bij het toevoegen van een deelnemer aan het deelnemersbestand")
                    
                    #deelnemerselectie openen
                    self._windowdeelnemer = Deelnemerselectie(self.book, self._logger)
                    self._windowdeelnemer.show()
                elif controle == "fout":
                    self._logger.info("Deelnemer wil zijn ingevulde gegevens aanpassen. Deelnemer toevoegen scherm blijft open")
                    #als niet op "ja" wordt geklikt, wordt de messagebox gesloten en het invoerveld weer getoont
            else:
                #melding dat deelnemer al bekend is in deelnemersbestand
                titel = "Deelnemer al bekend"
                tekst = "De deelnemer die u wil toevoegen is al bekend in het deelnemersbestand.\n U kunt contact opnemen met de beheerder om de gegevens van deze deelnemer aan te passen."
                functions.Mbox(titel, tekst, 0)    #messagebox met alleen OK knop
                
        else: 
            self._logger.info("Niet alle deelnemersgegevens zijn goed ingevuld. De deelnemer moet zijn gegevens aanpassen")
            #foutmelding tonen
            self.ui.lblFoutmeldingGegevens.setText(foutmeldingGegevens)
            self.ui.lblFoutmeldingPensioen.setText(foutmeldingPensioen)
        
    
    
    def onChange(self): functions.maanddag(self)



class Flexmenu(QtWidgets.QMainWindow):
    
    def __init__(self, book, deelnemer, logger):
        self._logger = logger
        self._logger.info("Flexmenu scherm geopend")
        Ui_MainWindow5, QtBaseClass5 = uic.loadUiType("{}\\flexmenu.ui".format(sys.path[0]))
        super(Flexmenu, self).__init__()
        self.book = book
        self.opslaanCount = 0 #Teller voor aantal opgeslagen flexibilisaties.
        self.opslaanList = list()
        self.zoekFlexibilisaties()
        
        # Setup van UI
        self.ui = Ui_MainWindow5()
        self.ui.setupUi(self)
        self.setWindowTitle("Flexibilisatie menu") #Het moet na de setup, daarom staat het nu even hier
        
        # Deelnemer
        self.deelnemerObject = deelnemer
        
        # Regeling selectie
        self._regelingenActiefKort = list()
        self.dropdownRegelingen()
        
        # Knoppen
        self.ui.btnAndereDeelnemer.clicked.connect(self.btnAndereDeelnemerClicked)
        self.ui.btnVergelijken.clicked.connect(self.btnVergelijkenClicked)
        self.ui.btnOpslaan.clicked.connect(self.btnOpslaanClicked)
        
        # Aanpassing: pensioenleeftijd
        self.ui.CheckLeeftijdWijzigen.stateChanged.connect(self.invoerVerandering)
        self.ui.sbJaar.valueChanged.connect(self.invoerVerandering)
        self.ui.sbMaand.valueChanged.connect(self.invoerVerandering)
        
        # Aanpassing: OP/PP
        self.ui.CheckUitruilen.stateChanged.connect(self.invoerVerandering)
        self.ui.cbUitruilenVan.activated.connect(self.invoerVerandering)
        self.ui.cbUMethode.activated.connect(self.invoerVerandering)
        self.ui.txtUVerhoudingOP.textEdited.connect(self.invoerVerandering)
        self.ui.txtUVerhoudingPP.textEdited.connect(self.invoerVerandering)
        self.ui.txtUPercentage.textEdited.connect(self.invoerVerandering)
           
        # Aanpassing: hoog-laag constructie
        self.ui.CheckHoogLaag.stateChanged.connect(self.invoerVerandering)
        self.ui.cbHLVolgorde.activated.connect(self.invoerVerandering)
        self.ui.cbHLMethode.activated.connect(self.invoerVerandering)
        self.ui.txtHLVerhoudingHoog.textEdited.connect(self.invoerVerandering)
        self.ui.txtHLVerhoudingLaag.textEdited.connect(self.invoerVerandering)
        self.ui.txtHLVerschil.textEdited.connect(self.invoerVerandering)
        self.ui.sbHLJaar.valueChanged.connect(self.invoerVerandering)

        # Aanpassing: Regeling
        self.ui.cbRegeling.activated.connect(self.wijzigVelden)
        
        # Laatste UI update
        self.ui.sbMaand.setValue(0)
        self.invoerVerandering()
        self.wijzigVelden()
        
        # Berekening sheet klaarmaken
        functions.berekeningen_init(book.sheets["Berekeningen"], self.deelnemerObject, self._logger)
    
    def zoekFlexibilisaties(self):
        self.opslaanList = self.book.sheets["Flexopslag"].range("3:3")[1:500].value # Zoekt maximaal tot 500 anders wordt het langzaam
        self.opslaanList = [int(value) for value in self.opslaanList if type(value) == float]

    def zoekNieuwID(self):
        if len(self.opslaanList) > 0:
            for i in range(len(self.opslaanList)):
                if (i+1) not in self.opslaanList:
                    return (i+1)
            return len(self.opslaanList)+1
        else:
            return 1
    
    def dropdownRegelingen(self):
        regelingenActief = list()
        regelingenActiefKort = list()
        
        for regeling in self.deelnemerObject.pensioenen:
            if regeling.ouderdomsPensioen != None:
                if regeling.ouderdomsPensioen > 0:
                    regelingenActief.append(regeling.pensioenVolNaam)
                    regelingenActiefKort.append(regeling.pensioenNaam)

        self.ui.cbRegeling.addItems(regelingenActief)
        self._regelingenActiefKort = regelingenActiefKort
    
    def blokkeerSignalen(self, actief):
        """
        Een verandering in een invoervakje geeft een 'signaal'.
        Deze functie kan de 'signalen' van invoervakjes blokkeren en activeren.
        
        actief : bool
            Als True, dan zullen wijzigingen in de window NIET geregistreerd worden.
            Als False, dan zullen wijzigingen in de window WEL geregistreerd worden.
        """
        
        if actief == True:
            self._logger.info("Veld signalen worden geblokkerd.")
        elif actief == False:
            self._logger.info("Veld signalen worden geactiveerd.")
        
        # Aanpassing: pensioenleeftijd
        self.ui.CheckLeeftijdWijzigen.blockSignals(actief)
        self.ui.sbJaar.blockSignals(actief)
        self.ui.sbMaand.blockSignals(actief)
 
        # Aanpassing: OP/PP
        self.ui.CheckUitruilen.blockSignals(actief)
        self.ui.cbUitruilenVan.blockSignals(actief)
        self.ui.cbUMethode.blockSignals(actief)
        self.ui.txtUVerhoudingOP.blockSignals(actief)
        self.ui.txtUVerhoudingPP.blockSignals(actief)
        self.ui.txtUPercentage.blockSignals(actief)
           
        # Aanpassing: hoog-laag constructie
        self.ui.CheckHoogLaag.blockSignals(actief)
        self.ui.cbHLVolgorde.blockSignals(actief)
        self.ui.cbHLMethode.blockSignals(actief)
        self.ui.txtHLVerhoudingHoog.blockSignals(actief)
        self.ui.txtHLVerhoudingLaag.blockSignals(actief)
        self.ui.txtHLVerschil.blockSignals(actief)
        self.ui.sbHLJaar.blockSignals(actief)
    
    def invoerCheck(self):
        """
        Deze functie checkt voor de volgende velden of de invoer klopt:
            - OP verhouding
            - PP verhouding
            - OP/PP uitruil percentage
            - Hoog verhouding
            - Laag verhouding
            - Hoog/laag verschil
        
        Er wordt voor al deze velden gecheckt of er letters staan.
        Er wordt alleen voor de relevante velden voor de methode gecheckt of er missende invoer is.
        
        Bij missende invoer en/of letters returnt deze functie False.
        Als alles klopt, returnt deze functie True.
        """
        
        melding_OP, OK_OP = functions.checkVeldInvoer(self.ui.cbUMethode.currentText(),
                                  self.ui.txtUPercentage.text(),
                                  self.ui.txtUVerhoudingOP.text(),
                                  self.ui.txtUVerhoudingPP.text())
        
        self.ui.lblFoutmeldingUitruilen.setText(melding_OP)
        
        melding_HL, OK_HL = functions.checkVeldInvoer(self.ui.cbHLMethode.currentText(),
                                                      self.ui.txtHLVerschil.text(),
                                                      self.ui.txtHLVerhoudingHoog.text(),
                                                      self.ui.txtHLVerhoudingLaag.text())
        
        self.ui.lblFoutmeldingHoogLaag.setText(melding_HL)

        if (OK_OP == True and OK_HL == True):
            return True
        else:
            return False

    def wijzigVelden(self):
        """
        Deze functie wordt geactiveerd als de regeling in de dropdown aangepast wordt.
        Deze functie moet alle invoervelden aanpassen naar eerder ingevoerde keuzes voor gekozen regeling.
        Als een regeling geselecteerd wordt waar nog niet eerder aanpassingen voor gemaakt zijn dan moet 
        deze functie alle velden weer leegmaken.
        """
        
        self.blokkeerSignalen(True)

        self._logger.info("Veldwijziging geïnitialiseerd.")
        
        for flexibilisatie in self.deelnemerObject.flexibilisaties:
            if flexibilisatie.pensioen.pensioenVolNaam == str(self.ui.cbRegeling.currentText()):
                self.regelingCode = flexibilisatie
                break

        # --- Leeftijd velden ---
        self._logger.info("Leeftijdvelden wijzigen...")
        try:
            self.ui.CheckLeeftijdWijzigen.setChecked(self.regelingCode.leeftijd_Actief)
            self.ui.sbJaar.setValue(int(self.regelingCode.leeftijdJaar))
            self.ui.sbMaand.setValue(int(self.regelingCode.leeftijdMaand))
            
        except Exception as e:
            self._logger.exception("Probleem bij het wijzigen van leeftijdvelden.")

        # --- OP/PP velden ---
        self._logger.info("OP/PP velden wijzigen...")
        try:
            self.ui.CheckUitruilen.setChecked(self.regelingCode.OP_PP_Actief)
            
            if self.regelingCode.OP_PP_UitruilenVan == "OP naar PP":
                self.ui.cbUitruilenVan.setCurrentIndex(0)
            elif self.regelingCode.OP_PP_UitruilenVan == "PP naar OP":
                self.ui.cbUitruilenVan.setCurrentIndex(1)
            
            if self.regelingCode.OP_PP_Methode == "Percentage":
                self.ui.cbUMethode.setCurrentIndex(0)
            elif self.regelingCode.OP_PP_Methode == "Verhouding":
                self.ui.cbUMethode.setCurrentIndex(1)
            
            self.ui.txtUVerhoudingOP.setText(str(self.regelingCode.OP_PP_Verhouding_OP))
            self.ui.txtUVerhoudingPP.setText(str(self.regelingCode.OP_PP_Verhouding_PP))
            self.ui.txtUPercentage.setText(str(self.regelingCode.OP_PP_Percentage))
            
        except Exception as e:
            self._logger.exception("Probleem bij het wijzigen van OP/PP velden.")
        
        # --- Hoog/Laag velden ---
        self._logger.info("Hoog/laag velden wijzigen...")
        try:
            self.ui.CheckHoogLaag.setChecked(self.regelingCode.HL_Actief)
            if self.regelingCode.HL_Volgorde == "Hoog-laag":
                self.ui.cbHLVolgorde.setCurrentIndex(0)
            elif self.regelingCode.HL_Volgorde == "Laag-hoog":
                self.ui.cbHLVolgorde.setCurrentIndex(1)
            
            if self.regelingCode.HL_Methode == "Opvullen AOW":
                self.ui.cbHLMethode.setCurrentIndex(0)
            elif self.regelingCode.HL_Methode == "Verhouding":
                self.ui.cbHLMethode.setCurrentIndex(1)
            elif self.regelingCode.HL_Methode == "Verschil":
                self.ui.cbHLMethode.setCurrentIndex(2)
    
            self.ui.txtHLVerhoudingHoog.setText(str(self.regelingCode.HL_Verhouding_Hoog))
            self.ui.txtHLVerhoudingLaag.setText(str(self.regelingCode.HL_Verhouding_Laag))
            self.ui.txtHLVerschil.setText(str(self.regelingCode.HL_Verschil))
            self.ui.sbHLJaar.setValue(int(self.regelingCode.HL_Jaar))
            
        except Exception as e:
            self._logger.exception("Probleem bij het wijzigen van hoog/laag velden.")
        
        self._logger.info("Veldwijziging afgerond.")
        
        self.blokkeerSignalen(False)
        
    def invoerVerandering(self):
        """ 
        Deze functie activeert zodra de gebruiker een verandering maakt in het flexmenu scherm.
        Zo kan het scherm live aanpassen op basis van input van de gebruiker.
        """
        # Functie voor invoer check
        #  > Is alles wat ingevoerd wel correct? (dus geen letters waar cijfers horen enzo)
        #  > Is alles ingevoerd waar invoer moet staan?
        # Als beide eisen voldoen, kunnen de volgende functies doorgevoerd worden
      
        if self.invoerCheck() == True:
            self.ui.lbl_opslaanMelding.setText("") # Opslaan melding verdwijnt.
            self.flexkeuzesOpslaan() # Sla flex keuzes op
            self.samenvattingUpdate() # Update de samenvatting
        
    def flexkeuzesOpslaan(self):
        """
        Deze functie slaat huidig ingevulde flex opties op in het flexibiliseringsobject.
        """
        
        # self.regelingCode = functions.regelingNaamCode(str(self.ui.cbRegeling.currentText()))
        
        self._logger.info("Flexkeuze opslaan geïnitialiseerd.")

        # Selecteer huidige regeling-object voor flex keuzes
        for flexibilisatie in self.deelnemerObject.flexibilisaties:
            if flexibilisatie.pensioen.pensioenVolNaam == str(self.ui.cbRegeling.currentText()):
                self.regelingCode = flexibilisatie
                break
        
        # --- Leeftijd wijzigen ---
        try:
            self.regelingCode.leeftijd_Actief = self.ui.CheckLeeftijdWijzigen.isChecked()
            self.regelingCode.leeftijdJaar = int(self.ui.sbJaar.value())
            self.regelingCode.leeftijdMaand = int(self.ui.sbMaand.value())
        except Exception as e:
            self._logger.exception("Er gaat iets fout bij het opslaan van de pensioenleeftijd in flexmenu.ui")
        
        # --- OP/PP uitruiling ---
        try:
            self.regelingCode.OP_PP_Actief = self.ui.CheckUitruilen.isChecked() 
            self.regelingCode.OP_PP_UitruilenVan = str(self.ui.cbUitruilenVan.currentText()) 
            self.regelingCode.OP_PP_Methode = str(self.ui.cbUMethode.currentText()) 
            
            if str(self.ui.cbUMethode.currentText()) == "Verhouding":
                self.regelingCode.OP_PP_Verhouding_OP = int(self.ui.txtUVerhoudingOP.text())
                self.regelingCode.OP_PP_Verhouding_PP = int(self.ui.txtUVerhoudingPP.text())
                
                if str(self.ui.txtUPercentage.text()) == "":
                    self.regelingCode.OP_PP_Percentage = 0
                else:
                    self.regelingCode.OP_PP_Percentage = int(self.ui.txtUPercentage.text())
            
            elif str(self.ui.cbUMethode.currentText()) == "Percentage":
                self.regelingCode.OP_PP_Percentage = int(self.ui.txtUPercentage.text())
                
                if str(self.ui.txtUVerhoudingOP.text()) == "":
                    self.regelingCode.OP_PP_Verhouding_OP = 0
                else:
                    self.regelingCode.OP_PP_Verhouding_OP = int(self.ui.txtUVerhoudingOP.text())
                
                if str(self.ui.txtUVerhoudingPP.text()) == "":
                    self.regelingCode.OP_PP_Verhouding_PP = 0
                else:
                    self.regelingCode.OP_PP_Verhouding_PP = int(self.ui.txtUVerhoudingPP.text())

        except Exception as e:
            self._logger.exception("Huidig geselecteerde OP/PP flexibilisaties in flexmenu.ui kunnen niet opgeslagen worden.")
        
        # --- Hoog/laag constructie ---
        try:
            self.regelingCode.HL_Actief = self.ui.CheckHoogLaag.isChecked() 
            self.regelingCode.HL_Volgorde = str(self.ui.cbHLVolgorde.currentText()) 
            self.regelingCode.HL_Methode = str(self.ui.cbHLMethode.currentText()) 
            self.regelingCode.HL_Jaar = int(self.ui.sbHLJaar.value()) 
            
            if str(self.ui.cbHLMethode.currentText()) == "Verhouding":
                self.regelingCode.HL_Verhouding_Hoog = int(self.ui.txtHLVerhoudingHoog.text())
                self.regelingCode.HL_Verhouding_Laag = int(self.ui.txtHLVerhoudingLaag.text())
                
                if str(self.ui.txtHLVerhoudingHoog.text()) == "":
                    self.regelingCode.HL_Verhouding_Hoog = 0
                else:
                    self.regelingCode.HL_Verschil = int(self.ui.txtHLVerschil.text())
                
            elif str(self.ui.cbHLMethode.currentText()) == "Verschil":
                self.regelingCode.HL_Verschil = int(self.ui.txtHLVerschil.text())
                
                if str(self.ui.txtHLVerhoudingHoog.text()) == "":
                    self.regelingCode.HL_Verhouding_Hoog = 0
                else:
                    self.regelingCode.HL_Verhouding_Hoog = int(self.ui.txtHLVerhoudingHoog.text())
                
                if str(self.ui.txtHLVerhoudingLaag.text()) == "":
                    self.regelingCode.HL_Verhouding_Laag = 0
                else:
                    self.regelingCode.HL_Verhouding_Laag = int(self.ui.txtHLVerhoudingLaag.text())
                    
            elif str(self.ui.cbHLMethode.currentText()) == "Opvullen AOW":
                if str(self.ui.txtHLVerhoudingHoog.text()) == "":
                    self.regelingCode.HL_Verhouding_Hoog = 0
                else:
                    self.regelingCode.HL_Verschil = int(self.ui.txtHLVerschil.text())
                    
                if str(self.ui.txtHLVerhoudingHoog.text()) == "":
                    self.regelingCode.HL_Verhouding_Hoog = 0
                else:
                    self.regelingCode.HL_Verhouding_Hoog = int(self.ui.txtHLVerhoudingHoog.text())
                
                if str(self.ui.txtHLVerhoudingLaag.text()) == "":
                    self.regelingCode.HL_Verhouding_Laag = 0
                else:
                    self.regelingCode.HL_Verhouding_Laag = int(self.ui.txtHLVerhoudingLaag.text())
                 
        except Exception as e:
            self._logger.exception("Huidig geselecteerde hoog/laag flexibilisaties in flexmenu.ui kunnen niet opgeslagen worden.")

    def berekenen(self):
        """
        Deze functie zal alle flexibiliseringswaarden naar de Excel sheet plaatsen voor berekenen.
        """
        pass
    
    def diagramUpdate(self):
        """
        Deze functie update het diagram.
        """
        pass
    
    def samenvattingUpdate(self):
        """
        Deze functie update de waarden in de samenvatting boxes.
        """
        
        self._logger.info("Samenvatting updaten...")
        
        # ZwitserLeven
        if "ZL" in self._regelingenActiefKort:
            self.regelingCode = next(flexibilisatie for flexibilisatie in self.deelnemerObject.flexibilisaties if flexibilisatie.pensioen.pensioenNaam == "ZL")
            
            self.ui.lbl_ZL.setText("ZL")
            
            if (self.regelingCode.leeftijd_Actief == False
                and self.regelingCode.OP_PP_Actief == False
                and self.regelingCode.HL_Actief == False):
                self.ui.lbl_ZL_OP.setText("€"+f"{self.regelingCode.pensioen.ouderdomsPensioen:,}".replace(',','.')+",-")
                self.ui.lbl_ZL_PP.setText("€"+f"{self.regelingCode.pensioen.partnerPensioen:,}".replace(',','.')+",-")
            else:
                self.ui.lbl_ZL_OP.setText("€—")
                self.ui.lbl_ZL_PP.setText("€—")
            
            if self.regelingCode.leeftijd_Actief == True:
                self.ui.lbl_ZL_pLeeftijd.setText(str(self.regelingCode.leeftijdJaar)+" jaar en "+str(self.regelingCode.leeftijdMaand)+" maanden")
            elif self.regelingCode.leeftijd_Actief == False:
                self.ui.lbl_ZL_pLeeftijd.setText("Leeftijd nog bepalen.")
            
            if self.regelingCode.OP_PP_Actief == True:
                self.ui.lbl_ZL_OP_PP.setText(str(self.regelingCode.OP_PP_UitruilenVan))
            elif self.regelingCode.OP_PP_Actief == False:
                self.ui.lbl_ZL_OP_PP.setText("OP/PP uitruiling n.v.t.")
            
            if self.regelingCode.HL_Actief == True:
                self.ui.lbl_ZL_hlConstructie.setText(str(self.regelingCode.HL_Volgorde))
            elif self.regelingCode.HL_Actief == False:
                self.ui.lbl_ZL_hlConstructie.setText("H/L constructie n.v.t.")
                                          
        else:
            self.ui.lbl_ZL.setText("ZL (n.v.t.)")
            self.ui.lbl_ZL_OP.setText("€—")
            self.ui.lbl_ZL_PP.setText("€—")
            
            self.ui.lbl_ZL_pLeeftijd.setText("Leeftijd n.v.t.")
            self.ui.lbl_ZL_OP_PP.setText("OP/PP uitruiling n.v.t.")
            self.ui.lbl_ZL_hlConstructie.setText("H/L constructie n.v.t.")
        
        # Aegon OP65
        if "Aegon OP65" in self._regelingenActiefKort:
            self.regelingCode = next(flexibilisatie for flexibilisatie in self.deelnemerObject.flexibilisaties if flexibilisatie.pensioen.pensioenNaam == "Aegon OP65")
            
            self.ui.lbl_A65.setText("Aegon 65")
            
            if (self.regelingCode.leeftijd_Actief == False
                and self.regelingCode.OP_PP_Actief == False
                and self.regelingCode.HL_Actief == False):
                self.ui.lbl_A65_OP.setText("€"+f"{self.regelingCode.pensioen.ouderdomsPensioen:,}".replace(',','.')+",-")
                self.ui.lbl_A65_PP.setText("€"+f"{self.regelingCode.pensioen.partnerPensioen:,}".replace(',','.')+",-")
            else:
                self.ui.lbl_A65_OP.setText("€—")
                self.ui.lbl_A65_PP.setText("€—")
            
            if self.regelingCode.leeftijd_Actief == True:
                self.ui.lbl_A65_pLeeftijd.setText(str(self.regelingCode.leeftijdJaar)+" jaar en "+str(self.regelingCode.leeftijdMaand)+" maanden")
            elif self.regelingCode.leeftijd_Actief == False:
                self.ui.lbl_A65_pLeeftijd.setText("65 jaar en 0 maanden")
            
            if self.regelingCode.OP_PP_Actief == True:
                self.ui.lbl_A65_OP_PP.setText(str(self.regelingCode.OP_PP_UitruilenVan))
            elif self.regelingCode.OP_PP_Actief == False:
                self.ui.lbl_A65_OP_PP.setText("OP/PP uitruiling n.v.t.")
            
            if self.regelingCode.HL_Actief == True:
                self.ui.lbl_A65_hlConstructie.setText(str(self.regelingCode.HL_Volgorde))
            elif self.regelingCode.HL_Actief == False:
                self.ui.lbl_A65_hlConstructie.setText("H/L constructie n.v.t.")
                
        else:
            self.ui.lbl_A65.setText("Aegon 65 (n.v.t.)")
            self.ui.lbl_A65_OP.setText("€—")
            self.ui.lbl_A65_PP.setText("€—")
            
            self.ui.lbl_A65_pLeeftijd.setText("Leeftijd n.v.t.")
            self.ui.lbl_A65_OP_PP.setText("OP/PP uitruiling n.v.t.")
            self.ui.lbl_A65_hlConstructie.setText("H/L constructie n.v.t.")
        
        # Aegon OP67
        if "Aegon OP67" in self._regelingenActiefKort:
            self.regelingCode = next(flexibilisatie for flexibilisatie in self.deelnemerObject.flexibilisaties if flexibilisatie.pensioen.pensioenNaam == "Aegon OP67")
            
            self.ui.lbl_A67.setText("Aegon 67")
            
            if (self.regelingCode.leeftijd_Actief == False
                and self.regelingCode.OP_PP_Actief == False
                and self.regelingCode.HL_Actief == False):
                self.ui.lbl_A67_OP.setText("€"+f"{self.regelingCode.pensioen.ouderdomsPensioen:,}".replace(',','.')+",-")
                self.ui.lbl_A67_PP.setText("€"+f"{self.regelingCode.pensioen.partnerPensioen:,}".replace(',','.')+",-")
            else:
                self.ui.lbl_A67_OP.setText("€—")
                self.ui.lbl_A67_PP.setText("€—")
            
            if self.regelingCode.leeftijd_Actief == True:
                self.ui.lbl_A67_pLeeftijd.setText(str(self.regelingCode.leeftijdJaar)+" jaar en "+str(self.regelingCode.leeftijdMaand)+" maanden")
            elif self.regelingCode.leeftijd_Actief == False:
                self.ui.lbl_A67_pLeeftijd.setText("67 jaar en 0 maanden")
            
            if self.regelingCode.OP_PP_Actief == True:
                self.ui.lbl_A67_OP_PP.setText(str(self.regelingCode.OP_PP_UitruilenVan))
            elif self.regelingCode.OP_PP_Actief == False:
                self.ui.lbl_A67_OP_PP.setText("OP/PP uitruiling n.v.t.")
            
            if self.regelingCode.HL_Actief == True:
                self.ui.lbl_A67_hlConstructie.setText(str(self.regelingCode.HL_Volgorde))
            elif self.regelingCode.HL_Actief == False:
                self.ui.lbl_A67_hlConstructie.setText("H/L constructie n.v.t.")
                
        else:
            self.ui.lbl_A67.setText("Aegon 67 (n.v.t.)")
            self.ui.lbl_A67_OP.setText("€—")
            self.ui.lbl_A67_PP.setText("€—")
            
            self.ui.lbl_A67_pLeeftijd.setText("Leeftijd n.v.t.")
            self.ui.lbl_A67_OP_PP.setText("OP/PP uitruiling n.v.t.")
            self.ui.lbl_A67_hlConstructie.setText("H/L constructie n.v.t.")
        
        # NN OP65
        if "NN OP65" in self._regelingenActiefKort:
            self.regelingCode = next(flexibilisatie for flexibilisatie in self.deelnemerObject.flexibilisaties if flexibilisatie.pensioen.pensioenNaam == "NN OP65")
            
            self.ui.lbl_NN65.setText("NN 65")
            
            if (self.regelingCode.leeftijd_Actief == False
                and self.regelingCode.OP_PP_Actief == False
                and self.regelingCode.HL_Actief == False):
                self.ui.lbl_NN65_OP.setText("€"+f"{self.regelingCode.pensioen.ouderdomsPensioen:,}".replace(',','.')+",-")
                self.ui.lbl_NN65_PP.setText("€"+f"{self.regelingCode.pensioen.partnerPensioen:,}".replace(',','.')+",-")
            else:
                self.ui.lbl_NN65_OP.setText("€—")
                self.ui.lbl_NN65_PP.setText("€—")
            
            if self.regelingCode.leeftijd_Actief == True:
                self.ui.lbl_NN65_pLeeftijd.setText(str(self.regelingCode.leeftijdJaar)+" jaar en "+str(self.regelingCode.leeftijdMaand)+" maanden")
            elif self.regelingCode.leeftijd_Actief == False:
                self.ui.lbl_NN65_pLeeftijd.setText("65 jaar en 0 maanden")
            
            if self.regelingCode.OP_PP_Actief == True:
                self.ui.lbl_NN65_OP_PP.setText(str(self.regelingCode.OP_PP_UitruilenVan))
            elif self.regelingCode.OP_PP_Actief == False:
                self.ui.lbl_NN65_OP_PP.setText("OP/PP uitruiling n.v.t.")
            
            if self.regelingCode.HL_Actief == True:
                self.ui.lbl_NN65_hlConstructie.setText(str(self.regelingCode.HL_Volgorde))
            elif self.regelingCode.HL_Actief == False:
                self.ui.lbl_NN65_hlConstructie.setText("H/L constructie n.v.t.")
                
        else:
            self.ui.lbl_NN65.setText("NN 65 (n.v.t.)")
            self.ui.lbl_NN65_OP.setText("€—")
            self.ui.lbl_NN65_PP.setText("€—")
            
            self.ui.lbl_NN65_pLeeftijd.setText("Leeftijd n.v.t.")
            self.ui.lbl_NN65_OP_PP.setText("OP/PP uitruiling n.v.t.")
            self.ui.lbl_NN65_hlConstructie.setText("H/L constructie n.v.t.")
        
        # NN OP67
        if "NN OP67" in self._regelingenActiefKort:
            self.regelingCode = next(flexibilisatie for flexibilisatie in self.deelnemerObject.flexibilisaties if flexibilisatie.pensioen.pensioenNaam == "NN OP67")
            
            self.ui.lbl_NN67.setText("NN 67")
            
            if (self.regelingCode.leeftijd_Actief == False
                and self.regelingCode.OP_PP_Actief == False
                and self.regelingCode.HL_Actief == False):
                self.ui.lbl_NN67_OP.setText("€"+f"{self.regelingCode.pensioen.ouderdomsPensioen:,}".replace(',','.')+",-")
                self.ui.lbl_NN67_PP.setText("€"+f"{self.regelingCode.pensioen.partnerPensioen:,}".replace(',','.')+",-")
            else:
                self.ui.lbl_NN67_OP.setText("€—")
                self.ui.lbl_NN67_PP.setText("€—")
            
            if self.regelingCode.leeftijd_Actief == True:
                self.ui.lbl_NN67_pLeeftijd.setText(str(self.regelingCode.leeftijdJaar)+" jaar en "+str(self.regelingCode.leeftijdMaand)+" maanden")
            elif self.regelingCode.leeftijd_Actief == False:
                self.ui.lbl_NN67_pLeeftijd.setText("67 jaar en 0 maanden")
            
            if self.regelingCode.OP_PP_Actief == True:
                self.ui.lbl_NN67_OP_PP.setText(str(self.regelingCode.OP_PP_UitruilenVan))
            elif self.regelingCode.OP_PP_Actief == False:
                self.ui.lbl_NN67_OP_PP.setText("OP/PP uitruiling n.v.t.")
            
            if self.regelingCode.HL_Actief == True:
                self.ui.lbl_NN67_hlConstructie.setText(str(self.regelingCode.HL_Volgorde))
            elif self.regelingCode.HL_Actief == False:
                self.ui.lbl_NN67_hlConstructie.setText("H/L constructie n.v.t.")
                
        else:
            self.ui.lbl_NN67.setText("NN 67 (n.v.t.)")
            self.ui.lbl_NN67_OP.setText("€—")
            self.ui.lbl_NN67_PP.setText("€—")
            
            self.ui.lbl_NN67_pLeeftijd.setText("Leeftijd n.v.t.")
            self.ui.lbl_NN67_OP_PP.setText("OP/PP uitruiling n.v.t.")
            self.ui.lbl_NN67_hlConstructie.setText("H/L constructie n.v.t.")
        
        # PF VLC OP68
        if "PF VLC OP68" in self._regelingenActiefKort:
            self.regelingCode = next(flexibilisatie for flexibilisatie in self.deelnemerObject.flexibilisaties if flexibilisatie.pensioen.pensioenNaam == "PF VLC OP68")
            
            self.ui.lbl_VLC.setText("PF VLC 68")
            
            if (self.regelingCode.leeftijd_Actief == False
                and self.regelingCode.OP_PP_Actief == False
                and self.regelingCode.HL_Actief == False):
                self.ui.lbl_VLC_OP.setText("€"+f"{self.regelingCode.pensioen.ouderdomsPensioen:,}".replace(',','.')+",-")
                self.ui.lbl_VLC_PP.setText("€"+f"{self.regelingCode.pensioen.partnerPensioen:,}".replace(',','.')+",-")
            else:
                self.ui.lbl_VLC_OP.setText("€—")
                self.ui.lbl_VLC_PP.setText("€—")
        
            if self.regelingCode.leeftijd_Actief == True:
                self.ui.lbl_VLC_pLeeftijd.setText(str(self.regelingCode.leeftijdJaar)+" jaar en "+str(self.regelingCode.leeftijdMaand)+" maanden")
            elif self.regelingCode.leeftijd_Actief == False:
                self.ui.lbl_VLC_pLeeftijd.setText("68 jaar en 0 maanden")
            
            if self.regelingCode.OP_PP_Actief == True:
                self.ui.lbl_VLC_OP_PP.setText(str(self.regelingCode.OP_PP_UitruilenVan))
            elif self.regelingCode.OP_PP_Actief == False:
                self.ui.lbl_VLC_OP_PP.setText("OP/PP uitruiling n.v.t.")
            
            if self.regelingCode.HL_Actief == True:
                self.ui.lbl_VLC_hlConstructie.setText(str(self.regelingCode.HL_Volgorde))
            elif self.regelingCode.HL_Actief == False:
                self.ui.lbl_VLC_hlConstructie.setText("H/L constructie n.v.t.")
                
        else:
            self.ui.lbl_VLC.setText("PF VLC 68 (n.v.t.)")
            self.ui.lbl_VLC_OP.setText("€—")
            self.ui.lbl_VLC_PP.setText("€—")
            
            self.ui.lbl_VLC_pLeeftijd.setText("Leeftijd n.v.t.")
            self.ui.lbl_VLC_OP_PP.setText("OP/PP uitruiling n.v.t.")
            self.ui.lbl_VLC_hlConstructie.setText("H/L constructie n.v.t.")
    
    def btnAndereDeelnemerClicked(self):
        self.close()
        self._logger.info("Flexmenu scherm gesloten")
        self._windowdeelnemer = Deelnemerselectie(self.book, self._logger)
        self._windowdeelnemer.show()
        
    def btnVergelijkenClicked(self):
        # Sheet van vergelijkingen openen
        self.close()
        self._logger.info("Flexmenu scherm gesloten")
        
        #drop down op vergelijkingssheet updaten
        functions.vergelijken_keuzes()
        self._logger.info("drop down op vergelijkingssheet geüpdate")
        
    def btnOpslaanClicked(self): 
        # Alle huidige flexibiliserignen opslaan in een Excel sheet
        # Huidig diagram opslaan en plaats in vergelijking sheet    
        
        if self.invoerCheck() == True:
            nieuwID = self.zoekNieuwID()
            offsetID = len(self.opslaanList)
            
            # Persoonsgegevens opslaan als dit de eerste flexibilisatie is
            if len(self.opslaanList) < 1:
                functions.persoonOpslag(self.book.sheets["Flexopslag"],self.deelnemerObject)
                
            
            # ID van de flexibilisatie in Excel opslaan
            flexID = [["Naam flexibilisatie",f"Flexibilisatie {nieuwID}"],
                     ["AfbeeldingID",f"{nieuwID}"]]
            self.book.sheets["Flexopslag"].range((2,4+4*offsetID),(3,5+4*offsetID)).options(ndims = 2).value = flexID
            self.book.sheets["Flexopslag"].range((2,4+4*offsetID),(3,5+4*offsetID)).color = (150,150,150)
            
            # Flexibilisatiekeuzes opslaan in Excel
            for regelingCount,flexibilisatie in enumerate(self.deelnemerObject.flexibilisaties):
                functions.flexOpslag(self.book.sheets["Flexopslag"],flexibilisatie,offsetID,regelingCount) 
                
            self.opslaanList.append(nieuwID)
            self.opslaanCount += 1
            
            #self.close()
            self._logger.info("Flexibilisatie opgeslagen.")
            
        elif self.invoerCheck() == False:
            self.ui.lbl_opslaanMelding.setText("Opslaan niet mogelijk bij foute invoer.")