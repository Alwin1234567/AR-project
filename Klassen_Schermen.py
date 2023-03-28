"""
Header
Hier komen alle libraries die in het programma gebruikt worden
"""
from PyQt5 import QtWidgets, uic
import functions
from datetime import datetime
from decimal import getcontext, Decimal

"""
Body
Hier komen alle functies
"""
class Functiekeus(QtWidgets.QMainWindow):
    def __init__(self, book, logger):
        self._logger = logger
        self._logger.info("Functiekeus scherm geopend")
        #De 0 staat op het einde, zodat hij de QTbaseClass niet meeneemt, deze wordt namelijk
        #Nergens gebruikt
        Functiekeus_UI = uic.loadUiType("{}\\1AdviseurBeheerder.ui".format(functions.krijgpad()))[0]
        #Ui_MainWindow, QtBaseClass = uic.loadUiType("{}\\1AdviseurBeheerder.ui".format(functions.krijgpad()))
        super(Functiekeus, self).__init__()
        self.book = book
        self.ui = Functiekeus_UI()
        self.ui.setupUi(self)
        self.setWindowTitle("Kies uw functie")
        self.ui.btnAdviseur.clicked.connect(self.btnAdviseurClicked)
        self.ui.btnBeheerder.clicked.connect(self.btnBeheerderClicked)
            
    def btnAdviseurClicked(self):
        #scherm sluiten
        self.close()
        self._logger.info("Functiekeus scherm gesloten")
        self._windowdeelnemer = Deelnemerselectie(self.book, self._logger)
        self._windowdeelnemer.show()
    def btnBeheerderClicked(self): 
        #scherm sluiten
        self.close()
        self._logger.info("Functiekeus scherm gesloten")
        if functions.isBeheerder(self.book):
            self._windowBeheerder = Beheerderkeuzes(self.book, self._logger)
            self._windowBeheerder.show()
        else:
            self._windowinlog = Inloggen(self.book, self._logger)
            self._windowinlog.show()
        


class Inloggen(QtWidgets.QMainWindow):
    def __init__(self, book, logger):
        self._logger = logger
        self._logger.info("Inloggen scherm geopend")
        inlog_UI = uic.loadUiType("{}\\2InlogBeheerder.ui".format(functions.krijgpad()))[0]
        #Ui_MainWindow2, QtBaseClass2 = uic.loadUiType("{}\\2InlogBeheerder.ui".format(functions.krijgpad()))
        super(Inloggen, self).__init__()
        self.book = book
        self.ui = inlog_UI()
        self.ui.setupUi(self)
        self.setWindowTitle("Inloggen als beheerder")
        self.ui.btnTerug.clicked.connect(self.btnTerugClicked)
        self.ui.btnInloggen.clicked.connect(self.btnInloggenClicked)
        #op enter klikken met cursur in txtBeheerderscode voert zelfde uit als klikken op knop inloggen
        self.ui.txtBeheerderscode.returnPressed.connect(self.ui.btnInloggen.click)

        self._Wachtwoord = functions.wachtwoord()
        #voorkom dat scherm gesloten kan worden met kruisje
        self._want_to_close = False
        
    def closeEvent(self, event):
        #functie die voorkomt dat het scherm gesloten kan worden met het kruisje
        if self._want_to_close == False:
            event.ignore()
            functions.Mbox("Sluiten niet mogelijk", "U kunt dit scherm niet sluiten met het kruisje.\nU kunt naar de sheet 'Vergelijken' komen door terug te gaan naar het keuzemenu en deze te sluiten of verder te gaan met inloggen.", 0)
            
    def btnInloggenClicked(self):
        if self.ui.txtBeheerderscode.text() == self._Wachtwoord:
            self._logger.info("Inloggen scherm gesloten")
            #scherm sluiten
            self._want_to_close = True
            self.close()
            #Aangeven dat beheerder ingelogd is
            beheerder = self.book.sheets["Beheerder"]
            beheerder.api.Unprotect(Password = functions.wachtwoord())
            beheerder.cells(1, 1).value = "Beheerder"
            beheerder.api.Protect(Password = functions.wachtwoord())
            self._windowBeheerder = Beheerderkeuzes(self.book, self._logger)
            self._windowBeheerder.show()
            
        else:
            self.ui.lblFoutmeldingInlog.setText("Wachtwoord incorrrect")
    def btnTerugClicked(self):
        #scherm sluiten
        self._want_to_close = True
        self.close()
        self._logger.info("Inloggen scherm gesloten")
        self._windowkeus = Functiekeus(self.book, self._logger)
        self._windowkeus.show()


class Beheerderkeuzes(QtWidgets.QMainWindow):
    def __init__(self, book, logger):
        self._logger = logger
        self._logger.info("Beheerderkeuzes scherm geopend")
        Beheerkeuzes_UI = uic.loadUiType("{}\\Beheerderkeuzes.ui".format(functions.krijgpad()))[0]
        #Ui_MainWindow2, QtBaseClass2 = uic.loadUiType("{}\\Beheerderkeuzes.ui".format(functions.krijgpad()))
        super(Beheerderkeuzes, self).__init__()
        self.book = book
        self.ui = Beheerkeuzes_UI()
        self.ui.setupUi(self)
        self.setWindowTitle("Beheerderkeuzes")
        self.ui.btnGegevensWijzigen.clicked.connect(self.btnGegevensWijzigenClicked)
        self.ui.btnBeheren.clicked.connect(self.btnBeherenClicked)
        self.ui.btnAdviseren.clicked.connect(self.btnAdviserenClicked)
        self.ui.btnUitloggen.clicked.connect(self.btnUitloggenClicked)
        
        #sheets definieren
        self.sheets = ["Sterftetafels", "AG2020", "Berekeningen", "deelnemersbestand", "Gegevens pensioencontracten", "Flexopslag", "Flexopslag"]
        self.vergelijken = self.book.sheets["Vergelijken"]
        #self.flexopslag = self.book.sheets["Flexopslag"]
        self.beheerder = self.book.sheets["Beheerder"]
    
    
    def btnGegevensWijzigenClicked(self):
        #scherm sluiten
        self.close()
        self._logger.info("Beheerderkeuzes scherm gesloten")
        self._windowWijzigen = DeelnemerselectieBeheerder(self.book, self._logger)
        self._windowWijzigen.show()
        
    
    def btnBeherenClicked(self):
        #scherm sluiten
        self.close()
        self._logger.info("Beheerderkeuzes scherm gesloten")
        #beveiliging sheets ongedaan maken
        for i in self.sheets:
            self.book.sheets[i].api.Unprotect(Password = functions.wachtwoord())
            self.book.sheets[i].visible = True
        #vergelijken unprotecten
        self.vergelijken.api.Unprotect(Password = functions.wachtwoord())
        #sheets leesbaar maken
        functions.tekstkleurSheets(self.book, self.sheets, zicht = 1)
                
    def btnAdviserenClicked(self):
        #scherm sluiten
        self.close()
        self._logger.info("Beheerderkeuzes scherm gesloten")
        self._windowdeelnemer = Deelnemerselectie(self.book, self._logger)
        self._windowdeelnemer.show()
    
    def btnUitloggenClicked(self):
        #Aangeven dat beheerder uitgelogd is
        self.beheerder.api.Unprotect(Password = functions.wachtwoord())
        self.beheerder.cells(1, 1).value = ""
        self.beheerder.api.Protect(Password = functions.wachtwoord())
        self.beheerder.visible = False
        #scherm sluiten
        self.close()
        #sheets onleesbaar maken
        functions.tekstkleurSheets(self.book, self.sheets, zicht = 0)
        #sheets beveiligen en hidden
        for i in self.sheets:
            self.book.sheets[i].api.Protect(Password = functions.wachtwoord())
            self.book.sheets[i].visible = False
        #vergelijken sheet protecten
        self.vergelijken.api.Protect(Password = functions.wachtwoord(), Contents=False)
        
        self._logger.info("Beheerderkeuzes scherm gesloten")
        self._windowkeus = Functiekeus(self.book, self._logger)
        self._windowkeus.show()
        functions.Mbox("Uitgelogd", "U bent nu uitgelogd.", 0)
               


class Deelnemerselectie(QtWidgets.QMainWindow):
    def __init__(self, book, logger):
        self._logger = logger
        self._logger.info("Deelnemerselectie scherm geopend")
        Deelnemerselectie_UI = uic.loadUiType("{}\\deelnemerselectie.ui".format(functions.krijgpad()))[0]
        #Ui_MainWindow3, QtBaseClass3 = uic.loadUiType("{}\\deelnemerselectie.ui".format(functions.krijgpad()))
        super(Deelnemerselectie, self).__init__()
        self.book = book
        self.ui = Deelnemerselectie_UI()
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
        #voorkom dat scherm gesloten kan worden met kruisje
        self._want_to_close = False
        
    def closeEvent(self, event):
        #functie die voorkomt dat het scherm gesloten kan worden met het kruisje
        if self._want_to_close == False:
            event.ignore()
            functions.Mbox("Sluiten niet mogelijk", "U kunt dit scherm niet sluiten met het kruisje.\nU kunt naar de sheet 'Vergelijken' komen door terug te gaan naar het keuzemenu en deze te sluiten of door een deelnemer te selecteren en hiervoor een flexibilisatie te starten.", 0)
            
    def btnDeelnemerToevoegenClicked(self):
        #scherm sluiten
        self._want_to_close = True
        self.close()
        self._logger.info("Deelnemerselectie scherm gesloten")
        self._windowtoevoeg = Deelnemertoevoegen(self.book, self._logger)
        self._windowtoevoeg.show()
        
    def btnStartFlexibiliserenClicked(self):
        if self.ui.lwKeuzes.currentRow() == -1: 
            self.ui.lblFoutmeldingKiezen.setText("Gelieve een deelnemer te selecteren voordat u gaat flexibiliseren")
            return
        #opgeslagen flexibilisaties van vorige deelnemer verwijderen uit opslag en vergelijken sheet
        functions.opslagLegen(self.book, self._logger)
                
        #nieuwe deelnemer aanmaken
        deelnemer = self.kleinDeelnemerlijst[self.ui.lwKeuzes.currentRow()]
        deelnemer.activeerFlexibilisatie()
        #scherm sluiten
        self._want_to_close = True
        self.close()
        self._logger.info("Deelnemerselectie scherm gesloten")
        self._windowflex = Flexmenu(self.book, deelnemer, self._logger)
        self._windowflex.show()
        
        #afbeelding huidige pensioen op vergelijken sheet plaatsen
        try: functions.maak_afbeelding(deelnemer, sheet = self.book.sheets["Vergelijken"], ID = 0, titel = "0 - Originele pensioen")
        except: 
            self._logger.exception("Fout bij het genereren van de afbeelding op Vergelijkenscherm")
        
    def btnTerugClicked(self):
        #scherm sluiten
        self._want_to_close = True
        self.close()
        self._logger.info("Deelnemerselectie scherm gesloten")
        #controleren of beheerder is ingelogd
        if functions.isBeheerder(self.book):
            self.windowBeheerder = Beheerderkeuzes(self.book, self._logger)
            self.windowBeheerder.show()
        else:
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
        deelnemertoevoegen_UI = uic.loadUiType("{}\\4DeelnemerToevoegen.ui".format(functions.krijgpad()))[0]
        #Ui_MainWindow4, QtBaseClass4 = uic.loadUiType("{}\\4DeelnemerToevoegen.ui".format(functions.krijgpad()))
        super(Deelnemertoevoegen, self).__init__()
        self.book = book
        self.ui = deelnemertoevoegen_UI()
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
        #voorkom dat scherm gesloten kan worden met kruisje
        self._want_to_close = False
        
    def closeEvent(self, event):
        #functie die voorkomt dat het scherm gesloten kan worden met het kruisje
        if self._want_to_close == False:
            event.ignore()
            functions.Mbox("Sluiten niet mogelijk", "U kunt dit scherm niet sluiten met het kruisje.\nU kunt naar de sheet 'Vergelijken' komen door terug te gaan naar het keuzemenu en deze te sluiten.", 0)
            
    def btnTerugClicked(self):
        #scherm sluiten
        self._want_to_close = True
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
                    if functions.isfloat(x.text()):   #Er is een getal-waarde ingevuld
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
                if Pensioensgegevens[i+1] == "": Pensioensgegevens[i+1] = 0
                try: verhouding = (Pensioensgegevens[i+1])/(Pensioensgegevens[i])
                except: 
                    verhouding = 0
                    verhoudingFout = verhoudingFout + ", " + PPRegelingen[int(i/2 - 2)]
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
                    #scherm sluiten
                    self._want_to_close = True
                    self.close()
                    self._logger.info("Deelnemer toevoegen scherm gesloten")
                    
                    #het parttime percentage delen door 100, zodat het in excel als % komt
                    if gegevens[7] != "": 
                        gegevens[7] = float(gegevens[7])/100
                    try: #toevoegen van de gegevens van een deelnemer aan het deelnemersbestand
                        self.book.sheets["deelnemersbestand"].api.Unprotect(Password = functions.wachtwoord())
                        functions.ToevoegenDeelnemer(gegevens)
                        functions.ProtectBeheer(self.book.sheets["deelnemersbestand"]) #.api.Protect(Password = functions.wachtwoord())
                        self._logger.info("Nieuwe deelnemer is toegevoegd aan het deelnemersbestand")
                    except:
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
    
    def __init__(self, book, deelnemer, logger, titel = ""):
        self._logger = logger
        self._logger.info("Flexmenu scherm geopend")
        Flexmenu_UI = uic.loadUiType("{}\\flexmenu.ui".format(functions.krijgpad()))[0]
        #Ui_MainWindow5, QtBaseClass5 = uic.loadUiType("{}\\flexmenu.ui".format(functions.krijgpad()))
        super(Flexmenu, self).__init__()
        self.book = book
        self.opslaanCount = 0 #Teller voor aantal opgeslagen flexibilisaties.
        self.opslaanList = list()
        self.zoekFlexibilisaties()
        
        # Deelnemer
        self.deelnemerObject = deelnemer
        
        #voorkom dat scherm gesloten kan worden met kruisje
        self._want_to_close = False
        #berekeningen sheet unprotecten
        self.book.sheets["Berekeningen"].api.Unprotect(Password = functions.wachtwoord())
        
        # Setup AOW-leeftijd knop
        self.AOWjaar = 60 # Deze wordt aangepast naar echte AOW leeftijd met functie self.getAOWleeftijd()
        self.AOWmaand = 0 # Deze wordt aangepast naar echte AOW leeftijd met functie self.getAOWleeftijd()
        self.AOW = None
        self.getAOWinformatie()
        
        
        # Setup van UI
        self.ui = Flexmenu_UI()
        self.ui.setupUi(self)
        self.setWindowTitle("Flexibilisatie menu") #Het moet na de setup, daarom staat het nu even hier
        
        #vul meegegeven titel in
        self.ui.txtTitel.setText(titel)
        self._titel = titel
        
        # Regeling selectie
        self._regelingenActiefKort = list()
        self.dropdownRegelingen()
        
        # Knoppen
        self.ui.btnAndereDeelnemer.clicked.connect(self.btnAndereDeelnemerClicked)
        self.ui.btnVergelijken.clicked.connect(self.btnVergelijkenClicked)
        self.ui.btnOpslaan.clicked.connect(self.btnOpslaanClicked)
        
        # Aanpassing: pensioenleeftijd
        self.ui.CheckLeeftijdWijzigen.stateChanged.connect(lambda: self.invoerVerandering(1))
        self.ui.sbJaar.valueChanged.connect(lambda: self.invoerVerandering(1))
        self.ui.sbMaand.valueChanged.connect(lambda: self.invoerVerandering(1))
        self.ui.btnAOWleeftijd.clicked.connect(self.btnAOWClicked)
        
        # Aanpassing: OP/PP
        self.ui.CheckUitruilen.stateChanged.connect(lambda: self.invoerVerandering(2))
        self.ui.cbUitruilenVan.activated.connect(lambda: self.invoerVerandering(2))
        self.ui.cbUMethode.activated.connect(lambda: self.invoerVerandering(2))
        self.ui.txtUVerhoudingOP.textEdited.connect(lambda: self.invoerVerandering(2))
        self.ui.txtUVerhoudingPP.textEdited.connect(lambda: self.invoerVerandering(2))
        self.ui.txtUPercentage.textEdited.connect(lambda: self.invoerVerandering(2))
           
        # Aanpassing: hoog-laag constructie
        self.ui.CheckHoogLaag.stateChanged.connect(lambda: self.invoerVerandering(3, methode = True))
        self.ui.cbHLVolgorde.activated.connect(lambda: self.invoerVerandering(3))
        self.ui.cbHLMethode.activated.connect(lambda: self.invoerVerandering(3, methode = True))
        self.ui.txtHLVerhoudingHoog.textEdited.connect(lambda: self.invoerVerandering(3))
        self.ui.txtHLVerhoudingLaag.textEdited.connect(lambda: self.invoerVerandering(3))
        self.ui.txtHLVerschil.textEdited.connect(lambda: self.invoerVerandering(3))
        self.ui.sbHLJaar.valueChanged.connect(lambda: self.invoerVerandering(3))

        # Aanpassing: Regeling
        self.ui.cbRegeling.activated.connect(self.wijzigVelden)
        
        # Aanpassing: Titel
        self.ui.txtTitel.textEdited.connect(lambda: self.invoerVerandering(4))
        
        
        # Laatste UI update
        self.ui.sbMaand.setValue(0)
        self.samenvattingUpdate()
        self.wijzigVelden()
        
        # Berekening sheet klaarmaken
        functions.berekeningen_init(book.sheets["Berekeningen"], self.deelnemerObject, self._logger)
        functions.leesOPPP(book.sheets["Berekeningen"], self.deelnemerObject.flexibilisaties)
        
        # Afbeelding genereren
        try:
            functions.maak_afbeelding(self.deelnemerObject, ax = self.ui.wdt_pltAfbeelding.canvas.ax, titel = self._titel)
            self.ui.wdt_pltAfbeelding.canvas.draw()
        except: self._logger.exception("Fout bij het genereren van de afbeelding")
        
        # Persoonsgegevens opslaan als dit de eerste flexibilisatie is
        if len(self.opslaanList) < 1:
            functions.persoonOpslag(self.book.sheets["Flexopslag"],self.deelnemerObject)
            self._logger.info("Persoonsgegevens deelnemer opgeslagen in flexopslag")
    
    def closeEvent(self, event):
        #functie die voorkomt dat het scherm met het kruisje gesloten kan worden
        if self._want_to_close == False:
            event.ignore()
            functions.Mbox("Sluiten niet mogelijk", "U kunt dit scherm niet sluiten met het kruisje.\nU kunt naar de sheet 'Vergelijken' komen door te klikken op de knop 'Vergelijken'.", 0)
            
    def blokkeerSignalen(self, actief):
        """
        Een verandering in een invoervakje geeft een 'signaal'.
        Deze functie kan de 'signalen' van invoervakjes blokkeren en activeren.
        
        actief : bool
            Als True, dan zullen wijzigingen in de window NIET geregistreerd worden.
            Als False, dan zullen wijzigingen in de window WEL geregistreerd worden.
        """
        
        if actief == True:
            self._logger.info("Veld signalen worden geblokkeerd.")
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
    
    def dropdownRegelingen(self):
        regelingenActief = list() # Lijst met lange namen van regelingen
        regelingenActiefKort = list() # Lijst met verkorte namen van regelingen
        
        for regeling in self.deelnemerObject.pensioenen:
            if regeling.ouderdomsPensioen > 0 or regeling.koopsom > 0:
                regelingenActief.append(regeling.pensioenVolNaam) # Lange regeling naam opslaan
                regelingenActiefKort.append(regeling.pensioenNaam) # Korte regeling naam opslaan

        self.ui.cbRegeling.addItems(regelingenActief) # Dropdown krijgt lijst met lange namen van regelingen
        self._regelingenActiefKort = regelingenActiefKort # Wordt apart een lijst met korte namen van regelingen opgeslagen
    
    def getAOWinformatie(self):
        """
        Deze functie zorgt ervoor dat er niet elke keer bij het klikken op de AOW knop opnieuw de 
        AOW leeftijd ingeladen moet worden, dit vergt namelijk veel tijd. In de __init__ wordt deze
        functie opgeroepen zodat het maar 1 keer opgeslagen hoeft te worden.
        """
        

        for pensioenV in self.deelnemerObject.pensioenen:
            if pensioenV.pensioenNaam == "AOW":
                self.AOWjaar = int(pensioenV.pensioenleeftijd)
                self.AOWmaand = int(round((pensioenV.pensioenleeftijd-self.AOWjaar)*12))
                self.AOW = pensioenV
        
    def wijzigVelden(self,aanpassing_aow=False):
        """
        Deze functie wordt geactiveerd als de regeling in de dropdown aangepast wordt.
        Deze functie moet alle invoervelden aanpassen naar eerder ingevoerde keuzes voor gekozen regeling.
        Als een regeling geselecteerd wordt waar nog niet eerder aanpassingen voor gemaakt zijn dan moet 
        deze functie alle velden weer leegmaken.
        """
        
        self.blokkeerSignalen(True)

        self._logger.info("Veldwijziging geïnitialiseerd.")
        
        self.ui.lblFoutmeldingLeeftijd.setText("")
        self.ui.lblFoutmeldingUitruilen.setText("")
        self.ui.lblFoutmeldingHoogLaag.setText("")
        
        # Zoek flexibilisatie-object die hoort bij huidig geselecteerde regeling in flexmenu dropdown.
        for flexibilisatie in self.deelnemerObject.flexibilisaties:
            if flexibilisatie.pensioen.pensioenVolNaam == str(self.ui.cbRegeling.currentText()):
                self.regelingCode = flexibilisatie
                break

        # --- Leeftijd velden ---
        self._logger.info("Leeftijdvelden wijzigen...")
        try:
            if self.regelingCode.HL_Actief == True and self.regelingCode.HL_Methode == "Opvullen AOW":
                self.ui.CheckLeeftijdWijzigen.setChecked(True)
                if aanpassing_aow:
                    self.ui.sbJaar.setValue(self.regelingCode.leeftijdJaar)
                else:
                    self.ui.sbJaar.setValue(self.regelingCode.AOWjaar)
                self.ui.sbMaand.setValue(self.AOWmaand)
            else:
                self.ui.CheckLeeftijdWijzigen.setChecked(self.regelingCode.leeftijd_Actief)
                self.ui.sbJaar.setValue(int(self.regelingCode.leeftijdJaar))
                self.ui.sbMaand.setValue(int(self.regelingCode.leeftijdMaand))
            
        except:
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
            
        except:
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
                if self.regelingCode.HL_Actief == True:
                    self.ui.lblFoutmeldingHoogLaag.setText("Leeftijd staat nu ingesteld op leeftijd voor AOW opvullen.")
            elif self.regelingCode.HL_Methode == "Verhouding":
                self.ui.cbHLMethode.setCurrentIndex(1)
            elif self.regelingCode.HL_Methode == "Verschil":
                self.ui.cbHLMethode.setCurrentIndex(2)
    
            self.ui.txtHLVerhoudingHoog.setText(str(self.regelingCode.HL_Verhouding_Hoog))
            self.ui.txtHLVerhoudingLaag.setText(str(self.regelingCode.HL_Verhouding_Laag))
            self.ui.txtHLVerschil.setText(str(self.regelingCode.HL_Verschil))
            self.ui.sbHLJaar.setValue(int(self.regelingCode.HL_Jaar))
            
        except:
            self._logger.exception("Probleem bij het wijzigen van hoog/laag velden.")
        
        self._logger.info("Veldwijziging afgerond.")
        
        self.blokkeerSignalen(False)
        
        self.invoerVerandering(0)
        
    def invoerCheck(self,num,checkMax=False):
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
        
        Parameters
        ----------
        num : int (0,1,2,3)
            0 verandert alle onderdelen
            1 voor verandering bij leeftijd
            2 voor verandering bij OP/PP
            3 voor verandering bij Hoog/Laag
            
        Returns
        -------
        OK : bool
            True als invoer klopt. False als invoer fout is.
        
        """
        for flexibilisatie in self.deelnemerObject.flexibilisaties:
            if flexibilisatie.pensioen.pensioenVolNaam == str(self.ui.cbRegeling.currentText()):
                self.regelingCode = flexibilisatie
                break
        
        if checkMax:
            limietList = functions.leesLimietMeldingen(self.book.sheets["Berekeningen"], 
                                                        self.deelnemerObject.flexibilisaties, 
                                                        self.regelingCode.pensioen.pensioenNaam)

        if num == 1 or num == 0: # Check of invoer klopt van leeftijd blok
            if int(self.ui.sbJaar.value()) > (self.AOWjaar+5):
                self.ui.sbJaar.setValue(self.AOWjaar+5)
                self.ui.sbMaand.setValue(self.AOWmaand)
                self.ui.lblFoutmeldingLeeftijd.setText(f"Maximum leeftijd is {self.AOWjaar+5} jaar en {self.AOWmaand} maanden.")
            elif int(self.ui.sbJaar.value()) == (self.AOWjaar+5):
                if int(self.ui.sbMaand.value()) > self.AOWmaand:
                    self.ui.sbMaand.setValue(self.AOWmaand)
                    self.ui.lblFoutmeldingLeeftijd.setText(f"Maximum leeftijd is {self.AOWjaar+5} jaar en {self.AOWmaand} maanden.")
                else:
                    self.ui.lblFoutmeldingLeeftijd.setText("")
            else:
                self.ui.lblFoutmeldingLeeftijd.setText("")
            
            if num == 1:
                return True
        
        if num == 2 or num == 0: # Check of invoer klopt van OP/PP blok
            melding, OK = functions.checkVeldInvoer("OP-PP",
                                                    self.ui.cbUMethode.currentText(),
                                                    self.ui.txtUPercentage.text(),
                                                    self.ui.txtUVerhoudingOP.text(),
                                                    self.ui.txtUVerhoudingPP.text())
            
            meldingMax = ""
            if OK and str(self.ui.cbUMethode.currentText()) == "Percentage" and checkMax:
                try:
                    if float(limietList[0][1]) > float(limietList[0][3]):
                        if float(limietList[0][3]) > 0:
                            meldingMax = f"Percentage te hoog, {round(100*limietList[0][3],2)}% wordt gehanteerd."
                            self.regelingCode.OP_PP_PercentageMax = 100*limietList[0][3]
                        elif float(limietList[0][3]) <= 0:
                            meldingMax = "OP kan niet verder uitgeruild worden naar PP."
                except:
                    pass

            totMelding = melding + " " + meldingMax
            
            self.ui.lblFoutmeldingUitruilen.setText(totMelding)
            
            if num == 2:
                return OK
            else:
                OK_OPPP = OK
   
        if num == 3 or num == 0: # Check of invoer klopt van Hoog/Laag blok
            melding, OK = functions.checkVeldInvoer("hoog-laag",
                                                    self.ui.cbHLMethode.currentText(),
                                                    self.ui.txtHLVerschil.text(),
                                                    self.ui.txtHLVerhoudingHoog.text(),
                                                    self.ui.txtHLVerhoudingLaag.text())
            
            meldingAOW = ""
            if (OK and self.ui.cbHLMethode.currentText() == "Opvullen AOW" 
                and self.ui.CheckHoogLaag.isChecked() == True):
                meldingAOW = "Leeftijd staat nu ingesteld op leeftijd voor AOW opvullen."
            
            meldingMax = ""
            if OK and str(self.ui.cbHLMethode.currentText()) == "Verschil" and checkMax:
                try:
                    if float(limietList[2][2]) > float(limietList[2][3]):
                        if float(limietList[2][3]) > 0:
                            meldingMax = f"Verschil te groot, {round(float(limietList[2][3]),2)} wordt gehanteerd."
                            self.regelingCode.HL_VerschilMax = float(limietList[2][3])
                except:
                    pass

            totMelding = melding + " " + meldingAOW + " " + meldingMax
            
            self.ui.lblFoutmeldingHoogLaag.setText(totMelding)
            
            if num == 3:
                return OK
            else:
                OK_HL = OK
        
        if num == 0:
            if OK_OPPP and OK_HL:
                return True
            else:
                return False
                
        
    def invoerVerandering(self, num, methode = False, aanpassing = False):
        """ 
        Deze functie activeert zodra de gebruiker een verandering maakt in het flexmenu scherm.
        Zo kan het scherm live aanpassen op basis van input van de gebruiker.
        
        Parameters
        ----------
        num : int (0,1,2,3)
            0 verandert alle onderdelen
            1 voor verandering bij leeftijd
            2 voor verandering bij OP/PP
            3 voor verandering bij Hoog/Laag
            4 voor verandering bij Titel
        
        methode : bool
            True betekent dat de HL methode gewijzigd is
        
        """

        if num != 4:
            if self.invoerCheck(num):
                #self.ui.lbl_opslaanMelding.setText("") # Opslaan melding verdwijnt.
                
                if aanpassing:
                    self.wijzigVelden(aanpassing_aow = True)
                    
                self.flexkeuzesOpslaan(num) # Sla flex keuzes op
                if methode: self.berekeningenDoorvoeren(1)
                else: self.berekeningenDoorvoeren(num)
                functions.leesOPPP(self.book.sheets["Berekeningen"], self.deelnemerObject.flexibilisaties) # lees de nieuwe OP en PP waardes
            
                # set leeftijd op juiste variabele
                if methode:
                    self.blokkeerSignalen(True)
                    if self.regelingCode.HL_Actief and self.regelingCode.HL_Methode == "Opvullen AOW":
                        try:
                            self.ui.CheckLeeftijdWijzigen.setChecked(True)
                            self.ui.sbJaar.setValue(self.regelingCode.AOWJaar)
                            self.ui.sbMaand.setValue(self.AOWmaand)
                            self.ui.sbHLJaar.setValue(int(self.AOWjaar-self.regelingCode.AOWJaar))
                        except: self._logger.exception("Fout bij het opslaan van opvullen AOW leeftijd.")
                    else:
                        try:
                            self.ui.CheckLeeftijdWijzigen.setChecked(self.regelingCode.leeftijd_Actief)
                            self.ui.sbJaar.setValue(self.regelingCode.leeftijdJaar)
                            self.ui.sbMaand.setValue(self.regelingCode.leeftijdMaand)
                        except: self._logger.exception("Fout bij het opslaan van normale pensioenleeftijd.")
                    self.blokkeerSignalen(False)
                
                self.samenvattingUpdate() # Update de samenvatting
        else:
            self._titel = str(self.ui.txtTitel.text())
            #print(self._titel)
                
        try: # probeer een nieuwe afbeelding te maken
            functions.maak_afbeelding(self.deelnemerObject, ax = self.ui.wdt_pltAfbeelding.canvas.ax, titel = self._titel)
            self.ui.wdt_pltAfbeelding.canvas.draw()
        except: self._logger.exception("Fout bij het genereren van de afbeelding")
            
        self.invoerCheck(num,True)

    def zoekFlexibilisaties(self):
        self.opslaanList = self.book.sheets["Flexopslag"].range("3:3")[1:500].value # Zoekt maximaal tot 500 anders wordt het langzaam
        self.opslaanList = [int(value) for value in self.opslaanList if type(value) == float] # Flex IDs worden uit Excel opgehaald als type Float, moet opgeslagen worden als type int

    def zoekNieuwID(self):
        if len(self.opslaanList) > 0: # Als de lijst niet leeg is, zijn er al ID's opgeslagen en moet de eerstvolgende "lege" ID gevonden worden
            for i in range(len(self.opslaanList)):
                if (i+1) not in self.opslaanList: # Als getal niet in lijst staat, is dit het nieuwe ID
                    return (i+1)
            return len(self.opslaanList)+1
        else:
            return 1 # Als lijst leeg is, zijn er nog geen ID's opgeslagen. De eerste moet ID waarde 1 krijgen.
    
    
    def flexkeuzesOpslaan(self, num):
        """
        Deze functie slaat huidig ingevulde flex opties op in het flexibiliseringsobject.
        
        Parameters
        ----------
        num : int (0,1,2,3)
            0 verandert alle onderdelen
            1 voor verandering bij leeftijd
            2 voor verandering bij OP/PP
            3 voor verandering bij Hoog/Laag
        
        """
        
        # self.regelingCode = functions.regelingNaamCode(str(self.ui.cbRegeling.currentText()))
        
        self._logger.info("Flexkeuze opslaan geïnitialiseerd.")

        # Zoek flexibilisatie-object die hoort bij huidig geselecteerde regeling in flexmenu dropdown.
        for flexibilisatie in self.deelnemerObject.flexibilisaties:
            if flexibilisatie.pensioen.pensioenVolNaam == str(self.ui.cbRegeling.currentText()):
                self.regelingCode = flexibilisatie
                break

        if num == 1 or num == 0: # Leeftijd wijziging opslaan
            if self.regelingCode.HL_Actief and self.regelingCode.HL_Methode == "Opvullen AOW":
                try: # Hier worden de spinboxes eerst naar de juiste waardes gezet voordat ze opgeslagen kunnen worden.
                    self.blokkeerSignalen(True)
                    if self.ui.sbJaar.value() > (self.AOWjaar-2):
                        self.ui.sbJaar.setValue(int(self.AOWjaar-1))
                    self.ui.sbMaand.setValue(int(self.AOWmaand))
                    self.blokkeerSignalen(False)
                    
                    self.ui.CheckLeeftijdWijzigen.setChecked(True)
                    jaar = int(self.ui.sbJaar.value())
                    maand = int(self.ui.sbMaand.value())
                    self.deelnemerObject.setAOWLeeftijd(jaar, maand, self.AOWjaar)
                    
                    self.regelingCode.HL_Jaar = int(self.AOWjaar-self.regelingCode.AOWJaar)
                    
                    self.blokkeerSignalen(True)
                    self.ui.sbHLJaar.setValue(self.regelingCode.HL_Jaar)
                    self.blokkeerSignalen(False)
                    
                except: self._logger.exception("Er gaat iets fout bij het corrigeren van de leeftijd spinboxes voor AOW opvullen in flexmenu.ui")
            else:
                try:
                    self.regelingCode.leeftijd_Actief = self.ui.CheckLeeftijdWijzigen.isChecked()
                    self.regelingCode.leeftijdJaar = int(self.ui.sbJaar.value())
                    self.regelingCode.leeftijdMaand = int(self.ui.sbMaand.value())
                except: self._logger.exception("Er gaat iets fout bij het opslaan van de pensioenleeftijd in flexmenu.ui")
        
        if num == 2 or num == 0: # OP/PP uitruiling opslaan
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
                        self.regelingCode.OP_PP_Percentage = float(self.ui.txtUPercentage.text())
                
                elif str(self.ui.cbUMethode.currentText()) == "Percentage":
                    self.regelingCode.OP_PP_Percentage = float(self.ui.txtUPercentage.text())
                    
                    if str(self.ui.txtUVerhoudingOP.text()) == "":
                        self.regelingCode.OP_PP_Verhouding_OP = 0
                    else:
                        self.regelingCode.OP_PP_Verhouding_OP = int(self.ui.txtUVerhoudingOP.text())
                    
                    if str(self.ui.txtUVerhoudingPP.text()) == "":
                        self.regelingCode.OP_PP_Verhouding_PP = 0
                    else:
                        self.regelingCode.OP_PP_Verhouding_PP = int(self.ui.txtUVerhoudingPP.text())
    
            except:
                self._logger.exception("Huidig geselecteerde OP/PP flexibilisaties in flexmenu.ui kunnen niet opgeslagen worden.")
        
        if num == 3 or num == 0: # Hoog/Laag constructie opslaan
            try:
                # Onderstaande if-elif statements regelen het aanpassen van de hoog-laag eerste periode spinbox op basis van de leeftijd voor opvullen AOW
                # Het houdt ook rekening met of de spinbox veranderd moet worden of dat de leeftijd veranderd moet worden als de spinbox verandert.
                if self.regelingCode.HL_Methode != "Opvullen AOW" and str(self.ui.cbHLMethode.currentText()) == "Opvullen AOW" and self.regelingCode.HL_Actief:
                    # Dit activeert als Opvullen AOW nu geselecteerd is en de hoog-laag constructie staat al geactiveerd.
                    self.blokkeerSignalen(True)
                    self.ui.sbHLJaar.setValue(int(self.AOWjaar-self.regelingCode.AOWJaar))
                    self.regelingCode.HL_Jaar = int(self.AOWjaar-self.regelingCode.AOWJaar)
                    self.blokkeerSignalen(False)
                
                elif self.regelingCode.HL_Methode == "Opvullen AOW" and self.regelingCode.HL_Actief == False and self.ui.CheckHoogLaag.isChecked():
                    # Dit activeert als Opvullen AOW al geselcteerd was en op dit moment hoog-laag constructie geactiveerd wordt.
                    self.blokkeerSignalen(True)
                    self.ui.sbHLJaar.setValue(int(self.AOWjaar-self.regelingCode.AOWJaar))
                    self.regelingCode.HL_Jaar = int(self.AOWjaar-self.regelingCode.AOWJaar)
                    self.blokkeerSignalen(False)

                elif (self.regelingCode.HL_Methode == "Opvullen AOW" and 
                      str(self.ui.cbHLMethode.currentText()) == "Opvullen AOW" and self.regelingCode.HL_Actief):
                    # Dit activeert als opvullen AOW al geselecteerd is en de hoog-laag constructie ook al op actief stond.
                    
                    self.blokkeerSignalen(True)
                    self.ui.CheckLeeftijdWijzigen.setChecked(True)
                    
                    if int(self.ui.sbHLJaar.value()) > (self.AOWjaar-60):
                        self.ui.sbHLJaar.setValue(int(self.AOWjaar-60))
                        self.regelingCode.HL_Jaar = int(self.AOWjaar-60)
                    else:
                        self.regelingCode.HL_Jaar = int(self.ui.sbHLJaar.value())
                    
                    self.ui.sbJaar.setValue(int(self.AOWjaar-self.regelingCode.HL_Jaar))
                    self.ui.sbMaand.setValue(int(self.AOWmaand))
                    
                    self.blokkeerSignalen(False)
                    
                    jaar = int(self.ui.sbJaar.value())
                    maand = int(self.ui.sbMaand.value())
                    
                    self.deelnemerObject.setAOWLeeftijd(jaar, maand, self.AOWjaar)
                    
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
                        self.regelingCode.HL_Verschil = float(self.ui.txtHLVerschil.text())
                    
                elif str(self.ui.cbHLMethode.currentText()) == "Verschil":
                    self.regelingCode.HL_Verschil = float(self.ui.txtHLVerschil.text())
                    
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
                        self.regelingCode.HL_Verschil = float(self.ui.txtHLVerschil.text())
                        
                    if str(self.ui.txtHLVerhoudingHoog.text()) == "":
                        self.regelingCode.HL_Verhouding_Hoog = 0
                    else:
                        self.regelingCode.HL_Verhouding_Hoog = int(self.ui.txtHLVerhoudingHoog.text())
                    
                    if str(self.ui.txtHLVerhoudingLaag.text()) == "":
                        self.regelingCode.HL_Verhouding_Laag = 0
                    else:
                        self.regelingCode.HL_Verhouding_Laag = int(self.ui.txtHLVerhoudingLaag.text())
                     
            except: self._logger.exception("Huidig geselecteerde hoog/laag flexibilisaties in flexmenu.ui kunnen niet opgeslagen worden.")
    
    def AOWberekenen(self, overbruggingen):
        for flexcombo in overbruggingen:
            instellingen = functions.berekeningen_instellingen()
            blokhoogte = instellingen["pensioeninfohoogte"] + instellingen["afstandtotblokken"] + len(self.deelnemerObject.flexibilisaties) + flexcombo[0] * (instellingen["blokgrootte"] + instellingen["afstandtussenblokken"])
            factoren = self.book.sheets["Berekeningen"].range((blokhoogte + 12, 2), (blokhoogte + 16, 2)).value
            if self.deelnemerObject.burgelijkeStaat == "Samenwonend": bedrag = self.AOW.samenwondendAOW
            else: bedrag = self.AOW.alleenstaandAOW
            for combo in overbruggingen:
                if combo == flexcombo: continue
                bedrag += combo[1].ouderdomsPensioenHoog
                bedrag -= combo[1].ouderdomsPensioenLaag
            self.book.sheets["Berekeningen"].range((blokhoogte + 1, 2)).value = flexcombo[1].AOWJaar + flexcombo[1].AOWMaand / 12
            blok = list()
            if bedrag < 0 :
                flexcombo[1].HL_Volgorde = "Laag-hoog"
                blok.append(["Verschil", "Laag-hoog"])
                blok.append([self.AOWjaar - flexcombo[1].AOWJaar, -bedrag])
            else: 
                flexcombo[1].HL_Volgorde = "Hoog-laag"
                blok.append(["Verschil", "Hoog-laag"])
                blok.append([self.AOWjaar - flexcombo[1].AOWJaar, bedrag])
            updaterange = self.book.sheets["Berekeningen"].range((blokhoogte + 4, 2),\
                                                                 (blokhoogte + 5, 3))
            try: updaterange.value = blok
            except: self._logger.exception("error bij het updaten van de Verekeningsheet")    

    def berekeningenDoorvoeren(self, num):
        instellingen = functions.berekeningen_instellingen()
        overbruggingen = list()
        for i, flexibilisatie in enumerate(self.deelnemerObject.flexibilisaties):
            if flexibilisatie.HL_Actief and flexibilisatie.HL_Methode == "Opvullen AOW": overbruggingen.append((i, flexibilisatie))
        
        overbruggingen.sort(key = lambda flex: flex[1].pensioen.rente)
        
        for i, flexibilisatie in enumerate(self.deelnemerObject.flexibilisaties):
            blokhoogte = instellingen["pensioeninfohoogte"] + instellingen["afstandtotblokken"] + len(self.deelnemerObject.flexibilisaties) + i * (instellingen["blokgrootte"] + instellingen["afstandtussenblokken"])
            blok = list()
            
            # Leeftijd doorvoeren
            if flexibilisatie.HL_Actief and flexibilisatie.HL_Methode == "Opvullen AOW": blok.append([flexibilisatie.AOWJaar + flexibilisatie.AOWMaand / 12, ""])
            elif flexibilisatie.leeftijd_Actief: blok.append([flexibilisatie.leeftijdJaar + flexibilisatie.leeftijdMaand / 12, ""])
            else: blok.append([flexibilisatie.pensioen.pensioenleeftijd, ""])
            
            if num == 1 and flexibilisatie.pensioen.actieveRegeling:
                bedragen = functions.regelingBedrag(self.deelnemerObject, flexibilisatie)
                if flexibilisatie.pensioen.pensioenSoortRegeling == "DC":
                    try: self.book.sheets["Berekeningen"].range((blokhoogte + 6, 4)). value = bedragen[2]
                    except Exception as e: self._logger.exception("error bij het updaten van de Verekeningsheet")
                elif flexibilisatie.pensioen.pensioenSoortRegeling == "DB":
                    try: self.book.sheets["Berekeningen"].range((blokhoogte + 6, 2)). value = bedragen[0]
                    except Exception as e: self._logger.exception("error bij het updaten van de Verekeningsheet")
                elif flexibilisatie.pensioen.pensioenSoortRegeling == "DB met PP":
                    try: self.book.sheets["Berekeningen"].range((blokhoogte + 6, 2), (blokhoogte + 6, 3)). value = [bedragen[0], bedragen[1]]
                    except Exception as e: self._logger.exception("error bij het updaten van de Verekeningsheet")
                else: self._logger.warning("onbekende actieve regeling")
            
            # OP/PP doorvoeren
            if flexibilisatie.OP_PP_Actief:
                if flexibilisatie.OP_PP_Methode == "Verhouding":
                    blok.append([flexibilisatie.OP_PP_Methode, ""])
                    blok.append(["1", str(min(flexibilisatie.OP_PP_Verhouding_PP / flexibilisatie.OP_PP_Verhouding_OP, 0.7))])
                else:
                    blok.append(["{} {}".format(flexibilisatie.OP_PP_UitruilenVan, flexibilisatie.OP_PP_Methode), ""])
                    blok.append([flexibilisatie.OP_PP_Percentage / 100, ""])
            else:
                blok.append(["", ""])
                blok.append(["", ""])
                
            # Hoog/Laag doorvoeren
            if i not in [flex[0] for flex in overbruggingen]:
                if flexibilisatie.HL_Actief:
                    blok.append([flexibilisatie.HL_Methode, flexibilisatie.HL_Volgorde])
                    if flexibilisatie.HL_Methode == "Verhouding": blok.append([flexibilisatie.HL_Jaar, min(max(flexibilisatie.HL_Verhouding_Laag / flexibilisatie.HL_Verhouding_Hoog, 3/4), 1)])
                    else: blok.append([flexibilisatie.HL_Jaar, max(flexibilisatie.HL_Verschil, 0)])
                else:
                    blok.append(["", ""])
                    blok.append(["", ""])
            else:
                blok.append(["Verschil", "Laag-hoog"])
                blok.append([self.AOWjaar - flexibilisatie.AOWJaar, flexibilisatie.ouderdomsPensioenHoog])

            updaterange = self.book.sheets["Berekeningen"].range((blokhoogte + 1, 2),\
                                                                 (blokhoogte + 5, 3))
            try: updaterange.value = blok
            except: self._logger.exception("error bij het updaten van de Verekeningsheet")
        if len(overbruggingen) > 0: self.AOWberekenen(overbruggingen)
       
        
    def samenvattingUpdate(self):
        """
        Deze functie update de waarden in de samenvatting boxes.
        """
        
        self._logger.info("Samenvatting updaten...")
        
        regelingenList = ["ZL","Aegon OP65","Aegon OP67","NN OP65","NN OP67","PF VLC OP68"]

        for regeling in regelingenList:
            for flexibilisatie in self.deelnemerObject.flexibilisaties:
                if flexibilisatie.pensioen.pensioenNaam == str(regeling):
                    self.regelingCode = flexibilisatie
                    break
            
            regelingDict = functions.samenvattingDict(regeling,self.ui)
            
            if regeling in self._regelingenActiefKort:
                regelingDict["lbl"].setText(f"{regeling}")
                
                self.update_samenvatting(regelingDict["lbl_OP"], regelingDict["lbl_PP"])
                
                if self.regelingCode.leeftijd_Actief: regelingDict["lbl_pLeeftijd"].setText(str(self.regelingCode.leeftijdJaar)+" jaar en "+str(self.regelingCode.leeftijdMaand)+" maanden")
                else:
                    regelingDict["lbl_pLeeftijd"].setText(str(int(self.regelingCode.pensioen.pensioenleeftijd)) + " jaar en " + str(0) + " maanden")
                
                if self.regelingCode.OP_PP_Actief: regelingDict["lbl_OP_PP"].setText(str(self.regelingCode.OP_PP_UitruilenVan))
                else: regelingDict["lbl_OP_PP"].setText("OP/PP uitruiling n.v.t.")
                
                if self.regelingCode.HL_Actief: regelingDict["lbl_hlConstructie"].setText(str(self.regelingCode.HL_Volgorde))
                else: regelingDict["lbl_hlConstructie"].setText("H/L constructie n.v.t.")
                                              
            else:
                regelingDict["lbl"].setText(f"{regeling} (n.v.t.)")
                regelingDict["lbl_OP"].setText("€—")
                regelingDict["lbl_PP"].setText("€—")
                
                regelingDict["lbl_pLeeftijd"].setText("Leeftijd n.v.t.")
                regelingDict["lbl_OP_PP"].setText("OP/PP uitruiling n.v.t.")
                regelingDict["lbl_hlConstructie"].setText("H/L constructie n.v.t.")
    
    def update_samenvatting(self, lblOP, lblPP):
        # leest de informatie uit het regelingscode object en zet het in de juiste samenvatting
        if self.regelingCode.HL_Actief: lblOP.setText("€{},-/{},-".format(self.regelingCode.ouderdomsPensioenHoog, self.regelingCode.ouderdomsPensioenLaag))
        else: lblOP.setText("€{},-".format(self.regelingCode.ouderdomsPensioenHoog))
        lblPP.setText("€{},-".format(self.regelingCode.partnerPensioen))
    
    def btnAOWClicked(self):
        self.blokkeerSignalen(True)
        self._logger.info("AOW-leeftijd button in flexmenu geklikt.")

        try:
            self.ui.sbJaar.setValue(self.AOWjaar)
            self.ui.sbMaand.setValue(self.AOWmaand)
        except:
            self._logger.exception("Probleem bij het wijzigen van leeftijdvelden naar AOW-leeftijd.")

        self.blokkeerSignalen(False)
        self.invoerVerandering(0)
        
    def btnAndereDeelnemerClicked(self):
        #controleren of gebruiker echt andere deelnemer wil selecteren
        controle = functions.Mbox("Andere deelnemer selecteren", "Door een andere deelnemer te selecteren zullen de opgeslagen flexibiliseringen voor de huidige deelnemer verwijderd worden.\nU kunt deze actie niet ongedaan maken.", 1)
        if controle == "OK Clicked":
            #scherm sluiten
            self._want_to_close = True
            self.close()
            #berekeningensheet protecten (als geen beheerder is ingelogd)
            functions.ProtectBeheer(self.book.sheets["Berekeningen"])
            self._logger.info("Flexmenu scherm gesloten")
            self._windowdeelnemer = Deelnemerselectie(self.book, self._logger)
            self._windowdeelnemer.show()
        
    def btnVergelijkenClicked(self):
        #scherm sluiten
        self._want_to_close = True
        self.close()
        #berekeningensheet protecten (als geen beheerder is ingelogd)
        functions.ProtectBeheer(self.book.sheets["Berekeningen"])
        self._logger.info("Flexmenu scherm gesloten")
        
        # Drop down op vergelijkingssheet updaten
        functions.vergelijken_keuzes()
        
        # Open Vergelijken sheet
        self.book.sheets["Vergelijken"].activate()
        self._logger.info("drop down op vergelijkingssheet geüpdate")
        
    def btnOpslaanClicked(self): 
        # Alle huidige flexibiliserignen opslaan in een Excel sheet
        # Huidig diagram opslaan en plaats in vergelijking sheet    

        nieuwID = self.zoekNieuwID()
        offsetID = len(self.opslaanList)
        
        # ID nummer met laatste opgeslagen flexibilisaties ophogen
        flexopslag = self.book.sheets["Flexopslag"]
        if str(flexopslag.cells(2, 5).value) != "None":   # Alleen als er nog flexibilisaties opgeslagen zijn
            # ID-nummer van laatste opgeslagen flexibilisatie vinden
            kolomLaatsteOpslag = functions.FlexopslagVinden(self.book)[1]
            IDLaatsteOpslag = flexopslag.cells(3, kolomLaatsteOpslag).value
            IDoud = int(IDLaatsteOpslag.split()[-1])
            if nieuwID != IDoud + 1:
                nieuwID = nieuwID + IDoud
                offsetID = (kolomLaatsteOpslag-1)/4
        
        # Afbeelding op vergelijkingsSheet zetten
        try: functions.maak_afbeelding(self.deelnemerObject, sheet = self.book.sheets["Vergelijken"], ID = nieuwID, titel = f"{nieuwID} - {self._titel}")
        except: self._logger.exception("Fout bij het genereren van de afbeelding op Vergelijkenscherm")
        
        
        #sheet flexopslag unprotecten
        self.book.sheets["Flexopslag"].api.Unprotect(Password = functions.wachtwoord())
        # ID van de flexibilisatie in Excel opslaan     # oude titel = f"Flexibilisatie {nieuwID}"
        flexID = [["Naam flexibilisatie",f"{nieuwID} - {self._titel}"],
                 ["AfbeeldingID",f"Vergelijking {nieuwID}"]]
        self.book.sheets["Flexopslag"].range((2,4+4*offsetID),(3,5+4*offsetID)).options(ndims = 2).value = flexID
        self.book.sheets["Flexopslag"].range((2,4+4*offsetID),(3,5+4*offsetID)).color = (150,150,150)
        
        # Flexibilisatiekeuzes opslaan in Excel
        for regelingCount,flexibilisatie in enumerate(self.deelnemerObject.flexibilisaties):
            functions.flexOpslag(self.book.sheets["Flexopslag"],flexibilisatie,offsetID,regelingCount) 
        #sheet flesopslag protecten
        functions.ProtectBeheer(self.book.sheets["Flexopslag"])
        
        # Melding geven dat flexibilisatie opgeslagen is
        tekst = "Deze flexibilisatie is opgeslagen. \nU kunt nu verder flexibiliseren. \nMet de knop 'Vergelijken' kunt u uw opgeslagen flexibilisaties vergelijken."
        functions.Mbox("Flexibilisatie opgeslagen", tekst, 0)
        
        self.opslaanList.append(nieuwID)
        self.opslaanCount += 1

        self._logger.info("Flexibilisatie opgeslagen.")
            

class DeelnemerselectieBeheerder(QtWidgets.QMainWindow):
    def __init__(self, book, logger):
        self._logger = logger
        self._logger.info("DeelnemerselectieBeheerder scherm geopend")
        DeelnemerselectieBeheerder_UI = uic.loadUiType("{}\\deelnemerselectie.ui".format(functions.krijgpad()))[0]
        #Ui_MainWindow3, QtBaseClass3 = uic.loadUiType("{}\\deelnemerselectie.ui".format(functions.krijgpad()))
        super(DeelnemerselectieBeheerder, self).__init__()
        self.book = book
        self.ui = DeelnemerselectieBeheerder_UI()
        self.ui.setupUi(self)
        self.setWindowTitle("Deelnemerselectie")
        self.deelnemerlijst = functions.getDeelnemersbestand(self.book)
        self.kleinDeelnemerlijst = list()
        #naam button aanpassen
        self.ui.btnStartFlexibiliseren.setText("Gegevens wijzigen")
        #button deelnemer toevoegen verbergen
        self.ui.btnDeelnemerToevoegen.hide()
        #buttons connecten
        self.ui.btnStartFlexibiliseren.clicked.connect(self.btnGegevensWijzigenClicked)
        self.ui.btnTerug.clicked.connect(self.btnTerugClicked)
        self.ui.sbDag.valueChanged.connect(lambda: self.onChange(False))
        self.ui.sbMaand.valueChanged.connect(lambda: self.onChange(True))
        self.ui.sbJaar.valueChanged.connect(lambda: self.onChange(True))
        self.ui.txtVoorletters.textChanged.connect(lambda: self.onChange(False))
        self.ui.txtTussenvoegsel.textChanged.connect(lambda: self.onChange(False))
        self.ui.txtAchternaam.textChanged.connect(lambda: self.onChange(False))
        self.ui.cbGeslacht.currentTextChanged.connect(lambda: self.onChange(False))
        self.ui.lwKeuzes.currentItemChanged.connect(self.clearError)
        #voorkom dat scherm gesloten kan worden met kruisje
        self._want_to_close = False
    
    def closeEvent(self, event):
        #functie die voorkomt dat het scherm met het kruisje gesloten kan worden
        if self._want_to_close == False:
            event.ignore()
            functions.Mbox("Sluiten niet mogelijk", "U kunt dit scherm niet sluiten met het kruisje.\nU kunt naar de sheets inzien door terug te gaan naar de beheerderskeuzes en deze te sluiten of te klikken op 'Beheren'.", 0)
            
    def btnGegevensWijzigenClicked(self):
        if self.ui.lwKeuzes.currentRow() == -1: 
            self.ui.lblFoutmeldingKiezen.setText("Gelieve een deelnemer te selecteren om de gegevens te wijzigen")
            return
        
        #nieuwe deelnemer aanmaken
        deelnemer = self.kleinDeelnemerlijst[self.ui.lwKeuzes.currentRow()]
        deelnemer.activeerFlexibilisatie()
        #scherm sluiten
        self._want_to_close = True
        self.close()
        self._logger.info("DeelnemerselectieBeheerder scherm gesloten")
        self._windowWijzigen = DeelnemerWijzigen(self.book, self._logger, deelnemer)
        self._windowWijzigen.show()
        
        
    def btnTerugClicked(self):
        #scherm sluiten
        self._want_to_close = True
        self.close()
        self._logger.info("DeelnemerselectieBeheerder scherm gesloten")
        #controleren of beheerder is ingelogd
        if functions.isBeheerder(self.book):
            self.windowBeheerder = Beheerderkeuzes(self.book, self._logger)
            self.windowBeheerder.show()
        else:
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
  

class DeelnemerWijzigen(QtWidgets.QMainWindow):
    def __init__(self, book, logger, deelnemer):
        self._logger = logger
        self._logger.info("Deelnemer wijzigen scherm geopend")
        DeelnemerWijzigen_UI = uic.loadUiType("{}\\4DeelnemerToevoegen.ui".format(functions.krijgpad()))[0]
        #Ui_MainWindow4, QtBaseClass4 = uic.loadUiType("{}\\4DeelnemerToevoegen.ui".format(functions.krijgpad()))
        super(DeelnemerWijzigen, self).__init__()
        self.book = book
        self.ui = DeelnemerWijzigen_UI()
        self.ui.setupUi(self)
        self.setWindowTitle("Deelnemer wijzigen")
        #naam button aanpassen
        self.ui.btnToevoegen.setText("Wijzigen")
        #buttons connecten
        self.ui.btnTerug.clicked.connect(self.btnTerugClicked)
        self.ui.btnToevoegen.clicked.connect(self.btnWijzigenClicked)
        self.ui.sbMaand.valueChanged.connect(self.onChange)
        self.ui.sbJaar.valueChanged.connect(self.onChange)
        self._30maand = [4,6,9,11]
        #voorkom dat scherm gesloten kan worden met kruisje
        self._want_to_close = False
        #voeg schaduwtekst toe aan de invoervelden
        self.ui.txtVoorletters.setPlaceholderText("A.B.")
        self.ui.txtTussenvoegsel.setPlaceholderText("van")
        self.ui.txtAchternaam.setPlaceholderText("Albert")
        self.ui.txtParttimePercentage.setPlaceholderText("70")
        for i in [self.ui.txtFulltimeLoon, self.ui.txtOPAegon65, self.ui.txtOPAegon67, self.ui.txtOPNN65, self.ui.txtOPNN67, 
                  self.ui.txtOPVLC68, self.ui.txtOPZL, self.ui.txtPPNN65, self.ui.txtPPNN67, self.ui.txtPPVLC68]:
            i.setPlaceholderText("500,00")
        #voeg deelnemergegevens toe aan invoervelden
        self.ui.txtVoorletters.setText(deelnemer.voorletters)
        self.ui.txtTussenvoegsel.setText(deelnemer.tussenvoegsels)
        self.ui.txtAchternaam.setText(deelnemer.achternaam)
        self.ui.sbDag.setValue(deelnemer.geboortedatum.day)
        self.ui.sbMaand.setValue(deelnemer.geboortedatum.month)
        self.ui.sbJaar.setValue(deelnemer.geboortedatum.year)
        self.ui.cbGeslacht.setCurrentText(deelnemer.geslacht)               
        self.ui.cbBurgerlijkeStaat.setCurrentText(deelnemer.burgelijkeStaat)
        self.ui.cbHuidigeRegeling.setCurrentText(deelnemer.regeling)
        if str(deelnemer.regeling) != "Inactief": 
            self.ui.txtFulltimeLoon.setText(str(deelnemer.ftLoon).replace(".", ","))
            self.ui.txtParttimePercentage.setText(str(deelnemer.pt*100).replace(".", ","))
        else:
            self.ui.txtFulltimeLoon.setText("")
            self.ui.txtParttimePercentage.setText("")
        pensioenfondsen = [["ZL", self.ui.CheckZL, self.ui.txtOPZL], ["Aegon OP65", self.ui.CheckAegon65, self.ui.txtOPAegon65], 
                            ["Aegon OP67", self.ui.CheckAegon67, self.ui.txtOPAegon67], ["NN OP65", self.ui.CheckNN65, self.ui.txtOPNN65, self.ui.txtPPNN65], 
                            ["NN OP67", self.ui.CheckNN67, self.ui.txtOPNN67, self.ui.txtPPNN67], ["PF VLC OP68", self.ui.CheckPFVLC68, self.ui.txtOPVLC68, self.ui.txtPPVLC68]]
        
        #OP en PP invullen en aanvinken als ingevuld
        x = 1   #begin op 1, omdat op plek 0 AOW zit
        for i in range(6):
            if x < len(deelnemer.pensioenen):
                if pensioenfondsen[i][0] == deelnemer.pensioenen[x].pensioenNaam:
                    pensioenfondsen[i][2].setText(str(deelnemer.pensioenen[x].ouderdomsPensioen).replace(".", ","))
                    if str(deelnemer.pensioenen[x].ouderdomsPensioen) != "":
                        pensioenfondsen[i][1].setChecked(True)
                    if i > 2:
                        pensioenfondsen[i][3].setText(str(deelnemer.pensioenen[x].partnerPensioen).replace(".", ","))
                    x += 1
        
        self.rijNr = deelnemer.rijNr
        
    def closeEvent(self, event):
        #functie die voorkomt dat het scherm gesloten kan worden met het kruisje
        if self._want_to_close == False:
            event.ignore()
            functions.Mbox("Sluiten niet mogelijk", "U kunt dit scherm niet sluiten met het kruisje.\nU kunt naar de sheets inzien door terug te gaan naar de beheerderskeuzes en deze te sluiten of te klikken op 'Beheren'.", 0)
            
    def btnTerugClicked(self):
        #scherm sluiten
        self._want_to_close = True
        self.close()
        self._logger.info("Deelnemer wijzigen scherm gesloten")
        self._windowdeelnemer = DeelnemerselectieBeheerder(self.book, self._logger)
        self._windowdeelnemer.show()
        
    def btnWijzigenClicked(self):
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
                        
                        
            #geboortedatum in goede notatie voor invoer in excel
            geboortedatum = datetime(int(self.ui.sbJaar.text()), int(self.ui.sbMaand.text()), int(self.ui.sbDag.text())).strftime("%m-%d-%Y")
            Deelnemersgegevens[3] = geboortedatum
            #lijst met alle gegevens
            gegevens = Deelnemersgegevens + Pensioensgegevens
            
            #deelnemer zijn gegevens laten controleren
            self._logger.info("Ingevulde gegevens worden getoont voor controle")
            controle = functions.gegevenscontrole(gegevens)
            if controle == "correct":
                #scherm sluiten
                self._want_to_close = True
                self.close()
                self._logger.info("Deelnemer toevoegen scherm gesloten")
                
                #het parttime percentage delen door 100, zodat het in excel als % komt
                if gegevens[7] != "": 
                    gegevens[7] = float(gegevens[7])/100
                try: #toevoegen van de gegevens van een deelnemer aan het deelnemersbestand
                    self.book.sheets["deelnemersbestand"].api.Unprotect(Password = functions.wachtwoord())
                    functions.ToevoegenDeelnemer(gegevens, regel = self.rijNr)
                    functions.ProtectBeheer(self.book.sheets["deelnemersbestand"]) #.api.Protect(Password = functions.wachtwoord())
                    self._logger.info("Deelnemersgegevens zijn gewijzigd in het deelnemersbestand")
                except:
                    self._logger.exception("Er is iets fout gegaan bij het wijzigen van een deelnemer in het deelnemersbestand")
                
                #deelnemerselectie openen
                self._windowBeheerder = Beheerderkeuzes(self.book, self._logger)
                self._windowBeheerder.show()
            elif controle == "fout":
                self._logger.info("Deelnemer wil zijn ingevulde gegevens aanpassen. Deelnemer wijzigen scherm blijft open")
                #als niet op "ja" wordt geklikt, wordt de messagebox gesloten en het invoerveld weer getoont
        
                
        else: 
            self._logger.info("Niet alle deelnemersgegevens zijn goed ingevuld. De deelnemer moet zijn gegevens aanpassen")
            #foutmelding tonen
            self.ui.lblFoutmeldingGegevens.setText(foutmeldingGegevens)
            self.ui.lblFoutmeldingPensioen.setText(foutmeldingPensioen)
        
    
    
    def onChange(self): functions.maanddag(self)


