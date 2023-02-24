"""
Header
Hier komen alle libraries die in het programma gebruikt worden
"""
import sys
from PyQt5 import QtWidgets, uic
from functions import maanddag, regelingenophalen, regelingCodeNaam, regelingNaamCode
from flex_keuzes import flexibilisering

from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.backends.backend_qt5agg import NavigationToolbar2QT as NavigationToolbar
import matplotlib.pyplot as plt

"""
Body
Hier komen alle functies
"""
class Functiekeus(QtWidgets.QMainWindow):
    def __init__(self):
        Ui_MainWindow, QtBaseClass = uic.loadUiType("{}\\1AdviseurBeheerder.ui".format(sys.path[0]))
        super(Functiekeus, self).__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.ui.btnAdviseur.clicked.connect(self.btnAdviseurClicked)
        self.ui.btnBeheerder.clicked.connect(self.btnBeheerderClicked)
        
        
    def btnAdviseurClicked(self):
        self.close()
        self._windowdeelnemer = Deelnemerselectie()
        self._windowdeelnemer.show()
    def btnBeheerderClicked(self): 
        self.close()
        self._windowinlog = Inloggen()
        self._windowinlog.show()
        


class Inloggen(QtWidgets.QMainWindow):
    def __init__(self):
        Ui_MainWindow2, QtBaseClass2 = uic.loadUiType("{}\\2InlogBeheerder.ui".format(sys.path[0]))
        super(Inloggen, self).__init__()
        self.ui = Ui_MainWindow2()
        self.ui.setupUi(self)
        self.ui.btnTerug.clicked.connect(self.btnTerugClicked)
        self.ui.btnInloggen.clicked.connect(self.btnInloggenClicked)
        self._Wachtwoord = "wachtwoord"

        
    def btnInloggenClicked(self):
        if self.ui.txtBeheerderscode.text() == self._Wachtwoord:
            self.close()
            self._windowdeelnemer = Deelnemerselectie()
            self._windowdeelnemer.show()
        else:
            self.ui.lblFoutmeldingInlog.setText("Wachtwoord incorrrect")
    def btnTerugClicked(self):
        self.close()
        self._windowkeus = Functiekeus()
        self._windowkeus.show()
        
        

class Deelnemerselectie(QtWidgets.QMainWindow):
    def __init__(self):
        Ui_MainWindow3, QtBaseClass3 = uic.loadUiType("{}\\deelnemerselectie.ui".format(sys.path[0]))
        super(Deelnemerselectie, self).__init__()
        self.ui = Ui_MainWindow3()
        self.ui.setupUi(self)    
        self.ui.btnDeelnemerToevoegen.clicked.connect(self.btnDeelnemerToevoegenClicked)
        self.ui.btnStartFlexibiliseren.clicked.connect(self.btnStartFlexibiliserenClicked)
        self.ui.btnTerug.clicked.connect(self.btnTerugClicked)
        self.ui.sbMaand.valueChanged.connect(self.maandchange)
        self.ui.sbJaar.valueChanged.connect(self.jaarchange)
        
    def btnDeelnemerToevoegenClicked(self):
        self.close()
        self._windowtoevoeg = Deelnemertoevoegen()
        self._windowtoevoeg.show()
        
    def btnStartFlexibiliserenClicked(self):
        self.close()
        self._windowflex = Flexmenu()
        self._windowflex.show()
        
    def btnTerugClicked(self):
        self.close()
        self.windowstart = Functiekeus()
        self.windowstart.show()
        
    def maandchange(self):
        maanddag(self)
                
    def jaarchange(self):
        maanddag(self)
        
        
        
class Deelnemertoevoegen(QtWidgets.QMainWindow):
    def __init__(self):
        Ui_MainWindow4, QtBaseClass4 = uic.loadUiType("{}\\4DeelnemerToevoegen.ui".format(sys.path[0]))
        super(Deelnemertoevoegen, self).__init__()
        self.ui = Ui_MainWindow4()
        self.ui.setupUi(self)
        self.ui.btnTerug.clicked.connect(self.btnTerugClicked)
        self.ui.btnToevoegen.clicked.connect(self.btnToevoegenClicked)
        self.ui.sbMaand.valueChanged.connect(self.maandchange)
        self.ui.sbJaar.valueChanged.connect(self.jaarchange)
        self._30maand = [4,6,9,11]
        
    def btnTerugClicked(self):
        self.close()
        self._windowdeelnemer = Deelnemerselectie()
        self._windowdeelnemer.show()
        
    def btnToevoegenClicked(self):
        if self.ui.txtVoorletters.text() == "" or self.ui.txtAchternaam.text() == "":
            print("Naam gegevens incompleet")
        elif self.ui.txtFulltimeLoon.text() == "" or self.ui.txtParttimePercentage.text() == "":
            print("Loon informatie incompleet")
        else:
            self.close()
            self._windowdeelnemer = Deelnemerselectie()
            self._windowdeelnemer.show()
    
    def maandchange(self):
        maanddag(self)
        
    def jaarchange(self):
        maanddag(self)



class Flexmenu(QtWidgets.QMainWindow):
    def __init__(self):
        Ui_MainWindow5, QtBaseClass5 = uic.loadUiType("{}\\flexmenu.ui".format(sys.path[0]))
        super(Flexmenu, self).__init__()
        
        # Setup van UI
        self.ui = Ui_MainWindow5()
        self.ui.setupUi(self)
        
        # Deelnemer
        self.deelnemer = 4 #Dit moet een variabel worden, het getal is de regel waar de deelnemer staat in het bestand
        
        # Regeling selectie
        self.ui.cbRegeling.addItems(regelingenophalen(self.deelnemer)[0])
    
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
        
    def regelingenObject(self):
        """
        Deze functie maakt voor elke regeling een flexibilisering-object aan 
        uit flex_keuzes.py. Functie checkt ook welke regelingen actief zijn. 
        """
    
        if "ZL" in regelingenophalen(self.deelnemer)[1]:
            self._ZL = flexibilisering("ZL",True)
        else:
            self._ZL = flexibilisering("ZL",False)
            
        if "A65" in regelingenophalen(self.deelnemer)[1]:
            self._A65 = flexibilisering("A65",True)
        else:
            self._A65 = flexibilisering("A65",True)
            
        if "A67" in regelingenophalen(self.deelnemer)[1]:
            self._A67 = flexibilisering("A67",True)
        else:
            self._A67 = flexibilisering("A67",True)
            
        if "NN65" in regelingenophalen(self.deelnemer)[1]:
            self._NN65 = flexibilisering("NN65",True)
        else:
            self._NN65 = flexibilisering("NN65",True)
            
        if "NN67" in regelingenophalen(self.deelnemer)[1]:
            self._NN67 = flexibilisering("NN67",True)
        else:
            self._NN67 = flexibilisering("NN67",True)
            
        if "VLC68" in regelingenophalen(self.deelnemer)[1]:
            self._VLC68 = flexibilisering("VLC68",True)
        else:
            self._VLC68 = flexibilisering("VLC68",False)
        
    def invoerVerandering(self):
        self.regelingCode = regelingNaamCode(str(self.ui.cbRegeling.currentText()))
        
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
        self._windowdeelnemer = Deelnemerselectie()
        self._windowdeelnemer.show()
        
    def btnVergelijkenClicked(self):
        # Sheet van vergelijkingen openen
        self.close()
        
    def btnOpslaanClicked(self): 
        # Alle huidige flexibiliserignen opslaan in een Excel sheet
        # Huidig diagram opslaan en plaats in vergelijking sheet
        self.close()