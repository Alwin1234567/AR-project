"""
Header
Hier komen alle libraries die in het programma gebruikt worden
"""
import sys
from PyQt5 import QtWidgets, uic
import functions
from datetime import datetime

"""
Body
Hier komen alle functies
"""
class Functiekeus(QtWidgets.QMainWindow):
    def __init__(self, book):
        Ui_MainWindow, QtBaseClass = uic.loadUiType("{}\\1AdviseurBeheerder.ui".format(sys.path[0]))
        super(Functiekeus, self).__init__()
        self.book = book
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.ui.btnAdviseur.clicked.connect(self.btnAdviseurClicked)
        self.ui.btnBeheerder.clicked.connect(self.btnBeheerderClicked)
        
        
    def btnAdviseurClicked(self):
        self.close()
        self._windowdeelnemer = Deelnemerselectie(self.book)
        self._windowdeelnemer.show()
    def btnBeheerderClicked(self): 
        self.close()
        self._windowinlog = Inloggen(self.book)
        self._windowinlog.show()
        


class Inloggen(QtWidgets.QMainWindow):
    def __init__(self, book):
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
            self.close()
            self._windowdeelnemer = Deelnemerselectie(self.book)
            self._windowdeelnemer.show()
        else:
            self.ui.lblFoutmeldingInlog.setText("Wachtwoord incorrrect")
    def btnTerugClicked(self):
        self.close()
        self._windowkeus = Functiekeus(self.book)
        self._windowkeus.show()
        
        

class Deelnemerselectie(QtWidgets.QMainWindow):
    def __init__(self, book):
        Ui_MainWindow3, QtBaseClass3 = uic.loadUiType("{}\\deelnemerselectie.ui".format(sys.path[0]))
        super(Deelnemerselectie, self).__init__()
        self.book = book
        self.ui = Ui_MainWindow3()
        self.ui.setupUi(self)
        self.deelnemersbestand = functions.getDeelnemersbestand(self.book, ["Geboortedatum", "voorletter", "tussenvoegsels", "Naam", "Geslacht"])
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
        
        
    def btnDeelnemerToevoegenClicked(self):
        self.close()
        self._windowtoevoeg = Deelnemertoevoegen(self.book)
        self._windowtoevoeg.show()
        
    def btnStartFlexibiliserenClicked(self):
        self.close()
        self._windowflex = Flexmenu(self.book)
        self._windowflex.show()
        
    def btnTerugClicked(self):
        self.close()
        self.windowstart = Functiekeus(self.book)
        self.windowstart.show()
        
    def onChange(self, datumChange):
        if datumChange: functions.maanddag(self)
        kleindeelnemersbestand = self.deelnemersbestand
        kleindeelnemersbestand = functions.filterkolom(kleindeelnemersbestand, self.ui.txtVoorletters.text(), "voorletter")
        kleindeelnemersbestand = functions.filterkolom(kleindeelnemersbestand, self.ui.txtTussenvoegsel.text(), "tussenvoegsels")
        kleindeelnemersbestand = functions.filterkolom(kleindeelnemersbestand, self.ui.txtAchternaam.text(), "Naam")
        kleindeelnemersbestand = functions.filterkolom(kleindeelnemersbestand, datetime(self.ui.sbJaar.value(), self.ui.sbMaand.value(), self.ui.sbDag.value()), "Geboortedatum")
        kleindeelnemersbestand = functions.filterkolom(kleindeelnemersbestand, self.ui.cbGeslacht.currentText(), "Geslacht")
        self.ui.lwKeuzes.clear()
        for i, naam in enumerate(self.deelnemersbestand[0]):
            if naam == "voorletter": voorletters = i
            elif naam == "tussenvoegsels": tussenvoegsels = i
            elif naam == "Naam": achternaam = i
            elif naam == "Geboortedatum": geboortedatum = i
            elif naam == "Geslacht": geslacht = i
        for rij in kleindeelnemersbestand[1:]:
            weergave = "{} {}".format(rij[voorletters], rij[achternaam])
            if rij[tussenvoegsels] != None: weergave += ", {}".format(rij[tussenvoegsels])
            weergave += " | {} | {}".format(rij[geboortedatum].date(), rij[geslacht])
            self.ui.lwKeuzes.addItem(weergave)
        self.ui.lwKeuzes.repaint()
        
        
        
        
class Deelnemertoevoegen(QtWidgets.QMainWindow):
    def __init__(self, book):
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
        
    def btnTerugClicked(self):
        self.close()
        self._windowdeelnemer = Deelnemerselectie(self.book)
        self._windowdeelnemer.show()
        
    def btnToevoegenClicked(self):
        if self.ui.txtVoorletters.text() == "" or self.ui.txtAchternaam.text() == "":
            print("Naam gegevens incompleet")
        elif self.ui.txtFulltimeLoon.text() == "" or self.ui.txtParttimePercentage.text() == "":
            print("Loon informatie incompleet")
        else:
            self.close()
            self._windowdeelnemer = Deelnemerselectie(self.book)
            self._windowdeelnemer.show()
    
    def onChange(self): functions.maanddag(self)



class Flexmenu(QtWidgets.QMainWindow):
    def __init__(self, book):
        Ui_MainWindow5, QtBaseClass5 = uic.loadUiType("{}\\flexmenu.ui".format(sys.path[0]))
        super(Flexmenu, self).__init__()
        self.book = book
        self.ui = Ui_MainWindow5()
        self.ui.setupUi(self)
        self.ui.btnAndereDeelnemer.clicked.connect(self.btnAndereDeelnemerClicked)
        self.ui.btnVergelijken.clicked.connect(self.btnVergelijkenClicked)
        self.ui.btnOpslaan.clicked.connect(self.btnOpslaanClicked)
        
    def btnAndereDeelnemerClicked(self):
        self.close()
        self._windowdeelnemer = Deelnemerselectie(self.book)
        self._windowdeelnemer.show()
        
    def btnVergelijkenClicked(self):
        self.close()
        
    def btnOpslaanClicked(self):
        self.close()