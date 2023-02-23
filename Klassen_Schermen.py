# -*- coding: utf-8 -*-


import sys
from PyQt5 import QtWidgets, uic
from functions import maanddag

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
        self.ui = Ui_MainWindow5()
        self.ui.setupUi(self)
        self.ui.btnAndereDeelnemer.clicked.connect(self.btnAndereDeelnemerClicked)
        self.ui.btnVergelijken.clicked.connect(self.btnVergelijkenClicked)
        self.ui.btnOpslaan.clicked.connect(self.btnOpslaanClicked)
        
    def btnAndereDeelnemerClicked(self):
        self.close()
        self._windowdeelnemer = Deelnemerselectie()
        self._windowdeelnemer.show()
        
    def btnVergelijkenClicked(self):
        self.close()
        
    def btnOpslaanClicked(self):
        self.close()