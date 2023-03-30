"""
Header
Hier komen alle libraries die in het programma gebruikt worden
"""

import xlwings as xw
from xlwings.constants import DVType
from xlwings.utils import rgb_to_int
from datetime import datetime, date
from Deelnemer import Deelnemer
from Pensioenfonds import Pensioenfonds
import ctypes
import logging
import os
from os.path import exists
import sys
from string import ascii_uppercase
import matplotlib.pyplot as plt
from reportlab.lib.units import cm
from pathlib import Path
from io import BytesIO 
from svglib.svglib import svg2rlg



"""
Body
Hier komen alle functies
"""
def wachtwoord():
    '''
    functie waarin het wachtwoord voor het protecten en unprotecten van sheets staat

    Returns
    -------
    wachtwoord

    '''
    wachtwoord = "wachtwoord"
    return wachtwoord

def ProtectBeheer(sheet):
    '''
    functie die de sheet protect als er geen beheerder is ingelogd en niets doet als er een beheerder is ingelogd.

    Parameters
    ----------
    sheet : xlwings.Book.sheets["naam sheet"]
        De excel sheet waarin het programma runned.
    

    Returns
    -------
    None.

    '''
    
    beheerder = isBeheerder(sheet.book)
    if beheerder == False:
        if sheet.name == "Vergelijken":
            sheet.api.Protect(Password = wachtwoord(), Contents=False)
        else:
            sheet.api.Protect(Password = wachtwoord())
    

def isBeheerder(book):
    '''
    functie die controleert of er een beheerder is ingelogd

    Parameters
    ----------
    book : xlwings.Book
        Het excel bestand waarin het programma runned.

    Returns
    -------
    beheerder : bool
        True als beheerder is ingelogd, False als geen beheerder is ingelogd

    '''
    beheerder = book.sheets["Beheerder"].cells(1, 1).value == "Beheerder"
    return beheerder

def getaltotijd(getal):
    """
    Functie die een float met daarin een jaartal en een maand omzet in een string die het mooier weergeeft

    Parameters
    ----------
    getal : Float
        variabele met daarin een jaar en maand.

    Returns
    -------
    tijd : String
        variabele waarin de jaar en maand gescheiden zijn in format XXj XXm.

    """
    jaar = int(getal)
    maand = round((getal - jaar) * 12)
    tijd = "{}j".format(jaar)
    if maand > 0: tijd = tijd + " {}m".format(maand)
    return tijd

def getaltogeld(getal): return "€{:.2f}".format(float(getal)).replace(".",",")

def blokkentellen(beginrij, beginkolom, blokafstand, sheet):
    """
    Een functie die het aantal blokken met OP of PP informatie telt

    Parameters
    ----------
    beginrij : Int
        De rij vanaf waar het moet gaan rekenen.
    beginkolom : Int
        De kolom vanaf waar het moet gaan rekenen.
    blokafstand : Int
        De afstand tussen twee blokken.
    sheet : Book.Sheet Type
        De Sheet waarop de blokken staan.

    Returns
    -------
    blokaantal : Int
        De hoeveelheid blokken die het algoritme geteld heeft.

    """
    blokaantal = 0
    leescell = [beginrij, beginkolom]
    while sheet.range(tuple(leescell)).value != None:
        blokaantal += 1
        leescell[0] +=blokafstand
    return blokaantal

# def kleurinvoer(kleur):
#     """
#     Een functie die een string met rgb waardes veranderd naar een tuple met rgb waardes.

#     Parameters
#     ----------
#     kleur : String
#         Bevat drie rgb waardes in een string gescheiden met een ",".

#     Returns
#     -------
#     tuple(kleuren) : Tuple(List)
#         De drie rgb waardes als integer in een tuple.

#     """
#     rgb = kleur.split(",")
#     kleuren = list()
#     for i in range(len(rgb)):
#         kleuren.append(int(rgb[i])/255)    
#     return tuple(kleuren)

def maanddag(interface):
    """
    Een functie die kijkt naar de maand en het jaar van de op dit moment ingevulde datum.
    En past dan toe dat er geen hogere dag mag worden gekozen dan de maand heeft.
    Zorft bijvoorbeeld dat 31 juni niet kan

    Parameters
    ----------
    interface : object/UI
        Is een object van een user interface uit qtdisigner

    Returns
    -------
    De max van de spinbox waar de dagen voor de datum in getoond worden

    """
    maand30 = [4,6,9,11]
    if interface.ui.sbMaand.value() in maand30:
        interface.ui.sbDag.setMaximum(30)
    elif interface.ui.sbMaand.value() == 2:
        if interface.ui.sbJaar.value()%4 == 0: interface.ui.sbDag.setMaximum(29)
        else: interface.ui.sbDag.setMaximum(28)
    else: interface.ui.sbDag.setMaximum(31)

def regelingenophalen(rij):
    """
    Een functie die de regelingen van een deelnemer opzoekt.

    Parameters
    ----------
    rij : integer
        Het rij nummer waar de deelnemer staat in het deelnemersbestand in Excel.
    
    Returns
    -------
    regelingen: List met de volledige namen van pensioenregelingen van de betreffende deelnemer
    regelingCode: List met de codenamen van pensioenregelingen van de betreffende deelnemer
    """
    
    # Sheet ophalen
    book = xw.Book.caller()
    deelnemersBestand = book.sheets["deelnemersbestand"]
    
    regelingen = []
    regelingCode = []
    
    if deelnemersBestand.cells(rij,"J").value is not None:
        regelingen.append("ZwitserLeven")
        regelingCode.append("ZL")
    if deelnemersBestand.cells(rij,"K").value is not None:
        regelingen.append("Aegon OP65")
        regelingCode.append("Aegon67")
    if deelnemersBestand.cells(rij,"L").value is not None:
        regelingen.append("Aegon OP67")
        regelingCode.append("Aegon67")
    if deelnemersBestand.cells(rij,"M").value is not None:
        regelingen.append("Nationale Nederlanden OP65")
        regelingCode.append("NN65")
    if deelnemersBestand.cells(rij,"O").value is not None:
        regelingen.append("Nationale Nederlanden OP67")
        regelingCode.append("NN67")
    if deelnemersBestand.cells(rij,"Q").value is not None:
        regelingen.append("Pensioenfonds VLC OP68")
        regelingCode.append("PF_VLC68")
    
    return regelingen, regelingCode

def regelingNaamCode(naam):
    """
    Een functie die de volledige regeling naam omzet naar de codenaam.

    Parameters
    ----------
    naam : string
        Volledige naam van de regeling.
    
    Returns
    -------
    code: String met de codenaam van de regeling.
    """
    
    if naam == "ZwitserLeven":
        code = "ZL"
    elif naam == "Aegon OP65":
        code = "Aegon65"
    elif naam == "Aegon OP67":
        code = "Aegon67"
    elif naam == "Nationale Nederlanden OP65":
        code = "NN65"
    elif naam == "Nationale Nederlanden OP67":
        code = "NN67"
    elif naam == "Pensioenfonds VLC OP68":
        code = "PF_VLC68"
    
    return code
    
def regelingCodeNaam(code):
    """
    Een functie die de codenaam van de regeling omzet naar de volledige naam.

    Parameters
    ----------
    code : string
        Codenaam van de regeling.
    
    Returns
    -------
    naam: String met de volledige naam van de regeling.
    """
    
    if code == "ZL":
        naam = "ZwitserLeven"
    elif code == "Aegon65":
        naam = "Aegon OP65"
    elif naam == "Aegon67":
        naam = "Aegon OP67"
    elif code == "NN65":
        naam = "Nationale Nederlanden OP65"
    elif code == "NN67":
        naam = "Nationale Nederlanden OP67"
    elif code == "PF_VLC68":
        naam = "Pensioenfonds VLC OP68"
    
    return naam

def getDeelnemersbestand(book, rij = 0):
    """
    functie die het deelnemersbestand inleest uit excel en dit terug geeft.

    Parameters
    ----------
    book : xlwings.Book
        Het excel bestand waarin het programma runned.
    rij : integer
        specifieke rij die uit het deelnemersbestand gelezen moet worden
        default = 0 (geen specifieke rij nodig - alles uitgelezen)
    
    Returns
    -------
    deelnemerlijst : list(Deelnemer)
        lijst met daarin Deelnemer objecten van alle deelnemers.

    """
    deelnemersbestand = book.sheets["deelnemersbestand"].range((1,1)).expand("down").expand("right").value
    
    deelnemersbestand[0].append("rijNr")    #kolom rijnummer toevoegen 
    for i in range(len(deelnemersbestand) - 1):
        deelnemersbestand[i+1].append(i + 2)    #rijnummer per deelnemer toevoegen
    pensioeninformatie = getPensioeninformatie(book) 
    deelnemerlijst = list()
    for deelnemer in deelnemersbestand[1:]:
        informatie = [deelnemersbestand[0], deelnemer]
        deelnemerlijst.append(Deelnemer(informatie, pensioeninformatie))
    
    if rij == 0: 
        return deelnemerlijst
    elif rij != 0:  #er is een rijnummer meegegeven
        return deelnemerlijst[rij-2]


def getPensioeninformatie(book):
    """
    functie die de informatie van de pensioenen verzameld

    Parameters
    ----------
    book : xlwings.Book
        Het excel bestand waarin het programma runned.

    Returns
    -------
    pensioeninformatie : list(Pensioenfonds)
        lijst met de pensioenfonds objecten van alle verschillende pensioenen.

    """    
    kolommen = dict()
    kolommen["naamkolom"] = 2
    kolommen["soortRegeling"] = 3
    kolommen["pensioenleeftijdkolom"] = 4
    kolommen["rentekolom"] = 5
    kolommen["sterftetafelkolom"] = 6
    kolommen["opbouwpercentage"] = 7
    kolommen["franchise"] = 8
    kolommen["opmerking"] = 9
    kolommen["kleurzachtkolom"] = 10
    kolommen["kleurhardkolom"] = 11
    
    gegevens_pensioenenSheet = book.sheets["Gegevens pensioencontracten"]
    
    pensioenen = dict()
    pensioenen["AOW"] = ((None, None), 3)
    pensioenen["ZL"] = ((9, None), 4)
    pensioenen["Aegon65"] = ((10, None), 5)
    pensioenen["Aegon67"] = ((11, None), 6)
    pensioenen["NN65"] = ((12, 13), 7)
    pensioenen["NN67"] = ((14, 15), 8)
    pensioenen["PF_VLC68"] = ((16, 17), 9)
    
    pensioeninformatie = list()
    for pensioen in pensioenen.values():
        pensioeninformatie.append(Pensioenfonds(gegevens_pensioenenSheet, kolommen, pensioen))
    return pensioeninformatie


def filterkolom(deelnemerlijst, zoekterm, attribuutnaam):
    """
    functie die door de kolommen filterd om te kijken of ze aan de zoekterm voldoen.

    Parameters
    ----------
    deelnemerlijst : list(Deelnemer)
        Lijst met daarin de deelneemers waarover gefilterd moet worden
    zoekterm : None of str of datetime
        de waarde waarop gefilterd moet worden.
    attribuutnaam : str
        De naam van de attribuut waarop wordt gezocht.

    Returns
    -------
    kleinDeelnemerlijst : list(Deelnemer)
        gefilterde versie van het deelnemersbestand.

    """
    if zoekterm == "" or zoekterm == datetime(1950, 1, 1): return deelnemerlijst
    kleinDeelnemerlijst = list()
    if type(zoekterm) == str:
        for deelnemer in deelnemerlijst:
            attribuut = getattr(deelnemer, attribuutnaam)
            if attribuut == None: pass
            elif len(attribuut) < len(zoekterm): pass
            else:
                if attribuut[0:len(zoekterm)] == zoekterm: kleinDeelnemerlijst.append(deelnemer)
    elif type(zoekterm) == datetime:
        for deelnemer in deelnemerlijst:
            attribuut = getattr(deelnemer, attribuutnaam)
            if attribuut == zoekterm: kleinDeelnemerlijst.append(deelnemer)
    return kleinDeelnemerlijst


def pensioensdatum():
    """
    Functie die het geboortejaar en maand van mensen op pensioensleeftijd uitrekent en teruggeeft
    input : -
    output : geboortejaar en maand van mensen op pensioensleeftijd
    """
    
    huidigeDatum = datetime.today()
    #huidige pensioensleeftijd 67 jaar en 3 maanden
    pensioenJaar = 67
    pensioenMaand = 3
    #huidige datum opsplitsen in jaar en maand
    nieuweJaar = huidigeDatum.year
    nieuweMaand = huidigeDatum.month
    if nieuweMaand <= pensioenMaand:
        nieuweJaar = nieuweJaar - pensioenJaar - 1
        nieuweMaand = 12 - (pensioenMaand - nieuweMaand)
    else: 
        nieuweJaar = huidigeDatum.year - pensioenJaar
        
    pensioensdatum = datetime(nieuweJaar, nieuweMaand, 1)
    pensioensdatum = pensioensdatum.strftime("%m-%Y")
    
    return pensioensdatum

def isfloat(string):
    """
    Functie die controleert of een gegeven string een float is.
    input : string
    output: boolean (True of False)
    """
    
    stringNum = string.replace(",", "").replace(".", "") #verwijder komma's en punten
    resultaat = stringNum.isdigit() #check of de string (zonder punten en komma's) een getal is.
    
    return resultaat


def DeelnemerVinden(book, persoonsgegevens):
    """
    functie die een deelnemer kan vinden in het deelnemersbestand

    Parameters
    ----------
    book : xlwings.Book
        Het excel bestand waarin het programma runned.
    persoonsgegevens : list(persoonsgegevens)
        bevat [achternaam, tussenvoegsel, voorletters, geboortedatum, geslacht] mogelijk met meer erachteraan

    Returns
    -------
    deelnemerlijst : list(Deelnemer)
        gefilterde versie van het deelnemersbestand met de deelnemer met exact dezelfde persoonsgegevens

    """
    deelnemerlijst = getDeelnemersbestand(book)
    attributen = ["achternaam", "tussenvoegsels", "voorletters", "geboortedatum", "geslacht"]
    #persoonsgegevens[3] = datetime(persoonsgegevens[3]).strftime("%d-%m-%Y")
    for i in [0,1,2,3,4]:
        deelnemerlijst = filterkolom(deelnemerlijst, persoonsgegevens[i], attributen[i])

    return deelnemerlijst


def ToevoegenDeelnemer(gegevens, regel = 0):
    """
    Functie die een deelnemer toevoegt onderaan het deelnemersbestand
    input : lijst met gegevens van de deelnemer
    output : deelnemer toegevoegd aan excel-bestand
    """
    
    book = xw.Book.caller()
    deelnemersbestand = book.sheets["deelnemersbestand"]
    
    if regel == 0:   #geen specifieke regel meegegeven
        #check wat eerstvolgende lege regel is
        #bereken het aantal deelnemers door het aantal volle rijen na 1e regel te tellen
        aantalDeelnemers = len(deelnemersbestand.cells(1,1).expand().value)
        regel = aantalDeelnemers + 1
    #sheet unprotecten
    deelnemersbestand.api.Unprotect(Password = wachtwoord())
    #gegevens deelnemer invullen in de lege regel
    deelnemersbestand.cells(regel, 1).value = gegevens
    #sheet protecten
    ProtectBeheer(deelnemersbestand) #.api.Protect(Password = wachtwoord())

def Mbox(title, text, style):
    """
    functie die een messagebox maakt

    Parameters
    ----------
    title : string
        Wordt de titel van de messagebox
    text : string
        bevat de tekst die in de messagebox moet komen
    style : integer
        Geeft aan welke knoppen er op de messagebox komen
        ##  Styles:
        ##  0 : OK
        ##  1 : OK | Cancel
        ##  2 : Abort | Retry | Ignore
        ##  3 : Yes | No | Cancel
        ##  4 : Yes | No
        ##  5 : Retry | Cancel 
        ##  6 : Cancel | Try Again | Continue

   
    Returns
    -------
    str
        string die aangeeft welke op knop is geklikt als reactie op de messagebox.

    """
    returnValue = ctypes.windll.user32.MessageBoxW(0, text, title, style)
    if returnValue == 0:
        raise Exception('fout in messagebox')
    #controleren op welke knop gedrukt is
    elif returnValue == 1: #OK
        return "OK Clicked"
    elif returnValue == 2: #Cancel
        return "Cancel Clicked"
    elif returnValue == 3: #Abort
        return "Abort Clicked"
    elif returnValue == 4: #Retry
        return "Retry Clicked"
    elif returnValue == 5: #Ignore
        return "Ignore Clicked"
    elif returnValue == 6: #Yes
        return "Ja"
    elif returnValue == 7: #No
        return "Nee"
    

def gegevenscontrole(gegevenslijst):
    """
    functie die een messagebox met alle gegevens van een deelnemer maakt

    Parameters
    ----------
    gegevenslijst : list
        een lijst met alle gegevens van de deelnemer in de vorm:
            ["Achternaam", "tussen", "voor", "geboorte", "geslacht", "burg", "fulltimeLoon", "PT% als kans", "regeling", "Zl", "Aegon", "Aegon", "NN", "NN", "NN", "NN", "VLC", "VLC"]

    Returns
    -------
    string "correct" of "fout"
        Bij "correct" zijn de gegevens goed, bij "fout" zijn niet alle gegevens goed

    """
    #alle gegevens omzetten naar een string
    for g in range(0,len(gegevenslijst)):
        gegevenslijst[g] = str(gegevenslijst[g])
        
    invoer = [] #lijst met alle deelnemersgegevens met uitleg in juiste volgorde
    invoer.append("Naam: " + gegevenslijst[2] + " " + gegevenslijst[1] + " " + gegevenslijst[0])
    #geboortedatum van maand-dag-jaar naar dag-maand-jaar notatie
    datumSplit = gegevenslijst[3].split("-")
    geboortedatum = datumSplit[1] + "-" +  datumSplit[0] + "-" +  datumSplit[2]
    invoer.append("Geboortedatum: " + geboortedatum)
    invoer.append("Geslacht: " + gegevenslijst[4])
    invoer.append("Burgerlijke staat: " + gegevenslijst[5])
    invoer.append("Fulltime loon: €" + gegevenslijst[6])
    invoer.append("Parttime percentage: " + gegevenslijst[7] + "%")
    invoer.append("Huidige regeling: " + gegevenslijst[8])
    invoer.append("\n")     #lege regel tussenvoegen voor opgebouwde pensioenen
    
    #opgebouwde pensioenen toevoegen, als deze ingevuld zijn
    teller = 0  #bijhouden welk pensioen het is
    for p1 in gegevenslijst[9:12]:
        if p1 != "":
            regeling = ["ZL: ", "Aegon 65: ", "Aegon 67: "][teller]
            invoer.append(regeling + "OP = €" + p1)
        teller += 1
    
    teller = 0
    for p2 in range(12,17,2):
        if gegevenslijst[p2] != "":
            regeling = ["NN 65: ", "NN 67: ", "PF VLC 68: "][teller]
            invoer.append(regeling + "OP = €" + gegevenslijst[p2] + " en PP = €" + gegevenslijst[p2+1])
        teller += 1
    
    #alle gegevens met uitleg op een nieuwe regel in een string
    message = "Uw gegevens: "        
    for i in invoer:
        message = message + i + "\n"
    message = message + "\nKloppen bovenstaande gegevens?"
    
    #messagebox tonen
    controle = Mbox("Gegevenscontrole", message, 4)
    if controle == "Ja":
        return "correct"
    else:
        return "fout"

def setup_logger(name):
    """
    functie die de logger met handlers maakt

    Parameters
    ----------
    naam : str
        naam van de logger.

    Returns
    -------
    logger : Logger
        Het logger object dat gebruikt wordt.

    """
    logger = logging.getLogger(name)

    logger.setLevel(logging.DEBUG)
    today = date.today().strftime("%Y_%m_%d")
    if not exists("{}\\Logs\\{}.log".format(krijgpad(), today)): os.makedirs(os.path.dirname("{}\\Logs\\{}.log".format(krijgpad(), today)), exist_ok=True)
    if not exists("{}\\Logs\\Errors\\{}.log".format(krijgpad(), today)): os.makedirs(os.path.dirname("{}\\Logs\\Errors\\{}.log".format(krijgpad(), today)), exist_ok=True)
    filename = "{}\\Logs\\{}.log".format(krijgpad(), today)
    errorname = "{}\\Logs\\Errors\\{}.log".format(krijgpad(), today)
    
    chat_logger = logging.StreamHandler()
    file_logger = logging.FileHandler(filename)
    error_logger = logging.FileHandler(errorname)
    
    chat_logger.setLevel(logging.WARNING)
    file_logger.setLevel(logging.INFO)
    error_logger.setLevel(logging.ERROR)

    chat_logger.setFormatter(logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s"))
    file_logger.setFormatter(logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s"))
    error_logger.setFormatter(logging.Formatter("%(asctime)s - %(name)s - %(message)s"))

    logger.addHandler(chat_logger)
    logger.addHandler(file_logger)
    logger.addHandler(error_logger)
    logger.info("Setup logger is done")
    return logger

def isNotFloat(veldInput):
    try:
        veldInput = float(veldInput)
        return False
    except ValueError:
        return True
        
def checkVeldInvoer(soort,methode,veld1,veld2,veld3):
    intProblem = False
    emptyProblem = False
    message = ""
    OK = True
    
    if str(methode) == "Percentage" or str(methode) == "Verschil":
        if str(veld1) == "": emptyProblem = True
        elif isNotFloat(veld1): intProblem = True
        
        if str(veld2) == "": pass
        elif isNotFloat(veld2): intProblem = True
            
        if str(veld3) == "": pass
        elif isNotFloat(veld3): intProblem = True  
        
    elif str(methode) == "Verhouding":
        if str(veld1) == "": pass
        elif isNotFloat(veld1): intProblem = True

        if str(veld2) == "": emptyProblem = True
        elif isNotFloat(veld2): intProblem = True
            
        if str(veld3) == "": emptyProblem = True
        elif isNotFloat(veld3): intProblem = True
    
    elif str(methode) == "Opvullen AOW":
        if isNotFloat(veld1): intProblem = True

        if isNotFloat(veld2): intProblem = True

        if isNotFloat(veld3): intProblem = True

    if intProblem == True and emptyProblem == True: # Er zijn letters ingevuld & er zijn lege vakjes
        return ["Er is foute invoer en missende invoer.",False]
    elif intProblem == True and emptyProblem == False: # Er zijn letters ingevuld
        return ["Invoer mag alleen een positief getal zijn.",False]
    elif intProblem == False and emptyProblem == True: # Er zijn lege vakjes
        return ["Er is missende invoer.",False]
    elif intProblem == False and emptyProblem == False: # Alle vakjes zijn met gehele getallen ingevuld
        if soort == "OP-PP":
            if str(methode) == "Verhouding":
                if float(veld2) < 0 or float(veld3) < 0: # Getallen mogen niet negatief zijn
                    message = "Getallen mogen niet negatief zijn."
                    OK = False
                elif float(veld3)/float(veld2) > 0.70: # Verhouding moet voldoen aan PP max. 70% van OP regel
                    message = "Verhouding ongeldig (PP maximaal 70% van OP)"
                    OK = False
            elif str(methode) == "Percentage":
                if float(veld1) < 0: # Getallen mogen niet negatief zijn
                    message = "Getallen mogen niet negatief zijn."
                    OK = False
                elif float(veld1) > 100: # Percentage kan niet hoger dan 100% zijn
                    message = "Percentage ongeldig (kan niet hoger dan 100%)"
                    OK = False
        elif soort == "hoog-laag":
            if str(methode) == "Verhouding":
                if float(veld2) < 0 and int(veld3) < 0: # Getallen mogen niet negatief zijn
                    message = "Getallen mogen niet negatief zijn."
                    OK = False
                elif float(veld3)/float(veld2) < 0.75 or float(veld3)/float(veld2) > 1: # Verhouding moet voldoen aan hoog-laag 4:3 regel
                    message = "Verhouding ongeldig (3:4 regel)"
                    OK = False
            elif str(methode) == "Verschil":
                if float(veld1) < 0: # Getallen mogen niet negatief zijn
                    message = "Getallen mogen niet negatief zijn."
                    OK = False
            
        return [message,OK]
    
# def tpxFormule(sterftetafel, rij, leeftijdKolomLetter, jaarKolom, tpxKolom):
#     if sterftetafel == "AG_2020": return '=if({0}{1}<>"", (1-INDEX(INDIRECT("{2}"),{0}{3}+1,{4}{3}-2018))*{5}{3},"")'.format(leeftijdKolomLetter, rij + 3,  sterftetafel, rij + 2, jaarKolom, tpxKolom)
#     else: return '=if({0}{1}<>"", INDEX(INDIRECT("{2}"),{0}{1}+1,1)/ INDEX(INDIRECT("{2}"),${0}$2+1,1),"")'.format(leeftijdKolomLetter, rij + 3, sterftetafel)

def persoonOpslag(sheet, persoonObject):
    """
    Functie die persoonsgegevens opslaat in de flexopslag sheet.
    
    Parameters
    ----------
    sheet : xlwings.Book.sheets["sheetname"]
        sheet van excel bestand waarin het programma runned.
    
    persoonObject : Python-object
        Object waaruit gegevens gehaald worden.
    
    Returns
    -------
    None : None
        Het plakt alle gegevens in de Excel sheet.
    """
   
    persopslag = []
    
    for i in range(10):
        persopslag.append(["",""])
    
    persopslag[0][0] = "Voorletters"
    persopslag[0][1] = str(persoonObject.voorletters)
    
    persopslag[1][0] = "Tussenvoegsels"
    persopslag[1][1] = str(persoonObject.tussenvoegsels)
    
    persopslag[2][0] = "Achternaam"
    persopslag[2][1] = str(persoonObject.achternaam)
    
    persopslag[3][0] = "Geboortedatum"
    persopslag[3][1] = persoonObject.geboortedatum
    
    persopslag[4][0] = "Geslacht"
    persopslag[4][1] = str(persoonObject.geslacht)
    
    persopslag[5][0] = "Burg. staat"
    persopslag[5][1] = str(persoonObject.burgelijkeStaat)
    
    persopslag[6][0] = "FT loon"
    persopslag[6][1] = persoonObject.ftLoon
    
    persopslag[7][0] = "PT%"
    persopslag[7][1] = persoonObject.pt
    
    persopslag[8][0] = "Regeling"
    persopslag[8][1] = str(persoonObject.regeling)
    
    persopslag[9][0] = "Rij nr"
    persopslag[9][1] = persoonObject.rijNr
    
    #sheet unprotecten
    sheet.api.Unprotect(Password = wachtwoord())
    #persoonopslag invoegen
    sheet.range((6,1),(15,2)).options(ndims = 2).value = persopslag
    sheet.range((6,1),(15,2)).color = (150,150,150)
    #sheet protecten
    ProtectBeheer(sheet)

def flexOpslag(sheet,flexibilisatie,countOpslaan,countRegeling):
    """
    Functie waar een lege 2D lijst wordt gecreëerd om flexibilisaties in op te slaan.
    Deze lijst moet vervolgens in de Flexopslag sheet geplakt worden.
    
    Parameters
    ----------
    book : xlwings.Book
        Het excel bestand waarin het programma runned.
    
    flexibilisatie: Python-object
        Het regeling-specifieke object met flexibilisatiekeuzes
    
    countOpslaan : integer
        Aantal eerder ogeslagen flexbilisaties (verplaatst deze flexibilisatie met 
                                                een aantal stappen naar rechts in de sheet)
    
    countRegeling: integer
        Hoeveelste regeling waarvoor flexibilisaties worden opgeslagen (verplaatst deze flexibilisatie met
                                                                        een aantal stappen omlaag in de sheet)
    
    Returns
    -------
    None : None
        Het plakt alle flex keuzes in de Excel sheet
    """
    
    
    flexopslag = []
    
    for i in range(19):
        flexopslag.append(["","",""])
    
    flexopslag[0][0] = "Pensioenfonds"
    
    flexopslag[2][0] = "Wijzigen"
    flexopslag[3][0] = "Pensioenleeftijd"
    
    flexopslag[5][0] = "Uitruilen"
    flexopslag[6][0] = "Volgorde"
    flexopslag[7][0] = "Methode"
    flexopslag[8][0] = "Verschil/verhouding"
    
    flexopslag[10][0] = "Hoog/Laag"
    flexopslag[11][0] = "Volgorde"
    flexopslag[12][0] = "Duur"
    flexopslag[13][0] = "Methode"
    flexopslag[14][0] = "Vers/Verh/Opv"
    
    flexopslag[16][0] = "OP H|L"
    flexopslag[17][0] = "PP"
    
    flexopslag[18][0] = "Kleur"
    
    # Pensioennaam invullen
    flexopslag[0][1] = str(flexibilisatie.pensioen.pensioenNaam)
    
    # Pensioenleeftijd wijzigen J/N
    if flexibilisatie.leeftijd_Actief: flexopslag[2][1] = "Ja"
    else: flexopslag[2][1] = "Nee"
    
    # Pensioenleeftijd: Jaar & Maand
    if flexibilisatie.HL_Methode == "Opvullen AOW":
        flexopslag[2][1] = "Ja"
        flexopslag[3][1] = flexibilisatie.AOWJaar
        flexopslag[3][2] = flexibilisatie.AOWMaand
    else:
        flexopslag[3][1] = flexibilisatie.leeftijdJaar
        flexopslag[3][2] = flexibilisatie.leeftijdMaand
    
    # OP/PP Uitruilen wijzigen J/N
    if flexibilisatie.OP_PP_Actief: flexopslag[5][1] = "Ja"
    else: flexopslag[5][1] = "Nee"
    
    # OP/PP uitruiling opslaan
    flexopslag[6][1] = flexibilisatie.OP_PP_UitruilenVan
    #methode opslaan
    flexopslag[7][1] = flexibilisatie.OP_PP_Methode
    if flexibilisatie.OP_PP_Methode == "Verhouding":
        flexopslag[8][1] = flexibilisatie.OP_PP_Verhouding_OP
        flexopslag[8][2] = flexibilisatie.OP_PP_Verhouding_PP
    elif flexibilisatie.OP_PP_Methode == "Percentage":
        flexopslag[8][1] = flexibilisatie.OP_PP_Percentage
        if flexibilisatie.OP_PP_Percentage > flexibilisatie.OP_PP_PercentageMax:
            flexopslag[8][2] = flexibilisatie.OP_PP_PercentageMax
    #else: logger.info("OP/PP methode wordt niet herkend bij opslaan naar excel.")
    
    # Hoog/Laag constructie opslaan
    if flexibilisatie.HL_Actief: flexopslag[10][1] = "Ja"
    else: flexopslag[10][1] = "Nee"
    
    flexopslag[11][1] = flexibilisatie.HL_Volgorde
    
    flexopslag[12][1] = flexibilisatie.HL_Jaar
    
    flexopslag[13][1] = flexibilisatie.HL_Methode
    if flexibilisatie.HL_Methode == "Verhouding":
        flexopslag[14][1] = flexibilisatie.HL_Verhouding_Hoog
        flexopslag[14][2] = flexibilisatie.HL_Verhouding_Laag
    elif flexibilisatie.HL_Methode == "Verschil":
        flexopslag[14][1] = flexibilisatie.HL_Verschil
        
        if flexibilisatie.HL_Verschil > flexibilisatie.HL_VerschilMax:
            flexopslag[14][2] = flexibilisatie.HL_VerschilMax
    #else: logger.info("H/L methode wordt niet herkend bij opslaan naar excel.")
    
    # Nieuwe OP en PP opslaan
    flexopslag[16][1] = flexibilisatie.ouderdomsPensioenHoog
    flexopslag[16][2] = flexibilisatie.ouderdomsPensioenLaag
    flexopslag[17][1] = flexibilisatie.partnerPensioen
    
    # RGB opslaan
    flexopslag[18][1] = str(flexibilisatie.pensioen.pensioenKleurZacht)
    
    #sheet unprotecten
    sheet.api.Unprotect(Password = wachtwoord())
    # Waardes in sheet plakken & celkleur instellen
    sheet.range((5+20*countRegeling,4+4*countOpslaan),(23+20*countRegeling,6+4*countOpslaan)).options(ndims = 2).value = flexopslag
    sheet.range((5+20*countRegeling,4+4*countOpslaan),(23+20*countRegeling,6+4*countOpslaan)).color = flexibilisatie.pensioen.pensioenKleurZacht
    #sheet protecten
    ProtectBeheer(sheet) #.api.Protect(Password = wachtwoord())
    
    # # Pensioenleeftijd wijzigen J/N
    # if flexibilisatie.leeftijd_Actief: flexopslag[2][1] = "J"
    # else: flexopslag[2][1] = "N"
    
    # # Pensioenleeftijd: Jaar & Maand
    # flexopslag[3][1] = flexibilisatie.leeftijdJaar
    # flexopslag[3][2] = flexibilisatie.leeftijdMaand
    
    # # OP/PP Uitruilen wijzigen J/N
    # if flexibilisatie.OP_PP_Actief: flexopslag[5][1] = "J"
    # else: flexopslag[5][1] = "N"
    
    # # OP/PP uitruiling opslaan
    # if flexibilisatie.OP_PP_UitruilenVan == "OP naar PP": flexopslag[6][1] = "OP/PP"
    # elif flexibilisatie.OP_PP_UitruilenVan == "PP naar OP": flexopslag[6][1] = "PP/OP"
    
    # if flexibilisatie.OP_PP_Methode == "Verhouding":
    #     flexopslag[7][1] = "Verh"
    #     flexopslag[8][1] = flexibilisatie.OP_PP_Verhouding_OP
    #     flexopslag[8][2] = flexibilisatie.OP_PP_Verhouding_PP
    # elif flexibilisatie.OP_PP_Methode == "Percentage":
    #     flexopslag[7][1] = "Perc"
    #     flexopslag[8][1] = flexibilisatie.OP_PP_Percentage
    # else:
    #     logger.info("OP/PP methode wordt niet herkend bij opslaan naar excel.")
    
    # # Hoog/Laag constructie opslaan
    # if flexibilisatie.HL_Actief: flexopslag[9][1] = "J"
    # else: flexopslag[10][1] = "N"
    
    # if flexibilisatie.HL_Volgorde == "Hoog-laag": flexopslag[11][1] = "Hoog/Laag"
    # elif flexibilisatie.HL_Volgorde == "Laag-hoog": flexopslag[11][1] = "Laag/Hoog"
    
    # flexopslag[12][1] = flexibilisatie.HL_Jaar
    
    # if flexibilisatie.HL_Methode == "Verhouding":
    #     flexopslag[13][1] = "Verh"
    #     flexopslag[14][1] = flexibilisatie.HL_Verhouding_Hoog
    #     flexopslag[15][2] = flexibilisatie.HL_Verhouding_Laag
    # elif flexibilisatie.HL_Methode == "Verschil":
    #     flexopslag[13][1] = "Verh"
    #     flexopslag[14][1] = flexibilisatie.HL_Verschil
    # elif flexibilisatie.HL_Methode == "Opvullen AOW":
    #     flexopslag[13][1] = "Opv"
    # else:
    #     logger.info("H/L methode wordt niet herkend bij opslaan naar excel.")

def FlexopslagVinden(book, naamFlex = "Geen"):
    '''
    functie die telt hoeveel flexibilisaties opgeslagen zijn, 
    hoeveel pensioenen deze persoon heeft
    en op welke plek de flexibilisatie met als naam naamFlex zit

    Parameters
    ----------
    book : xlwings.Book
        Het excel bestand waarin het programma runned.
    naamFlex : string
        naam van de flexibilisatie die gezocht moet worden.

    Returns
    -------
    list
        flexKolom = kolom waarop naamFlex zit
        zoekKolom -4 = kolom met laatste flexibilisatie
        aantalPensionenen = aantal pensioenen van deze deelnemer

    '''
    #sheet definiëren
    flexopslag = book.sheets["Flexopslag"]
    
    flexKolom = 0
    aantalPensioenen = 0
    #startkolom voor het zoeken van flexibilisatie
    zoekKolom = 5
    #alle blokken langsgaan op zoek naar flexibilisatie met naam naamFlex
    while str(flexopslag.cells(2,zoekKolom).value) != "None":
        naam = str(flexopslag.cells(2,zoekKolom).value)
        if naam == naamFlex:
            flexKolom = zoekKolom
        zoekKolom += 4
    if flexKolom != 0:
        #opzoeken hoeveel pensioenen deze deelnemer heeft
        aantalPensioenen = blokkentellen(5, flexKolom, 20, flexopslag)
    return [flexKolom, zoekKolom-4, aantalPensioenen]

#zelfde als functie opslagLegen
# def flexopslagLegen(book):
#     #sheet unprotecten
#     #book.sheets["Flexopslag"].api.Unprotect(Password = wachtwoord())
#     #opgeslagen flexibilisaties van vorige deelnemer verwijderen uit opslag
#     book.sheets["Flexopslag"].clear()
#     #sheet protecten
#     #ProtectBeheer(book.sheets["Flexopslag"]) #.api.Protect(Password = wachtwoord())
    
#     #laatste opslag is verwijderd, dus drop down legen
#     book.sheets["Vergelijken"]["B6"].value = ""

def UitlezenFlexopslag(book, naamFlex):
    """
    functie om in de flexopslag de flexibilisatie met de naam "naamFlex" te vinden
    en deze gegevens op te halen

    Parameters
    ----------
    book : xlwings.Book
        Het excel bestand waarin het programma runned.
    naamFlex : string
        naam van de flexibilisatie die gezocht moet worden.

    Returns
    -------
    flexgegevens : list(pensioen)
        lijst met de pensioenen (ook list).
        pensioen is lijst met de fexibilisatiegegevens

    """
    #sheet definiëren
    flexopslag = book.sheets["Flexopslag"]
    
    opslag = FlexopslagVinden(book, naamFlex)
    flexKolom = opslag[0]
    aantalPensioenen = opslag[2]
    
    flexgegevens = []
    rij = 0
    while rij < aantalPensioenen*20:
        #lijst met gegevens van 1 pensioen aanmaken
        #pensioen = [0-pensioenfonds, 1-wijzigen, 2-leeftijd-jaar, 3-leeftijd-maand, 4-uitruilen, 5-volgorde, 6-methode,
        #7-verhouding/percentage, 8-verhouding/max, 9-hoog/laag, 10-volgorde, 11-duur, 12-methode, 13-vers/verh/opvullen, 14-verhouding/max, 15-OP, 16-PP, 17-kleur]
        pensioen = []
        rijAdd = [5,7,8,8,10,11,12,13,13,15,16,17,18,19,19,21,21,22,23]
        kolomAdd = [0,0,0,1,0,0,0,0,1,0,0,0,0,0,1,0,1,0,0]
        for i in range(19):
            pensioen.append(str(flexopslag.cells(rij+rijAdd[i] ,flexKolom+kolomAdd[i]).value))
            #pensioensgegevens toevoegen aan lijst met totale flexibilisatiegegevens
        flexgegevens.append(pensioen)
        #rij ophogen met 20 -> naar volgende blok
        rij += 20
    
    return flexgegevens

def flexopslagNaamNaarID(book, naamFlex):
    """
    functie om in de flexopslag de flexibilisatie met de naam "naamFlex" te vinden
    en hiervan de ID uit te lezen

    Parameters
    ----------
    book : xlwings.Book
        Het excel bestand waarin het programma runned.
    naamFlex : string
        naam van de flexibilisatie die gezocht moet worden.

    Returns
    -------
    ID : string
        ID van de afbeelding van de flexibilisatie met naam "naamFlex"

    """
    #sheet definiëren
    flexopslag = book.sheets["Flexopslag"]
    
    opslag = FlexopslagVinden(book, naamFlex)
    flexKolom = opslag[0]
    
    ID = flexopslag.cells(3,flexKolom).value
    if type(ID) == int: ID = int(ID)
    else: ID = str(ID)
    return ID

# def zoekRGB(book,regeling):
#     i = 1
#     rgb = "Geen rgb gevonden."
    
#     while i < 11:
#         if str(book.sheets["Gegevens pensioencontracten"].range(i,2).value) == regeling:
#             rgb = str(book.sheets["Gegevens pensioencontracten"].range(i,10).value)
#         i += 1
    
#     return rgb
    
def berekeningen_init(sheet, deelnemer, logger):
    """
    Functie die het berekenings sheet klaar zet voor de deelnemer om mee te flexibiliseren

    Parameters
    ----------
    sheet : xw.Book.Sheet
        Sheet "Berekeningen" waar de berekeningen op worden gedaan.
    deelnemer : Deelnemer
        het deelnemer object van de deelnemer waarmee geflexibiliseerd wordt.
    logger : Logger
        Het logger object om informatie te loggen.

    Returns
    -------
    None.

    """
    #sheet unprotecten
    #sheet.api.Unprotect(Password = wachtwoord())
    
    logger.info("start berekenscherm init")
    # verkrijg berekeningen instellingen
    aantalpensioenen = len(deelnemer.flexibilisaties)
    instellingen = berekeningen_instellingen()
    
    # clear sheet
    sheet.clear()
    
    # pensioen info
    infotitel = ["Regeling", "OP_H", "OP_L", "PP", "Pensioenleeftijd"]
    sheet.range((max(instellingen["pensioeninfohoogte"] - 1, 1), instellingen["pensioeninfokolom"]),\
                (max(instellingen["pensioeninfohoogte"] - 1, 1), instellingen["pensioeninfokolom"] + 3)).value = infotitel
    for i, flexibilisatie in enumerate(deelnemer.flexibilisaties):
        blokhoogte = instellingen["pensioeninfohoogte"] + instellingen["afstandtotblokken"] + aantalpensioenen + i * (instellingen["blokgrootte"] + instellingen["afstandtussenblokken"])
        pensioeninfo = list()
        pensioeninfo.append(flexibilisatie.pensioen.pensioenVolNaam)
        pensioeninfo.append('=B{}'.format(blokhoogte + 10))
        pensioeninfo.append('=IF(B{0} = "", "", C{1})'.format(blokhoogte + 4, blokhoogte + 10))
        pensioeninfo.append('=C{}'.format(blokhoogte + 9))
        pensioeninfo.append(flexibilisatie.pensioen.pensioenleeftijd)
        inforange = sheet.range((instellingen["pensioeninfohoogte"] + i, instellingen["pensioeninfokolom"]),\
                            (instellingen["pensioeninfohoogte"] + i, instellingen["pensioeninfokolom"] + len(pensioeninfo) - 1))
        inforange.formula = pensioeninfo
        inforange.color = flexibilisatie.pensioen.pensioenKleurZacht
    
    # pensioen blok
    for i, flexibilisatie in enumerate(deelnemer.flexibilisaties):
        blokhoogte = instellingen["pensioeninfohoogte"] + instellingen["afstandtotblokken"] + aantalpensioenen + i * (instellingen["blokgrootte"] + instellingen["afstandtussenblokken"])
        rekenblokstart = instellingen["afstandtotrekenkolom"] + i * (instellingen["rekenblokbreedte"] + instellingen["afstandtussenrekenblokken"])
        blok = list()
        blok.append(["Naam", flexibilisatie.pensioen.pensioenVolNaam, "", ""])
        if flexibilisatie.leeftijd_Actief: blok.append(["Start Pensioenjaar", flexibilisatie.leeftijdJaarMaand, "", ""])
        else: blok.append(["Start Pensioenjaar", flexibilisatie.pensioen.pensioenleeftijd, "", ""])
        
        blok.append(["Uitruilen soort", "", "", ""])
        blok.append(["Uitruilen waarde", "", "", '=IF(B{0} = "OP naar PP Percentage", (0.7 * B{1}  - C{1})  / ((0.7 + B{2} / B{3}) * B{1}), "")'.format(blokhoogte + 2, blokhoogte + 8, blokhoogte + 12, blokhoogte + 14)])
        
        blok.append(["Hoog Laag", "", "", ""])
        blok.append(["Hoog Laag waarde", "", "", '=IF(B{0} = "Verschil", IF(C{0} = "Hoog-laag", (B{1} * B{2}) / (4 * B{2} - B{3}), (B{1} * B{2}) / (4 * B{2} - B{4})), "")'.format(blokhoogte + 4, blokhoogte + 9, blokhoogte + 12, blokhoogte + 15, blokhoogte + 16)])    
        
        blok.append(["", "", "", ""])
        
        if flexibilisatie.pensioen.pensioenSoortRegeling == "DC": blok.append(["OP en PP Origineel", "=ROUND(D{0} / B{1}, 0)".format(blokhoogte + 7, blokhoogte + 11), "0", regelingBedrag(deelnemer, flexibilisatie)[2]])
        else: blok.append(["OP en PP Origineel", regelingBedrag(deelnemer, flexibilisatie)[0], regelingBedrag(deelnemer, flexibilisatie)[1], "=B{0} * B{1} + C{0} * B{2}".format(blokhoogte + 7, blokhoogte + 11, blokhoogte + 13)])
        blok.append(["OP en PP na Uitstellen", '=ROUND(B{0} * B{1} / B{2}, 0)'.format(blokhoogte + 7, blokhoogte + 11, blokhoogte + 12),\
                     '=ROUND(C{0} * B{1} / B{2}, 0)'.format(blokhoogte + 7, blokhoogte + 13, blokhoogte + 14), "formuletekst"])
        blok.append(["OP en PP na uitruilen", '=IF(B{0} =  "", B{5}, IF(B{0} = "Verhouding", ROUND(D{1} /  (B{2} * B{3} + C{2} *  B{4}), 0), IF(B{0} = "OP naar PP Percentage", ROUND(B{5} * (1 - MIN(B{2}, D{2})), 0), ROUND(B{5} + C{5} * B{2} * B{4} / B{3}, 0))))'.format(blokhoogte + 2, blokhoogte + 7, blokhoogte + 3, blokhoogte + 12, blokhoogte + 14, blokhoogte + 8),\
                     '=IF(B{0} =  "", C{5}, IF(B{0} = "Verhouding", ROUND(C{2} * D{1} /  (B{2} * B{3} + C{2} *  B{4}), 0), IF(B{0} = "OP naar PP Percentage", ROUND(C{5} + B{5} * MIN(B{2}, D{2}) * B{3} / B{4}, 0), ROUND(C{5} * (1 - B{2}), 0))))'.format(blokhoogte + 2, blokhoogte + 7, blokhoogte + 3, blokhoogte + 12, blokhoogte + 14, blokhoogte + 8), "formuletekst"])
        blok.append(["Op met hoog laag", '=IF(B{0} =  "", B{2}, IF(B{0} = "Verhouding",  ROUND((B{2} * B{3}) / IF(C{0} = "Hoog-laag", B{4} + C{1} * B{5}, C{1} * B{4} + B{5}), 0), ROUND(B{2} + IF(C{0} = "Hoog-laag", MIN(C{1}, D{1}) * B{4}, MIN(C{1}, D{1}) * B{5}) / B{3}, 0)))'.format(blokhoogte + 4, blokhoogte + 5, blokhoogte + 9, blokhoogte + 12, blokhoogte + 15, blokhoogte + 16),\
                     '=IF(B{0} =  "", B{2}, IF(B{0} = "Verhouding", ROUND(C{1} * (B{2} * B{3}) / IF(C{0} = "Hoog-laag", B{4} + C{1} * B{5}, C{1} * B{4} + B{5}), 0), ROUND(B{6} - MIN(C{1}, D{1}), 0)))'.format(blokhoogte + 4, blokhoogte + 5, blokhoogte + 9, blokhoogte + 12, blokhoogte + 15, blokhoogte + 16, blokhoogte + 10), "formuletekst"])
        
        blok.append(["Sommatie OP origineel", '=ROUND(SUMPRODUCT(INDIRECT("{0}"& MAX(ROUNDUP(E{2} - B{3}, 0) + 3, 3)):{0}{4}, INDIRECT("{1}"& MAX(ROUNDUP(E{2} - B{3}, 0) + 3, 3)):{1}{4}), 3)'.format(inttoletter(rekenblokstart + 10), inttoletter(rekenblokstart + 11), instellingen["pensioeninfohoogte"] + i, blokhoogte + 1, instellingen["rekenblokgrootte"]), "",\
                     '=CONCAT("=SUMPRODUCT( {0}", MAX(ROUNDUP(E{2} - B{3}, 0) + 3, 3), ":{0}{4}, {1}", MAX(ROUNDUP(E{2} - B{3}, 0) + 3, 3), ":{1}{4})")'.format(inttoletter(rekenblokstart + 10), inttoletter(rekenblokstart + 11), instellingen["pensioeninfohoogte"] + i, blokhoogte + 1, instellingen["rekenblokgrootte"])])
        blok.append(["Sommatie OP uitstellen", '=ROUND(SUMPRODUCT(INDIRECT("{0}"& MAX(ROUNDUP(B{3} - E{2}, 0) + 3, 3)):{0}{4}, INDIRECT("{1}"& MAX(ROUNDUP(B{3} - E{2}, 0) + 3, 3)):{1}{4}), 3)'.format(inttoletter(rekenblokstart + 3), inttoletter(rekenblokstart + 6), instellingen["pensioeninfohoogte"] + i, blokhoogte + 1, instellingen["rekenblokgrootte"]), "",\
                     '=CONCAT("=SUMPRODUCT( {0}", MAX(ROUNDUP(B{3} - E{2}, 0) + 3, 3), ":{0}{4}, {1}", MAX(ROUNDUP(B{3} - E{2}, 0) + 3, 3), ":{1}{4})")'.format(inttoletter(rekenblokstart + 3), inttoletter(rekenblokstart + 6), instellingen["pensioeninfohoogte"] + i, blokhoogte + 1, instellingen["rekenblokgrootte"])])
        blok.append(["Sommatie PP origineel", '=ROUND(SUMPRODUCT(INDIRECT("{0}"& MAX(ROUNDUP(E{3} - B{4}, 0) + 3, 3)):{0}{5}, INDIRECT("{1}"& MAX(ROUNDUP(E{3} - B{4}, 0) + 3, 3)):{1}{5}, INDIRECT("{2}"& MAX(ROUNDUP(E{3} - B{4}, 0) + 3, 3)):{2}{5}), 3)'.format(inttoletter(rekenblokstart + 10), inttoletter(rekenblokstart + 5), inttoletter(rekenblokstart + 7), instellingen["pensioeninfohoogte"] + i, blokhoogte + 1, instellingen["rekenblokgrootte"]), "",\
                     '=CONCAT("=SUMPRODUCT( {0}", MAX(ROUNDUP(E{3} - B{4}, 0) + 3, 3), ":{0}{5}, {1}", MAX(ROUNDUP(E{3} - B{4}, 0) + 3, 3), ":{1}{5}, {2}", MAX(ROUNDUP(E{3} - B{4}, 0) + 3, 3), ":{2}{5})")'.format(inttoletter(rekenblokstart + 10), inttoletter(rekenblokstart + 5), inttoletter(rekenblokstart + 7), instellingen["pensioeninfohoogte"] + i, blokhoogte + 1, instellingen["rekenblokgrootte"])])
        blok.append(["Sommatie PP uitstellen", '=ROUND(SUMPRODUCT(INDIRECT("{0}"& MAX(ROUNDUP(B{4} - E{3}, 0) + 3, 3)):{0}{5}, INDIRECT("{1}"& MAX(ROUNDUP(B{4} - E{3}, 0) + 3, 3)):{1}{5}, INDIRECT("{2}"& MAX(ROUNDUP(B{4} - E{3}, 0) + 3, 3)):{2}{5}), 3)'.format(inttoletter(rekenblokstart + 3), inttoletter(rekenblokstart + 5), inttoletter(rekenblokstart + 7), instellingen["pensioeninfohoogte"] + i, blokhoogte + 1, instellingen["rekenblokgrootte"]), "",\
                     '=CONCAT("=SUMPRODUCT( {0}", MAX(ROUNDUP(B{4} - E{3}, 0) + 3, 3), ":{0}{5}, {1}", MAX(ROUNDUP(B{4} - E{3}, 0) + 3, 3), ":{1}{5}, {2}", MAX(ROUNDUP(B{4} - E{3}, 0) + 3, 3), ":{2}{5})")'.format(inttoletter(rekenblokstart + 3), inttoletter(rekenblokstart + 5), inttoletter(rekenblokstart + 7), instellingen["pensioeninfohoogte"] + i, blokhoogte + 1, instellingen["rekenblokgrootte"])])
        blok.append(["Sommatie HL eerste", '=IF(B{5} = "", "", ROUND(SUMPRODUCT(INDIRECT("{0}"& MAX(ROUNDUP(B{3} - E{2}, 0) + 3, 3)):INDIRECT("{0}"&MAX(ROUNDUP(B{3} - E{2}, 0) + 2, 2) + B{4}), INDIRECT("{1}"& MAX(ROUNDUP(B{3} - E{2}, 0) + 3, 3)):INDIRECT("{1}"&MAX(ROUNDUP(B{3} - E{2}, 0) + 2, 2) + B{4})), 3))'.format(inttoletter(rekenblokstart + 3), inttoletter(rekenblokstart + 6), instellingen["pensioeninfohoogte"] + i, blokhoogte + 1, blokhoogte + 5, blokhoogte + 4), "",\
                     '=IF(B{5} = "", "",CONCAT("=SUMPRODUCT({0}", MAX(ROUNDUP(B{3} - E{2}, 0) + 3, 3), ":{0}", MAX(ROUNDUP(B{3} - E{2}, 0) + 2, 2) + B{4}, ", {1}", MAX(ROUNDUP(B{3} - E{2}, 0) + 3, 3),":{1}", MAX(ROUNDUP(B{3} - E{2}, 0) + 2, 2) + B{4},")"))'.format(inttoletter(rekenblokstart + 3), inttoletter(rekenblokstart + 6), instellingen["pensioeninfohoogte"] + i, blokhoogte + 1, blokhoogte + 5, blokhoogte + 4)])
        blok.append(["Sommatie HL tweede", '=IF(B{5} = "", "", ROUND(SUMPRODUCT(INDIRECT("{0}"& MAX(ROUNDUP(B{3} - E{2}, 0) + 3, 3) + B{6}):{0}{4}, INDIRECT("{1}"& MAX(ROUNDUP(B{3} - E{2}, 0) + 3, 3) + B{6}):{1}{4}), 3))'.format(inttoletter(rekenblokstart + 3), inttoletter(rekenblokstart + 6), instellingen["pensioeninfohoogte"] + i, blokhoogte + 1, instellingen["rekenblokgrootte"], blokhoogte + 4, blokhoogte + 5), "",\
                     '=IF(B{5} = "", "", CONCAT("=SUMPRODUCT({0}", MAX(ROUNDUP(B{3} - E{2}, 0) + 3, 3) + B{6}, ":{0}{4}, {1}", MAX(ROUNDUP(B{3} - E{2}, 0) + 3, 3) + B{6}, ":{1}{4})"))'.format(inttoletter(rekenblokstart + 3), inttoletter(rekenblokstart + 6), instellingen["pensioeninfohoogte"] + i, blokhoogte + 1, instellingen["rekenblokgrootte"], blokhoogte + 4, blokhoogte + 5)])
        
        if sum([len(rij) for rij in blok]) == len(blok) * 4:
            blokruimte = sheet.range((blokhoogte, instellingen["pensioenblokkolom"]),\
                                     (blokhoogte + instellingen["blokgrootte"] - 1, instellingen["pensioenblokkolom"] + len(blok[0]) - 1)).options(ndims = 2)
            # geldblok = sheet.range((blokhoogte + 7, instellingen["pensioenblokkolom"] + 1),\
            #                          (blokhoogte + 10, instellingen["pensioenblokkolom"] + 2))
            # geldblok.api.NumberFormat = "Currency"
            blokruimte.formula = blok
            blokruimte.color = flexibilisatie.pensioen.pensioenKleurZacht
        else:
            logger.warning("berekeningen rekenblok niet allemaal gelijk\n{}".format([len(rij) for rij in blok]))
            logger.debug([len(rij) for rij in blok])
        
        # rekenblok header
    for i, flexibilisatie in enumerate(deelnemer.flexibilisaties):
        blokhoogte = instellingen["pensioeninfohoogte"] + instellingen["afstandtotblokken"] + aantalpensioenen + i * (instellingen["blokgrootte"] + instellingen["afstandtussenblokken"])
        rekenblokstart = instellingen["afstandtotrekenkolom"] + i * (instellingen["rekenblokbreedte"] + instellingen["afstandtussenrekenblokken"])
        blok = list()
        blok.append([flexibilisatie.pensioen.pensioenVolNaam] + [""] * (instellingen["rekenblokbreedte"] - 1))
        blok.append(["t", "jaar", "Leeftijd", "tpx", "tqx", "tqx op 1 juli", "dt", "dt op 1 juli", "t'", "leeftijd'", "tpx'", "dt'"])
        rij = list()
        rij.append("0")
        rij.append("={} + ROUNDDOWN({} + {}3, 0)".format(deelnemer.geboortedatum.year, deelnemer.geboortedatum.month / 12,inttoletter(rekenblokstart + 2)))
        rij.append("=min(E{},B{})".format(instellingen["pensioeninfohoogte"] + i, blokhoogte + 1))
        rij.append("1")
        rij.append('=if({0}3<>"", 1-{0}3, "")'.format(inttoletter(rekenblokstart + 3)))
        rij.append('=if({0}4<>"", (((12 - MOD( 7 - {2} - ({0}3 - TRUNC({1}3)) * 12,12)) * {1}3) + MOD( 7 - {2} - ({0}3 - TRUNC({0}3)) * 12, 12) * {1}4) / 12, "")'.format(inttoletter(rekenblokstart + 2), inttoletter(rekenblokstart + 4), deelnemer.geboortedatum.month))
        rij.append('=if({0}3<>"", (1+{1})^-{2}3, "")'.format(inttoletter(rekenblokstart + 2), flexibilisatie.pensioen.rente / 100, inttoletter(rekenblokstart)))
        rij.append('=if({0}4<>"", (1+{1})^-({2}3 + (7 - {3}) / 12), "")'.format(inttoletter(rekenblokstart + 2), flexibilisatie.pensioen.rente / 100, inttoletter(rekenblokstart), deelnemer.geboortedatum.month))
        rij.append("0")
        rij.append("={}3".format(inttoletter(rekenblokstart + 2)))
        rij.append("1")
        rij.append('=if({0}3<>"", (1+{1})^-{2}3, "")'.format(inttoletter(rekenblokstart + 9), flexibilisatie.pensioen.rente / 100, inttoletter(rekenblokstart + 8)))
        blok.append(rij)
        
        if sum([len(rij) for rij in blok]) == len(blok) * instellingen["rekenblokbreedte"]:
            blokruimte = sheet.range((1, instellingen["afstandtotrekenkolom"] + i * (instellingen["rekenblokbreedte"] + instellingen["afstandtussenrekenblokken"] )),\
                                     (3, instellingen["afstandtotrekenkolom"] + i * (instellingen["rekenblokbreedte"] + instellingen["afstandtussenrekenblokken"] ) + instellingen["rekenblokbreedte"] - 1))
            mergeruimte = sheet.range((1, instellingen["afstandtotrekenkolom"] + i * (instellingen["rekenblokbreedte"] + instellingen["afstandtussenrekenblokken"] )),\
                                     (1, instellingen["afstandtotrekenkolom"] + i * (instellingen["rekenblokbreedte"] + instellingen["afstandtussenrekenblokken"] ) + instellingen["rekenblokbreedte"] - 1))
            blokruimte.formula = blok
            blokruimte.color = flexibilisatie.pensioen.pensioenKleurZacht
            mergeruimte.merge()
            mergeruimte.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
        else:
            logger.warning("berekeningen rekenblok niet allemaal gelijk\n{}".format([len(rij) for rij in blok]))
            logger.debug([len(rij) for rij in blok])
            
    # rekenblok Body
    for i, flexibilisatie in enumerate(deelnemer.flexibilisaties):
        blokhoogte = instellingen["pensioeninfohoogte"] + instellingen["afstandtotblokken"] + aantalpensioenen + i * (instellingen["blokgrootte"] + instellingen["afstandtussenblokken"])
        rekenblokstart = instellingen["afstandtotrekenkolom"] + i * (len(blok[0]) + instellingen["afstandtussenrekenblokken"])
        rij = list()
        rij.append('=IF({0}4 <> "", {0}4 -{0}$3, "")'.format(inttoletter(rekenblokstart + 2)))
        rij.append("={}3 + 1".format(inttoletter(rekenblokstart + 1)))
        rij.append('=IF({0}3<118, IF(AND(B${1} - TRUNC(B${1}) <> 0, {2}4 - {2}$3 = 1,  {0}3 - TRUNC({0}3) = 0), {0}3 + B${1} - TRUNC(B${1}), {0}3 + 1), "")'.format(inttoletter(rekenblokstart + 2), blokhoogte + 1, inttoletter(rekenblokstart + 1)))
        if flexibilisatie.pensioen.sterftetafel == "AG_2020": rij.append('=IF({0}4 <> "", ((1-@INDEX(INDIRECT("{1}"), {0}3+1, {2}3 - 2018)) *  (1 - {0}3 + TRUNC({0}3)) + (1-@INDEX(INDIRECT("{1}"), {0}3+2, {2}3 - 2018)) * ({0}3 - TRUNC({0}3))) * {3}3,"")'.format(inttoletter(rekenblokstart + 2),  flexibilisatie.pensioen.sterftetafel, inttoletter(rekenblokstart + 1), inttoletter(rekenblokstart + 3)))
        else: rij.append('=IF({0}4<>"", (INDEX(INDIRECT("{1}"),TRUNC({0}4)+1,1) * (1 - {0}4 + TRUNC({0}4)) + INDEX(INDIRECT("{1}"),TRUNC({0}4)+2,1) * ({0}4 - TRUNC({0}4))) / (INDEX(INDIRECT("{1}"),TRUNC(${0}$3)+1,1) * (1 - ${0}$3 + TRUNC(${0}$3)) + INDEX(INDIRECT("{1}"),TRUNC(${0}$3)+2,1) * (${0}$3 - TRUNC(${0}$3))),"")'.format(inttoletter(rekenblokstart + 2), flexibilisatie.pensioen.sterftetafel))
        rij.append('=IF({0}4<>"", 1-{0}4, "")'.format(inttoletter(rekenblokstart + 3)))
        rij.append('=if({0}5<>"", (((12 - MOD( 7 - {2} - ({0}4 - TRUNC({1}4)) * 12,12)) * {1}4) + MOD( 7 - {2} - ({0}4 - TRUNC({0}4)) * 12, 12) * {1}5) / 12, "")'.format(inttoletter(rekenblokstart + 2), inttoletter(rekenblokstart + 4), deelnemer.geboortedatum.month))
        rij.append('=IF({0}4<>"", (1+{1})^-{2}4, "")'.format(inttoletter(rekenblokstart + 2), flexibilisatie.pensioen.rente / 100, inttoletter(rekenblokstart)))
        rij.append('=IF({0}5<>"", (1+{1})^-({2}4 + (7 - {3}) / 12), "")'.format(inttoletter(rekenblokstart + 2), flexibilisatie.pensioen.rente / 100, inttoletter(rekenblokstart), deelnemer.geboortedatum.month))
        
        rij.append('=IF({0}4 <> "", {0}4 -{0}$3, "")'.format(inttoletter(rekenblokstart + 9)))
        rij.append('=IF({0}3<118, TRUNC({0}3) + 1,"")'.format(inttoletter(rekenblokstart + 9)))
        if flexibilisatie.pensioen.sterftetafel == "AG_2020": rij.append('=IF({0}4 <> "", ((1-@INDEX(INDIRECT("{1}"), {0}3+1, {2}3 - 2018)) *  (1 - {0}3 + TRUNC({0}3)) + (1-@INDEX(INDIRECT("{1}"), {0}3+2, {2}3 - 2018)) * ({0}3 - TRUNC({0}3))) * {3}3,"")'.format(inttoletter(rekenblokstart + 9),  flexibilisatie.pensioen.sterftetafel, inttoletter(rekenblokstart + 1), inttoletter(rekenblokstart + 10)))
        else: rij.append('=IF({0}4<>"", (INDEX(INDIRECT("{1}"),TRUNC({0}4)+1,1) * (1 - {0}4 + TRUNC({0}4)) + INDEX(INDIRECT("{1}"),TRUNC({0}4)+2,1) * ({0}4 - TRUNC({0}4))) / (INDEX(INDIRECT("{1}"),TRUNC(${0}$3)+1,1) * (1 - ${0}$3 + TRUNC(${0}$3)) + INDEX(INDIRECT("{1}"),TRUNC(${0}$3)+2,1) * (${0}$3 - TRUNC(${0}$3))),"")'.format(inttoletter(rekenblokstart + 9), flexibilisatie.pensioen.sterftetafel))
        rij.append('=IF({0}4<>"", (1+{1})^-{2}4, "")'.format(inttoletter(rekenblokstart + 9), flexibilisatie.pensioen.rente / 100, inttoletter(rekenblokstart)))
        
        blokruimte = sheet.range((4, instellingen["afstandtotrekenkolom"] + i * (instellingen["rekenblokbreedte"] + instellingen["afstandtussenrekenblokken"] )),\
                                 (max(4, instellingen["rekenblokgrootte"]), instellingen["afstandtotrekenkolom"] + i * (instellingen["rekenblokbreedte"] + instellingen["afstandtussenrekenblokken"]) + instellingen["rekenblokbreedte"] - 1))
        blokruimte.formula = rij
        blokruimte.color = flexibilisatie.pensioen.pensioenKleurZacht
        
    #sheet protecten
    #ProtectBeheer(sheet) #.api.Protect(Password = wachtwoord())
    
    logger.info("berekenscherm init afgerond")
    

def inttoletter(getal):
    """
    Functie die een kolom getal naar kolom letter vertaald

    Parameters
    ----------
    getal : INT
        een getal dat de kolom waarde weergeeft.

    Returns
    -------
    STR
        De string van de kolomnaam, 1 -> "A", 27 -> "AA" etc.

    """
    
    letter = ""
    while True:
        if getal > 26:
            letter = "{}{}".format(ascii_uppercase[(getal%26) - 1], letter)
            getal = (getal-getal%26)//26 - (1 - min(getal%26, 1))
        else: return "{}{}".format(ascii_uppercase[(getal%26) - 1], letter)

def berekeningen_instellingen():
    """
    Functie die de settings van het berekenings sheet meegeeft.

    Parameters
    ----------
    None

    Returns
    -------
    Instellingen: dict
        een dictionary met daarin de waardes van de lokaties van de berekening blokken.

    """
    instellingen = dict()
    instellingen["pensioeninfohoogte"] = 2
    instellingen["pensioeninfokolom"] = 1
    instellingen["pensioenblokkolom"] = 1

    instellingen["afstandtotblokken"] = 6
    instellingen["afstandtussenblokken"] = 2
    instellingen["blokgrootte"] = 17
    
    instellingen["afstandtotrekenkolom"] = 8
    instellingen["afstandtussenrekenblokken"] = 1
    instellingen["rekenblokgrootte"] = 63
    instellingen["rekenblokbreedte"] = 12
    
    return instellingen

def leesOPPP(sheet, flexibilisaties):
    """
    leest de OP en PP waardes uit het berekeningssheet en slaat ze op in flexibilisaties.

    Parameters
    ----------
    sheet : Book.Sheet
        Berekeningen sheet.
    flexibilisaties : flex_keuzes
        object met alle flexibilisatie eigenschappen.

    Returns
    -------
    None.

    """
    
    instellingen = berekeningen_instellingen()
    for i, flexibilisatie in enumerate(flexibilisaties):
        blokhoogte = instellingen["pensioeninfohoogte"] + instellingen["afstandtotblokken"] + len(flexibilisaties) + i * (instellingen["blokgrootte"] + instellingen["afstandtussenblokken"])
        bereik = sheet.range((blokhoogte + 9, 2), (blokhoogte + 10, 3)).options(ndims = 2, numbers = int).value
        OPU = bereik[0][0]
        PP = bereik[0][1]
        OPH = bereik[1][0]
        OPL = bereik[1][1]
        flexibilisatie.ouderdomsPensioenUitruilen = OPU
        flexibilisatie.partnerPensioen = PP
        flexibilisatie.ouderdomsPensioenHoog = OPH
        flexibilisatie.ouderdomsPensioenLaag = OPL
               
def leesLimietMeldingen(sheet, flexibilisaties, huidigRegelingNaam):
    """
    leest de flexibilisaties uit de rekensheet en slaat eventuele andere gehanteerde waardes op
    Bijvoorbeeld als een OP naar PP percentage te hoog was, dan wordt het nieuw gehanteerde percentage opgeslagen.

    Parameters
    ----------
    sheet : Book.Sheet
        Berekeningen sheet.
    flexibilisaties : flex_keuzes
        object met alle flexibilisatie eigenschappen.
    limietMelding : bool
        True als er gecheckt moet worden of er een limiet is bereikt.

    Returns
    -------
    Lijst met per regeling de gehanteerde limieten.
    """

    instellingen = berekeningen_instellingen()
    for i, flexibilisatie in enumerate(flexibilisaties):
        if huidigRegelingNaam == flexibilisatie.pensioen.pensioenNaam:
            blokhoogte = instellingen["pensioeninfohoogte"] + instellingen["afstandtotblokken"] + len(flexibilisaties) + i * (instellingen["blokgrootte"] + instellingen["afstandtussenblokken"])
            bereik = sheet.range((blokhoogte + 3, 1), (blokhoogte + 5, 4)).options(ndims = 2, numbers = float).value
            
            return bereik

        
def maak_afbeelding(deelnemer, sheet = None, ax = None, ID = 0, pdf = False, titel = ""):
    """
    Maakt de afbeelding in het flexscherm.

    Parameters
    ----------
    flexibilisaties : list(flex_keuzes)
        lijst met daarin de flexibilisaties.
    ax : axes
        subplot object uit het flexscherm.

    Returns
    -------
    None.

    """
    
    # verkrijg AOW
    AOW = None
    for pensioen in deelnemer.pensioenen:
        if pensioen.pensioenSoortRegeling == "AOW": 
            AOW = pensioen
            break
    # Lijst met alle voorkomende jaren van OP
    allejaren = set()
    if AOW != None: allejaren.add(AOW.pensioenleeftijd)
    for flexibilisatie in deelnemer.flexibilisaties:
        if flexibilisatie.HL_Actief and flexibilisatie.HL_Methode == "Opvullen AOW": allejaren.add(flexibilisatie.AOWJaar + flexibilisatie.AOWMaand / 12)
        elif flexibilisatie.leeftijd_Actief: allejaren.add(flexibilisatie.leeftijdJaar + flexibilisatie.leeftijdMaand / 12)
        else: allejaren.add(flexibilisatie.pensioen.pensioenleeftijd)
        if flexibilisatie.HL_Actief and flexibilisatie.HL_Methode != "Opvullen AOW": 
            if flexibilisatie.leeftijd_Actief: allejaren.add(flexibilisatie.leeftijdJaar + flexibilisatie.leeftijdMaand / 12 + flexibilisatie.HL_Jaar)
            else: allejaren.add(flexibilisatie.pensioen.pensioenleeftijd + flexibilisatie.HL_Jaar)
    # geeft de breedte aan van alle hoogtes
    randen = list(allejaren)
    randen.sort()
    randen.append(randen[-1] + 10)
    
    # een lijst met alle verzekeringsnamen
    naamlijst = list()
    if AOW != None: naamlijst.append(AOW.pensioenNaam)
    for flexibilisatie in deelnemer.flexibilisaties: naamlijst.append(flexibilisatie.pensioen.pensioenVolNaam) 
    
    # bepaald de kleuren
    kleuren = list()
    if AOW != None: kleuren.append(tuple([kleur / 255 for kleur in AOW.pensioenKleurHard]))
    for flexibilisatie in deelnemer.flexibilisaties: kleuren.append(tuple([kleur / 255 for kleur in flexibilisatie.pensioen.pensioenKleurHard]))
    
    #berekent de hoogte van elke staaf
    hoogtes = [[0 for i in range(len(randen)-1)]]
    ywaardes = set()
    ywaardes.add(0)
    
    #AOW toevoegen
    if AOW != None:
        hoogtes.append(list())
        startjaar = AOW.pensioenleeftijd
        for j, leeftijd in enumerate(randen[:-1]):
            if leeftijd < startjaar: hoogtes[1].append(hoogtes[0][j])
            else: 
                if deelnemer.burgelijkeStaat == "Samenwonend": bedrag = float(hoogtes[0][j] + AOW.samenwondendAOW)
                else: bedrag = float(hoogtes[0][j] + AOW.alleenstaandAOW)
                hoogtes[1].append(bedrag)
                ywaardes.add(bedrag)
    
    # De flexibilisaties toevoegen
    for i, flexibilisatie in enumerate(deelnemer.flexibilisaties):
        if AOW != None: i += 1
        if flexibilisatie.HL_Actief and flexibilisatie.HL_Methode == "Opvullen AOW" and AOW != None: startjaar = flexibilisatie.AOWJaar + flexibilisatie.AOWMaand / 12
        elif flexibilisatie.leeftijd_Actief: startjaar = flexibilisatie.leeftijdJaar + flexibilisatie.leeftijdMaand / 12
        else: startjaar = flexibilisatie.pensioen.pensioenleeftijd
        aanspraakHoog = flexibilisatie.ouderdomsPensioenHoog
        aanspraakLaag = flexibilisatie.ouderdomsPensioenLaag
        if flexibilisatie.HL_Actief and flexibilisatie.HL_Methode == "Opvullen AOW" and AOW != None: HoogLaagjaren = AOW.pensioenleeftijd - startjaar
        else: HoogLaagjaren = flexibilisatie.HL_Jaar
        HoogLaagVolgorde = flexibilisatie.HL_Volgorde
        
        hoogtes.append(list())

        for j, leeftijd in enumerate(randen[:-1]):
            if leeftijd < startjaar: hoogtes[i+1].append(hoogtes[i][j])
            elif flexibilisatie.HL_Actief:
                if leeftijd < startjaar + HoogLaagjaren:
                    if HoogLaagVolgorde == "Hoog-laag": bedrag = float(hoogtes[i][j] + aanspraakHoog)
                    else: bedrag = float(hoogtes[i][j] + aanspraakLaag)
                else:
                    if HoogLaagVolgorde == "Hoog-laag": bedrag = float(hoogtes[i][j] + aanspraakLaag)
                    else: bedrag = float(hoogtes[i][j] + aanspraakHoog)
                hoogtes[i+1].append(bedrag)
                ywaardes.add(bedrag)
                        
            else: 
                bedrag = float(hoogtes[i][j] + aanspraakHoog)
                hoogtes[i+1].append(bedrag)
                ywaardes.add(bedrag)
    ywaardes = list(ywaardes)
    ywaardes.sort()
    
    # bereken PP
    PPtotaal = 0
    for flexibilisatie in deelnemer.flexibilisaties: PPtotaal += flexibilisatie.partnerPensioen
    
    # maak titel - titel al meegegeven bij aanroepen functie (Standaard "Een super coole title")
    #titel = "Een super coole title"
    # maak de afbeeling
    if ax != None:
        ax.clear()
        for i in range(len(hoogtes) - 1): ax.stairs(hoogtes[i+1], edges = randen, baseline = hoogtes[i], fill=True, label = naamlijst[i], color = kleuren[i])
        
        ax.set_xticks(randen[:-1], [getaltotijd(rand) for rand in randen[:-1]])
        ax.set_xticklabels([getaltotijd(rand) for rand in randen[:-1]], rotation=30, horizontalalignment='right')
        ax.set_yticks(ywaardes, [getaltogeld(ywaarde) for ywaarde in ywaardes])
    
        handles, labels = ax.get_legend_handles_labels()
        if AOW != None: order = range(len(deelnemer.flexibilisaties), -1, -1)
        else: order = range(len(deelnemer.flexibilisaties) - 1, -1, -1)
        ax.legend(handles = [handles[idx] for idx in order], labels = [labels[idx] for idx in order]) 
    
        ax.set_xlabel("Totale partnerpensioen: €{:.2f}".format(PPtotaal).replace(".",","))
        ax.set_title(titel, fontweight='bold')
    
    if sheet != None or pdf:

        afbeelding = plt.figure()
        for i in range(len(hoogtes) - 1): plt.stairs(hoogtes[i+1],edges = randen,  baseline=hoogtes[i], fill=True, label = naamlijst[i], color = kleuren[i])
        
        plt.xticks(randen[:-1], [getaltotijd(rand) for rand in randen[:-1]])
        plt.setp(plt.gca().get_xticklabels(), rotation=30, horizontalalignment='right')
        plt.yticks(ywaardes, [getaltogeld(ywaarde) for ywaarde in ywaardes])

        handles, labels = plt.gca().get_legend_handles_labels()
        if AOW != None: order = range(len(deelnemer.flexibilisaties), -1, -1)
        else: order = range(len(deelnemer.flexibilisaties) - 1, -1, -1)
        plt.legend([handles[idx] for idx in order],[labels[idx] for idx in order]) 
        
        plt.suptitle(titel, fontweight='bold')
        plt.xlabel("Totale partnerpensioen: €{:.2f}".format(PPtotaal).replace(".",","))
        
        if sheet != None:
            if ID == 0:
                locatie = sheet.range((14,2))
            else:
                teller = ID-1
                locatieTop = int(12 + (teller%4)*22) #12 + 22*int(ID)    #afbeeldingsformaat in cellen = 22 hoog, 8 breed
                locatieLeft = int(17 + ((teller - teller%4)/4)*8) #maximaal 4 afbeeldingen onder elkaar, daarna ernaast verder
                locatie = sheet.range((locatieTop,locatieLeft))
            #sheet unprotecten
            sheet.api.Unprotect(Password = wachtwoord())
            #afbeelding opslaan op sheet
            if ID == 0:
                try:
                    sheet.pictures.add(afbeelding, update = True, top = locatie.top, left = locatie.left, height = 300, name = "Vergelijking {}".format(ID))
                except:
                    sheet.pictures.add(afbeelding, top = locatie.top, left = locatie.left, height = 300, name = "Vergelijking {}".format(ID))
            else:
                sheet.pictures.add(afbeelding, top = locatie.top, left = locatie.left, height = 300, name = "Vergelijking {}".format(ID))
            #sheet protecten
            ProtectBeheer(sheet) #.api.Protect(Password = wachtwoord(), Contents=False)
        if pdf:
            # Credits to Philipp on Stackoverflow
            # https://stackoverflow.com/questions/18897511/how-to-drawimage-a-matplotlib-figure-in-a-reportlab-canvas
            afbeelding.set_size_inches(300/72, 231/72)
            afbeeldingData = BytesIO()
            
            afbeelding.savefig(afbeeldingData, format='svg')
            afbeeldingData.seek(0)  # rewind the data
            
            Image = svg2rlg(afbeeldingData)
            return Image

        

def vergelijken_keuzes():
    """
    functie die de drop down list in de vergelijkingssheet vult met de namen van de opgeslagen afbeeldingen

    Parameters
    ----------
    box : integer
        0 - alle drie de keuzecellen updaten
        1 - 1e keuzecel updaten
        2 - linker vergelijken keuzecel updaten
        3 - rechter vergelijken keuzecel updaten
        
    Returns
    -------
    drop down list gevuld met namen uit de flexopslag

    """
    
    #sheets en book opslaan in variabelen
    book = xw.Book.caller()
    invoer = book.sheets["Flexopslag"]
    uitvoer = book.sheets["Vergelijken"]
    #list maken waarin de opgeslagen pensioenen worden bijgehouden
    pensioenlist = ["-"]
    celKolom = 5 
    #sheet unprotecten
    uitvoer.api.Unprotect(Password = wachtwoord())
    
    
    keuzecellen = ["B6", "J13", "B37", "J37"]
    if str(invoer.cells(2,celKolom).value) != "None":   #alleen als er nog flexibilisaties opgeslagen zijn
        #rij met flexibilisatienaam langsgaan en elke naam toevoegen aan pensioenlist
        while str(invoer.cells(2,celKolom).value) != "None":
            naam = str(invoer.cells(2,celKolom).value)
            pensioenlist.append(naam)
            celKolom += 4
        #lijst omzetten naar string, gescheiden door komma
        pensioenopties = ','.join(pensioenlist[1:])
        
        #keuzeveld1
        #verwijder bestaande datavalidatie uit cel
        uitvoer[keuzecellen[0]].api.Validation.Delete()
        #voeg nieuwe datavalidatie toe aan cel
        uitvoer[keuzecellen[0]].api.Validation.Add(Type=DVType.xlValidateList, Formula1=pensioenopties)
        #vul keuzeveld1 met eerste opties uit pensioenlist
        uitvoer[keuzecellen[0]].value = pensioenlist[1]
        #lege optie toevoegne aan pensioenopties
        pensioenopties = ','.join(pensioenlist)
        for cel in keuzecellen[1:]:
            #verwijder bestaande datavalidatie uit cel
            uitvoer[cel].api.Validation.Delete()
            #voeg nieuwe datavalidatie toe aan cel
            uitvoer[cel].api.Validation.Add(Type=DVType.xlValidateList, Formula1=pensioenopties)
        #vul keuzeveld1 met eerste opties uit pensioenlist
        #uitvoer[keuzecellen[0]].value = pensioenlist[0]
    else:   #geen flexibilisaties opgeslagen
        for cel in keuzecellen:
            #verwijder bestaande datavalidatie uit cel
            uitvoer[cel].api.Validation.Delete()
            #maak keuzeveld leeg
            uitvoer[cel].value = ""
            #voeg nieuwe datavalidatie toe aan cel
            pensioenopties = ','.join([" "])
            uitvoer[cel].api.Validation.Add(Type=DVType.xlValidateCustom, Formula1=pensioenopties)
    #sheet protecten
    ProtectBeheer(uitvoer) #.api.Protect(Password=wachtwoord(), Contents=False)
   
def opslagLegen(book, logger):
    flexopslag = book.sheets["Flexopslag"]
    vergelijken = book.sheets["Vergelijken"]
    if str(flexopslag.cells(2, 5).value) != "None":   #alleen als er nog flexibilisaties opgeslagen zijn
        #Vergelijken sheet legen
        #ID-nummer van laatste opgeslagen flexibilisatie vinden
        kolomLaatsteOpslag = FlexopslagVinden(book)[1]
        IDLaatsteOpslag = flexopslag.cells(3, kolomLaatsteOpslag).value
        stopNummer = int(IDLaatsteOpslag.split()[-1])
        #vergelijken sheet unprotecten
        vergelijken.api.Unprotect(Password=wachtwoord())
        #alle ID's tot laatste ID afgaan om (mogelijke) afbeelding te verwijderen
        for i in range(0,stopNummer+1):
            ID = f"Vergelijking {i}"
            try:
                vergelijken.pictures[ID].delete()
            except:
                pass
        #vergelijken sheet protecten
        ProtectBeheer(vergelijken) #.api.Protect(Password=wachtwoord(), Contents=False)
        logger.info("Afbeeldingen op vergelijken sheet verwijderd")
        #flexopslag unprotecten
        flexopslag.api.Unprotect(Password = wachtwoord())
        #opgeslagen flexibilisaties van vorige deelnemer verwijderen uit opslag
        flexopslag.clear()
        #flexopslag protecten
        ProtectBeheer(flexopslag) #.api.Protect(Password = wachtwoord())
        
        #laatste opslag is verwijderd, dus drop down legen
        vergelijken_keuzes()
        logger.info("Flexopslag is geleegd")
    else:
        logger.info("Flexopslag legen niet nodig, was al leeg")

def dictAssign(lbl,lbl_OP,lbl_PP,lbl_pLeeftijd,lbl_OP_PP,lbl_hlConstructie):
    """
    Parameters
    ----------
    lbl,lbl_OP,lbl_PP,lbl_pLeeftijd,lbl_OP_PP,lbl_hlConstructie : self.ui labels
        Dit zijn verwijzingen naar samenvatting labels uit het flexmenu
    
    Returns
    -------
    rDict: dict
        Dictionary met labels van de samenvatting in het flexmenu voor specifieke regeling
    """
    rDict = dict()
    
    rDict["lbl"] = lbl
    rDict["lbl_OP"] = lbl_OP
    rDict["lbl_PP"] = lbl_PP
    rDict["lbl_pLeeftijd"] = lbl_pLeeftijd
    rDict["lbl_OP_PP"] = lbl_OP_PP
    rDict["lbl_hlConstructie"] = lbl_hlConstructie
    
    return rDict

def samenvattingDict(regeling,UI):
    """
    Parameters
    ----------
    regeling : str
        Korte naam van de regeling
    
    UI : self.ui object
        Meegegeven vanuit flexmenu class
    
    Returns
    -------
    regelingDict : dict
        Dictionary met labels van de samenvatting in het flexmenu voor specifieke regeling
    """
    
    regelingDict = dict()
    
    if regeling == "ZL":
        regelingDict = dictAssign(UI.lbl_ZL,
                   UI.lbl_ZL_OP,
                   UI.lbl_ZL_PP,
                   UI.lbl_ZL_pLeeftijd,
                   UI.lbl_ZL_OP_PP,
                   UI.lbl_ZL_hlConstructie)
    elif regeling == "Aegon OP65":
        regelingDict = dictAssign(UI.lbl_A65,
                   UI.lbl_A65_OP,
                   UI.lbl_A65_PP,
                   UI.lbl_A65_pLeeftijd,
                   UI.lbl_A65_OP_PP,
                   UI.lbl_A65_hlConstructie)
    elif regeling == "Aegon OP67":
        regelingDict = dictAssign(UI.lbl_A67,
                   UI.lbl_A67_OP,
                   UI.lbl_A67_PP,
                   UI.lbl_A67_pLeeftijd,
                   UI.lbl_A67_OP_PP,
                   UI.lbl_A67_hlConstructie)
    elif regeling == "NN OP65":
        regelingDict = dictAssign(UI.lbl_NN65,
                   UI.lbl_NN65_OP,
                   UI.lbl_NN65_PP,
                   UI.lbl_NN65_pLeeftijd,
                   UI.lbl_NN65_OP_PP,
                   UI.lbl_NN65_hlConstructie)
    elif regeling == "NN OP67":
        regelingDict = dictAssign(UI.lbl_NN67,
                   UI.lbl_NN67_OP,
                   UI.lbl_NN67_PP,
                   UI.lbl_NN67_pLeeftijd,
                   UI.lbl_NN67_OP_PP,
                   UI.lbl_NN67_hlConstructie)
    elif regeling == "PF VLC OP68":
        regelingDict = dictAssign(UI.lbl_VLC,
                   UI.lbl_VLC_OP,
                   UI.lbl_VLC_PP,
                   UI.lbl_VLC_pLeeftijd,
                   UI.lbl_VLC_OP_PP,
                   UI.lbl_VLC_hlConstructie)
    
    return regelingDict

    
    
    
def nieuwe_pagina(pdf, halfbreedte):
    """
    Functie die alles op de pagina van de pdf zet dat op elke pagina moet komen

    Parameters
    ----------
    pdf : Canvas van reportlap
        Een pdf vanuit canvas
    halfbreedte: float
        De helft van de breedte van een a4

    Returns
    -------
    Pdf
        Het VLC-logo rechtsboven in de hoek en een lijn door het midden van de pagina in de pdf

    """
    breedte_logo = 183.2
    hoogte_logo = 40
    image = ("{}\\logo.png".format(krijgpad()))
    #Zet het Vlc logo rechtsboven op de pagina
    pdf.drawImage(image, cm*21 -breedte_logo, cm* 29.7-hoogte_logo, breedte_logo, hoogte_logo)
    #Maakt een streep in het midden van de pagina om zo oud en nieuw te splitsen
    pdf.line(halfbreedte, 0, halfbreedte, cm* 29.7)
    pdf.setFont("Helvetica", 30)
    pdf.drawString(30, 770, "Nieuw")
    pdf.drawString(halfbreedte + 30, 770, "Oud")
    pdf.setFont("Helvetica", 11)
    
    
def leeftijd_notatie(jaar, maand):
    """
    Functie die een pensioen leeftijd mooi neerzet in taal

    Parameters
    ----------
    jaar : str
        Het jaar van de pensioenleeftijd
    maand: str
        De maand van de pensioenleeftijd

    Returns
    -------
    str
        In woorden wat de pensioenleeftijd van een pensioen is rekening houdend met gramatica ev/mv

    """
    maand = str(int(float(maand)))
    jaar = str(int(float(jaar)))

    if maand == "0":
        antwoord = jaar + " jaar"
    elif maand == "1":
        antwoord = jaar + " jaar en 1 maand"
    else:
        antwoord = jaar + " jaar en " + maand + " maanden"
    return antwoord  


    
    
def geld_per_leeftijd(oud_pensioen, nieuw_pensioen):
    """
    Functie die per leeftijd aangeeft hoeveel geld erbij of af gaat

    Parameters
    ----------
    oud_pensioen : list
        Lijst met lijsten van oud pensioen. 
    
    nieuw_pensioen: list
        Lijst met lijsten van oud pensioen.

    Returns
    -------
    list
        Een lijst met de lijst van oud en nieuw pensioen. 
        oud en nieuwpensioen zijn beide een lijst met lijsten van 2 lang 
        bestaande uit de pensioen leeftijd en het verschil in bedrag met de leeftijd ervoor. 
        Deze lijst is van jong naar oud gesoorteerd.
        In deze lijst staan geen dubbele leeftijden meer.
        
    """
    datum_en_geldnieuw = []
    
    p = 1 #hoeveelste pensioen
    for i in nieuw_pensioen:
        if i[1] == "Ja": #pensioenleeftijd aanpassen
            startjaar = str(int(float(i[2])))
            startmaand = str(int(float(i[3])))
        else:
            startjaar = str(int(float(oud_pensioen[p][1])))
            startmaand = "0"
        
        if i[9] == "Ja": #hooglaag staat aan
            duur = int(float(i[11]))
            hl_jaar = str(int(startjaar) + duur)
            datum1 = leeftijd_notatie(startjaar, startmaand)
            datum2 = leeftijd_notatie(hl_jaar, startmaand)
            if i[10] == "Hoog-laag":
                OP2 = int(float(i[16])) - int(float(i[15])) #op tweede gedeelte hl
            else:
                OP2 = int(float(i[15])) - int(float(i[16])) 
            OP1 = int(float(i[15])) #op eerste gedeelte hl
            datum_en_geldnieuw.append([datum1, OP1])
            datum_en_geldnieuw.append([datum2, OP2])
            
        else:
            datum = leeftijd_notatie(startjaar, startmaand)
            geld = int(float(i[15]))
            datum_en_geldnieuw.append([datum, geld])
        p += 1            
                    
            
    datum_en_geldoud = []
    oud_pensioen.pop(0) #voor nu nog even gedaan omdat AOW er nog niet goed in lijkt te staan
    for i in oud_pensioen:
        datum = leeftijd_notatie(i[1], "0")
        geld = i[3]
        datum_en_geldoud.append([datum, geld])
    
    datum_en_geld = [datum_en_geldoud, datum_en_geldnieuw]
    
    oud_en_nieuw = [] #een lijst waar de enkellijsten van een oud en nieuw pensioen komen
                        #met oud op index 0 en nieuw op index 1
    
    for lijst in datum_en_geld:
        enkellijst = [["",""]]
        dubbel = 0
        for i in lijst:
            for j in range(len(enkellijst)):
                if i[0] == enkellijst[j][0]:
                    enkellijst[j] = [enkellijst[j][0], enkellijst[j][1]+ i[1]]
                    dubbel = 1
            if dubbel == 0:
                enkellijst.append(i)
            dubbel = 0
            
        enkellijst.pop(0)
        enkellijst.sort()
        
        oud_en_nieuw.append(enkellijst)
        
    return oud_en_nieuw
    
def tekstkleurSheets(book, sheets, zicht):
    """
    functie die de tekstkleur van de tekst in meegegeven sheets aanpast. Hierdoor wordt de tekst onleesbaar.

    Parameters
    ----------
    book : xw.book
        
    sheets : list("naam sheet")
        lijst met daarin namen van de sheets die aangepast moeten worden.
    zicht : integer
        0 - tekst onleesbaar
        1 - tekst leesbaar.

    Returns
    -------
    veranderd de tekstkleur van de tekst in de sheet om deze leesbaar of onleesbaar te maken.

    """
    #rgb_int definieren voor wit en zwart
    wit = 16777215
    zwart = 0
    for sheetnaam in sheets:
        sheet = book.sheets[sheetnaam]
        #mogelijk maken om sheet te wijzigen
        sheet.api.Unprotect(Password = wachtwoord())
        try:
            #grootte van gegevensblok inlezen
            aantalRegels = len(sheet.cells(1,1).expand().value) + 1
            aantalKolommen = len(sheet.cells(1,1).expand().value[0]) + 1
        except:
            aantalRegels = 1
            aantalKolommen = 1
        
        if zicht == 0:  #tekstkleur zelfde als achtegrondkleur -> onleesbaar maken
            if sheetnaam == "Sterftetafels":
                sheet.range((1,1),(aantalRegels,aantalKolommen)).api.Font.Color = wit
                sheet.range((1,2),(2,3)).api.Font.Color = rgb_to_int((146,208,80))  #groen
                sheet.range((3,1)).api.Font.Color = rgb_to_int((146,208,80))        #groen
            
            elif sheetnaam == "AG2020":
                sheet.range((1,1),(aantalRegels,aantalKolommen)).api.Font.Color = rgb_to_int((255,153,0))   #oranje
                sheet.range((2,2),(aantalRegels,aantalKolommen)).api.Font.Color = rgb_to_int((0,128,128))   #turquoise
            
            elif sheetnaam == "deelnemersbestand":
                kleuren = [(225,211,212), (229,220,255), (206,232,255), (255,227,194), (255,227,194), (255,255,197), (255,255,197), (222,255,250), (222,255,250)]
                sheet.range((1,1),(aantalRegels,9)).api.Font.Color = wit
                for i in range(0,(len(kleuren))):
                    sheet.range((1,i+10),(aantalRegels,i+10)).api.Font.Color = rgb_to_int(kleuren[i])
            
            elif sheetnaam == "Gegevens pensioencontracten":
                sheet.range((1,1),(aantalRegels,aantalKolommen)).api.Font.Color = wit
                sheet.range((2,2),(2,aantalKolommen)).api.Font.Color = rgb_to_int((146,208,80))  #groen
                
            elif sheetnaam == "Berekeningen": 
                sheet.shapes["VerbergBerekeningen"].api.Fill.Visible = True
            
            elif sheetnaam  == "Flexopslag":
                sheet.shapes["VerbergBerekeningen"].api.Fill.Visible = True
            
            #sheet weer beveiligen, omdat gebruiker gegevens niet mag zien
            if sheetnaam != "Vergelijken":
                ProtectBeheer(book.sheets[sheet]) #.api.Protect(Password = wachtwoord())
            else:
                ProtectBeheer(book.sheets[sheet]) #.api.Protect(Password = wachtwoord(), Contents=False)
        
        elif zicht == 1:    #tekstkleur zwart maken -> leesbaar maken
            if sheetnaam in ["Sterftetafels", "AG2020", "deelnemersbestand", "Gegevens pensioencontracten"]:
                sheet.range((1,1),(aantalRegels,aantalKolommen)).api.Font.Color = zwart
            elif sheetnaam in ["Berekeningen", "Flexopslag"]:
                sheet.shapes["VerbergBerekeningen"].api.Fill.Visible = False

def regelingBedrag(deelnemer, flexibilisatie):
    if flexibilisatie.HL_Actief and flexibilisatie.HL_Methode == "Opvullen AOW": jaren = flexibilisatie.AOWJaar - (date.today().year - deelnemer.geboortedatum.year)
    elif flexibilisatie.leeftijd_Actief: jaren = flexibilisatie.leeftijdJaar + flexibilisatie.leeftijdMaand - (date.today().year - deelnemer.geboortedatum.year)
    else: jaren = flexibilisatie.pensioen.pensioenleeftijd - (date.today().year - deelnemer.geboortedatum.year)
    if flexibilisatie.pensioen.pensioenSoortRegeling == "DC":
        if not flexibilisatie.pensioen.actieveRegeling: return (0, 0, flexibilisatie.pensioen.koopsom)
        bedrag = flexibilisatie.pensioen.koopsom
        for i in range(int(jaren)):
            bedrag += flexibilisatie.pensioen.regelingsFactor
            bedrag *= 1 + flexibilisatie.pensioen.rente
        return (0, 0, bedrag)
    elif flexibilisatie.pensioen.pensioenSoortRegeling == "DB":
        if not flexibilisatie.pensioen.actieveRegeling: return (flexibilisatie.pensioen.ouderdomsPensioen, 0, 0)
        bedrag = flexibilisatie.pensioen.ouderdomsPensioen
        bedrag += flexibilisatie.pensioen.regelingsFactor * jaren
        return (round(bedrag), 0, 0)
    elif flexibilisatie.pensioen.pensioenSoortRegeling == "DB met PP":
        if not flexibilisatie.pensioen.actieveRegeling: return (flexibilisatie.pensioen.ouderdomsPensioen, flexibilisatie.pensioen.partnerPensioen, 0)
        bedragOP = flexibilisatie.pensioen.ouderdomsPensioen
        bedragPP = flexibilisatie.pensioen.partnerPensioen
        bedragOP += flexibilisatie.pensioen.regelingsFactor * jaren / 1.7
        bedragPP += flexibilisatie.pensioen.regelingsFactor * jaren / 1.7 * 0.7
        return (round(bedragOP), round(bedragPP), 0)
    else: return (0, 0, 0)

def GegevensNaarFlexibilisatie(deelnemer, opslag):
    
    #lijst met pensioennamen van de deelnemer 
    pensioennamen = []  
    for i in opslag:
        pensioennamen.append(i[0])
    
    #lijst met pensioennamen langsgaan en opgeslagen flexibilisatiegegevens per pensioen toevoegne aan flexibiliseringsobject van het deelnemersobject
    for i,p in enumerate(pensioennamen):
        for flexibilisatie in deelnemer.flexibilisaties:
            #als het flexibilisatieobject bij het pensioen uit de lijst pensioennamen hoort
            if flexibilisatie.pensioen.pensioenNaam == p:
                #met properties flexibilisaties opslaan in objecten flexibilisatie
                pensioengegevens = opslag[i]
                #leeftijd aanpassen
                if pensioengegevens[1] == "Ja":
                    flexibilisatie.leeftijd_Actief = True
                elif pensioengegevens[1] == "Nee":
                    flexibilisatie.leeftijd_Actief = False
                flexibilisatie.leeftijdJaar = int(float(pensioengegevens[2]))
                flexibilisatie.leeftijdMaand = int(float(pensioengegevens[3]))
                
                #uitruilen
                if pensioengegevens[4] == "Ja":
                    flexibilisatie.OP_PP_Actief = True
                elif pensioengegevens[4] == "Nee":
                    flexibilisatie.OP_PP_Actief = False
                    #volgorde
                flexibilisatie.OP_PP_UitruilenVan = pensioengegevens[5]
                    #methode
                flexibilisatie.OP_PP_Methode = pensioengegevens[6]
                if pensioengegevens[6] == "Verhouding":
                    flexibilisatie.OP_PP_Verhouding_OP = int(float(pensioengegevens[7]))
                    flexibilisatie.OP_PP_Verhouding_PP = int(float(pensioengegevens[8]))
                elif pensioengegevens[6] == "Percentage":
                    flexibilisatie.OP_PP_Percentage = int(float(pensioengegevens[7]))
                
                
                #hoog-laag-constructie
                if pensioengegevens[9] == "Ja":
                    flexibilisatie.HL_Actief = True
                elif pensioengegevens[9] == "Nee":
                    flexibilisatie.HL_Actief = False
                    #volgorde
                flexibilisatie.HL_Volgorde = pensioengegevens[10]
                    #duur
                flexibilisatie.HL_Jaar = int(float(pensioengegevens[11]))
                    #methode
                flexibilisatie.HL_Methode = pensioengegevens[12]
                if pensioengegevens[12] == "Verhouding":
                    flexibilisatie.HL_Verhouding_Hoog = int(float(pensioengegevens[13]))
                    flexibilisatie.HL_Verhouding_Laag = int(float(pensioengegevens[14]))
                elif pensioengegevens[12] == "Verschil":
                    flexibilisatie.HL_Verschil = int(float(pensioengegevens[13]))       

        
def krijgpad():
    if sys.executable[-10:] == "python.exe": return sys.path[0]
    return Path(sys.executable).parent.parent.absolute()
        
        
