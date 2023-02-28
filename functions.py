"""
Header
Hier komen alle libraries die in het programma gebruikt worden
"""

import xlwings as xw
from datetime import datetime, date
from Deelnemer import Deelnemer
from Pensioenfonds import Pensioenfonds
import ctypes
import logging
import os
import sys


"""
Body
Hier komen alle functies
"""


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

def kleurinvoer(kleur):
    """
    Een functie die een string met rgb waardes veranderd naar een tuple met rgb waardes.

    Parameters
    ----------
    kleur : String
        Bevat drie rgb waardes in een string gescheiden met een ",".

    Returns
    -------
    tuple(kleuren) : Tuple(List)
        De drie rgb waardes als integer in een tuple.

    """
    rgb = kleur.split(",")
    kleuren = list()
    for i in range(len(rgb)):
        kleuren.append(int(rgb[i])/255)    
    return tuple(kleuren)

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
    List met de volledige namen van pensioenregelingen van de betreffende deelnemer
    List met de codenamen van pensioenregelingen van de betreffende deelnemer
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
    String met de codenaam van de regeling.
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
    String met de volledige naam van de regeling.
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

def getDeelnemersbestand(book):
    """
    functie die het deelnemersbestand inleest uit excel en dit terug geeft.

    Parameters
    ----------
    book : xlwings.Book
        Het excel bestand waarin het programma runned.
    
    Returns
    -------
    deelnemerlijst : list(Deelnemer)
        lijst met daarin Deelnemer objecten van alle deelnemers.

    """
    deelnemersbestand = book.sheets["deelnemersbestand"].range((1,1)).expand("down").expand("right").value
    
    deelnemersbestand[0].append("rijNr")
    for i in range(len(deelnemersbestand) - 1):
        deelnemersbestand[i+1].append(i + 2)
    pensioeninformatie = getPensioeninformatie(book)
    deelnemerlijst = list()
    for deelnemer in deelnemersbestand[1:]:
        informatie = [deelnemersbestand[0], deelnemer]
        deelnemerlijst.append(Deelnemer(informatie, pensioeninformatie))
    return deelnemerlijst


def getPensioeninformatie(book):
    kolommen = dict()
    kolommen["naamkolom"] = 2
    kolommen["pensioenleeftijdkolom"] = 4
    kolommen["rentekolom"] = 5
    kolommen["sterftetafelkolom"] = 6
    
    gegevens_pensioenenSheet = book.sheets["Gegevens pensioencontracten"]
    
    pensioenen = dict()
    # pensioenen["ZL"] = ((9, None), 3)
    pensioenen["Aegon65"] = ((10, None), 4)
    pensioenen["Aegon67"] = ((11, None), 5)
    pensioenen["NN65"] = ((12, 13), 6)
    pensioenen["NN67"] = ((14, 15), 7)
    pensioenen["PF_VLC68"] = ((16, 17), 8)
    
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
    if zoekterm == "" or zoekterm == "-" or zoekterm == datetime(1950, 1, 1): return deelnemerlijst
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


def ToevoegenDeelnemer(gegevens):
    """
    Functie die een deelnemer toevoegt onderaan het deelnemersbestand
    input : lijst met gegevens van de deelnemer
    output : deelnemer toegevoegd aan excel-bestand
    """
    
    book = xw.Book.caller() #xw.Book("Main.xlsm")
    deelnemersbestand = book.sheets["deelnemersbestand"]
    
    #check wat eerstvolgende lege regel is
    #bereken het aantal deelnemers door het aantal volle rijen na 1e regel te tellen
    aantalDeelnemers = len(deelnemersbestand.cells(1,1).expand().value)
    legeRegel = aantalDeelnemers + 1
     
    #gegevens deelnemer invullen in de lege regel
    deelnemersbestand.cells(legeRegel, 1).value = gegevens


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
        DESCRIPTION.

    """
    returnValue = ctypes.windll.user32.MessageBoxW(0, text, title, style)
    if returnValue == 0:
        raise Exception('Oops')
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
    

def gegevenscontrole(gegevens):
    """
    functie die een messagebox met alle gegevens van een deelnemer maakt

    Parameters
    ----------
    gegevens : list
        een lijst met alle gegevens van de deelnemer in de vorm:
            ["Achternaam", "tussen", "voor", "geboorte", "geslacht", "burg", "fulltimeLoon", "PT% als kans", "regeling", "Zl", "Aegon", "Aegon", "NN", "NN", "NN", "NN", "VLC", "VLC"]

    Returns
    -------
    string "correct" of "fout"
        Bij "correct" zijn de gegevens goed, bij "fout" zijn niet alle gegevens goed

    """
    #het parttime percentage in de vorm van een percentage zetten
    gegevens[7] = str(gegevens[7] * 100)
    #alle gegevens omzetten naar een string
    for g in range(0,len(gegevens)):
        gegevens[g] = str(gegevens[g])
        
    invoer = [] #lijst met alle deelnemersgegevens met uitleg in juiste volgorde
    invoer.append("Naam: " + gegevens[2] + " " + gegevens[1] + " " + gegevens[0])
    invoer.append("Geboortedatum: " + gegevens[3])
    invoer.append("Geslacht: " + gegevens[4])
    invoer.append("Burgerlijke staat: " + gegevens[5])
    invoer.append("Fulltime loon: €" + gegevens[6])
    invoer.append("Parttime percentage: " + gegevens[7] + "%")
    invoer.append("Huidige regeling: " + gegevens[8])
    invoer.append("\n")     #lege regel tussenvoegen voor opgebouwde pensioenen
    
    #opgebouwde pensioenen toevoegen, als deze ingevuld zijn
    teller = 0  #bijhouden welk pensioen het is
    for p1 in gegevens[9:12]:
        if p1 != "":
            regeling = ["ZL: ", "Aegon 65: ", "Aegon 67: "][teller]
            invoer.append(regeling + "OP = €" + p1)
        teller += 1
    
    teller = 0
    for p2 in range(12,17,2):
        if gegevens[p2] != "":
            regeling = ["NN 65: ", "NN 67: ", "PF VLC 68: "][teller]
            invoer.append(regeling + "OP = €" + gegevens[p2] + " en PP = €" + gegevens[p2+1])
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
    logger = logging.getLogger(name)

    logger.setLevel(logging.DEBUG)
    today = date.today().strftime("%Y_%m_%d")
    os.makedirs(os.path.dirname("{}\\Logs\\{}.log".format(sys.path[0], today)))
    os.makedirs(os.path.dirname("{}\\Logs\\Errors\\{}.log".format(sys.path[0], today)))
    filename = "{}\\Logs\\{}.log".format(sys.path[0], today)
    errorname = "{}\\Logs\\Errors\\{}.log".format(sys.path[0], today)
    
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


