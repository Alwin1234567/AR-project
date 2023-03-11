"""
Header
Hier komen alle libraries die in het programma gebruikt worden
"""

import xlwings as xw
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
    kolommen["pensioenleeftijdkolom"] = 4
    kolommen["rentekolom"] = 5
    kolommen["sterftetafelkolom"] = 6
    kolommen["Kleurzachtkolom"] = 10
    kolommen["Kleurhardkolom"] = 11
    
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
    if not exists("{}\\Logs\\{}.log".format(sys.path[0], today)): os.makedirs(os.path.dirname("{}\\Logs\\{}.log".format(sys.path[0], today)), exist_ok=True)
    if not exists("{}\\Logs\\Errors\\{}.log".format(sys.path[0], today)): os.makedirs(os.path.dirname("{}\\Logs\\Errors\\{}.log".format(sys.path[0], today)), exist_ok=True)
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

def isInteger(veldInput):
    try:
        veldInput = int(veldInput)
        return True
    except:
        return False
        
def checkVeldInvoer(methode,veld1,veld2,veld3):
    intProblem = False
    emptyProblem = False
    
    if str(methode) == "Percentage" or str(methode) == "Verschil":
        if str(veld1) == "":
            emptyProblem = True
        elif isInteger(veld1) == False:
            intProblem = True
        
        if str(veld2) == "":
            pass
        elif isInteger(veld2) == False:
            intProblem = True
            
        if str(veld3) == "":
            pass
        elif isInteger(veld3) == False:
            intProblem = True  
        
        
    elif str(methode) == "Verhouding":
        if str(veld1) == "":
            pass
        elif isInteger(veld1) == False:
            intProblem = True
        
        if str(veld2) == "":
            emptyProblem = True
        elif isInteger(veld2) == False:
            intProblem = True
            
        if str(veld3) == "":
            emptyProblem = True
        elif isInteger(veld3) == False:
            intProblem = True       
    
    elif str(methode) == "Opvullen AOW":
        if isInteger(veld1) == False:
            intProblem = True

        if isInteger(veld2) == False:
            intProblem = True

        if isInteger(veld3) == False:
            intProblem = True   


    if intProblem == True and emptyProblem == True:
        return ["Er is foute invoer en missende invoer.",False]
    elif intProblem == True and emptyProblem == False:
        return ["Invoer mag alleen een geheel getal zijn.",False]
    elif intProblem == False and emptyProblem == True:
        return ["Er is missende invoer.",False]
    else:
        return ["",True]
    
def tpxFormule(sterftetafel, rij, leeftijdKolomLetter, jaarKolom, tpxKolom):
    if sterftetafel == "AG_2020": return '=if({0}{1}<>"", (1-INDEX(INDIRECT("{2}"),{0}{3}+1,{4}{3}-2018))*{5}{3},"")'.format(leeftijdKolomLetter, rij + 3,  sterftetafel, rij + 2, jaarKolom, tpxKolom)
    else: return '=if({0}{1}<>"", INDEX(INDIRECT("{2}"),{0}{1}+1,1)/ INDEX(INDIRECT("{2}"),${0}$2+1,1),"")'.format(leeftijdKolomLetter, rij + 3, sterftetafel)

def persoonOpslag(book, persoonObject):
    """
    Functie die persoonsgegevens opslaat in de flexopslag sheet.
    
    Parameters
    ----------
    book : xlwings.Book
        Het excel bestand waarin het programma runned.
    
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
    persopslag[8][1] = str(persoonObject.achternaam)
    
    persopslag[9][0] = "Rij nr"
    persopslag[9][1] = persoonObject.rijNr
    
    book.range((6,1),(15,2)).options(ndims = 2).value = persopslag
    book.range((6,1),(15,2)).api.Interior.Color = rgb_to_int((150,150,150))


def flexOpslag(book,flexibilisatie,countOpslaan,countRegeling):
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
    
    flexopslag[16][0] = "Jaarbedrag"
    
    flexopslag[18][0] = "Kleur"
    
    # Pensioennaam invullen
    flexopslag[0][1] = str(flexibilisatie.pensioen.pensioenNaam)
    
    # Pensioenleeftijd wijzigen J/N
    if flexibilisatie.leeftijd_Actief: flexopslag[2][1] = "J"
    else: flexopslag[2][1] = "N"
    
    # Pensioenleeftijd: Jaar & Maand
    flexopslag[3][1] = flexibilisatie.leeftijdJaar
    flexopslag[3][2] = flexibilisatie.leeftijdMaand
    
    # OP/PP Uitruilen wijzigen J/N
    if flexibilisatie.OP_PP_Actief: flexopslag[5][1] = "J"
    else: flexopslag[5][1] = "N"
    
    # OP/PP uitruiling opslaan
    if flexibilisatie.OP_PP_UitruilenVan == "OP naar PP": flexopslag[6][1] = "OP/PP"
    elif flexibilisatie.OP_PP_UitruilenVan == "PP naar OP": flexopslag[6][1] = "PP/OP"
    
    if flexibilisatie.OP_PP_Methode == "Verhouding":
        flexopslag[7][1] = "Verh"
        flexopslag[8][1] = flexibilisatie.OP_PP_Verhouding_OP
        flexopslag[8][2] = flexibilisatie.OP_PP_Verhouding_PP
    elif flexibilisatie.OP_PP_Methode == "Percentage":
        flexopslag[7][1] = "Perc"
        flexopslag[8][1] = flexibilisatie.OP_PP_Percentage
    else:
        logger.info("OP/PP methode wordt niet herkend bij opslaan naar excel.")
    
    # Hoog/Laag constructie opslaan
    if flexibilisatie.HL_Actief: flexopslag[9][1] = "J"
    else: flexopslag[10][1] = "N"
    
    if flexibilisatie.HL_Volgorde == "Hoog-laag": flexopslag[11][1] = "Hoog/Laag"
    elif flexibilisatie.HL_Volgorde == "Laag-hoog": flexopslag[11][1] = "Laag/Hoog"
    
    flexopslag[12][1] = flexibilisatie.HL_Jaar
    
    if flexibilisatie.HL_Methode == "Verhouding":
        flexopslag[13][1] = "Verh"
        flexopslag[14][1] = flexibilisatie.HL_Verhouding_Hoog
        flexopslag[15][2] = flexibilisatie.HL_Verhouding_Laag
    elif flexibilisatie.HL_Methode == "Verschil":
        flexopslag[13][1] = "Verh"
        flexopslag[14][1] = flexibilisatie.HL_Verschil
    elif flexibilisatie.HL_Methode == "Opvullen AOW":
        flexopslag[13][1] = "Opv"
    else:
        logger.info("H/L methode wordt niet herkend bij opslaan naar excel.")
    
    # Nieuwe OP en PP opslaan
    flexopslag[16][1] = "OP Onbekend"
    flexopslag[16][2] = "PP Onbekend"
    
    # RGB opslaan
    flexopslag[18][1] = str(flexibilisatie.pensioen.pensioenKleurHard)
    
    # Waardes in sheet plakken & celkleur instellen
    book.range((5+20*countRegeling,4+4*countOpslaan),(23+20*countRegeling,6+4*countOpslaan)).options(ndims = 2).value = flexopslag
    book.range((5+20*countRegeling,4+4*countOpslaan),(23+20*countRegeling,6+4*countOpslaan)).api.Interior.Color = rgb_to_int(flexibilisatie.pensioen.pensioenKleurHard)


def zoekRGB(book,regeling):
    i = 1
    rgb = "Geen rgb gevonden."
    
    while i < 11:
        if str(book.sheets["Gegevens pensioencontracten"].range(i,2).value) == regeling:
            rgb = str(book.sheets["Gegevens pensioencontracten"].range(i,10).value)
        i += 1
    
    return rgb
    
def berekeningen_init(sheet, deelnemer, logger):
    pensioeninfohoogte = 2
    pensioeninfokolom = 1
    pensioenblokkolom = 1

    afstandtotblokken = 6
    afstandtussenblokken = 2
    blokgrootte = 17
    
    afstandtotrekenkolom = 8
    afstandtussenrekenblokken = 1
    rekenblokgrootte = 63
    rekenblokbreedte = 8
    
    aantalpensioenen = len(deelnemer.flexibilisaties)
    
    
    # pensioen info
    for i, flexibilisatie in enumerate(deelnemer.flexibilisaties):
        pensioeninfo = list()
        pensioeninfo.append(flexibilisatie.pensioen.pensioenVolNaam)
        pensioeninfo.append("verwijzing/ formule")
        pensioeninfo.append("verwijzing/ formule")
        pensioeninfo.append(flexibilisatie.pensioen.pensioenleefijd)
        inforange = sheet.range((pensioeninfohoogte + i, pensioeninfokolom),\
                            (pensioeninfohoogte + i, pensioeninfokolom + len(pensioeninfo) - 1))
        inforange.value = pensioeninfo
        inforange.api.Interior.Color = rgb_to_int(flexibilisatie.pensioen.pensioenKleurHard)
    
    # pensioen blok
    for i, flexibilisatie in enumerate(deelnemer.flexibilisaties):
        blokhoogte = pensioeninfohoogte + afstandtotblokken + aantalpensioenen + i * (blokgrootte + afstandtussenblokken)
        rekenblokstart = afstandtotrekenkolom + i * (rekenblokbreedte + afstandtussenrekenblokken)
        blok = list()
        blok.append(["Naam", flexibilisatie.pensioen.pensioenVolNaam, "", ""])
        if flexibilisatie.leeftijd_Actief: blok.append(["Start Pensioenjaar", flexibilisatie.leeftijdJaarMaand, "", ""])
        else: blok.append(["Start Pensioenjaar", flexibilisatie.pensioen.pensioenleefijd, "", ""])
        if flexibilisatie.OP_PP_Actief:
            rij1 = list()
            rij2 = list()
            rij1.append("Uitruilen soort")
            rij2.append("Uitruilen waarde")
            if flexibilisatie.OP_PP_Methode == "Percentage":
                rij1.append("{} {}".format(flexibilisatie.OP_PP_UitruilenVan, flexibilisatie.OP_PP_Methode))
                rij2.append(flexibilisatie.OP_PP_Percentage)
                rij2.append("")
            else:
                rij1.append(flexibilisatie.OP_PP_Methode)
                rij2.append("1")
                rij2.append(flexibilisatie.OP_PP_Verhouding_PP / flexibilisatie.OP_PP_Verhouding_OP)
            rij1.append("")
            rij1.append("")
            rij2.append("")
            blok.append(rij1)
            blok.append(rij2)
        else:
            blok.append(["Uitruilen soort", "", "", ""])
            blok.append(["Uitruilen waarde", "", "", ""])
        
        if flexibilisatie.HL_Actief:
            blok.append(["Hoog Laag", flexibilisatie.HL_Methode, "", ""])
            rij = list()
            rij.append("Hoog Laag waarde")
            rij.append(flexibilisatie.HL_Jaar)
            if flexibilisatie.HL_Methode == "Verhouding":
                if flexibilisatie.HL_Volgorde == "Hoog-laag": rij.append(flexibilisatie.HL_Verhouding_Laag / flexibilisatie.HL_Verhouding_Hoog)
                else: rij.append(flexibilisatie.HL_Verhouding_Hoog / flexibilisatie.HL_Verhouding_Laag) 
            else:
                if flexibilisatie.HL_Volgorde == "Hoog-laag": rij.append(flexibilisatie.HL_Verschil)
                else: rij.append(-1 * flexibilisatie.HL_Verschil) 
            rij.append("")
            blok.append(rij)
        else:
            blok.append(["Hoog Laag", "", "", ""])
            blok.append(["Hoog Laag waarde", "", "", ""])    
        
        blok.append(["", "", "", ""])
        
        blok.append(["OP / PP", flexibilisatie.pensioen.ouderdomsPensioen, flexibilisatie.pensioen.partnerPensioen, "=B{0} * B{1} + C{0} * B{2}".format(blokhoogte + 7, blokhoogte + 11, blokhoogte + 13)])
        blok.append(["OP' / PP'", '=B{0} * B{1} / B{2}'.format(blokhoogte + 7, blokhoogte + 11, blokhoogte + 12),\
                     '=C{0} * B{1} / B{2}'.format(blokhoogte + 7, blokhoogte + 13, blokhoogte + 14), "formuletekst"])
        blok.append(["OP'' / PP''", '=IF(B{0} =  "", B{5}, IF(B{0} = "Verhouding", ROUND(D{1} /  (B{2} * B{3} + C{2} *  B{4}), 0), IF(B{0} = "OP naar PP Percentage", ROUND(B{5} * (1 - B{2}), 0), ROUND(B{5} + C{5} * B{2} * B{4} / B{3}, 0))))'.format(blokhoogte + 2, blokhoogte + 7, blokhoogte + 3, blokhoogte + 12, blokhoogte + 14, blokhoogte + 8),\
                     '=IF(B{0} =  "", C{5}, IF(B{0} = "Verhouding", ROUND(C{2} * D{1} /  (B{2} * B{3} + C{2} *  B{4}), 0), IF(B{0} = "OP naar PP Percentage", ROUND(C{5} + B{5} * B{2} * B{3} / B{4}, 0), ROUND(C{5} * (1 - B{2}), 0))))'.format(blokhoogte + 2, blokhoogte + 7, blokhoogte + 3, blokhoogte + 12, blokhoogte + 14, blokhoogte + 8), "formuletekst"])
        blok.append(["OP''H / PP''L", "formule", "formule", "formuletekst"])
        
        blok.append(["rode a", '=ROUND(SUMPRODUCT(INDIRECT("{0}"& MAX(D{2} - B{3} + 3, 3)):{0}{4}, INDIRECT("{1}"& MAX(D{2} - B{3} + 3, 3)):{1}{4}), 3)'.format(inttoletter(rekenblokstart + 3), inttoletter(rekenblokstart + 6), pensioeninfohoogte + i, blokhoogte + 1, rekenblokgrootte), "", "formule"])
        blok.append(["zwarte a", '=ROUND(SUMPRODUCT(INDIRECT("{0}"& MAX(B{3} - D{2} + 3, 3)):{0}{4}, INDIRECT("{1}"& MAX(B{3} - D{2} + 3, 3)):{1}{4}), 3)'.format(inttoletter(rekenblokstart + 3), inttoletter(rekenblokstart + 6), pensioeninfohoogte + i, blokhoogte + 1, rekenblokgrootte), "", "formule"])
        blok.append(["PP rode a", '=ROUND(SUMPRODUCT(INDIRECT("{0}"& MAX(D{3} - B{4} + 3, 3)):{0}{5}, INDIRECT("{1}"& MAX(D{3} - B{4} + 3, 3)):{1}{5}, INDIRECT("{2}"& MAX(D{3} - B{4} + 3, 3)):{2}{5}), 3)'.format(inttoletter(rekenblokstart + 3), inttoletter(rekenblokstart + 5), inttoletter(rekenblokstart + 7), pensioeninfohoogte + i, blokhoogte + 1, rekenblokgrootte), "", "formuletekst"])
        blok.append(["groene a", '=ROUND(SUMPRODUCT(INDIRECT("{0}"& MAX(B{4} - D{3} + 3, 3)):{0}{5}, INDIRECT("{1}"& MAX(B{4} - D{3} + 3, 3)):{1}{5}, INDIRECT("{2}"& MAX(B{4} - D{3} + 3, 3)):{2}{5}), 3)'.format(inttoletter(rekenblokstart + 3), inttoletter(rekenblokstart + 5), inttoletter(rekenblokstart + 7), pensioeninfohoogte + i, blokhoogte + 1, rekenblokgrootte), "", "formuletekst"])
        blok.append(["m|zwarte a", "formule", "", "formuletekst"])
        blok.append(["zwarte a (m-1)|", "formule", "", "formuletekst"])
        
        if sum([len(rij) for rij in blok]) == len(blok) * 4:
            blokruimte = sheet.range((blokhoogte, pensioenblokkolom),\
                                     (blokhoogte + blokgrootte - 1, pensioenblokkolom + len(blok[0]) - 1)).options(ndims = 2)
            # geldblok = sheet.range((blokhoogte + 7, pensioenblokkolom + 1),\
            #                          (blokhoogte + 10, pensioenblokkolom + 2))
            # geldblok.api.NumberFormat = "Currency"
            blokruimte.formula = blok
            blokruimte.api.Interior.Color = rgb_to_int(flexibilisatie.pensioen.pensioenKleurHard)
        else:
            logger.warning("berekeningen pensioenblok niet allemaal gelijk")
            logger.debug([len(rij) for rij in blok])
        
        # rekenblok header
    for i, flexibilisatie in enumerate(deelnemer.flexibilisaties):
        blokhoogte = pensioeninfohoogte + afstandtotblokken + aantalpensioenen + i * (blokgrootte + afstandtussenblokken)
        rekenblokstart = afstandtotrekenkolom + i * (rekenblokbreedte + afstandtussenrekenblokken)
        blok = list()
        blok.append([flexibilisatie.pensioen.pensioenVolNaam] + [""] * 7)
        blok.append(["t", "jaar", "Leeftijd", "tpx", "tqx", "tqx op 1 juli", "dt", "dt op 1 juli"])
        rij = list()
        rij.append("0")
        rij.append("={} + {}3".format(deelnemer.geboortedatum.year, inttoletter(rekenblokstart + 2)))
        rij.append("=min(D{},B{})".format(pensioeninfohoogte + i, blokhoogte + 1))
        rij.append("1")
        rij.append('=if({0}3<>"", 1-{0}3, "")'.format(inttoletter(rekenblokstart + 3)))
        rij.append('=if({0}4<>"", (((13 - {2}) * {1}3) + (({2}) - 1) * {1}4) / 12, "")'.format(inttoletter(rekenblokstart + 2), inttoletter(rekenblokstart + 4), deelnemer.geboortedatum.month))
        rij.append('=if({0}3<>"", (1+{1})^-{2}3, "")'.format(inttoletter(rekenblokstart + 2), flexibilisatie.pensioen.rente / 100, inttoletter(rekenblokstart)))
        rij.append('=if({0}4<>"", (1+{1})^-({2}3 + (7 - {3}) / 12), "")'.format(inttoletter(rekenblokstart + 2), flexibilisatie.pensioen.rente / 100, inttoletter(rekenblokstart), deelnemer.geboortedatum.month))
        blok.append(rij)
        
        if sum([len(rij) for rij in blok]) == len(blok) * 8:
            blokruimte = sheet.range((1, afstandtotrekenkolom + i * (len(blok[0]) + afstandtussenrekenblokken )),\
                                     (3, afstandtotrekenkolom + i * (len(blok[0]) + afstandtussenrekenblokken ) + 7))
            mergeruimte = sheet.range((1, afstandtotrekenkolom + i * (len(blok[0]) + afstandtussenrekenblokken )),\
                                     (1, afstandtotrekenkolom + i * (len(blok[0]) + afstandtussenrekenblokken ) + 7))
            blokruimte.formula = blok
            blokruimte.api.Interior.Color = rgb_to_int(flexibilisatie.pensioen.pensioenKleurHard)
            mergeruimte.merge()
            mergeruimte.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
        else:
            logger.warning("berekeningen rekenblok niet allemaal gelijk")
            logger.debug([len(rij) for rij in blok])
            
    # rekenblok Body
    for i, flexibilisatie in enumerate(deelnemer.flexibilisaties):
        rekenblokstart = afstandtotrekenkolom + i * (len(blok[0]) + afstandtussenrekenblokken)
        rij = list()
        rij.append("={}3 + 1".format(inttoletter(rekenblokstart)))
        rij.append("={}3 + 1".format(inttoletter(rekenblokstart + 1)))
        rij.append('=if({0}3<119,{0}3 + 1,"")'.format(inttoletter(rekenblokstart + 2)))        
        if flexibilisatie.pensioen.sterftetafel == "AG_2020": rij.append('=if({0}4<>"", (1-INDEX(INDIRECT("{1}"),{0}3+1,{2}3-2018))*{3}3,"")'.format(inttoletter(rekenblokstart + 2),  flexibilisatie.pensioen.sterftetafel, rekenblokstart + 1, rekenblokstart + 3))
        else: rij.append('=if({0}4<>"", INDEX(INDIRECT("{1}"),{0}4+1,1)/ INDEX(INDIRECT("{1}"),${0}$3+1,1),"")'.format(inttoletter(rekenblokstart + 2), flexibilisatie.pensioen.sterftetafel))
        rij.append('=if({0}4<>"", 1-{0}4, "")'.format(inttoletter(rekenblokstart + 3)))
        rij.append('=if({0}5<>"", (((13 - {2}) * {1}4) + (({2}) - 1) * {1}5) / 12, "")'.format(inttoletter(rekenblokstart + 2), inttoletter(rekenblokstart + 4), deelnemer.geboortedatum.month))
        rij.append('=if({0}4<>"", (1+{1})^-{2}4, "")'.format(inttoletter(rekenblokstart + 2), flexibilisatie.pensioen.rente / 100, inttoletter(rekenblokstart)))
        rij.append('=if({0}5<>"", (1+{1})^-({2}4 + (7 - {3}) / 12), "")'.format(inttoletter(rekenblokstart + 2), flexibilisatie.pensioen.rente / 100, inttoletter(rekenblokstart), deelnemer.geboortedatum.month))
        
        blokruimte = sheet.range((4, afstandtotrekenkolom + i * (len(rij) + afstandtussenrekenblokken )),\
                                 (max(4, rekenblokgrootte), afstandtotrekenkolom + i * (len(rij) + afstandtussenrekenblokken ) + 7))
        blokruimte.formula = rij
        blokruimte.api.Interior.Color = rgb_to_int(flexibilisatie.pensioen.pensioenKleurHard)
        

def inttoletter(getal):
    letter = ""
    while True:
        if getal > 26:
            letter = "{}{}".format(ascii_uppercase[(getal%26) - 1], letter)
            getal = (getal-getal%26)//26 - (1 - min(getal%26, 1))
        else: return "{}{}".format(ascii_uppercase[(getal%26) - 1], letter)


def berekeningen_update():
    pass
    