"""
Header
Hier komen alle libraries die in het programma gebruikt worden
"""

from datetime import datetime

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

def getDeelnemersbestand(book, verzocht = None):
    """
    functie die het deelnemersbestand inleest uit excel en dit (gedeeltelijk) terug geeft.

    Parameters
    ----------
    book : xlwings.Book
        Het excel bestand waarin het programma runned.
    verzocht : list, optioneel
        lijst met kolomnamen uit deelnemersbestand die opgevraagd worden. de standaard is None.

    Returns
    -------
    kleinDeelnemersbestand : list(list())
        Netsed list structuur van het deelnemersbestand met de opgevraagde kolommen.

    """
    deelnemersbestand = book.sheets["deelnemersbestand"].range((1,1)).expand("down").expand("right").value
    if verzocht == None: return deelnemersbestand
    kolommen = list()
    for i, item in enumerate(deelnemersbestand[0]):
        if item in verzocht: kolommen.append(i)
    kleinDeelnemersbestand = list()
    for rij in deelnemersbestand:
        nieuweRij = list()
        for i, item in enumerate(rij): 
            if i in kolommen: nieuweRij.append(item)
        kleinDeelnemersbestand.append(nieuweRij)
    return kleinDeelnemersbestand


def filterkolom(deelnemersbestand, zoekterm, kolomnaam):
    """
    functie die door de kolommen filterd om te kijken of ze aan de zoekterm voldoen.

    Parameters
    ----------
    deelnemersbestand : list(list())
        lijst met daarin de rijen van het (mogelijk verkleind) deelnemersbestand.
    zoekterm : None of str of datetime
        de waarde waarop gefilterd moet worden.
    kolomnaam : str
        de kolomnaam van de kolom waarin gefilterd moet worden.

    Returns
    -------
    kleinDeelnemersbestand : list(list)
        gefilterde versie van het deelnemersbestand.

    """
    if zoekterm == "" or zoekterm == "-" or zoekterm == datetime(1950, 1, 1): return deelnemersbestand
    for i, naam in enumerate(deelnemersbestand[0]):
        if naam == kolomnaam: 
            kolom = i
            break
    if type(zoekterm) == str:
        kleinDeelnemersbestand = [deelnemersbestand[0]]
        for rij in deelnemersbestand[1:]:
            if rij[kolom] == None: pass
            elif len(rij[kolom]) < len(zoekterm): pass
            else:
                if rij[kolom][0:len(zoekterm)] == zoekterm: kleinDeelnemersbestand.append(rij)
        deelnemersbestand = kleinDeelnemersbestand
    if type(zoekterm) == datetime:
        kleinDeelnemersbestand = [deelnemersbestand[0]]
        for rij in deelnemersbestand[1:]:
            if rij[kolom] == zoekterm: kleinDeelnemersbestand.append(rij)
    return kleinDeelnemersbestand