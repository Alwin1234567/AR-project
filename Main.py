"""
Header
Hier komen alle libraries die in het programma gebruikt worden
"""
import xlwings as xw
import datetime as dt
import matplotlib.pyplot as plt


"""
Body
Hier komen alle functies
"""


@xw.sub
def vergelijken_afbeelding_generatie():
    """
    Functie die de data leest en vevolgens een afbeelding genereerd op basis van de data
    """
    
    #sheets en book opslaan in variabelen
    book = xw.Book.caller()
    invoer = book.sheets["Tijdelijk"]
    uitvoer = book.sheets["Vergelijken"]
    
    #data lezen
    beginrij = 5
    beginkolom = 2
    blokafstand = 6
    
    naam = 0
    beginjaar = 1
    jaarbedrag = 2
    hooglaaggrens = 3
    verhouding = 4
    
    #aantal blokken tellen
    blokaantal = 0
    leescell = [beginrij, beginkolom]
    while invoer.range(tuple(leescell)).value != None:
        blokaantal +=1
        leescell[0] +=blokafstand

    allejaren = set()
    for blok in range(blokaantal):
        allejaren.add(invoer.range((beginrij + beginjaar + blok * blokafstand, beginkolom)).options(numbers = float).value)
        allejaren.add(invoer.range((beginrij + hooglaaggrens + blok * blokafstand, beginkolom)).options(numbers = float).value)
    
    #geeft de breedte aan van alle hoogtes
    randen = list(allejaren)
    randen.sort()
    
    #een lijst met alle verzekeringsnamen
    naamlijst = list()
    for blok in range(blokaantal): naamlijst.append(invoer.range((beginrij + naam + blok * blokafstand, beginkolom)).value)
    
    #berekent de hoogte van elke staaf
    hoogtes = [[0 for i in range(len(randen)-1)]]
    ywaardes = set()
    ywaardes.add(0)
    
    for blok in range(blokaantal):
        startjaar = float(invoer.range((beginrij + beginjaar + blok * blokafstand, beginkolom)).options(numbers = float).value)
        toezegging = float(invoer.range((beginrij + jaarbedrag + blok * blokafstand, beginkolom)).options(numbers = float).value)
        laaghoogverhouding = float(invoer.range((beginrij + verhouding + blok * blokafstand, beginkolom)).options(numbers = float).value)
        alternatiefjaar = float(invoer.range((beginrij + hooglaaggrens + blok * blokafstand, beginkolom)).options(numbers = float).value)
        
        hoogtes.append(list())

        for i, leeftijd in enumerate(randen[:-1]):
            if leeftijd >= alternatiefjaar:
                bedrag = float(hoogtes[blok][i] + toezegging * laaghoogverhouding)
                hoogtes[blok+1].append(bedrag)
                ywaardes.add(bedrag)
            elif leeftijd >= startjaar:
                bedrag = float(hoogtes[blok][i] + toezegging)
                hoogtes[blok+1].append(bedrag)
                ywaardes.add(bedrag)
            else: hoogtes[blok+1].append(hoogtes[blok][i])
    ywaardes = list(ywaardes)
    ywaardes.sort()
    #maak de afbeeling
    afbeelding = plt.figure()
    for i in range(len(hoogtes) - 1):
        plt.stairs(hoogtes[i+1],edges = randen,  baseline=hoogtes[i], fill=True, label = naamlijst[i])
    
    plt.xticks(randen[:-1], [getaltotijd(rand) for rand in randen[:-1]])
    plt.setp(plt.gca().get_xticklabels(), rotation=30, horizontalalignment='right')
    plt.yticks(ywaardes, [getaltogeld(ywaarde) for ywaarde in ywaardes])

    handles, labels = plt.gca().get_legend_handles_labels()
    order = range(blokaantal-1, -1, -1)
    plt.legend([handles[idx] for idx in order],[labels[idx] for idx in order]) 
    
    uitvoer.pictures.add(afbeelding)
    
    

def getaltotijd(getal):
    jaar = int(getal)
    maand = round((getal - jaar) * 12)
    tijd = "{}j".format(jaar)
    if maand > 0: tijd = tijd + " {}m".format(maand)
    return tijd

def getaltogeld(getal): return "€{:.2f}".format(float(getal))
    