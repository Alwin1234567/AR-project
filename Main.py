"""
Header
Hier komen alle libraries die in het programma gebruikt worden
"""
import xlwings as xw
import matplotlib.pyplot as plt
import sys
from PyQt5 import QtWidgets
from string import ascii_uppercase
import functions
import Klassen_Schermen
from logging import getLogger


"""
Body
Hier komen alle functies
"""
@xw.sub
def Schermen():
    logger = functions.setup_logger("Main") if not getLogger("Main").hasHandlers() else getLogger("Main")
    app = 0
    app = QtWidgets.QApplication(sys.argv)
    window = Klassen_Schermen.Functiekeus(xw.Book.caller(), logger)
    window.show()
    app.exec_()



@xw.sub
def vergelijken_afbeelding_generatie():
    """
    Functie die de data leest en vevolgens een afbeelding genereerd op basis van de data
    """
    
    #sheets en book opslaan in variabelen
    book = xw.Book.caller()
    invoer = book.sheets["Tijdelijk afbeelding"]
    uitvoer = book.sheets["Vergelijken"]
    
    #data lezen
    titel = (3,1)
    
    beginrij = 5
    OPbeginkolom = 2
    blokafstand = 7
    
    PPbeginkolom = 5
    PPblokafstand = 3
    
    PPnaam = 0
    PPjaarbedrag = 1
    
    OPnaam = 0
    OPkleur = 1
    OPbeginjaar = 2
    OPjaarbedrag = 3
    OPhooglaaggrens = 4
    OPverhouding = 5
    
    #Aantal blokken tellen
    blokaantal = functions.blokkentellen(beginrij, OPbeginkolom, blokafstand, invoer)
    PPblokaantal = functions.blokkentellen(beginrij, PPbeginkolom, PPblokafstand, invoer)

    #Lijst met alle voorkomende jaren van OP
    allejaren = set()
    for blok in range(blokaantal):
        allejaren.add(invoer.range((beginrij + OPbeginjaar + blok * blokafstand, OPbeginkolom)).options(numbers = float).value)
        allejaren.add(invoer.range((beginrij + OPhooglaaggrens + blok * blokafstand, OPbeginkolom)).options(numbers = float).value)
    
    #geeft de breedte aan van alle hoogtes
    randen = list(allejaren)
    randen.sort()
    randen.append(randen[-1] + 10)
    
    #een lijst met alle verzekeringsnamen
    naamlijst = list()
    for blok in range(blokaantal): naamlijst.append(invoer.range((beginrij + OPnaam + blok * blokafstand, OPbeginkolom)).value)
    
    #bepaald de kleuren
    kleuren = list()
    for blok in range(blokaantal): kleuren.append(functions.kleurinvoer(invoer.range((beginrij + OPkleur + blok * blokafstand, OPbeginkolom)).value))

    #berekent de hoogte van elke staaf
    hoogtes = [[0 for i in range(len(randen)-1)]]
    ywaardes = set()
    ywaardes.add(0)
    
    for blok in range(blokaantal):
        startjaar = float(invoer.range((beginrij + OPbeginjaar + blok * blokafstand, OPbeginkolom)).options(numbers = float).value)
        toezegging = float(invoer.range((beginrij + OPjaarbedrag + blok * blokafstand, OPbeginkolom)).options(numbers = float).value)
        laaghoogverhouding = float(invoer.range((beginrij + OPverhouding + blok * blokafstand, OPbeginkolom)).options(numbers = float).value)
        alternatiefjaar = float(invoer.range((beginrij + OPhooglaaggrens + blok * blokafstand, OPbeginkolom)).options(numbers = float).value)
        
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

    #bereken PP
    PPtotaal = 0
    for blok in range(PPblokaantal):
        PPtotaal += float(invoer.range((beginrij + PPjaarbedrag + blok * PPblokaantal, PPbeginkolom)).options(numbers = float).value)
        


    #maak de afbeeling
    afbeelding = plt.figure()
    for i in range(len(hoogtes) - 1):
        plt.stairs(hoogtes[i+1],edges = randen,  baseline=hoogtes[i], fill=True, label = naamlijst[i], color = kleuren[i])
    
    plt.xticks(randen[:-1], [functions.getaltotijd(rand) for rand in randen[:-1]])
    plt.setp(plt.gca().get_xticklabels(), rotation=30, horizontalalignment='right')
    plt.yticks(ywaardes, [functions.getaltogeld(ywaarde) for ywaarde in ywaardes])

    handles, labels = plt.gca().get_legend_handles_labels()
    order = range(blokaantal-1, -1, -1)
    plt.legend([handles[idx] for idx in order],[labels[idx] for idx in order]) 
    
    plt.gca().set_title("Totale partnerpensioen: €{:.2f}".format(PPtotaal).replace(".",","))
    plt.suptitle(invoer.range(titel).value, fontweight='bold')
    
    
    uitvoer.pictures.add(afbeelding, top = uitvoer.range((3,3)).top, left = uitvoer.range((3,3)).left, height = 300, name = "testnaam")
    
    
@xw.sub
def delimage():
    book = xw.Book.caller()
    uitvoer = book.sheets["Vergelijken"]
    print(uitvoer.pictures["testnaam"].height)
    uitvoer.pictures["testnaam"].api.Delete()
    
    




@xw.sub
#Idee voor berekeningen uitvoeren: Functies schrijven
def invoer_test_klikken():
    #sheets en book opslaan in variabelen
    book = xw.Book.caller()
    invoer = book.sheets["Tijdelijk invoerscherm"]
    
    kolommen= invoer.range((1,8), (61,130))
    #Kolommen legen waar de berekeningen komen
    kolommen.clear_contents()
    
    regeling_range= invoer.range((10,1), (18,1))
    pensioenbedragen=invoer.range((10,2),(18,2))
    sterftetafel_range = invoer.range((10,3),(18,3))
    rentes= invoer.range((10,4),(18,4))
    pensioenleeftijd_range= invoer.range((10,5),(18,5))
    koopsomfactor_range= invoer.range((10,6),(18,6))
    
    basis_koopsom= invoer.range((10,7),(18,7))

    
    pensioenleeftijd=[]
    rente=[]

    
    letters=[]
    for p in range(4):
        if len(letters)>=26:
            for i in ascii_uppercase:
                letters.append( letters[p-1] + i)
        else:
            for i in ascii_uppercase:
                letters.append(i)   
    

    counter=1
    for i in range(1,10):

        if pensioenbedragen(i).value != None and pensioenbedragen(i).value != 0:
            kolom_t= (counter-1)*10+9
            kolom_jaar= (counter-1)*10+10
            kolom_leeftijd= (counter-1)*10+11
            kolom_tpx= (counter-1)*10+12
            kolom_tqx= (counter-1)*10+13
            kolom_tqx_juli= (counter-1)*10+14
            kolom_dt= (counter-1)*10+15
            kolom_dt_juli= (counter-1)*10+16
            
            rente.append(rentes(i).value)
            pensioenleeftijd.append(pensioenleeftijd_range(i).value)
            
            #'=if(' + letters[kolom_leeftijd-1] + '2+1<=119;
            # invoer.range((1, kolom_t), (2, kolom_t)).options(ndims = 2).formula = [["t"], [0]]
            # invoer.range((2, kolom_t)).formula = 0
            blok1 = list()
            blok1.append(["t", "Jaar", "Leeftijd", "tpx", "tqx", "tqx op 1 juli", "dt", "dt op 1 juli"])
            blok1.append(["0", '=year(B4)+' + str(int(pensioenleeftijd[i-1])), '=$E' + str(i+9), 1, 0, '=if({0}3<>"", (((13 - month($B$4)) * {1}2) + ((month($B$4) - 1) * {1}3)) / 12, "")'.format(letters[kolom_leeftijd - 1], letters[kolom_tqx - 1]), '=if({0}2<>"", (1+$D${1})^-{2}2, "")'.format(letters[kolom_leeftijd-1], str(i+9), letters[kolom_t-1]),\
                          '=if({0}3<>"", (1+$D${1})^-({2}2 + (month($B$4)-1)/12), "")'.format(letters[kolom_leeftijd-1], str(i + 9), letters[kolom_t-1])])
            for rij in range(59): blok1.append(['=1+{}{}'.format(letters[kolom_t-1], str(rij+2)),\
                                               '=1+{}{}'.format(letters[kolom_jaar-1], str(rij+2)),\
                                                   '=if({0}{1}<119, 1+{0}{1},"")'.format(letters[kolom_leeftijd-1], str(rij+2)),\
                                                       functions.tpxFormule(sterftetafel_range(i).value, rij, letters[kolom_leeftijd-1], letters[kolom_jaar-1], letters[kolom_tpx-1]),\
                                                           '=if({0}{1}<>"", 1-{2}{1}, "")'.format(letters[kolom_leeftijd-1], str(rij+3), letters[kolom_tpx-1]),\
                                                               '=if({0}{1}<>"", (((13 - month($B$4)) * {2}{3}) + ((month($B$4) - 1) * {2}{1})) / 12, "")'.format(letters[kolom_leeftijd - 1], str(rij + 4), letters[kolom_tqx - 1], str(rij + 3)),\
                                                                   '=if({0}{1}<>"", (1+$D${2})^-{3}{1}, "")'.format(letters[kolom_leeftijd-1], str(rij+3), str(i+9), letters[kolom_t-1]),\
                                                                       '=if({0}{1}<>"", (1+$D${2})^-({3}{4} + (month($B$4)-1)/12), "")'.format(letters[kolom_leeftijd-1], str(rij+4), str(i + 9), letters[kolom_t-1], str(rij+3))\
                                                                           ])
            # print([len(x) for x in blok1])
            invoer.range((1, kolom_t), (61, kolom_t+7)).options(ndims = 2).formula = blok1
            
            # blok2 = list()
            # blok2.append([""])
            # blok2.append([])
            # for rij in range(59): blok2.append([
            #                                         ])
            # print(blok2)
            # invoer.range((1, kolom_t+5), (61, kolom_t+6)).options(ndims = 2).formula = blok2
            # invoer.range((1, kolom_jaar)).value= "Jaar"
            # invoer.range((2, kolom_jaar)).formula= [['=year(B4)+' + str(int(pensioenleeftijd[i-1]))]]
            # invoer.range((3, kolom_jaar), (61, kolom_jaar)).formula= [['=1+' + letters[kolom_jaar-1] + '2']]
            
            # invoer.range((1, kolom_leeftijd)).value= "Leeftijd"
            # invoer.range((2, kolom_leeftijd)).formula= [['=$E' + str(i+9)]]
            # invoer.range((3, kolom_leeftijd), (61, kolom_leeftijd)).formula= [['=if(' + letters[kolom_leeftijd-1] + '2<119, 1+' + letters[kolom_leeftijd-1] + '2,"")']]
            
            # invoer.range((1, kolom_tqx)).value= "tqx"
            # invoer.range((2, kolom_tqx), (61, kolom_tqx)).formula= [['=if(' + letters[kolom_leeftijd-1] + '2<>"", 1-' + letters[kolom_tpx-1] + '2, "")']]
            
            # invoer.range((1, kolom_tqx_juli)).value= "tqx op 1 juli"
            # invoer.range((2, kolom_tqx_juli), (60, kolom_tqx_juli)).formula= [['=if(' + letters[kolom_leeftijd-1] + '3<>"", (((13-month($B$4))*' + letters[kolom_tqx-1] + '2)+((month($B$4)-1)*' + letters[kolom_tqx-1] + '3))/12, "")']]
            
            # invoer.range((1, kolom_dt)).value= "dt"
            # invoer.range((2, kolom_dt), (61, kolom_dt)).formula= [['=if(' + letters[kolom_leeftijd-1] + '2<>"", (1+$D$' + str(i+9) + ')^-' + letters[kolom_t-1] + '2, "")']]
            
            # invoer.range((1, kolom_dt_juli)).value= "dt op 1 juli"
            # invoer.range((2, kolom_dt_juli), (60, kolom_dt_juli)).formula= [['=if(' + letters[kolom_leeftijd-1] + '3<>"", (1+$D$' + str(i+9) + ')^-(' + letters[kolom_t-1] + '2+(month($B$4)-1)/12), "")']]
            
            
            # if sterftetafel_range(i).value== "AG_2020":
            #     # invoer.range((1, kolom_tpx)).value= "tpx"
            #     # invoer.range((2, kolom_tpx)).value= 1
            #     print('=if(' + letters[kolom_leeftijd-1] + '3<>"", (1-INDEX(INDIRECT($C$' + str(i+9) + '),' + letters[kolom_leeftijd-1] + '2+1, ' + letters[kolom_jaar-1] + '2-2018))*' + letters[kolom_tpx-1] + '2,"")')
            #     # invoer.range((3, kolom_tpx), (61, kolom_tpx)).formula= [['=if(' + letters[kolom_leeftijd-1] + '3<>"", (1-INDEX(INDIRECT($C$' + str(i+9) + '),' + letters[kolom_leeftijd-1] + '2+1, ' + letters[kolom_jaar-1] + '2-2018))*' + letters[kolom_tpx-1] + '2,"")']]
                
                
            # else:
            #     # invoer.range((1, kolom_tpx)).value= "tpx"
            #     print('=if(' + letters[kolom_leeftijd-1] + '2<>"", INDEX(INDIRECT($C$' + str(i+9) + '),' + letters[kolom_leeftijd-1] + '2+1,1)/ INDEX(INDIRECT($C$' + str(i+9) + '),$' + letters[kolom_leeftijd-1] + '$2+1,1),"")')
            #     # invoer.range((2, kolom_tpx), (61, kolom_tpx)).formula= [['=if(' + letters[kolom_leeftijd-1] + '2<>"", INDEX(INDIRECT($C$' + str(i+9) + '),' + letters[kolom_leeftijd-1] + '2+1,1)/ INDEX(INDIRECT($C$' + str(i+9) + '),$' + letters[kolom_leeftijd-1] + '$2+1,1),"")']]
    
            
            # if regeling_range(i).value== "ZL":
            #     koopsomfactor_range(i).value= 0
            #     basis_koopsom(i).value= pensioenbedragen(i).value
                
            # elif "OP" in regeling_range(i).value:
            #     koopsomfactor_range(i).formula= [['=SUMPRODUCT(' + letters[kolom_tpx-1] + '2:' + letters[kolom_tpx-1] + '61,' + letters[kolom_dt-1] + '2:' + letters[kolom_dt-1] + '61)']]
            #     basis_koopsom(i).value= pensioenbedragen(i).value*koopsomfactor_range(i).value
                
            # else:
            #     koopsomfactor_range(i).formula= [['=SUMPRODUCT(' + letters[kolom_tpx-1] + '2:' + letters[kolom_tpx-1] + '61,' + letters[kolom_tqx_juli-1] +'2:' + letters[kolom_tqx_juli-1] + '61,'+ letters[kolom_dt_juli-1] + '2:' + letters[kolom_dt_juli-1] + '61)']]
            #     basis_koopsom(i).value= pensioenbedragen(i).value*koopsomfactor_range(i).value
            
            
            counter+= 1
            
        else:
            basis_koopsom(i).value= 0
            rente.append(0)
            pensioenleeftijd.append(0)
            
            
@xw.sub
def AfbeeldingKiezen():
    """
    Functie die de gekozen afbeelding op de vergelijkings sheet kiest
    """
    
    #sheets en book opslaan in variabelen
    book = xw.Book.caller()
    sheet = book.sheets["Vergelijken"]
    #gekozen afbeelding inlezen
    gekozenAfbeelding = sheet.cells(6,"B").value
    #naam van gekozen afbeelding op sheet printen
    sheet.cells(8, "M").value = gekozenAfbeelding
    

@xw.sub
def AfbeeldingVerwijderen():
    """
    Functie die de gekozen afbeelding op de vergelijkings sheet verwijderd
    """
    
    #sheets en book opslaan in variabelen
    book = xw.Book.caller()
    sheet = book.sheets["Vergelijken"]
    #gekozen afbeelding inlezen
    gekozenAfbeelding = sheet.cells(6,"B").value
    #naam van gekozen afbeelding op sheet printen
    sheet.cells(11, "M").value = gekozenAfbeelding
    
    #gekozen afbeelding verwijderen
    #sheet.pictures[naam].delete()

@xw.sub
def AfbeeldingAanpassen():
    """
    Functie die de gekozen afbeelding op de vergelijkings sheet als basis neemt voor nieuwe flexibilisaties
    """
    
    #sheets en book opslaan in variabelen
    book = xw.Book.caller()
    sheet = book.sheets["Vergelijken"]
    #gekozen afbeelding inlezen
    gekozenAfbeelding = sheet.cells(6,"B").value
    #naam van gekozen afbeelding op sheet printen
    sheet.cells(14, "M").value = gekozenAfbeelding
           
        
            
@xw.sub
def flexibilisaties_testen():
   book = xw.Book.caller()
   invoer = book.sheets["Tijdelijk invoerscherm"] 
   
   flexibilisaties= invoer.range((21,2), (55,6))
   
   regeling_range= invoer.range((10,1), (18,1))
   pensioenleeftijd= invoer.range((10,5), (18,5))
   koopsomfactor_range= invoer.range((10,6), (18,6))
   basis_koopsom= invoer.range((10,7), (18,7))

   koopsommen= []
   koopsommen= basis_koopsom.value
   
   for i in range(1,10):
       regeling= (i-1)*8+1
       soort= (i-1)*8+2
       hoeveelheid= (i-1)*8+3
       factor= (i-1)*8+4
       aanspraak_OP= (i-1)*8+5
       aanspraak_PP= (i-1)*8+6
       
       for c in range(1,5):
           if flexibilisaties(regeling, c).value!= None:
               pass
               
               
               
               
               
               
               
           # if flexibilisaties(regeling).value[:-2] in invoer.range((11,1)).value and flexibilisaties(regeling).value[-2:] in invoer.range((11,1)).value:
           #     print("test")
           #     # if flexibilisaties(2).value== "vervroegen" or flexibilisaties(2).value== "verlaten":
           #     #     #Bij corresponderende pensioenregelingsrij de hoeveelheid vervroegen/verlaten optellen
           #     #     pensioenleeftijd(2).value= pensioenleeftijd(2).value+ flexibilisaties(3).value
                   
                   
                   
           #     # elif  flexibilisaties(2).value== "uitruilen":
           #     #     print("Uitruilen")
                   
           #     # elif flexibilisaties(2).value== "hoog" or flexibilisaties(2).value== "laag":
           #     #      print("Hoogconstructie of laagconstructie") 
                    
           #     # else:
           #     #     print("AOW overbruggen")
                   
           # else:
           #     print("Het werkt niet")
       
       
       
       
       
       
       
   

   
   
   
       
       
       
       
       
       
       
       