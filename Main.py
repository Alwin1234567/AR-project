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
    # ywaardes.add(0)
    
    # for blok in range(blokaantal):
    #     startjaar = float(invoer.range((beginrij + OPbeginjaar + blok * blokafstand, OPbeginkolom)).options(numbers = float).value)
    #     toezegging = float(invoer.range((beginrij + OPjaarbedrag + blok * blokafstand, OPbeginkolom)).options(numbers = float).value)
    #     laaghoogverhouding = float(invoer.range((beginrij + OPverhouding + blok * blokafstand, OPbeginkolom)).options(numbers = float).value)
    #     alternatiefjaar = float(invoer.range((beginrij + OPhooglaaggrens + blok * blokafstand, OPbeginkolom)).options(numbers = float).value)
        
    #     hoogtes.append(list())

    #     for i, leeftijd in enumerate(randen[:-1]):
    #         if leeftijd >= alternatiefjaar:
    #             bedrag = float(hoogtes[blok][i] + toezegging * laaghoogverhouding)
    #             hoogtes[blok+1].append(bedrag)
    #             ywaardes.add(bedrag)
    #         elif leeftijd >= startjaar:
    #             bedrag = float(hoogtes[blok][i] + toezegging)
    #             hoogtes[blok+1].append(bedrag)
    #             ywaardes.add(bedrag)
    #         else: hoogtes[blok+1].append(hoogtes[blok][i])
    # ywaardes = list(ywaardes)
    # ywaardes.sort()

    # #bereken PP
    # PPtotaal = 0
    # for blok in range(PPblokaantal):
    #     PPtotaal += float(invoer.range((beginrij + PPjaarbedrag + blok * PPblokaantal, PPbeginkolom)).options(numbers = float).value)
        


    # #maak de afbeeling
    # afbeelding = plt.figure()
    # for i in range(len(hoogtes) - 1):
    #     plt.stairs(hoogtes[i+1],edges = randen,  baseline=hoogtes[i], fill=True, label = naamlijst[i], color = kleuren[i])
    
    # plt.xticks(randen[:-1], [functions.getaltotijd(rand) for rand in randen[:-1]])
    # plt.setp(plt.gca().get_xticklabels(), rotation=30, horizontalalignment='right')
    # plt.yticks(ywaardes, [functions.getaltogeld(ywaarde) for ywaarde in ywaardes])

    # handles, labels = plt.gca().get_legend_handles_labels()
    # order = range(blokaantal-1, -1, -1)
    # plt.legend([handles[idx] for idx in order],[labels[idx] for idx in order]) 
    
    # plt.gca().set_title("Totale partnerpensioen: €{:.2f}".format(PPtotaal).replace(".",","))
    # plt.suptitle(invoer.range(titel).value, fontweight='bold')
    
    
    # uitvoer.pictures.add(afbeelding, top = uitvoer.range((3,3)).top, left = uitvoer.range((3,3)).left, height = 300, name = "testnaam")
    
    
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
    

    #Berekeningskolommen leegmaken
    kolommen = invoer.range((1,8), (80,130))
    Uitkomst_kolommen = invoer.range((10,6), (20,6))
    #Kolommen legen waar de berekeningen komen
    kolommen.clear_contents()
    Uitkomst_kolommen.clear_contents()
    
    regeling_range = invoer.range((10,1), (20,1))
    pensioenbedragen = invoer.range((10,2), (20,2))
    sterftetafel_range = invoer.range((10,3), (20,3))
    rentes = invoer.range((10,4), (20,4))
    pensioenleeftijd_range = invoer.range((10,5), (20,5))
    koopsomfactor_range = invoer.range((10,6), (20,6))
    basis_koopsom = invoer.range((10,7), (20,7))

    pensioenleeftijd=[]
    rente=[]
    letters=[]
    for p in range(5):
        if len(letters) >= 26:
            for i in ascii_uppercase:
                letters.append(letters[p-1] + i)
        else:
            for i in ascii_uppercase:
                letters.append(i)  
      
    for c in range(2,12):
        pensioenleeftijd_range(c).value= regeling_range(c).value[-2:]

    counter=1
    for i in range(1,12):
        if pensioenbedragen(i).value != None:
            kolom_t = (counter-1)*10+9
            kolom_jaar = (counter-1)*10+10
            kolom_leeftijd = (counter-1)*10+11
            kolom_tpx = (counter-1)*10+12
            kolom_tqx = (counter-1)*10+13
            kolom_tqx_juli = (counter-1)*10+14
            kolom_dt = (counter-1)*10+15
            kolom_dt_juli = (counter-1)*10+16

            rente.append(rentes(i).value)
            pensioenleeftijd.append(pensioenleeftijd_range(i).value)
            

            # blok1 = list()
            # blok1.append(["t", "Jaar", "Leeftijd", "tpx", "tqx", "tqx op 1 juli", "dt", "dt op 1 juli"])
            # blok1.append(["0", '=year(B4)+' + str(int(pensioenleeftijd[i-1])), '=$E' + str(i+9), 1, 0, '=if({0}3<>"", (((13 - month($B$4)) * {1}2) + ((month($B$4) - 1) * {1}3)) / 12, "")'.format(letters[kolom_leeftijd - 1], letters[kolom_tqx - 1]), '=if({0}2<>"", (1+$D${1})^-{2}2, "")'.format(letters[kolom_leeftijd-1], str(i+9), letters[kolom_t-1]),\
            #               '=if({0}3<>"", (1+$D${1})^-({2}2 + (month($B$4)-1)/12), "")'.format(letters[kolom_leeftijd-1], str(i + 9), letters[kolom_t-1])])
            # for rij in range(59): blok1.append(['=1+{}{}'.format(letters[kolom_t-1], str(rij+2)),\
            #                                    '=1+{}{}'.format(letters[kolom_jaar-1], str(rij+2)),\
            #                                        '=if({0}{1}<119, 1+{0}{1},"")'.format(letters[kolom_leeftijd-1], str(rij+2)),\
            #                                            functions.tpxFormule(sterftetafel_range(i).value, rij, letters[kolom_leeftijd-1], letters[kolom_jaar-1], letters[kolom_tpx-1]),\
            #                                                '=if({0}{1}<>"", 1-{2}{1}, "")'.format(letters[kolom_leeftijd-1], str(rij+3), letters[kolom_tpx-1]),\
            #                                                    '=if({0}{1}<>"", (((13 - month($B$4)) * {2}{3}) + ((month($B$4) - 1) * {2}{1})) / 12, "")'.format(letters[kolom_leeftijd - 1], str(rij + 4), letters[kolom_tqx - 1], str(rij + 3)),\
            #                                                        '=if({0}{1}<>"", (1+$D${2})^-{3}{1}, "")'.format(letters[kolom_leeftijd-1], str(rij+3), str(i+9), letters[kolom_t-1]),\
            #                                                            '=if({0}{1}<>"", (1+$D${2})^-({3}{4} + (month($B$4)-1)/12), "")'.format(letters[kolom_leeftijd-1], str(rij+4), str(i + 9), letters[kolom_t-1], str(rij+3))\
            #                                                                ])
            # invoer.range((1, kolom_t), (61, kolom_t+7)).options(ndims = 2).formula = blok1
            

            invoer.range((1, kolom_t)).formula = [["t"]]
            invoer.range((2, kolom_t)).formula = [["0"]]
            invoer.range((3, kolom_t), (61, kolom_t)).formula = [['=1+' + letters[kolom_t-1] + '2']]
            
            invoer.range((1, kolom_jaar)).value = "Jaar"
            invoer.range((2, kolom_jaar)).formula = [['=year(B4)+' + str(int(pensioenleeftijd[i-1]))]]
            invoer.range((3, kolom_jaar), (61, kolom_jaar)).formula = [['=1+' + letters[kolom_jaar-1] + '2']]
            
            invoer.range((1, kolom_leeftijd)).value= "Leeftijd"
            invoer.range((2, kolom_leeftijd)).formula= [['=$E' + str(i+9)]]
            invoer.range((3, kolom_leeftijd), (61, kolom_leeftijd)).formula= [['=if(' + letters[kolom_leeftijd-1] + '2<119, 1+' + letters[kolom_leeftijd-1] + '2,"")']]
            
            invoer.range((1, kolom_tqx)).value= "tqx"
            invoer.range((2, kolom_tqx), (61, kolom_tqx)).formula= [['=if(' + letters[kolom_leeftijd-1] + '2<>"", 1-' + letters[kolom_tpx-1] + '2, "")']]

            invoer.range((1, kolom_tqx_juli)).value= "tqx op 1 juli"
            invoer.range((2, kolom_tqx_juli), (61, kolom_tqx_juli)).formula= [['=if(' + letters[kolom_leeftijd-1] + '3<>"", (((13-month($B$4))*' + letters[kolom_tqx-1] + '2)+((month($B$4)-1)*' + letters[kolom_tqx-1] + '3))/12, "")']]

            
            invoer.range((1, kolom_dt)).value= "dt"
            invoer.range((2, kolom_dt), (61, kolom_dt)).formula= [['=if(' + letters[kolom_leeftijd-1] + '2<>"", (1+$D$' + str(i+9) + ')^-' + letters[kolom_t-1] + '2, "")']]
            
            invoer.range((1, kolom_dt_juli)).value= "dt op 1 juli"
            invoer.range((2, kolom_dt_juli), (61, kolom_dt_juli)).formula= [['=if(' + letters[kolom_leeftijd-1] + '3<>"", (1+$D$' + str(i+9) + ')^-(' + letters[kolom_t-1] + '2+(month($B$4)-1)/12), "")']]

            if sterftetafel_range(i).value== "AG_2020":
                invoer.range((1, kolom_tpx)).value= "tpx"
                invoer.range((2, kolom_tpx)).value= 1
                print('=if(' + letters[kolom_leeftijd-1] + '3<>"", (1-INDEX(INDIRECT($C$' + str(i+9) + '),' + letters[kolom_leeftijd-1] + '2+1, ' + letters[kolom_jaar-1] + '2-2018))*' + letters[kolom_tpx-1] + '2,"")')
                invoer.range((3, kolom_tpx), (61, kolom_tpx)).formula= [['=if(' + letters[kolom_leeftijd-1] + '3<>"", (1-INDEX(INDIRECT($C$' + str(i+9) + '),' + letters[kolom_leeftijd-1] + '2+1, ' + letters[kolom_jaar-1] + '2-2018))*' + letters[kolom_tpx-1] + '2,"")']]

            else:
                invoer.range((1, kolom_tpx)).value= "tpx"
                invoer.range((2, kolom_tpx), (61, kolom_tpx)).formula= [['=if(' + letters[kolom_leeftijd-1] + '2<>"", INDEX(INDIRECT($C$' + str(i+9) + '),' + letters[kolom_leeftijd-1] + '2+1,1) / INDEX(INDIRECT($C$' + str(i+9) + '),$' + letters[kolom_leeftijd-1] + '$2+1,1),"")']]

            if regeling_range(i).value== "ZL":
                koopsomfactor_range(i).value= 0
                basis_koopsom(i).value= pensioenbedragen(i).value

    
            
            if regeling_range(i).value== "ZL":
                koopsomfactor_range(i).value= 0
                basis_koopsom(i).value= pensioenbedragen(i).value
                
            elif "OP" in regeling_range(i).value:
                koopsomfactor_range(i).formula= [['=ROUND(SUMPRODUCT(' + letters[kolom_tpx-1] + '2:' + letters[kolom_tpx-1] + '61,' + letters[kolom_dt-1] + '2:' + letters[kolom_dt-1] + '61),3)']]
                basis_koopsom(i).value= float(pensioenbedragen(i).value)*koopsomfactor_range(i).value
                
            else:
                koopsomfactor_range(i).formula = [['=ROUND(SUMPRODUCT(' + letters[kolom_tpx-1] + '2:' + letters[kolom_tpx-1] + '61,' + letters[kolom_tqx_juli-1] +'2:' + letters[kolom_tqx_juli-1] + '61,'+ letters[kolom_dt_juli-1] + '2:' + letters[kolom_dt_juli-1] + '61),3)']]
                basis_koopsom(i).value = float(pensioenbedragen(i).value)*koopsomfactor_range(i).value

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
    
    #ID van de gekozen afbeelding opzoeken
    ID = functions.flexopslagNaamNaarID(book, gekozenAfbeelding)
    
    #gekozen afbeelding verwijderen
    sheet.pictures[ID].delete()

@xw.sub
def afbeelding_aanpassen():
    """
    Functie die de gekozen afbeelding op de vergelijkings sheet als basis neemt voor nieuwe flexibilisaties
    """
    
    #sheets en book opslaan in variabelen
    book = xw.Book.caller()
    sheet = book.sheets["Vergelijken"]
    flexopslag = book.sheets["Flexopslag"]
    #gekozen afbeelding inlezen
    gekozenAfbeelding = sheet.cells(6,"B").value 
    
    #gegevens van gekozen afbeelding inladen
    opslag = functions.UitlezenFlexopslag(book, gekozenAfbeelding)
    #rijnummer deelnemer zoeken
    rijNr = int(float(flexopslag.cells(15,"B").value))
    
    #deelnemerobject inladen
    deelnemer = functions.getDeelnemersbestand(book, rijNr)
    deelnemer.activeerFlexibilisatie()      #maak pensioenobjecten aan
    
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
                if pensioengegevens[1] == "J":
                    flexibilisatie.leeftijd_Actief = True
                elif pensioengegevens[1] == "N":
                    flexibilisatie.leeftijd_Actief = False
                flexibilisatie.leeftijdJaar = int(float(pensioengegevens[2]))
                flexibilisatie.leeftijdMaand = int(float(pensioengegevens[3]))
                
                #uitruilen
                if pensioengegevens[4] == "J":
                    flexibilisatie.OP_PP_Actief = True
                elif pensioengegevens[4] == "N":
                    flexibilisatie.OP_PP_Actief = False
                    #volgorde
                if pensioengegevens[5] == "OP/PP": flexibilisatie.OP_PP_UitruilenVan = "OP naar PP"
                elif pensioengegevens[5] == "PP/OP": flexibilisatie.OP_PP_UitruilenVan = "PP naar OP"
                    #methode
                if pensioengegevens[6] == "Verh":
                    flexibilisatie.OP_PP_Methode = "Verhouding"
                    flexibilisatie.OP_PP_Verhouding_OP = int(float(pensioengegevens[7]))
                    flexibilisatie.OP_PP_Verhouding_PP = int(float(pensioengegevens[8]))
                elif pensioengegevens[6] == "Perc":
                    flexibilisatie.OP_PP_Methode = "Percentage"
                    flexibilisatie.OP_PP_Percentage = int(float(pensioengegevens[7]))
                
                flexibilisatie.OP_PP_UitruilenVan = pensioengegevens[6]
                
                
                flexibilisatie.leeftijdJaar = int(float(pensioengegevens[2]))
                
                #hoog-laag-constructie
                if pensioengegevens[9] == "J":
                    flexibilisatie.HL_Actief = True
                elif pensioengegevens[9] == "N":
                    flexibilisatie.HL_Actief = False
                    #volgorde
                if pensioengegevens[10] == "Hoog/Laag": flexibilisatie.HL_Volgorde = "Hoog-laag"
                elif pensioengegevens[10] == "Laag/Hoog": flexibilisatie.HL_Volgorde = "Laag-hoog"
                    #duur
                flexibilisatie.HL_Jaar = int(float(pensioengegevens[11]))
                    #methode
                if pensioengegevens[12] == "Verh":
                    flexibilisatie.HL_Methode = "Verhouding"
                    flexibilisatie.HL_Verhouding_Hoog = int(float(pensioengegevens[13]))
                    flexibilisatie.HL_Verhouding_Laag = int(float(pensioengegevens[14]))
                elif pensioengegevens[12] == "Verh":
                    flexibilisatie.HL_Methode = "Verschil"
                    flexibilisatie.HL_Verschil = int(float(pensioengegevens[13]))
                elif pensioengegevens[12] == "Opv":
                    flexibilisatie.HL_Methode = "Opvullen AOW"
                
    
    
    #scherm flexmenu openen
    logger = functions.setup_logger("Main") if not getLogger("Main").hasHandlers() else getLogger("Main")
    app = 0
    app = QtWidgets.QApplication(sys.argv)
    window = Klassen_Schermen.Flexmenu(xw.Book.caller(), deelnemer, logger)
    window.invoerVerandering()
    window.show()
    app.exec_()
    
@xw.sub
def NieuweFlexibilisatie():
    """
    Functie die het flexmenu scherm opnieuw opent voor de juiste deelnemer
    """
    
    #sheet en book opslaan in variabelen
    book = xw.Book.caller()
    flexopslag = book.sheets["Flexopslag"]
    
    #rijnummer deelnemer zoeken
    rijNr = int(float(flexopslag.cells(15,"B").value))
    #deelnemerobject inladen
    deelnemer = functions.getDeelnemersbestand(book, rijNr)
    deelnemer.activeerFlexibilisatie()      #maak pensioenobjecten aan
    
    #scherm flexmenu openen
    logger = functions.setup_logger("Main") if not getLogger("Main").hasHandlers() else getLogger("Main")
    app = 0
    app = QtWidgets.QApplication(sys.argv)
    window = Klassen_Schermen.Flexmenu(xw.Book.caller(), deelnemer, logger)
    window.invoerVerandering()
    window.show()
    app.exec_()
        
    
                  
@xw.sub
def flexibilisaties_testen():
   book = xw.Book.caller()
   invoer = book.sheets["Tijdelijk invoerscherm"] 
   
   flexibilisaties= invoer.range((24,2), (55,6))
   
   regeling_range= invoer.range((10,1), (20,1))
   aanspraken = invoer.range((10,2), (20,2))
   pensioenleeftijd= invoer.range((10,5), (20,5))
   koopsomfactor= invoer.range((10,6), (20,6))
   basis_koopsom= invoer.range((10,7), (20,7))
   
   koopsommen = basis_koopsom.value
   regelingen= regeling_range.value  

   #loop voor rijen
   for i in range(1,5):
       regeling = (i-1)*10+1
       soort = (i-1)*10+2
       verhouding = (i-1)*10+3
       duur = (i-1)*10+4
       factor_OP = (i-1)*10+5
       factor_PP = (i-1)*10+6
       aanspraak_OP = (i-1)*10+7
       aanspraak_PP = (i-1)*10+8
       
       factor_deel1 = (i-1)*10+5
       factor_deel2 = (i-1)*10+6
       aanspraak_deel1 = (i-1)*10+7
       aanspraak_deel2 = (i-1)*10+8

       flex_vak = invoer.range((factor_OP+23, 2), (aanspraak_PP+23, 5))
       flex_vak.clear_contents()

       #loop voor kolommen
       for c in range(1,5):
           if flexibilisaties(regeling, c).value != None:
               rij= regelingen.index(flexibilisaties(regeling, c).value)
               if "vervroegen" in flexibilisaties(soort, c).value or "verlaten" in flexibilisaties(soort, c).value:
                   #Bij corresponderende pensioenregelingsrij de hoeveelheid vervroegen/verlaten optellen
                   pensioenleeftijd(rij+1).value = pensioenleeftijd(rij+1).value + flexibilisaties(duur, c).value
                   pensioenleeftijd(rij+2).value = pensioenleeftijd(rij+2).value + flexibilisaties(duur, c).value
                   flexibilisaties(factor_OP, c).formula = koopsomfactor(rij+1).formula
                   flexibilisaties(factor_PP, c).formula = koopsomfactor(rij+2).formula
                   flexibilisaties(aanspraak_OP, c).value = round(float(koopsommen[rij]) / flexibilisaties(factor_OP, c).value)
                   flexibilisaties(aanspraak_PP, c).value = round(float(koopsommen[rij+1]) / flexibilisaties(factor_PP, c).value)
                   
                   OP_nieuw = float(flexibilisaties(aanspraak_OP, c).value)
                   PP_nieuw = float(flexibilisaties(aanspraak_PP, c).value)

               elif "AOW" in flexibilisaties(soort, c).value:
                   flexibilisaties(factor_OP, c).value = "Berekeningen komen later"
                       
               
               elif  "uitruilen" in flexibilisaties(soort, c).value:
                   uitruilen_naar = flexibilisaties(soort, c).value[-2:]
                   flexibilisaties(factor_OP, c).formula = koopsomfactor(rij+1).formula
                   flexibilisaties(factor_PP, c).formula = koopsomfactor(rij+2).formula
                   
                   if ":" in str(flexibilisaties(verhouding, c).value):
                       verhouding_uitruilen = int(flexibilisaties(verhouding, c).value[-2:])/100
                       if uitruilen_naar == "PP":
                           flexibilisaties(aanspraak_OP, c).value = float(koopsommen[rij]) / (flexibilisaties(factor_OP, c).value + verhouding_uitruilen * flexibilisaties(factor_PP, c).value)
                           flexibilisaties(aanspraak_PP, c).value = round(float(flexibilisaties(aanspraak_OP, c).value) * verhouding_uitruilen + PP_nieuw)
                           flexibilisaties(aanspraak_OP, c).value = round(flexibilisaties(aanspraak_OP, c).value)
                           
                           OP_nieuw = float(flexibilisaties(aanspraak_OP, c).value)
                           PP_nieuw = float(flexibilisaties(aanspraak_PP, c).value)
                       else:
                           flexibilisaties(aanspraak_PP, c).value = float(koopsommen[rij+1]) / (flexibilisaties(factor_PP, c).value + verhouding_uitruilen * flexibilisaties(factor_OP, c).value)
                           flexibilisaties(aanspraak_OP, c).value = round(float(flexibilisaties(aanspraak_PP, c).value) * verhouding_uitruilen + OP_nieuw)
                           flexibilisaties(aanspraak_PP, c).value = round(flexibilisaties(aanspraak_PP, c).value)
                           
                           OP_nieuw = float(flexibilisaties(aanspraak_OP, c).value)
                           PP_nieuw = float(flexibilisaties(aanspraak_PP, c).value)
                           
                   else:
                       if uitruilen_naar == "PP":
                           verschil_uitruilen = flexibilisaties(verhouding, c).value * OP_nieuw
                           flexibilisaties(aanspraak_OP, c).value = round(OP_nieuw - verschil_uitruilen)
                           flexibilisaties(aanspraak_PP, c).value = round(verschil_uitruilen * flexibilisaties(factor_OP, c).value / flexibilisaties(factor_PP, c).value + PP_nieuw)
                           
                           OP_nieuw = float(flexibilisaties(aanspraak_OP, c).value)
                           PP_nieuw = float(flexibilisaties(aanspraak_PP, c).value)
                       else:
                           verschil_uitruilen = flexibilisaties(verhouding, c).value * PP_nieuw
                           flexibilisaties(aanspraak_PP, c).value = round(PP_nieuw - verschil_uitruilen)
                           flexibilisaties(aanspraak_OP, c).value = round(verschil_uitruilen * flexibilisaties(factor_PP, c).value / flexibilisaties(factor_OP, c).value + OP_nieuw)
                           
                           OP_nieuw = float(flexibilisaties(aanspraak_OP, c).value)
                           PP_nieuw = float(flexibilisaties(aanspraak_PP, c).value)
    
               else:
                   flex_duur = int(flexibilisaties(duur, c).value)
                   x = koopsomfactor(rij+1).formula
                   
                   if x.index('61') != None:
                       y = x.replace('61', str(flex_duur+1))

                   if x.index('2') != None:
                       z = x.replace('2', str(flex_duur+2))
                       
                   #Hoog-Laag constructie    
                   if "hoog" in flexibilisaties(soort, c).value:
                       flexibilisaties(factor_deel1, c).formula = y 
                       if ":" in str(flexibilisaties(verhouding, c).value):
                           soort_HL = int(flexibilisaties(verhouding, c).value[-2:])/100
                           flexibilisaties(factor_deel2, c).formula = y + '+' + z[1:] + '*' + str(soort_HL)
                           flexibilisaties(aanspraak_deel1, c).value = (OP_nieuw * koopsomfactor(rij+1).value) / flexibilisaties(factor_deel2, c).value
                           flexibilisaties(aanspraak_deel2, c).value = round(float(flexibilisaties(aanspraak_deel1, c).value) * soort_HL)
                           flexibilisaties(aanspraak_deel1, c).value = round(flexibilisaties(aanspraak_deel1, c).value)

                       else:
                           soort_HL = flexibilisaties(verhouding, c).value
                           flexibilisaties(factor_deel2, c).formula = z
                           flexibilisaties(aanspraak_deel1, c).value = ((OP_nieuw * koopsomfactor(rij+1).value) + soort_HL*flexibilisaties(factor_deel2, c).value)/koopsomfactor(rij+1).value
                           flexibilisaties(aanspraak_deel2, c).value = round(flexibilisaties(aanspraak_deel1, c).value - soort_HL)
                           flexibilisaties(aanspraak_deel1, c).value = round(flexibilisaties(aanspraak_deel1, c).value)

                   #Laag-Hoog constructie        
                   else:
                       if ":" in str(flexibilisaties(verhouding, c).value):
                           soort_HL = int(flexibilisaties(verhouding, c).value[:2])/100
                           y = y + '*' + str(soort_HL)
                           flexibilisaties(factor_deel1, c).formula = y 
                           flexibilisaties(factor_deel2, c).formula = y + '+' + z[1:] 
                           flexibilisaties(aanspraak_deel2, c).value = (OP_nieuw * koopsomfactor(rij+1).value) / flexibilisaties(factor_deel2, c).value
                           flexibilisaties(aanspraak_deel1, c).value = round(float(flexibilisaties(aanspraak_deel2, c).value) * soort_HL)
                           flexibilisaties(aanspraak_deel2, c).value = round(flexibilisaties(aanspraak_deel2, c).value)
                           
                       else:
                           flexibilisaties(factor_deel1, c).formula = y 
                           soort_HL = flexibilisaties(verhouding, c).value
                           flexibilisaties(factor_deel2, c).formula = z
                           flexibilisaties(aanspraak_deel1, c).value = (OP_nieuw * koopsomfactor(rij+1).value - soort_HL * flexibilisaties(factor_deel2, c).value)/koopsomfactor(rij+1).value
                           flexibilisaties(aanspraak_deel2, c).value = round(flexibilisaties(aanspraak_deel1, c).value + soort_HL)
                           flexibilisaties(aanspraak_deel1, c).value = round(flexibilisaties(aanspraak_deel1, c).value)
                    
                           
       