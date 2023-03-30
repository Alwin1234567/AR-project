"""
Header
Hier komen alle libraries die in het programma gebruikt worden
"""
import xlwings as xw
import sys
from PyQt5 import QtWidgets
import functions
import Klassen_Schermen
from logging import getLogger
from reportlab.pdfgen.canvas import Canvas
from reportlab.lib.units import cm
import mplwidget #zodat deze ook in de Frozen variant geïmporteerd word
from reportlab.graphics import renderPDF
import os

"""
Body
Hier komen alle functies
"""


@xw.sub
def VergelijkenHelp():
    book = xw.Book.caller()
    
    #mogelijke help-berichten definiëren
    Kiezen = "'Kiezen' maakt een pdf bestand van de gekozen flexibilisatie aan. Hierin staan de gegevens van het originele en het geselecteerde pensioen en kunnen deze vergeleken worden.\n"
    Verwijderen = "'Verwijderen' verwijderd de gekozen flexibilisatie van de vergelijken sheet en uit de opslag.\n"
    Aanpassen = "'Aanpassen' opent het flexibilisatiemenu met daarin de flexibilisaties van de gekozen flexibilisatie ingeladen. Dit kunt u gebruiken om verder te flexibiliseren. De gekozen flexibilisatie blijft opgeslagen staan op de vergelijken sheet.\n"
    NieuweFlex = "'Nieuwe flexibilisatie' opent het flexibilisatiemenu, waardoor er een nieuwe flexibilisatie voor de huidige deelnemer uitgevoerd kan worden.\n"
    AndereDeelnemer = "'Andere deelnemer' opent het deelnemerselectiescherm om een flexibilisatie voor een nieuwe deelnemer te starten.\n"
    Inloggen = "'Inloggen' opent het inlogscherm voor beheerders. Een beheerder kan gegevens inzien en wijzigen.\n"
    Beheerderkeuzes = "'Beheerderkeuzes' opent het beheerderkeuzes scherm waarmee gegevens ingezien en gewijzigd kunnen worden.\n"
    Uitloggen = "'Uitloggen' logt u als beheerder uit. U kunt nog steeds flexibilisaties uitvoeren en de huidige flexibilisaties blijven bewaard, maar u kunt niet alle gegevens meer inzien of wijzigen.\n"
    Vergelijken = "U kunt flexibilisaties met elkaar vergelijken door deze in de drop down menu's in de vakken linksonder het knoppenmenu te selecteren. Hierdoor worden de afbeeldingen verplaatst en kunt u ze naast elkaar zetten. Door '-' of een andere afbeelding te selecteren verplaatst de vorige afbeelding terug naar zijn originele plek.\n"
    
    bericht = f"Dit is een uitleg van wat u kunt op de vergelijken sheet: \n\n{Kiezen}\n{Verwijderen}\n{Aanpassen}\n{NieuweFlex}\n{AndereDeelnemer}\n"
    if functions.isBeheerder(book):
        bericht = bericht + f"{Uitloggen}\n{Beheerderkeuzes}\n{Vergelijken}\n"
    else:
        bericht = bericht + f"{Inloggen}\n{Vergelijken}\n"
    
    #messagebox met help-bericht maken
    functions.Mbox("Help bij vergelijken", bericht, 0)

@xw.sub
def AfbeeldingVerplaatsen(vak):
    book = xw.Book.caller()
    vergelijken = book.sheets["Vergelijken"]
    
    if vak == 1:
        vlak = vergelijken.range((14,10))
        keuzecel = "J13"
    elif vak == 2:
        vlak = vergelijken.range((38,2))
        keuzecel = "B37"
    else: #vak == 3 of iets anders
        vlak = vergelijken.range((38,10))
        keuzecel = "J37"
    
    try:
        gekozenAfbeelding = vergelijken[keuzecel].value
        ID = functions.flexopslagNaamNaarID(book, gekozenAfbeelding)
    except:
        ID = "-"
        vergelijken["M2"].value == "ID is leeg of -" 
    #vergelijken sheet unprotecten
    vergelijken.api.Unprotect(Password = functions.wachtwoord())
    for pic in vergelijken.pictures:
        if round(pic.top,1) == round(vlak.top,1) and round(pic.left,1) == round(vlak.left,1): #als er een afbeelding al in staat
            if pic.name != ID: #als naam afbeelding in box niet gelijk is aan naam gekozen afbeelding
                #afbeelding terugverplaatsen
                
                teller = int(pic.name.split()[-1])-1
                rij = int(12 + (teller%4)*22)
                kolom = int(17 + ((teller - teller%4)/4)*8)
                afbeelding = vergelijken.pictures[pic.name]
                afbeelding.top = vergelijken.range((rij,kolom)).top
                afbeelding.left = vergelijken.range((rij,kolom)).left
                
        if ID != "-":
            #juiste afbeelding op vlak zetten
            afbeelding = vergelijken.pictures[ID]
            afbeelding.top = vlak.top
            afbeelding.left = vlak.left
    #vergelijken sheet protecten
    functions.ProtectBeheer(vergelijken)

      
@xw.sub
def AfbeeldingKiezen():
    """
    Functie die de gekozen afbeelding op de vergelijkings sheet kiest
    """
    
    #sheets en book opslaan in variabelen
    book = xw.Book.caller()
    sheet = book.sheets["Vergelijken"]
    flexopslag = book.sheets["Flexopslag"]
    if str(flexopslag.cells(2, 5).value) != "None":   #alleen als er nog flexibilisaties opgeslagen zijn
        #gekozen afbeelding inlezen
        gekozenAfbeelding = sheet.cells(6,"B").value
        
        #Deelnemer toegevoegd zodat er gekeken kan worden naar het oude pensioen. 
        #rijnummer deelnemer zoeken
        rijNr = int(float(flexopslag.cells(15,"B").value))
        
        #deelnemerobject inladen
        deelnemer = functions.getDeelnemersbestand(book, rijNr)
        deelnemer.activeerFlexibilisatie()      #maak pensioenobjecten aan
        
        oudpensioen = []
        for i in deelnemer.pensioenen:
            pensioen = []
            pensioen.append(i.pensioenNaam)             #0
            pensioen.append(str(i.pensioenleeftijd))    #1
            pensioen.append("0") #De maand              #2
            pensioen.append(i.ouderdomsPensioen)        #3
            pensioen.append(i.partnerPensioen)          #4
            oudpensioen.append(pensioen)            
            
            
        Fonds = functions.getPensioeninformatie(book)
        puntleeftijd = oudpensioen[0][1]
        AOWjaar = float(puntleeftijd)//1 
        AOWmaand = (float(puntleeftijd)%1) * 12
        AOWleeftijd = functions.leeftijd_notatie(AOWjaar, AOWmaand)
        
        if deelnemer.burgelijkeStaat == "Samenwonend":    
            AOW = "€" + str(int(float(Fonds[0].samenwondendAOW)))
        else:
            AOW = "€" + str(int(float(Fonds[0].alleenstaandAOW)))
            

        nieuwpensioen = functions.UitlezenFlexopslag(book, gekozenAfbeelding)
        
        eenperjaar = functions.geld_per_leeftijd(oudpensioen, nieuwpensioen)
        
        eenperjaaroud = eenperjaar[0]
        eenperjaarnieuw = eenperjaar[1]
        
        
        naam_pdf = gekozenAfbeelding + ".pdf"
        
        save_pad = os.path.join(functions.krijgpad(), naam_pdf)
        pdf_canvas = Canvas(save_pad)
        pdf_canvas.setFont("Helvetica", 11)
        halfbreedte = cm*10.5
        
       
        
        oudpensioenimg = functions.maak_afbeelding(deelnemer, pdf = True, titel = "Oudpensioen")
        
        functions.GegevensNaarFlexibilisatie(deelnemer, nieuwpensioen)
        nieuwpensioenimg = functions.maak_afbeelding(deelnemer, pdf = True, titel = "Nieuw pensioen")
        
        
        renderPDF.draw(oudpensioenimg, pdf_canvas, 105, 540)
        renderPDF.draw(nieuwpensioenimg, pdf_canvas, 105, 200)
        
        pdf_canvas.showPage()
        functions.nieuwe_pagina(pdf_canvas, halfbreedte)
        
        verhaalstart = 706
        verhaallijn = verhaalstart
        
        totOPoud = 0
        pdf_canvas.drawString(30+halfbreedte, 720, "Met uw oude pensioen")
        for i in eenperjaaroud:
            totOPoud = totOPoud + i[1]
            oudverhaal = "ontving u vanaf  "+ i[0]+ " €" + str(totOPoud)+ " per jaar aan OP"
            pdf_canvas.drawString(30+halfbreedte, verhaallijn, oudverhaal)
            verhaallijn -= 14
            
        totPPoud = 0
        for i in oudpensioen:
            totPPoud = totPPoud + int(i[4])
        
        PPoudverhaal = "Uw oude partner pensioen was €" + str(totPPoud) + " per jaar"
        pdf_canvas.drawString(30 + halfbreedte, verhaallijn-14, PPoudverhaal)
        
        verhaallijn = verhaalstart
        
        pdf_canvas.drawString(30, 720, "Met uw nieuwe pensioen")
        
        totOPnieuw = 0
        for i in eenperjaarnieuw:
            totOPnieuw = totOPnieuw + i[1]
            if "maand" in i[0]:
                nieuwverhaal = "ontvangt u vanaf  "+ i[0] 
                pdf_canvas.drawString(30, verhaallijn, nieuwverhaal)
                verhaallijn -= 14
                nieuwverhaal = "€" + str(totOPnieuw)+ " per jaar aan OP"
                pdf_canvas.drawString(110, verhaallijn, nieuwverhaal)
            else:
                nieuwverhaal = "ontvangt u vanaf  "+ i[0]+ " €" + str(totOPnieuw)+ " per jaar aan OP"
                pdf_canvas.drawString(30, verhaallijn, nieuwverhaal)
            verhaallijn -= 14
        
        
        totPPnieuw = 0
        for i in nieuwpensioen:
            totPPnieuw = totPPnieuw + int(float(i[17]))
        
        PPnieuwverhaal = "Uw nieuwe partner pensioen is €" + str(totPPnieuw) + " per jaar"
        pdf_canvas.drawString(40, verhaallijn-14, PPnieuwverhaal)
        
        
        AOWverhaal1 = "Vanaf " + AOWleeftijd + " krijgt u"
        AOWverhaal2 = AOW + " aan AOW per jaar"
        pdf_canvas.drawString(40, verhaallijn-42, AOWverhaal1)
        pdf_canvas.drawString(40, verhaallijn-56, AOWverhaal2)
        functions.nieuwe_pagina(pdf_canvas, halfbreedte)
        
        startschrijfhoogte = 720
        schrijfhoogte = startschrijfhoogte
        
        
        #Samenvattingen
        benodigde_indexen = [0, 1, 2, 15, 17, 4, 5, 7, 9, 10, 11, 13]
        #4 start 5, 7(8)
        #9 start 10,11,13(14) past 15(16) aan
        
        
        p = 1 # hoeveelste pensioen neergezet wordt
        for pensioen in nieuwpensioen:
            labels = ["Pensioenfonds", "Vervroegen/Uitstellen?", "Pensioenleeftijd", "Nieuw OP", "Nieuw PP", "Uitruilen?", "Volgorde", pensioen[6],"Hoog/Laag?", "volgorde",  "Duur", pensioen[12]]
            benodigdeantwoorden = []
            for l in benodigde_indexen:
                if l == 2:
                    if pensioen[1] == "Ja":
                        antwoord = functions.leeftijd_notatie(pensioen[2], pensioen[3])
                    else:
                        antwoord =  functions.leeftijd_notatie(oudpensioen[p-1][1], "0")
                        
                elif l == 15:
                    if pensioen[9] == "Ja":
                        hoog = "€" + str(int(float(pensioen[l]))) + " hoog"
                        laag = "€" + str(int(float(pensioen[l+1]))) + " laag"
                        if pensioen[10] == "Hoog-laag":
                            antwoord = hoog + " " + laag
                        else:
                            antwoord = laag + " " + hoog
                    else:
                        antwoord = "€" + str(int(float(pensioen[l])))
                elif l == 17:
                    antwoord = "€" + str(int(float(pensioen[l])))
                        
                elif l == 5 or l == 7: #opties die alleen bij uitruilen horen
                    if pensioen[4] == "Ja":
                        if l == 7:
                            if pensioen[l-1] == "Verhouding":
                                antwoord = pensioen[l] + ":" + pensioen[l+1]
                            else: #pencentage
                                if pensioen[l+1] == "None": #geen maxpercentage gebruikt
                                    antwoord = str(int(float(pensioen[l]))) + "%"
                                else: #wel maxpercentage gebruikt
                                    antwoord = str(round(float(pensioen[l+1]), 2)) + "%"
                                
                        else:
                            antwoord = pensioen[l]
                    else:
                        antwoord = ""
                elif l == 10 or l == 11 or l == 13: #opties die alleen bij hoog-laag horen
                    if pensioen[9] == "Ja":
                        if l == 11:
                            antwoord = functions.leeftijd_notatie(pensioen[l], "0")
                        elif l == 13:
                            if pensioen[l-1] == "Verhouding":
                                antwoord = pensioen[l] + ":" + pensioen[l+1]
                            elif pensioen[l-1] == "Verschil":
                                if pensioen[l+1] == "None": #geen maxbedrag gebruikt
                                    antwoord = "€" + str(int(float(pensioen[l])))
                                else:  #wel maxbedrag gebruikt
                                    antwoord = "€" + str(round(float(pensioen[l+1]), 2))
                                    
                            else: #Opvullen AOW
                                antwoord = ""
                        else:
                            antwoord = pensioen[l]
                    else:
                        antwoord = ""
                
                else:
                    antwoord = pensioen[l]
                benodigdeantwoorden.append(antwoord)
                
            #antwoorden noteren
            if p%3 == 1:
                pdf_canvas.showPage()
                functions.nieuwe_pagina(pdf_canvas, halfbreedte)
                schrijfhoogte = startschrijfhoogte
            k = 0
            for j in labels:
                if k == 0: # fonds
                    pdf_canvas.drawString(30 + halfbreedte, schrijfhoogte, j)
                    pdf_canvas.drawString(170 + halfbreedte, schrijfhoogte, oudpensioen[p-1][0])
                elif k == 2: #leeftijd
                    pdf_canvas.drawString(30 + halfbreedte, schrijfhoogte, j)
                    oud_leeftijd = functions.leeftijd_notatie(oudpensioen[p-1][1], "0")
                    pdf_canvas.drawString(170 + halfbreedte, schrijfhoogte, oud_leeftijd)
                elif k == 3: #OP
                    pdf_canvas.drawString(30 + halfbreedte, schrijfhoogte, "Oud OP")
                    pdf_canvas.drawString(170 + halfbreedte, schrijfhoogte, "€" + str(oudpensioen[p-1][3]))
                elif k == 4: #PP
                    pdf_canvas.drawString(30 + halfbreedte, schrijfhoogte, "Oud PP")
                    pdf_canvas.drawString(170 + halfbreedte, schrijfhoogte, "€" + str(oudpensioen[p-1][4]))
                pdf_canvas.drawString(30, schrijfhoogte, j)
                #antwoord = self._pensioen[self._benodigde_indexen[k]]
                antwoord = benodigdeantwoorden[k]
                pdf_canvas.drawString(170, schrijfhoogte, antwoord)
                schrijfhoogte = schrijfhoogte -14
                k+=1
            schrijfhoogte = schrijfhoogte -25
            p+=1
            
        pdf_canvas.save()
        functions.Mbox("Opgeslagen", "De keuze is opgeslagen", 0)
        
    else:
        functions.Mbox("foutmelding", "Er zijn geen flexibilisaties opgeslagen. \nMaak eerst een nieuwe flexibilisatie aan.", 0)


@xw.sub
def AfbeeldingVerwijderen():
    """
    Functie die de gekozen afbeelding op de vergelijkings sheet verwijderd
    """
    
    #sheets en book opslaan in variabelen
    book = xw.Book.caller()
    Vergelijken = book.sheets["Vergelijken"]
    Opslag = book.sheets["Flexopslag"]
    if str(Opslag.cells(2, 5).value) != "None":   #alleen als er nog flexibilisaties opgeslagen zijn
        #gekozen afbeelding inlezen
        gekozenAfbeelding = Vergelijken.cells(6,"B").value
        #vragen of echt verwijderd moet worden
        controle = functions.Mbox("Afbeelding verwijderen", f"Wilt u de flexibilisatie '{gekozenAfbeelding}' echt verwijderen?\nU kunt deze actie niet ongedaan maken.", 4)
        if controle == "Ja":
            #ID van de gekozen afbeelding opzoeken
            ID = functions.flexopslagNaamNaarID(book, gekozenAfbeelding)
            
            #vergelijken sheet unprotecten
            Vergelijken.api.Unprotect(Password = functions.wachtwoord())
            #gekozen afbeelding verwijderen
            try:
                Vergelijken.pictures[ID].delete()
            except:
                pass
            #vergelijken sheet protecten
            functions.ProtectBeheer(Vergelijken) #.api.Protect(Password = functions.wachtwoord(), Contents=False)
            
            #tellen hoeveel opgeslagen flexibiliseringen en hoeveel pensioenen
            Flexopslag = functions.FlexopslagVinden(xw.Book.caller(), gekozenAfbeelding)
            
            startKolom = Flexopslag[0]
            laatsteKolom = Flexopslag[1]
            aantalPensioenen = Flexopslag[2]
            rijen = aantalPensioenen*20 + 4
            #opslag sheet unprotecten
            Opslag.api.Unprotect(Password = functions.wachtwoord())
            if startKolom != laatsteKolom: #er zijn meer dan 1 flexibilisaties opgeslagen
                #verwijderen gegevens verwijderde flexibilisatie
                Opslag.range((1,startKolom-1),(rijen,startKolom+1)).clear_contents()
                #flexibilisaties na verwijderde blok opschuiven
                Opslag.cells(1,startKolom-1).value = Opslag.range((1,startKolom+3),(rijen,laatsteKolom+1)).value
            #laatste (of enige) kolom verwijderen
            Opslag.range((1,laatsteKolom-1),(rijen,laatsteKolom+1)).clear()
            #opslag sheet protecten
            functions.ProtectBeheer(Opslag) #.api.Protect(Password = functions.wachtwoord())
            
            try:
                #drop down op vergelijkingssheet updaten
                functions.vergelijken_keuzes()
            except:
                #laatste opslag is verwijderd, dus drop down legen
                Vergelijken["B6"].value = ""
    
    else: #er zijn geen flexibilisaties opgeslagen
        #keuzecel in vergelijkingssheet legen
        Vergelijken["B6"].value = ""
        functions.Mbox("foutmelding", "Er zijn geen flexibilisaties opgeslagen. \nMaak eerst een nieuwe flexibilisatie aan.", 0)
    
    
@xw.sub
def afbeelding_aanpassen():
    """
    Functie die de gekozen afbeelding op de vergelijkings sheet als basis neemt voor nieuwe flexibilisaties
    """
    
    #sheets en book opslaan in variabelen
    book = xw.Book.caller()
    sheet = book.sheets["Vergelijken"]
    flexopslag = book.sheets["Flexopslag"]
    if str(flexopslag.cells(2, 5).value) != "None":   #alleen als er nog flexibilisaties opgeslagen zijn
        #gekozen afbeelding inlezen
        gekozenAfbeelding = sheet.cells(6,"B").value 
        
        #gegevens van gekozen afbeelding inladen
        opslag = functions.UitlezenFlexopslag(book, gekozenAfbeelding)
        #rijnummer deelnemer zoeken
        rijNr = int(float(flexopslag.cells(15,"B").value))
        
        #deelnemerobject inladen
        deelnemer = functions.getDeelnemersbestand(book, rijNr)
        deelnemer.activeerFlexibilisatie()      #maak pensioenobjecten aan
        
        
        functions.GegevensNaarFlexibilisatie(deelnemer, opslag)
        
        #lijst met pensioennamen langsgaan en opgeslagen flexibilisatiegegevens per pensioen toevoegne aan flexibiliseringsobject van het deelnemersobject
        
        
        # for i,p in enumerate(pensioennamen):
            # for flexibilisatie in deelnemer.flexibilisaties:
            #     #als het flexibilisatieobject bij het pensioen uit de lijst pensioennamen hoort
            #     if flexibilisatie.pensioen.pensioenNaam == p:
            #         #met properties flexibilisaties opslaan in objecten flexibilisatie
            #         pensioengegevens = opslag[i]
            #         #leeftijd aanpassen
            #         if pensioengegevens[1] == "Ja":
            #             flexibilisatie.leeftijd_Actief = True
            #         elif pensioengegevens[1] == "Nee":
            #             flexibilisatie.leeftijd_Actief = False
            #         flexibilisatie.leeftijdJaar = int(float(pensioengegevens[2]))
            #         flexibilisatie.leeftijdMaand = int(float(pensioengegevens[3]))
                    
            #         #uitruilen
            #         if pensioengegevens[4] == "Ja":
            #             flexibilisatie.OP_PP_Actief = True
            #         elif pensioengegevens[4] == "Nee":
            #             flexibilisatie.OP_PP_Actief = False
            #             #volgorde
            #         flexibilisatie.OP_PP_UitruilenVan = pensioengegevens[5]
            #             #methode
            #         flexibilisatie.OP_PP_Methode = pensioengegevens[6]
            #         if pensioengegevens[6] == "Verhouding":
            #             flexibilisatie.OP_PP_Verhouding_OP = int(float(pensioengegevens[7]))
            #             flexibilisatie.OP_PP_Verhouding_PP = int(float(pensioengegevens[8]))
            #         elif pensioengegevens[6] == "Percentage":
            #             flexibilisatie.OP_PP_Percentage = int(float(pensioengegevens[7]))
                    
                    
            #         #hoog-laag-constructie
            #         if pensioengegevens[9] == "Ja":
            #             flexibilisatie.HL_Actief = True
            #         elif pensioengegevens[9] == "Nee":
            #             flexibilisatie.HL_Actief = False
            #             #volgorde
            #         flexibilisatie.HL_Volgorde = pensioengegevens[10]
            #             #duur
            #         flexibilisatie.HL_Jaar = int(float(pensioengegevens[11]))
            #             #methode
            #         flexibilisatie.HL_Methode = pensioengegevens[12]
            #         if pensioengegevens[12] == "Verhouding":
            #             flexibilisatie.HL_Verhouding_Hoog = int(float(pensioengegevens[13]))
            #             flexibilisatie.HL_Verhouding_Laag = int(float(pensioengegevens[14]))
            #         elif pensioengegevens[12] == "Verschil":
            #             flexibilisatie.HL_Verschil = int(float(pensioengegevens[13]))       
        if len(gekozenAfbeelding)>4:
            titelAfbeelding = gekozenAfbeelding[4:]
        else:
            titelAfbeelding = ""
                
        #scherm flexmenu openen
        logger = functions.setup_logger("Main") if not getLogger("Main").hasHandlers() else getLogger("Main")
        app = 0
        app = QtWidgets.QApplication(sys.argv)
        window = Klassen_Schermen.Flexmenu(xw.Book.caller(), deelnemer, logger, titel = titelAfbeelding)
        window.invoerVerandering(num = 0, aanpassing = True)
        
        
        window.show()
        app.exec_()
    else:
        functions.Mbox("foutmelding", "Er zijn geen flexibilisaties opgeslagen. \nMaak eerst een nieuwe flexibilisatie aan.", 0)
    
@xw.sub
def NieuweFlexibilisatie():
    """
    Functie die het flexmenu scherm opnieuw opent voor de juiste deelnemer
    """
    
    #sheet en book opslaan in variabelen
    book = xw.Book.caller()
    flexopslag = book.sheets["Flexopslag"]
    if str(flexopslag.cells(15,"B").value) != "None":   #alleen als er nog flexibilisaties opgeslagen zijn
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
        
        window.show()
        app.exec_()
    else:
        functions.Mbox("foutmelding", "Er is geen deelnemer opgeslagen. \nGelieve eerst een deelnemer te selecteren via de knop 'Andere deelnemer'.", 0)
        
@xw.sub
def AndereDeelnemer():
    """
    Functie die het deelnemerselectie scherm opent
    """
    controle = functions.Mbox("Andere deelnemer selecteren", "Door een andere deelnemer te selecteren zullen de huidige gegevens op de vergelijken sheet verwijderd worden.\nU kunt deze actie niet ongedaan maken.", 1)
    if controle == "OK Clicked":
        #scherm Deelnemerselectie openen
        logger = functions.setup_logger("Main") if not getLogger("Main").hasHandlers() else getLogger("Main")
        app = 0
        app = QtWidgets.QApplication(sys.argv)
        window = Klassen_Schermen.Deelnemerselectie(xw.Book.caller(), logger)
        window.show()
        app.exec_()
    
@xw.sub
def BeheerderskeuzesOpenen():
    """
    Functie die het scherm met de beheerderskeuzes opent
    """
    
    #scherm Beheerderkeuzes openen
    logger = functions.setup_logger("Main") if not getLogger("Main").hasHandlers() else getLogger("Main")
    app = 0
    app = QtWidgets.QApplication(sys.argv)
    windowBeheerder = Klassen_Schermen.Beheerderkeuzes(xw.Book.caller(), logger)
    windowBeheerder.show()
    app.exec_()
    
@xw.sub
def InEnUitloggen() :
    """
    functie om in of uit te loggen als beheerder

    Returns
    -------
    None.

    """
    book = xw.Book.caller()
    logger = functions.setup_logger("Main") if not getLogger("Main").hasHandlers() else getLogger("Main")
    if functions.isBeheerder(book):
        #beheerder is ingelogd, dus wil uitloggen
        controle = functions.Mbox("Uitloggen", "Weet u zeker dat u wilt uitloggen?",4)
        if controle == "Ja":
            app = 0
            app = QtWidgets.QApplication(sys.argv)
            windowBeheerder = Klassen_Schermen.Beheerderkeuzes(xw.Book.caller(), logger)
            windowBeheerder.btnUitloggenClicked()
            app.exec_()
        
    else:
        #beheerder is niet ingelogd, dus wil inloggen
        app = 0
        app = QtWidgets.QApplication(sys.argv)
        windowInloggen = Klassen_Schermen.Inloggen(xw.Book.caller(), logger)
        windowInloggen.show()
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

if __name__ == '__main__':
    if sys.argv[1] == "AfbeeldingVerplaatsen": AfbeeldingVerplaatsen(sys.argv[2])
    elif sys.argv[1] == "AfbeeldingKiezen": AfbeeldingKiezen()
    elif sys.argv[1] == "AfbeeldingVerwijderen": AfbeeldingVerwijderen()
    elif sys.argv[1] == "afbeelding_aanpassen": afbeelding_aanpassen()
    elif sys.argv[1] == "NieuweFlexibilisatie": NieuweFlexibilisatie()
    elif sys.argv[1] == "AndereDeelnemer": AndereDeelnemer()
    elif sys.argv[1] == "BeheerderskeuzesOpenen": BeheerderskeuzesOpenen()
    elif sys.argv[1] == "InEnUitloggen": InEnUitloggen()
    elif sys.argv[1] == "flexibilisaties_testen": flexibilisaties_testen()