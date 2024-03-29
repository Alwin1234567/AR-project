"""
Header
Hier komen alle libraries die in het programma gebruikt worden
"""
from Pensioen import Pensioen
from flex_keuzes import Flexibilisering
from datetime import datetime
"""
Body
Hier komen alle functies
"""
class Deelnemer():
    """
    Een class waarin de informatie over een deelnemer staat opgeslagen.
    
    informatie : list
        Een lijst met daarin de naam van een attribuut en de waarde van de attribuut.
    pensioeninformatie : list
        Een lijst met daarin pensioenfonds objecten van de verschillende fondsen.
    """
    
    def __init__(self, informatie, pensioeninformatie):
        self._achternaam = self.informatieOpslaan(informatie, "Naam")
        self._tussenvoegsels = self.informatieOpslaan(informatie, "tussenvoegsels")
        self._voorletters = self.informatieOpslaan(informatie, "voorletter")
        self._geboortedatum = self.informatieOpslaan(informatie, "Geboortedatum")
        self._geslacht = self.informatieOpslaan(informatie, "Geslacht")
        self._burgelijkeStaat = self.informatieOpslaan(informatie, "Burg. Staat")
        self._ftLoon = self.informatieOpslaan(informatie, "FT loon")
        if self._ftLoon != None: self._ftLoon = int(self._ftLoon)
        self._pt = self.informatieOpslaan(informatie, "PT%")
        if self._pt != None: self._pt = float(self._pt)
        self._regeling = self.informatieOpslaan(informatie, "Regeling")
        self._rijNr = self.informatieOpslaan(informatie, "rijNr")
        self._pensioenen = self.pensioenenOpslaan(informatie, pensioeninformatie)
        self.actieveRegeling()
        self._flexibilisaties = list()
        
    def informatieOpslaan(self, informatie, kolomNaam):
        index = None
        for i, kolom in enumerate(informatie[0]):
            if kolom == kolomNaam: 
                index = i
                break
        if index == None:
            print("geen kolom gevonden met naam {}".format(kolomNaam))
            return
        return informatie[1][index]
    
    def pensioenenOpslaan(self, informatie, pensioeninformatie):
        pensioenen = list()
        for pensioen in pensioeninformatie:
            if pensioen.soortRegeling == "AOW":
                AOWpensioen = Pensioen(pensioen, 0, 0)
                AOWpensioen.pensioenleeftijd = self.pensioenLeeftijd()
                pensioenen.append(AOWpensioen)
            else:
                if informatie[1][pensioen.ouderdomsPensioen] != None: 
                    OP = int(informatie[1][pensioen.ouderdomsPensioen])
                    if pensioen.partnerPensioen == None: PP = 0
                    elif informatie[1][pensioen.partnerPensioen] == None: PP = 0
                    else: PP = int(informatie[1][pensioen.partnerPensioen])
                    pensioenen.append(Pensioen(pensioen, OP, PP))
        return pensioenen
    
    def activeerFlexibilisatie(self):
        flexibilisaties = list()
        for pensioen in self._pensioenen: 
            if pensioen.pensioenSoortRegeling != "AOW": flexibilisaties.append(Flexibilisering(pensioen))
        self._flexibilisaties = flexibilisaties
    
    def setAOWLeeftijd(self, jaar, maand, AOWjaar):
        for flexibilisatie in self._flexibilisaties:
            flexibilisatie.AOWJaar = jaar
            flexibilisatie.AOWMaand = maand
            if flexibilisatie.HL_Methode == "Opvullen AOW":
                flexibilisatie.HL_Jaar = int(AOWjaar-jaar)
    
    def actieveRegeling(self):
        if self._regeling == "Inactief": return
        selectie = list()
        for pensioen in self._pensioenen:
            if self._regeling in pensioen.pensioenNaam: selectie.append(pensioen)
        if len(selectie) == 0: return
        selectie.sort(reverse = True, key = lambda pensioen: pensioen.pensioenleeftijd)
        pensioen = selectie[0]
        pensioen.extraPensioen(self._ftLoon * self._pt)
    
    def pensioenLeeftijd(self):
        if self._geboortedatum > datetime(1997, 3, 31): return 70
        if self._geboortedatum > datetime(1993, 6, 30): return 69.75
        if self._geboortedatum > datetime(1989, 9, 30): return 69.5
        if self._geboortedatum > datetime(1985, 12, 31): return 69.25
        if self._geboortedatum > datetime(1982, 3, 31): return 69
        if self._geboortedatum > datetime(1979, 6, 30): return 68.75
        if self._geboortedatum > datetime(1976, 9, 30): return 68.5
        if self._geboortedatum > datetime(1972, 12, 31): return 68.25
        if self._geboortedatum > datetime(1970, 3, 31): return 68
        if self._geboortedatum > datetime(1967, 6, 30): return 67.75
        if self._geboortedatum > datetime(1963, 9, 30): return 67.5
        if self._geboortedatum > datetime(1961, 9, 30): return 67.25
        if self._geboortedatum > datetime(1960, 12, 31): return 67.25
        if self._geboortedatum > datetime(1957, 2, 28): return 67
        if self._geboortedatum > datetime(1956, 5, 31): return 66 + 10 / 12
        if self._geboortedatum > datetime(1955, 8, 31): return 66 + 7 / 12
        else: return 66 + 4 / 12
        
    @property
    def achternaam(self): return self._achternaam
    
    @property
    def tussenvoegsels(self): return self._tussenvoegsels
    
    @property
    def voorletters(self): return self._voorletters
    
    @property
    def geboortedatum(self): return self._geboortedatum
    
    @property
    def geslacht(self): return self._geslacht
    
    @property
    def burgelijkeStaat(self): return self._burgelijkeStaat
    
    @property
    def ftLoon(self): return self._ftLoon
    
    @property
    def pt(self): return self._pt
    
    @property
    def regeling(self): return self._regeling
    
    @property
    def rijNr(self): return self._rijNr
    
    @property
    def pensioenen(self): return self._pensioenen
    
    @property
    def flexibilisaties(self): return self._flexibilisaties
    