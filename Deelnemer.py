"""
Header
Hier komen alle libraries die in het programma gebruikt worden
"""
from Pensioenfonds import Pensioenfonds
from Pensioeninformatie import Pensioeninformatie
from flex_keuzes import Flexibilisering
"""
Body
Hier komen alle functies
"""
class Deelnemer():
    """
    Een class waarin de informatie over een deelnemer staat opgeslagen.
    
    book : xlwings.Book
        Het excel bestand waarin het programma runned.
    informatie : list
        Een lijst met daarin de naam van een attribuut en de waarde van de attribuut.
    """
    
    def __init__(self, book, informatie):
        self.book = book
        self._achternaam = self.informatieOpslaan(informatie, "Naam")
        self._tussenvoegsels = self.informatieOpslaan(informatie, "tussenvoegsels")
        self._voorletters = self.informatieOpslaan(informatie, "voorletter")
        self._geboortedatum = self.informatieOpslaan(informatie, "Geboortedatum")
        self._geslacht = self.informatieOpslaan(informatie, "Geslacht")
        self._burgelijkeStaat = self.informatieOpslaan(informatie, "Burg. Staat")
        self._ftLoon = self.informatieOpslaan(informatie, "FT loon")
        self._pt = self.informatieOpslaan(informatie, "PT%")
        self._regeling = self.informatieOpslaan(informatie, "Regeling")
        self._rijNr = self.informatieOpslaan(informatie, "rijNr")
        self._pensioenen = self.pensioenenOpslaan(informatie)
        self._flexibilsaties = list()
        
        
        
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
    
    def pensioenenOpslaan(self, informatie):
        pensioenenlijst = self.pensioenenInit()
        pensioenen = list()
        for pensioen in pensioenenlijst:
            bedragen = list()
            for kolom in pensioen.deelnemersbestandKolommen:
                if kolom == None: bedragen.append(0)
                else: bedragen.append(informatie[1][kolom])
            pensioenen.append(Pensioenfonds(self.book, pensioen.gegevensRij, bedragen))
        return pensioenen
                
    def pensioenenInit(self):
        # ZL = Pensioeninformatie("ZL", (10, None), 3)
        Aegon65 = Pensioeninformatie("Aegon65", (11, None), 4)
        Aegon67 = Pensioeninformatie("Aegon67", (12, None), 5)
        NN65 = Pensioeninformatie("NN65", (13, 14), 6)
        NN67 = Pensioeninformatie("NN67", (15, 16), 7)
        PF_VLC68 = Pensioeninformatie("PF_VLC68", (17, 18), 8)
        return [Aegon65, Aegon67, NN65, NN67, PF_VLC68]
    
    def actieveerFlexibilisatie(self):
        flexibilsaties = list()
        for pensioen in self._pensioenen: flexibilsaties.append(Flexibilisering(pensioen))
        self._flexibilsaties = flexibilsaties
            
        
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
    