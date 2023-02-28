"""
Header
Hier komen alle libraries die in het programma gebruikt worden
"""
from Pensioen import Pensioen
from flex_keuzes import Flexibilisering
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
        self._pt = self.informatieOpslaan(informatie, "PT%")
        self._regeling = self.informatieOpslaan(informatie, "Regeling")
        self._rijNr = self.informatieOpslaan(informatie, "rijNr")
        self._pensioenen = self.pensioenenOpslaan(informatie, pensioeninformatie)
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
            if informatie[1][pensioen.ouderdomsPensioen] != None: 
                OP = informatie[1][pensioen.ouderdomsPensioen]
                if pensioen.partnerPensioen == None: PP = 0
                elif informatie[1][pensioen.partnerPensioen] == None: PP = 0
                else: PP = informatie[1][pensioen.partnerPensioen]
                pensioenen.append(Pensioen(pensioen, OP, PP))
        return pensioenen
    
    def actieveerFlexibilisatie(self):
        flexibilisaties = list()
        for pensioen in self._pensioenen: flexibilisaties.append(Flexibilisering(pensioen))
        self._flexibilisaties = flexibilisaties
            
        
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
    