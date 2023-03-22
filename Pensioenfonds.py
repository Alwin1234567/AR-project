"""
Header
Hier komen alle libraries die in het programma gebruikt worden
"""


"""
Body
Hier komen alle functies
"""
class Pensioenfonds():
    """
    Een class waarin de informatie over een pensioenfonds staat opgeslagen.
    
    gegevensSheet : xlwings.Book.sheets
        Het excel sheet waarin de pensioen informatie staat.
    kolommen : dict
        dictionary met daarin de kolommen nummers.
    pensioen : tuple(int)
        Een tuple met de kolom van het ouderdomspensioen en het partnerpensioen.
    """
    
    def __init__(self, gegevensSheet, kolommen, pensioen):
        OPenPP = pensioen[0]
        gegevensRij = pensioen[1]
        
        
        self._naam = gegevensSheet.range((gegevensRij, kolommen["naamkolom"])).value
        self._volNaam = ""
        self.volledigeNaam()
        
        self._soortRegeling = gegevensSheet.range((gegevensRij, kolommen["soortRegeling"])).value
        self._pensioenleefijd = gegevensSheet.range((gegevensRij, kolommen["pensioenleeftijdkolom"])).value
        self._rente = float(gegevensSheet.range((gegevensRij, kolommen["rentekolom"])).options(numbers = float).value)
        self._sterftetafel = gegevensSheet.range((gegevensRij, kolommen["sterftetafelkolom"])).value
        self._opbouwpercentage = gegevensSheet.range((gegevensRij, kolommen["opbouwpercentage"])).value
        self._franchise = gegevensSheet.range((gegevensRij, kolommen["franchise"])).options(numbers = int).value
        self._opmerking = gegevensSheet.range((gegevensRij, kolommen["opmerking"])).value
        self._kleurZacht = tuple([int(kleur) for kleur in gegevensSheet.range((gegevensRij, kolommen["kleurzachtkolom"])).value.split(",")])
        self._kleurHard = tuple([int(kleur) for kleur in gegevensSheet.range((gegevensRij, kolommen["kleurhardkolom"])).value.split(",")])
        if type(self._opmerking) == str: reduced = self._opmerking.replace(",", ".").split("; ")
        else: reduced = str()
        if self._soortRegeling == "AOW": self.setvars(0, 0, 0, round(float(reduced[0])), round(float(reduced[1])))
        else: self.setvars(OPenPP[0], OPenPP[1], 0, 0, 0)
            
    def setvars(self, OP, PP, koopsom, alleenstaand, samenwonend):
        self._ouderdomsPensioen = OP
        self._partnerPensioen = PP
        self._koopsom = koopsom
        self._alleenstaandAOW = alleenstaand
        self._samenwondendAOW = samenwonend
    
    def volledigeNaam(self):
        if self._naam == "ZL": self._volNaam = "ZwitserLeven"
        elif self._naam == "Aegon OP65": self._volNaam = "Aegon 65"
        elif self._naam == "Aegon OP67": self._volNaam = "Aegon 67"
        elif self._naam == "NN OP65": self._volNaam = "Nationale Nederlanden 65"
        elif self._naam == "NN OP67": self._volNaam = "Nationale Nederlanden 67"
        elif self._naam == "PF VLC OP68": self._volNaam = "Pensioenfonds VLC 68"
        elif self._naam == "AOW": self._volnaam = "AOW"
        else: print(self._naam)
            
    @property
    def pensioenVolNaam(self): return self._volNaam
    
    @property
    def pensioenNaam(self): return self._naam
    
    @property
    def soortRegeling(self): return self._soortRegeling
    
    @property
    def pensioenleeftijd(self): return self._pensioenleefijd
    
    @property
    def rente(self): return self._rente
    
    @property
    def sterftetafel(self): return self._sterftetafel
    
    @property
    def opbouwpercentage(self): return self._opbouwpercentage
    
    @property
    def franchise(self): return self._franchise
    
    @property
    def opmerking(self): return self._opmerking
    
    @property
    def kleurZacht(self): return self._kleurZacht
    
    @property
    def kleurHard(self): return self._kleurHard
    
    @property
    def ouderdomsPensioen(self): return self._ouderdomsPensioen
    
    @property
    def partnerPensioen(self): return self._partnerPensioen
    
    @property
    def alleenstaandAOW(self): return self._alleenstaandAOW
    
    @property
    def samenwondendAOW(self): return self._samenwondendAOW