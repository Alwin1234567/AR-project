"""
Header
Hier komen alle libraries die in het programma gebruikt worden
"""
from datetime import date

"""
Body
Hier komen alle functies
"""
class Pensioen():
    """
    Een class waarin de informatie van een pensioen van een deelnemer staat opgeslagen.
    
    pensioen : Pensioen class
        hierin staat alle pensioen informatie behalve OP en PP.
    OP : float
        Het OP bedrag van de deelnemer bij dit pensioen.
    PP : float
        Het PP bedrag van de deelnemer bij dit pensioen, kan ook 0 zijn.
    """
    
    def __init__(self, pensioen, kolom1, kolom2):
        self._pensioen = pensioen
        self._actieveRegeling = False
        self._regelingFactor = 0
        if self._pensioen.soortRegeling == "DC":
            self._OP = 0
            self._PP = 0
            self._koopsom = kolom1
        elif self._pensioen.soortRegeling != "AOW":
            self._OP = kolom1
            self._PP = kolom2
            self._koopsom = 0
        else:
            self._OP = 0
            self._PP = 0
            self._koopsom = 0
            
    def extraPensioen(self, inkomen):
        self._actieveRegeling = True
        self._regelingsFactor = (inkomen - self._pensioen.franchise) * self._pensioen.opbouwpercentage
        
    
    @property
    def pensioenVolNaam(self):
        return self._pensioen.pensioenVolNaam
    
    @property
    def pensioenNaam(self): return self._pensioen.pensioenNaam
    
    @property
    def pensioenSoortRegeling(self): return self._pensioen.soortRegeling
    
    @property
    def pensioenleeftijd(self): return self._pensioen.pensioenleeftijd
    
    @property
    def rente(self): return self._pensioen.rente
    
    @property
    def sterftetafel(self): return self._pensioen.sterftetafel
    
    @property
    def opbouwpercentage(self): return self._pensioen.opbouwpercentage
    
    @property
    def franchise(self): return self._pensioen.franchise
    
    @property
    def opmerking(self): return self._pensioen.opmerking
    
    @property
    def pensioenKleurZacht(self): return self._pensioen.kleurZacht
    
    @property
    def actieveRegeling(self): return self._actieveRegeling
    
    @property
    def regelingsFactor(self): return self._regelingsFactor
    
    @property
    def pensioenKleurHard(self): return self._pensioen.kleurHard
    
    @property
    def ouderdomsPensioen(self): return self._OP
    
    @property
    def partnerPensioen(self): return self._PP
    
    @property
    def koopsom(self): return self._koopsom            
    
    @property
    def alleenstaandAOW(self): return self._pensioen.alleenstaandAOW
    
    @property
    def samenwondendAOW(self): return self._pensioen.samenwondendAOW