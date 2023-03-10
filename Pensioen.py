"""
Header
Hier komen alle libraries die in het programma gebruikt worden
"""


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
    
    def __init__(self, pensioen, OP, PP):
        self._pensioen = pensioen
        self._OP = OP
        self._PP = PP
        
    @property
    def pensioenVolNaam(self):
        return self._pensioen.pensioenVolNaam
    
    @property
    def pensioenNaam(self):
        return self._pensioen.pensioenNaam
    
    @property
    def pensioenleefijd(self):
        return self._pensioen.pensioenleefijd
    
    @property
    def rente(self):
        return self._pensioen.rente
    
    @property
    def sterftetafel(self):
        return self._pensioen.sterftetafel
    
    @property
    def pensioenKleurZacht(self):
        return self._pensioen.kleurZacht
    
    @property
    def pensioenKleurHard(self):
        return self._pensioen.kleurHard
    
    @property
    def ouderdomsPensioen(self):
        return self._OP
    
    @property
    def partnerPensioen(self):
        return self._PP