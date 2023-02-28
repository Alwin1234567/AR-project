"""
Header
Hier komen alle libraries die in het programma gebruikt worden
"""


"""
Body
Hier komen alle functies
"""
class Pensioen():
    
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
    def ouderdomsPensioen(self):
        return self._OP
    
    @property
    def partnerPensioen(self):
        return self._PP