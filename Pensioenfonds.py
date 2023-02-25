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
    Een class waarin de informatie over een pensioenfonds van een deelnemer staat opgeslagen.
    
    book : xlwings.Book
        Het excel bestand waarin het programma runned.
    pensioenrij : int
        De rij waarin de informatie over het pensioenfonds staat.
    OPenPP : tuple
        Een tuple met de waarde van het ouderdomspensioen en het partnerpensioen.
    """
    
    def __init__(self, book, pensioenrij, OPenPP):
        pensioenleeftijdkolom = 4
        rentekolom = 5
        sterftetafelkolom = 6
        
        self.book = book
        self.gegevensPensioencontracten = book.sheets["Gegevens pensioencontracten"]
        
        self._pensioenleefijd = self.gegevensPensioencontracten.range((pensioenrij, pensioenleeftijdkolom)).value
        self._rente = float(self.gegevensPensioencontracten.range((pensioenrij, rentekolom)).options(numbers = float).value)
        self._sterftetafel = self.gegevensPensioencontracten.range((pensioenrij, sterftetafelkolom)).value
        
        self._ouderdomsPensioen = OPenPP[0]
        self._partnerPensioen = OPenPP[1]
    
    @property
    def pensioenleefijd(self):
        return self._pensioenleefijd
    
    @property
    def rente(self):
        return self._rente
    
    @property
    def sterftetafel(self):
        return self._sterftetafel
    
    @property
    def ouderdomsPensioen(self):
        return self._ouderdomsPensioen
    
    @property
    def partnerPensioen(self):
        return self._partnerPensioen
    