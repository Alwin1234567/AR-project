"""
Header
Hier komen alle libraries die in het programma gebruikt worden
"""


"""
Body
Hier komen alle functies
"""

class Pensioeninformatie():
    """
    Een class waarin de informatie over een pensioenfonds staat opgeslagen.
    
    naam : str
        De naam van het pensioenfonds.
    deelnemersbestandKolommen : tuple
        Een tulpl met de informatie over de kolom van het ouderdomspensioen en mogelijk partnerpensioen.
    gegevensRij : int
        De rij waarin informatie over het pensioenfonds staat.
    """
    
    def __init__(self, naam, deelnemersbestandKolommen, gegevensRij):
        self._naam = naam
        self._deelnemersbestandKolommen = deelnemersbestandKolommen
        self._gegevensRij = gegevensRij
        
    
    @property
    def naam(self):
        return self._naam
    
    @property
    def deelnemersbestandKolommen(self):
        return self._deelnemersbestandKolommen
    
    @property
    def gegevensRij(self):
        return self._gegevensRij