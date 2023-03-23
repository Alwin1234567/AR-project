"""
Header
Hier komen alle libraries die in het programma gebruikt worden
"""


"""
Body
Hier komen alle functies
"""

class Flexibilisering:

    """
    Een object van deze klasse bevat alle huidige flexibiliseringskeuzes waarmee
    verder gerekend moet worden. Voor elke regeling kan een object van deze klasse
    gemaakt worden.
    
    naam : str
        Naam van de pensioenregeling
    
    actief : bool
        Geeft aan of de deelnemer geld heeft staan bij deze pensioenregeling
    """


    def __init__(self, pensioen, actief=False):
        # self._naam = naam
        self._actief = actief
        
        self._pensioen = pensioen
        # Pensioenleeftijd
        self._leeftijd_Actief = False
        self._pLeeftijdJaar = 60
        self._pLeeftijdMaand = 0
        
        # OP/PP verhouding
        self._OP_PP_Actief = False
        self._OP_PP_UitruilenVan = "OP naar PP"
        self._OP_PP_Methode = "Percentage"
        self._OP_PP_Verhouding_OP = 100
        self._OP_PP_Verhouding_PP = 70
        self._OP_PP_Percentage = 0
        
        # Hoog/laag constructie
        self._HL_Actief = False
        self._HL_Volgorde = "Hoog-laag"
        self._HL_Methode = "Verhouding" 
        self._HL_Verhouding_Hoog = 100
        self._HL_Verhouding_Laag = 75
        self._HL_Verschil = 0
        self._HL_Jaar = 5
        self._HL_Maand = 0
        
        # OP en PP waardes
        self._OP_Hoog = self._pensioen.ouderdomsPensioen
        self._OP_Laag = self._pensioen.ouderdomsPensioen
        self._PP = self._pensioen.partnerPensioen
        
        # aanvullen AOW
        self._AOWJaar = 60
        self.AOWMaand = 0
        
        
        
    @property
    def naam(self):
        return self._naam
    
    @property
    def actief(self):
        return self._actief

    # --- pensioenleeftijd ---
    @property
    def leeftijd_Actief(self):
        return self._leeftijd_Actief
    
    @leeftijd_Actief.setter
    def leeftijd_Actief(self, leeftijd_Actief):
        self._leeftijd_Actief = leeftijd_Actief
    
    @property
    def leeftijdJaar(self):
        return self._pLeeftijdJaar
    
    @leeftijdJaar.setter
    def leeftijdJaar(self, pLeeftijdJaar):
        self._pLeeftijdJaar = pLeeftijdJaar
    
    @property
    def leeftijdMaand(self):
        return self._pLeeftijdMaand
    
    @leeftijdMaand.setter
    def leeftijdMaand(self, pLeeftijdMaand):
        self._pLeeftijdMaand = pLeeftijdMaand
    
    @property
    def leeftijdJaarMaand(self): return self._pLeeftijdJaar + self._pLeeftijdMaand / 12
    
    # --- OP/PP uitruiling ---
    # >>> Actief
    @property
    def OP_PP_Actief(self):
        return self._OP_PP_Actief
    
    @OP_PP_Actief.setter
    def OP_PP_Actief(self, uitruiling_Actief):
        self._OP_PP_Actief = uitruiling_Actief
    
    # >>> Uitruilen van ... naar ...
    @property
    def OP_PP_UitruilenVan(self):
        return self._OP_PP_UitruilenVan
    
    @OP_PP_UitruilenVan.setter
    def OP_PP_UitruilenVan(self, UitruilenVan):
        self._OP_PP_UitruilenVan = UitruilenVan
    
    # >>> Methode
    @property
    def OP_PP_Methode(self):
        return self._OP_PP_Methode
    
    @OP_PP_Methode.setter
    def OP_PP_Methode(self, OP_PP_Methode):
        self._OP_PP_Methode = OP_PP_Methode
    
    # >>> Verhouding OP
    @property
    def OP_PP_Verhouding_OP(self):
        return self._OP_PP_Verhouding_OP
    
    @OP_PP_Verhouding_OP.setter
    def OP_PP_Verhouding_OP(self, OP_PP_Verhouding_OP):
        self._OP_PP_Verhouding_OP = OP_PP_Verhouding_OP

    # >>> Verhouding PP
    @property
    def OP_PP_Verhouding_PP(self):
        return self._OP_PP_Verhouding_PP
    
    @OP_PP_Verhouding_PP.setter
    def OP_PP_Verhouding_PP(self, OP_PP_Verhouding_PP):
        self._OP_PP_Verhouding_PP = OP_PP_Verhouding_PP
    
    # >>> Percentage OP/PP of PP/OP
    @property
    def OP_PP_Percentage(self):
        return self._OP_PP_Percentage
    
    @OP_PP_Percentage.setter
    def OP_PP_Percentage(self, OP_PP_Percentage):
        self._OP_PP_Percentage = OP_PP_Percentage
    
    # --- Hoog/laag constructie ---
    
    # >>> Actief
    @property
    def HL_Actief(self):
        return self._HL_Actief
    
    @HL_Actief.setter
    def HL_Actief(self, HL_Actief):
        self._HL_Actief = HL_Actief
        
    
    # >>> Volgorde
    @property
    def HL_Volgorde(self):
        return self._HL_Volgorde
    
    @HL_Volgorde.setter
    def HL_Volgorde(self, HL_Volgorde):
        self._HL_Volgorde = HL_Volgorde
    
    # >>> Methode
    @property
    def HL_Methode(self):
        return self._HL_Methode
    
    @HL_Methode.setter
    def HL_Methode(self, HL_Methode):
        self._HL_Methode = HL_Methode
    
    # >>> Verhouding Hoog
    @property
    def HL_Verhouding_Hoog(self):
        return self._HL_Verhouding_Hoog
    
    @HL_Verhouding_Hoog.setter
    def HL_Verhouding_Hoog(self, HL_Verhouding_Hoog):
        self._HL_Verhouding_Hoog = HL_Verhouding_Hoog

    # >>> Verhouding Laag
    @property
    def HL_Verhouding_Laag(self):
        return self._HL_Verhouding_Laag
    
    @HL_Verhouding_Laag.setter
    def HL_Verhouding_Laag(self, HL_Verhouding_Laag):
        self._HL_Verhouding_Laag = HL_Verhouding_Laag
    
    # >>> Percentage H/L of L/H
    @property
    def HL_Verschil(self):
        return self._HL_Verschil
    
    @HL_Verschil.setter
    def HL_Verschil(self, HL_Verschil):
        self._HL_Verschil = HL_Verschil
    
    # >>> Aantal jaren eerste periode
    @property
    def HL_Jaar(self):
        return self._HL_Jaar
    
    @HL_Jaar.setter
    def HL_Jaar(self, HL_Jaar):
        self._HL_Jaar = HL_Jaar
    
    @property
    def pensioen(self): return self._pensioen
    
    # --- OP/PP waardes ---
    
    # >>> ouderdomsensioen hoog 
    @property
    def ouderdomsPensioenHoog(self): return self._OP_Hoog
    
    @ouderdomsPensioenHoog.setter
    def ouderdomsPensioenHoog(self, OP_Hoog): self._OP_Hoog = OP_Hoog
    
    # >>> ouderdomsensioen laag
    @property
    def ouderdomsPensioenLaag(self): return self._OP_Laag
    
    @ouderdomsPensioenLaag.setter
    def ouderdomsPensioenLaag(self, OP_Laag): self._OP_Laag = OP_Laag
    
    # >>> partnerpensioen
    @property
    def partnerPensioen(self): return self._PP
    
    @partnerPensioen.setter
    def partnerPensioen(self, PP): self._PP = PP
    
    # --- AOW waardes ---
    
    # >>> aow jaar
    @property
    def AOWJaar(self): return self._AOWJaar
    
    @AOWJaar.setter
    def AOWJaar(self, jaar): self._AOWJaar = jaar
    
    # >>> aow maand
    @property
    def AOWMaand(self): return self._AOWMaand
    
    @AOWMaand.setter
    def AOWMaand(self, maand): self._AOWMaand = maand
    