from appreciation import Appreciation

class Canevas:
    """ Classe définissant un canevas """
    
    def __init__(self,widthCritere,widthMarge1,widthPoints,widthPourcentage,widthNoteFinale,widthMarge2,widthMarge3,widthCommentaireJure,widthNoteJure,colors, heightJure,fontTitre,fontAppreciation,fontTitreCritere,fontDescriptionCritere,fontNoteCritere,fontNoteAppreciation,fontNotePourcentage,fontPourcentage,fontNoteFinale,fontTitreNoteFinale,fontTitreJure,fontColonneCommentaireJure,fontCommentaireJure,fontColonneNoteJure,fontNoteAppreciationJure,fontNoteCritereJure,titreColonneCommentaireJure,titreColonneNoteJure,titrePourcentage,titreNoteFinale,width_nom,width_prenom, width_projet,border):
        self.appreciations = []
        self.widthCritere = widthCritere
        self.widthMarge1 = widthMarge1
        self.widthPoints = widthPoints
        self.widthPourcentage = widthPourcentage
        self.widthNoteFinale = widthNoteFinale
        self.widthMarge2 = widthMarge2
        self.widthMarge3 = widthMarge3
        self.widthCommentaireJure = widthCommentaireJure
        self.widthNoteJure = widthNoteJure
        self.colors = colors
        self.heightJure = heightJure
        self.fontTitre = fontTitre
        self.fontAppreciation = fontAppreciation
        self.fontTitreCritere = fontTitreCritere
        self.fontDescriptionCritere = fontDescriptionCritere
        self.fontNoteCritere = fontNoteCritere
        self.fontNoteAppreciation = fontNoteAppreciation
        self.fontNotePourcentage = fontNotePourcentage
        self.fontPourcentage = fontPourcentage
        self.fontNoteFinale = fontNoteFinale
        self.fontTitreNoteFinale = fontTitreNoteFinale
        self.fontTitreJure = fontTitreJure
        self.fontColonneCommentaireJure = fontColonneCommentaireJure
        self.fontCommentaireJure = fontCommentaireJure
        self.fontColonneNoteJure = fontColonneNoteJure
        self.fontNoteAppreciationJure = fontNoteAppreciationJure
        self.fontNoteCritereJure = fontNoteCritereJure
        self.titreColonneCommentaireJure = titreColonneCommentaireJure
        self.titreColonneNoteJure = titreColonneNoteJure
        self.titrePourcentage = titrePourcentage
        self.titreNoteFinale = titreNoteFinale
        self.width_nom = width_nom
        self.width_prenom = width_prenom
        self.width_projet = width_projet
        self.border = border
        
    def add_appreciation(self,appreciation):
        if isinstance(appreciation, Appreciation):
            self.appreciations.append(appreciation)
            
            
    def to_string(self):
        res = ""
        for i in self.appreciations:
            res += i.to_string()
            res += "\n"
        return(res)