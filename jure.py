class Jure:
    """ Classe définissant un juré """
    
    def __init__(self,nom,prenom):
        self.nom = nom
        self.prenom = prenom
        
        
    def to_string(self):
        return(str(self.nom) + " " + str(self.prenom))