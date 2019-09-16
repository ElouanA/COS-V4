class Critere:
    """ Classe définissant un critère """
    
    def __init__(self,titre,description,points,hauteur):
        self.titre = titre
        self.description = description
        self.points = points
        self.hauteur = hauteur
        
    def to_string(self):
        res = self.titre + "\n"
        res += "    "
        res += str(self.description)
        return(res) 