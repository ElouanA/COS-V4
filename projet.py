from etudiant import Etudiant

class Projet:
    """ Classe définissant un projet """
    
    def __init__(self,nom,id):
        self.nom = nom
        self.id = id
        
    def to_string(self):
        return(str(self.nom) + " : " + str(self.id))
            
            