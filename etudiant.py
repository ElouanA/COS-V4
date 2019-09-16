from jure import Jure

class Etudiant(Jure):
    """ Classe définissant un étudiant """
    
    def __init__(self,nom,prenom,idprojet):
        Jure.__init__(self,nom,prenom)
        self.idprojet = idprojet
        
    def to_string(self):
        return(str(self.nom) + " " + str(self.prenom)+ " " + str(self.idprojet))