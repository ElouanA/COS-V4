from critere import Critere

class Appreciation:
    """ Classe définissant une appréciation """
    
    def __init__(self,titre,nbPoints,partDeLaNote):
        self.titre = titre
        self.criteres = []
        self.nbPoints = nbPoints
        self.partDeLaNote = partDeLaNote
        
    def add_critere(self,critere):
        if isinstance(critere, Critere):
            self.criteres.append(critere)
        
    def to_string(self):
        res = self.titre + " \n"
        for i in self.criteres:
            res += "    "
            res += i.to_string()
            res += "\n"
            
        return(res)