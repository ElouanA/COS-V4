from jure import Jure

class Encadrant(Jure):
    """ Classe définissant un encadrant d'un projet """
    
    def __init__(self,nom,prenom,note,commentaire):
        Jure.__init__(self,nom,prenom,note,commentaire)
        