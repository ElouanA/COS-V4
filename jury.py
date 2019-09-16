from jure import Jure

class Jury:
    """ Classe définissant un jury """
    
    def __init__(self,id):
        self.jures = []
        self.id = id
        
    def add_jure(self,jure):
        if isinstance(jure, Jure):
            self.jures.append(jure)
        