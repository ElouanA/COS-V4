import unittest
from cos import Cos
from projet import Projet
from canevas import Canevas
from jure import Jure

class CosTest(unittest.TestCase):
    
    def setUp(self):
        """Initialisation des tests."""
        self.cos = Cos()
        
    def test_add_canevas(self):
        canevas = Canevas(None,None)
        self.cos.add_canevas(canevas)
        self.assertEqual(len(self.cos.canevas),1)
        
    def test_add_projet(self):
        projet = Projet(None,None)
        self.cos.add_projet(projet)
        self.assertEqual(len(self.cos.projets),1)
        
    def test_add2_projet(self):
        projet = Projet(None,1)
        self.cos.add_projet(projet)
        self.assertEqual(len(self.cos.projets),1)
        
    def test_add3_projet(self):
        projet = Projet(None,1)
        self.cos.add_projet(projet)
        projet = Projet(None,2)
        self.cos.add_projet(projet)
        self.assertEqual(len(self.cos.projets),2)
        
    def test_add_jure(self):
        jure = Jure(None,None,None,None,)
        self.cos.add_jure(jure)
        self.assertEqual(len(self.cos.jures),1)
        

if __name__ == '__main__':
    unittest.main()