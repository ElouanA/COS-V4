from etudiant import Etudiant
from jure import Jure
from cos import Cos
from openpyxl import load_workbook
path='FichierEntree_parJury.xlsx'
wb = load_workbook(filename = path)

cos = Cos(wb)

cos.initialize_cos_jury()

print(cos.to_string())
cos.generate_sortie_etudiant()
cos.generate_sortie_jury()
  
