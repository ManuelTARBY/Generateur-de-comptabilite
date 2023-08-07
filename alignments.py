from openpyxl.styles import Alignment

# Définition des alignements
align_base = Alignment(vertical='center')
align_droite = Alignment(vertical='center', horizontal='right')
align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
align_center_adjust = Alignment(horizontal='center', vertical='center', shrinkToFit=True)
align_titre = Alignment(horizontal='center', vertical='center')
