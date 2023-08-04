from openpyxl.styles import Font

# Police de caractère du document
police = 'Arial'

# Font des en-têtes
font_huit = Font(name=police, size=8)
font_huit_bold = Font(name=police, size=8, bold=True)
font_huit_green = Font(name=police, size=8, color='00B050')
font_huit_red = Font(name=police, size=8, color='9C0006')
font_neuf = Font(name=police, size=9)
font_dix = Font(name=police, size=10)
font_dix_bold = Font(name=police, size=10, bold=True)
font_dix_rouge = Font(name=police, size=10, color='00FF0000')
font_douze_bold = Font(name=police, size=12, bold=True)
