import glob
import calendar
import locale
from tkinter.filedialog import askdirectory
from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

locale.setlocale(locale.LC_ALL, 'fr_FR')

_MSG_ERROR_INT_ = "Veuillez saisir un nombre entier."
alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'


def genererfichiercompta(annee: int):
    """
    Génère le fichier excel avec ses onglets
    :param annee: Année pour laquelle le fichier de comptabilité est créé
    :return: Fichier de comptabilité
    """
    doc = Workbook()
    for i in range(12):
        sheet = doc.create_sheet(f"{calendar.month_name[i + 1].capitalize()} {annee}", i)
        mettreenformesheetmois(sheet)
        remplirsheetmois(doc, sheet)
    doc['Sheet'].title = f"Bilan {annee}"
    mettreenformesheetbilan(doc[f'Bilan {annee}'])
    remplirsheetbilan(doc)
    return doc


def verifnom(nom: str, path: str):
    """
    Attribut un nom de fichier n'existant pas dans l'emplacement courant
    :param nom: Nom du fichier à vérifier
    :param path: Dossier de destination du fichier
    :return: Nom définitif du fichier
    """
    # Récupération de la liste des fichiers présents dans le répertoire courant
    liste_fic = []
    for file in glob.glob(f"{path}/*.xlsx"):
        file = file.removesuffix('.xlsx')
        liste_fic.append(file.removeprefix(f'{path}\\'))

    # Compteur pour le nom modifié
    i = 1
    ajout = ''
    # Boucle cherchant si le nom de fichier existe déjà dans le répertoire
    while True:
        nomok = True
        for fic in liste_fic:
            if fic == f'{nom}{ajout}':
                ajout = f'({i})'
                i += 1
                nomok = False
            if not nomok:
                break
        if nomok:
            break
    return f'{nom}{ajout}'


def mettreenformesheetmois(sheet):
    """
    Dimensionne une feuille de calcul de comptabilité mensuelle
    :param sheet: Feuille de calcul sur laquelle doivent s'appliquer les propriétés
    :return:
    """
    # Fusion des cellules
    list_cells_a_merge = ('A1:D2', 'E1:AA1', 'E2:G2', 'H2:J2', 'K2:M2', 'N2:Z2')
    for i in range(len(list_cells_a_merge)):
        sheet.merge_cells(list_cells_a_merge[i])

    # Dimensionnement des colonnes
    sheet.column_dimensions['A'].width = 8.5
    sheet.column_dimensions['B'].width = 4.73
    sheet.column_dimensions['C'].width = 8.82
    sheet.column_dimensions['D'].width = 41.5
    sheet.column_dimensions['AA'].width = 10.91

    # Dimensionnement des lignes
    list_hauteur_ligne = (16, 15, 35, 24.5, 3.5)
    for i in range(len(list_hauteur_ligne)):
        sheet.row_dimensions[i + 1].height = list_hauteur_ligne[i]
    for i in range(6, nbligne):
        sheet.row_dimensions[i].height = 12.5

    # Font des en-têtes
    font_base = Font(name='Arial', size=8)
    font_intitule = Font(name='Arial', size=10)
    font_diff = Font(name='Arial', size=9)
    font_total = Font(name='Arial', size=10, bold=True)
    font_totaux = Font(name='Arial', size=10, color='00FF0000')
    font_titre = Font(name='Arial', size=12, bold=True)
    font_en_tete_niv_un = Font(name='Arial', size=8, bold=True)
    font_green = Font(name='Arial', size=8, color='00B050')
    font_red_alert = Font(name='Arial', size=8, color='9C0006')

    # Styles d'alignement
    align_base = Alignment(vertical='center')
    align_totaux_dates = Alignment(vertical='center', horizontal='right')
    align_ligne_trois = Alignment(horizontal='center', vertical='center', wrap_text=True)
    align_titre = Alignment(horizontal='center', vertical='center')

    # Définition des PatternFill
    fill_jaune = PatternFill(fgColor="00FFFF00", fill_type="solid")
    fill_gris = PatternFill(fgColor="C0C0C0", fill_type="solid")

    # Application des couleurs de fond de cellule
    for i in range(len(alphabet)):
        sheet[f'{alphabet[i]}5'].fill = fill_jaune
    sheet['AA5'].fill = fill_jaune
    for i in range(4):
        sheet[f'{alphabet[i]}4'].fill = fill_gris

    # Définition des mises en forme conditionnelles
    cond_format_red_alert = CellIsRule(operator='lessThan', formula=[0], stopIfTrue=False, font=font_red_alert)
    cond_format_green = CellIsRule(operator='greaterThanOrEqual', formula=[0], stopIfTrue=False, font=font_green)

    # Application des mises en forme conditionnelles
    liste_cell_cond_format = ('G4', 'J4')
    for cell in liste_cell_cond_format:
        sheet.conditional_formatting.add(cell, cond_format_red_alert)
        sheet.conditional_formatting.add(cell, cond_format_green)

    # Application des propriétés générales
    for row in sheet[f'A1:AA{nbligne - 1}']:
        for cell in row:
            cell.font = font_base
            cell.alignment = align_base
    sheet['AA2'].font = font_diff

    # Propriétés de la colonne AA (totaux)
    for i in range(6, nbligne):
        sheet[f'AA{i}'].font = font_totaux
        sheet[f'AA{i}'].alignment = align_totaux_dates
        sheet[f'A{i}'].alignment = align_totaux_dates

    # Propriétés de la colonne D (intitulé)
    for i in range(3, nbligne):
        sheet[f'D{i}'].font = font_intitule
    sheet['AA4'].font = font_intitule

    # Propriétés des en-têtes
    for row in sheet['A1:Z2']:
        for cell in row:
            cell.font = font_titre
    for row in sheet['A1:AA3']:
        for cell in row:
            cell.alignment = align_titre
    sheet['D4'].alignment = align_titre
    sheet['AA4'].alignment = align_titre
    cell = sheet['AA3'].font = font_total
    cell.alignment = align_titre

    # Application du retour à la ligne à la ligne 3
    for row in sheet['E3:Z3']:
        for cell in row:
            cell.alignment = align_ligne_trois

    list_en_tete_niv_un = ('E2', 'H2', 'K2', 'N2')
    for i in range(len(list_en_tete_niv_un)):
        sheet[list_en_tete_niv_un[i]].font = font_en_tete_niv_un

    # Définition des modèles de bordures
    thin = Side(border_style="thin")
    medium = Side(border_style="medium")
    thin_border = Border(left=thin, right=thin, top=thin, bottom=thin)
    medium_border_right = Border(left=thin, right=medium, top=thin, bottom=thin)
    medium_border_left = Border(left=medium, right=thin, top=thin, bottom=thin)
    medium_border_bottom = Border(left=thin, right=thin, top=thin, bottom=medium)
    medium_all_borders = Border(left=medium, right=medium, top=medium, bottom=medium)
    medium_top_bottom = Border(left=thin, right=thin, top=medium, bottom=medium)
    medium_coin_haut_gauche = Border(left=medium, right=thin, top=medium, bottom=medium)
    medium_coin_haut_droit = Border(left=thin, right=medium, top=medium, bottom=medium)
    medium_coin_bas_gauche = Border(left=medium, right=thin, top=thin, bottom=medium)
    medium_coin_bas_droit = Border(left=thin, right=medium, top=thin, bottom=medium)

    # Modèle de base
    for row in sheet[f'A3:AA{nbligne - 1}']:
        for cell in row:
            cell.border = thin_border
    # Applications des bordures spécifiques
    # Colonnes gauches
    list_left = ("A", "H", "K", "N")
    for col in list_left:
        for i in range(3, nbligne):
            sheet[f'{col}{i}'].border = medium_border_left
    # Colonnes droites
    list_right = ('D', 'Z', 'AA')
    for col in list_right:
        for i in range(3, nbligne):
            sheet[f'{col}{i}'].border = medium_border_right
    # Ligne 2
    for i in range(alphabet.index('E'), len(alphabet)):
        sheet[f'{alphabet[i]}2'].border = medium_all_borders
    # Ligne 3
    coin_gauche = ('A', 'E', 'H', 'K', 'N')
    coin_droit = ('D', 'G', 'J', 'M', 'Z')
    for i in range(0, len(alphabet)):
        if alphabet[i] in coin_gauche:
            sheet[f'{alphabet[i]}3'].border = medium_coin_haut_gauche
        elif alphabet[i] in coin_droit:
            sheet[f'{alphabet[i]}3'].border = medium_coin_haut_droit
        else:
            sheet[f'{alphabet[i]}3'].border = medium_top_bottom
    sheet['AA3'].border = medium_all_borders
    # Dernière ligne
    for i in range(0, len(alphabet)):
        if alphabet[i] in coin_gauche:
            sheet[f'{alphabet[i]}{nbligne - 1}'].border = medium_coin_bas_gauche
        elif alphabet[i] in coin_droit:
            sheet[f'{alphabet[i]}{nbligne - 1}'].border = medium_coin_bas_droit
        else:
            sheet[f'{alphabet[i]}{nbligne - 1}'].border = medium_border_bottom
    sheet[f'AA{nbligne - 1}'].border = Border(left=medium, right=medium, top=thin, bottom=medium)

    # Formats du contenu des cellules
    monetaire_euro = '#,##0.00 €'
    date_fr = 'dd/mm/yyyy'
    # Colonne AA (monétaire euros deux chiffre après la virgule
    for row in sheet[f'E4:AA{nbligne - 1}']:
        for cell in row:
            cell.number_format = monetaire_euro
    # Colonne A (dates)
    for i in range(6, nbligne):
        sheet[f'A{i}'].number_format = date_fr


def remplirsheetmois(doc, sheet):
    """
    Remplit la feuille de calcul avec le texte par défaut
    :param doc: Docuement au format .xlsx
    :param sheet: Feuille de calcul à remplir
    :return:
    """
    # Remplissage des titres
    sheet['A1'].value = 'Compte chèques'
    sheet['E1'].value = 'Feuille de comptabilité'
    sheet['D4'].value = 'totaux :'
    sheet['E2'].value = 'Caisse'
    sheet['H2'].value = 'Banque'
    sheet['K2'].value = 'Recettes'
    sheet['N2'].value = 'Dépenses'
    contenu_champs = ('DATE', 'N°', 'N° chq', 'Intitulé', 'Recettes', 'Dépenses', 'Situation', 'Recettes', 'Dépenses',
                      'Situation', 'Recettes diverses', 'Compte à régulariser', 'Virements internes',
                      'Virements internes', 'Epargne', 'Alimentat°', 'Produits entretien',
                      'Transport', 'Hygiène', 'Invest.', 'Santé', 'Assurances', 'Divers', 'Electricité',
                      'Eau', 'Impôts')
    # Remplissage des champs de la ligne 3 (en-têtes)
    for i in range(len(alphabet)):
        sheet[f'{alphabet[i]}3'].value = contenu_champs[i]
    sheet['AA3'].value = 'TOTAL'

    # Remplissage des formules de calcul de la ligne 4 (totaux)
    for i in range(4, len(alphabet)):
        sheet[f'{alphabet[i]}4'].value = f'=SUM({alphabet[i]}6:{alphabet[i]}71)'
    sheet['G4'].value = '=E4-F4'
    sheet['J4'].value = '=H4-I4'

    # Remplissage des formules de calcul de la colonne AA
    sheet['AA4'].value = '=SUM(N4:Z4)'
    for i in range(6, nbligne):
        sheet[f'AA{i}'].value = f'=E{i}+H{i}-SUM(K{i}:M{i})-F{i}-I{i}+SUM(N{i}:Z{i})'

    # Remplissage de la première ligne d'enregistrement
    sheet['D6'].value = 'Ouverture'
    month_dict = {'Janvier': '01', 'Février': '02', 'Mars': '03', 'Avril': '04', 'Mai': '05', 'Juin': '06',
                  'Juillet': '07', 'Août': '08', 'Septembre': '09', 'Octobre': '10', 'Novembre': '11', 'Décembre': '12'}
    mois = month_dict[sheet.title[:-5]]
    annee = int(sheet.title[-4:])
    sheet['A6'].value = f'01/{mois}/{annee}'

    # Application du report des situations des mois précédents
    if sheet.title[:-5] == 'Janvier':
        sheet['E6'].value = 0
        sheet['H6'].value = 0
    else:
        liste_sheet = doc.sheetnames
        indice_sheet = liste_sheet.index(sheet.title)
        sheet['E6'].value = f'=\'{liste_sheet[indice_sheet - 1]}\'!G4'
        sheet['H6'].value = f'=\'{liste_sheet[indice_sheet - 1]}\'!J4'


def definirparamfichier():
    """
    Permet à l'utilisateur de définir l'année comptable pour laquelle le document sera créé
    :return: Année de comptabilité définie par l'utilisateur et nombre de lignes de saisie
    """
    while True:
        try:
            annee = int(input("Veuillez saisir l'année pour laquelle vous souhaitez créer le document : "))
            nb = definirnblignesaisie()
            return annee, nb
        except ValueError:
            print(_MSG_ERROR_INT_)


def definirnblignesaisie():
    """
    Permet à l'utilisateur de définir le nombre de ligne qu'il pourra saisir pour chaque feuille de calcul
    :return: Nombre de ligne de saisie
    """
    while True:
        try:
            nb = int(input("Veuillez saisir le nombre de ligne de saisie : "))
            return nb
        except ValueError:
            print(_MSG_ERROR_INT_)


def mettreenformesheetbilan(sheet):
    """
    Met en forme la feuille de bilan comptable de l'année
    :param sheet: Feuille de calcul concernée
    :return:
    """

    # Dimensionnement des colonnes
    for i in range(1, alphabet.index('Q')):
        sheet.column_dimensions[f'{alphabet[i]}'].width = 9
    sheet.column_dimensions['A'].width = 12
    sheet.column_dimensions['Q'].width = 12

    # Dimensionnement des lignes
    list_hauteur_ligne = (16, 35, 24.5, 3.5)

    # Tableau général
    for i in range(len(list_hauteur_ligne)):
        sheet.row_dimensions[i+1].height = list_hauteur_ligne[i]
    for i in range(5, 16):
        sheet.row_dimensions[i + 1].height = 12.5

    # Tableau totaux
    list_hauteur_ligne = (15, 20.5, 24.5)
    for i in range(len(list_hauteur_ligne)):
        sheet.row_dimensions[i + 19].height = list_hauteur_ligne[i]

    # Fusion des cellules
    list_cells_a_merge = ('B1:D1', 'E1:Q1', 'E19:P19')
    for i in range(len(list_cells_a_merge)):
        sheet.merge_cells(list_cells_a_merge[i])

    # Définition des modèles de bordures
    thin = Side(border_style="thin")
    medium = Side(border_style="medium")
    thin_border = Border(left=thin, right=thin, top=thin, bottom=thin)
    medium_border_left = Border(left=medium, right=thin, top=thin, bottom=thin)
    medium_all_borders = Border(left=medium, right=medium, top=medium, bottom=medium)
    medium_top_bottom = Border(left=thin, right=thin, top=medium, bottom=medium)
    medium_contour_gauche = Border(left=medium, right=thin, top=medium, bottom=medium)
    medium_contour_droit = Border(left=thin, right=medium, top=medium, bottom=medium)
    medium_coin_bas_gauche = Border(left=medium, right=thin, top=thin, bottom=medium)
    medium_coin_haut_gauche = Border(left=medium, right=thin, top=medium, bottom=thin)
    medium_gauche_simple = Border(left=medium)

    # Définition des Fonts
    font_general = Font(name='Arial', size=8)
    font_general_gras = Font(name='Arial', size=8, bold=True)
    font_totaux = Font(name='Arial', size=10, color='00FF0000')
    font_dix = Font(name='Arial', size=10)
    font_dix_gras = Font(name='Arial', size=10, bold=True)

    # Définition des alignements
    align_center = Alignment(vertical='center', horizontal='center', wrap_text=True)
    align_droite = Alignment(vertical='center', horizontal='right')

    # Définition des PatternFill
    fill_jaune = PatternFill(fgColor="00FFFF00", fill_type="solid")
    fill_gris = PatternFill(fgColor="C0C0C0", fill_type="solid")

    # Modèle de base
    for row in sheet['A1:Q16']:
        for cell in row:
            cell.border = thin_border
    for row in sheet['E19:P21']:
        for cell in row:
            cell.border = thin_border
            cell.alignment = align_center

    # Application des alignements spécifiques
    for row in sheet['A1:Q2']:
        for cell in row:
            cell.alignment = align_center
    for row in sheet['A3:Q16']:
        for cell in row:
            cell.alignment = align_droite
    sheet['Q3'].alignment = align_center

    # Application des Fonts
    for row in sheet['A2:P21']:
        for cell in row:
            cell.font = font_general
    sheet['B1'].font = font_general_gras
    sheet['E1'].font = font_general_gras
    sheet['Q2'].font = font_dix_gras
    sheet['Q3'].font = font_dix
    sheet['E19'].font = font_dix
    for i in range(4, 16):
        sheet[f'Q{i+1}'].font = font_totaux

    # Application des couleurs de fonds de cellules
    for i in range(alphabet.index('R')):
        sheet[f'{alphabet[i]}4'].fill = fill_jaune
    sheet['A3'].fill = fill_gris

    # Formats du contenu des cellules
    monetaire_euro = '#,##0.00 €'
    for row in sheet['B3:Q16']:
        for cell in row:
            cell.number_format = monetaire_euro
    for i in range(alphabet.index('E'), alphabet.index('P') + 1):
        sheet[f'{alphabet[i]}21'].number_format = monetaire_euro

    # Récupère la liste des dépenses
    liste_depenses = []
    for i in range(alphabet.index('N'), alphabet.index('Z')):
        liste_depenses.append(sheet[f'{alphabet[i + 1]}3'].value)

    # Application des bordures spécifiques
    # Tableau récapitulatif
    for i in range(alphabet.index('E'), alphabet.index('E') + len(liste_depenses)):
        sheet[f'{alphabet[i]}19'].border = medium_all_borders
        if i == alphabet.index('E'):
            sheet[f'{alphabet[i]}20'].border = medium_contour_gauche
            sheet[f'{alphabet[i]}21'].border = medium_contour_gauche
        elif i == alphabet.index('E') + len(liste_depenses) - 1:
            sheet[f'{alphabet[i]}20'].border = medium_contour_droit
            sheet[f'{alphabet[i]}21'].border = medium_contour_droit
        else:
            sheet[f'{alphabet[i]}20'].border = medium_top_bottom
            sheet[f'{alphabet[i]}21'].border = medium_top_bottom
    # Tableau détaillé
    list_border_left = ('A', 'B', 'E', 'Q')
    for col in list_border_left:
        for i in range(16):
            if i == 0:
                rep = medium_coin_haut_gauche
            elif i == 15:
                rep = medium_coin_bas_gauche
            else:
                rep = medium_border_left
            sheet[f'{col}{i+1}'].border = rep
    for i in range(16):
        sheet[f'R{i+1}'].border = medium_gauche_simple
    for i in range(alphabet.index('Q') + 1):
        sheet[f'{alphabet[i]}1'].border = medium_all_borders
        if alphabet[i] == 'B' or alphabet[i] == 'E':
            border = medium_contour_gauche
        elif alphabet[i] == 'Q':
            border = medium_all_borders
        else:
            border = medium_top_bottom
        sheet[f'{alphabet[i]}2'].border = border
    for i in range(alphabet.index('P') + 1):
        sheet[f'{alphabet[i]}17'].border = Border(top=medium)


def remplirsheetbilan(doc):
    """
    Remplit la feuille de calcul du bilan annuel
    :param doc: Document au format .xlxs
    :return: 
    """
    # Récupère la liste des onglets
    liste_sheet = []
    for sheetname in doc.sheetnames:
        liste_sheet.append(sheetname)
    sheet = doc[liste_sheet[0]]

    # Récupère la liste des dépenses
    liste_depenses = []
    for i in range(alphabet.index('N'), alphabet.index('Z')):
        liste_depenses.append(sheet[f'{alphabet[i + 1]}3'].value)

    # Remplit la liste des dépenses dans l'onglet Bilan
    sheet = doc[liste_sheet[len(liste_sheet) - 1]]
    for i in range(len(liste_depenses)):
        sheet[f'{alphabet[i + 4]}2'].value = liste_depenses[i]
        sheet[f'{alphabet[i + 4]}20'].value = liste_depenses[i]

    # Remplit les en-têtes
    sheet['B1'].value = 'Banque'
    sheet['E1'].value = 'Dépenses'
    sheet['E19'].value = 'Dépenses mensuelles'
    sheet['Q2'].value = 'TOTAL'
    list_entete = ('DATE', 'Recettes', 'Dépenses', 'Situation')
    for i in range(len(list_entete)):
        sheet[f'{alphabet[i]}2'].value = list_entete[i]

    # Remplit les formules de calcul des sommes et des moyennes des dépenses
    sheet['B3'].value = '=SUM(B5:B16)'
    sheet['C3'].value = '=SUM(C5:C16)'
    sheet['D3'].value = '=B3-C3'
    for i in range(alphabet.index('E'), alphabet.index('E') + len(liste_depenses)):
        sheet[f'{alphabet[i]}3'].value = f'=SUM({alphabet[i]}5:{alphabet[i]}16)'
        sheet[f'{alphabet[i]}21'].value = f'=AVERAGE({alphabet[i]}5:{alphabet[i]}16)'

    # Remplit la cellule de total des dépenses
    index = alphabet.index('E')
    sheet['Q3'].value = f'=SUM(E3:{alphabet[index + len(liste_depenses) - 1]}3)'

    # Remplit la colonne des totaux
    for i in range(5, 5 + len(liste_depenses)):
        sheet[f'Q{i}'].value = f'=C{i}-SUM(E{i}:P{i})'

    # Remplit la colonne des mois
    for i in range(len(liste_sheet) - 1):
        sheet[f'A{i+5}'].value = liste_sheet[i]

    # Remplit le contenu du tableau principal
    # Partie "Banque"
    for i in range(5, 17):
        sheet[f'B{i}'].value = \
            f'=SUM(\'{liste_sheet[i - 5]}\'!E7:\'{liste_sheet[i - 5]}\'!E{nbligne - 1})+' \
            f'SUM(\'{liste_sheet[i - 5]}\'!H7:\'{liste_sheet[i - 5]}\'!H{nbligne - 1})'
        sheet[f'C{i}'].value = \
            f'=SUM(\'{liste_sheet[i - 5]}\'!F7:\'{liste_sheet[i - 5]}\'!F{nbligne - 1})+' \
            f'SUM(\'{liste_sheet[i - 5]}\'!I7:\'{liste_sheet[i - 5]}\'!I{nbligne - 1})'
        sheet[f'D{i}'].value = f'=B{i}-C{i}'
    # Partie "Dépenses"
        for j in range(alphabet.index('E'), alphabet.index('P') + 1):
            sheet[f'{alphabet[j]}{i}'].value = f'=\'{liste_sheet[i - 5]}\'!{alphabet[j + 10]}4'


def cheminfichier():
    """
    Définit le chemin complet du fichier (répertoire de destination et nom de fichier)
    :return: Chemin complet du fichier
    """
    # Chemin vers le répertoire de destination
    path = f'{askdirectory(title="Choix du dossier de destination")}'
    # Attribution du nom de fichier
    name = input("Veuillez donner un nom à votre fichier : ")
    # Vérification de la disponibilité du nom de fichier dans le répertoire
    name = verifnom(name, path)
    return f'{path}/{name}.xlsx'


# Définition du chemin de destination du fichier
chemin = cheminfichier()
# Définition des paramètres du fichier (année comptable et nombre de lignes de saisie)
donnees = definirparamfichier()
nbligne = donnees[1] + 6
document = genererfichiercompta(donnees[0])
# Enregistre le document dans l'endroit spécifié
document.save(chemin)
