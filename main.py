import glob
import calendar
import locale
import tkinter
from tkinter import *
from tkinter.filedialog import askdirectory
from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Font, Alignment, PatternFill
from borders import *

locale.setlocale(locale.LC_ALL, 'fr_FR')

alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'


def genererfichiercompta():
    """
    Génère le fichier excel avec ses onglets
    :return: Fichier de comptabilité
    """
    annee = int(lannee.get())
    doc = Workbook()
    for i in range(12):
        sheet = doc.create_sheet(f"{calendar.month_name[i + 1].capitalize()} {annee}", i)
        mettreenformesheetmois(sheet)
        remplirsheetmois(doc, sheet)
    doc['Sheet'].title = f"Bilan {annee}"
    mettreenformesheetbilan(doc[f'Bilan {annee}'])
    remplirsheetbilan(doc)
    return doc


def verifnom():
    """
    Attribut un nom de fichier n'existant pas dans le repertoire choisi
    :return: Nom définitif du fichier
    """
    # Récupère la liste des fichiers présents dans le répertoire courant
    liste_fic = []
    # for file in glob.glob(f"{path}/*.xlsx"):
    path = lblpath['text']
    for file in glob.glob(f"{path}/*.xlsx"):
        file = file.removesuffix('.xlsx')
        # liste_fic.append(file.removeprefix(f'{path}\\'))
        liste_fic.append(file.removeprefix(f'{path}\\'))

    # Compteur pour le nom modifié
    i = 1
    ajout = ''
    nom = lenom.get()

    # Cherche si le nom de fichier existe déjà dans le répertoire à partir de la liste créée juste avant
    while True:
        nomok = True
        for fic in liste_fic:
            # if fic == f'{nom}{ajout}':
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
    :param sheet: Feuille de calcul sur laquelle les propriétés doivent s'appliquer
    :return:
    """
    nbligne = int(lignes.get()) + 6
    # Fusion des cellules
    list_cells_a_merge = ('A1:D2', 'E1:AA1', 'E2:G2', 'H2:J2', 'K2:M2', 'N2:Z2')
    for plage in list_cells_a_merge:
        sheet.merge_cells(plage)

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

    # Définition des zones multi-cellules
    bottom = nbligne - 1
    liste_zones = ('E2:G2', 'H2:J2', 'K2:M2', 'N2:Z2', 'A3:D3', 'E3:G3', 'H3:J3', 'K3:M3', 'N3:Z3', f'A4:D{bottom}',
                   f'E4:G{bottom}', f'H4:J{bottom}', f'K4:M{bottom}', f'N4:Z{bottom}', f'AA4:AA{bottom}')

    # Application des bordures aux zones multi-cellules
    for zone in liste_zones:
        appliquerbordures(sheet[zone])

    # Définition des zones mono-cellules
    liste_cellules = 'AA3'

    # Application des règles aux zones mono-cellules
    sheet[liste_cellules].border = medium_all_borders

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
    nbligne = int(lignes.get()) + 6
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
            cell.border = thin_all_borders
    for row in sheet['E19:P21']:
        for cell in row:
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

    # Définition des zones multi-cellules
    liste_zones = ('A1:Q16', 'B3:D16', 'E3:P16', 'B1:D1', 'E1:Q1', 'B2:D2', 'E2:P2', 'A3:A16', 'Q3:Q16',
                   'E19:P19', 'E20:P20', 'E21:P21')

    # Application des bordures aux zones multi-cellules
    for zone in liste_zones:
        appliquerbordures(sheet[zone])

    # Définition des zones mono-cellules
    liste_cellules = 'A2'

    # Application des règles aux zones mono-cellules
    sheet[liste_cellules].border = medium_all_borders


def remplirsheetbilan(doc):
    """
    Remplit la feuille de calcul du bilan annuel
    :param doc: Document au format .xlxs
    :return: 
    """
    nbligne = int(lignes.get()) + 6
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
    """
    # Chemin vers le répertoire de destination
    path = f'{askdirectory(title="Choix du dossier de destination")}'
    # Attribution du nom de fichier
    lblpath['text'] = f'{path}'


def appliquerbordures(zone):
    """
    Applique les bordures à une zone de feuille de calcul (medium aux bordures extérieures et thin aux autres)
    :param zone: Zone sur laquelle porte l'application des bordures
    :return:
    """
    # Obtient le nombre de lignes dans la zone
    nb = 0
    for _ in zone:
        nb += 1
    idrow = 0
    if nb == 1:
        for row in zone:
            idcell = 0
            for cell in row:
                if idcell == 0:
                    cell.border = medium_contour_left
                elif idcell == len(row) - 1:
                    cell.border = medium_contour_right
                else:
                    cell.border = medium_top_bottom
                idcell += 1
    else:
        for row in zone:
            idcell = 0
            if len(row) == 1:
                for cell in row:
                    if idrow == 0:
                        cell.border = medium_contour_top
                    elif idrow == nb - 1:
                        cell.border = medium_contour_bottom
                    else:
                        cell.border = medium_left_right
            else:
                for cell in row:
                    if idrow == 0:
                        if idcell == 0:
                            cell.border = medium_coin_top_left
                        elif idcell == len(row) - 1:
                            cell.border = medium_coin_top_right
                        else:
                            cell.border = medium_border_top
                    elif idrow == nb - 1:
                        if idcell == 0:
                            cell.border = medium_coin_bottom_left
                        elif idcell == len(row) - 1:
                            cell.border = medium_coin_bottom_right
                        else:
                            cell.border = medium_border_bottom
                    else:
                        if idcell == 0:
                            cell.border = medium_border_left
                        elif idcell == len(row) - 1:
                            cell.border = medium_border_right
                        else:
                            cell.border = thin_all_borders
                    idcell += 1
            idrow += 1


def verifcontenu():
    """
    Vérifie si le contenu des champs de la fenêtre sont corrects
    :return: True si ok, False dans le cas contraire
    """
    # Vérifie si l'année renseignée est correcte
    if lannee.get() == '':
        lblerror['text'] = "Vous devez saisir une année"
        txtannee.focus()
        return False
    else:
        try:
            year = int(lannee.get())
            if not 999 < year < 10000:
                lblerror['text'] = "L'année doit être comprise entre 1000 et 9999"
                return False
        except:
            lblerror['text'] = "L'année doit être un nombre entier"
            return False
        # Vérifie si le nombre de lignes de saisie renseigné est correct
        if lignes.get() == '':
            lblerror['text'] = "Vous devez entrer un nombre de lignes de saisie"
            txtnblign.focus()
            return False
        else:
            try:
                ln = int(lignes.get())
                if not 9 < ln < 1000:
                    lblerror['text'] = "Le nombre de lignes doit être compris entre 10 et 999"
                    return False
            except:
                lblerror['text'] = "Le nombre de lignes doit être un nombre entier"
                return False
            # Vérifie si un nom de fichier a été saisi
            if lenom.get() == '':
                lblerror['text'] = "Vous devez entrer un nom de fichier"
                txtnom.focus()
                return False
            # Vérifie si un répertoire de destination a été choisi
            elif lblpath['text'] == '':
                lblerror['text'] = "Vous devez sélectionner un dossier de destination"
                btndir.focus()
                return False
            else:
                lblerror['text'] = ''
                return True


def creerfichier():
    """
    Génère le fichier de comptabilité
    :return:
    """
    if verifcontenu():
        document = genererfichiercompta()
        # Crée le chemin complet pour l'enregistrement du fichier
        nom = verifnom()
        dest = lblpath['text']
        path = f'{dest}/{nom}.xlsx'
        # Enregistre le document dans le répertoire choisi
        document.save(path)
        fenetre.destroy()


# Création de l'interface graphique
fenetre = Tk()
lenom = tkinter.StringVar(name='nomfic', master=fenetre)
lannee = tkinter.StringVar(name='annee', master=fenetre)
lignes = tkinter.StringVar(name='nbligne', master=fenetre)
fenetre.geometry('380x285')
fenetre.title("Création d'un fichier de comptabilité")
fenetre.resizable(width=False, height=False)

# Contenu de l'interface graphique
# Titre
lblprs = Label(fenetre, text="Renseignez les informations", font=('Arial', 14, "bold"))
lblprs.pack(pady=10)

# Zone du choix de l'année
fenannee = Frame(fenetre)
fenannee.pack(pady=6)
lblannee = Label(fenannee, text="Année comptable :", font=('Arial', 10))
lblannee.pack(side=LEFT)
txtannee = Entry(fenannee, border=2, width=7, justify='center', font=('Arial', 10), textvariable=lannee)
txtannee.pack(side=LEFT)

# Zone du choix de nbde lignes
fennblign = Frame(fenetre)
fennblign.pack(pady=6)
lblnblign = Label(fennblign, text="Nombre de lignes de saisie :", font=('Arial', 10))
lblnblign.pack(side=LEFT)
txtnblign = Entry(fennblign, border=2, width=5, justify='center', font=('Arial', 10), textvariable=lignes)
txtnblign.pack(side=LEFT)

# Zone du choix de nom
fennom = Frame(fenetre)
fennom.pack(pady=6)
lblnom = Label(fennom, text="Nom du fichier :", font=('Arial', 10))
lblnom.pack(side=LEFT)
txtnom = Entry(fennom, border=2, width=25, font=('Arial', 10), textvariable=lenom)
txtnom.pack(side=LEFT)

# Zone du choix de répertoire
fendir = Frame(fenetre)
fendir.pack()
lbldir = Label(fendir, text="Destination :", font=('Arial', 10))
lbldir.pack(side=LEFT)
btndir = Button(fendir, text="...", width=3, font=('Arial', 10), command=cheminfichier)
btndir.pack(side=LEFT, padx=2)
lblpath = Label(fendir, width=25, text='', font=('Arial', 10))
lblpath.pack(side=LEFT)

# Bouton d'envoi des réponses
btnsend = Button(fenetre, text="Lancer la génération", font=('Arial', 12, 'bold'), bg='darkgreen', fg='white',
                 command=creerfichier)
btnsend.pack(pady=20)

# Label qui affiche les messages d'erreur
lblerror = Label(fenetre, text='', font=('Arial', 10), fg='red')
lblerror.pack()

fenetre.mainloop()
