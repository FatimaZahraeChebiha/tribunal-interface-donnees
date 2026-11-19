"""Export des tableaux vers Excel : en-tête avec logo, bordures, ligne المجموع mise en valeur."""

import os
import tempfile
import warnings

from openpyxl import Workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.drawing.image import Image as ImageExcel
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from . import config

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl.styles.stylesheet")

LIGNE_ENTETE = 6  # les 5 premières lignes sont réservées au logo
CENTRE = Alignment(horizontal="center", vertical="center")
BORDURE = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
FOND_ENTETE = PatternFill(start_color="B7DEE8", end_color="B7DEE8", fill_type="solid")
FOND_TOTAL = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")


def generer_nom_fichier(base_name="النتائج", dossier=None):
    """Propose un nom de fichier libre : base.xlsx, puis base_1.xlsx, base_2.xlsx..."""
    dossier = dossier or os.path.expanduser("~")
    chemin = os.path.join(dossier, f"{base_name}.xlsx")
    compteur = 1
    while os.path.exists(chemin):
        chemin = os.path.join(dossier, f"{base_name}_{compteur}.xlsx")
        compteur += 1
    return chemin


def _ajouter_logo(ws, nb_colonnes):
    """Place le logo (facultatif) au centre de l'en-tête."""
    if not config.LOGO.exists():
        return
    image = ImageExcel(str(config.LOGO))
    image.width, image.height = config.TAILLE_LOGO
    ws.add_image(image, get_column_letter(nb_colonnes // 2 + 1) + "2")


def ecrire_feuille(ws, df):
    """Écrit un DataFrame dans une feuille, à partir de la ligne 6, avec la mise en forme du tribunal."""
    nb_colonnes = len(df.columns)
    ws.merge_cells(start_row=1, start_column=1, end_row=LIGNE_ENTETE - 1, end_column=nb_colonnes)
    ws.cell(row=1, column=1, value="").alignment = CENTRE
    _ajouter_logo(ws, nb_colonnes)

    if str(df.columns[0]).lower() in ["index", "unnamed: 0"]:
        df = df.rename(columns={df.columns[0]: "السنة"})

    for num_col, nom in enumerate(df.columns, 1):
        cell = ws.cell(row=LIGNE_ENTETE, column=num_col, value=str(nom))
        cell.font = Font(bold=True, color="000000")
        cell.fill = FOND_ENTETE
        cell.alignment = CENTRE
        cell.border = BORDURE

    a_colonne_rmz = "الرمز" in df.columns
    for num_ligne, ligne in enumerate(df.itertuples(index=False), start=LIGNE_ENTETE + 1):
        valeurs = list(ligne)
        if a_colonne_rmz and valeurs[list(df.columns).index("الرمز")] == "المجموع":
            # Tableau des statistiques générales : « المجموع » fusionné sur les deux premières colonnes
            ws.merge_cells(start_row=num_ligne, start_column=1, end_row=num_ligne, end_column=2)
            cell = ws.cell(row=num_ligne, column=1, value="المجموع")
            cell.fill, cell.font, cell.alignment, cell.border = FOND_TOTAL, Font(bold=True), CENTRE, BORDURE
            ws.cell(row=num_ligne, column=2).border = BORDURE
            debut = 3
        else:
            debut = 1
        ligne_total = str(valeurs[0]).strip() == "المجموع"
        for num_col in range(debut, nb_colonnes + 1):
            cell = ws.cell(row=num_ligne, column=num_col, value=valeurs[num_col - 1])
            cell.alignment = CENTRE
            cell.border = BORDURE
            if ligne_total:
                cell.fill = FOND_TOTAL
                cell.font = Font(bold=True)

    for num_col, cellules in enumerate(ws.columns, start=1):
        if all(isinstance(c, MergedCell) for c in cellules):
            continue
        largeur = max(len(str(c.value)) if c.value is not None else 0 for c in cellules)
        ws.column_dimensions[get_column_letter(num_col)].width = largeur + 2


def enregistrer_classeur(fichier, feuilles, figure=None, titre_figure="الرسم البياني"):
    """Crée un classeur avec une feuille par (titre, DataFrame) et, si besoin, une feuille avec le graphique."""
    wb = Workbook()
    wb.remove(wb.active)
    for titre, df in feuilles:
        ws = wb.create_sheet(title=titre)
        if not df.empty:
            ecrire_feuille(ws, df)

    chemin_image = None
    if figure is not None:
        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
            chemin_image = tmp.name
        figure.savefig(chemin_image)
        wb.create_sheet(title=titre_figure).add_image(ImageExcel(chemin_image), "B2")
    try:
        wb.save(fichier)
    finally:
        if chemin_image:
            os.remove(chemin_image)
