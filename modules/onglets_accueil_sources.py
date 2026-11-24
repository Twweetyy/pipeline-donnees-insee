"""Ordre et noms des premiers onglets du classeur.

- « Accueil » (ex-Portrait Roubaix / Portrait MEL) : chiffres clés, en premier ;
- « Sources » (ex-Présentation) : millésimes, liens, précisions, puis le détail des calculs
  (ex-onglet « Detail calcul », supprimé), en deuxième ;
- les onglets thématiques suivent, dans leur ordre habituel.

Pour Roubaix, Sources porte aussi, à droite (colonnes G à I), la correspondance
grand quartier -> IRIS (Code IRIS, libellé, grand quartier).

Étape finale du pipeline : les modules précédents travaillent encore sur les noms du modèle Excel.
"""
from copy import copy

from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

NOM_ACCUEIL = "Accueil"
NOM_SOURCES = "Sources"
ANCIEN_PRESENTATION = "Présentation"
ANCIEN_DETAIL = "Detail calcul"
SEPARATEUR_FORMULE = " -> "
MENTION_NON_CALCULE = "(non calculé"


def _lire_detail_calcul(ws) -> list[tuple[str, list[tuple[str, str, bool]]]]:
    """[(en-tête du thème, [(indicateur, formule, calculé), ...]), ...] depuis l'onglet Detail calcul."""
    themes = []
    for colonne in range(1, ws.max_column + 1):
        entete = ws.cell(row=2, column=colonne).value
        if not entete:
            continue
        indicateurs = []
        for ligne in range(3, ws.max_row + 1):
            texte = ws.cell(row=ligne, column=colonne).value
            if not texte:
                continue
            texte = str(texte)
            calcule = MENTION_NON_CALCULE not in texte
            texte = texte.split("\n" + MENTION_NON_CALCULE)[0].strip()
            indicateur, _, formule = texte.partition(SEPARATEUR_FORMULE)
            indicateurs.append((indicateur.strip(), formule.strip(), calcule))
        if indicateurs:
            themes.append((str(entete).strip(), indicateurs))
    return themes


def _copier_style(source, cible) -> None:
    cible.font, cible.fill, cible.border = copy(source.font), copy(source.fill), copy(source.border)
    cible.alignment = copy(source.alignment)


def _largeur(ws, lettres: str) -> float:
    return sum(ws.column_dimensions[lettre].width or 10 for lettre in lettres)


def _hauteur(textes: list[tuple[str, float]], taille: float) -> float:
    """Hauteur de ligne estimée pour des textes renvoyés à la ligne (Excel ne la recalcule pas)."""
    lignes = 1
    for texte, largeur in textes:
        capacite = max(1, int(largeur * 1.1 * 10 / taille))
        lignes = max(lignes, sum(-(-max(len(morceau), 1) // capacite) for morceau in texte.split("\n")))
    return min(409, lignes * taille * 1.35 + 6)


def _derniere_ligne_remplie(ws) -> int:
    """Dernière ligne remplie de la page (colonnes A à E : le tableau des grands quartiers est à droite)."""
    derniere = ws.max_row
    while derniere > 1 and all(c.value in (None, "") for c in ws[derniere][:5]):
        derniere -= 1
    return derniere


def table_grands_quartiers(dico_fichiers: dict) -> list[tuple[str, str, str]]:
    """[(code IRIS, libellé, grand quartier), ...] triés par code, depuis les colonnes de Périmètres_MAS.csv
    déjà jointes aux données (mêmes valeurs que la colonne « Grands quartiers » des onglets)."""
    lignes = {}
    for df in dico_fichiers.values():
        colonnes = getattr(df, "columns", [])
        if not {"IRIS", "Grands quartiers"} <= set(colonnes):
            continue
        libelles = df["Libellé de l'Iris"] if "Libellé de l'Iris" in colonnes else [None] * len(df)
        for code, libelle, quartier in zip(df["IRIS"], libelles, df["Grands quartiers"]):
            # IRIS non localisé (59512) ou hors référentiel : pas de grand quartier
            if not isinstance(quartier, str) or not quartier.strip():
                continue
            code = str(code).strip()
            lignes.setdefault(code, (code, libelle if isinstance(libelle, str) else "", quartier.strip()))
    return [lignes[code] for code in sorted(lignes)]


def _ajouter_grands_quartiers(ws_sources, table) -> None:
    """Tableau Code IRIS / IRIS / Grand quartier à droite de la page Sources (colonnes G à I).

    Il reprend la charte de la page : bandeau titre comme A1, en-têtes comme la ligne 3,
    lignes comme le tableau des onglets. Un grand quartier à plusieurs IRIS est fusionné
    verticalement, et les grands quartiers alternent fond bleu clair / fond blanc.
    """
    premiere = 7  # colonne G ; F sert de marge
    colonnes = range(premiere, premiere + 3)
    style_titre, style_entete, style_ligne = ws_sources["A1"], ws_sources["D3"], ws_sources["D4"]
    police = style_entete.font.name or "Arial"
    bordure = copy(style_ligne.border)
    sans_fond = PatternFill(fill_type=None)

    ws_sources.column_dimensions[get_column_letter(premiere - 1)].width = 3
    for colonne, largeur in zip(colonnes, (13, 24, 36)):
        ws_sources.column_dimensions[get_column_letter(colonne)].width = largeur

    # Bandeau titre (ligne 1, comme « Base MAS - INSEE ... »)
    ws_sources.merge_cells(start_row=1, start_column=premiere, end_row=1, end_column=premiere + 2)
    titre = ws_sources.cell(row=1, column=premiere, value="Grands quartiers et IRIS")
    for colonne in colonnes:
        _copier_style(style_titre, ws_sources.cell(row=1, column=colonne))
    titre.alignment = Alignment(horizontal="center", vertical="center")

    # En-têtes (ligne 3, alignés sur « Onget disponible / Annee / Description »)
    for colonne, texte in zip(colonnes, ("Code IRIS", "IRIS", "Grand quartier")):
        cellule = ws_sources.cell(row=3, column=colonne, value=texte)
        _copier_style(style_entete, cellule)
        cellule.border = bordure
        cellule.alignment = Alignment(horizontal="center", vertical="center")

    # Une ligne par IRIS ; groupes de grands quartiers consécutifs (ordre des codes)
    groupes = []
    for code, libelle, quartier in table:
        if groupes and groupes[-1][0] == quartier:
            groupes[-1][1].append((code, libelle))
        else:
            groupes.append((quartier, [(code, libelle)]))

    ligne = 4
    for rang, (quartier, iris) in enumerate(groupes):
        fond = copy(style_ligne.fill) if rang % 2 == 0 else sans_fond
        debut = ligne
        for code, libelle in iris:
            # Lignes séparatrices très basses de la partie gauche : remontées pour rester lisibles
            if (ws_sources.row_dimensions[ligne].height or 15) < 15:
                ws_sources.row_dimensions[ligne].height = 15
            for colonne, valeur in zip(colonnes, (code, libelle, quartier if ligne == debut else None)):
                cellule = ws_sources.cell(row=ligne, column=colonne, value=valeur)
                cellule.font = Font(name=police, size=10, color=copy(style_ligne.font.color),
                                    bold=colonne != premiere + 1)
                cellule.fill = copy(fond)
                cellule.border = bordure
                cellule.alignment = Alignment(
                    horizontal="center" if colonne == premiere else "left",
                    vertical="center", wrap_text=True, indent=0 if colonne == premiere else 1,
                )
            ws_sources.cell(row=ligne, column=premiere).number_format = "@"
            ligne += 1
        if ligne - debut > 1:
            ws_sources.merge_cells(start_row=debut, start_column=premiere + 2, end_row=ligne - 1, end_column=premiere + 2)

    nb_quartiers = len({quartier for _, _, quartier in table})
    note = ws_sources.cell(
        row=ligne + 1, column=premiere,
        value=f"{nb_quartiers} grands quartiers, {len(table)} IRIS. Mêmes noms que la colonne "
              "« Grands quartiers » des onglets (fichier Périmètres_MAS.csv).",
    )
    ws_sources.merge_cells(start_row=ligne + 1, start_column=premiere, end_row=ligne + 2, end_column=premiere + 2)
    note.font = Font(name=police, italic=True, size=9, color="595959")
    note.alignment = Alignment(wrap_text=True, vertical="top")

    if ws_sources.print_area:
        bas = max(ws_sources.max_row, ligne + 2)
        ws_sources.print_area = f"A1:{get_column_letter(premiere + 2)}{bas}"


def _ajouter_detail_calcul(ws_sources, themes) -> None:
    """Ajoute le détail des calculs sous le contenu de la page Sources (colonnes A à E)."""
    derniere = _derniere_ligne_remplie(ws_sources)
    # Styles repris de la page : bandeau « Précision » et sous-titres « # ... »
    style_bandeau = next(
        (c for c in ws_sources["A"] if isinstance(c.value, str) and c.value.strip().startswith("Précision")),
        ws_sources["A3"],
    )
    style_sous_titre = next(
        (c for c in ws_sources["A"] if isinstance(c.value, str) and c.value.startswith("#")),
        style_bandeau,
    )
    police = style_sous_titre.font.name or "Arial"
    largeur_indicateur = _largeur(ws_sources, "ABC")
    largeur_formule = _largeur(ws_sources, "DE")

    ligne = derniere + 2
    ws_sources.merge_cells(start_row=ligne, start_column=1, end_row=ligne, end_column=5)
    bandeau = ws_sources.cell(row=ligne, column=1, value="Détail des calculs")
    _copier_style(style_bandeau, bandeau)
    bandeau.alignment = Alignment(horizontal="center", vertical="center")
    ws_sources.row_dimensions[ligne].height = 18
    ligne += 1
    note = ws_sources.cell(
        row=ligne, column=1,
        value="Indicateurs calculés ajoutés à la fin des onglets (tout à droite, fond orange) : "
              "nom de la colonne, puis formule appliquée à chaque ligne.",
    )
    ws_sources.merge_cells(start_row=ligne, start_column=1, end_row=ligne, end_column=5)
    note.font = Font(name=police, italic=True, size=9, color="595959")
    note.alignment = Alignment(wrap_text=True, vertical="center")
    ws_sources.row_dimensions[ligne].height = 18
    ligne += 1

    for entete, indicateurs in themes:
        ligne += 1
        ws_sources.merge_cells(start_row=ligne, start_column=1, end_row=ligne, end_column=5)
        sous_titre = ws_sources.cell(row=ligne, column=1, value="# " + entete.replace("Details", "Détails", 1))
        _copier_style(style_sous_titre, sous_titre)
        ws_sources.row_dimensions[ligne].height = 16
        ligne += 1
        for indicateur, formule, calcule in indicateurs:
            ws_sources.merge_cells(start_row=ligne, start_column=1, end_row=ligne, end_column=3)
            ws_sources.merge_cells(start_row=ligne, start_column=4, end_row=ligne, end_column=5)
            if not calcule:
                formule += "\n(non calculé : variable absente de ce millésime)"
            gauche = ws_sources.cell(row=ligne, column=1, value=indicateur)
            droite = ws_sources.cell(row=ligne, column=4, value=formule)
            couleur = None if calcule else "C00000"
            gauche.font = Font(name=police, bold=True, italic=not calcule, size=9, color=couleur)
            droite.font = Font(name=police, italic=not calcule, size=9, color=couleur)
            gauche.alignment = Alignment(wrap_text=True, vertical="center")
            droite.alignment = Alignment(wrap_text=True, vertical="center")
            ws_sources.row_dimensions[ligne].height = _hauteur(
                [(indicateur, largeur_indicateur), (formule, largeur_formule)], 9
            )
            ligne += 1

    if ws_sources.print_area:
        ws_sources.print_area = f"A1:E{ligne - 1}"


def organiser_onglets_accueil_sources(workbook, grands_quartiers=None) -> None:
    """Accueil (chiffres clés) puis Sources (présentation + détail des calculs) en tête du classeur.

    grands_quartiers : lignes de table_grands_quartiers (Roubaix), ajoutées à Sources avant le détail des calculs.
    """
    portrait = next((ws for ws in workbook.worksheets if ws.title.startswith("Portrait ")), None)
    if portrait is not None:
        portrait.title = NOM_ACCUEIL

    if ANCIEN_PRESENTATION in workbook.sheetnames:
        ws_sources = workbook[ANCIEN_PRESENTATION]
        ws_sources.title = NOM_SOURCES
        if ANCIEN_DETAIL in workbook.sheetnames:
            _ajouter_detail_calcul(ws_sources, _lire_detail_calcul(workbook[ANCIEN_DETAIL]))
        # Après le détail des calculs : la zone d'impression couvre alors toute la page
        if grands_quartiers:
            _ajouter_grands_quartiers(ws_sources, grands_quartiers)
    if ANCIEN_DETAIL in workbook.sheetnames:
        del workbook[ANCIEN_DETAIL]

    for position, nom in enumerate(n for n in (NOM_ACCUEIL, NOM_SOURCES) if n in workbook.sheetnames):
        ws = workbook[nom]
        workbook.move_sheet(ws, offset=position - workbook.sheetnames.index(nom))

    # Le classeur s'ouvre sur l'Accueil
    for ws in workbook.worksheets:
        ws.sheet_view.tabSelected = False
    workbook.active = 0
    workbook.worksheets[0].sheet_view.tabSelected = True
