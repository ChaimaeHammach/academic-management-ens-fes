

import re
import copy
import os
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


def _detecter_named_ranges(ws_src) -> set[tuple]:
    """
    Détecte les cellules (row, col) dont les formules utilisent des named ranges
    définis dans le workbook source.
    Ces cellules seront remplacées par leur valeur calculée dans la copie.
    """
    import re as _re
    cellules = set()
    try:
        wb = ws_src.parent
        # Récupérer tous les noms définis dans le workbook (named ranges)
        # L'API openpyxl: wb.defined_names est un DefinedNameDict (dict-like), clés = noms
        defined_names = set()
        if hasattr(wb, 'defined_names'):
            for name in wb.defined_names.keys():
                defined_names.add(str(name).upper())
        if not defined_names:
            return cellules
        # Scanner les cellules avec formules
        for row in ws_src.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith('='):
                    formula_upper = cell.value.upper()
                    for name in defined_names:
                        # Vérifier que le nom est utilisé comme identifiant (pas sous-chaîne)
                        if _re.search(r'\b' + _re.escape(name) + r'\b', formula_upper):
                            cellules.add((cell.row, cell.column))
                            break
    except Exception:
        pass
    return cellules


def copier_fiche_vers_wb(chemin: str, wb_dest, nom_feuille: str) -> None:
    """
    Copie une fiche Excel complète (avec formules + styles) dans wb_dest
    sous le nom nom_feuille.

    Correction: les formules utilisant des named ranges (ex: CC, EXAM)
    sont remplacées par leur valeur calculée (data_only=True) pour éviter
    les erreurs #NOM? dans le workbook destination où ces noms ne sont pas définis.
    """
    # Charger avec formules (pour les styles et la valeur littérale des cellules)
    wb_src = load_workbook(chemin, data_only=False)
    ws_src = wb_src.active

    # Charger avec valeurs calculées (pour remplacer les formules problématiques)
    wb_val = load_workbook(chemin, data_only=True)
    ws_val = wb_val.active

    # Détecter les cellules dont les formules utilisent des named ranges
    cellules_named_range = _detecter_named_ranges(ws_src)

    # Éviter les noms de feuilles en double
    existing = [s.title for s in wb_dest.worksheets]
    sheet_name = nom_feuille[:31]
    if sheet_name in existing:
        sheet_name = sheet_name[:28] + f"_{len(existing)}"

    ws_dst = wb_dest.create_sheet(title=sheet_name)

    # ── Dimensions colonnes et lignes ─────────────────────────────────────
    for col_letter, col_dim in ws_src.column_dimensions.items():
        ws_dst.column_dimensions[col_letter].width = col_dim.width or 12
    for row_idx, row_dim in ws_src.row_dimensions.items():
        if row_dim.height:
            ws_dst.row_dimensions[row_idx].height = row_dim.height

    # ── Cellules fusionnées ────────────────────────────────────────────────
    for merged in ws_src.merged_cells.ranges:
        try:
            ws_dst.merge_cells(
                start_row=merged.min_row, start_column=merged.min_col,
                end_row=merged.max_row,   end_column=merged.max_col
            )
        except Exception:
            pass

    # ── Copier toutes les cellules ────────────────────────────────────────
    for row in ws_src.iter_rows():
        for src_cell in row:
            coord = (src_cell.row, src_cell.column)
            # Si la cellule utilise un named range → écrire la valeur calculée
            if coord in cellules_named_range:
                val = ws_val.cell(src_cell.row, src_cell.column).value
            else:
                val = src_cell.value
            dst_cell = ws_dst.cell(
                row=src_cell.row,
                column=src_cell.column,
                value=val
            )
            _copier_style(src_cell, dst_cell)
        ws_dst.row_dimensions[src_cell.row].height = (
            ws_src.row_dimensions[src_cell.row].height or 15
        )

    # Vue RTL si nécessaire
    if ws_src.sheet_view.rightToLeft:
        ws_dst.sheet_view.rightToLeft = True

    wb_val.close()


def _copier_style(src, dst):
    """Copie le style complet d'une cellule source vers destination."""
    try:
        if src.font:
            dst.font = copy.copy(src.font)
    except Exception:
        pass
    try:
        if src.fill and src.fill.fill_type not in (None, "none"):
            dst.fill = copy.copy(src.fill)
    except Exception:
        pass
    try:
        if src.alignment:
            dst.alignment = copy.copy(src.alignment)
    except Exception:
        pass
    try:
        if src.border:
            dst.border = copy.copy(src.border)
    except Exception:
        pass
    try:
        if src.number_format:
            dst.number_format = src.number_format
    except Exception:
        pass


def ajuster_formule(formule: str, offset: int) -> str:
    """
    Décale les références de lignes relatives dans une formule.
    Ex : =IF(F8>0, ...) avec offset=+2 → =IF(F10>0, ...)
    Les références absolues ($F$8) ne sont PAS modifiées.
    """
    if offset == 0 or not formule.startswith("="):
        return formule

    def replacer(match):
        col_part = match.group(1)
        row_part = match.group(2)
        # Référence absolue $COL$ROW → ne pas toucher
        if "$" in col_part and match.group(0).count("$") >= 2:
            return match.group(0)
        new_row = int(row_part) + offset
        return f"{col_part}{new_row}"

    return re.sub(r"(\$?[A-Za-z]{1,3})(\d+)", replacer, formule)


# ─── Injection depuis PV précédent ────────────────────────────────────────────

def lire_lignes_modules_depuis_pv_prec(chemin_pv_prec: str) -> dict:
    """
    Lit un fichier PV précédent (généré par cet outil) et extrait,
    pour chaque feuille module, les lignes brutes de chaque étudiant.

    Retourne :
    {
      nom_module_normalisé: {
        "ws":              ws_src (openpyxl worksheet, data_only=False),
        "wb":              wb_src,
        "hdr_row":         int,          # ligne d'en-tête (1-based)
        "col_cin":         int | None,
        "col_massar":      int | None,
        "col_note_fin":    int | None,   # Note Finale (1-based)
        "col_dec_fin":     int | None,   # Décision Finale (1-based)
        "etudiants":       { CLE_upper: row_int },   # CLE → numéro de ligne Excel
      },
      ...
    }
    """
    import re as _re

    if not chemin_pv_prec or not os.path.exists(chemin_pv_prec):
        return {}

    try:
        wb = load_workbook(chemin_pv_prec, data_only=False)
    except Exception as e:
        print(f"  ⚠ Impossible de lire le PV précédent pour copie fiche : {e}")
        return {}

    result = {}

    for sname in wb.sheetnames:
        # Ignorer les feuilles PV (on veut les fiches modules)
        if sname.upper().startswith("PV"):
            continue

        ws = wb[sname]
        max_row = ws.max_row
        max_col = ws.max_column

        # Trouver la ligne d'en-tête (contient CIN ou Massar)
        hdr_row = None
        col_cin = col_massar = col_note_fin = col_dec_fin = col_cc = col_exam = col_annee = None

        for r in range(1, min(10, max_row + 1)):
            has_cin = has_mas = False
            col_annee = None 
            for c in range(1, min(20, max_col + 1)):
                v_raw = str(ws.cell(r, c).value or "").strip()
                v = v_raw.lower()
                if v == "cin" or "بطاقة" in v_raw:
                    has_cin = True
                    col_cin = c
                if any(k in v for k in ("massar", "cne")) or "مسار" in v_raw or ("وطني" in v_raw and "بطاقة" not in v_raw):
                    has_mas = True
                    col_massar = c
            if has_cin or has_mas:
                hdr_row = r
                break

        if hdr_row is None:
            continue

        # Identifier toutes les colonnes utiles
        # Identifier toutes les colonnes utiles
        for c in range(1, max_col + 1):
            v_raw = str(ws.cell(hdr_row, c).value or "").strip()
            v = v_raw.lower()
            if "note finale" in v or "note final" in v or "علامة" in v_raw or "النقطة النهائية" in v_raw:
                col_note_fin = c
            elif v_raw == "سنة" or "annee_val" in v or (v in ("av",)):
                col_annee = c
            elif "c.c" in v or (v.startswith("cc") and "%" in str(ws.cell(hdr_row, c).value or "")) or "مراقبة" in v_raw or "مستمرة" in v_raw:
                if col_cc is None:
                    col_cc = c
            elif "exm" in v or ("exam" in v and "%" in str(ws.cell(hdr_row, c).value or "")) or "امتحان" in v_raw:
                col_exam = c
            elif "décision" in v or "decision" in v or "قرار" in v_raw:
                if "finale" in v or "final" in v or "النهائي" in v_raw:
                    col_dec_fin = c
                elif col_dec_fin is None:
                    col_dec_fin = c
            elif "rattrapage" in v or "rattr" in v or "استدراك" in v_raw:
                pass  # on garde pour référence future si besoin

        # Fallback : si note_fin absente, chercher "note finale" ou "moyenne" en fin de ligne
        if col_note_fin is None:
            # Chercher la dernière colonne numérique significative avant les décisions
            for c in range(max_col, 0, -1):
                v_raw2 = str(ws.cell(hdr_row, c).value or "").strip()
                v = v_raw2.lower()
                if any(k in v for k in ("note finale", "note final", "note\nfinale", "moy")) or "علامة" in v_raw2 or "معدل" in v_raw2:
                    col_note_fin = c
                    break

        cle_col = col_cin if col_cin else col_massar
        if cle_col is None:
            continue

        # Construire l'index CLE → ligne
        # Règle : UNE seule entrée par étudiant pour éviter les doubles injections.
        # Clé primaire = Massar/CNE (identifiant unique par étudiant).
        # Si Massar absent, on utilise le CIN comme clé de secours.
        # cin_par_massar permet la résolution croisée lors de la comparaison.
        etudiants = {}        # { massar_ou_cin: row }
        cin_par_massar = {}   # { massar: cin } et { cin: massar }
        INVALIDES = ("", "NONE", "NAN", "CIN", "MASSAR", "CNE")
        for r in range(hdr_row + 1, max_row + 1):
            v_cin    = str(ws.cell(r, col_cin).value    or "").strip().upper() if col_cin    else ""
            v_massar = str(ws.cell(r, col_massar).value or "").strip().upper() if col_massar else ""
            cin_ok    = v_cin    and v_cin    not in INVALIDES
            massar_ok = v_massar and v_massar not in INVALIDES
            if cin_ok and massar_ok:
                cin_par_massar[v_massar] = v_cin
                cin_par_massar[v_cin]    = v_massar
            # Clé principale = Massar (unique par étudiant) ; CIN seulement si Massar absent
            if massar_ok:
                etudiants[v_massar] = r
            elif cin_ok:
                etudiants[v_cin] = r

        nom_norm = _re.sub(r"[\s_\-\.]+", "", sname.lower())
        result[nom_norm] = {
            "sname":          sname,
            "wb":             wb,
            "ws":             ws,
            "hdr_row":        hdr_row,
            "col_cin":        col_cin,
            "col_massar":     col_massar,
            "col_note_fin":   col_note_fin,
            "col_dec_fin":    col_dec_fin,
            "col_cc":         col_cc,
            "col_exam":       col_exam,
            "col_annee":      col_annee,
            "max_col":        max_col,
            "etudiants":      etudiants,
            "cin_par_massar": cin_par_massar,
        }

    print(f"  PV précédent (fiches) : {len(result)} modules trouvés → "
          f"{[v['sname'] for v in result.values()]}")
    return result


def injecter_etudiants_depuis_pv_prec(
    ws_dest,
    info_module_prec: dict,
    cles_manquantes: list[str],
    wb_dest_val=None,
) -> dict[str, int]:
    """
    Pour chaque CLE manquante dans ws_dest (feuille module du workbook courant),
    copie toute la ligne correspondante depuis la feuille module du PV précédent.

    Retourne { CLE: nouvelle_ligne_excel } pour mettre à jour liaison_modules.
    Préserve les styles de la ligne modèle (dernière ligne de données du dest).
    """
    if not cles_manquantes or not info_module_prec:
        return {}

    ws_src   = info_module_prec["ws"]
    src_etuds = info_module_prec["etudiants"]
    src_max_col = info_module_prec["max_col"]
    dst_max_col = ws_dest.max_column

    nouvelles_lignes: dict[str, int] = {}

    # Ligne modèle pour le style : dernière ligne de données du dest
    # (la dernière ligne non vide après l'en-tête)
    ligne_modele = ws_dest.max_row
    while ligne_modele > 1 and all(
        ws_dest.cell(ligne_modele, c).value is None
        for c in range(1, min(6, dst_max_col + 1))
    ):
        ligne_modele -= 1

    # Dédupliquer : ne pas injecter deux fois la même ligne source
    # (peut arriver si CIN et Massar d'un même étudiant passent tous les deux)
    lignes_src_deja_injectees: set[int] = set()

    for cle in cles_manquantes:
        src_row = src_etuds.get(cle)
        if src_row is None:
            continue

        # Déduplication : si cette ligne source a déjà été injectée, on saute
        if src_row in lignes_src_deja_injectees:
            print(f"    ⚠ CLE={cle} : même ligne source déjà injectée — ignoré")
            continue
        lignes_src_deja_injectees.add(src_row)

        new_row = ws_dest.max_row + 1

        # Copier chaque cellule de la ligne source → dest
        for c in range(1, src_max_col + 1):
            src_cell = ws_src.cell(src_row, c)
            dst_cell = ws_dest.cell(new_row, c)

            # Valeur : copier telle quelle (formule ou valeur)
            val = src_cell.value
            # Si c'est une formule, la garder telle quelle (elle pointe vers
            # les colonnes de la même feuille, donc c'est valide dans le dest)
            dst_cell.value = val

            # Style
            _copier_style(src_cell, dst_cell)

        # Hauteur de ligne
        ws_dest.row_dimensions[new_row].height = (
            ws_src.row_dimensions.get(src_row, None) and
            ws_src.row_dimensions[src_row].height
        ) or 15

        nouvelles_lignes[cle] = new_row
        print(f"    ✓ Ligne injectée depuis PV préc. : CLE={cle} → row {new_row}")

    return nouvelles_lignes


def _module_norm(nom: str) -> str:
    import re as _re
    return _re.sub(r"[\s_\-\.]+", "", nom.lower())
