

import os
import re
import copy
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from modules.reader     import lire_module, lire_pv_precedent
from modules.merger     import construire_tableau, completer_depuis_pv_precedent as _completer_notes
from modules.calculator import construire_pv
from modules.exporter   import exporter_excel as _ecrire_pv_feuille
from modules.copie_fiche import copier_fiche_vers_wb


# ─── Palette ──────────────────────────────────────────────────────────────────
BLEU_DARK = "4A4A4A"
BLEU_MID  = "7A7A7A"
GRIS      = "F5F5F5"
BLANC     = "FFFFFF"
VERT      = "FFFFFF"; VERT_T  = "000000"
ORANGE    = "FFFFFF"; ORANGE_T= "000000"
ROUGE     = "FFFFFF"; ROUGE_T = "000000"


def _f(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def _b():
    s = Side(style="thin", color="AAAAAA")
    return Border(left=s, right=s, top=s, bottom=s)

def _is_na(v):
    if v is None: return True
    try:
        import math; return math.isnan(float(v))
    except: return False

def _safe(v):
    if _is_na(v): return ""
    return v


# ─── Copier toutes les feuilles d'un fichier Excel dans un workbook ───────────

def _copier_toutes_feuilles(chemin_source: str, wb_dest: Workbook) -> None:
    """Copie toutes les feuilles du fichier source dans wb_dest (conserve les formules)."""
    _copier_toutes_feuilles_avec_noms(chemin_source, wb_dest, label_pv=None)


def _copier_toutes_feuilles_avec_noms(
    chemin_source: str,
    wb_dest:       Workbook,
    label_pv:      str | None = None,
) -> str | None:
    """
    Copie toutes les feuilles du fichier source dans wb_dest.
    Conserve les formules (data_only=False).
    Gère les doublons de noms en suffixant _2, _3, etc.

    Si label_pv est fourni, retourne le nom final attribué à la feuille
    dont le titre commence par "PV {label_pv}" (ou n'importe quelle feuille
    "PV ..." si aucune ne correspond exactement) dans wb_dest.
    Retourne None si aucune feuille PV n'est trouvée ou label_pv est None.
    """
    if not chemin_source or not os.path.exists(chemin_source):
        return None

    wb_src   = load_workbook(chemin_source, data_only=False)
    existing = {s.title for s in wb_dest.worksheets}
    nom_pv_dest: str | None = None

    for ws_src in wb_src.worksheets:
        nom = ws_src.title[:31]
        candidat, cpt = nom, 2
        while candidat in existing:
            candidat = nom[:28] + f"_{cpt}"; cpt += 1
        existing.add(candidat)

        ws_dest = wb_dest.create_sheet(title=candidat)

        for row in ws_src.iter_rows():
            for cell in row:
                nc = ws_dest.cell(row=cell.row, column=cell.column, value=cell.value)
                if cell.has_style:
                    nc.font         = copy.copy(cell.font)
                    nc.fill         = copy.copy(cell.fill)
                    nc.border       = copy.copy(cell.border)
                    nc.alignment    = copy.copy(cell.alignment)
                    nc.number_format = cell.number_format

        for mr in ws_src.merged_cells.ranges:
            ws_dest.merge_cells(str(mr))
        for cl, cd in ws_src.column_dimensions.items():
            ws_dest.column_dimensions[cl].width = cd.width
        for rn, rd in ws_src.row_dimensions.items():
            ws_dest.row_dimensions[rn].height = rd.height
        if ws_src.freeze_panes:
            ws_dest.freeze_panes = ws_src.freeze_panes

        # Identifier la feuille PV dans le fichier source
        if label_pv is not None and nom_pv_dest is None:
            src_title = ws_src.title
            if src_title == f"PV {label_pv}" or src_title.startswith(f"PV {label_pv}"):
                nom_pv_dest = candidat
            elif src_title.startswith("PV "):
                # Fallback : première feuille "PV ..." trouvée
                nom_pv_dest = candidat

    wb_src.close()
    return nom_pv_dest


# ─── Génération PV Semestre ───────────────────────────────────────────────────

def generer_pv_semestre(
    fichiers_modules:   list[tuple[str, str]],
    label_semestre:     str,
    chemin_pv_prec:     str | None,
    dossier_sortie:     str,
    ecole:              str,
    filiere:            str,
    annee:              str,
    langue:             str = "fr",
    chemin_pv_sem_prec: str | None = None,
) -> str:
    """
    PV Semestre : fiches originales + feuille PV SANS colonnes de synthèse finales.

    chemin_pv_sem_prec : PV du même semestre de l'année précédente.
      → Pour chaque étudiant absent des fiches modules courants,
        sa ligne est copiée depuis la feuille module correspondante du PV précédent
        directement dans la feuille module du workbook généré.
      → Le PV contient ainsi toutes les notes disponibles avec formules liées.
    """
    os.makedirs(dossier_sortie, exist_ok=True)

    modules_lus = []
    for nom, chemin in fichiers_modules:
        info = lire_module(chemin, nom)
        info["_chemin"] = chemin
        modules_lus.append(info)

    noms_modules = [m["nom_module"] for m in modules_lus]

    # ── Charger l'index des lignes depuis le PV semestre précédent ─────────
    from modules.copie_fiche import (
        lire_lignes_modules_depuis_pv_prec,
        injecter_etudiants_depuis_pv_prec,
        _module_norm,
    )
    index_pv_prec: dict = {}
    if chemin_pv_sem_prec and os.path.exists(chemin_pv_sem_prec):
        print("\n  Chargement PV semestre précédent pour injection…")
        index_pv_prec = lire_lignes_modules_depuis_pv_prec(chemin_pv_sem_prec)

    # ── Construire la table de liaison : module → {cle → ligne_excel} ──────
    liaison_modules: dict[str, dict] = {}
    for info in modules_lus:
        nom = info["nom_module"]
        df_m = info["df"]
        ligne_debut = info.get("xlsx_ligne_debut")
        col_note    = info.get("xlsx_col_note")
        if ligne_debut is None or col_note is None:
            continue
        cle_to_row = {}
        for df_idx, row in df_m.iterrows():
            ligne_excel = ligne_debut + df_idx
            for champ in ("CLE", "Massar", "CIN"):
                val = str(row.get(champ, "") or "").strip().upper()
                if val and val not in ("NAN", "NONE", ""):
                    cle_to_row.setdefault(val, ligne_excel)
        liaison_modules[nom] = {
            "cle_to_row": cle_to_row,
            "col_note":   col_note,
            "col_dec":    info.get("xlsx_col_dec"),
            "col_annee":  info.get("xlsx_col_annee"),
            "col_cin":    info.get("xlsx_col_cin",    1),
            "col_massar": info.get("xlsx_col_massar", 2),
            "col_nom":    info.get("xlsx_col_nom",    3),
            "col_prenom": info.get("xlsx_col_prenom", 4),
        }

    # ── Compléter df depuis le PV précédent (pour les notes manquantes) ────
    df_pv_precedent_sem = None
    if index_pv_prec:
        # Construire un df compatible avec completer_depuis_pv_precedent
        # en lisant les notes via l'index qu'on vient de charger
        rows_prec = []
        for norm_mod, info_mod in index_pv_prec.items():
            ws_p = info_mod["ws"]
            col_nf = info_mod["col_note_fin"]
            for cle, row_p in info_mod["etudiants"].items():
                if col_nf:
                    note_raw = ws_p.cell(row_p, col_nf).value
                    import pandas as _pd
                    note = None
                    try:
                        note = float(note_raw)
                    except (TypeError, ValueError):
                        pass
                    rows_prec.append({
                        "CLE":    cle,
                        "Massar": str(ws_p.cell(row_p, info_mod["col_massar"]).value or "").strip() if info_mod["col_massar"] else "",
                        "CIN":    str(ws_p.cell(row_p, info_mod["col_cin"]).value    or "").strip() if info_mod["col_cin"]    else "",
                        "Nom":    "",
                        "Prénom": "",
                        "Module": info_mod["sname"],
                        "Note":   note,
                    })
        if rows_prec:
            import pandas as _pd
            df_pv_precedent_sem = _pd.DataFrame(rows_prec)

    # PV précédent de semestre utilisé pour compléter les notes manquantes
    df, _, _ = construire_tableau(
        modules_si=modules_lus, modules_sp=[],
        df_pv_precedent=df_pv_precedent_sem,
    )

    # df_pv sera construit ici si index_pv_prec est absent,
    # sinon il est déjà construit dans le bloc injection ci-dessus
    coeffs = {m: 1.0 for m in noms_modules}
    if not index_pv_prec:
        df_pv = construire_pv(
            df=df, modules_si=noms_modules, modules_sp=[],
            coefficients=coeffs, label_si=label_semestre, label_sp="",
        )

    wb = Workbook()
    wb.remove(wb.active)

    for info in modules_lus:
        copier_fiche_vers_wb(info["_chemin"], wb, info["nom_module"][:31])

    # ── Injecter les étudiants manquants depuis le PV semestre précédent ──
    # Pour chaque module, on cherche quels étudiants sont dans le PV précédent
    # mais absents du fichier courant. On copie leur ligne dans la feuille module
    # du workbook courant et on met à jour la liaison (formules dynamiques).
    if index_pv_prec:
        print("\n  Injection des étudiants manquants depuis PV semestre précédent…")
        # ── Construire le mapping global Massar/CIN de tous les modules courants ──
        # Le Massar est l'identifiant UNIQUE d'un étudiant (le CIN peut être partagé).
        # Règle : un étudiant du PV précédent est un REDOUBLANT pour le module X si :
        #   1. Son Massar est présent dans AU MOINS UN module courant (il est inscrit cette année)
        #   2. Son Massar est ABSENT du module X courant (sa note manque pour ce module)
        massars_presents_global: set[str] = set()
        cin_to_massars_global:   dict     = {}   # cin → set(massars) pour résolution croisée
        for info in modules_lus:
            for _, row in info["df"].iterrows():
                mas = str(row.get("Massar", "") or "").strip().upper()
                cin = str(row.get("CIN",    "") or "").strip().upper()
                if mas and mas not in ("", "NONE", "NAN"):
                    massars_presents_global.add(mas)
                if cin and mas and cin not in ("", "NONE", "NAN"):
                    cin_to_massars_global.setdefault(cin, set()).add(mas)

        for info in modules_lus:
            nom_mod = info["nom_module"]
            norm    = _module_norm(nom_mod)

            # Trouver la fiche correspondante dans le PV précédent
            info_prec = index_pv_prec.get(norm)
            if info_prec is None:
                # Essai de correspondance partielle (8 premiers caractères)
                for k, v in index_pv_prec.items():
                    if len(norm) >= 6 and (norm[:8] in k or k[:8] in norm):
                        info_prec = v
                        break
            if info_prec is None:
                print(f"    ⚠ Module '{nom_mod}' absent du PV précédent — ignoré")
                continue

            # Identifiants présents dans CE module courant (Massar + CIN + CLE)
            ids_module_courant: set[str] = set()
            for _, row_df in info["df"].iterrows():
                for champ in ("Massar", "CIN", "CLE"):
                    val = str(row_df.get(champ, "") or "").strip().upper()
                    if val and val not in ("", "NONE", "NAN"):
                        ids_module_courant.add(val)

            cpm = info_prec.get("cin_par_massar", {})

            # Détection des vrais redoublants :
            # - leur Massar (clé dans etudiants du PV prec) est dans massars_presents_global
            #   (= ils sont inscrits cette année dans au moins un module)
            # - leur Massar est ABSENT de ce module courant
            cles_manquantes = []
            for massar_prec in sorted(info_prec["etudiants"].keys()):
                # Critère 1 : étudiant inscrit cette année
                cin_prec = cpm.get(massar_prec, "")
                in_this_year = massar_prec in massars_presents_global
                if not in_this_year and cin_prec:
                    # Résolution croisée : CIN → Massars cette année
                    massars_this_cin = cin_to_massars_global.get(cin_prec, set())
                    in_this_year = massar_prec in massars_this_cin
                if not in_this_year:
                    continue  # pas inscrit cette année → ignorer

                # Critère 2 : absent de CE module (vérifier Massar + CIN + alias)
                cin_prec_alias = cpm.get(massar_prec, "")
                if (massar_prec in ids_module_courant
                        or (cin_prec and cin_prec in ids_module_courant)
                        or (cin_prec_alias and cin_prec_alias in ids_module_courant)):
                    continue  # déjà présent → pas redoublant ici

                cles_manquantes.append(massar_prec)

            if not cles_manquantes:
                print(f"    ✓ Module '{nom_mod}' : aucun redoublant manquant")
                continue

            print(f"    Module '{nom_mod}' : {len(cles_manquantes)} redoublant(s) à injecter : {cles_manquantes}")

            # Trouver la feuille module dans le workbook courant
            ws_mod_dest = None
            sheet_name_dest = nom_mod[:31]
            for ws_candidate in wb.worksheets:
                if ws_candidate.title == sheet_name_dest:
                    ws_mod_dest = ws_candidate
                    break
                # Correspondance partielle (suffixe _PV possible)
                if ws_candidate.title.startswith(sheet_name_dest[:20]):
                    ws_mod_dest = ws_candidate
            if ws_mod_dest is None:
                print(f"    ⚠ Feuille '{sheet_name_dest}' introuvable dans le workbook")
                continue

            # Injecter les lignes manquantes
            nouvelles_lignes = injecter_etudiants_depuis_pv_prec(
                ws_dest=ws_mod_dest,
                info_module_prec=info_prec,
                cles_manquantes=cles_manquantes,
            )

            # Mettre à jour la liaison pour les nouvelles lignes
            # Enregistrer aussi le CIN correspondant pour la résolution dans _formule_note
            if nouvelles_lignes and nom_mod in liaison_modules:
                lien = liaison_modules[nom_mod]
                lien["cle_to_row"].update(nouvelles_lignes)
                if not lien.get("col_annee"):
                    lien["col_annee"] = info_prec.get("col_annee")
                cpm = info_prec.get("cin_par_massar", {})
                for massar_inj, row_inj in nouvelles_lignes.items():
                    cin_inj = cpm.get(massar_inj, "")
                    if cin_inj:
                        lien["cle_to_row"].setdefault(cin_inj, row_inj)

            # Mettre à jour df pour inclure les notes injectées dans le PV
            # (elles seront affichées via formules dans la feuille PV)
            col_nf    = info_prec.get("col_note_fin")
            col_mas   = info_prec.get("col_massar")
            col_cin_p = info_prec.get("col_cin")
            ws_src    = info_prec["ws"]
            cpm       = info_prec.get("cin_par_massar", {})

            for massar_inj, row_src in nouvelles_lignes.items():
                note = None
                if col_nf:
                    try:
                        note = float(ws_src.cell(row_src, col_nf).value or 0)
                    except (TypeError, ValueError):
                        pass
                massar = str(ws_src.cell(row_src, col_mas).value or "").strip() if col_mas else massar_inj
                cin_v  = str(ws_src.cell(row_src, col_cin_p).value or "").strip() if col_cin_p else ""

                # Chercher l'étudiant dans df d'abord par Massar, puis par CIN
                import pandas as _pd2
                mask = _pd2.Series([False] * len(df))
                if massar:
                    mask = mask | (df.get("Massar", _pd2.Series(dtype=str)).str.upper() == massar.upper())
                if cin_v:
                    mask = mask | (df.get("CIN", _pd2.Series(dtype=str)).str.upper() == cin_v.upper())
                # Aussi par CLE
                mask = mask | (df["CLE"].str.upper() == massar_inj.upper())
                cin_alias = cpm.get(massar_inj, "")
                if cin_alias:
                    mask = mask | (df["CLE"].str.upper() == cin_alias.upper())

                if mask.any():
                    # Étudiant déjà dans df (présent dans d'autres modules) → juste remplir la note
                    if note is not None:
                        df.loc[mask, f"{nom_mod}_Note_Finale"] = note
                        df.loc[mask, f"{nom_mod}_Validé"]      = note >= 10
                else:
                    # Vrai redoublant absent de tous les modules courants → ajouter ligne
                    new_row_df: dict = {
                        "CLE":    massar if massar else massar_inj,
                        "Massar": massar,
                        "CIN":    cin_v,
                        "Nom":    str(ws_src.cell(row_src, col_cin_p + 2 if col_cin_p else 4).value or "").strip(),
                        "Prénom": str(ws_src.cell(row_src, col_cin_p + 3 if col_cin_p else 5).value or "").strip(),
                    }
                    for col_df in df.columns:
                        if col_df not in new_row_df:
                            new_row_df[col_df] = None
                    if note is not None:
                        new_row_df[f"{nom_mod}_Note_Finale"] = note
                        new_row_df[f"{nom_mod}_Validé"]      = note >= 10
                    df = _pd2.concat([df, _pd2.DataFrame([new_row_df])], ignore_index=True)

        # Reconstruire df_pv après les injections
        coeffs = {m: 1.0 for m in noms_modules}
        df_pv  = construire_pv(
            df=df, modules_si=noms_modules, modules_sp=[],
            coefficients=coeffs, label_si=label_semestre, label_sp="",
        )

    _ecrire_pv_dans_wb(
        wb=wb, df_pv=df_pv,
        modules_si=noms_modules, modules_sp=[],
        label_si=label_semestre, label_sp="",
        ecole=ecole, filiere=filiere, annee=annee, langue=langue,
        titre_feuille=f"PV {label_semestre}",
        inclure_synthese=False,  # ← PAS de colonnes finales
        liaison_modules=liaison_modules,
    )

    slug = re.sub(r"[^\w\-]", "_", annee) if annee else "annee"
    chemin_out = os.path.join(dossier_sortie, f"PV_{label_semestre}_{slug}.xlsx")
    wb.save(chemin_out)
    print(f"PV Semestre généré : {chemin_out}")

    # Injection des notes redoublants dans le PV Semestre
    if chemin_pv_prec and os.path.exists(chemin_pv_prec):
        try:
            from modules.inject_redoublants import inject_notes_redoublants
            print("\n  Injection des notes redoublants (PV Semestre)...")
            inject_notes_redoublants(
                chemin_xlsx=chemin_out,
                chemin_pv_prec=chemin_pv_prec,
                chemin_sortie=chemin_out,
            )
        except Exception as e:
            import traceback
            print(f"  ⚠ Injection redoublants échouée : {e}")
            print(traceback.format_exc())

    return chemin_out


# ─── Lecture de la feuille PV d'un fichier de semestre ───────────────────────

def _lire_df_depuis_feuille_pv(chemin_pv: str, label: str) -> tuple[pd.DataFrame, list[str]]:
    """
    Lit la feuille "PV {label}" d'un fichier PV de semestre.
    Retourne (DataFrame structuré, liste de noms de modules).

    Structure de la feuille PV :
      Ligne 1 : entête général
      Ligne 2 : noms des modules (cellules fusionnées) + groupes identité/récap/synthèse
      Ligne 3 : sous-colonnes (Moy | AV | Déc pour chaque module, puis récap)
      Ligne 4+ : données étudiants
    """
    if not chemin_pv or not os.path.exists(chemin_pv):
        return pd.DataFrame(), []

    wb = load_workbook(chemin_pv, data_only=True)
    # Ouvrir aussi avec formules pour résoudre les références cross-sheet
    wb_formulas = load_workbook(chemin_pv, data_only=False)

    def _resolve_crosssheet(wb_f, wb_d, sheet_title, row, col):
        """
        Si la cellule (row, col) dans sheet_title contient une formule
        cross-sheet (='Autre'!$E4) et que data_only retourne None,
        va lire la valeur réelle dans la feuille source.
        """
        import re as _re
        from openpyxl.utils import column_index_from_string
        ws_d = wb_d[sheet_title] if sheet_title in wb_d.sheetnames else None
        ws_f = wb_f[sheet_title] if sheet_title in wb_f.sheetnames else None
        if ws_d is None:
            return None
        val = ws_d.cell(row, col).value
        if val is not None:
            return val
        # Valeur None → peut-être une formule non évaluée
        if ws_f is None:
            return None
        formula = ws_f.cell(row, col).value
        if not isinstance(formula, str) or not formula.startswith("="):
            return val
        m = _re.match(r"='?([^!']+)'?![$]?([A-Z]+)(\d+)", formula)
        if not m:
            return val
        src_sheet = m.group(1).replace("''", "'")
        src_col   = column_index_from_string(m.group(2))
        src_row   = int(m.group(3))
        if src_sheet in wb_d.sheetnames:
            src_val = wb_d[src_sheet].cell(src_row, src_col).value
            if src_val is None and src_sheet in wb_f.sheetnames:
                # Récursion un niveau : formule dans la feuille source aussi
                src_formula = wb_f[src_sheet].cell(src_row, src_col).value
                if isinstance(src_formula, str) and src_formula.startswith("="):
                    m2 = _re.match(r"='?([^!']+)'?![$]?([A-Z]+)(\d+)", src_formula)
                    if m2:
                        s2 = m2.group(1).replace("''", "'")
                        c2 = column_index_from_string(m2.group(2))
                        r2 = int(m2.group(3))
                        if s2 in wb_d.sheetnames:
                            src_val = wb_d[s2].cell(r2, c2).value
            return src_val
        return val

    # Chercher la feuille PV du bon semestre
    feuille_pv = None
    for ws in wb.worksheets:
        if ws.title == f"PV {label}" or ws.title.startswith(f"PV {label}"):
            feuille_pv = ws
            break
    if feuille_pv is None:
        for ws in wb.worksheets:
            if ws.title.startswith("PV "):
                feuille_pv = ws
                break

    if feuille_pv is None:
        wb.close()
        wb_formulas.close()
        return pd.DataFrame(), []

    pv_title = feuille_pv.title
    max_col = feuille_pv.max_column

    # Résoudre les cellules fusionnées pour lignes 2 et 3
    merged_map = {}
    for mr in feuille_pv.merged_cells.ranges:
        val = feuille_pv.cell(mr.min_row, mr.min_col).value
        for r in range(mr.min_row, mr.max_row + 1):
            for c in range(mr.min_col, mr.max_col + 1):
                merged_map[(r, c)] = val

    def gcell(row, col):
        v = merged_map.get((row, col), feuille_pv.cell(row, col).value)
        return v

    row2 = [gcell(2, c) for c in range(1, max_col + 1)]
    row3 = [gcell(3, c) for c in range(1, max_col + 1)]

    # Mots-clés pour catégoriser les colonnes
    mots_id      = {"cin", "massar", "nom", "prénom", "prenom", "identité", "identite",
                    "رقم", "الاسم", "النسب", "هوية", "هوية الطالب"}
    mots_synthese = {"moyenne générale", "modules validés", "crédits", "décision",
                     "rattrapages", "mention", "المعدل العام", "الوحدات", "الأرصدة",
                     "القرار", "الاستدراك", "الميزة"}
    mots_recap    = {"semestre", "récapitulatif", "recap", "الفصل"}

    # Construire col_info : index 0-based → (nom, type)
    # type : "CIN"|"Massar"|"Nom"|"Prénom" pour identité, "note"|"av"|"dec" pour modules
    col_info  = {}
    cur_group = None

    for i in range(max_col):
        v2 = str(row2[i]).strip() if row2[i] else ""
        v3 = str(row3[i]).strip() if row3[i] else ""
        v2l, v3l = v2.lower(), v3.lower()

        # Déterminer le groupe depuis ligne 2
        if v2:
            if any(k in v2l for k in mots_id):
                cur_group = "__id__"
            elif any(k in v2l for k in mots_synthese):
                cur_group = "__synth__"
            elif any(k in v2l for k in mots_recap):
                cur_group = "__recap__"
            else:
                cur_group = v2  # nom du module

        if cur_group == "__id__":
            if "cin" in v3l or "بطاقة" in v3l:
                col_info[i] = ("CIN", "id")
            elif "massar" in v3l or "مسار" in v3l:
                col_info[i] = ("Massar", "id")
            elif v3l in ("nom", "الاسم"):
                col_info[i] = ("Nom", "id")
            elif v3l in ("prénom", "prenom", "النسب"):
                col_info[i] = ("Prénom", "id")
        elif cur_group and cur_group not in ("__synth__", "__recap__"):
            if any(k in v3l for k in ("moy", "moyenne", "المعدل")):
                col_info[i] = (cur_group, "note")
            elif any(k in v3l for k in ("av", "avis", "الرأي")):
                col_info[i] = (cur_group, "av")
            elif any(k in v3l for k in ("déc", "dec", "قرار")):
                col_info[i] = (cur_group, "dec")

    # Liste des modules dans l'ordre d'apparition
    noms_modules = []
    seen = set()
    for i in sorted(col_info.keys()):
        nom, typ = col_info[i]
        if typ != "id" and nom not in seen:
            noms_modules.append(nom)
            seen.add(nom)

    # Lire les lignes de données — on s'arrête dès 3 lignes vides consécutives
    records = []
    lignes_vides = 0
    row_idx = 3
    while True:
        row_idx += 1
        vals = [_resolve_crosssheet(wb_formulas, wb, pv_title, row_idx, c)
                for c in range(1, max_col + 1)]
        if all(v is None for v in vals):
            lignes_vides += 1
            if lignes_vides >= 3:
                break
            continue
        lignes_vides = 0

        rec = {}
        for i, val in enumerate(vals):
            if i not in col_info:
                continue
            nom, typ = col_info[i]
            if typ == "id":
                rec[nom] = val
            elif typ == "note":
                # Convertir les symboles vides ("—", tiret, etc.) en None
                VIDES = {"—", "-", "—", "N/A", "NA", "ABS", "None", "nan"}
                val_clean = val.strip() if isinstance(val, str) else val
                if isinstance(val_clean, str) and val_clean in VIDES:
                    val_clean = None
                rec[f"{nom}_Note_Finale"] = val_clean
                try:
                    rec[f"{nom}_Validé"] = float(val_clean) >= 10 if val_clean not in (None, "") else False
                except (TypeError, ValueError):
                    rec[f"{nom}_Validé"] = False
            elif typ == "av":
                rec[f"{nom}_AV"] = val
            elif typ == "dec":
                rec[f"{nom}_Decision"] = val

        if rec:
            records.append(rec)

    wb.close()
    wb_formulas.close()

    if not records:
        return pd.DataFrame(), noms_modules

    df = pd.DataFrame(records)

    # Colonne CLE pour la fusion
    if "Massar" in df.columns and df["Massar"].notna().any():
        df["CLE"] = df["Massar"].astype(str).str.strip().str.upper()
    elif "CIN" in df.columns:
        df["CLE"] = df["CIN"].astype(str).str.strip().str.upper()
    else:
        df["CLE"] = df.index.astype(str)

    return df, noms_modules


# ─── Génération PV Annuel ─────────────────────────────────────────────────────

def _lire_pv_semestre_pour_annuel(
    chemin_pv: str,
    label:     str,
    nom_feuille_dans_dest: str,
) -> tuple[pd.DataFrame, list[str], dict]:
    """
    Lit la feuille PV d'un fichier PV semestre pour préparer le PV annuel.

    Retourne :
      - df         : DataFrame avec colonnes identité + {mod}_Note_Finale/_Validé/_AV/_Decision
      - noms_modules : liste ordonnée des noms de modules trouvés dans la feuille
      - liaison    : dict {nom_module: {cle_to_row, col_note, col_dec}}
                     où cle_to_row → ligne Excel réelle dans nom_feuille_dans_dest,
                     col_note / col_dec → colonnes 1-based de Moy. et Déc. du module
                     dans la feuille PV (pas la feuille module originale).

    La liaison permet à _ecrire_pv_dans_wb de générer des formules
    ='PV S1'!$E4  au lieu de valeurs statiques.
    """
    if not chemin_pv or not os.path.exists(chemin_pv):
        return pd.DataFrame(), [], {}

    wb = load_workbook(chemin_pv, data_only=True)
    wb_formulas = load_workbook(chemin_pv, data_only=False)

    def _resolve_cs(sheet_title, row, col):
        """Résout une formule cross-sheet si data_only retourne None."""
        import re as _re
        from openpyxl.utils import column_index_from_string
        ws_d = wb[sheet_title] if sheet_title in wb.sheetnames else None
        ws_f = wb_formulas[sheet_title] if sheet_title in wb_formulas.sheetnames else None
        if ws_d is None:
            return None
        val = ws_d.cell(row, col).value
        if val is not None:
            return val
        if ws_f is None:
            return val
        formula = ws_f.cell(row, col).value
        if not isinstance(formula, str) or not formula.startswith("="):
            return val
        m = _re.match(r"='?([^!']+)'?![$]?([A-Z]+)[$]?(\d+)", formula)
        if not m:
            return val
        src_sheet = m.group(1).replace("''", "'")
        src_col   = column_index_from_string(m.group(2))
        src_row   = int(m.group(3))
        if src_sheet in wb.sheetnames:
            src_val = wb[src_sheet].cell(src_row, src_col).value
            if src_val is None and src_sheet in wb_formulas.sheetnames:
                src_formula = wb_formulas[src_sheet].cell(src_row, src_col).value
                if isinstance(src_formula, str) and src_formula.startswith("="):
                    m2 = _re.match(r"='?([^!']+)'?![$]?([A-Z]+)[$]?(\d+)", src_formula)
                    if m2:
                        s2 = m2.group(1).replace("''", "'")
                        c2 = column_index_from_string(m2.group(2))
                        r2 = int(m2.group(3))
                        if s2 in wb.sheetnames:
                            src_val = wb[s2].cell(r2, c2).value
            return src_val
        return val

    # Trouver la feuille PV
    feuille_pv = None
    for ws in wb.worksheets:
        if ws.title == f"PV {label}" or ws.title.startswith(f"PV {label}"):
            feuille_pv = ws
            break
    if feuille_pv is None:
        for ws in wb.worksheets:
            if ws.title.startswith("PV "):
                feuille_pv = ws
                break
    if feuille_pv is None:
        wb.close()
        wb_formulas.close()
        return pd.DataFrame(), [], {}

    pv_title = feuille_pv.title
    max_col  = feuille_pv.max_column

    # Résoudre les fusions pour lignes 2 et 3
    merged_map: dict[tuple, object] = {}
    for mr in feuille_pv.merged_cells.ranges:
        val = feuille_pv.cell(mr.min_row, mr.min_col).value
        for r in range(mr.min_row, mr.max_row + 1):
            for c in range(mr.min_col, mr.max_col + 1):
                merged_map[(r, c)] = val

    def gcell(row, col):
        return merged_map.get((row, col), feuille_pv.cell(row, col).value)

    row2 = [gcell(2, c) for c in range(1, max_col + 1)]
    row3 = [gcell(3, c) for c in range(1, max_col + 1)]

    # Catégoriser les colonnes
    mots_id       = {"cin", "massar", "nom", "prénom", "prenom", "identité", "identite",
                     "رقم", "الاسم", "النسب", "هوية", "هوية الطالب"}
    mots_synthese = {"moyenne générale", "modules validés", "crédits", "décision",
                     "rattrapages", "mention", "المعدل العام", "الوحدات", "الأرصدة",
                     "القرار", "الاستدراك", "الميزة"}
    mots_recap    = {"semestre", "récapitulatif", "recap", "الفصل"}

    # col_info : index 0-based → (nom_groupe, type)
    col_info:  dict[int, tuple] = {}
    # col_excel : index 0-based → numéro de colonne 1-based dans la feuille PV
    cur_group = None

    for i in range(max_col):
        v2  = str(row2[i]).strip() if row2[i] else ""
        v3  = str(row3[i]).strip() if row3[i] else ""
        v2l, v3l = v2.lower(), v3.lower()

        if v2:
            if any(k in v2l for k in mots_id):
                cur_group = "__id__"
            elif any(k in v2l for k in mots_synthese):
                cur_group = "__synth__"
            elif any(k in v2l for k in mots_recap):
                cur_group = "__recap__"
            else:
                cur_group = v2  # nom du module

        if cur_group == "__id__":
            if "cin" in v3l or "بطاقة" in v3l:
                col_info[i] = ("CIN", "id")
            elif "massar" in v3l or "مسار" in v3l:
                col_info[i] = ("Massar", "id")
            elif v3l in ("nom", "الاسم"):
                col_info[i] = ("Nom", "id")
            elif v3l in ("prénom", "prenom", "النسب"):
                col_info[i] = ("Prénom", "id")
        elif cur_group and cur_group not in ("__synth__", "__recap__"):
            if any(k in v3l for k in ("moy", "moyenne", "المعدل")):
                col_info[i] = (cur_group, "note")
            elif any(k in v3l for k in ("av", "avis", "الرأي")):
                col_info[i] = (cur_group, "av")
            elif any(k in v3l for k in ("déc", "dec", "قرار")):
                col_info[i] = (cur_group, "dec")

    # Liste ordonnée des modules
    noms_modules: list[str] = []
    seen: set[str] = set()
    for i in sorted(col_info.keys()):
        nom, typ = col_info[i]
        if typ != "id" and nom not in seen:
            noms_modules.append(nom)
            seen.add(nom)

    # Colonnes Excel 1-based pour chaque module (note = Moy., dec = Déc.)
    col_note_pv: dict[str, int] = {}   # nom_module → colonne note 1-based dans feuille PV
    col_dec_pv:  dict[str, int] = {}   # nom_module → colonne déc  1-based dans feuille PV
    for i, (nom, typ) in col_info.items():
        if typ == "note" and nom not in col_note_pv:
            col_note_pv[nom] = i + 1   # 0-based → 1-based
        elif typ == "dec" and nom not in col_dec_pv:
            col_dec_pv[nom]  = i + 1

    # Lire les lignes de données (ligne 4+)
    VIDES = {"—", "-", "—", "N/A", "NA", "ABS", "None", "nan"}
    records: list[dict] = []
    lignes_vides = 0
    row_idx = 3
    while True:
        row_idx += 1
        vals = [_resolve_cs(pv_title, row_idx, c) for c in range(1, max_col + 1)]
        if all(v is None for v in vals):
            lignes_vides += 1
            if lignes_vides >= 3:
                break
            continue
        lignes_vides = 0

        rec: dict = {}
        for i, val in enumerate(vals):
            if i not in col_info:
                continue
            nom, typ = col_info[i]
            if typ == "id":
                rec[nom] = val
            elif typ == "note":
                val_c = val.strip() if isinstance(val, str) else val
                if isinstance(val_c, str) and val_c in VIDES:
                    val_c = None
                rec[f"{nom}_Note_Finale"] = val_c
                try:
                    rec[f"{nom}_Validé"] = float(val_c) >= 10 if val_c not in (None, "") else False
                except (TypeError, ValueError):
                    rec[f"{nom}_Validé"] = False
            elif typ == "av":
                rec[f"{nom}_AV"] = val
            elif typ == "dec":
                rec[f"{nom}_Decision"] = val
        if rec:
            records.append(rec)

    wb.close()
    wb_formulas.close()

    if not records:
        return pd.DataFrame(), noms_modules, {}

    df = pd.DataFrame(records)

    # Colonne CLE
    if "Massar" in df.columns and df["Massar"].notna().any():
        df["CLE"] = df["Massar"].astype(str).str.strip().str.upper()
    elif "CIN" in df.columns:
        df["CLE"] = df["CIN"].astype(str).str.strip().str.upper()
    else:
        df["CLE"] = df.index.astype(str)

    # Construire la liaison : chaque module pointe vers la feuille PV copiée
    # Les données commencent à la ligne 4 dans la feuille PV (lignes 1-3 = headers)
    LIGNE_DEBUT_PV = 4
    liaison: dict[str, dict] = {}
    for nom in noms_modules:
        if nom not in col_note_pv:
            continue
        cle_to_row: dict[str, int] = {}
        for df_idx, row in df.iterrows():
            cle = str(row.get("CLE", "")).strip().upper()
            if cle:
                cle_to_row[cle] = LIGNE_DEBUT_PV + df_idx   # ligne Excel réelle
        liaison[nom] = {
            "cle_to_row": cle_to_row,
            "col_note":   col_note_pv[nom],
            "col_dec":    col_dec_pv.get(nom),
            "feuille":    nom_feuille_dans_dest,   # nom de la feuille PV dans le wb dest
        }

    return df, noms_modules, liaison


def generer_pv_annuel(
    fichier_pv_si:   str,
    fichier_pv_sp:   str,
    label_si:        str,
    label_sp:        str,
    chemin_pv_prec:  str | None,
    dossier_sortie:  str,
    ecole:           str,
    filiere:         str,
    annee:           str,
    langue:          str = "fr",
) -> str:
    """
    PV Annuel :
      1. Copie TOUTES les feuilles des deux fichiers PV semestre (modules + feuille PV).
      2. Lit la feuille PV de chaque semestre pour récupérer données et liaison.
      3. Génère la feuille "PV Année" avec formules pointant vers les feuilles PV copiées.
    """
    os.makedirs(dossier_sortie, exist_ok=True)

    # ── 1. Workbook de sortie ──────────────────────────────────────────────
    wb = Workbook()
    wb.remove(wb.active)

    # ── 2. Copier toutes les feuilles des deux PV semestre ─────────────────
    # On mémorise le nom final attribué à la feuille "PV {label}" dans wb_dest
    nom_feuille_pv_si = _copier_toutes_feuilles_avec_noms(fichier_pv_si, wb, label_si)
    nom_feuille_pv_sp = _copier_toutes_feuilles_avec_noms(fichier_pv_sp, wb, label_sp)

    if not nom_feuille_pv_si and not nom_feuille_pv_sp:
        raise ValueError(
            "Aucune feuille PV trouvée dans les fichiers fournis.\n"
            f"Fichier SI : {fichier_pv_si}\n"
            f"Fichier SP : {fichier_pv_sp}"
        )

    # ── 3. Lire données + construire liaison depuis les feuilles PV ────────
    df_si, noms_si, liaison_si = _lire_pv_semestre_pour_annuel(
        fichier_pv_si, label_si, nom_feuille_pv_si or f"PV {label_si}"
    )
    df_sp, noms_sp, liaison_sp = _lire_pv_semestre_pour_annuel(
        fichier_pv_sp, label_sp, nom_feuille_pv_sp or f"PV {label_sp}"
    )

    if df_si.empty and df_sp.empty:
        raise ValueError(
            "Aucune donnée trouvée dans les feuilles PV.\n"
            f"Vérifiez que les fichiers contiennent une feuille 'PV {label_si}' et 'PV {label_sp}'."
        )

    # ── 4. Fusionner sur la clé CLE ────────────────────────────────────────
    cols_id = ["CIN", "Massar", "Nom", "Prénom", "CLE"]

    if df_si.empty:
        df = df_sp.copy()
    elif df_sp.empty:
        df = df_si.copy()
    else:
        # Inclure les colonnes d'identité de df_sp dans le merge pour récupérer
        # les coordonnées des étudiants présents uniquement au S2 (outer join).
        cols_sp_id_existantes = [c for c in ("CIN", "Massar", "Nom", "Prénom") if c in df_sp.columns]
        cols_sp_notes = [c for c in df_sp.columns if c not in cols_id]
        df = pd.merge(
            df_si,
            df_sp[["CLE"] + cols_sp_id_existantes + cols_sp_notes],
            on="CLE", how="outer", suffixes=("_x", "_y")
        )
        for col in ("CIN", "Massar", "Nom", "Prénom"):
            col_x, col_y = f"{col}_x", f"{col}_y"
            if col_x in df.columns and col_y in df.columns:
                # Priorité à S1 (_x), compléter avec S2 (_y) quand S1 est vide
                df[col] = df[col_x].combine_first(df[col_y])
                df.drop(columns=[col_x, col_y], inplace=True, errors="ignore")
            elif col_x in df.columns:
                df.rename(columns={col_x: col}, inplace=True)
            elif col_y in df.columns:
                df.rename(columns={col_y: col}, inplace=True)
            # Si col existe déjà sans suffixe (pas de doublon), ne rien faire

    # ── 4b. Compléter les notes manquantes depuis le PV de l'année précédente ──
    # Cas redoublants : si une note est absente dans les fiches de cette année,
    # on la récupère depuis le PV précédent (note déjà validée).
    if chemin_pv_prec:
        try:
            df_pv_prec = lire_pv_precedent(chemin_pv_prec)
            df = _completer_notes(df, df_pv_prec, noms_si + noms_sp)
            print("  ✓ Notes complétées depuis le PV précédent.")
        except Exception as e:
            print(f"  ⚠ PV précédent ignoré (erreur de lecture) : {e}")

    # ── 5. Calculer la synthèse annuelle ───────────────────────────────────
    coeffs = {m: 1.0 for m in noms_si + noms_sp}
    df_pv  = construire_pv(
        df=df,
        modules_si=noms_si,
        modules_sp=noms_sp,
        coefficients=coeffs,
        label_si=label_si,
        label_sp=label_sp,
    )

    # ── 6. Liaison combinée SI + SP ────────────────────────────────────────
    liaison_annuel = {**liaison_si, **liaison_sp}

    # ── 7. Feuille "PV Année" ─────────────────────────────────────────────
    _ecrire_pv_dans_wb(
        wb=wb, df_pv=df_pv,
        modules_si=noms_si, modules_sp=noms_sp,
        label_si=label_si, label_sp=label_sp,
        ecole=ecole, filiere=filiere, annee=annee, langue=langue,
        titre_feuille="PV Année",
        inclure_synthese=True,
        liaison_modules=liaison_annuel,
    )

    # ── 8. Sauvegarder ────────────────────────────────────────────────────
    slug = re.sub(r"[^\w\-]", "_", annee) if annee else "annee"
    chemin_out = os.path.join(dossier_sortie, f"PV_Annuel_{label_si}_{label_sp}_{slug}.xlsx")
    wb.save(chemin_out)
    print(f"PV Annuel généré : {chemin_out}")

    # ── 9. Injection des notes redoublants (PV S1, PV S2, PV Année) ───────
    # Les étudiants absents des fiches modules de cette année (redoublants)
    # ont des cellules '—' statiques. On les remplace par les notes du PV
    # précédent, en gérant le croisement CIN↔Massar entre les deux années.
    if chemin_pv_prec and os.path.exists(chemin_pv_prec):
        try:
            from modules.inject_redoublants import inject_notes_redoublants
            print("\n  Injection des notes redoublants…")
            inject_notes_redoublants(
                chemin_xlsx=chemin_out,
                chemin_pv_prec=chemin_pv_prec,
                chemin_sortie=chemin_out,
            )
        except Exception as e:
            import traceback
            print(f"  ⚠ Injection redoublants échouée (fichier conservé tel quel) : {e}")
            print(traceback.format_exc())

    return chemin_out


# ─── Écriture d'une feuille PV dans un workbook ───────────────────────────────


def _trouver_ligne_liaison(lien: dict, row) -> int | None:
    """
    Cherche la ligne Excel d'un étudiant dans la liaison en testant
    CLE, Massar et CIN successivement.
    """
    cle_to_row = lien.get("cle_to_row", {})
    if not cle_to_row:
        return None
    for champ in ("CLE", "Massar", "CIN"):
        cle = str(row.get(champ, "") or "").strip().upper()
        if cle and cle in cle_to_row:
            return cle_to_row[cle]
    return None


def _formule_note(nom_module: str, row, liaison_modules: dict | None) -> str | None:
    """Formule dynamique vers Note Finale du module pour cet étudiant."""
    if not liaison_modules or nom_module not in liaison_modules:
        return None
    lien = liaison_modules[nom_module]
    ligne_xlsx = _trouver_ligne_liaison(lien, row)
    if not ligne_xlsx:
        return None
    col_xlsx = lien.get("col_note")
    if not col_xlsx:
        return None
    nom_feuille = lien.get("feuille", nom_module)[:31].replace("'", "''")
    return f"='{nom_feuille}'!${get_column_letter(col_xlsx)}{ligne_xlsx}"


def _formule_dec(nom_module: str, row, liaison_modules: dict | None) -> str | None:
    """Formule dynamique vers Décision Finale du module pour cet étudiant."""
    if not liaison_modules or nom_module not in liaison_modules:
        return None
    lien    = liaison_modules[nom_module]
    col_dec = lien.get("col_dec")
    if not col_dec:
        return None
    ligne_xlsx = _trouver_ligne_liaison(lien, row)
    if not ligne_xlsx:
        return None
    nom_feuille = lien.get("feuille", nom_module)[:31].replace("'", "''")
    return f"='{nom_feuille}'!${get_column_letter(col_dec)}{ligne_xlsx}"


def _formule_annee(nom_module: str, row, liaison_modules: dict | None) -> str | None:
    """Formule dynamique vers colonne سنة/AV du module pour cet étudiant."""
    if not liaison_modules or nom_module not in liaison_modules:
        return None
    lien      = liaison_modules[nom_module]
    col_annee = lien.get("col_annee")
    if not col_annee:
        return None
    ligne_xlsx = _trouver_ligne_liaison(lien, row)
    if not ligne_xlsx:
        return None
    nom_feuille = lien.get("feuille", nom_module)[:31].replace("'", "''")
    return f"='{nom_feuille}'!${get_column_letter(col_annee)}{ligne_xlsx}"


def _ecrire_pv_dans_wb(
    wb:               Workbook,
    df_pv:            pd.DataFrame,
    modules_si:       list[str],
    modules_sp:       list[str],
    label_si:         str,
    label_sp:         str,
    ecole:            str,
    filiere:          str,
    annee:            str,
    langue:           str,
    titre_feuille:    str,
    inclure_synthese: bool = True,
    liaison_modules:  dict | None = None,
) -> None:
    """
    Écrit le PV dans une nouvelle feuille du workbook.
    inclure_synthese=False → on s'arrête après le bloc semestre (sans colonnes finales).
    inclure_synthese=True  → on ajoute Moy Générale, Mods Validés, Crédits, Décision, Rattrapages, Mention.
    """
    from modules.exporter import (
        _labels, _cell, _merge, _fmt_note, _fmt_dec_module,
        _safe, _est_valide_row,
        BLEU_DARK, BLEU_MID, GRIS, BLANC,
        VERT, VERT_T, ORANGE, ORANGE_T, ROUGE, ROUGE_T
    )

    L     = _labels(langue, label_si, label_sp, ecole, filiere, annee)
    is_ar = (langue == "ar")

    n_si    = len(modules_si)
    n_sp    = len(modules_sp)
    recap_si = 3 if n_si > 0 else 0
    recap_sp = 3 if n_sp > 0 else 0
    # PV Semestre : 1 colonne Mention par semestre recap (en plus des 3 standard)
    # PV Annuel   : 6 colonnes de synthèse finales
    n_mention_sem = (1 if not inclure_synthese and n_si > 0 else 0) + \
                    (1 if not inclure_synthese and n_sp > 0 else 0)
    n_synth  = 6 if inclure_synthese else 0
    n_total  = 4 + n_si*3 + recap_si + n_sp*3 + recap_sp + n_synth + n_mention_sem

    existing   = [s.title for s in wb.worksheets]
    sheet_name = titre_feuille[:31]
    if sheet_name in existing:
        sheet_name = sheet_name[:28] + "_PV"

    ws = wb.create_sheet(title=sheet_name)
    if is_ar:
        ws.sheet_view.rightToLeft = True

    # ── Ligne 1 : En-tête ──
    _merge(ws, 1, 1, 1, n_total,
           value=L["entete"], size=11, bold=True, fill=BLEU_DARK, color=BLANC)
    ws.row_dimensions[1].height = 28

    # ── Ligne 2 : groupes ──
    col = 1
    _merge(ws, 2, col, 2, col+3, L["identite"], fill=BLEU_DARK, size=9)
    col += 4

    for mod in modules_si:
        _merge(ws, 2, col, 2, col+2, mod, fill=BLEU_MID, size=8, wrap=True)
        col += 3
    if modules_si:
        _n_recap_si = 4 if not inclure_synthese else 3   # +1 pour Mention semestre
        _merge(ws, 2, col, 2, col+_n_recap_si-1, L["recap_si"], fill=BLEU_DARK, size=9)
        col += _n_recap_si

    for mod in modules_sp:
        _merge(ws, 2, col, 2, col+2, mod, fill=BLEU_MID, size=8, wrap=True)
        col += 3
    if modules_sp:
        _n_recap_sp = 4 if not inclure_synthese else 3   # +1 pour Mention semestre
        _merge(ws, 2, col, 2, col+_n_recap_sp-1, L["recap_sp"], fill=BLEU_DARK, size=9)
        col += _n_recap_sp

    if inclure_synthese:
        finales = [L["moy_gen"], L["mods_valides"], L["credits"],
                   L["decision"], L["rattrapages"], L["mention"]]
        for i, h in enumerate(finales):
            _merge(ws, 2, col+i, 3, col+i, h, fill=BLEU_DARK, size=9, wrap=True)

    ws.row_dimensions[2].height = 30

    # ── Ligne 3 : sous-colonnes ──
    col = 1
    for h in [L["cin"], L["massar"], L["nom"], L["prenom"]]:
        _cell(ws, 3, col, h, bold=True, fill=BLEU_DARK); col += 1

    for _ in modules_si:
        for h in [L["moy"], L["av"], L["dec"]]:
            _cell(ws, 3, col, h, bold=True, fill=BLEU_MID); col += 1
    if modules_si:
        _headers_recap_si = [L["moy_si"], L["dec_si"], L["av_si"]]
        if not inclure_synthese:
            _headers_recap_si.append(L["mention"])
        for h in _headers_recap_si:
            _cell(ws, 3, col, h, bold=True, fill=BLEU_DARK); col += 1

    for _ in modules_sp:
        for h in [L["moy"], L["av"], L["dec"]]:
            _cell(ws, 3, col, h, bold=True, fill=BLEU_MID); col += 1
    if modules_sp:
        _headers_recap_sp = [L["moy_sp"], L["dec_sp"], L["av_sp"]]
        if not inclure_synthese:
            _headers_recap_sp.append(L["mention"])
        for h in _headers_recap_sp:
            _cell(ws, 3, col, h, bold=True, fill=BLEU_DARK); col += 1

    ws.row_dimensions[3].height = 18

    # ── Précalculer les colonnes note de chaque module dans la feuille PV ──
    # Nécessaire pour générer les formules AVERAGE(note_mod1, note_mod2, ...)
    # dans les colonnes Moy.Si et Moy.Sp.
    # On simule le compteur col pour identifier la colonne PV de chaque module.
    _col_note_pv: dict[str, int] = {}   # {nom_module: colonne_note_dans_pv (1-based)}
    _col_dec_pv:  dict[str, int] = {}   # {nom_module: colonne_dec_dans_pv (1-based)}
    _col_sim = 5   # après les 4 colonnes identité
    for mod in modules_si:
        _col_note_pv[mod] = _col_sim
        _col_dec_pv[mod]  = _col_sim + 2
        _col_sim += 3
    _col_recap_si = _col_sim if modules_si else None   # colonne Moy.Si
    _col_dec_si   = _col_sim + 1 if modules_si else None  # colonne Déc.Si
    _col_av_si    = _col_sim + 2 if modules_si else None  # colonne AV Si
    if modules_si:
        _col_sim += 4 if not inclure_synthese else 3   # +Mention en mode semestre
    for mod in modules_sp:
        _col_note_pv[mod] = _col_sim
        _col_dec_pv[mod]  = _col_sim + 2
        _col_sim += 3
    _col_recap_sp = _col_sim if modules_sp else None   # colonne Moy.Sp
    _col_dec_sp   = _col_sim + 1 if modules_sp else None  # colonne Déc.Sp
    _col_av_sp    = _col_sim + 2 if modules_sp else None  # colonne AV Sp

    # ── Données ──
    for i, (_, row) in enumerate(df_pv.iterrows()):
        r   = 4 + i
        bg  = GRIS if i % 2 == 0 else BLANC
        col = 1

        for champ in ["CIN", "Massar", "Nom", "Prénom"]:
            _cell(ws, r, col, _safe(row.get(champ, "")),
                  fill=bg, color="000000", align="left"); col += 1

        for mod in modules_si:
            # MOY / AV / DEC : toujours dynamiques via formule Excel
            note_val  = _formule_note(mod, row, liaison_modules)
            dec_src   = _formule_dec(mod, row, liaison_modules)
            annee_val = _formule_annee(mod, row, liaison_modules)
            if note_val  is None: note_val  = ""
            if dec_src   is None: dec_src   = ""
            # سنة du module : dynamique si liaison dispo, sinon fallback statique
            if annee_val is None:
                annee_val = str(annee or "").strip()
            if note_val is None: note_val = ""
            if dec_src  is None: dec_src  = ""
            _cell(ws, r, col,   note_val,  fill=bg, color="000000", bold=True)
            _cell(ws, r, col+1, annee_val, fill=bg, color="000000")
            _cell(ws, r, col+2, dec_src,   fill=bg, color="000000")
            col += 3

        if modules_si:
            moy_si_disp = _fmt_note(row.get(f"Moy_{label_si}"))
            dec_si_calc = _safe(row.get(f"Dec_{label_si}", ""))
            men_si_calc = _safe(row.get(f"Mention_{label_si}", ""))
            si_ok = (dec_si_calc == L["valide"] or dec_si_calc == "Validé")
            try:
                if not si_ok:
                    si_ok = moy_si_disp != "—" and float(moy_si_disp) >= 10 and dec_si_calc not in (L["non_valide"], "Non Validé")
            except Exception:
                pass
            bs = bg; ts = "000000"
            # Formule dynamique : AVERAGE sur les colonnes notes SI de cette ligne
            _refs_si = ",".join(f"{get_column_letter(_col_note_pv[m])}{r}" for m in modules_si)
            _col_moy_si_lettre = get_column_letter(col)
            _formule_moy_si = f"=ROUND(AVERAGE({_refs_si}),2)"

           # Décision semestrielle dynamique (Arabic) :
            # ر dans un module → غير مكتسب immédiat
            # غ م dans plus de 2 modules → غير مكتسب
            # sinon : مكتسب si moyenne >= 10
            if is_ar:
                _count_ghm = "+".join(
                    f'IF({get_column_letter(_col_dec_pv[m])}{r}="غ م",1,0)'
                    for m in modules_si
                )
                _count_r = "+".join(
                    f'IF({get_column_letter(_col_dec_pv[m])}{r}="ر",1,0)'
                    for m in modules_si
                )
                _formule_dec_si = (
                    f'=IF(OR(({_count_r})>0,({_count_ghm})>2),'
                    f'"{L["non_valide"]}",'
                    f'IF({_col_moy_si_lettre}{r}>=10,"{L["valide"]}","{L["non_valide"]}"))'
                )
            else:
                _count_nv = "+".join(f'IF({get_column_letter(_col_dec_pv[m])}{r}="NV",1,0)' for m in modules_si)
                _count_elim = "+".join(f'IF(AND(ISNUMBER({get_column_letter(_col_note_pv[m])}{r}),{get_column_letter(_col_note_pv[m])}{r}<5),1,0)' for m in modules_si)
                _formule_dec_si = (
                    f'=IF(OR(({_count_elim})>0,({_count_nv})>2),'
                    f'"{L["non_valide"]}",'
                    f'IF({_col_moy_si_lettre}{r}>=10,"{L["valide"]}","{L["non_valide"]}"))'
                )

            # سنة التحقق S1 : TOUJOURS statique (année actuelle), arabe ou français
            import re as _re_av
            _annee_static = str(annee or "").strip()
            _m_av = _re_av.match(r'^(\d{2,4})-(\d{2,4})$', _annee_static)
            if _m_av:
                _annee_static = f"{_m_av.group(1)[-2:]}-{_m_av.group(2)[-2:]}"
            annee_si_val = _annee_static   # statique dans tous les cas

            # Mention dynamique : basée sur la moyenne semestrielle (arabe ou français)
            if is_ar:
                _formule_mention_si = (
                    f'=IF({_col_moy_si_lettre}{r}>=16,"{L["tb"]}",'
                    f'IF({_col_moy_si_lettre}{r}>=14,"{L["b"]}",'
                    f'IF({_col_moy_si_lettre}{r}>=12,"{L["ab"]}",'
                    f'IF({_col_moy_si_lettre}{r}>=10,"{L["p"]}",""))))'
                )
            else:
                _formule_mention_si = (
                    f'=IF({_col_moy_si_lettre}{r}>=16,"Très Bien",'
                    f'IF({_col_moy_si_lettre}{r}>=14,"Bien",'
                    f'IF({_col_moy_si_lettre}{r}>=12,"Assez Bien",'
                    f'IF({_col_moy_si_lettre}{r}>=10,"Passable",""))))'
                )

            _cell(ws, r, col,   _formule_moy_si,  fill=bs, bold=True, color=ts)
            _cell(ws, r, col+1, _formule_dec_si,  fill=bs, color=ts)
            _cell(ws, r, col+2, annee_si_val,      fill=bs, color=ts)
            col += 3
            if not inclure_synthese:
                _cell(ws, r, col, _formule_mention_si, fill=bs, bold=True, color=ts)
                col += 1

        for mod in modules_sp:
            # MOY / AV / DEC : toujours dynamiques via formule Excel
            note_val  = _formule_note(mod, row, liaison_modules)
            dec_src   = _formule_dec(mod, row, liaison_modules)
            annee_val = _formule_annee(mod, row, liaison_modules)
            if note_val  is None: note_val  = ""
            if dec_src   is None: dec_src   = ""
            # سنة du module : dynamique si liaison dispo, sinon fallback statique
            if annee_val is None:
                annee_val = str(annee or "").strip()
            if note_val is None: note_val = ""
            if dec_src  is None: dec_src  = ""
            _cell(ws, r, col,   note_val,  fill=bg, color="000000", bold=True)
            _cell(ws, r, col+1, annee_val, fill=bg, color="000000")
            _cell(ws, r, col+2, dec_src,   fill=bg, color="000000")
            col += 3

        if modules_sp:
            moy_sp_disp = _fmt_note(row.get(f"Moy_{label_sp}"))
            dec_sp_calc = _safe(row.get(f"Dec_{label_sp}", ""))
            men_sp_calc = _safe(row.get(f"Mention_{label_sp}", ""))
            sp_ok = (dec_sp_calc == L["valide"] or dec_sp_calc == "Validé")
            try:
                if not sp_ok:
                    sp_ok = moy_sp_disp != "—" and float(moy_sp_disp) >= 10 and dec_sp_calc not in (L["non_valide"], "Non Validé")
            except Exception:
                pass
            bp = bg; tp = "000000"
            # Formule dynamique : AVERAGE sur les colonnes notes SP de cette ligne
            _refs_sp = ",".join(f"{get_column_letter(_col_note_pv[m])}{r}" for m in modules_sp)
            _col_moy_sp_lettre = get_column_letter(col)
            _formule_moy_sp = f"=ROUND(AVERAGE({_refs_sp}),2)"

           # Décision semestrielle SP dynamique (Arabic) :
            # ر dans un module → غير مكتسب immédiat
            # غ م dans plus de 2 modules → غير مكتسب
            # sinon : مكتسب si moyenne >= 10
            if is_ar:
                _count_ghm_sp = "+".join(
                    f'IF({get_column_letter(_col_dec_pv[m])}{r}="غ م",1,0)'
                    for m in modules_sp
                )
                _count_r_sp = "+".join(
                    f'IF({get_column_letter(_col_dec_pv[m])}{r}="ر",1,0)'
                    for m in modules_sp
                )
                _formule_dec_sp = (
                    f'=IF(OR(({_count_r_sp})>0,({_count_ghm_sp})>2),'
                    f'"{L["non_valide"]}",'
                    f'IF({_col_moy_sp_lettre}{r}>=10,"{L["valide"]}","{L["non_valide"]}"))'
                )
            else:
                _count_nv_sp   = "+".join(f'IF({get_column_letter(_col_dec_pv[m])}{r}="NV",1,0)' for m in modules_sp)
                _count_elim_sp = "+".join(f'IF(AND(ISNUMBER({get_column_letter(_col_note_pv[m])}{r}),{get_column_letter(_col_note_pv[m])}{r}<5),1,0)' for m in modules_sp)
                _formule_dec_sp = (
                    f'=IF(OR(({_count_elim_sp})>0,({_count_nv_sp})>2),'
                    f'"{L["non_valide"]}",'
                    f'IF({_col_moy_sp_lettre}{r}>=10,"{L["valide"]}","{L["non_valide"]}"))'
                )

            # سنة التحقق S2 : TOUJOURS statique (année actuelle), arabe ou français
            import re as _re_av_sp
            _annee_static_sp = str(annee or "").strip()
            _m_av_sp = _re_av_sp.match(r'^(\d{2,4})-(\d{2,4})$', _annee_static_sp)
            if _m_av_sp:
                _annee_static_sp = f"{_m_av_sp.group(1)[-2:]}-{_m_av_sp.group(2)[-2:]}"
            annee_sp_val = _annee_static_sp   # statique dans tous les cas

            # Mention dynamique SP (arabe ou français)
            if is_ar:
                _formule_mention_sp = (
                    f'=IF({_col_moy_sp_lettre}{r}>=16,"{L["tb"]}",'
                    f'IF({_col_moy_sp_lettre}{r}>=14,"{L["b"]}",'
                    f'IF({_col_moy_sp_lettre}{r}>=12,"{L["ab"]}",'
                    f'IF({_col_moy_sp_lettre}{r}>=10,"{L["p"]}",""))))'
                )
            else:
                _formule_mention_sp = (
                    f'=IF({_col_moy_sp_lettre}{r}>=16,"Très Bien",'
                    f'IF({_col_moy_sp_lettre}{r}>=14,"Bien",'
                    f'IF({_col_moy_sp_lettre}{r}>=12,"Assez Bien",'
                    f'IF({_col_moy_sp_lettre}{r}>=10,"Passable",""))))'
                )

            _cell(ws, r, col,   _formule_moy_sp,  fill=bp, bold=True, color=tp)
            _cell(ws, r, col+1, _formule_dec_sp,  fill=bp, color=tp)
            _cell(ws, r, col+2, annee_sp_val,      fill=bp, color=tp)
            col += 3
            if not inclure_synthese:
                _cell(ws, r, col, _formule_mention_sp, fill=bp, bold=True, color=tp)
                col += 1

        # ── Colonnes de synthèse (PV Annuel uniquement) ────────────────────
        if inclure_synthese:
            tous = modules_si + modules_sp
            n_tous = len(tous)

            # Colonnes DEC de tous les modules (dans la feuille PV courante)
            _dec_cols_tous = (
                [_col_dec_pv[m] for m in modules_si if m in _col_dec_pv] +
                [_col_dec_pv[m] for m in modules_sp if m in _col_dec_pv]
            )

            # ── Colonne AU : Moyenne Générale = (Moy.S1 + Moy.S2) / 2 ──────
            _col_moy_gen = col   # colonne AU dans la feuille courante
            _col_moy_gen_lettre = get_column_letter(_col_moy_gen)
            if _col_recap_si and _col_recap_sp:
                _formule_moy_gen = (
                    f"=ROUND(({get_column_letter(_col_recap_si)}{r}"
                    f"+{get_column_letter(_col_recap_sp)}{r})/2,2)"
                )
            elif _col_recap_si:
                _formule_moy_gen = f"={get_column_letter(_col_recap_si)}{r}"
            elif _col_recap_sp:
                _formule_moy_gen = f"={get_column_letter(_col_recap_sp)}{r}"
            else:
                _formule_moy_gen = ""

            # ── Colonne AV : Modules Validés (formule Excel dynamique) ───────
            if _dec_cols_tous:
                if is_ar:
                    # Arabe : "مكتسب" ou "مكتسب"
                    _valid_val = L.get("valide", "مكتسب")
                    _countif_parts = "+".join(
                        f'((UPPER({get_column_letter(c)}{r})="{_valid_val.upper()}")>0)*1'
                        for c in _dec_cols_tous
                    )
                else:
                    _countif_parts = "+".join(
                        f'((UPPER({get_column_letter(c)}{r})="V")+(UPPER({get_column_letter(c)}{r})="VALIDÉ")+(UPPER({get_column_letter(c)}{r})="VALIDÉE")>0)*1'
                        for c in _dec_cols_tous
                    )
                _formule_mods_val = f"=SUMPRODUCT({_countif_parts})"
            else:
                _formule_mods_val = "0"

            # ── Colonne AW : Crédits (formule Excel dynamique) ───────────────
            # Crédits = modules_validés × (60 / nombre_total_modules)
            # Ex : 12 modules → 60/12 = 5 crédits par module
            _credits_par_module = 60 / len(tous) if tous else 5
            _col_mods_val_lettre = get_column_letter(col + 1)
            cred = f"={_col_mods_val_lettre}{r}*{int(_credits_par_module)}"

            # ── Colonne AX : Décision annuelle (formule Excel dynamique) ─────
            if _formule_moy_gen:
                _mg_ref = f"{_col_moy_gen_lettre}{r}"
                if is_ar:
                    _formule_decision = (
                        f'=IF({_mg_ref}>=12,"{L.get("admis","ناجح")}",'
                        f'IF({_mg_ref}>=10,"{L.get("rachat","قابل للاستدراك")}","'
                        f'{L.get("ajourne","راسب")}"))'
                    )
                else:
                    # Colonnes notes de tous les modules (SI + SP)
                    _note_cols_tous = (
                        [_col_note_pv[m] for m in modules_si if m in _col_note_pv] +
                        [_col_note_pv[m] for m in modules_sp if m in _col_note_pv]
                    )
                    # Condition 2.1 : au moins un module < 7
                    _count_elim = "+".join(
                        f"IF(AND(ISNUMBER({get_column_letter(c)}{r}),{get_column_letter(c)}{r}<7),1,0)"
                        for c in _note_cols_tous
                    )
                    # Condition 2.2 : modules non validés (7 ≤ note < 10), compte > 2
                    _count_nv = "+".join(
                        f"IF(AND(ISNUMBER({get_column_letter(c)}{r}),{get_column_letter(c)}{r}>=7,{get_column_letter(c)}{r}<10),1,0)"
                        for c in _note_cols_tous
                    )
                    _formule_decision = (
                        f'=IF(OR({_mg_ref}<10,({_count_elim})>0,({_count_nv})>2),'
                        f'"NON ADMIS","ADMIS")'
                    )
            else:
                _formule_decision = ""

            # ── Colonne AY : Rattrapages (formule Excel dynamique) ────────────
            # Construire la liste (module, col_dec) dans le BON ordre : SI puis SP
            _modules_dec_ordonnes = (
                [(m, _col_dec_pv[m]) for m in modules_si if m in _col_dec_pv] +
                [(m, _col_dec_pv[m]) for m in modules_sp if m in _col_dec_pv]
            )
            if _modules_dec_ordonnes:
                _parts_rattrap = []
                for nom_mod, c_dec in _modules_dec_ordonnes:
                    _ref_dec = f"{get_column_letter(c_dec)}{r}"
                    _nom_safe = nom_mod.replace('"', '""')
                    if is_ar:
                        _nv = L.get("non_valide", "غير مكتسب")
                        _cond_rattrap = (
                            f'IF(OR(UPPER({_ref_dec})="R",'
                            f'UPPER({_ref_dec})="NE",'
                            f'UPPER({_ref_dec})="{_nv.upper()}"),'
                            f'"{_nom_safe} | ","")'
                        )
                    else:
                        _cond_rattrap = (
                            f'IF(OR(UPPER({_ref_dec})="R",'
                            f'UPPER({_ref_dec})="NE",'
                            f'UPPER({_ref_dec})="NV",'
                            f'UPPER({_ref_dec})="NON VALIDÉ",'
                            f'UPPER({_ref_dec})="NON VALIDÉE",'
                            f'UPPER({_ref_dec})="NON VALIDEE"),'
                            f'"{_nom_safe} | ","")'
                        )
                    _parts_rattrap.append(_cond_rattrap)

                _concat_rattrap = "&".join(_parts_rattrap)
                _formule_rattrap = (
                    f'=IF(LEN(TRIM({_concat_rattrap}))=0,"",'
                    f'LEFT(TRIM({_concat_rattrap}),'
                    f'LEN(TRIM({_concat_rattrap}))-2))'
                )
            else:
                _formule_rattrap = '=""'

            # ── Colonne AZ : Mention (formule Excel dynamique) ────────────────
            if _formule_moy_gen:
                _mg_ref = f"{_col_moy_gen_lettre}{r}"
                if is_ar:
                    _formule_mention = (
                        f'=IF({_mg_ref}>=16,"{L.get("tb","ممتاز")}",'
                        f'IF({_mg_ref}>=14,"{L.get("b","جيد جداً")}",'
                        f'IF({_mg_ref}>=12,"{L.get("ab","جيد")}",'
                        f'IF({_mg_ref}>=10,"{L.get("p","مقبول")}",""))))'
                    )
                else:
                    _formule_mention = (
                        f'=IF({_mg_ref}>=16,"Très Bien",'
                        f'IF({_mg_ref}>=14,"Bien",'
                        f'IF({_mg_ref}>=12,"Assez Bien",'
                        f'IF({_mg_ref}>=10,"Passable",""))))'
                    )
            else:
                _formule_mention = '=""'

            _cell(ws, r, col,   _formule_moy_gen,   fill=bg, bold=True,  color="000000")
            _cell(ws, r, col+1, _formule_mods_val,  fill=bg,              color="000000")
            _cell(ws, r, col+2, cred,               fill=bg,              color="000000")
            _cell(ws, r, col+3, _formule_decision,  fill=bg, bold=True,  color="000000")
            _cell(ws, r, col+4, _formule_rattrap,   fill=bg, wrap=True, size=8, color="000000")
            _cell(ws, r, col+5, _formule_mention,   fill=bg, bold=True,  color="000000")

        ws.row_dimensions[r].height = 15

    # ── Largeurs colonnes ──
    for col_cells in ws.columns:
        mx = 0
        cl = get_column_letter(col_cells[0].column)
        for cell in col_cells:
            if cell.value:
                mx = max(mx, len(str(cell.value)))
        ws.column_dimensions[cl].width = min(max(mx + 2, 6), 28)

    ws.freeze_panes = "E4"