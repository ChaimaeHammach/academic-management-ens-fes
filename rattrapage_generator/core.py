"""
core.py — Rattrapage Manager
=============================
Lecture des fiches Excel + génération des fichiers de sortie.
Supporte :
  - Format FR standard  : N° en col A, Decision col 8
  - Format AR inversé   : رقم en dernière colonne, القرار النهائي en col 1
  - Format multi-elem   : ex. Langues (Décision finale, NO en col A)
"""

import os
import re
import copy
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Décisions VALIDÉ → à exclure ────────────────────────────────────────────────
VALIDES = {
    # Français
    "validé", "valide", "v", "admis", "admise",
    # Arabe long
    "مقبول", "مقبول ",
    # Arabe abrégé
    "م",
    # Autres variantes
    "ن",
}

# ─── Palette ──────────────────────────────────────────────────────────────────
BLEU_DARK  = "1A3A5C"
BLEU_HDR   = "2563A8"
BLEU_LIGHT = "DBEAFE"
ROUGE_BG   = "FFC7CE"
ROUGE_TXT  = "9C0006"
VERT_BG    = "C6EFCE"
VERT_TXT   = "276221"
ORANGE_BG  = "FFEB9C"
ORANGE_TXT = "9C5700"
GRIS_BG    = "F1F5F9"
BLANC      = "FFFFFF"

GROUPE_COLORS = [
    "BDD7EE","FCE4D6","E2EFDA","FFF2CC",
    "D9D9D9","F4CCFF","D0E4F7","FFD966",
]


def _f(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def _b(style="thin", color="CBD5E1"):
    s = Side(style=style, color=color)
    return Border(left=s, right=s, top=s, bottom=s)

def _is_na(v):
    if v is None: return True
    try:
        import math; return math.isnan(float(v))
    except: return False

def _est_valide(decision: str) -> bool:
    d = decision.strip().lower()
    return d in VALIDES or re.sub(r"[\s\.\-é]","",d) in {
        re.sub(r"[\s\.\-é]","",v) for v in VALIDES
    }


# ─── Lecture d'une fiche ──────────────────────────────────────────────────────

def lire_fiche(chemin: str) -> dict:
    wb = load_workbook(chemin, data_only=True)   # pour lire les valeurs calculées
    wb_formulas = load_workbook(chemin, data_only=False)  # pour conserver les formules
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))

    # ── Détecter langue ────────────────────────────────────────────────────
    langue = "fr"
    for row in rows[:8]:
        line = " ".join(str(v or "") for v in row)
        if any(c > "\u0600" for c in line):
            langue = "ar"
            break

    # ── Extraire métadonnées (filière, année, module) ──────────────────────
    filiere, annee, module = "", "", ""
    for row in rows[:8]:
        v0 = str(row[0] or "").strip()
        if not v0:
            continue
        vl = v0.lower()
        if any(k in vl for k in ["licence","إجازة","license","filière"]):
            filiere = v0
        elif any(k in vl for k in ["année","annee","السنة","universitaire"]):
            annee = v0.split(":")[-1].strip() if ":" in v0 else v0
        elif any(k in vl for k in ["module","الوحدة","matière"]):
            module = v0.split(":")[-1].strip() if ":" in v0 else v0
        # Cas Langues-S6 : filière est directement ligne 0 sans mot clé
        elif not filiere and len(v0) > 20:
            filiere = v0

    # ── Trouver la ligne d'en-têtes ────────────────────────────────────────
    header_row_idx = None
    headers        = []

    # Marqueurs d'en-tête connus
    MARKERS_FR = {"n°","no","nom","prénom","prenom","cin","massar","cne"}
    MARKERS_AR = {"رقم","الاسم","الاسم العائلي","الاسم الشخصي","رقم البطاقة الوطنية"}

    for i, row in enumerate(rows):
        vals_str = [str(v or "").strip().lower() for v in row]
        non_empty = [v for v in vals_str if v]
        if len(non_empty) < 3:
            continue
        if any(v in MARKERS_FR for v in vals_str) or \
           any(v in MARKERS_AR for v in vals_str):
            header_row_idx = i
            headers = [str(v or "").strip() for v in row]
            break

    if header_row_idx is None:
        raise ValueError(
            f"Impossible de trouver la ligne d'en-têtes dans {os.path.basename(chemin)}"
        )

    # ── Identifier la colonne décision ─────────────────────────────────────
    col_decision        = None
    col_decision_finale = None
    col_num             = None   # colonne du N° étudiant
    col_note_finale     = None   # colonne note finale (fallback si décision vide)

    DECISION_FR = {"decision","décision","قرار"}
    DECISION_FIN= {"decisionfinal","décisionfinal","décisionfinale","decisionfinale",
                   "القرارالنهائي","القرار\nالنهائي","القرار النهائي"}

    for idx, h in enumerate(headers):
        hn = h.lower().replace(" ","").replace("\n","")
        h_raw = h.strip()

        # N° (colonne numéro étudiant)
        if h_raw in ("N°","NO","No","N","رقم") or hn in ("n°","no","رقم"):
            col_num = idx

        # Décision finale (priorité) — FR ou AR
        if (hn in DECISION_FIN or "القرار النهائي" in h
                or "décision finale" in h.lower()
                or "decision finale" in h.lower()
                or re.sub(r"[\s\n]","",h.lower()) in DECISION_FIN):
            col_decision_finale = idx
        # Décision initiale
        elif hn in DECISION_FR or "القرار" in h:
            if col_decision is None:
                col_decision = idx

        # Note finale (pour fallback quand toutes les décisions sont vides)
        if col_note_finale is None and any(k in hn for k in (
            "notefinal","notefinale","العلامةالنهائية","العلامةالنهائي",
            "moyennefinale","moyennefinal","notefinale","العلامةالأخيرة"
        )):
            col_note_finale = idx

    # Si pas de col_num détecté → colonne 0 par défaut (FR) ou dernière (AR)
    if col_num is None:
        col_num = len(headers)-1 if langue == "ar" else 0

    # Décision à utiliser : finale si disponible, sinon initiale
    col_dec_active = col_decision_finale if col_decision_finale is not None else col_decision

    # ── Lire les lignes étudiants ──────────────────────────────────────────
    df_all  = []
    df_ratt = []

    for i in range(header_row_idx + 1, len(rows)):
        row = rows[i]
        if all(v is None for v in row):
            continue
        # Vérifier que la cellule N° contient un entier
        try:
            num_val = row[col_num]
            if num_val is None: continue
            int(float(str(num_val)))
        except (TypeError, ValueError):
            continue

        row_dict = {"_row_idx": i, "_values": list(row)}
        df_all.append(row_dict)

        # ── Déterminer si non validé ───────────────────────────────────────
        # Lire les deux colonnes décision
        dec_init  = ""
        dec_finale = ""
        if col_decision is not None and col_decision < len(row):
            dec_init  = str(row[col_decision] or "").strip()
        if col_decision_finale is not None and col_decision_finale < len(row):
            dec_finale = str(row[col_decision_finale] or "").strip()

        # Choisir la décision la plus explicite :
        # Si dec_finale est une abréviation courte (م / غ م / V / NV ≤ 3 chars)
        # et que dec_init est plus longue et reconnue → préférer dec_init
        def _decision_active(d_finale, d_init):
            if d_finale:
                return d_finale
            if d_init:
                return d_init
            return ""

        dec = _decision_active(dec_finale, dec_init)
        # Fallback supplémentaire : si dec_finale est inconnue mais dec_init connue
        if dec_finale and dec_init and not _est_valide(dec_finale) and _est_valide(dec_init):
            dec = dec_init
        if dec_finale and dec_init and _est_valide(dec_finale) and not _est_valide(dec_init):
            dec = dec_finale

        if dec:
            est_ratt = not _est_valide(dec)
        else:
            # Décision vide (formules non calculées par openpyxl)
            # → fallback sur la note finale : non validé si note < 10
            est_ratt = False
            if col_note_finale is not None and col_note_finale < len(row):
                note = row[col_note_finale]
                try:
                    if note is not None and float(note) < 10:
                        est_ratt = True
                except (TypeError, ValueError):
                    pass
            # Si pas de colonne note finale non plus, chercher toute colonne
            # numérique qui ressemble à une note (valeur entre 0 et 20)
            if not est_ratt and col_note_finale is None:
                for ci, v in enumerate(row):
                    if ci == col_num:
                        continue
                    try:
                        f = float(v)
                        if 0 <= f < 10:
                            est_ratt = True
                            break
                    except (TypeError, ValueError):
                        pass

        if est_ratt:
            df_ratt.append(row_dict)

    if not module:
        module = os.path.splitext(os.path.basename(chemin))[0]

    # Nettoyer le nom du module des préfixes parasites
    for prefix in ["Module :", "Module:", "الوحدة :", "الوحدة:"]:
        if module.startswith(prefix):
            module = module[len(prefix):].strip()

    return {
        "module":               module,
        "filiere":              filiere,
        "annee":                annee,
        "langue":               langue,
        "header_row_idx":       header_row_idx,
        "headers":              headers,
        "col_num":              col_num,
        "col_decision":         col_decision,
        "col_decision_finale":  col_decision_finale,
        "col_dec_active":       col_dec_active,
        "col_note_finale":      col_note_finale,
        "df_all":               df_all,
        "df_ratt":              df_ratt,
        "nb_total":             len(df_all),
        "chemin":               chemin,
        "sheet_name":           ws.title,
        "wb":                   wb,           # valeurs calculées (pour lecture)
        "wb_formulas":          wb_formulas,  # formules originales (pour export)
    }


# ─── Génération ───────────────────────────────────────────────────────────────

def traiter_fiches(modules_info: list, dossier_sortie: str) -> tuple:
    os.makedirs(dossier_sortie, exist_ok=True)

    # ── Construire un nom significatif ─────────────────────────────────────
    # Année : prendre la première renseignée
    annees = [m.get("annee", "") for m in modules_info if m.get("annee", "")]
    annee_label = annees[0] if annees else "rattrapage"

    # Semestres : déduire les semestres distincts (S1, S2...)
    semestres = sorted({
        m.get("semestre", "")
        for m in modules_info
        if m.get("semestre", "")
    })
    sem_label = "_".join(semestres) if semestres else ""

    # Modules : liste des noms (max 3 pour éviter un nom trop long)
    noms = [m.get("module", "") or "" for m in modules_info if m.get("module", "")]
    if len(noms) <= 3:
        mods_label = "_".join(noms)
    else:
        mods_label = "_".join(noms[:3]) + f"_et_{len(noms)-3}_autres"

    # Nettoyer : remplacer les espaces et caractères spéciaux
    def _slug(s):
        return re.sub(r"[^\w\-]", "_", s).strip("_")

    annee_s = _slug(annee_label)
    sem_s   = _slug(sem_label)
    mods_s  = _slug(mods_label)

    # Construire le préfixe : Rattrapage_S1_2025-2026  ou  Rattrapage_2025-2026
    parts = ["Rattrapage"]
    if sem_s:
        parts.append(sem_s)
    if annee_s:
        parts.append(annee_s)
    prefixe = "_".join(parts)

    # Nom complet : prefixe + modules (tronqué à 100 chars)
    if mods_s:
        base = f"{prefixe}_{mods_s}"[:100]
    else:
        base = prefixe

    f1 = os.path.join(dossier_sortie, f"{base}_fiches.xlsx")
    f2 = os.path.join(dossier_sortie, f"{base}_conflits.xlsx")

    _generer_fiches(modules_info, f1)
    _generer_recap(modules_info, f2)
    return f1, f2


# ── Fichier 1 : fiches originales avec seulement les non validés ──────────────

def _generer_fiches(modules_info: list, chemin: str):
    wb_out = Workbook()
    wb_out.remove(wb_out.active)

    for info in modules_info:
        wb_src = info["wb_formulas"]   # workbook avec formules originales
        ws_src = wb_src.active
        module = info["module"]
        is_ar  = (info["langue"] == "ar")

        sheet_name = module[:31]
        existing   = [s.title for s in wb_out.worksheets]
        if sheet_name in existing:
            sheet_name = sheet_name[:28] + f"_{len(existing)}"

        ws_dst = wb_out.create_sheet(title=sheet_name)
        if is_ar:
            ws_dst.sheet_view.rightToLeft = True

        # Copier dimensions
        for col_letter, col_dim in ws_src.column_dimensions.items():
            ws_dst.column_dimensions[col_letter].width = col_dim.width or 12
        for row_idx, row_dim in ws_src.row_dimensions.items():
            if row_dim.height:
                ws_dst.row_dimensions[row_idx].height = row_dim.height

        header_row_1based = info["header_row_idx"] + 1

        # Copier en-têtes (toutes les lignes jusqu'à et y compris la ligne header)
        for r in range(1, header_row_1based + 1):
            for c in range(1, ws_src.max_column + 1):
                src_cell = ws_src.cell(row=r, column=c)
                dst_cell = ws_dst.cell(row=r, column=c, value=src_cell.value)
                _copier_style(src_cell, dst_cell)

        # Copier merged cells de l'en-tête
        for merged in ws_src.merged_cells.ranges:
            if merged.min_row <= header_row_1based:
                try:
                    ws_dst.merge_cells(str(merged))
                except Exception:
                    pass

        # Écrire les étudiants non validés avec renumérotation
        dst_row = header_row_1based + 1
        num     = 1
        col_num = info["col_num"]

        for etu in info["df_ratt"]:
            src_r  = etu["_row_idx"] + 1      # ligne source 1-based
            offset = dst_row - src_r           # décalage à appliquer aux formules
            for c in range(1, ws_src.max_column + 1):
                src_cell = ws_src.cell(row=src_r, column=c)
                val = src_cell.value
                # Renuméroter N°
                if (c - 1) == col_num:
                    val = num
                # Ajuster les références de ligne dans les formules
                elif isinstance(val, str) and val.startswith("=") and offset != 0:
                    val = _ajuster_formule(val, offset)
                dst_cell = ws_dst.cell(row=dst_row, column=c, value=val)
                _copier_style(src_cell, dst_cell)
            dst_row += 1
            num     += 1

        # Ligne total
        nb_ratt   = len(info["df_ratt"])
        nb_cols   = ws_src.max_column
        total_txt = (f"المجموع : {nb_ratt} طالب(ة)"
                     if is_ar else f"Total : {nb_ratt} étudiant(s) en rattrapage")
        # Unmerger dst_row au cas où des merged cells copiées l'occupent
        to_remove = [str(m) for m in ws_dst.merged_cells.ranges
                     if m.min_row <= dst_row <= m.max_row]
        for m in to_remove:
            try: ws_dst.unmerge_cells(m)
            except: pass
        ct = ws_dst.cell(row=dst_row, column=1, value=total_txt)
        if nb_cols > 1:
            try:
                ws_dst.merge_cells(start_row=dst_row, start_column=1,
                                   end_row=dst_row,   end_column=nb_cols)
            except Exception:
                pass
        ct.font      = Font(name="Arial", bold=True, size=9, color=BLANC)
        ct.fill      = _f(BLEU_HDR)
        ct.alignment = Alignment(horizontal="center", vertical="center")
        ct.border    = _b()
        ws_dst.row_dimensions[dst_row].height = 16

    wb_out.save(chemin)


def _copier_style(src, dst):
    try:
        if src.font:      dst.font      = copy.copy(src.font)
    except: pass
    try:
        if src.fill and src.fill.fill_type not in (None,"none"):
            dst.fill = copy.copy(src.fill)
    except: pass
    try:
        if src.alignment: dst.alignment = copy.copy(src.alignment)
    except: pass
    try:
        if src.border:    dst.border    = copy.copy(src.border)
    except: pass
    try:
        if src.number_format: dst.number_format = src.number_format
    except: pass


def _ajuster_formule(formule: str, offset: int) -> str:
    """
    Décale les références de lignes relatives dans une formule Excel.
    Ex : =IF(AND(F8>0,G8>0),...) avec offset=+2 → =IF(AND(F10>0,G10>0),...)
    Les références absolues ($F$8) ne sont PAS modifiées.
    """
    if offset == 0:
        return formule

    def replacer(match):
        col_part = match.group(1)   # ex: "F"  ou "$F"
        row_part = match.group(2)   # ex: "8"
        # Si la colonne est absolue ($F$8) → on garde tel quel
        if col_part.startswith("$") and match.group(0).count("$") == 2:
            return match.group(0)
        # Si la ligne est absolue ($8) → on garde tel quel
        if match.group(0).count("$") >= 1 and not col_part.startswith("$"):
            return match.group(0)
        new_row = int(row_part) + offset
        return f"{col_part}{new_row}"

    # Pattern : colonne (lettre optionnellement préfixée $) + ligne (chiffres)
    # Ne touche pas aux références absolues de ligne ($F$8)
    adjusted = re.sub(
        r'(\$?[A-Za-z]{1,3})(\d+)',
        replacer,
        formule
    )
    return adjusted


# ── Fichier 2 : récap + matrice + planning ────────────────────────────────────

def _generer_recap(modules_info: list, chemin: str):
    wb = Workbook()
    wb.remove(wb.active)
    for info in modules_info:
        _feuille_recap_module(wb, info)
    _feuille_matrice(wb, modules_info)
    _feuille_planning(wb, modules_info)
    wb.save(chemin)


def _feuille_recap_module(wb: Workbook, info: dict):
    module  = info["module"]
    is_ar   = (info["langue"] == "ar")
    headers = info["headers"]
    nb_cols = max(len([h for h in headers if h]), 6)

    sheet_name = ("R_" + module)[:31]
    existing   = [s.title for s in wb.worksheets]
    if sheet_name in existing:
        sheet_name = sheet_name[:28] + f"_{len(existing)}"

    ws = wb.create_sheet(title=sheet_name)
    if is_ar:
        ws.sheet_view.rightToLeft = True

    nb_ratt  = len(info["df_ratt"])
    nb_total = info["nb_total"]

    # Titre
    titre = f"{info['filiere']}  —  {info['annee']}  —  {module}"
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=nb_cols)
    c = ws.cell(1, 1, titre)
    c.font      = Font(name="Arial", bold=True, size=10, color=BLANC)
    c.fill      = _f(BLEU_DARK)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border    = _b()
    ws.row_dimensions[1].height = 22

    # Sous-titre
    sous = (f"{'قائمة الاستدراك' if is_ar else 'Liste de rattrapage'}  —  "
            f"{nb_ratt} {'طالب(ة)' if is_ar else 'étudiant(s)'} / {nb_total}")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=nb_cols)
    c2 = ws.cell(2, 1, sous)
    c2.font      = Font(name="Arial", size=9, color=BLANC)
    c2.fill      = _f(BLEU_HDR)
    c2.alignment = Alignment(horizontal="center", vertical="center")
    c2.border    = _b()
    ws.row_dimensions[2].height = 16

    # En-têtes
    for ci, h in enumerate(headers, start=1):
        if ci > nb_cols: break
        c = ws.cell(3, ci, h.replace("\n"," "))
        c.font      = Font(name="Arial", bold=True, size=9, color=BLANC)
        c.fill      = _f(BLEU_DARK)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border    = _b()
    ws.row_dimensions[3].height = 18

    if nb_ratt == 0:
        ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=nb_cols)
        msg = "لا يوجد طالب مستدرك." if is_ar else "Aucun étudiant en rattrapage."
        c = ws.cell(4, 1, msg)
        c.font      = Font(name="Arial", size=9, color="555555")
        c.fill      = _f(GRIS_BG)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border    = _b()
    else:
        col_dec   = info["col_dec_active"]
        for i, etu in enumerate(info["df_ratt"]):
            r   = 4 + i
            bg  = GRIS_BG if i % 2 == 0 else BLANC
            vals = etu["_values"]

            for ci in range(1, nb_cols + 1):
                val     = vals[ci-1] if (ci-1) < len(vals) else None
                cell_bg = bg
                cell_tc = "1C2B3A"
                bold    = False

                if col_dec is not None and (ci-1) == col_dec and val is not None:
                    dec_str = str(val).strip()
                    if dec_str and not _est_valide(dec_str):
                        cell_bg = ROUGE_BG
                        cell_tc = ROUGE_TXT
                        bold    = True

                # Renuméroter N°
                if (ci-1) == info["col_num"]:
                    val = i + 1

                c = ws.cell(r, ci, val)
                c.font      = Font(name="Arial", size=9, color=cell_tc, bold=bold)
                c.fill      = _f(cell_bg)
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.border    = _b()
            ws.row_dimensions[r].height = 14

        r_tot = 4 + nb_ratt
        ws.merge_cells(start_row=r_tot, start_column=1, end_row=r_tot, end_column=nb_cols)
        txt = (f"المجموع : {nb_ratt} طالب(ة)"
               if is_ar else f"Total : {nb_ratt} étudiant(s)")
        c = ws.cell(r_tot, 1, txt)
        c.font      = Font(name="Arial", bold=True, size=9, color=BLANC)
        c.fill      = _f(BLEU_HDR)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border    = _b()
        ws.row_dimensions[r_tot].height = 15

    # Largeurs colonnes
    for ci, h in enumerate(headers, start=1):
        if ci > nb_cols: break
        ltr = get_column_letter(ci)
        hn  = h.lower()
        if any(k in hn for k in ["nom","اسم","prenom","prénom"]):
            ws.column_dimensions[ltr].width = 22
        elif any(k in hn for k in ["cin","cne","massar","وطني","بطاقة"]):
            ws.column_dimensions[ltr].width = 16
        elif h.strip() in ("N°","NO","رقم",""):
            ws.column_dimensions[ltr].width = 5
        else:
            ws.column_dimensions[ltr].width = 13


def _build_sets(modules_info):
    """Construire les sets d'identifiants d'étudiants (CIN/Massar) par module."""
    sets = {}
    for info in modules_info:
        cles  = set()
        col_n = info["col_num"]
        for etu in info["df_ratt"]:
            vals = etu["_values"]
            # Chercher une valeur non nulle autre que le N°
            for ci in range(len(vals)):
                if ci == col_n: continue
                v = str(vals[ci] or "").strip()
                if v and v.lower() not in ("none","nan",""):
                    cles.add(v)
                    break
        sets[info["module"]] = cles
    return sets


def _feuille_matrice(wb: Workbook, modules_info: list):
    ws = wb.create_sheet(title="Matrice des conflits")
    modules = [m["module"] for m in modules_info]
    n       = len(modules)
    if n == 0: return

    sets    = _build_sets(modules_info)
    matrice = {
        mi: {mj: (len(sets[mi]) if mi==mj else len(sets[mi]&sets[mj]))
             for mj in modules}
        for mi in modules
    }
    max_conf = max(
        (matrice[mi][mj] for mi in modules for mj in modules if mi!=mj),
        default=1
    ) or 1

    nb = n + 2

    # Titre
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=nb)
    c = ws.cell(1, 1, "Matrice des conflits — Rattrapage")
    c.font      = Font(name="Arial", bold=True, size=11, color=BLANC)
    c.fill      = _f(BLEU_DARK)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border    = _b()
    ws.row_dimensions[1].height = 22

    legende = ("Diagonale = effectif  |  Cellule(i,j) = étudiants communs  "
               "|  Vert = pas de conflit  |  Orange/Rouge = conflit")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=nb)
    c2 = ws.cell(2, 1, legende)
    c2.font      = Font(name="Arial", size=8, color=BLANC)
    c2.fill      = _f(BLEU_HDR)
    c2.alignment = Alignment(horizontal="center", vertical="center")
    c2.border    = _b()
    ws.row_dimensions[2].height = 14

    # En-têtes colonnes
    for j, label in enumerate(["N°", "Module"] + [f"M{j+1}" for j in range(n)], start=1):
        c = ws.cell(3, j, label)
        c.font      = Font(bold=True, color=BLANC, size=8)
        c.fill      = _f(BLEU_DARK if j <= 2 else BLEU_HDR)
        c.alignment = Alignment(horizontal="center")
        c.border    = _b()
        if j > 2:
            ws.column_dimensions[get_column_letter(j)].width = 6
    ws.row_dimensions[3].height = 14

    # Corps
    for i, mod_i in enumerate(modules):
        r = 4 + i
        c1 = ws.cell(r, 1, f"M{i+1}")
        c1.font = Font(bold=True, color=BLANC, size=8)
        c1.fill = _f(BLEU_DARK); c1.alignment = Alignment(horizontal="center")
        c1.border = _b()

        c2 = ws.cell(r, 2, mod_i)
        c2.font = Font(size=8); c2.fill = _f(GRIS_BG)
        c2.alignment = Alignment(horizontal="left"); c2.border = _b()

        for j, mod_j in enumerate(modules):
            col = j + 3
            val = matrice[mod_i][mod_j]
            if mod_i == mod_j:
                bg, tc, bold = BLEU_HDR, BLANC, True
            elif val == 0:
                bg, tc, bold = VERT_BG, VERT_TXT, False
            else:
                ratio = val / max_conf
                bg, tc, bold = (ROUGE_BG, ROUGE_TXT, True) if ratio >= 0.5 \
                               else (ORANGE_BG, ORANGE_TXT, True)
            c = ws.cell(r, col, int(val) if val > 0 else "")
            c.font      = Font(size=9, color=tc, bold=bold)
            c.fill      = _f(bg)
            c.alignment = Alignment(horizontal="center")
            c.border    = _b()
        ws.row_dimensions[r].height = 14

    # Légende modules
    r_leg = 4 + n + 1
    ws.merge_cells(start_row=r_leg, start_column=1, end_row=r_leg, end_column=nb)
    c = ws.cell(r_leg, 1, "Légende des modules")
    c.font = Font(bold=True, color=BLANC, size=9)
    c.fill = _f(BLEU_DARK)
    c.alignment = Alignment(horizontal="center"); c.border = _b()
    ws.row_dimensions[r_leg].height = 14

    for i, mod in enumerate(modules):
        r  = r_leg + 1 + i
        bg = GRIS_BG if i % 2 == 0 else BLANC
        c1 = ws.cell(r, 1, f"M{i+1}")
        c1.font = Font(bold=True, color=BLANC, size=8)
        c1.fill = _f(BLEU_HDR)
        c1.alignment = Alignment(horizontal="center"); c1.border = _b()
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=nb)
        c2 = ws.cell(r, 2, mod)
        c2.font = Font(size=8); c2.fill = _f(bg)
        c2.alignment = Alignment(horizontal="left"); c2.border = _b()
        ws.row_dimensions[r].height = 13

    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 38


def _feuille_planning(wb: Workbook, modules_info: list):
    ws = wb.create_sheet(title="Planning suggéré")
    modules = [m["module"] for m in modules_info]
    n       = len(modules)
    if n == 0: return

    sets    = _build_sets(modules_info)
    matrice = {
        mi: {mj: len(sets[mi]&sets[mj]) for mj in modules}
        for mi in modules
    }

    # Greedy graph coloring
    groupes: dict = {}
    for mod in modules:
        voisins = {groupes[m] for m in modules
                   if m != mod and m in groupes and matrice[mod][m] > 0}
        g = 0
        while g in voisins: g += 1
        groupes[mod] = g

    # Titre
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=4)
    c = ws.cell(1, 1, "Planning suggéré — Rattrapage")
    c.font = Font(name="Arial", bold=True, size=11, color=BLANC)
    c.fill = _f(BLEU_DARK)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border = _b()
    ws.row_dimensions[1].height = 22

    expl = ("Les modules du même groupe peuvent se dérouler simultanément "
            "(aucun étudiant commun). Des groupes différents → créneaux séparés.")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=4)
    c2 = ws.cell(2, 1, expl)
    c2.font = Font(name="Arial", size=8, color=BLANC)
    c2.fill = _f(BLEU_HDR)
    c2.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c2.border = _b()
    ws.row_dimensions[2].height = 26

    # En-têtes
    for ci, h in enumerate(["Groupe","Module","Nb rattrapants","Conflits avec"],start=1):
        c = ws.cell(3, ci, h)
        c.font      = Font(bold=True, color=BLANC, size=9)
        c.fill      = _f(BLEU_DARK)
        c.alignment = Alignment(horizontal="center")
        c.border    = _b()
    ws.row_dimensions[3].height = 15

    modules_tries = sorted(modules, key=lambda m: groupes[m])
    for i, mod in enumerate(modules_tries):
        r  = 4 + i
        g  = groupes[mod]
        bg = GROUPE_COLORS[g % len(GROUPE_COLORS)]
        nb = len(sets[mod])
        conflits = [f"M{modules.index(m)+1}" for m in modules
                    if m != mod and matrice[mod][m] > 0]
        conflits_str = ", ".join(conflits) if conflits else "Aucun"

        for ci, val in enumerate([f"Groupe {g+1}", mod, int(nb), conflits_str], start=1):
            c = ws.cell(r, ci, val)
            c.font      = Font(size=9, color="1C2B3A", bold=(ci in (1,3)))
            c.fill      = _f(bg)
            c.alignment = Alignment(
                horizontal="left" if ci in (2,4) else "center",
                vertical="center"
            )
            c.border    = _b()
        ws.row_dimensions[r].height = 14

    r_sum = 4 + n + 1
    nb_grp = max(groupes.values()) + 1 if groupes else 0
    ws.merge_cells(start_row=r_sum, start_column=1, end_row=r_sum, end_column=4)
    c = ws.cell(r_sum, 1,
                f"→  {nb_grp} créneau(x) minimum nécessaire(s) pour éviter tous les conflits.")
    c.font      = Font(bold=True, color=BLANC, size=9)
    c.fill      = _f(BLEU_DARK)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border    = _b()
    ws.row_dimensions[r_sum].height = 16

    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 40
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 28
