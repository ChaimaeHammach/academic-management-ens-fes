

import pandas as pd
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ─── Palette neutre (minimaliste) ────────────────────────────────────────────
# En-têtes principaux : gris foncé fond, texte blanc
# En-têtes modules   : gris moyen fond, texte blanc
# Lignes données     : blanc / gris très clair alternés
# Résultats          : sans couleur de fond, texte noir — gras pour les notes

BLEU_DARK  = "4A4A4A"   # gris foncé → en-têtes principaux
BLEU_MID   = "7A7A7A"   # gris moyen → en-têtes modules
BLEU_LIGHT = "D0D0D0"   # gris clair (conservé pour compatibilité)
JAUNE      = "FFFFFF"
VERT       = "FFFFFF"; VERT_T  = "000000"
ORANGE     = "FFFFFF"; ORANGE_T= "000000"
ROUGE      = "FFFFFF"; ROUGE_T = "000000"
GRIS       = "F5F5F5"; BLANC   = "FFFFFF"

# Décisions prof considérées comme "validé"
DECISIONS_VALIDE = {"V", "VAR", "VALIDÉ", "VALIDE", "س.م", "س م"}


def _est_valide_row(row, mod: str) -> bool:
    """
    Détermine si un module est validé pour un étudiant.
    Priorité : flag Validé calculé par reader > décision prof > note >= 10.
    """
    flag = row.get(f"{mod}_Validé")
    if flag is not None and not _is_na(flag):
        return bool(flag)
    dec = str(row.get(f"{mod}_Decision", "")).strip().upper()
    if dec and dec not in ("NAN", "NONE", ""):
        return dec in DECISIONS_VALIDE
    note = row.get(f"{mod}_Note_Finale")
    if not _is_na(note):
        try:
            return float(note) >= 10
        except (TypeError, ValueError):
            pass
    return False


def _is_na(value) -> bool:
    if value is None:
        return True
    try:
        return pd.isna(value)
    except (TypeError, ValueError):
        return False


def _b(style="thin", color="AAAAAA"):
    s = Side(style=style, color=color)
    return Border(left=s, right=s, top=s, bottom=s)


def _f(hex_color):
    return PatternFill("solid", fgColor=hex_color)


def _safe(value):
    """Convertit toute valeur NA/nan/NaT en chaîne vide."""
    if _is_na(value):
        return ""
    return value


def _cell(ws, row, col, value="", bold=False, size=9, color=BLANC,
          fill=BLEU_DARK, align="center", wrap=False, border=True):
    cell_coord = f"{get_column_letter(col)}{row}"
    for merged in list(ws.merged_cells.ranges):
        if cell_coord in merged:
            ws.unmerge_cells(str(merged))
            break
    value = _safe(value)
    c = ws.cell(row, col, value)
    c.font      = Font(name="Arial", bold=bold, size=size, color=color)
    c.fill      = _f(fill)
    c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    if border:
        c.border = _b()
    return c


def _merge(ws, r1, c1, r2, c2, value="", bold=True, size=9,
           color=BLANC, fill=BLEU_DARK, align="center", wrap=True):
    ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
    c = ws.cell(r1, c1, value)
    c.font      = Font(name="Arial", bold=bold, size=size, color=color)
    c.fill      = _f(fill)
    c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    c.border    = _b()
    return c


def _fmt_note(value):
    """Formate une note numérique, retourne '—' si absente."""
    if _is_na(value):
        return "—"
    try:
        return round(float(value), 2)
    except (TypeError, ValueError):
        return "—"


def _fmt_dec_module(value):
    """
    Formate la décision d'un module pour l'affichage.
    V/VAR → Validé, NV/NE → Non Validé, ABS → Absent, sinon tel quel.
    """
    v = str(_safe(value)).strip().upper()
    MAP = {
        "V":      "V",
        "VAR":    "VAR",
        "VALIDÉ": "V",
        "VALIDE": "V",
        "NV":     "NV",
        "NE":     "NE",
        "ABS":    "Abs",
        "":       "",
        "NAN":    "",
        "NONE":   "",
    }
    return MAP.get(v, v)


# ─── Labels bilingues ─────────────────────────────────────────────────────────

def _labels(langue: str, label_si: str, label_sp: str, ecole: str,
            filiere: str, annee: str):
    is_ar = (langue == "ar")

    if is_ar:
        return {
            "entete": (
                f"{ecole}    |    السنة الجامعية : {annee}    |    "
                f"محضر المداولات السنوية ({label_si}+{label_sp})    |    الشعبة : {filiere}"
            ),
            "identite":      "هوية الطالب(ة)",
            "cin":           "الرقم الوطني",
            "massar":        "رقم المسار",
            "nom":           "الاسم",
            "prenom":        "النسب",
            "moy":           "المعدل",
            "av":            "سنة",
            "dec":           "القرار",
            "recap_si":      f"الفصل {label_si}",
            "recap_sp":      f"الفصل {label_sp}",
            "moy_si":        f"م.{label_si}",
            "dec_si":        f"ق.{label_si}",
            "av_si":         f"سنة",
            "moy_sp":        f"م.{label_sp}",
            "dec_sp":        f"ق.{label_sp}",
            "av_sp":         f"سنة",
            "moy_gen":       "المعدل العام",
            "mods_valides":  "الوحدات المكتسبة",
            "credits":       "الأرصدة",
            "decision":      "القرار",
            "rattrapages":   "الاستدراك",
            "mention":       "الميزة",
            "valide":        "مكتسب",
            "non_valide":    "غير مكتسب",
            "label_pv":      "محضر المداولات السنوية",
            "label_annee":   f"السنة الجامعية : {annee}",
            "label_filiere": f"الشعبة : {filiere}",
            "label_sem":     f"({label_si}+{label_sp})",
            "admis":         "مقبول",
            "rachat":        "استدراك ممكن",
            "ajourne":       "راسب",
            "tb":            "ممتاز",
            "b":             "جيد جداً",
            "ab":            "جيد",
            "p":             "مقبول",
        }
    else:
        return {
            "entete": (
                f"{ecole}    |    Année Universitaire : {annee}    |    "
                f"PV Annuel ({label_si}+{label_sp})    |    Filière : {filiere}"
            ),
            "identite":      "Identité de l'étudiant(e)",
            "cin":           "CIN",
            "massar":        "Massar",
            "nom":           "Nom",
            "prenom":        "Prénom",
            "moy":           "Moy.",
            "av":            "AV",
            "dec":           "Déc.",
            "recap_si":      f"Semestre {label_si}",
            "recap_sp":      f"Semestre {label_sp}",
            "moy_si":        f"Moy.{label_si}",
            "dec_si":        f"Déc.{label_si}",
            "av_si":         f"AV {label_si}",
            "moy_sp":        f"Moy.{label_sp}",
            "dec_sp":        f"Déc.{label_sp}",
            "av_sp":         f"AV {label_sp}",
            "moy_gen":       "Moyenne\nGénérale",
            "mods_valides":  "Modules\nValidés",
            "credits":       "Crédits",
            "decision":      "Décision",
            "rattrapages":   "Rattrapages",
            "mention":       "Mention",
            "valide":        "Validé",
            "non_valide":    "Non Validé",
            "label_pv":      "Procès-Verbal Annuel",
            "label_annee":   f"Année Universitaire : {annee}",
            "label_filiere": f"Filière : {filiere}",
            "label_sem":     f"({label_si}+{label_sp})",
            "admis":         "Admis",
            "rachat":        "Rachat possible",
            "ajourne":       "Ajourné",
            "tb":            "Très Bien",
            "b":             "Bien",
            "ab":            "Assez Bien",
            "p":             "Passable",
        }


# ─── Export Excel ─────────────────────────────────────────────────────────────

def exporter_excel(
    df_pv: pd.DataFrame,
    modules_si: list[str],
    modules_sp: list[str],
    label_si: str = "S5",
    label_sp: str = "S6",
    ecole:   str  = "École Normale Supérieure de Fès",
    filiere: str  = "",
    annee:   str  = "",
    chemin_sortie: str = "output/pv_final.xlsx",
    langue:  str  = "fr"
) -> str:

    os.makedirs(os.path.dirname(chemin_sortie) if os.path.dirname(chemin_sortie) else ".", exist_ok=True)

    L     = _labels(langue, label_si, label_sp, ecole, filiere, annee)
    is_ar = (langue == "ar")

    wb = Workbook()
    ws = wb.active
    ws.title = "PV Final"

    if is_ar:
        ws.sheet_view.rightToLeft = True

    n_si     = len(modules_si)
    n_sp     = len(modules_sp)
    recap_si = 3 if n_si > 0 else 0
    recap_sp = 3 if n_sp > 0 else 0
    n_total  = 4 + n_si*3 + recap_si + n_sp*3 + recap_sp + 6

    # ── Ligne 1 : En-tête ──
    _merge(ws, 1, 1, 1, n_total,
           value=L["entete"],
           size=11, bold=True, fill=BLEU_DARK, color=BLANC)
    ws.row_dimensions[1].height = 28

    # ── Ligne 2 : groupes de colonnes ──
    col = 1
    _merge(ws, 2, col, 2, col+3, L["identite"], fill=BLEU_DARK, size=9)
    col += 4

    for mod in modules_si:
        _merge(ws, 2, col, 2, col+2, mod, fill=BLEU_MID, size=8, wrap=True)
        col += 3

    if modules_si:
        _merge(ws, 2, col, 2, col+2, L["recap_si"], fill=BLEU_DARK, size=9)
        col += 3

    for mod in modules_sp:
        _merge(ws, 2, col, 2, col+2, mod, fill=BLEU_MID, size=8, wrap=True)
        col += 3

    if modules_sp:
        _merge(ws, 2, col, 2, col+2, L["recap_sp"], fill=BLEU_DARK, size=9)
        col += 3

    finales = [
        L["moy_gen"], L["mods_valides"], L["credits"],
        L["decision"], L["rattrapages"], L["mention"]
    ]
    for i, h in enumerate(finales):
        _merge(ws, 2, col+i, 3, col+i, h, fill=BLEU_DARK, size=9, wrap=True)

    ws.row_dimensions[2].height = 30

    # ── Ligne 3 : sous-colonnes ──
    col = 1
    for h in [L["cin"], L["massar"], L["nom"], L["prenom"]]:
        _cell(ws, 3, col, h, bold=True, fill=BLEU_DARK)
        col += 1

    for _ in modules_si:
        for h in [L["moy"], L["av"], L["dec"]]:
            _cell(ws, 3, col, h, bold=True, fill=BLEU_MID)
            col += 1

    if modules_si:
        for h in [L["moy_si"], L["dec_si"], L["av_si"]]:
            _cell(ws, 3, col, h, bold=True, fill=BLEU_DARK)
            col += 1

    for _ in modules_sp:
        for h in [L["moy"], L["av"], L["dec"]]:
            _cell(ws, 3, col, h, bold=True, fill=BLEU_MID)
            col += 1

    if modules_sp:
        for h in [L["moy_sp"], L["dec_sp"], L["av_sp"]]:
            _cell(ws, 3, col, h, bold=True, fill=BLEU_DARK)
            col += 1

    ws.row_dimensions[3].height = 18

    # ── Données ──
    for i, (_, row) in enumerate(df_pv.iterrows()):
        r   = 4 + i
        bg  = GRIS if i % 2 == 0 else BLANC
        col = 1

        # Identité
        for champ in ["CIN", "Massar", "Nom", "Prénom"]:
            _cell(ws, r, col, _safe(row.get(champ, "")),
                  fill=bg, color="000000", align="left")
            col += 1

        # Modules SI
        for mod in modules_si:
            note_disp = _fmt_note(row.get(f"{mod}_Note_Finale"))
            # سنة depuis la source
            annee_val = _safe(row.get(f"{mod}_Annee_Val", ""))
            if not annee_val or annee_val in ("nan", "None", ""):
                annee_val = _safe(row.get(f"{mod}_AV", ""))
            # القرار النهائي depuis la source directement (م / غ م / ر)
            dec_raw = row.get(f"{mod}_Decision", "")
            dec_disp_mod = _safe(dec_raw) if dec_raw else ""
            valide    = _est_valide_row(row, mod)

            _cell(ws, r, col,   note_disp,    fill=bg, color="000000", bold=True)
            _cell(ws, r, col+1, annee_val,    fill=bg, color="000000")
            _cell(ws, r, col+2, dec_disp_mod, fill=bg, color="000000")
            col += 3

        # Récap SI
        if modules_si:
            moy_si_disp = _fmt_note(row.get(f"Moy_{label_si}"))
            dec_si_calc = _safe(row.get(f"Dec_{label_si}", ""))
            si_valide   = (dec_si_calc == L["valide"] or dec_si_calc == "Validé")
            if not si_valide and dec_si_calc not in (L["non_valide"], "Non Validé"):
                try:
                    si_valide = moy_si_disp != "—" and float(moy_si_disp) >= 10
                except Exception:
                    pass
            dec_si_val  = L["valide"] if si_valide else L["non_valide"]
            bg_si       = GRIS
            tc_si       = "000000"
            # سنة : prendre depuis le premier module SI non vide
            annee_si_val = ""
            for mod_si in modules_si:
                av_m = _safe(row.get(f"{mod_si}_Annee_Val", ""))
                if av_m and av_m not in ("nan", "None", ""):
                    annee_si_val = av_m
                    break
            if not annee_si_val:
                for mod_si in modules_si:
                    av_m = _safe(row.get(f"{mod_si}_AV", ""))
                    if av_m and av_m not in ("nan", "None", ""):
                        annee_si_val = av_m
                        break
            _cell(ws, r, col,   moy_si_disp, fill=bg_si, bold=True, color=tc_si)
            _cell(ws, r, col+1, dec_si_val,  fill=bg_si, color=tc_si)
            _cell(ws, r, col+2, annee_si_val, fill=bg_si, color=tc_si)
            col += 3

        # Modules SP
        for mod in modules_sp:
            note_disp = _fmt_note(row.get(f"{mod}_Note_Finale"))
            # سنة depuis la source
            annee_val = _safe(row.get(f"{mod}_Annee_Val", ""))
            if not annee_val or annee_val in ("nan", "None", ""):
                annee_val = _safe(row.get(f"{mod}_AV", ""))
            # القرار النهائي depuis la source directement (م / غ م / ر)
            dec_raw = row.get(f"{mod}_Decision", "")
            dec_disp_mod = _safe(dec_raw) if dec_raw else ""
            valide    = _est_valide_row(row, mod)

            _cell(ws, r, col,   note_disp,    fill=bg, color="000000", bold=True)
            _cell(ws, r, col+1, annee_val,    fill=bg, color="000000")
            _cell(ws, r, col+2, dec_disp_mod, fill=bg, color="000000")
            col += 3

        # Récap SP
        if modules_sp:
            moy_sp_disp = _fmt_note(row.get(f"Moy_{label_sp}"))
            dec_sp_calc = _safe(row.get(f"Dec_{label_sp}", ""))
            sp_valide   = (dec_sp_calc == L["valide"] or dec_sp_calc == "Validé")
            if not sp_valide and dec_sp_calc not in (L["non_valide"], "Non Validé"):
                try:
                    sp_valide = moy_sp_disp != "—" and float(moy_sp_disp) >= 10
                except Exception:
                    pass
            dec_sp_val  = L["valide"] if sp_valide else L["non_valide"]
            bg_sp       = GRIS
            tc_sp       = "000000"
            # سنة : prendre depuis le premier module SP non vide
            annee_sp_val = ""
            for mod_sp in modules_sp:
                av_m = _safe(row.get(f"{mod_sp}_Annee_Val", ""))
                if av_m and av_m not in ("nan", "None", ""):
                    annee_sp_val = av_m
                    break
            if not annee_sp_val:
                for mod_sp in modules_sp:
                    av_m = _safe(row.get(f"{mod_sp}_AV", ""))
                    if av_m and av_m not in ("nan", "None", ""):
                        annee_sp_val = av_m
                        break
            _cell(ws, r, col,   moy_sp_disp, fill=bg_sp, bold=True, color=tc_sp)
            _cell(ws, r, col+1, dec_sp_val,  fill=bg_sp, color=tc_sp)
            _cell(ws, r, col+2, annee_sp_val, fill=bg_sp)
            col += 3

        # Colonnes finales
        moy_ann     = row.get("Moyenne_Annuelle")
        dec         = _safe(row.get("Decision", ""))
        men         = _safe(row.get("Mention", ""))
        cred        = _safe(row.get("Credits_Totaux", 0))
        rattrap     = _safe(row.get("Rattrapages", ""))
        moy_ann_disp = _fmt_note(moy_ann)

        tous_modules = modules_si + modules_sp
        mods_valides = sum(1 for m in tous_modules if _est_valide_row(row, m))

        if is_ar:
            dec_map = {"Admis": L["admis"], "Rachat possible": L["rachat"], "Ajourné": L["ajourne"]}
            men_map = {"Très Bien": L["tb"], "Bien": L["b"], "Assez Bien": L["ab"], "Passable": L["p"],
                       L["tb"]: L["tb"], L["b"]: L["b"], L["ab"]: L["ab"], L["p"]: L["p"]}
            dec_disp = dec_map.get(str(dec), str(dec))
            men_disp = men_map.get(str(men), str(men))
        else:
            dec_disp = dec
            men_disp = men

        _cell(ws, r, col,   moy_ann_disp,                    fill=bg, bold=True, color="000000")
        _cell(ws, r, col+1, f"{mods_valides}/{len(tous_modules)}", fill=bg, color="000000")
        _cell(ws, r, col+2, cred,                                  fill=bg, color="000000")
        _cell(ws, r, col+3, dec_disp,                              fill=bg, bold=True, color="000000")
        _cell(ws, r, col+4, rattrap,                               fill=bg, wrap=True, size=8, color="000000")
        _cell(ws, r, col+5, men_disp,                              fill=bg, bold=True, color="000000")

        ws.row_dimensions[r].height = 15

    # ── Ajuster largeurs ──
    for col_cells in ws.columns:
        max_len = 0
        cl = get_column_letter(col_cells[0].column)
        for cell in col_cells:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[cl].width = min(max(max_len + 2, 6), 28)

    ws.freeze_panes = "E4"
    wb.save(chemin_sortie)
    print(f"PV Excel exporté : {chemin_sortie}")
    return chemin_sortie


# ─── Export PDF ───────────────────────────────────────────────────────────────

def exporter_pdf(
    df_pv: pd.DataFrame,
    modules_si: list[str],
    modules_sp: list[str],
    label_si: str = "S5",
    label_sp: str = "S6",
    ecole:    str = "École Normale Supérieure de Fès",
    filiere:  str = "",
    annee:    str = "",
    chemin_sortie: str = "output/pv_final.pdf",
    langue:   str = "fr"
) -> str:
    try:
        from reportlab.lib.pagesizes import A3, landscape
        from reportlab.lib import colors
        from reportlab.lib.units import cm
        from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.enums import TA_CENTER, TA_RIGHT
    except ImportError:
        raise ImportError("reportlab non installé : pip install reportlab")

    os.makedirs(os.path.dirname(chemin_sortie) if os.path.dirname(chemin_sortie) else ".", exist_ok=True)

    L     = _labels(langue, label_si, label_sp, ecole, filiere, annee)
    is_ar = (langue == "ar")

    doc = SimpleDocTemplate(chemin_sortie, pagesize=landscape(A3),
                            rightMargin=0.8*cm, leftMargin=0.8*cm,
                            topMargin=1*cm, bottomMargin=1*cm)

    styles = getSampleStyleSheet()
    align_titre = TA_RIGHT if is_ar else TA_CENTER
    titre_s = ParagraphStyle("t", fontSize=11, alignment=align_titre,
                              fontName="Helvetica-Bold", spaceAfter=6)
    sous_s  = ParagraphStyle("s", fontSize=9,  alignment=align_titre,
                              fontName="Helvetica", spaceAfter=10)

    elements = []
    elements.append(Paragraph(ecole, titre_s))
    elements.append(Paragraph(
        f"{L['label_annee']}  |  {L['label_pv']} {L['label_sem']}  |  {L['label_filiere']}",
        sous_s
    ))
    elements.append(Spacer(1, 0.3*cm))

    headers = [L["cin"], L["massar"], L["nom"], L["prenom"]]
    for mod in modules_si:
        headers += [f"{mod}\n{L['moy']}", f"{mod}\n{L['av']}", f"{mod}\n{L['dec']}"]
    if modules_si:
        headers += [L["moy_si"], L["dec_si"], L["av_si"]]
    for mod in modules_sp:
        headers += [f"{mod}\n{L['moy']}", f"{mod}\n{L['av']}", f"{mod}\n{L['dec']}"]
    if modules_sp:
        headers += [L["moy_sp"], L["dec_sp"], L["av_sp"]]
    headers += [L["moy_gen"], L["mods_valides"], L["credits"],
                L["decision"], L["rattrapages"], L["mention"]]

    data = [headers]

    def v(x):
        if _is_na(x): return ""
        if isinstance(x, float): return str(round(x, 2))
        return str(x)

    for _, row in df_pv.iterrows():
        line = [v(row.get("CIN","")), v(row.get("Massar","")),
                v(row.get("Nom","")),  v(row.get("Prénom",""))]

        for mod in modules_si:
            line += [
                v(row.get(f"{mod}_Note_Finale")),
                v(row.get(f"{mod}_AV", "")),
                _fmt_dec_module(row.get(f"{mod}_Decision", "")),
            ]

        if modules_si:
            msi    = row.get(f"Moy_{label_si}")
            si_ok  = not _is_na(msi) and float(msi) >= 10
            line  += [v(msi), L["valide"] if si_ok else L["non_valide"], ""]

        for mod in modules_sp:
            line += [
                v(row.get(f"{mod}_Note_Finale")),
                v(row.get(f"{mod}_AV", "")),
                _fmt_dec_module(row.get(f"{mod}_Decision", "")),
            ]

        if modules_sp:
            msp    = row.get(f"Moy_{label_sp}")
            sp_ok  = not _is_na(msp) and float(msp) >= 10
            line  += [v(msp), L["valide"] if sp_ok else L["non_valide"], ""]

        tous_modules = modules_si + modules_sp
        mods_v       = sum(1 for m in tous_modules if _est_valide_row(row, m))
        dec_raw      = row.get("Decision", "")
        men_raw      = row.get("Mention",  "")

        if is_ar:
            dec_map  = {"Admis": L["admis"], "Rachat possible": L["rachat"], "Ajourné": L["ajourne"]}
            men_map  = {"Très Bien": L["tb"], "Bien": L["b"], "Assez Bien": L["ab"], "Passable": L["p"]}
            dec_disp = dec_map.get(str(dec_raw), v(dec_raw))
            men_disp = men_map.get(str(men_raw), v(men_raw))
        else:
            dec_disp = v(dec_raw)
            men_disp = v(men_raw)

        line += [
            v(row.get("Moyenne_Annuelle")),
            f"{mods_v}/{len(tous_modules)}",
            v(row.get("Credits_Totaux", "")),
            dec_disp,
            v(row.get("Rattrapages", "")),
            men_disp,
        ]
        data.append(line)

    style = TableStyle([
        ("BACKGROUND",    (0,0), (-1,0),  colors.HexColor(f"#{BLEU_DARK}")),
        ("TEXTCOLOR",     (0,0), (-1,0),  colors.white),
        ("FONTNAME",      (0,0), (-1,0),  "Helvetica-Bold"),
        ("FONTSIZE",      (0,0), (-1,-1), 6),
        ("ALIGN",         (0,0), (-1,-1), "RIGHT" if is_ar else "CENTER"),
        ("VALIGN",        (0,0), (-1,-1), "MIDDLE"),
        ("ROWBACKGROUNDS",(0,1), (-1,-1), [colors.white, colors.HexColor(f"#{GRIS}")]),
        ("GRID",          (0,0), (-1,-1), 0.3, colors.HexColor("#CBD5E1")),
        ("TOPPADDING",    (0,0), (-1,-1), 2),
        ("BOTTOMPADDING", (0,0), (-1,-1), 2),
    ])

    table = Table(data, repeatRows=1)
    table.setStyle(style)
    elements.append(table)
    doc.build(elements)

    print(f"PV PDF exporté : {chemin_sortie}")
    return chemin_sortie