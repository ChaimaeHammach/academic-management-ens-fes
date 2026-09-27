
import re
import pandas as pd
from openpyxl import load_workbook


# ─── Utilitaires ──────────────────────────────────────────────────────────────

def _norm(s: str) -> str:
    """Normalise une chaîne : minuscules, sans espaces/tirets/points."""
    return re.sub(r"[\s\-_./·،]", "", str(s)).lower().strip()


def _is_na(value) -> bool:
    if value is None:
        return True
    try:
        return pd.isna(value)
    except (TypeError, ValueError):
        return False


def _safe_float(value):
    if _is_na(value):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _safe_str(value) -> str:
    if _is_na(value):
        return ""
    return str(value).strip()


# ─── Lecture d'un PV annuel Excel ─────────────────────────────────────────────

def _lire_en_tetes_pv(ws):
    """
    Lit les 4 premières lignes du PV exporté (format openpyxl).
    Construit un mapping colonne_index -> texte_entete en résolvant les fusions.
    Retourne (headers_row2, headers_row3) : tuples indexés 1-based.
    """
    # Résoudre les cellules fusionnées
    merge_map = {}
    for merged in ws.merged_cells.ranges:
        val = ws.cell(merged.min_row, merged.min_col).value
        for row in range(merged.min_row, merged.max_row + 1):
            for col in range(merged.min_col, merged.max_col + 1):
                merge_map[(row, col)] = val

    def cell_val(row, col):
        if (row, col) in merge_map:
            return merge_map[(row, col)]
        return ws.cell(row, col).value

    max_col = ws.max_column
    row2 = [cell_val(2, c) for c in range(1, max_col + 1)]
    row3 = [cell_val(3, c) for c in range(1, max_col + 1)]
    return row2, row3


def _trouver_colonnes_pv(ws, label_si: str, label_sp: str):
    """
    Identifie les numéros de colonnes (1-based) pour :
      - CIN, Massar, Nom, Prénom
      - Moy_{label_si}, Dec_{label_si}, AV_{label_si}
      - Moy_{label_sp}, Dec_{label_sp}, AV_{label_sp}
      - Decision (annuelle), Mention

    Stratégie : on combine row2 (groupes) + row3 (sous-colonnes) pour construire
    un header composite, puis on normalise pour matcher.
    """
    row2, row3 = _lire_en_tetes_pv(ws)
    max_col = len(row3)

    # Header composite : "groupe__souscolonne"
    composites = []
    for i in range(max_col):
        g = _norm(str(row2[i] or ""))
        s = _norm(str(row3[i] or ""))
        composites.append((g, s, g + "__" + s))

    def match_any(idx, *patterns):
        g, s, c = composites[idx]
        for p in patterns:
            pn = _norm(p)
            if pn in g or pn in s or pn in c:
                return True
        return False

    # Noms de labels normalisés
    si_n = _norm(label_si)
    sp_n = _norm(label_sp)

    cols = {
        "cin": None, "massar": None, "nom": None, "prenom": None,
        f"moy_{label_si}": None, f"dec_{label_si}": None, f"av_{label_si}": None,
        f"moy_{label_sp}": None, f"dec_{label_sp}": None, f"av_{label_sp}": None,
        "decision": None, "mention": None,
    }

    # ── Identité ──
    for i in range(max_col):
        if cols["cin"] is None and match_any(i, "cin", "ب-ت-و", "bto"):
            cols["cin"] = i + 1
        if cols["massar"] is None and match_any(i, "massar", "cne", "nummassar"):
            cols["massar"] = i + 1
        if cols["nom"] is None and match_any(i, "nom", "lastname", "النسب"):
            cols["nom"] = i + 1
        if cols["prenom"] is None and match_any(i, "prénom", "prenom", "firstname", "الاسم"):
            cols["prenom"] = i + 1

    # ── Récapitulatif semestriel ──
    # IMPORTANT : ce sont les colonnes RÉCAP (groupe « Semestre Sx »), PAS les
    # modules. Un module comme « LanguesS5 » contient « s5 » dans son groupe,
    # mais sa sous-colonne est juste « Moy. » (sans suffixe semestre).
    # Les colonnes récap portent le semestre DANS la sous-colonne :
    #   « Moy.S5 » → moys5,  « Déc.S5 » → décs5,  « AV S5 » → avs5
    # On exige donc le token semestre dans la sous-colonne (row3),
    # avec fallback sur un groupe « Semestre Sx ».

    MOY_KW = ("moy", "moyenne", "معدل", "م")
    DEC_KW = ("déc", "dec", "décision", "decision", "قرار", "ق")
    AV_KW  = ("av", "a/v", "année", "annee", "validé", "م.و")

    def _detecter_recap(sem_n, label):
        """Cherche les 3 colonnes récap (moy/déc/av) d'un semestre donné."""
        # Tier 1 : sous-colonne (row3) contient le token semestre
        for i in range(max_col):
            g, s, c = composites[i]
            if sem_n not in s:
                continue
            if cols[f"moy_{label}"] is None and any(s.startswith(k) or k in s for k in MOY_KW) \
               and not any(k in s for k in ("déc", "dec", "av")):
                cols[f"moy_{label}"] = i + 1
            elif cols[f"dec_{label}"] is None and any(k in s for k in DEC_KW):
                cols[f"dec_{label}"] = i + 1
            elif cols[f"av_{label}"] is None and any(s.startswith(k) or k in s for k in ("av", "a/v", "année", "annee")):
                cols[f"av_{label}"] = i + 1

        # Tier 2 : fallback — groupe « semestre Sx »
        if cols[f"moy_{label}"] is None:
            for i in range(max_col):
                g, s, c = composites[i]
                if "semestre" in g and sem_n in g:
                    if cols[f"moy_{label}"] is None and any(k in s for k in MOY_KW) \
                       and not any(k in s for k in ("déc", "dec")):
                        cols[f"moy_{label}"] = i + 1
                    elif cols[f"dec_{label}"] is None and any(k in s for k in DEC_KW):
                        cols[f"dec_{label}"] = i + 1
                    elif cols[f"av_{label}"] is None and any(s.startswith(k) for k in ("av",)):
                        cols[f"av_{label}"] = i + 1

    _detecter_recap(si_n, label_si)
    _detecter_recap(sp_n, label_sp)

    # Decision et Mention globales : cherche dans row2 ou row3 seul
    # (dernières colonnes, pas associées à un semestre)
    for i in range(max_col - 1, -1, -1):
        g, s, c = composites[i]
        # Decision globale
        if cols["decision"] is None:
            if any(x in g or x in s for x in ["decision", "décision", "قرار"]):
                # Ne pas confondre avec déc.S1 etc.
                if si_n not in g and si_n not in s and sp_n not in g and sp_n not in s:
                    cols["decision"] = i + 1
        # Mention
        if cols["mention"] is None:
            if any(x in g or x in s for x in ["mention", "ميزة", "ميزه"]):
                cols["mention"] = i + 1

    # Fallback : si decision non trouvée, chercher en avant
    if cols["decision"] is None:
        for i in range(max_col):
            g, s, c = composites[i]
            if any(x in g or x in s for x in ["decision", "décision", "قرار"]):
                if si_n not in g and si_n not in s and sp_n not in g and sp_n not in s:
                    cols["decision"] = i + 1
                    break

    return cols


def lire_pv_annuel(chemin_xlsx: str, label_si: str, label_sp: str) -> pd.DataFrame:
    """
    Lit un PV annuel Excel exporté par ce système et retourne un DataFrame avec :
      CIN, Massar, Nom, Prénom,
      Moy_{label_si}, Dec_{label_si}, AV_{label_si},
      Moy_{label_sp}, Dec_{label_sp}, AV_{label_sp},
      Decision_Annuelle, Mention_Annuelle

    Les données commencent à la ligne 4 (après 3 lignes d'en-tête).
    """
    wb = load_workbook(chemin_xlsx, read_only=False, data_only=True)

    # Chercher la feuille "PV Année"
    ws = None
    ws_si = ws_sp = None
    for sname in wb.sheetnames:
        sl = sname.lower().strip()
        if "pv ann" in sl or "pv année" in sl or "pv annee" in sl:
            ws = wb[sname]
        elif sl == f"pv {label_si.lower()}" or sl.startswith(f"pv {label_si.lower()}"):
            ws_si = wb[sname]
        elif sl == f"pv {label_sp.lower()}" or sl.startswith(f"pv {label_sp.lower()}"):
            ws_sp = wb[sname]
    if ws is None:
        for sname in reversed(wb.sheetnames):
            if sname.upper().startswith("PV"):
                ws = wb[sname]; break
    if ws is None:
        ws = wb.active

    cols = _trouver_colonnes_pv(ws, label_si, label_sp)

    # Construire index CIN→row pour PV S1 et PV S2 (résolution formules)
    def _build_sem_index(ws_sem):
        """{ CIN_upper: { col_idx_annuel: valeur } } depuis une feuille semestre."""
        if ws_sem is None:
            return {}
        cols_sem = _trouver_colonnes_pv(ws_sem, label_si, label_sp)
        col_cin_sem = cols_sem.get("cin", 1)
        col_mas_sem = cols_sem.get("massar", 2)
        idx = {}
        for ri in range(4, ws_sem.max_row + 1):
            cin_v = _safe_str(ws_sem.cell(ri, col_cin_sem).value).upper()
            mas_v = _safe_str(ws_sem.cell(ri, col_mas_sem).value).upper()
            for key_val in (cin_v, mas_v):
                if key_val and key_val not in ("", "NONE", "NAN"):
                    idx[key_val] = ri
        return idx, cols_sem

    idx_si, cols_si = _build_sem_index(ws_si) if ws_si else ({}, {})
    idx_sp, cols_sp = _build_sem_index(ws_sp) if ws_sp else ({}, {})

    def _resolve_sem_val(cin, massar, col_key, is_si):
        """Lit la valeur depuis PV S1 ou S2 si dispo, sinon None."""
        ws_sem  = ws_si if is_si else ws_sp
        idx_sem = idx_si if is_si else idx_sp
        cols_s  = cols_si if is_si else cols_sp
        if ws_sem is None:
            return None
        key = cin.upper() if cin else massar.upper()
        row_sem = idx_sem.get(key) or idx_sem.get(massar.upper() if massar else "")
        if row_sem is None:
            return None
        col = cols_s.get(col_key)
        if col is None:
            return None
        return ws_sem.cell(row_sem, col).value

    # Données à partir de la ligne 4
    rows_out = []
    for row_idx in range(4, ws.max_row + 1):
        def cv(col_key):
            c = cols.get(col_key)
            if c is None:
                return None
            return ws.cell(row_idx, c).value

        cin    = _safe_str(cv("cin"))
        massar = _safe_str(cv("massar"))
        nom    = _safe_str(cv("nom"))
        prenom = _safe_str(cv("prenom"))

        # Sauter lignes vides ou parasites
        if not cin and not massar and not nom:
            continue
        if cin.lower() in ("cin", "ب-ت-و", "") and massar.lower() in ("massar", ""):
            continue

        cin_v    = _safe_str(cv("cin"))
        massar_v = _safe_str(cv("massar"))

        # Résoudre Moy. depuis PV semestre si la cellule est vide/formule non calc.
        raw_moy_si = cv(f"moy_{label_si}")
        raw_moy_sp = cv(f"moy_{label_sp}")
        raw_moy_si = raw_moy_si if not _is_na(raw_moy_si) and not isinstance(raw_moy_si, str) else                      _resolve_sem_val(cin_v, massar_v, f"moy_{label_si}", True)
        raw_moy_sp = raw_moy_sp if not _is_na(raw_moy_sp) and not isinstance(raw_moy_sp, str) else                      _resolve_sem_val(cin_v, massar_v, f"moy_{label_sp}", False)

        moy_si = _safe_float(raw_moy_si)
        dec_si = _safe_str(cv(f"dec_{label_si}"))
        av_si  = _safe_str(cv(f"av_{label_si}"))

        moy_sp = _safe_float(raw_moy_sp)
        dec_sp = _safe_str(cv(f"dec_{label_sp}"))
        av_sp  = _safe_str(cv(f"av_{label_sp}"))

        dec_ann = _safe_str(cv("decision"))
        men_ann = _safe_str(cv("mention"))

        rows_out.append({
            "CIN":    cin,
            "Massar": massar,
            "Nom":    nom,
            "Prénom": prenom,
            f"Moy_{label_si}":  moy_si,
            f"Dec_{label_si}":  dec_si,
            f"AV_{label_si}":   av_si,
            f"Moy_{label_sp}":  moy_sp,
            f"Dec_{label_sp}":  dec_sp,
            f"AV_{label_sp}":   av_sp,
            "Decision_Annuelle": dec_ann,
            "Mention_Annuelle":  men_ann,
        })

    wb.close()
    df = pd.DataFrame(rows_out)
    df.dropna(subset=["CIN", "Massar", "Nom"], how="all", inplace=True)
    df.reset_index(drop=True, inplace=True)
    return df


# ─── Fusion des 3 années ──────────────────────────────────────────────────────

def _cle_etudiant(row) -> str:
    """Clé de jointure : préfère Massar, sinon CIN."""
    m = _safe_str(row.get("Massar", ""))
    c = _safe_str(row.get("CIN", ""))
    k = m if m and m not in ("nan", "NaN", "None", "") else c
    return k.upper().strip()


def fusionner_pvs_licence(
    df1a: pd.DataFrame,   # PV 1ère année : S1 + S2
    df2a: pd.DataFrame,   # PV 2ème année : S3 + S4
    df3a: pd.DataFrame,   # PV 3ème année : S5 + S6
    annee_1a: str = "",
    annee_2a: str = "",
    annee_3a: str = "",
) -> pd.DataFrame:
    """
    Fusionne les 3 DataFrames annuels sur la clé Massar/CIN.
    Retourne un DataFrame avec les 6 semestres + moyennes + décision licence.
    """
    SEM_COLS = ["S1", "S2", "S3", "S4", "S5", "S6"]

    # Ajouter une clé de jointure à chaque df
    for df in (df1a, df2a, df3a):
        df["_cle"] = df.apply(_cle_etudiant, axis=1)

    # Index par clé
    idx1 = df1a.set_index("_cle")
    idx2 = df2a.set_index("_cle") if df2a is not None and len(df2a) > 0 else None
    idx3 = df3a.set_index("_cle") if df3a is not None and len(df3a) > 0 else None

    # Union de toutes les clés
    cles = set(df1a["_cle"].tolist())
    if idx2 is not None: cles |= set(df2a["_cle"].tolist())
    if idx3 is not None: cles |= set(df3a["_cle"].tolist())
    cles.discard("")

    def _row_unique(idx, cle):
        """Retourne toujours une Series (1ère ligne si duplicats)."""
        if cle not in idx.index:
            return None
        r = idx.loc[cle]
        # Si plusieurs lignes (duplicats), prendre la première
        if isinstance(r, pd.DataFrame):
            r = r.iloc[0]
        return r

    def _get(r, field, default=""):
        """Lecture sécurisée d'un champ depuis une Series ou dict."""
        if r is None:
            return default
        try:
            v = r[field] if field in r.index else default
        except (KeyError, TypeError):
            v = default
        return v

    rows_out = []
    for cle in sorted(cles):
        r1 = _row_unique(idx1, cle)
        r2 = _row_unique(idx2, cle) if idx2 is not None else None
        r3 = _row_unique(idx3, cle) if idx3 is not None else None

        # Identité : prendre la première source non vide
        def _id(field):
            for r in (r1, r2, r3):
                if r is not None:
                    v = _safe_str(_get(r, field))
                    if v and v not in ("nan", "NaN", "None"):
                        return v
            return ""

        cin    = _id("CIN")
        massar = _id("Massar")
        nom    = _id("Nom")
        prenom = _id("Prénom")

        # Moyennes, décisions, AV des 6 semestres
        def _sem(r, sem, annee_uv):
            if r is None:
                return None, "", ""
            moy = _safe_float(_get(r, f"Moy_{sem}"))
            dec = _safe_str(_get(r, f"Dec_{sem}"))
            av  = _safe_str(_get(r, f"AV_{sem}"))
            if not av or av in ("nan", ""):
                av = annee_uv
            return moy, dec, av

        moy_s1, dec_s1, av_s1 = _sem(r1, "S1", annee_1a)
        moy_s2, dec_s2, av_s2 = _sem(r1, "S2", annee_1a)
        moy_s3, dec_s3, av_s3 = _sem(r2, "S3", annee_2a)
        moy_s4, dec_s4, av_s4 = _sem(r2, "S4", annee_2a)
        moy_s5, dec_s5, av_s5 = _sem(r3, "S5", annee_3a)
        moy_s6, dec_s6, av_s6 = _sem(r3, "S6", annee_3a)

        # Moyenne Licence = moyenne des 6 semestres disponibles
        moys_dispo = [m for m in [moy_s1, moy_s2, moy_s3, moy_s4, moy_s5, moy_s6] if m is not None]
        moy_licence = round(sum(moys_dispo) / len(moys_dispo), 2) if moys_dispo else None

        # Décision Licence
        if moy_licence is None:
            dec_licence = "Incomplet"
        elif moy_licence >= 10:
            dec_licence = "Admis"
        elif moy_licence >= 8:
            dec_licence = "Rachat possible"
        else:
            dec_licence = "Ajourné"

        # Mention Licence
        def _mention(m):
            if m is None: return ""
            if m >= 16: return "Très Bien"
            if m >= 14: return "Bien"
            if m >= 12: return "Assez Bien"
            if m >= 10: return "Passable"
            return ""

        rows_out.append({
            "CIN":    cin,
            "Massar": massar,
            "Nom":    nom,
            "Prénom": prenom,
            # Semestre 1
            "Moy_S1": moy_s1, "Dec_S1": dec_s1, "AV_S1": av_s1,
            # Semestre 2
            "Moy_S2": moy_s2, "Dec_S2": dec_s2, "AV_S2": av_s2,
            # Semestre 3
            "Moy_S3": moy_s3, "Dec_S3": dec_s3, "AV_S3": av_s3,
            # Semestre 4
            "Moy_S4": moy_s4, "Dec_S4": dec_s4, "AV_S4": av_s4,
            # Semestre 5
            "Moy_S5": moy_s5, "Dec_S5": dec_s5, "AV_S5": av_s5,
            # Semestre 6
            "Moy_S6": moy_s6, "Dec_S6": dec_s6, "AV_S6": av_s6,
            # Résumé Licence
            "Moy_Licence":   moy_licence,
            "Dec_Licence":   dec_licence,
            "Mention_Licence": _mention(moy_licence),
        })

    df_out = pd.DataFrame(rows_out)
    df_out.reset_index(drop=True, inplace=True)
    return df_out


# ─── Export Excel PV Global Licence ───────────────────────────────────────────

def exporter_pv_licence_excel(
    df: pd.DataFrame,
    ecole:   str = "",
    filiere: str = "",
    chemin_sortie: str = "pv_global_licence.xlsx",
    langue:  str = "fr",
) -> str:
    """
    Génère un fichier Excel du PV Global de Licence avec 6 colonnes semestrielles.
    """
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    import os

    BLEU_DARK  = "1A3A5C"
    BLEU_MID   = "2563A8"
    BLEU_LIGHT = "B8D4F0"
    VERT   = "C6EFCE"; VERT_T  = "276221"
    ORANGE = "FFEB9C"; ORANGE_T= "9C5700"
    ROUGE  = "FFC7CE"; ROUGE_T = "9C0006"
    GRIS   = "F1F5F9"; BLANC   = "FFFFFF"
    OR     = "F59E0B"

    is_ar = (langue == "ar")

    def _b():
        s = Side(style="thin", color="CBD5E1")
        return Border(left=s, right=s, top=s, bottom=s)

    def _f(hex_color):
        return PatternFill("solid", fgColor=hex_color)

    def _cell(ws, row, col, value="", bold=False, size=9, color=BLANC,
              fill=BLEU_DARK, align="center", wrap=False):
        from openpyxl.utils import get_column_letter
        coord = f"{get_column_letter(col)}{row}"
        for merged in list(ws.merged_cells.ranges):
            if coord in merged:
                ws.unmerge_cells(str(merged))
                break
        v = "" if _is_na(value) else value
        c = ws.cell(row, col, v)
        c.font      = Font(name="Arial", bold=bold, size=size, color=color)
        c.fill      = _f(fill)
        c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
        c.border    = _b()
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

    def _fmt(v):
        if _is_na(v): return "—"
        try:
            return round(float(v), 2)
        except:
            return str(v).strip() or "—"

    wb = Workbook()
    ws = wb.active
    ws.title = "PV Global Licence"
    if is_ar:
        ws.sheet_view.rightToLeft = True

    # Titres header
    if is_ar:
        titre_pv  = f"محضر المداولات — دبلوم الإجازة  |  الشعبة : {filiere}"
        lbl_id    = "هوية الطالب"
        lbl_sems  = ["الفصل 1", "الفصل 2", "الفصل 3", "الفصل 4", "الفصل 5", "الفصل 6"]
        lbl_lic   = "إجمالي الإجازة"
        lbl_moy   = "م.عام"; lbl_dec = "قرار"; lbl_av = "م.و"
        lbl_cin   = "ب-ت-و"; lbl_mas = "مسار"; lbl_nom = "الاسم"; lbl_pre = "النسب"
        lbl_moy_l = "معدل الإجازة"; lbl_dec_l = "القرار"; lbl_men_l = "الميزة"
        lbl_admis = "مقبول"; lbl_rachat = "استدراك"; lbl_ajourne = "راسب"
    else:
        titre_pv  = f"PV Global de Licence  |  Filière : {filiere}"
        lbl_id    = "Identité"
        lbl_sems  = ["Semestre 1", "Semestre 2", "Semestre 3", "Semestre 4", "Semestre 5", "Semestre 6"]
        lbl_lic   = "Résultat Licence"
        lbl_moy   = "Moy."; lbl_dec = "Décision"; lbl_av = "Année Val."
        lbl_cin   = "CIN"; lbl_mas = "Massar"; lbl_nom = "Nom"; lbl_pre = "Prénom"
        lbl_moy_l = "Moy. Licence"; lbl_dec_l = "Décision"; lbl_men_l = "Mention"
        lbl_admis = "Admis"; lbl_rachat = "Rachat possible"; lbl_ajourne = "Ajourné"

    # Colonnes : 4 identité + 6×3 semestres + 3 licence = 4+18+3 = 25
    N_TOTAL = 25

    # ── Ligne 1 : en-tête global ──
    if ecole:
        _merge(ws, 1, 1, 1, N_TOTAL, ecole, size=11, bold=True, fill=BLEU_DARK, color=BLANC)
    else:
        _merge(ws, 1, 1, 1, N_TOTAL, titre_pv, size=11, bold=True, fill=BLEU_DARK, color=BLANC)
    ws.row_dimensions[1].height = 26

    # ── Ligne 2 : titre PV ──
    _merge(ws, 2, 1, 2, N_TOTAL, titre_pv, size=10, bold=True, fill=BLEU_MID, color=BLANC)
    ws.row_dimensions[2].height = 20

    # ── Ligne 3 : groupes ──
    col = 1
    _merge(ws, 3, col, 4, col+3, lbl_id, fill=BLEU_DARK, size=9)
    col += 4
    sem_fills = [BLEU_DARK, "1E4B8E", BLEU_DARK, "1E4B8E", BLEU_DARK, "1E4B8E"]
    for s_idx, s_label in enumerate(lbl_sems):
        _merge(ws, 3, col, 3, col+2, s_label, fill=sem_fills[s_idx], size=9, color=BLANC)
        col += 3
    _merge(ws, 3, col, 3, col+2, lbl_lic, fill=OR, size=9, color=BLANC)
    col += 3
    ws.row_dimensions[3].height = 18

    # ── Ligne 4 : sous-colonnes ──
    col = 1
    for h in [lbl_cin, lbl_mas, lbl_nom, lbl_pre]:
        _cell(ws, 4, col, h, bold=True, fill=BLEU_DARK)
        col += 1
    for s_idx in range(6):
        for h in [lbl_moy, lbl_dec, lbl_av]:
            _cell(ws, 4, col, h, bold=True, fill=sem_fills[s_idx])
            col += 1
    for h in [lbl_moy_l, lbl_dec_l, lbl_men_l]:
        _cell(ws, 4, col, h, bold=True, fill=OR, color="1C2B3A")
        col += 1
    ws.row_dimensions[4].height = 18

    # ── Données ──
    SEMS = ["S1", "S2", "S3", "S4", "S5", "S6"]
    for i, (_, row) in enumerate(df.iterrows()):
        r   = 5 + i
        bg  = GRIS if i % 2 == 0 else BLANC
        col = 1

        # Identité
        for f in ["CIN", "Massar", "Nom", "Prénom"]:
            _cell(ws, r, col, _fmt(row.get(f, "")), fill=bg, color="1C2B3A", align="left", size=9)
            col += 1

        # 6 semestres
        for sem in SEMS:
            moy = row.get(f"Moy_{sem}")
            dec = _fmt(row.get(f"Dec_{sem}", ""))
            av  = _fmt(row.get(f"AV_{sem}", ""))

            moy_v = _safe_float(moy)
            if moy_v is None:
                bg_s = ROUGE; tc_s = ROUGE_T
            elif moy_v >= 10:
                bg_s = VERT;  tc_s = VERT_T
            else:
                bg_s = ROUGE; tc_s = ROUGE_T

            _cell(ws, r, col,   _fmt(moy) if moy_v is not None else "—", fill=bg_s, bold=True, color=tc_s)
            _cell(ws, r, col+1, dec, fill=bg_s, color=tc_s, size=8)
            _cell(ws, r, col+2, av,  fill=bg_s, color=tc_s, size=8)
            col += 3

        # Résultat Licence
        moy_l = row.get("Moy_Licence")
        moy_l_v = _safe_float(moy_l)
        dec_l   = _fmt(row.get("Dec_Licence", ""))
        men_l   = _fmt(row.get("Mention_Licence", ""))

        if dec_l in (lbl_admis, "Admis"):
            bg_l = VERT; tc_l = VERT_T
        elif dec_l in (lbl_rachat, "Rachat possible"):
            bg_l = ORANGE; tc_l = ORANGE_T
        else:
            bg_l = ROUGE; tc_l = ROUGE_T

        # Traduire décision si arabe
        if is_ar:
            dec_map = {"Admis": lbl_admis, "Rachat possible": lbl_rachat, "Ajourné": lbl_ajourne}
            dec_l = dec_map.get(dec_l, dec_l)

        _cell(ws, r, col,   _fmt(moy_l) if moy_l_v is not None else "—", fill=bg_l, bold=True, color=tc_l, size=10)
        _cell(ws, r, col+1, dec_l, fill=bg_l, bold=True, color=tc_l)
        _cell(ws, r, col+2, men_l, fill=bg_l, bold=True, color=tc_l)
        col += 3

        ws.row_dimensions[r].height = 15

    # Largeurs colonnes
    widths = [10, 14, 18, 14]  # identité
    for _ in range(6):
        widths += [7, 12, 12]
    widths += [12, 16, 12]
    for idx, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(idx)].width = w

    ws.freeze_panes = "E5"

    os.makedirs(os.path.dirname(chemin_sortie) if os.path.dirname(chemin_sortie) else ".", exist_ok=True)
    wb.save(chemin_sortie)
    return chemin_sortie


# ─── Fonction principale (point d'entrée public) ──────────────────────────────

def generer_pv_global_licence(
    pv_1a: str,          # chemin PV annuel 1ère année
    pv_2a: str,          # chemin PV annuel 2ème année
    pv_3a: str,          # chemin PV annuel 3ème année
    labels_1a: tuple,    # (label_si, label_sp) ex: ("S1", "S2")
    labels_2a: tuple,    # ("S3", "S4")
    labels_3a: tuple,    # ("S5", "S6")
    annee_1a:  str = "",
    annee_2a:  str = "",
    annee_3a:  str = "",
    ecole:     str = "",
    filiere:   str = "",
    annee_lib: str = "",
    langue:    str = "fr",
    dossier_sortie: str = ".",
) -> str:
    """
    Pipeline complet : lit les 3 PV annuels, fusionne, exporte le PV Global Licence.
    Retourne le chemin du fichier généré.
    """
    import os, re

    # Lecture des 3 PV annuels
    df1 = lire_pv_annuel(pv_1a, labels_1a[0], labels_1a[1]) if pv_1a else pd.DataFrame()
    df2 = lire_pv_annuel(pv_2a, labels_2a[0], labels_2a[1]) if pv_2a else pd.DataFrame()
    df3 = lire_pv_annuel(pv_3a, labels_3a[0], labels_3a[1]) if pv_3a else pd.DataFrame()

    if df1.empty and df2.empty and df3.empty:
        raise ValueError("Aucune donnée trouvée dans les fichiers PV fournis.")

    # Fusion
    df_licence = fusionner_pvs_licence(
        df1a=df1, df2a=df2, df3a=df3,
        annee_1a=annee_1a,
        annee_2a=annee_2a,
        annee_3a=annee_3a,
    )

    # Export
    os.makedirs(dossier_sortie, exist_ok=True)
    slug = re.sub(r"[^\w\-]", "_", annee_lib) if annee_lib else "licence"
    chemin_out = os.path.join(dossier_sortie, f"PV_Global_Licence_{slug}.xlsx")

    exporter_pv_licence_excel(
        df=df_licence,
        ecole=ecole,
        filiere=filiere,
        chemin_sortie=chemin_out,
        langue=langue,
    )

    print(f"PV Global Licence généré : {chemin_out}")
    return chemin_out
