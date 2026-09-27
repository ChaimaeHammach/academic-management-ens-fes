
from __future__ import annotations
import copy, math, re
from pathlib import Path
from typing import Optional
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

VALEURS_VIDES = {"—", "-", "–", "—", "N/A", "NA", "ABS", "", "None", "nan"}
CREDITS_PAR_MODULE = 6

# ── Utilitaires ──────────────────────────────────────────────────────────────

def _est_vide(v) -> bool:
    if v is None: return True
    try:
        if math.isnan(float(v)): return True
    except (TypeError, ValueError): pass
    return str(v).strip() in VALEURS_VIDES

def _safe_float(v) -> Optional[float]:
    try: return round(float(v), 2)
    except (TypeError, ValueError): return None

def _norm(nom: str) -> str:
    """Normalise une chaîne : minuscules, sans espaces/tirets/points/underscores.
    Conserve les caractères Unicode (accents, arabe) pour éviter les collisions
    entre noms de modules arabes (ex: 'البلاغة العربية_S1' vs 'نحو 1_S1')."""
    return re.sub(r"[\s_\-\.]+", "", str(nom or "").strip().lower())

# ── Détection de la langue (arabe / français) ───────────────────────────────

_AR_RE = re.compile(r"[\u0600-\u06FF]")

def _is_arabic(ws) -> bool:
    """Détecte si une feuille PV est en arabe (en-têtes contenant de l'arabe)."""
    if ws is None: return False
    mg = _resolve_merges(ws)
    for r in (1, 2, 3):
        for c in range(1, min(ws.max_column, 25) + 1):
            v = _gc(ws, mg, r, c)
            if v and _AR_RE.search(str(v)):
                return True
    return False

def _resolve_merges(ws) -> dict:
    m = {}
    for mr in ws.merged_cells.ranges:
        val = ws.cell(mr.min_row, mr.min_col).value
        for r in range(mr.min_row, mr.max_row + 1):
            for c in range(mr.min_col, mr.max_col + 1):
                m[(r, c)] = val
    return m

def _gc(ws, mg, r, c):
    return mg.get((r, c), ws.cell(r, c).value)

# ── PV précédent – double-index CIN + Massar ────────────────────────────────

def _safe_str(v) -> str:
    """Convertit en str en gérant pd.NA / None sans lever d'exception."""
    try:
        if v is None: return ""
        import pandas as _pd
        if v is _pd.NA: return ""
    except Exception:
        pass
    try:
        if v != v: return ""   # NaN
    except Exception:
        pass
    return str(v)


def _lire_pv_prec(chemin: str) -> dict:
    """Construit index { par_cin, par_massar, cin_massar }.

    Voie 1 : lecture générique via modules.reader.lire_pv_precedent (fonctionne
             pour les PV français déjà supportés).
    Voie 2 : lecture directe (français ET arabe) en s'appuyant sur _carte_pv /
             _ids_pv / _note_fiche, qui résolvent les références croisées vers
             les feuilles modules du PV précédent. Complète la Voie 1 sans
             jamais écraser une valeur déjà trouvée (setdefault).
    """
    par_cin:    dict = {}
    par_massar: dict = {}
    cin_massar: dict = {}

    # ── Voie 1 : lecture générique (logique française existante) ───────────
    try:
        from modules.reader import lire_pv_precedent
        df = lire_pv_precedent(chemin)
        for _, row in df.iterrows():
            cin    = _safe_str(row.get("CIN",    "")).strip().upper()
            massar = _safe_str(row.get("Massar", "")).strip().upper()
            mod    = _safe_str(row.get("Module", "")).strip()
            n      = _safe_float(row.get("Note"))
            if not mod or n is None: continue
            if cin:    par_cin.setdefault(cin, {}).setdefault(mod, n)
            if massar: par_massar.setdefault(massar, {}).setdefault(mod, n)
            if cin and massar:
                cin_massar.setdefault(cin, massar)
                cin_massar.setdefault(massar, cin)
    except Exception as e:
        print(f"  ⚠ PV précédent (lecture générique) : {e}")

    # ── Voie 2 : lecture directe via structure PV (FR + AR) ─────────────────
    try:
        extra = _lire_pv_prec_direct(chemin)
        for cin, mods in extra.get("par_cin", {}).items():
            d = par_cin.setdefault(cin, {})
            for m, n in mods.items(): d.setdefault(m, n)
        for mas, mods in extra.get("par_massar", {}).items():
            d = par_massar.setdefault(mas, {})
            for m, n in mods.items(): d.setdefault(m, n)
        for k, v in extra.get("cin_massar", {}).items():
            cin_massar.setdefault(k, v)
    except Exception as e:
        print(f"  ⚠ PV précédent (lecture directe) : {e}")

    print(f"  PV préc. : {len(par_cin)} CIN, {len(par_massar)} Massar")
    return {"par_cin": par_cin, "par_massar": par_massar, "cin_massar": cin_massar}


def _lire_pv_prec_direct(chemin: str) -> dict:
    """Construit l'index { par_cin, par_massar, cin_massar } en lisant
    directement la feuille PV (S1/S2/Année) du fichier précédent et en
    résolvant les notes via les fiches modules correspondantes.
    Fonctionne pour les fichiers français ET arabes (utilise _carte_pv,
    _ids_pv et _note_fiche, qui supportent les deux langues)."""
    if not Path(chemin).exists():
        return {}

    wb = load_workbook(chemin, data_only=False)

    ws_pvs: list = []
    ws_mods_prec: dict = {}
    for sn in wb.sheetnames:
        ws = wb[sn]
        sl = sn.lower().strip()
        if sl == "pv" or sl.startswith("pv "):
            ws_pvs.append(ws)
        else:
            ws_mods_prec[_norm(sn)] = ws

    par_cin:    dict = {}
    par_massar: dict = {}
    cin_massar: dict = {}

    for ws_pv in ws_pvs:
        carte = _carte_pv(ws_pv)
        noms  = [k for k in carte if not k.startswith("__")]
        ic, _ = _ids_pv(ws_pv)
        for _, (row, _mas) in ic.items():
            identite = _id_row(ws_pv, row)
            cin_k = identite.get("CIN", "")
            mas_k = identite.get("Massar", "")
            if not cin_k and not mas_k: continue

            for nom in noms:
                md = carte.get(nom, {})
                cm = md.get("col_moy")
                if not cm: continue
                val  = ws_pv.cell(row, cm).value
                note = None
                if isinstance(val, str) and val.startswith("="):
                    m = re.match(r"='?([^!']+)'?!", val)
                    if m:
                        fiche_name = m.group(1).replace("''", "'")
                        ws_fiche = ws_mods_prec.get(_norm(fiche_name))
                        if ws_fiche is not None:
                            note = _note_fiche(ws_fiche, identite)
                else:
                    note = _safe_float(val)
                if note is None: continue

                if cin_k: par_cin.setdefault(cin_k, {}).setdefault(nom, note)
                if mas_k: par_massar.setdefault(mas_k, {}).setdefault(nom, note)

            if cin_k and mas_k:
                cin_massar.setdefault(cin_k, mas_k)
                cin_massar.setdefault(mas_k, cin_k)

    return {"par_cin": par_cin, "par_massar": par_massar, "cin_massar": cin_massar}


def _note_prec(idx: dict, identite: dict, mod: str) -> Optional[float]:
    """Cherche la note dans le PV précédent avec fallback CIN↔Massar."""
    if not idx: return None
    par_cin    = idx.get("par_cin", {})
    par_massar = idx.get("par_massar", {})
    cm         = idx.get("cin_massar", {})
    mod_n      = _norm(mod)

    def _dans(dico, cle):
        if not cle or cle not in dico: return None
        mods = dico[cle]
        if mod in mods: return mods[mod]
        for k, v in mods.items():
            kn = _norm(k)
            if kn == mod_n: return v
            if len(mod_n) >= 6 and (mod_n[:8] in kn or kn[:8] in mod_n): return v
        return None

    cin    = str(identite.get("CIN",    "") or "").strip().upper()
    massar = str(identite.get("Massar", "") or "").strip().upper()
    return (_dans(par_cin, cin)
            or _dans(par_massar, massar)
            or _dans(par_massar, cm.get(cin, ""))
            or _dans(par_cin, cm.get(massar, "")))

# ── Évaluation fiche module ──────────────────────────────────────────────────

def _note_fiche(ws_fiche, identite: dict) -> Optional[float]:
    """Calcule CC*0.4+Exam*0.6 (ou CC*0.2+TD*0.2+Exam*0.6, ou Note Finale)
    depuis la fiche module. Supporte les en-têtes en français et en arabe."""
    max_row = ws_fiche.max_row
    max_col = ws_fiche.max_column
    cin_t   = str(identite.get("CIN",    "") or "").strip().upper()
    mas_t   = str(identite.get("Massar", "") or "").strip().upper()

    hdr = col_cin = col_mas = col_cc = col_td = col_exam = col_moy = col_ratt = col_nf = None
    for r in range(1, min(8, max_row + 1)):
        for c in range(1, min(20, max_col + 1)):
            v = str(ws_fiche.cell(r, c).value or "").strip().lower()
            if v == "cin" or "بطاقة" in v: hdr = r; col_cin = c
    if hdr is None: return None

    for c in range(1, max_col + 1):
        v = str(ws_fiche.cell(hdr, c).value or "").strip().lower()
        if "massar" in v or "cne" in v or ("وطني" in v and "بطاقة" not in v):
            col_mas  = c
        elif "موجهة" in v:
            col_td = c
        elif ("c.c" in v or v.startswith("cc") or "مراقبة" in v or "مستمرة" in v) and col_cc is None:
            col_cc = c
        elif "exm" in v or "exam" in v or "امتحان" in v:
            col_exam = c
        elif v in ("moyenne", "moyenne initiale") or "معدل" in v:
            col_moy = c
        elif "rattrapage" in v or "rattr" in v or "استدراك" in v:
            col_ratt = c
        elif "note finale" in v or "note final" in v or "علامة" in v:
            col_nf = c

    for r in range(hdr + 1, max_row + 1):
        vc = str(ws_fiche.cell(r, col_cin).value or "").strip().upper() if col_cin else ""
        vm = str(ws_fiche.cell(r, col_mas).value or "").strip().upper() if col_mas else ""
        if vc not in (cin_t, mas_t) and vm not in (cin_t, mas_t): continue

        if col_nf:
            nf = _safe_float(ws_fiche.cell(r, col_nf).value)
            if nf is not None: return nf
        cc   = _safe_float(ws_fiche.cell(r, col_cc).value)   if col_cc   else None
        td   = _safe_float(ws_fiche.cell(r, col_td).value)   if col_td   else None
        exam = _safe_float(ws_fiche.cell(r, col_exam).value) if col_exam else None
        if cc is not None and exam is not None:
            if td is not None:
                mi = round(cc * 0.2 + td * 0.2 + exam * 0.6, 2)
            else:
                mi = round(cc * 0.4 + exam * 0.6, 2)
            rt = _safe_float(ws_fiche.cell(r, col_ratt).value) if col_ratt else None
            return round(max(mi, rt), 2) if rt is not None and 0 <= rt <= 20 else mi
        if cc is not None: return cc
        if col_moy: return _safe_float(ws_fiche.cell(r, col_moy).value)
    return None

# ── Cartographie feuille PV ──────────────────────────────────────────────────

def _carte_pv(ws) -> dict:
    mg = _resolve_merges(ws)
    mc = ws.max_column
    v2 = lambda c: _gc(ws, mg, 2, c)
    v3 = lambda c: _gc(ws, mg, 3, c)

    MOT_ID    = {"cin","massar","nom","prénom","prenom","identité","identite","هوية"}
    MOT_RECAP = {"semestre","récapitulatif","recap","الفصل"}
    MOT_SYNTH = {"moyenne\ngénérale","modules\nvalidés","crédits","décision",
                 "rattrapages","mention","modules validés","moyenne générale",
                 "معدل العام","وحدات","رصد","قرار","استدراك","ميزة"}

    carte: dict = {}; grp = None
    for c in range(1, mc + 1):
        v2s = str(v2(c) or "").strip(); v3s = str(v3(c) or "").strip()
        v2l = v2s.lower(); v3l = v3s.lower()
        if v2s:
            if any(k in v2l for k in MOT_ID):       grp = "__id__"
            elif any(k in v2l for k in MOT_SYNTH):  grp = "__synth__"
            elif any(k in v2l for k in MOT_RECAP):  grp = "__recap__"
            else:                                    grp = v2s
        if not grp or grp == "__id__": continue
        if grp == "__synth__":
            sd = carte.setdefault("__synthese__", {})
            if any(k in v2l for k in ("moyenne\ngénérale","moyenne générale","معدل العام")): sd.setdefault("col_moy_ann", c)
            elif any(k in v2l for k in ("modules\nvalidés","modules validés","وحدات")): sd.setdefault("col_mods_val", c)
            elif "crédit"    in v2l or "رصد" in v2l: sd.setdefault("col_credits", c)
            elif "décision"  in v2l or "decision" in v2l or "قرار" in v2l: sd.setdefault("col_dec_ann", c)
            elif "rattrapage" in v2l or "استدراك" in v2l: sd.setdefault("col_ratt", c)
            elif "mention"   in v2l or "ميزة" in v2l: sd.setdefault("col_mention", c)
            continue
        if grp == "__recap__":
            rd = carte.setdefault("__recap__", {})
            if "moy" in v3l or v3s.startswith("م."):  rd.setdefault("col_moy_sem", c)
            elif any(k in v3l for k in ("déc","dec")) or v3s.startswith("ق."): rd.setdefault("col_dec_sem", c)
            elif "av" in v3l or v3s == "سنة": rd.setdefault("col_av_sem", c)
            elif "mention" in v3l or "ميزة" in v3l: rd.setdefault("col_mention_sem", c)
            continue
        md = carte.setdefault(grp, {})
        if "moy" in v3l or "معدل" in v3l:  md.setdefault("col_moy", c)
        elif "av" in v3l or v3s == "سنة": md.setdefault("col_av",  c)
        elif any(k in v3l for k in ("déc","dec")) or "قرار" in v3l: md.setdefault("col_dec", c)
    return carte


def _ids_pv(ws) -> tuple[dict, dict]:
    mg = _resolve_merges(ws); mc = ws.max_column
    col_cin = col_mas = None
    for c in range(1, min(8, mc + 1)):
        # FIX-AR-1 : ne pas exclure "بطاقة" — "رقم البطاقة الوطنية" doit être reconnu comme CIN
        raw = str(_gc(ws, mg, 3, c) or "").strip()
        v = raw.lower()
        if "cin" in v or "بطاقة" in raw or ("وطني" in raw and "مسار" not in raw): col_cin = c
        elif "massar" in v or "مسار" in raw or ("cne" in v and col_mas is None): col_mas = c
    ids_c:  dict = {}
    ids_m:  dict = {}
    lignes_vues = set()
    for r in range(4, ws.max_row + 1):
        cin_raw = ws.cell(r, col_cin).value if col_cin else None
        mas_raw = ws.cell(r, col_mas).value if col_mas else None
        # Ignorer les cellules avec formules (elles n'ont pas été évaluées)
        if isinstance(cin_raw, str) and cin_raw.startswith("="): cin_raw = None
        if isinstance(mas_raw, str) and mas_raw.startswith("="): mas_raw = None
        cin = str(cin_raw or "").strip().upper()
        mas = str(mas_raw or "").strip().upper()
        if not cin and not mas:
            continue
        if r in lignes_vues:
            continue
        lignes_vues.add(r)
        if cin and cin not in ("","NONE","NAN"): ids_c[cin] = (r, mas)
        if mas and mas not in ("","NONE","NAN"): ids_m[mas] = (r, cin)
    return ids_c, ids_m


def _row_etud(ids_c, ids_m, identite) -> Optional[int]:
    cin = str(identite.get("CIN",    "") or "").strip().upper()
    mas = str(identite.get("Massar", "") or "").strip().upper()
    if cin and cin in ids_c:  return ids_c[cin][0]
    if mas and mas in ids_m:  return ids_m[mas][0]
    if cin and cin in ids_m:  return ids_m[cin][0]
    if mas and mas in ids_c:  return ids_c[mas][0]
    return None


def _id_row(ws, row: int) -> dict:
    """
    Lit les champs identité d'une ligne du PV.
    Détecte dynamiquement les colonnes via les en-têtes ligne 3
    (supporte français ET arabe). Fallback sur colonnes 1-4.
    """
    mg = _resolve_merges(ws)
    mc = min(ws.max_column, 8)

    # FIX-AR-2 : variantes arabes étendues pour Nom / Prénom
    # النسب / اللقب / الاسم العائلي  →  col_nom   (nom de famille)
    # الاسم / الاسم الشخصي           →  col_prenom (prénom)
    # NB : dans les PV arabes marocains, النسب (nom) vient souvent AVANT الاسم (prénom)
    _AR_NOM    = {"النسب", "اللقب", "الاسم العائلي"}
    _AR_PRENOM = {"الاسم", "الاسم الشخصي", "الإسم"}

    col_cin = col_mas = col_nom = col_prenom = None
    for c in range(1, mc + 1):
        v  = str(_gc(ws, mg, 3, c) or "").strip()
        vl = v.lower()
        if ("cin" in vl or "بطاقة" in v or ("وطني" in v and "مسار" not in v)) and col_cin is None:
            col_cin = c
        elif ("massar" in vl or "مسار" in v or ("cne" in vl and col_mas is None)) and col_mas is None:
            col_mas = c
        elif (vl == "nom" or v in _AR_NOM or any(k in v for k in _AR_NOM)) and col_nom is None:
            col_nom = c
        elif (vl in ("prénom", "prenom") or v in _AR_PRENOM
              or any(k in v for k in _AR_PRENOM)) and col_prenom is None:
            col_prenom = c

    # Fallback : colonnes fixes si aucun en-tête trouvé
    if col_cin is None and col_mas is None:
        col_cin, col_mas, col_nom, col_prenom = 1, 2, 3, 4

    return {
        "CIN":    str(ws.cell(row, col_cin    or 1).value or "").strip().upper(),
        "Massar": str(ws.cell(row, col_mas    or 2).value or "").strip().upper(),
        "Nom":    str(ws.cell(row, col_nom    or 3).value or "").strip(),
        "Prénom": str(ws.cell(row, col_prenom or 4).value or "").strip(),
    }

# ── Recalcul synthèse ────────────────────────────────────────────────────────

def _trouver_col_annee_fiche(ws_fiche) -> int | None:
    """
    Trouve la colonne سنة / AV dans une feuille module.
    Cherche dans les 8 premières lignes les en-têtes connus.
    Retourne le numéro de colonne 1-based, ou None si non trouvé.
    """
    mots_annee = {"سنة", "سنه", "av", "a/v", "année validée", "annee validee"}
    for r in range(1, min(9, ws_fiche.max_row + 1)):
        for c in range(1, ws_fiche.max_column + 1):
            v = str(ws_fiche.cell(r, c).value or "").strip()
            if v.lower() in mots_annee or v in mots_annee:
                return c
    return None


def _trouver_col_note_fiche(ws_fiche) -> int | None:
    """
    Trouve la colonne Note Finale dans une feuille module.
    Cherche dans les 8 premières lignes les en-têtes connus (FR + AR).
    Retourne le numéro de colonne 1-based, ou None si non trouvé.
    """
    mots_note = {
        "note finale", "note final", "note\nfinale",
        "علامة", "النقطة النهائية", "العلامة النهائية",
    }
    for r in range(1, min(9, ws_fiche.max_row + 1)):
        for c in range(1, ws_fiche.max_column + 1):
            v = str(ws_fiche.cell(r, c).value or "").strip().lower()
            if v in mots_note or any(k in v for k in ("note finale", "note final", "علامة")):
                return c
    return None


def _trouver_col_dec_fiche(ws_fiche) -> int | None:
    """
    Trouve la colonne Décision Finale dans une feuille module.
    Cherche dans les 8 premières lignes les en-têtes connus (FR + AR).
    Retourne le numéro de colonne 1-based, ou None si non trouvé.
    """
    mots_dec_fin = {
        "décision finale", "decision finale", "décision final", "decision final",
        "déc finale", "dec finale", "décisionfinal", "decisionfinal",
        "القرار النهائي", "قرار نهائي",
    }
    # Priorité : chercher "finale" d'abord
    for r in range(1, min(9, ws_fiche.max_row + 1)):
        for c in range(1, ws_fiche.max_column + 1):
            v = str(ws_fiche.cell(r, c).value or "").strip().lower()
            if v in mots_dec_fin or any(k in v for k in ("déc final", "dec final", "قرار نهائي", "decision final")):
                return c
    # Fallback : chercher "décision" tout court
    for r in range(1, min(9, ws_fiche.max_row + 1)):
        for c in range(1, ws_fiche.max_column + 1):
            v = str(ws_fiche.cell(r, c).value or "").strip().lower()
            if v in ("décision", "decision", "déc", "dec", "قرار"):
                return c
    return None


def _dec_mod(n, lang: str = "fr"):
    if lang == "ar":
        return "م" if (n is not None and n >= 10) else "غ م"
    return "V" if (n is not None and n >= 10) else "NV"

def _mention(m, lang: str = "fr"):
    if m is None: return ""
    if lang == "ar":
        if m >= 16: return "ممتاز"
        if m >= 14: return "جيد جداً"
        if m >= 12: return "جيد"
        if m >= 10: return "مقبول"
        return ""
    if m >= 16: return "Très Bien"
    if m >= 14: return "Bien"
    if m >= 12: return "Assez Bien"
    if m >= 10: return "Passable"
    return ""

def _recalc_sem(ws, carte, row, noms, lang: str = "fr"):
    rd = carte.get("__recap__", {})
    if not rd: return
    notes = []
    for n in noms:
        col = carte.get(n, {}).get("col_moy")
        if not col: continue
        val = ws.cell(row, col).value
        if isinstance(val, str) and val.startswith("="): continue
        v = _safe_float(val)
        if v is not None: notes.append(v)
    if not notes: return
    moy = round(sum(notes)/len(notes), 2)
    # Règle décision : NV si élim (<7) ou >2 NV, sinon V si moy>=10
    nb_nv   = sum(1 for v in notes if v < 10)
    has_elim = any(v < 7 for v in notes)
    valide, non_valide = ("مكتسب", "غير مكتسب") if lang == "ar" else ("Validé", "Non Validé")
    if has_elim or nb_nv > 2:
        dec = non_valide
    elif moy >= 10:
        dec = valide
    else:
        dec = non_valide
    men = _mention(moy, lang) if dec == valide else ""
    for ck, val in (("col_moy_sem", moy), ("col_dec_sem", dec), ("col_av_sem", None), ("col_mention_sem", men)):
        if val is None: continue   # col_av_sem géré séparément
        c = rd.get(ck)
        if c:
            cur = ws.cell(row, c).value
            # Ne pas écraser une formule dynamique déjà en place
            if not (isinstance(cur, str) and cur.startswith("=")):
                ws.cell(row, c).value = val

def _recalc_ann(ws, carte, row, noms, wb=None, lang: str = "fr"):
    """
    Recalcule la synthèse annuelle.
    Si wb est fourni, résout les formules cross-sheet ='PV Sx'!$Exx
    en lisant directement les valeurs dans les feuilles sources.
    Applique les nouvelles règles de validation :
      1. Moy < 10 → Ajourné
      2. Note < 7 dans un module → Rachat possible
      3. > 2 modules non validés (7 <= note < 10) → Rachat possible
    """
    sd = carte.get("__synthese__", {})
    notes = {}
    for n in noms:
        col = carte.get(n, {}).get("col_moy")
        if not col: notes[n] = None; continue
        val = ws.cell(row, col).value
        if isinstance(val, str) and val.startswith("="):
            # Tenter de résoudre ='SheetName'!$ColRow
            resolved = None
            if wb is not None:
                m = re.match(r"='?([^!']+)'?!\$?([A-Z]+)(\d+)", val)
                if m:
                    sh_name = m.group(1).replace("''", "'")
                    col_ltr = m.group(2)
                    row_n   = int(m.group(3))
                    if sh_name in wb.sheetnames:
                        from openpyxl.utils import column_index_from_string
                        c_idx = column_index_from_string(col_ltr)
                        src_val = wb[sh_name].cell(row_n, c_idx).value
                        resolved = _safe_float(src_val)
            notes[n] = resolved
        else:
            notes[n] = _safe_float(val)
    known = [v for v in notes.values() if v is not None]
    if not known: return
    moy      = round(sum(known)/len(known), 2)
    nb_val   = sum(1 for v in known if v >= 10)
    nb_total = len(noms)
    credits  = nb_val * CREDITS_PAR_MODULE
    nv_mods  = [n for n, v in notes.items() if v is None or v < 10]
    rattrap  = ", ".join(nv_mods) if nv_mods else ""

    # Nouvelles règles de validation annuelle
    if moy < 10:
        dec_key = "ajourn"
    else:
        # Vérifier note < 7 ou > 2 modules non validés
        nb_elim = sum(1 for v in known if v < 7)
        nb_nv_ann = sum(1 for v in known if 7 <= v < 10)
        if nb_elim > 0 or nb_nv_ann > 2:
            dec_key = "rachat"
        else:
            dec_key = "admis"

    if lang == "ar":
        dec_map = {"admis": "مقبول", "rachat": "استدراك ممكن", "ajourn": "راسب"}
    else:
        dec_map = {"admis": "Admis", "rachat": "Rachat possible", "ajourn": "Ajourné"}
    dec = dec_map[dec_key]

    admis_val = "مقبول" if lang == "ar" else "Admis"
    mention = _mention(moy, lang) if dec == admis_val else ""

    def _set(k, v):
        c = sd.get(k)
        if c:
            cur = ws.cell(row, c).value
            if not (isinstance(cur, str) and cur.startswith("=")):
                ws.cell(row, c).value = v

    _set("col_moy_ann",  moy)
    _set("col_mods_val", f"{nb_val}/{nb_total}")
    _set("col_credits",  credits)
    _set("col_dec_ann",  dec)
    _set("col_ratt",     rattrap)
    _set("col_mention",  mention)

# ── Injection d'une ligne ────────────────────────────────────────────────────

def _injecter_ligne(
    ws, carte, row, noms, identite, idx_prec,
    ws_mods, ws_s1=None, ws_s2=None,
    label="", lang: str = "fr"
) -> bool:
    """
    Version simplifiée :
    - modules déjà injectés → row_module toujours trouvable
    - logique claire et linéaire
    """

    ok = False

    cin_k = str(identite.get("CIN", "") or "").strip().upper()
    mas_k = str(identite.get("Massar", "") or "").strip().upper()

    for nom in noms:
        md  = carte.get(nom, {})
        cm  = md.get("col_moy")
        cd  = md.get("col_dec")
        cav = md.get("col_av")

        if not cm:
            continue

        # ── 1. Trouver feuille module + ligne ──
        ws_fiche = ws_mods.get(_norm(nom))
        row_module = None

        if ws_fiche:
            idx_mod = _lire_index_module_courant(ws_fiche)
            idx_index = idx_mod.get("index", {})
            cin_mas_map = idx_mod.get("cin_massar_map", {})

            row_module = (
                idx_index.get(cin_k)
                or idx_index.get(mas_k)
                or idx_index.get(cin_mas_map.get(cin_k, ""))
                or idx_index.get(cin_mas_map.get(mas_k, ""))
            )

        # ── 2. Récupérer note ──
        val = ws.cell(row, cm).value
        note = None

        if isinstance(val, str) and val.startswith("="):
            # formule → lire depuis fiche
            if ws_fiche:
                note = _note_fiche(ws_fiche, identite)

        else:
            # valeur directe
            note = _safe_float(val)

        # fallback PV précédent si vide
        if note is None:
            note = _note_prec(idx_prec, identite, nom)

        # ── 3. Injecter note et décision — formules dynamiques si possible ──
        if row_module and ws_fiche:
            from openpyxl.utils import get_column_letter as _gcl
            fiche_title_dyn = ws_fiche.title[:31].replace("'", "''")

            # Trouver les colonnes Note Finale et Décision Finale dans la feuille module
            col_note_fiche = _trouver_col_note_fiche(ws_fiche)
            col_dec_fiche  = _trouver_col_dec_fiche(ws_fiche)

            if col_note_fiche:
                # MOY → formule dynamique vers la Note Finale du module
                ws.cell(row, cm).value = (
                    f"='{fiche_title_dyn}'!${_gcl(col_note_fiche)}{row_module}"
                )
                ok = True

            if cd and col_dec_fiche:
                # DEC → formule dynamique vers la Décision Finale du module
                ws.cell(row, cd).value = (
                    f"='{fiche_title_dyn}'!${_gcl(col_dec_fiche)}{row_module}"
                )

            # Si colonnes non trouvées dans la fiche, fallback valeur statique
            if not col_note_fiche and note is not None:
                ws.cell(row, cm).value = note
                ok = True
            if cd and not col_dec_fiche and note is not None:
                ws.cell(row, cd).value = _dec_mod(note, lang)

        else:
            # Pas de feuille module disponible → fallback statique
            if note is not None:
                ws.cell(row, cm).value = note
                if cd:
                    ws.cell(row, cd).value = _dec_mod(note, lang)
                ok = True

        # ── 4. Injecter AV (colonne سنة réelle dans la feuille module) ──
        if cav and row_module and ws_fiche:
            fiche_title = ws_fiche.title[:31].replace("'", "''")
            # Trouver la colonne سنة / AV réelle dans la feuille module
            col_annee_fiche = _trouver_col_annee_fiche(ws_fiche)
            if col_annee_fiche:
                from openpyxl.utils import get_column_letter as _gcl
                col_lettre = _gcl(col_annee_fiche)
                ws.cell(row, cav).value = f"='{fiche_title}'!${col_lettre}{row_module}"
            else:
                # Fallback : lire la valeur directement et l'écrire en statique
                val_av = ws_fiche.cell(row_module, 1).value  # colonne A = fallback
                if val_av is not None:
                    ws.cell(row, cav).value = str(val_av).strip()

    # ── 5. Recalcul ──
    if ok:
        _recalc_sem(ws, carte, row, noms, lang)

    return ok
# ── Ajout des absents dans un PV semestre ────────────────────────────────────

def _ajouter_absents(ws_pv, ws_an, carte_pv, noms_pv,
                     ids_cin_pv, ids_mas_pv, idx_prec, lang: str = "fr") -> None:
    """
    Étudiants dans PV Année mais absents du PV semestre, avec '—' pour ses modules.
    → Ajoute une ligne dans PV semestre, met à jour les formules dans PV Année.
    """
    carte_an = _carte_pv(ws_an)
    ids_cin_an, _ = _ids_pv(ws_an)

    for cin_an, (row_an, mas_an) in ids_cin_an.items():
        identite = _id_row(ws_an, row_an)

        if _row_etud(ids_cin_pv, ids_mas_pv, identite) is not None:
            continue  # déjà présent

        # A-t-il des '—' dans les modules de ce semestre ?
        a_vide = any(
            _est_vide(ws_an.cell(row_an,
                carte_an.get(m, {}).get("col_moy") or 0).value)
            for m in noms_pv if carte_an.get(m, {}).get("col_moy")
        )
        if not a_vide:
            continue

        # Collecter notes depuis PV précédent
        notes: dict[str, float] = {}
        for m in noms_pv:
            n = _note_prec(idx_prec, identite, m)
            if n is not None: notes[m] = n

        if not notes:
            print(f"    ✗ {identite.get('CIN')} absent PV sem, non trouvé dans PV préc.")
            continue

        # FIX-AR-3 : détecter les vraies colonnes du PV semestre (arabe/français)
        # au lieu de supposer que CIN=1, Massar=2, Nom=3, Prénom=4
        mg_pv = _resolve_merges(ws_pv)
        mc_pv = min(ws_pv.max_column, 8)
        _AR_NOM    = {"النسب", "اللقب", "الاسم العائلي"}
        _AR_PRENOM = {"الاسم", "الاسم الشخصي", "الإسم"}
        _c_cin = _c_mas = _c_nom = _c_prenom = None
        for _c in range(1, mc_pv + 1):
            _hv  = str(_gc(ws_pv, mg_pv, 3, _c) or "").strip()
            _hvl = _hv.lower()
            if ("cin" in _hvl or "بطاقة" in _hv or ("وطني" in _hv and "مسار" not in _hv)) and _c_cin is None:
                _c_cin = _c
            elif ("massar" in _hvl or "مسار" in _hv or ("cne" in _hvl and _c_mas is None)) and _c_mas is None:
                _c_mas = _c
            elif (_hvl == "nom" or _hv in _AR_NOM or any(k in _hv for k in _AR_NOM)) and _c_nom is None:
                _c_nom = _c
            elif (_hvl in ("prénom", "prenom") or _hv in _AR_PRENOM
                  or any(k in _hv for k in _AR_PRENOM)) and _c_prenom is None:
                _c_prenom = _c
        _c_cin    = _c_cin    or 1
        _c_mas    = _c_mas    or 2
        _c_nom    = _c_nom    or 3
        _c_prenom = _c_prenom or 4

        # Ajouter la ligne
        nr = ws_pv.max_row + 1
        ws_pv.cell(nr, _c_cin   ).value = identite.get("CIN")
        ws_pv.cell(nr, _c_mas   ).value = identite.get("Massar")
        ws_pv.cell(nr, _c_nom   ).value = identite.get("Nom")
        ws_pv.cell(nr, _c_prenom).value = identite.get("Prénom")
        for m, n in notes.items():
            md = carte_pv.get(m, {})
            if md.get("col_moy"): ws_pv.cell(nr, md["col_moy"]).value = n
            if md.get("col_dec"): ws_pv.cell(nr, md["col_dec"]).value = _dec_mod(n, lang)
        _recalc_sem(ws_pv, carte_pv, nr, noms_pv, lang)
        print(f"    ✓ Ligne {nr} ajoutée dans PV sem : "
              f"{identite.get('CIN')} ({identite.get('Nom')} {identite.get('Prénom')})")

        # Mettre à jour PV Année : remplacer '—' par formule
        pv_t = ws_pv.title[:31].replace("'", "''")
        for m, n in notes.items():
            c_an = carte_an.get(m, {}).get("col_moy")
            c_pv = carte_pv.get(m, {}).get("col_moy")
            d_an = carte_an.get(m, {}).get("col_dec")
            d_pv = carte_pv.get(m, {}).get("col_dec")
            if c_an and c_pv:
                ws_an.cell(row_an, c_an).value = (
                    f"='{pv_t}'!${get_column_letter(c_pv)}{nr}")
            if d_an and d_pv:
                ws_an.cell(row_an, d_an).value = (
                    f"='{pv_t}'!${get_column_letter(d_pv)}{nr}")


# ══════════════════════════════════════════════════════════════════════════════
# ── NOUVEAU : Injection des redoublants dans les feuilles modules ─────────────
# ══════════════════════════════════════════════════════════════════════════════

def _lire_index_module_prec(ws_mod_prec) -> dict:
    """
    Lit une feuille module du PV précédent et retourne un index
    { CIN_ou_Massar_upper: numero_ligne } pour retrouver la ligne d'un étudiant.

    Détecte automatiquement :
      - la ligne d'en-tête (contient CIN / رقم البطاقة الوطنية / رقم الوطني)
      - la colonne CIN  (رقم البطاقة الوطنية ou « cin »)
      - la colonne Massar (رقم الوطني / رقم مسار / massar / cne)

    Retourne aussi hdr_row, col_cin, col_massar pour usage externe.
    """
    max_row = ws_mod_prec.max_row
    max_col = ws_mod_prec.max_column

    hdr_row = col_cin = col_massar = None

    # Recherche de l'en-tête sur les 10 premières lignes
    for r in range(1, min(11, max_row + 1)):
        found = False
        for c in range(1, min(20, max_col + 1)):
            v = str(ws_mod_prec.cell(r, c).value or "").strip()
            vl = v.lower()
            # FIX-AR-4 : "رقم البطاقة الوطنية" contient "بطاقة" ET "وطني" → ne pas les exclure
            if vl == "cin" or "بطاقة" in v or ("وطني" in v and "مسار" not in v):
                col_cin = c
                hdr_row = r
                found = True
            # Détection Massar / رقم مسار / رقم وطني (sans بطاقة)
            if ("massar" in vl or "cne" in vl or "مسار" in v
                    or ("وطني" in v and "بطاقة" in v and col_massar is None)):
                col_massar = c
                hdr_row = r
                found = True
        if found:
            break

    if hdr_row is None:
        return {"hdr_row": None, "col_cin": None, "col_massar": None, "index": {}}

    INVALIDES = {"", "NONE", "NAN"}
    index: dict[str, int] = {}          # { clé_upper: numéro_de_ligne }
    cin_massar_map: dict[str, str] = {} # croisement CIN ↔ Massar

    for r in range(hdr_row + 1, max_row + 1):
        cin_val = str(ws_mod_prec.cell(r, col_cin).value or "").strip().upper() \
                  if col_cin else ""
        mas_val = str(ws_mod_prec.cell(r, col_massar).value or "").strip().upper() \
                  if col_massar else ""
        cin_ok  = cin_val and cin_val not in INVALIDES
        mas_ok  = mas_val and mas_val not in INVALIDES

        if not cin_ok and not mas_ok:
            continue

        # Enregistrer les deux clés possibles → même ligne
        if cin_ok:
            index.setdefault(cin_val, r)
        if mas_ok:
            index.setdefault(mas_val, r)
        if cin_ok and mas_ok:
            cin_massar_map[cin_val] = mas_val
            cin_massar_map[mas_val] = cin_val

    return {
        "hdr_row":        hdr_row,
        "col_cin":        col_cin,
        "col_massar":     col_massar,
        "index":          index,
        "cin_massar_map": cin_massar_map,
    }


def _lire_index_module_courant(ws_mod_cur) -> dict:
    """
    Identique à _lire_index_module_prec mais pour une feuille du workbook courant.
    Retourne le même format de dictionnaire.
    """
    return _lire_index_module_prec(ws_mod_cur)


def _copier_style_cellule(src_cell, dst_cell) -> None:
    """Copie font, fill, alignment, border, number_format d'une cellule."""
    try:
        if src_cell.font:
            dst_cell.font = copy.copy(src_cell.font)
    except Exception:
        pass
    try:
        if src_cell.fill and src_cell.fill.fill_type not in (None, "none"):
            dst_cell.fill = copy.copy(src_cell.fill)
    except Exception:
        pass
    try:
        if src_cell.alignment:
            dst_cell.alignment = copy.copy(src_cell.alignment)
    except Exception:
        pass
    try:
        if src_cell.border:
            dst_cell.border = copy.copy(src_cell.border)
    except Exception:
        pass
    try:
        if src_cell.number_format:
            dst_cell.number_format = src_cell.number_format
    except Exception:
        pass


def _ajuster_formule_ligne(formule: str, src_row: int, dst_row: int) -> str:
    """
    Décale les références de lignes relatives dans une formule Excel.
    Seules les références relatives (A8, B8) sont ajustées ;
    les absolues ($A$8, $B$8) restent intactes.

    Exemple : formule sur src_row=8 copiée en dst_row=42
      =IF(F8>=10,"مقبول","غير مقبول")  →  =IF(F42>=10,"مقبول","غير مقبول")
    """
    if not formule or not formule.startswith("="):
        return formule
    offset = dst_row - src_row
    if offset == 0:
        return formule

    def _replacer(m):
        col_part = m.group(1)   # ex: "F" ou "$F"
        row_part = m.group(2)   # ex: "8"
        # Référence absolue sur la ligne → ne pas toucher
        # Une référence absolue de ligne ressemble à $F$8 (deux $) ou F$8 (un $ avant le numéro)
        full = m.group(0)
        dollar_before_row = "$" + row_part in full and col_part + "$" + row_part in full
        if dollar_before_row:
            return full
        new_row = int(row_part) + offset
        return f"{col_part}{new_row}"

    return re.sub(r"(\$?[A-Za-z]{1,3})(\d+)", _replacer, formule)


def inject_lignes_redoublants_dans_modules(
    wb_courant,
    wb_prec,
    identites_redoublants: list[dict],
) -> dict[str, list[str]]:
    """
    Pour chaque étudiant redoublant (identifié par CIN et/ou Massar),
    et pour chaque feuille module du workbook courant (toutes les feuilles
    qui ne sont pas des feuilles PV) :

      1. Cherche la feuille de même nom dans le workbook précédent.
      2. Localise la ligne de cet étudiant dans la feuille précédente.
      3. Si l'étudiant n'est PAS encore présent dans la feuille courante,
         copie toute la ligne (valeurs + formules ajustées + styles) dans
         la feuille courante, à la suite des données existantes.

    NB : La feuille PV S1 (et toutes les feuilles commençant par "PV")
         n'est JAMAIS modifiée par cette fonction.

    Paramètres
    ----------
    wb_courant           : workbook openpyxl de l'année en cours (déjà chargé)
    wb_prec              : workbook openpyxl du PV précédent (déjà chargé)
    identites_redoublants: liste de dicts {"CIN": ..., "Massar": ...,
                                           "Nom": ..., "Prénom": ...}

    Retourne
    --------
    Un dict { nom_feuille: [liste de CIN/Massar injectés] } pour le log.
    """
    if not identites_redoublants:
        return {}

    rapport: dict[str, list[str]] = {}

    # Index des feuilles modules du PV précédent (normalisé → ws)
    ws_mods_prec: dict[str, object] = {}
    for sn in wb_prec.sheetnames:
        sl = sn.lower().strip()
        if not (sl == "pv" or sl.startswith("pv ")):
            ws_mods_prec[_norm(sn)] = wb_prec[sn]

    # Parcourir les feuilles modules du workbook courant
    for sn_cur in wb_courant.sheetnames:
        sl_cur = sn_cur.lower().strip()
        # ⚠️ Ne jamais toucher aux feuilles PV
        if sl_cur == "pv" or sl_cur.startswith("pv "):
            continue

        ws_cur = wb_courant[sn_cur]
        sn_norm = _norm(sn_cur)

        # Feuille homologue dans le PV précédent
        ws_prec = ws_mods_prec.get(sn_norm)
        if ws_prec is None:
            print(f"  ⚠ Module [{sn_cur}] : pas de feuille homologue dans le PV précédent")
            continue

        # Index de la feuille précédente et courante
        idx_prec_mod = _lire_index_module_prec(ws_prec)
        idx_cur_mod  = _lire_index_module_courant(ws_cur)

        if idx_prec_mod["hdr_row"] is None:
            print(f"  ⚠ Module [{sn_cur}] : en-tête introuvable dans le PV précédent")
            continue
        if idx_cur_mod["hdr_row"] is None:
            print(f"  ⚠ Module [{sn_cur}] : en-tête introuvable dans la feuille courante")
            continue

        idx_prec_index = idx_prec_mod["index"]
        idx_cur_index  = idx_cur_mod["index"]
        cin_mas_prec   = idx_prec_mod.get("cin_massar_map", {})

        injectes: list[str] = []

        for identite in identites_redoublants:
            cin    = str(identite.get("CIN",    "") or "").strip().upper()
            massar = str(identite.get("Massar", "") or "").strip().upper()

            # L'étudiant est-il déjà dans la feuille courante ?
            deja_present = (
                (cin    and cin    in idx_cur_index) or
                (massar and massar in idx_cur_index)
            )
            if deja_present:
                continue

            # Chercher la ligne source dans le PV précédent
            src_row = (
                idx_prec_index.get(cin)
                or idx_prec_index.get(massar)
                or idx_prec_index.get(cin_mas_prec.get(cin, ""))
                or idx_prec_index.get(cin_mas_prec.get(massar, ""))
            )
            if src_row is None:
                print(f"    ✗ [{sn_cur}] {cin}/{massar} : pas trouvé dans le PV précédent")
                continue

            # Ligne de destination : à la suite des données existantes
            dst_row = ws_cur.max_row + 1

            # Copie cellule par cellule
            max_col_prec = ws_prec.max_column
            for c in range(1, max_col_prec + 1):
                src_cell = ws_prec.cell(src_row, c)
                dst_cell = ws_cur.cell(dst_row, c)

                val = src_cell.value
                if isinstance(val, str) and val.startswith("="):
                    # Ajuster les numéros de lignes dans les formules
                    val = _ajuster_formule_ligne(val, src_row, dst_row)
                dst_cell.value = val
                _copier_style_cellule(src_cell, dst_cell)

            # Hauteur de ligne
            try:
                h = ws_prec.row_dimensions[src_row].height
                if h:
                    ws_cur.row_dimensions[dst_row].height = h
            except Exception:
                pass

            cle_log = cin or massar
            injectes.append(cle_log)
            print(f"    ✓ [{sn_cur}] Ligne {src_row}→{dst_row} injectée : "
                  f"{cin or '?'}/{massar or '?'} "
                  f"({identite.get('Nom', '')} {identite.get('Prénom', '')})")

        if injectes:
            rapport[sn_cur] = injectes

    return rapport


def _collecter_identites_redoublants(wb, idx_prec: dict) -> list[dict]:
    """
    Parcourt les feuilles PV S1 et PV S2 du workbook courant et collecte
    les identités des étudiants dont au moins une note a été injected
    depuis le PV précédent (i.e. ils figurent dans idx_prec).

    Retourne une liste de dicts {"CIN": ..., "Massar": ..., "Nom": ..., "Prénom": ...}
    sans doublons (dédupliqués par CIN puis Massar).
    """
    par_cin    = idx_prec.get("par_cin", {})
    par_massar = idx_prec.get("par_massar", {})

    vus: set[str] = set()
    identites: list[dict] = []

    for sn in wb.sheetnames:
        sl = sn.lower().strip()
        if not (sl.startswith("pv s1") or sl.startswith("pv s2") or
                sl == "pv s1" or sl == "pv s2"):
            continue

        ws = wb[sn]
        ids_c, ids_m = _ids_pv(ws)

        for cin, (row, massar) in ids_c.items():
            # Est-ce un redoublant ? → il est dans le PV précédent
            est_redoublant = cin in par_cin or massar in par_massar
            if not est_redoublant:
                continue
            cle = cin or massar
            if cle in vus:
                continue
            vus.add(cle)
            identites.append({
                "CIN":    cin,
                "Massar": massar,
                "Nom":    str(ws.cell(row, 3).value or "").strip(),
                "Prénom": str(ws.cell(row, 4).value or "").strip(),
            })

    return identites


# ── Complétion PV semestriels pour les redoublants ──────────────────────────

def _build_mapping_annee(ws_an) -> dict:
    """
    Lit une ligne de référence (premier étudiant normal) du PV Année pour
    extraire le mapping dynamique :
        { col_annee → {"sheet": nom_feuille_sem, "col": col_dans_sem} }
    Entièrement générique — aucune colonne codée en dur.
    """
    import re as _re
    from openpyxl.utils import column_index_from_string as _c2i

    mapping = {}
    for r in range(4, ws_an.max_row + 1):
        found = False
        for c in range(1, ws_an.max_column + 1):
            v = ws_an.cell(r, c).value
            if isinstance(v, str) and v.startswith("='") and "!" in v:
                m = _re.match(r"='([^']+)'!\$([A-Z]+)(\d+)", v)
                if m:
                    sheet = m.group(1).replace("''", "'")
                    col_idx = _c2i(m.group(2))
                    mapping[c] = {"sheet": sheet, "col": col_idx}
                    found = True
        if found:
            break
    return mapping


def _index_massar_sem(ws) -> dict:
    """{ massar_upper: row } pour une feuille PV semestriel."""
    mg = _resolve_merges(ws)
    col_mas = None
    for c in range(1, min(ws.max_column + 1, 8)):
        v = str(_gc(ws, mg, 3, c) or "").strip().lower()
        if "massar" in v or "\u0645\u0633\u0627\u0631" in v or "cne" in v:
            col_mas = c
            break
    if col_mas is None:
        col_mas = 2
    idx = {}
    for r in range(4, ws.max_row + 1):
        massar = str(ws.cell(r, col_mas).value or "").strip().upper()
        if massar and massar not in ("", "NONE", "NAN"):
            idx[massar] = r
    return idx


def _modules_r2(ws) -> set:
    """Ensemble des noms de modules lus sur la ligne 2 (hors identité étudiant)."""
    mods = set()
    SKIP = {"\u0647\u0648\u064a\u0629 \u0627\u0644\u0637\u0627\u0644\u0628(\u0629)", ""}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(2, c).value
        if v and str(v).strip() not in SKIP:
            mods.add(str(v).strip())
    return mods


def completer_pv_semestre_redoublants(wb_cur, wb_prec_data, ws_an) -> dict:
    """
    Détecte tous les redoublants présents dans PV Année courant mais absents
    d'un (ou plusieurs) PV semestriel courant, pour N étudiants et N semestres.

    Algorithme :
      1. Extraire dynamiquement le mapping col_PVAnnée → (feuille_sem, col_sem)
         depuis un étudiant normal de référence.
      2. Pour chaque étudiant du PV Année, détecter les feuilles semestrielles
         où ses colonnes liées sont manquantes (None ou pas de formule ='...').
      3. Pour chaque cas manquant :
         a. Trouver la feuille précédente homologue (mêmes modules en R2).
         b. Copier la ligne de l'étudiant depuis le PV précédent (valeurs réelles).
         c. Ajouter les formules ='feuille_sem'!$ColRow dans PV Année.

    Retourne dict { massar: [feuilles_complétées] } pour le log.
    """
    from openpyxl.utils import get_column_letter as _gcl
    from collections import defaultdict

    rapport = {}

    # ── Mapping dynamique PV Année → feuilles semestrielles ──
    mapping = _build_mapping_annee(ws_an)
    if not mapping:
        print("  \u26a0 completer_pv_semestre: PV Ann\u00e9e sans formules cross-sheet.")
        return rapport

    # Grouper par feuille semestrielle : { nom_feuille: [(col_an, col_sem), ...] }
    par_feuille: dict = defaultdict(list)
    for col_an, info in mapping.items():
        par_feuille[info["sheet"]].append((col_an, info["col"]))

    # ── Index des feuilles PV semestrielles précédentes par modules ──
    pv_sem_prec = {
        sn: wb_prec_data[sn]
        for sn in wb_prec_data.sheetnames
        if sn.lower().strip().startswith("pv s") and "ann" not in sn.lower()
    }

    def _homologue_prec(ws_cur):
        mods_cur = _modules_r2(ws_cur)
        best_sn, best_ws, best_score = None, None, 0
        for sn_p, ws_p in pv_sem_prec.items():
            mods_p = _modules_r2(ws_p)
            if not mods_cur or not mods_p:
                continue
            score = len(mods_cur & mods_p) / max(len(mods_cur), len(mods_p))
            if score > best_score:
                best_score, best_sn, best_ws = score, sn_p, ws_p
        if best_score >= 0.5:
            return best_sn, best_ws
        return None, None

    # ── Colonne Massar dans PV Année ──
    mg_an = _resolve_merges(ws_an)
    col_mas_an = 2  # fallback
    for c in range(1, min(ws_an.max_column + 1, 8)):
        v = str(_gc(ws_an, mg_an, 3, c) or "").strip().lower()
        if "massar" in v or "\u0645\u0633\u0627\u0631" in v or "cne" in v:
            col_mas_an = c
            break

    # Cache des homologues pour éviter de recalculer
    _cache_homologue: dict = {}

    # ── Parcourir chaque étudiant du PV Année ──
    for row_an in range(4, ws_an.max_row + 1):
        massar = str(ws_an.cell(row_an, col_mas_an).value or "").strip().upper()
        if not massar or massar in ("", "NONE", "NAN"):
            continue

        for sn_cur, cols_pairs in par_feuille.items():
            if sn_cur not in wb_cur.sheetnames:
                continue

            ws_cur = wb_cur[sn_cur]

            # Détecter si les colonnes liées à cette feuille sont manquantes
            cols_an = [cp[0] for cp in cols_pairs]
            manquant = all(
                ws_an.cell(row_an, c).value is None
                or (isinstance(ws_an.cell(row_an, c).value, str)
                    and not ws_an.cell(row_an, c).value.startswith("="))
                for c in cols_an
            )
            if not manquant:
                continue

            # Trouver ou confirmer la ligne dans le PV semestriel courant
            idx_cur = _index_massar_sem(ws_cur)
            if massar in idx_cur:
                row_sem = idx_cur[massar]
                # Présent mais formules absentes dans PV Année → juste lier
            else:
                # Absent → copier depuis PV précédent
                if sn_cur not in _cache_homologue:
                    _cache_homologue[sn_cur] = _homologue_prec(ws_cur)
                sn_prec, ws_prec = _cache_homologue[sn_cur]

                if ws_prec is None:
                    print(f"  \u26a0 [{sn_cur}] Pas de feuille pr\u00e9c\u00e9dente homologue")
                    continue

                idx_prec = _index_massar_sem(ws_prec)
                if massar not in idx_prec:
                    print(f"  \u26a0 [{sn_cur}] {massar} absent de \'{sn_prec}\'")
                    continue

                row_prec = idx_prec[massar]
                row_sem  = ws_cur.max_row + 1
                for c in range(1, ws_prec.max_column + 1):
                    ws_cur.cell(row_sem, c).value = ws_prec.cell(row_prec, c).value
                print(f"  \u2713 [{sn_cur}] {massar} copi\u00e9 depuis \'{sn_prec}\' R{row_prec} \u2192 R{row_sem}")

            # Écrire les formules dynamiques dans PV Année
            sheet_title = sn_cur.replace("'", "''")
            for col_an, col_sem in cols_pairs:
                ws_an.cell(row_an, col_an).value = f"='{sheet_title}'!${_gcl(col_sem)}{row_sem}"
            rapport.setdefault(massar, []).append(sn_cur)
            print(f"  \u2713 [PV Ann\u00e9e] R{row_an} ({massar}) \u2192 formules \'{sn_cur}\' R{row_sem}")

    return rapport


# ── Fonction principale ──────────────────────────────────────────────────────

def inject_notes_redoublants(
    chemin_xlsx:    str,
    chemin_pv_prec: Optional[str] = None,
    chemin_sortie:  Optional[str] = None,
) -> str:
    """
    Injecte les notes manquantes dans un fichier PV Excel généré.

    🔥 ORDRE CORRIGÉ :
      1. Injection lignes redoublants dans modules
      2. PV S1
      3. PV S2
      4. PV Année
    """

    if not Path(chemin_xlsx).exists():
        raise FileNotFoundError(f"Fichier introuvable : {chemin_xlsx}")

    wb = load_workbook(chemin_xlsx, data_only=False)

    # ── PV précédent ───────────────────────────────────────────────
    idx_prec: dict = {}
    wb_prec = None

    if chemin_pv_prec and Path(chemin_pv_prec).exists():
        idx_prec = _lire_pv_prec(chemin_pv_prec)
        wb_prec  = load_workbook(chemin_pv_prec, data_only=False)
    else:
        print("  ⚠ Aucun PV précédent fourni.")

    # ── Identifier les feuilles ────────────────────────────────────
    # FIX-AR-5 : "PV S1_2" matchait startswith("pv s1") et écrasait ws_s1.
    # Détection en 4 passes pour gérer les noms non standards.
    ws_s1 = ws_s2 = ws_an = None
    ws_mods: dict = {}

    # Passe 1 : correspondances exactes (priorité absolue)
    for sn in wb.sheetnames:
        sl = sn.lower().strip()
        if sl == "pv s1":
            ws_s1 = wb[sn]
        elif sl == "pv s2":
            ws_s2 = wb[sn]
        elif "pv ann" in sl or "pv année" in sl:
            ws_an = wb[sn]

    # Passe 2 : préfixe "pv s1 " / "pv s2 " AVEC espace (ex: "PV S1 Arabe")
    # "PV S1_2" ne commence pas par "pv s1 " → non affecté ici.
    for sn in wb.sheetnames:
        sl = sn.lower().strip()
        if ws_s1 is None and sl.startswith("pv s1 "):
            ws_s1 = wb[sn]
        elif ws_s2 is None and sl.startswith("pv s2 "):
            ws_s2 = wb[sn]

    # Passe 3 : toute autre feuille "pv s…" non encore attribuée
    # → si ws_s2 vide, c'est le PV S2 de substitution (ex: "PV S1_2")
    for sn in wb.sheetnames:
        sl = sn.lower().strip()
        if sl.startswith("pv s") and wb[sn] not in (ws_s1, ws_s2, ws_an):
            if ws_s2 is None:
                ws_s2 = wb[sn]
                print(f"  ⚠ Feuille '{sn}' prise comme PV S2 (nom non standard)")
            elif ws_s1 is None:
                ws_s1 = wb[sn]
                print(f"  ⚠ Feuille '{sn}' prise comme PV S1 (nom non standard)")

    # Passe 4 : tout le reste → feuilles modules
    pv_ws_set = {ws_s1, ws_s2, ws_an} - {None}
    for sn in wb.sheetnames:
        if wb[sn] not in pv_ws_set:
            ws_mods[_norm(sn)] = wb[sn]

    print(f"\n  Feuilles : PV S1={'✓' if ws_s1 else '✗'}  "
          f"PV S2={'✓' if ws_s2 else '✗'}  "
          f"PV Année={'✓' if ws_an else '✗'}  "
          f"Modules={[sn for sn in wb.sheetnames if not sn.startswith('PV')]}")

    # ── Détection langue ───────────────────────────────────────────
    lang = "ar" if (_is_arabic(ws_s1) or _is_arabic(ws_s2) or _is_arabic(ws_an)) else "fr"
    print(f"  Langue détectée : {lang}")

    # ════════════════════════════════════════════════════════════════
    # 🔥 ÉTAPE 1 : INJECTION MODULES AVANT PV (SOLUTION 3)
    # ════════════════════════════════════════════════════════════════

    if wb_prec is not None and idx_prec:
        print("\n  === Pré-injection dans les feuilles modules ===")

        identites_redoublants = _collecter_identites_redoublants(wb, idx_prec)

        print(f"  Redoublants détectés : {len(identites_redoublants)}")

        for id_ in identites_redoublants:
            print(f"    - {id_.get('CIN','?')} / {id_.get('Massar','?')} "
                  f"({id_.get('Nom','')} {id_.get('Prénom','')})")

        if identites_redoublants:
            rapport = inject_lignes_redoublants_dans_modules(
                wb_courant=wb,
                wb_prec=wb_prec,
                identites_redoublants=identites_redoublants,
            )

            if rapport:
                print("\n  Résumé injection modules :")
                for feuille, cles in rapport.items():
                    print(f"    [{feuille}] : {len(cles)} ligne(s) → {cles}")
            else:
                print("  Aucun ajout (déjà présents ou introuvables).")

    elif wb_prec is None:
        print("\n  ⚠ Injection modules ignorée : PV précédent non fourni.")

    # ════════════════════════════════════════════════════════════════
    # 🔥 ÉTAPE 1.5 : COMPLÉTION PV SEMESTRIELS POUR LES REDOUBLANTS
    # Détecte tous les redoublants absents d'un PV semestriel courant
    # et les complète depuis le PV semestriel précédent correspondant,
    # puis met à jour les formules du PV Année pour les lier.
    # ════════════════════════════════════════════════════════════════

    if ws_an is not None and chemin_pv_prec and Path(chemin_pv_prec).exists():
        print("\n  === Complétion PV semestriels (redoublants absents) ===")
        wb_prec_data = load_workbook(chemin_pv_prec, data_only=True)
        rapport_sem = completer_pv_semestre_redoublants(wb, wb_prec_data, ws_an)
        if rapport_sem:
            print(f"  {len(rapport_sem)} redoublant(s) complété(s) : {list(rapport_sem.keys())}")
        else:
            print("  Aucun redoublant à compléter dans les PV semestriels.")

    # ════════════════════════════════════════════════════════════════
    # 🔥 ÉTAPE 2 : PV S1
    # ════════════════════════════════════════════════════════════════

    if ws_s1 is not None:
        print("\n  === PV S1 ===")
        carte = _carte_pv(ws_s1)
        noms  = [k for k in carte if not k.startswith("__")]

        ic, im = _ids_pv(ws_s1)

        if ws_an:
            _ajouter_absents(ws_s1, ws_an, carte, noms, ic, im, idx_prec, lang)
            ic, im = _ids_pv(ws_s1)

        for cin, (row, _) in ic.items():
            identite = _id_row(ws_s1, row)

            if not identite.get("CIN") and not identite.get("Massar"):
                continue

            _injecter_ligne(
                ws_s1, carte, row, noms,
                identite, idx_prec, ws_mods,
                label="S1", lang=lang
            )

    # ════════════════════════════════════════════════════════════════
    # 🔥 ÉTAPE 3 : PV S2
    # ════════════════════════════════════════════════════════════════

    if ws_s2 is not None:
        print("\n  === PV S2 ===")
        carte = _carte_pv(ws_s2)
        noms  = [k for k in carte if not k.startswith("__")]

        ic, im = _ids_pv(ws_s2)

        if ws_an:
            _ajouter_absents(ws_s2, ws_an, carte, noms, ic, im, idx_prec, lang)
            ic, im = _ids_pv(ws_s2)

        for cin, (row, _) in ic.items():
            identite = _id_row(ws_s2, row)

            if not identite.get("CIN") and not identite.get("Massar"):
                continue

            _injecter_ligne(
                ws_s2, carte, row, noms,
                identite, idx_prec, ws_mods,
                label="S2", lang=lang
            )

    # ════════════════════════════════════════════════════════════════
    # 🔥 ÉTAPE 4 : PV ANNÉE
    # ════════════════════════════════════════════════════════════════

    if ws_an is not None:
        print("\n  === PV Année ===")
        carte = _carte_pv(ws_an)
        noms  = [k for k in carte if not k.startswith("__")]

        ic, _ = _ids_pv(ws_an)

        for cin, (row, _) in ic.items():
            identite = _id_row(ws_an, row)

            _injecter_ligne(
                ws_an, carte, row, noms,
                identite, idx_prec, ws_mods,
                ws_s1=ws_s1, ws_s2=ws_s2,
                label="Année", lang=lang
            )

            _recalc_ann(ws_an, carte, row, noms, wb=wb, lang=lang)

    # ── Sauvegarde ────────────────────────────────────────────────
    dest = chemin_sortie or chemin_xlsx
    wb.save(dest)

    print(f"\n  ✓ Injection terminée → {dest}")
    return dest