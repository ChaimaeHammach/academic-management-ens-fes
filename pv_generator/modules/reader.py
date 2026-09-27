
import pandas as pd
import os
import re
from openpyxl import load_workbook




VARIANTES = {
    "massar": [
        "massar", "n°massar", "nummassar", "n° massar", "no massar",
        "cne", "n° cne", "n°cne",
        "الرقم الوطني", "رقم الوطني", "رقم مسار", "رقم",
    ],
    "cin": [
        "cin", "cin.", "carte nationale", "ب-ت-و",
        "رقم البطاقة الوطنية", "البطاقة الوطنية", "رقم البطاقة",
        "رقم بطاقة", "بطاقة وطنية",
    ],
    "nom": ["nom", "nom de famille", "lastname", "النسب", "الاسم العائلي", "اسم العائلة"],
    "prenom": ["prénom", "prenom", "firstname", "الاسم", "الاسم الشخصي", "اسم شخصي"],
    "no": ["no", "n°", "num", "numéro", "#"],
    "moyenne": [
        "moyenne finale", "moyennefinale", "moyenne final",
        "moyenne", "moy", "moy.", "moyenne module", "moy module",
        "moyenne générale", "moyennegénérale", "moyennemodule",
        "moyenne annuelle", "moyenneannuelle",
        "moyenne initiale", "moyenneinitiale",
        "المعدل", "معدل", "المعدل النهائي", "المعدل الاجمالي",
    ],
    "decision_finale": [
        "décision finale", "decision finale",
        "déc finale", "dec finale",
        "décisionfinal", "decisionfinal",
        "décision final", "decision final",
        "القرار النهائي", "قرار نهائي",
    ],
    "decision": [
        "décision initiale", "decision initiale",
        "déc initiale", "dec initiale",
        "décision", "decision", "déc", "dec", "déc.", "d1",
        "النتيجة", "النتيجه", "نتيجة", "نتيجه", "القرار",
    ],
    "note_finale": [
        "note finale", "notefinal", "note final", "note_finale",
        "résultat", "note final d'élément", "note final element",
        "العلامة النهائية", "علامة نهائية", "العلامة", "النتيجة النهائية",
    ],
    "rattrapage": ["rattrapage", "ratt", "rattrap", "note rattrapage"],
    "av": [
        "av", "a/v", "appréciation", "appreciation",
        "année validée", "annee validee",
        "الاستدراك", "استدراك",
    ],
    "annee_val": [
        "سنة", "سنه",
    ],
}

# Mots-clés qui indiquent que c'est un en-tête répété parasite (pas un étudiant)
MOTS_CLES_PARASITES = {
    "massar", "cin", "nom", "prénom", "prenom", "no", "n°",
    "nan", "none", "", "cne",
}

# Valeurs des cellules qui correspondent à des colonnes "décision validée"
DECISIONS_VALIDE = {"V", "VAR", "VALIDÉ", "VALIDE", "س.م", "س م"}


def _norm(s: str) -> str:
    """Normalise une chaîne : minuscules, sans espaces/tirets/points."""
    return re.sub(r"[\s_\-\.]+", "", str(s).strip().lower())


def _is_na(v) -> bool:
    if v is None:
        return True
    try:
        return bool(pd.isna(v))
    except (TypeError, ValueError):
        return False


def _est_numerique_ou_pct(s: str) -> bool:
    
    s = s.strip().rstrip("%")
    try:
        float(s)
        return True
    except ValueError:
        return False


def _trouver_col(df: pd.DataFrame, cle: str) -> str | None:
    norm_map = {_norm(c): c for c in df.columns}
    for variante in VARIANTES.get(cle, []):
        if _norm(variante) in norm_map:
            return norm_map[_norm(variante)]
    return None


def _normaliser_annee_val(v) -> str:
   
    import re as _re
    if _is_na(v):
        return ""
    s = str(v).strip()
    if s in ("nan", "None", ""):
        return ""

    # Déjà sous forme "AA-BB" ou "AAAA-BBBB" → formater en "AA-BB"
    m = _re.match(r'^(\d{2,4})-(\d{2,4})$', s)
    if m:
        y1, y2 = m.group(1), m.group(2)
        return f"{y1[-2:]}-{y2[-2:]}"  # normalise "2025-2026" → "25-26"

    # Valeur numérique : probablement un artefact (soustraction, 0, ROW(), etc.)
    try:
        f = float(s)
        # Valeur numérique négative ou faible → artefact de eval("25-26") = -1
        # ou valeur 0 (formule non calculée) → on rejette
        if f <= 0 or (0 < f < 100 and not _re.match(r'^\d{4}$', s)):
            # Sauf si c'est une année sur 4 chiffres (ex: 2025) → improbable ici
            return ""
        # Entier positif petit (<= 50) → artefact (numéro de ligne, etc.)
        if f == int(f) and int(f) <= 50:
            return ""
    except (ValueError, TypeError):
        pass

    # Texte quelconque non reconnu → retourner tel quel
    return s


# ─── Lecture brute avec openpyxl ─────────────────────────────────────────────

def _choisir_feuille_pv(wb):
    
    for nom in wb.sheetnames:
        if "pv" in nom.lower():
            return wb[nom]
    return wb.active


def _lire_brut(chemin: str, max_rows: int = 15, feuille: str | None = None) -> list[list]:
    
    wb = load_workbook(chemin, data_only=True)
    ws = wb[feuille] if feuille and feuille in wb.sheetnames else _choisir_feuille_pv(wb)

    merged_values: dict[tuple, object] = {}
    for mr in ws.merged_cells.ranges:
        top_val = ws.cell(mr.min_row, mr.min_col).value
        for r in range(mr.min_row, mr.max_row + 1):
            for c in range(mr.min_col, mr.max_col + 1):
                merged_values[(r, c)] = top_val

    rows = []
    for row_idx, row in enumerate(
        ws.iter_rows(max_row=min(max_rows, ws.max_row)), start=1
    ):
        row_data = [
            merged_values.get((row_idx, col_idx + 1), cell.value)
            for col_idx, cell in enumerate(row)
        ]
        rows.append(row_data)
    return rows


def _evaluer_formule(formule: str, ws, row_idx: int, cache: dict) -> object:
    
    import re
    if not isinstance(formule, str) or not formule.startswith('='):
        return formule

    expr = formule[1:]

    # Refuser les fonctions complexes (IF, ISNUMBER, ROUND, etc.)
    if re.search(r'[A-Z]{2,}\s*\(', expr):
        return None

    # Si la formule est une simple référence de cellule (ex: =$X$7, =A5, ='Sheet'!B3)
    # → retourner directement la valeur source sans évaluation arithmétique
    ref_simple = re.fullmatch(
        r"'?[^!']*'?!?\$?([A-Z]{1,3})\$?(\d+)", expr.strip()
    )
    if ref_simple:
        col_str = ref_simple.group(1)
        r = int(ref_simple.group(2))
        try:
            from openpyxl.utils import column_index_from_string
            c = column_index_from_string(col_str)
            v = ws.cell(r, c).value
            if isinstance(v, str) and v.startswith('='):
                v = _evaluer_formule(v, ws, r, cache)
            cache[(r, col_str)] = v
            return v  # ← valeur brute, pas de conversion float
        except Exception:
            return None

    # Pour les expressions arithmétiques pures : vérifier que toutes les
    # références pointent vers des nombres avant d'évaluer.
    raw_refs: dict[str, object] = {}

    def get_ref(m):
        col_str = m.group(1).replace('$', '')
        row_str = m.group(2).replace('$', '')
        r = int(row_str) if row_str.isdigit() else row_idx
        key = (r, col_str)
        if key in cache:
            v = cache[key]
        else:
            try:
                from openpyxl.utils import column_index_from_string
                c = column_index_from_string(col_str)
                v = ws.cell(r, c).value
                if isinstance(v, str) and v.startswith('='):
                    v = _evaluer_formule(v, ws, r, cache)
                cache[key] = v
            except Exception:
                v = None
        raw_refs[m.group(0)] = v
        # Si la valeur n'est pas numérique → on ne peut pas l'utiliser dans eval
        if v is None:
            return '0'
        try:
            float(v)
            return str(v)
        except (TypeError, ValueError):
            # Valeur texte dans une expression arithmétique → abandon
            raw_refs['__has_text__'] = True
            return '0'

    expr2 = re.sub(r'(\$?[A-Z]{1,3})(\$?\d+)', get_ref, expr)

    # Si au moins une référence était du texte, ne pas évaluer
    if raw_refs.get('__has_text__'):
        # Retourner la première valeur non-None trouvée (cas référence unique avec tiret)
        for v in raw_refs.values():
            if v and v != '__has_text__':
                return v
        return None

    try:
        result = eval(expr2)  # noqa: S307
        return round(float(result), 4)
    except Exception:
        return None


def _lire_donnees_brut(chemin: str, debut_ligne: int, n_cols: int,
                       feuille: str | None = None) -> list[list]:
   
    # Charger avec valeurs calculées
    wb_val = load_workbook(chemin, data_only=True)
    ws_val = (wb_val[feuille] if feuille and feuille in wb_val.sheetnames
              else _choisir_feuille_pv(wb_val))

    # Charger avec formules pour fallback
    wb_frm = load_workbook(chemin, data_only=False)
    ws_frm = (wb_frm[feuille] if feuille and feuille in wb_frm.sheetnames
              else _choisir_feuille_pv(wb_frm))

    merged_values: dict[tuple, object] = {}
    for mr in ws_val.merged_cells.ranges:
        top_val = ws_val.cell(mr.min_row, mr.min_col).value
        for r in range(mr.min_row, mr.max_row + 1):
            for c in range(mr.min_col, mr.max_col + 1):
                merged_values[(r, c)] = top_val

    eval_cache: dict = {}
    rows = []
    for row_idx, row in enumerate(
        ws_val.iter_rows(min_row=debut_ligne, values_only=False), start=debut_ligne
    ):
        row_data = []
        for col_idx, cell in enumerate(row):
            val = merged_values.get((row_idx, col_idx + 1), cell.value)

            # Cas 1 : valeur None → essayer d'évaluer la formule
            if val is None:
                frm_cell = ws_frm.cell(row_idx, col_idx + 1)
                if isinstance(frm_cell.value, str) and frm_cell.value.startswith('='):
                    val = _evaluer_formule(frm_cell.value, ws_frm, row_idx, eval_cache)

            # Cas 2 : valeur numérique dans une cellule formatée "@" (texte forcé)
            # ou dont le number_format ressemble à "texte" → conserver comme str
            # Exemple : "25-26" stocké dans une cellule format "@" peut revenir
            # comme entier si Excel l'a interprété comme date.
            if val is not None and not isinstance(val, str):
                nf = cell.number_format or ""
                # Format texte forcé : "@"
                if nf == "@":
                    val = str(val)
                # Format "dd-yy" ou similaire (date stockée comme "25-26")
                # → si la valeur numérique correspond à un numéro de série de date
                # improbable (< 1000 ou pattern NN-NN), conserver comme texte
                # Ne pas toucher aux vraies notes (0-20) sauf si format texte

            row_data.append(val)

        # Tronquer/padder à n_cols
        row_data = row_data[:n_cols] + [None] * max(0, n_cols - len(row_data))
        rows.append(row_data)

    wb_val.close()
    wb_frm.close()
    return rows


# ─── Détection de la ligne identité ──────────────────────────────────────────

def _trouver_ligne_identite(rows: list[list]) -> int:
    
    var_id = (
        VARIANTES["massar"] + VARIANTES["cin"] +
        VARIANTES["nom"]    + VARIANTES["prenom"]
    )

    def _est_pct_pur(row):
        
        vals = [str(v).strip() for v in row if v is not None and str(v).strip()]
        return len(vals) > 0 and all(_est_numerique_ou_pct(v) for v in vals)

    def _est_ligne_donnees(row):
       
        mots_labels = (
            VARIANTES["massar"] + VARIANTES["cin"] + VARIANTES["nom"] +
            VARIANTES["prenom"] + VARIANTES["moyenne"] + VARIANTES["decision_finale"] +
            VARIANTES["decision"] + VARIANTES["av"] + VARIANTES["no"] +
            ["cc", "contrôle", "controle", "examen", "exam", "note", "rattrapage",
             "tp", "td", "élément", "element", "m1", "m2", "m21", "m22"]
        )
        vals_texte = [str(v).strip() for v in row
                      if v is not None and isinstance(v, str) and str(v).strip()]
        vals_num   = [v for v in row if v is not None and isinstance(v, (int, float))]

        if not vals_texte and vals_num:
            return True  # que des nombres → données

        # Si tous les textes sont des labels connus → c'est une ligne d'en-tête
        nb_labels = sum(
            1 for v in vals_texte
            if any(_norm(lab) in _norm(v) or _norm(v) in _norm(lab)
                   for lab in mots_labels)
        )
        nb_texte = len(vals_texte)
        if nb_texte > 0 and nb_labels / nb_texte >= 0.5:
            return False  # majoritairement des labels → en-tête

        # Textes qui ressemblent à des noms propres ou numéros (données réelles)
        import re
        nb_noms_propres = sum(
            1 for v in vals_texte
            if re.match(r'^[A-ZÀ-Ü][A-ZÀ-Üa-zà-ü\- ]{2,}$', v) or
               re.match(r'^[A-Z]{1,2}\d{6,}$', v)  # CIN comme CD785623
        )
        if nb_noms_propres >= 2 and vals_num:
            return True  # noms propres + notes → données

        return False

    candidats = []
    for i, row in enumerate(rows):
        if _est_pct_pur(row):
            continue
        if _est_ligne_donnees(row):
            continue
        valeurs = [_norm(str(v)) for v in row if v is not None]
        hits = sum(
            1 for v in valeurs
            if any(_norm(var) in v or v in _norm(var) for var in var_id)
        )
        if hits >= 2:
            candidats.append(i)

    if candidats:
        return candidats[-1]   # Prendre la dernière ligne candidate (plus proche des données)

    # Fallback : ligne avec le plus de mots-clés de notes
    mots = ["cc", "contrôle", "controle", "examen", "exam", "tp",
            "moyenne", "moy", "décision", "decision", "note"]
    best_i, best_h = 0, 0
    for i, row in enumerate(rows):
        if _est_pct_pur(row):
            continue
        valeurs = [_norm(str(v)) for v in row if v is not None]
        h = sum(1 for v in valeurs if any(m in v for m in mots))
        if h > best_h:
            best_h, best_i = h, i
    return best_i


# ─── Construction des en-têtes combinés ──────────────────────────────────────

def _construire_entetes(rows: list[list], ligne_id: int) -> list[str]:
    """
    Pour chaque colonne :
      - Prend la valeur de ligne_id si elle est non vide et non numérique/pourcentage
      - Sinon, remonte ligne par ligne (jusqu'à 4 niveaux) pour trouver une valeur utile
      - Exclut les valeurs qui sont uniquement numériques ou des pourcentages
      - Exclut les valeurs qui ressemblent à des titres d'en-tête méta (école, année...)
    """
    n_cols = max(len(r) for r in rows)
    # Padder toutes les lignes
    padded = [r + [None] * (n_cols - len(r)) for r in rows]

    # Mots à exclure (en-têtes méta)
    mots_meta = {"universitaire", "licence", "semestre", "module", "école",
                 "ecole", "année", "annee", "spécialité", "specialite"}

    def _est_valeur_utile(v) -> bool:
        if v is None:
            return False
        s = str(v).strip()
        if not s or s in ("None", "nan"):
            return False
        if _est_numerique_ou_pct(s):
            return False
        # Exclure les longues chaînes méta
        s_norm = s.lower()
        if any(m in s_norm for m in mots_meta) and len(s) > 30:
            return False
        return True

    combined = []
    for col_idx in range(n_cols):
        val_id = padded[ligne_id][col_idx]
        if _est_valeur_utile(val_id):
            combined.append(str(val_id).strip())
            continue

        # Remonter jusqu'à 4 lignes au-dessus
        found = ""
        for offset in range(1, min(5, ligne_id + 1)):
            above = padded[ligne_id - offset][col_idx]
            if _est_valeur_utile(above):
                found = str(above).strip()
                break
        combined.append(found if found else f"_Col{col_idx}")

    # Dédupliquer
    seen: dict[str, int] = {}
    result = []
    for h in combined:
        if h in seen:
            seen[h] += 1
            result.append(f"{h}.{seen[h]}")
        else:
            seen[h] = 0
            result.append(h)
    return result


# ─── Lecture méta (école, semestre, module) ───────────────────────────────────

def _lire_entete_meta(chemin: str) -> dict:
    raw = pd.read_excel(chemin, header=None, nrows=6)
    info = {"annee": "", "semestre": "", "session": "Ordinaire", "module": ""}

    for _, row in raw.iterrows():
        ligne = " ".join(str(v) for v in row if pd.notna(v) and str(v).strip())
        ll = ligne.lower()

        if re.search(r"\d{4}-\d{4}", ligne):
            m = re.search(r"\d{4}-\d{4}", ligne)
            if m:
                info["annee"] = m.group()

        if "semestre" in ll:
            m = re.search(r"s\d", ll)
            if m:
                info["semestre"] = m.group().upper()
            if "rattrapage" in ll:
                info["session"] = "Rattrapage"
            elif "ordinaire" in ll:
                info["session"] = "Ordinaire"

        if "module" in ll:
            m = re.search(r"module\s*[:\-]?\s*(.+)", ligne, re.IGNORECASE)
            if m:
                c = m.group(1).strip()
                if c and not re.match(r"^\d+$", c):
                    info["module"] = c

    return info


# ─── Lecture principale ───────────────────────────────────────────────────────

def lire_module(chemin: str, nom_module: str | None = None) -> dict:
    """
    Lit un fichier Excel de notes.
    Gère toutes les structures rencontrées : simple, 2 niveaux, 3 niveaux.
    Copie directement les valeurs du prof (Moyenne finale, Décision finale, AV).
    """
    if not os.path.exists(chemin):
        raise FileNotFoundError(f"Fichier introuvable : {chemin}")

    meta = _lire_entete_meta(chemin)
    if nom_module:
        meta["module"] = nom_module
    elif not meta["module"]:
        meta["module"] = os.path.splitext(os.path.basename(chemin))[0]

    # 1. Lire les premières lignes avec fusions résolues
    rows = _lire_brut(chemin, max_rows=15)

    # 2. Trouver la ligne identité (dernière ligne avec Massar/CIN/Nom)
    ligne_id = _trouver_ligne_identite(rows)
    print(f"  [{meta['module']}] Ligne identité : {ligne_id + 1} "
          f"→ {[v for v in rows[ligne_id] if v is not None][:7]}")

    # 3. Construire les en-têtes combinés
    headers = _construire_entetes(rows, ligne_id)
    print(f"  [{meta['module']}] En-têtes : {headers}")

    n_cols = len(headers)

    # 4. Lire les données (à partir de la ligne suivant ligne_id, 1-based)
    debut = ligne_id + 2   # ligne_id est 0-based, +1 pour 1-based, +1 pour passer à la suivante
    data_rows = _lire_donnees_brut(chemin, debut, n_cols)

    df = pd.DataFrame(data_rows, columns=headers)
    df.dropna(how="all", inplace=True)
    df.reset_index(drop=True, inplace=True)

    # Supprimer la ligne de pourcentages (première ligne de données si toutes numériques/pct)
    # On note si on l'a sautée pour corriger l'offset de ligne Excel
    _ligne_pct_sautee = False
    if len(df) > 0:
        first_vals = [
            str(v).strip() for v in df.iloc[0]
            if v is not None and str(v).strip() not in ("", "None", "nan")
        ]
        if first_vals and all(_est_numerique_ou_pct(v) for v in first_vals):
            df = df.iloc[1:].reset_index(drop=True)
            _ligne_pct_sautee = True

    # Ligne Excel réelle du premier étudiant (1-based)
    _xlsx_ligne_debut = debut + (1 if _ligne_pct_sautee else 0)

    # 5. Identifier les colonnes
    col_massar   = _trouver_col(df, "massar")
    col_cin      = _trouver_col(df, "cin")
    col_nom      = _trouver_col(df, "nom")
    col_prenom   = _trouver_col(df, "prenom")
    col_moy      = _trouver_col(df, "moyenne")
    col_dec_fin  = _trouver_col(df, "decision_finale")
    col_dec_init = _trouver_col(df, "decision")
    col_av       = _trouver_col(df, "av")
    col_annee    = _trouver_col(df, "annee_val")

    print(f"  [{meta['module']}] Colonnes → "
          f"Massar={col_massar} | CIN={col_cin} | "
          f"Moy={col_moy} | DécFin={col_dec_fin} | "
          f"DécInit={col_dec_init} | AV={col_av} | Année={col_annee}")

    if col_massar is None and col_cin is None:
        raise ValueError(
            f"Aucune colonne Massar/CIN dans {chemin}.\n"
            f"Colonnes : {list(df.columns)}"
        )
    if col_nom is None:
        raise ValueError(f"Colonne Nom introuvable dans {chemin}.")

    # Si Moyenne Finale absente → dernière colonne numérique significative
    if col_moy is None:
        excl = {_norm(v) for v in VARIANTES["no"] + VARIANTES["massar"] + VARIANTES["cin"]}
        cols_num = [
            c for c in df.columns
            if _norm(c) not in excl
            and pd.to_numeric(df[c], errors="coerce").notna().sum() > len(df) * 0.3
        ]
        if cols_num:
            col_moy = cols_num[-1]
            print(f"  [{meta['module']}] Fallback Moyenne → '{col_moy}'")

    # 6. Construire le DataFrame résultat
    cle_col = col_massar if col_massar else col_cin

    result = pd.DataFrame()
    result["Massar"] = df[col_massar].astype(str).str.strip() if col_massar else pd.NA
    result["CIN"]    = df[col_cin].astype(str).str.strip()    if col_cin    else pd.NA
    result["Nom"]    = df[col_nom].astype(str).str.strip()
    result["Prénom"] = df[col_prenom].astype(str).str.strip() if col_prenom else pd.NA
    result["CLE"]    = df[cle_col].astype(str).str.strip()

    result["Moyenne"]     = pd.to_numeric(df[col_moy], errors="coerce").round(2) if col_moy else pd.NA
    result["Note_Finale"] = result["Moyenne"]

    # Décision finale prioritaire
    if col_dec_fin:
        result["Decision"] = df[col_dec_fin].astype(str).str.strip()
    elif col_dec_init:
        result["Decision"] = df[col_dec_init].astype(str).str.strip()
    else:
        result["Decision"] = pd.NA

    result["Decision_initiale"] = (
        df[col_dec_init].astype(str).str.strip() if col_dec_init else pd.NA
    )

    result["AV"] = (
        df[col_av].apply(_normaliser_annee_val)
        if col_av else pd.NA
    )

    # Colonne سنة (année validée, ex: "25-26")
    result["Annee_Val"] = (
        df[col_annee].apply(_normaliser_annee_val)
        if col_annee else pd.NA
    )

    # Validé : décision prof ou note >= 10
    def _est_valide(row):
        d = str(row.get("Decision", "")).strip().upper()
        if d and d not in ("NAN", "NONE", ""):
            return d in DECISIONS_VALIDE
        n = row.get("Moyenne")
        return not _is_na(n) and float(n) >= 10

    result["Validé"] = result.apply(_est_valide, axis=1)

    # 7. Nettoyage : supprimer les lignes parasites (en-têtes répétés, vides)
    result = result[~result["CLE"].str.lower().str.strip().isin(MOTS_CLES_PARASITES)]
    result = result[result["CLE"].str.len() > 1]
    result = result[~result["CLE"].isin(["nan", "NaN", "None", ""])]
    result.reset_index(drop=True, inplace=True)

    print(f"  [{meta['module']}] ✓ {len(result)} étudiants | "
          f"Décisions: {result['Decision'].value_counts().to_dict()} | "
          f"AV: {result['AV'].unique()[:3]}")

    # Index de colonne Excel (1-based) de la Note Finale (après rattrapage)
    # Priorité : col_note_finale (ex: "Note Finale") > col_dec_fin comme indicateur
    # On cherche d'abord la vraie colonne "note_finale" dans df
    col_note_fin = _trouver_col(df, "note_finale")
    # Si elle n'existe pas (fiche simple sans colonne séparée), fallback sur col_moy
    _col_pour_note = col_note_fin if col_note_fin else col_moy

    _xlsx_col_note = (
        list(df.columns).index(_col_pour_note) + 1
        if _col_pour_note and _col_pour_note in df.columns
        else None
    )

    # Index de colonne Excel (1-based) de la Décision Finale
    _col_pour_dec = col_dec_fin if col_dec_fin else col_dec_init
    _xlsx_col_dec = (
        list(df.columns).index(_col_pour_dec) + 1
        if _col_pour_dec and _col_pour_dec in df.columns
        else None
    )

    # Index de colonne Excel (1-based) de la colonne سنة
    _xlsx_col_annee = (
        list(df.columns).index(col_annee) + 1
        if col_annee and col_annee in df.columns
        else None
    )

    # Index de colonnes Excel (1-based) pour l'identité
    def _xlsx_col(col_name):
        if col_name and col_name in df.columns:
            return list(df.columns).index(col_name) + 1
        return None

    _xlsx_col_cin    = _xlsx_col(col_cin)
    _xlsx_col_massar = _xlsx_col(col_massar)
    _xlsx_col_nom    = _xlsx_col(col_nom)
    _xlsx_col_prenom = _xlsx_col(col_prenom)

    return {
        "nom_module":       meta["module"],
        "annee":            meta["annee"],
        "semestre":         meta["semestre"],
        "session":          meta["session"],
        "df":               result,
        # Métadonnées pour formules Excel (liaisons dynamiques PV ↔ feuilles modules)
        "xlsx_ligne_debut": _xlsx_ligne_debut,   # ligne Excel du 1er étudiant (1-based)
        "xlsx_col_note":    _xlsx_col_note,       # colonne Excel de la Note Finale (1-based)
        "xlsx_col_dec":     _xlsx_col_dec,        # colonne Excel de la Décision Finale (1-based)
        "xlsx_col_annee":   _xlsx_col_annee,      # colonne Excel de سنة (1-based)
        "xlsx_col_cin":     _xlsx_col_cin,        # colonne Excel CIN (1-based)
        "xlsx_col_massar":  _xlsx_col_massar,     # colonne Excel Massar (1-based)
        "xlsx_col_nom":     _xlsx_col_nom,        # colonne Excel Nom (1-based)
        "xlsx_col_prenom":  _xlsx_col_prenom,     # colonne Excel Prénom (1-based)
    }


# ─── PV précédent ─────────────────────────────────────────────────────────────

def lire_pv_precedent(chemin: str) -> pd.DataFrame:
    """
    Lit un fichier PV (généré par pv_generator ou autre format PV standard).
    
    Stratégie :
    1. Choisir la feuille PV (nom contenant 'PV', sinon feuille active)
    2. Repérer la ligne d'en-têtes (contient CIN/Massar, Nom, Prénom)
    3. Repérer la ligne de groupes juste au-dessus (noms des modules)
    4. Mapper groupe → colonnes Moy. pour nommer les notes par vrai nom de module
    5. Extraire les données étudiant + note par module
    """
    if not os.path.exists(chemin):
        raise FileNotFoundError(f"PV précédent introuvable : {chemin}")

    wb = load_workbook(chemin, data_only=True)
    ws = _choisir_feuille_pv(wb)

    # ── 1. Résoudre les cellules fusionnées ──────────────────────────────────
    merged_values: dict[tuple, object] = {}
    for mr in ws.merged_cells.ranges:
        top_val = ws.cell(mr.min_row, mr.min_col).value
        for r in range(mr.min_row, mr.max_row + 1):
            for c in range(mr.min_col, mr.max_col + 1):
                merged_values[(r, c)] = top_val

    def cell_val(r, c):
        return merged_values.get((r, c), ws.cell(r, c).value)

    # ── 2. Trouver la ligne d'en-têtes (CIN/Massar + Nom + Prénom) ──────────
    cle_kw   = set(VARIANTES["massar"] + VARIANTES["cin"])
    nom_kw   = set(VARIANTES["nom"])
    prenom_kw = set(VARIANTES["prenom"])

    ligne_hdr = None
    for r in range(1, min(10, ws.max_row + 1)):
        row_norm = [_norm(str(cell_val(r, c) or "")) for c in range(1, ws.max_column + 1)]
        has_cle  = any(v in cle_kw   for v in row_norm)
        has_nom  = any(v in nom_kw   for v in row_norm)
        has_pre  = any(v in prenom_kw for v in row_norm)
        if has_cle and has_nom and has_pre:
            ligne_hdr = r
            break

    if ligne_hdr is None:
        raise ValueError("Aucune colonne Massar/CIN dans le PV précédent.")

    # ── 3. Identifier les colonnes d'identité dans la ligne hdr ─────────────
    n_cols = ws.max_column
    col_massar = col_cin = col_nom = col_prenom = None
    for c in range(1, n_cols + 1):
        vn = _norm(str(cell_val(ligne_hdr, c) or ""))
        if col_massar is None and vn in set(VARIANTES["massar"]):
            col_massar = c
        elif col_cin is None and vn in set(VARIANTES["cin"]):
            col_cin = c
        elif col_nom is None and vn in set(VARIANTES["nom"]):
            col_nom = c
        elif col_prenom is None and vn in set(VARIANTES["prenom"]):
            col_prenom = c

    cle_col = col_massar if col_massar else col_cin

    # ── 4. Mapper groupe → colonnes note (ligne hdr-1 = groupes modules) ────
    # La ligne au-dessus de hdr contient les noms de modules fusionnés
    # ex: E2="Analyse 1..." fusionné sur E-G, donc E3=Moy., F3=AV, G3=Déc.
    # On veut : colonne "Moy." → nom du module correspondant
    ligne_grp = ligne_hdr - 1
    col_to_module: dict[int, str] = {}   # {col_note: nom_module}
    moy_norm = {"moy", "moy.", "moyenne", "note", "note finale"}

    if ligne_grp >= 1:
        current_module = None
        for c in range(1, n_cols + 1):
            # Valeur groupe (après résolution fusions)
            grp_val = cell_val(ligne_grp, c)
            if grp_val and _norm(str(grp_val)) not in {
                _norm("Identité de l'étudiant(e)"), "identitédel'étudiant(e)",
                "identite", "semestres1", "semestres2", "semestre", ""
            }:
                nom_grp = str(grp_val).strip()
                # Exclure les en-têtes génériques
                if not any(kw in _norm(nom_grp) for kw in ["identit", "semestre", "avs", "décsi"]):
                    current_module = nom_grp

            # Colonne sous-entête (ligne_hdr)
            hdr_val = _norm(str(cell_val(ligne_hdr, c) or ""))
            if hdr_val in moy_norm and current_module:
                col_to_module[c] = current_module

    # ── 5. Identifier les colonnes note : toutes celles de type Moy. ─────────
    # Si le mapping groupe→module est vide, fallback sur colonnes numériques
    excl_cols = {col_massar, col_cin, col_nom, col_prenom}

    if not col_to_module:
        # Fallback : toutes les colonnes numériques hors identité
        for c in range(1, n_cols + 1):
            if c in excl_cols:
                continue
            hdr_n = _norm(str(cell_val(ligne_hdr, c) or ""))
            if hdr_n in moy_norm:
                col_to_module[c] = f"Module_{c}"

    # ── 6. Lire les données ──────────────────────────────────────────────────
    data_debut = ligne_hdr + 1

    rows_out = []
    for r in range(data_debut, ws.max_row + 1):
        cle_val = cell_val(r, cle_col) if cle_col else None
        if not cle_val or str(cle_val).strip() in ("", "None", "nan"):
            continue
        cle = str(cle_val).strip()

        massar = str(cell_val(r, col_massar)).strip() if col_massar else pd.NA
        cin    = str(cell_val(r, col_cin)).strip()    if col_cin    else pd.NA
        nom    = str(cell_val(r, col_nom)).strip()    if col_nom    else pd.NA
        prenom = str(cell_val(r, col_prenom)).strip() if col_prenom else pd.NA

        # Masquer les "None"/"nan" résiduels
        massar = pd.NA if massar in ("None","nan","") else massar
        cin    = pd.NA if cin    in ("None","nan","") else cin
        nom    = pd.NA if nom    in ("None","nan","") else nom
        prenom = pd.NA if prenom in ("None","nan","") else prenom

        for c_note, nom_mod in col_to_module.items():
            raw = cell_val(r, c_note)
            note = pd.to_numeric(raw, errors="coerce")
            if pd.isna(note) or note > 20 or note < 0:
                # Ignorer les notes aberrantes (ex: >20) ou manquantes
                continue
            rows_out.append({
                "CLE": cle, "Massar": massar, "CIN": cin,
                "Nom": nom, "Prénom": prenom,
                "Module": nom_mod.strip(), "Note": round(float(note), 2),
            })

    if not rows_out:
        raise ValueError("Aucune note extraite du PV précédent.")

    result = pd.DataFrame(rows_out)
    result = result[result["CLE"].str.len() > 1]
    result = result[~result["CLE"].isin(["nan", "NaN", "None", ""])]
    result.reset_index(drop=True, inplace=True)
    print(f"  PV précédent : {result['CLE'].nunique()} étudiants, "
          f"{result['Module'].nunique()} modules")
    return result


# ─── Redoublants ─────────────────────────────────────────────────────────────

def lire_redoublants(chemin: str) -> pd.DataFrame:
    if not os.path.exists(chemin):
        raise FileNotFoundError(f"Fichier redoublants introuvable : {chemin}")
    df = pd.read_excel(chemin, header=0)
    df.dropna(how="all", inplace=True)
    df.columns = [c.strip() for c in df.columns]
    col_massar = _trouver_col(df, "massar")
    col_cin    = _trouver_col(df, "cin")
    col_nom    = _trouver_col(df, "nom")
    col_prenom = _trouver_col(df, "prenom")
    col_module = next((c for c in df.columns if "module" in c.lower()), None)
    col_note   = next((c for c in df.columns if "note" in c.lower()), None)
    if (col_massar is None and col_cin is None) or col_module is None or col_note is None:
        raise ValueError(f"Colonnes manquantes dans redoublants : {list(df.columns)}")
    cle_col = col_massar if col_massar else col_cin
    result  = pd.DataFrame()
    result["CLE"]    = df[cle_col].astype(str).str.strip()
    result["Massar"] = df[col_massar].astype(str).str.strip() if col_massar else pd.NA
    result["CIN"]    = df[col_cin].astype(str).str.strip()    if col_cin    else pd.NA
    result["Nom"]    = df[col_nom].astype(str).str.strip()    if col_nom    else pd.NA
    result["Prénom"] = df[col_prenom].astype(str).str.strip() if col_prenom else pd.NA
    result["Module"] = df[col_module].astype(str).str.strip()
    result["Note"]   = pd.to_numeric(df[col_note], errors="coerce")
    result = result[result["CLE"].str.len() > 1]
    result = result[~result["CLE"].isin(["nan", "NaN", "None", ""])]
    print(f"  Redoublants : {len(result)} entrées ({result['CLE'].nunique()} étudiants)")
    return result


def lire_modules_semestre(fichiers: dict[str, str]) -> list[dict]:
    modules = []
    for nom, chemin in fichiers.items():
        try:
            modules.append(lire_module(chemin, nom))
        except Exception as e:
            print(f"  ERREUR {nom} : {e}")
    return modules