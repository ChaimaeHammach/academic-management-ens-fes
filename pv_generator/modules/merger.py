

import pandas as pd


# ─── Fusion d'un semestre ─────────────────────────────────────────────────────

def fusionner_semestre(modules: list[dict]) -> pd.DataFrame:
    """
    Fusionne les DataFrames de tous les modules d'un semestre.

    Logique :
      1. Collecter toutes les identités connues (CLE → Massar/CIN/Nom/Prénom)
         depuis TOUS les modules non-vides.
      2. Construire la liste complète de CLE (union de toutes les fiches).
      3. Pour chaque module non-vide : joindre ses colonnes de notes.
      4. Appliquer les identités collectées sur toutes les lignes.
      5. Modules vides → colonnes ajoutées avec NA (pas d'étudiants ignorés).
    """
    if not modules:
        raise ValueError("Aucun module fourni.")

    cols_identite = ["CLE", "Massar", "CIN", "Nom", "Prénom"]

    # ── Étape 1 : collecter toutes les identités ───────────────────────────
    # dict CLE → {Massar, CIN, Nom, Prénom} depuis tous les modules
    identites: dict[str, dict] = {}
    for mod in modules:
        df = mod["df"]
        if len(df) == 0:
            continue
        for _, row in df.iterrows():
            cle = row.get("CLE")
            if pd.isna(cle) or cle is None:
                continue
            cle = str(cle).strip()
            if cle not in identites:
                identites[cle] = {}
            for col in ["Massar", "CIN", "Nom", "Prénom"]:
                val = row.get(col)
                if col not in identites[cle] or pd.isna(identites[cle].get(col, None)):
                    if val is not None and not (isinstance(val, float) and pd.isna(val)):
                        identites[cle][col] = val

    if not identites:
        # Tous les modules sont vides → retourner un DataFrame vide structuré
        print(f"  ⚠ Tous les modules sont vides.")
        cols_vides = cols_identite.copy()
        for mod in modules:
            nom = mod["nom_module"]
            for suf in ["_Note_Finale", "_AV", "_Decision", "_Validé", "_Annee_Val"]:
                cols_vides.append(f"{nom}{suf}")
        return pd.DataFrame(columns=cols_vides)

    # ── Étape 2 : DataFrame de base avec toutes les CLEs ──────────────────
    df_base = pd.DataFrame({"CLE": list(identites.keys())})

    # ── Étape 3 : joindre les notes de chaque module ──────────────────────
    noms_modules_vides = []
    for mod in modules:
        nom = mod["nom_module"]
        df  = mod["df"].copy()

        rename = {
            "Moyenne":           f"{nom}_Moyenne",
            "Note_Finale":       f"{nom}_Note_Finale",
            "Decision":          f"{nom}_Decision",
            "Decision_initiale": f"{nom}_Decision_initiale",
            "Validé":            f"{nom}_Validé",
            "AV":                f"{nom}_AV",
            "Annee_Val":         f"{nom}_Annee_Val",
        }
        df.rename(columns=rename, inplace=True)
        cols_notes = [c for c in df.columns if c.startswith(f"{nom}_")]

        if len(df) == 0:
            print(f"  ⚠ Module '{nom}' : aucun étudiant (colonnes vides ajoutées)")
            noms_modules_vides.append(nom)
            for col in cols_notes:
                df_base[col] = pd.NA
            continue

        # Joindre les notes (outer pour conserver les CLEs qui manquent dans ce module)
        df_base = pd.merge(
            df_base,
            df[["CLE"] + cols_notes],
            on="CLE",
            how="left"   # left : on garde toutes les CLEs déjà connues
        )

    # ── Étape 4 : appliquer les identités collectées ───────────────────────
    for col in ["Massar", "CIN", "Nom", "Prénom"]:
        df_base[col] = df_base["CLE"].map(
            lambda cle: identites.get(str(cle), {}).get(col, None)
        )

    # ── Ordre des colonnes : identité d'abord ─────────────────────────────
    cols_id_present = [c for c in cols_identite if c in df_base.columns]
    cols_notes_all  = [c for c in df_base.columns if c not in cols_identite]
    df_base = df_base[cols_id_present + cols_notes_all]
    df_base.reset_index(drop=True, inplace=True)

    return df_base


# ─── Compléter depuis le PV précédent ────────────────────────────────────────

def completer_depuis_pv_precedent(
    df: pd.DataFrame,
    df_pv_precedent: pd.DataFrame,
    noms_modules: list[str]
) -> pd.DataFrame:
    """
    Pour chaque étudiant × module où la note est manquante :
      - Cherche la note dans le PV de l'année précédente.
      - Si introuvable → laisse la case vide (NA).
    """
    if df_pv_precedent is None or len(df_pv_precedent) == 0:
        return df

    df = df.copy()

    # Normaliser les CLEs pour éviter mismatches casse/espaces
    def _norm_cle(s):
        return str(s).strip().upper() if pd.notna(s) and s is not None else ""

    df["CLE"] = df["CLE"].apply(_norm_cle)
    df_pv_precedent = df_pv_precedent.copy()
    df_pv_precedent["CLE"] = df_pv_precedent["CLE"].apply(_norm_cle)

    # Normaliser aussi les noms de modules dans le pivot (correspondance partielle)
    def _norm_mod(s):
        import re as _re
        return _re.sub(r"[^a-z0-9]", "", str(s).lower().strip())

    # Pivot du PV précédent : CLE → {Module: Note}
    pivot = df_pv_precedent.pivot_table(
        index="CLE", columns="Module", values="Note", aggfunc="max"
    ).reset_index()
    pivot.columns.name = None

    # Construire un mapping nom_module_norm → nom_colonne_pivot
    pivot_cols_norm = {_norm_mod(c): c for c in pivot.columns if c != "CLE"}

    # Compléter les notes manquantes module par module
    for mod in noms_modules:
        col_nf = f"{mod}_Note_Finale"
        # Correspondance exacte d'abord, puis partielle (8 premiers chars)
        pivot_col = mod
        if pivot_col not in pivot.columns:
            mod_n = _norm_mod(mod)
            pivot_col = pivot_cols_norm.get(mod_n)
            if pivot_col is None:
                # Correspondance partielle
                for pn, pc in pivot_cols_norm.items():
                    if len(mod_n) >= 6 and (mod_n[:8] in pn or pn[:8] in mod_n):
                        pivot_col = pc
                        break
        if pivot_col is None or pivot_col not in pivot.columns:
            continue

        notes_hist = pivot[["CLE", pivot_col]].rename(columns={pivot_col: f"{mod}_hist"})
        df = pd.merge(df, notes_hist, on="CLE", how="left")

        if col_nf not in df.columns:
            df[col_nf] = pd.NA

        mask_vide = df[col_nf].isna()
        df.loc[mask_vide, col_nf] = df.loc[mask_vide, f"{mod}_hist"]
        df.loc[mask_vide & df[col_nf].notna(), f"{mod}_Validé"] = (
            df.loc[mask_vide & df[col_nf].notna(), col_nf].apply(
                lambda n: pd.notna(n) and float(n) >= 10
            )
        )
        df.drop(columns=[f"{mod}_hist"], inplace=True, errors="ignore")

    return df


# ─── Point d'entrée principal ─────────────────────────────────────────────────

def construire_tableau(
    modules_si: list[dict],
    modules_sp: list[dict],
    df_pv_precedent: pd.DataFrame | None = None,
) -> tuple[pd.DataFrame, list[str], list[str]]:
    """
    Fusionne tous les modules SI et SP.
    Retourne (df_final, noms_si, noms_sp).
    """
    noms_si = [m["nom_module"] for m in modules_si]
    noms_sp = [m["nom_module"] for m in modules_sp]
    tous    = noms_si + noms_sp

    print("\n" + "="*55)
    print("  FUSION DES MODULES")
    print("="*55)

    df_si = fusionner_semestre(modules_si) if modules_si else pd.DataFrame()
    df_sp = fusionner_semestre(modules_sp) if modules_sp else pd.DataFrame()

    if df_si.empty and df_sp.empty:
        raise ValueError("Aucun module chargé.")
    elif df_si.empty:
        df = df_sp
    elif df_sp.empty:
        df = df_si
    else:
        # Fusionner SI et SP : outer join pour garder tous les étudiants
        cols_sp_notes = [c for c in df_sp.columns
                         if c not in ["CLE", "Massar", "CIN", "Nom", "Prénom"]]
        df = pd.merge(df_si, df_sp[["CLE"] + cols_sp_notes], on="CLE", how="outer")

        # Récupérer les identités manquantes depuis df_sp après outer join
        cols_id_sp = [c for c in ["Massar", "CIN", "Nom", "Prénom"] if c in df_sp.columns]
        if cols_id_sp:
            id_sp = df_sp[["CLE"] + cols_id_sp].copy()
            df = pd.merge(df, id_sp, on="CLE", how="left", suffixes=("", "_sp"))
            for col in cols_id_sp:
                col_sp = f"{col}_sp"
                if col_sp in df.columns:
                    df[col] = df[col].combine_first(df[col_sp])
                    df.drop(columns=[col_sp], inplace=True, errors="ignore")

    # Compléter les notes manquantes depuis le PV précédent
    if df_pv_precedent is not None:
        df = completer_depuis_pv_precedent(df, df_pv_precedent, tous)

    # Colonnes identité en premier
    cols_id    = ["CLE", "Massar", "CIN", "Nom", "Prénom"]
    cols_notes = [c for c in df.columns if c not in cols_id]
    df = df[[c for c in cols_id if c in df.columns] + cols_notes]
    df.reset_index(drop=True, inplace=True)

    print(f"  Total étudiants : {len(df)}")
    print(f"  Modules SI : {noms_si}")
    print(f"  Modules SP : {noms_sp}")
    print("="*55 + "\n")

    return df, noms_si, noms_sp


# ─── Détection des absents ────────────────────────────────────────────────────

def detecter_absents(df: pd.DataFrame, modules: list[str]) -> pd.DataFrame:
    """Retourne les étudiants avec au moins une note manquante."""
    alertes = []
    for _, row in df.iterrows():
        for mod in modules:
            nf = row.get(f"{mod}_Note_Finale")
            if nf is None or (isinstance(nf, float) and pd.isna(nf)):
                alertes.append({
                    "Massar":   row.get("Massar", ""),
                    "CIN":      row.get("CIN", ""),
                    "Nom":      row.get("Nom", ""),
                    "Module":   mod,
                    "Problème": "Note manquante",
                })
    return pd.DataFrame(alertes) if alertes else pd.DataFrame(
        columns=["Massar", "CIN", "Nom", "Module", "Problème"]
    )
