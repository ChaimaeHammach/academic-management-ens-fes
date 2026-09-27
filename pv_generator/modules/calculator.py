

import pandas as pd


def mention(moyenne: float | None) -> str:
    if moyenne is None or pd.isna(moyenne):
        return ""
    if moyenne >= 16: return "Très Bien"
    if moyenne >= 14: return "Bien"
    if moyenne >= 12: return "Assez Bien"
    if moyenne >= 10: return "Passable"
    return ""


# Valeurs non-numériques à ignorer (symboles affichés dans les PV)
_VALEURS_VIDES = {"—", "-", "—", "N/A", "NA", "ABS", "", "None", "nan"}

def _note_valide(note) -> bool:
    """Retourne True si la note est un nombre utilisable (non None, non NA, non NaN, non tiret)."""
    if note is None:
        return False
    if isinstance(note, str) and note.strip() in _VALEURS_VIDES:
        return False
    try:
        float(note)
        return True
    except (TypeError, ValueError):
        return False


def moyenne_semestre(
    row: pd.Series,
    modules: list[str],
    coefficients: dict[str, float],
) -> tuple[float | None, int]:
    """
    Calcule la moyenne pondérée d'un semestre à partir des Note_Finale copiées.
    Retourne (moyenne, credits_valides).
    """
    somme   = 0.0
    total_c = 0.0
    credits = 0

    for mod in modules:
        note   = row.get(f"{mod}_Note_Finale")
        coeff  = coefficients.get(mod, 1.0)
        valide = row.get(f"{mod}_Validé", False)

        if _note_valide(note):
            somme   += float(note) * coeff
            total_c += coeff
            if valide:
                credits += int(coeff * 4)

    moy = round(somme / total_c, 2) if total_c > 0 else None
    return moy, credits


def est_eliminatoire(note) -> bool:
    """Retourne True si la note est éliminatoire (strictement inférieure à 7)."""
    if not _note_valide(note):
        return False
    return float(note) < 7.0


def _decision_semestre(
    row: pd.Series,
    modules: list[str],
    moy: float | None,
    seuil_admis: float = 10.0,
) -> str:
    """
    Règle de validation du semestre :
      - Non Validé si : au moins un module éliminatoire (note < 7)
                        OU nombre de modules non validés > 2
      - Validé si     : aucune des conditions précédentes ET moyenne >= seuil_admis
      - Non Validé    : sinon
    """
    if moy is None:
        return "Non Validé"

    nb_non_valides = 0
    a_eliminatoire = False

    for mod in modules:
        nf     = row.get(f"{mod}_Note_Finale")
        valide = row.get(f"{mod}_Validé", False)
        if _note_valide(nf):
            if est_eliminatoire(nf):
                a_eliminatoire = True
            if not valide:
                nb_non_valides += 1

    if a_eliminatoire or nb_non_valides > 2:
        return "Non Validé"
    if moy >= seuil_admis:
        return "Validé"
    return "Non Validé"


def construire_pv(
    df: pd.DataFrame,
    modules_si: list[str],
    modules_sp: list[str],
    coefficients: dict[str, float] | None = None,
    label_si: str = "S1",
    label_sp: str = "S2",
    seuil_admis:  float = 10.0,
    seuil_rachat: float = 8.0,
) -> pd.DataFrame:
    """
    Calcule les colonnes récapitulatives du PV.
    Si un seul semestre est fourni (l'autre est vide), calcule seulement ce semestre.
    """
    tous_modules = modules_si + modules_sp
    if coefficients is None:
        coefficients = {m: 1.0 for m in tous_modules}

    df = df.copy()
    moy_si_l, cred_si_l = [], []
    moy_sp_l, cred_sp_l = [], []
    moy_ann_l, cred_tot_l = [], []
    decision_l, mention_l, rattrap_l = [], [], []
    dec_si_l, dec_sp_l = [], []
    men_si_l, men_sp_l = [], []

    for _, row in df.iterrows():

        # ── Semestre impair ──
        if modules_si:
            msi, csi = moyenne_semestre(row, modules_si, coefficients)
            dec_si = _decision_semestre(row, modules_si, msi, seuil_admis)
        else:
            msi, csi = None, 0
            dec_si = ""
        moy_si_l.append(msi)
        cred_si_l.append(csi)
        dec_si_l.append(dec_si)
        men_si_l.append(mention(msi))

        # ── Semestre pair ──
        if modules_sp:
            msp, csp = moyenne_semestre(row, modules_sp, coefficients)
            dec_sp = _decision_semestre(row, modules_sp, msp, seuil_admis)
        else:
            msp, csp = None, 0
            dec_sp = ""
        moy_sp_l.append(msp)
        cred_sp_l.append(csp)
        dec_sp_l.append(dec_sp)
        men_sp_l.append(mention(msp))

        # ── Annuel ──
        notes_ann = [x for x in [msi, msp] if x is not None]
        moy_ann   = round(sum(notes_ann) / len(notes_ann), 2) if notes_ann else None
        cred_tot  = csi + csp
        moy_ann_l.append(moy_ann)
        cred_tot_l.append(cred_tot)

        # ── Décision annuelle ──
        # Nouvelles règles :
        # 1. Moy < 10 → Ajourné (Non Validée)
        # 2. Même si Moy >= 10 :
        #    2.1 Au moins un module avec note < 7 → Rachat possible
        #    2.2 Plus de 2 modules non validés (7 <= note < 10) → Rachat possible
        tous_modules_ann = modules_si + modules_sp
        if moy_ann is None:
            dec = "Incomplet"
        elif moy_ann < seuil_admis:
            dec = "Ajourné"
        else:
            # Vérifier les conditions d'échec même avec Moy >= 10
            nb_modules_nv = 0
            a_elim = False
            for mod in tous_modules_ann:
                nf = row.get(f"{mod}_Note_Finale")
                if _note_valide(nf):
                    nf_float = float(nf)
                    if nf_float < 7:
                        a_elim = True
                    elif nf_float < 10:
                        nb_modules_nv += 1
            if a_elim or nb_modules_nv > 2:
                dec = "Rachat possible"
            else:
                dec = "Admis"
        decision_l.append(dec)

        mention_l.append(mention(moy_ann))

        # ── Rattrapages : modules avec Note_Finale < 10 ──
        rattrap = []
        for mod in tous_modules:
            nf = row.get(f"{mod}_Note_Finale")
            if _note_valide(nf) and float(nf) < 10:
                rattrap.append(mod)
        rattrap_l.append(", ".join(rattrap) if rattrap else "")

    if modules_si:
        df[f"Moy_{label_si}"]      = moy_si_l
        df[f"Credits_{label_si}"]  = cred_si_l
        df[f"Dec_{label_si}"]      = dec_si_l
        df[f"Mention_{label_si}"]  = men_si_l
    if modules_sp:
        df[f"Moy_{label_sp}"]      = moy_sp_l
        df[f"Credits_{label_sp}"]  = cred_sp_l
        df[f"Dec_{label_sp}"]      = dec_sp_l
        df[f"Mention_{label_sp}"]  = men_sp_l

    df["Moyenne_Annuelle"] = moy_ann_l
    df["Credits_Totaux"]   = cred_tot_l
    df["Decision"]         = decision_l
    df["Mention"]          = mention_l
    df["Rattrapages"]      = rattrap_l

    counts = df["Decision"].value_counts()
    print("\n" + "="*55)
    print("  RÉSULTATS PV")
    print("="*55)
    for dec, n in counts.items():
        print(f"  {dec:<22} : {n}")
    print("="*55 + "\n")

    return df