import streamlit as st
import pandas as pd
import unicodedata
import re
from io import BytesIO

# ================== CONFIG ================== #

TARGET_VIEWS = [
    ("coaching", "Coaching"),
    ("fichesdecours", "Fiches de cours"),
    ("fiches cours", "Fiches de cours"),
    ("professeurs", "Professeurs"),
    ("plateforme", "Plateforme"),
    ("organisationgenerale", "Organisation générale"),
    ("organisation generale", "Organisation générale"),
]

FAC_ORDER = ["UPC", "UPEC", "UPS", "UVSQ", "SU", "USPN"]
FAC_DISPLAY = {
    "UPC": "UPC",
    "UPEC": "UPEC L1",
    "UPS": "UPS",
    "UVSQ": "UVSQ",
    "SU": "SU",
    "USPN": "USPN",
}

RECO_COL_EXACT = (
    "Si vous avez des besoins, des demandes ou des améliorations à proposer avant le concours, écrivez-les ici !"
)

# ================== UTILS ================== #


def normalize(s: str) -> str:
    if s is None:
        return ""
    s = str(s)
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = s.lower().strip()
    s = re.sub(r"\s+", " ", s)
    return s


def parse_note(val):
    """Renvoie une note sur 5 en float ou None."""
    if pd.isna(val):
        return None
    s = str(val).strip()

    # 1) Forme "2/5"
    m = re.match(r"^\s*(\d+(?:[.,]\d+)?)\s*/\s*(\d+(?:[.,]\d+)?)\s*$", s)
    if m:
        num = float(m.group(1).replace(",", "."))
        den = float(m.group(2).replace(",", "."))
        return (num / den) * 5.0 if den else None

    # 2) Nombre simple "2" ou "2,5"
    try:
        return float(s.replace(",", "."))
    except ValueError:
        pass

    # 3) Nombre en début de chaîne "4 - Plutôt satisfait"
    m2 = re.match(r"^\s*(\d+(?:[.,]\d+)?)", s)
    if m2:
        return float(m2.group(1).replace(",", "."))

    return None


def read_all_sheets(file) -> pd.DataFrame:
    """Concatène toutes les feuilles d'un Excel uploadé."""
    try:
        xls = pd.ExcelFile(file)
        frames = []
        for sheet in xls.sheet_names:
            df_sheet = pd.read_excel(xls, sheet_name=sheet)
            if not df_sheet.empty:
                df_sheet["_source_sheet"] = sheet
                frames.append(df_sheet)
        return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    except Exception as e:
        raise RuntimeError(f"Erreur lecture Excel: {e}")


def find_identity_columns(df: pd.DataFrame):
    cols = {normalize(c): c for c in df.columns}
    prenom = next(
        (cols[k] for k in cols if "prenom" in k or "prénom" in k or "first" in k),
        None,
    )
    nom = next((cols[k] for k in cols if k.startswith("nom") or "last" in k), None)
    email = next(
        (cols[k] for k in cols if "email" in k or "mail" in k),
        None,
    )
    return prenom, nom, email


def find_pseudo_column(df: pd.DataFrame):
    for col in df.columns:
        n = normalize(col)
        if any(k in n for k in ["pseudo", "identifiant", "username", "login", "user"]):
            return col
    return None


def infer_faculty_from_value(val: str):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    s = normalize(val).upper()
    for fac in FAC_ORDER:
        if fac in s:
            return fac
    return None


def infer_faculty_for_row(
    row: pd.Series,
    pseudo_col: str | None,
    prenom_col: str | None,
    nom_col: str | None,
    email_col: str | None,
):
    # 1) Pseudo
    if pseudo_col and pseudo_col in row and pd.notna(row[pseudo_col]):
        f = infer_faculty_from_value(row[pseudo_col])
        if f:
            return f
    # 2) Email
    if email_col and email_col in row and pd.notna(row[email_col]):
        f = infer_faculty_from_value(row[email_col])
        if f:
            return f
    # 3) Prénom + Nom
    parts = []
    if prenom_col and prenom_col in row and pd.notna(row[prenom_col]):
        parts.append(str(row[prenom_col]))
    if nom_col and nom_col in row and pd.notna(row[nom_col]):
        parts.append(str(row[nom_col]))
    if parts:
        return infer_faculty_from_value(" ".join(parts))
    return None


def _is_comment_col(n: str) -> bool:
    return any(x in n for x in ["comment", "commentaire", "remarque", "avis"])


# ================== LOGIQUE "NOTES & VUES" ================== #


def build_pairs(df: pd.DataFrame):
    """Détecte les couples (colonne note/échelle, colonne commentaire)."""
    columns = list(df.columns)
    norm = [normalize(c) for c in columns]

    target_keys = {k for k, _ in TARGET_VIEWS}
    display_map = dict(TARGET_VIEWS)
    pairs: dict[str, tuple[str, str]] = {}

    for i, ncol in enumerate(norm):
        is_note = "note" in ncol
        is_scale = "echelle de 0 a 5" in ncol or ncol.startswith("sur une echelle")

        if not (is_note or is_scale):
            continue

        # Chercher la colonne de commentaire dans les 2 suivantes
        comment_col = None
        for j in (i + 1, i + 2):
            if j < len(columns) and _is_comment_col(norm[j]):
                comment_col = columns[j]
                break
        if not comment_col:
            continue

        # Catégorie
        if is_note:
            base = re.sub(r"\bnote\b|:|-", " ", ncol)
        else:
            base = re.sub(r"^sur une echelle( de)? 0 a 5", " ", ncol)

        cat_key = normalize(base)
        cat_key_simple = re.sub(r"[^a-z0-9 ]", "", cat_key)
        cat_key_simple = re.sub(r"\s+", "", cat_key_simple)

        match_key = None
        for tk in target_keys:
            if tk in cat_key_simple or cat_key_simple in tk:
                match_key = tk
                break
        if not match_key:
            continue

        display = display_map[match_key]
        pairs[display] = (columns[i], comment_col)

    return pairs


def compute_averages_by_fac(
    df: pd.DataFrame,
    pairs: dict[str, tuple[str, str]],
    pseudo_col: str | None,
    prenom_col: str | None,
    nom_col: str | None,
    email_col: str | None,
) -> pd.DataFrame:
    rows = []
    for view, (note_col, _) in pairs.items():
        notes = df[note_col].map(parse_note)
        facs = df.apply(
            lambda r: infer_faculty_for_row(r, pseudo_col, prenom_col, nom_col, email_col),
            axis=1,
        )
        tmp = pd.DataFrame({"fac": facs, "note": notes}).dropna(subset=["fac", "note"])
        if tmp.empty:
            rows.append({"Catégorie": view, **{FAC_DISPLAY[f]: None for f in FAC_ORDER}})
            continue

        mean_by_fac = tmp.groupby("fac")["note"].mean().to_dict()
        row = {"Catégorie": view}
        for f in FAC_ORDER:
            val = mean_by_fac.get(f)
            row[FAC_DISPLAY[f]] = round(float(val), 2) if val is not None else None
        rows.append(row)

    df_avg = pd.DataFrame(rows)
    ordered_cols = ["Catégorie"] + [FAC_DISPLAY[f] for f in FAC_ORDER]
    return df_avg.reindex(columns=ordered_cols).sort_values("Catégorie").reset_index(drop=True)


def build_views(
    df: pd.DataFrame,
    prenom_col: str | None,
    nom_col: str | None,
    email_col: str | None,
    pseudo_col: str | None,
    pairs: dict[str, tuple[str, str]],
):
    sheets: dict[str, pd.DataFrame] = {}
    for display, (note_col, comm_col) in pairs.items():
        cols = [
            c
            for c in [prenom_col, nom_col, email_col, pseudo_col, note_col, comm_col]
            if c and c in df.columns
        ]
        if not cols:
            continue

        temp = df[cols].copy()
        rename_map: dict[str, str] = {}
        if prenom_col in temp.columns:
            rename_map[prenom_col] = "Prénom"
        if nom_col in temp.columns:
            rename_map[nom_col] = "Nom"
        if email_col in temp.columns:
            rename_map[email_col] = "Email"
        if pseudo_col in temp.columns:
            rename_map[pseudo_col] = "Pseudo"
        rename_map[note_col] = "Note"
        rename_map[comm_col] = "Commentaire"
        temp.rename(columns=rename_map, inplace=True)

        # Ajout fac
        temp["Fac"] = df.apply(
            lambda r: infer_faculty_for_row(r, pseudo_col, prenom_col, nom_col, email_col),
            axis=1,
        ).map(lambda f: FAC_DISPLAY.get(f, f) if f else "")

        temp["__note_num"] = temp["Note"].map(parse_note)
        temp = temp[temp["__note_num"] < 3.0].drop(columns="__note_num")

        ordered = [
            c
            for c in ["Prénom", "Nom", "Email", "Pseudo", "Fac", "Note", "Commentaire"]
            if c in temp.columns
        ]
        sheets[display] = temp[ordered]
    return sheets


# ================== NOUVELLES VUES ================== #


def build_commentaires_view(
    df: pd.DataFrame,
    prenom_col: str | None,
    nom_col: str | None,
    email_col: str | None,
    pseudo_col: str | None,
) -> pd.DataFrame:
    comment_cols = [c for c in df.columns if _is_comment_col(normalize(c))]
    rows = []
    for _, r in df.iterrows():
        comments = []
        for col in comment_cols:
            val = r.get(col)
            if isinstance(val, str) and val.strip():
                comments.append(f"{col}: {val.strip()}")
        if not comments:
            continue

        fac = infer_faculty_for_row(r, pseudo_col, prenom_col, nom_col, email_col)
        rows.append(
            {
                "Prénom": r.get(prenom_col, ""),
                "Nom": r.get(nom_col, ""),
                "Email": r.get(email_col, ""),
                "Pseudo": r.get(pseudo_col, ""),
                "Fac": FAC_DISPLAY.get(fac, fac) if fac else "",
                "Commentaires": "\n".join(comments),
            }
        )
    return pd.DataFrame(rows)


def build_recommandations_view(
    df: pd.DataFrame,
    prenom_col: str | None,
    nom_col: str | None,
    email_col: str | None,
    pseudo_col: str | None,
) -> pd.DataFrame:
    if RECO_COL_EXACT not in df.columns:
        return pd.DataFrame(
            columns=["Prénom", "Nom", "Email", "Pseudo", "Fac", "Recommandation"]
        )

    rows = []
    for _, r in df.iterrows():
        rec = r.get(RECO_COL_EXACT)
        if isinstance(rec, str) and rec.strip():
            fac = infer_faculty_for_row(r, pseudo_col, prenom_col, nom_col, email_col)
            rows.append(
                {
                    "Prénom": r.get(prenom_col, ""),
                    "Nom": r.get(nom_col, ""),
                    "Email": r.get(email_col, ""),
                    "Pseudo": r.get(pseudo_col, ""),
                    "Fac": FAC_DISPLAY.get(fac, fac) if fac else "",
                    "Recommandation": rec.strip(),
                }
            )
    return pd.DataFrame(rows)


def build_excel_bytes(
    df_avg: pd.DataFrame,
    standard_views: dict[str, pd.DataFrame],
    commentaires_df: pd.DataFrame,
    reco_df: pd.DataFrame,
) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        # Moyennes
        df_avg.to_excel(writer, sheet_name="Moyennes", index=False)
        # Vues <3/5
        for view in ["Coaching", "Fiches de cours", "Professeurs", "Plateforme", "Organisation générale"]:
            df_view = standard_views.get(view, pd.DataFrame())
            df_view.to_excel(writer, sheet_name=view[:31], index=False)
        # Commentaires & Recommandations
        commentaires_df.to_excel(writer, sheet_name="Commentaires", index=False)
        reco_df.to_excel(writer, sheet_name="Recommandations", index=False)

    output.seek(0)
    return output.getvalue()


# ================== STREAMLIT APP ================== #

st.set_page_config(page_title="Feedback PASS", layout="centered")

st.title("📊 Générateur de vues de feedback")
st.write(
    "Uploade l’export brut (Excel) du formulaire, et je te génère un fichier "
    "**vues_feedback.xlsx** avec : Moyennes, vues <3/5, Commentaires & Recommandations."
)

uploaded = st.file_uploader("Fichier Excel exporté", type=["xlsx", "xls"])

if not uploaded:
    st.info("👉 Choisis un fichier .xlsx ou .xls pour commencer.")
    st.stop()

try:
    df = read_all_sheets(uploaded)
except Exception as e:
    st.error(f"Erreur lors de la lecture du fichier : {e}")
    st.stop()

if df.empty:
    st.error("Le fichier ne contient aucune donnée exploitable.")
    st.stop()

st.success(f"✅ Fichier chargé ({len(df)} lignes, {len(df.columns)} colonnes).")

# Détection colonnes identité / pseudo
prenom_col, nom_col, email_col = find_identity_columns(df)
pseudo_col = find_pseudo_column(df)

pairs = build_pairs(df)
if not pairs:
    st.error(
        "Impossible de détecter les colonnes de notes/commentaires.\n\n"
        "Vérifie que les questions sont bien du type "
        "'Sur une échelle de 0 à 5, comment évaluez-vous...'"
        " et que les commentaires sont dans des colonnes 'Commentaire'."
    )
    st.stop()

df_avg = compute_averages_by_fac(df, pairs, pseudo_col, prenom_col, nom_col, email_col)
standard_views = build_views(df, prenom_col, nom_col, email_col, pseudo_col, pairs)
commentaires_df = build_commentaires_view(df, prenom_col, nom_col, email_col, pseudo_col)
reco_df = build_recommandations_view(df, prenom_col, nom_col, email_col, pseudo_col)

st.subheader("Aperçu rapide des moyennes")
st.dataframe(df_avg)

excel_bytes = build_excel_bytes(df_avg, standard_views, commentaires_df, reco_df)

st.download_button(
    label="📥 Télécharger vues_feedback.xlsx",
    data=excel_bytes,
    file_name="vues_feedback.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.caption(
    "Le fichier contient : Moyennes, Coaching, Fiches de cours, Professeurs, "
    "Plateforme, Organisation générale, Commentaires, Recommandations."
)
