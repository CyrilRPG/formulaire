import streamlit as st
import pandas as pd
import unicodedata
import re
import altair as alt

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
    "UPEC": "UPEC",   # demandé
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
    if pd.isna(val):
        return None
    s = str(val).strip()

    m = re.match(r"^\s*(\d+(?:[.,]\d+)?)\s*/\s*(\d+(?:[.,]\d+)?)\s*$", s)
    if m:
        num = float(m.group(1).replace(",", "."))
        den = float(m.group(2).replace(",", "."))
        return (num / den) * 5.0 if den else None

    try:
        return float(s.replace(",", "."))
    except ValueError:
        pass

    m2 = re.match(r"^\s*(\d+(?:[.,]\d+)?)", s)
    if m2:
        return float(m2.group(1).replace(",", "."))
    return None


def read_all_sheets(file) -> pd.DataFrame:
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
    if pseudo_col and pseudo_col in row and pd.notna(row[pseudo_col]):
        f = infer_faculty_from_value(row[pseudo_col])
        if f:
            return f
    if email_col and email_col in row and pd.notna(row[email_col]):
        f = infer_faculty_from_value(row[email_col])
        if f:
            return f

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

        comment_col = None
        for j in (i + 1, i + 2):
            if j < len(columns) and _is_comment_col(norm[j]):
                comment_col = columns[j]
                break
        if not comment_col:
            continue

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
    """Vues <3/5 par catégorie.
       IMPORTANT : on enlève Email & Pseudo pour la lisibilité."""
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

        temp["Fac"] = df.apply(
            lambda r: infer_faculty_for_row(r, pseudo_col, prenom_col, nom_col, email_col),
            axis=1,
        ).map(lambda f: FAC_DISPLAY.get(f, f) if f else "")

        temp["__note_num"] = temp["Note"].map(parse_note)
        temp = temp[temp["__note_num"] < 3.0].drop(columns="__note_num")

        # On enlève Email & Pseudo ici pour gagner de la place
        ordered = [
            c
            for c in ["Prénom", "Nom", "Fac", "Note", "Commentaire"]
            if c in temp.columns
        ]
        sheets[display] = temp[ordered]
    return sheets


# ================== VUES COMMENTAIRES & RECO ================== #

def build_commentaires_view(
    df: pd.DataFrame,
    prenom_col: str | None,
    nom_col: str | None,
    email_col: str | None,
    pseudo_col: str | None,
    pairs: dict[str, tuple[str, str]],
) -> pd.DataFrame:
    """Tous les élèves ayant laissé au moins un commentaire.
       Chaque commentaire sur une nouvelle ligne,
       avec le nom de la section (Coaching, Professeurs, etc.)."""

    # Map colonne commentaire -> nom de vue (Coaching, Professeurs, etc.)
    comment_map = {}  # {comment_col: "Coaching"}
    for display, (_, comm_col) in pairs.items():
        if comm_col in df.columns:
            comment_map[comm_col] = display

    rows = []
    for _, r in df.iterrows():
        comments = []
        for comm_col, display_name in comment_map.items():
            val = r.get(comm_col)
            if isinstance(val, str) and val.strip():
                comments.append(f"{display_name} :\n{val.strip()}")

        if not comments:
            continue

        fac = infer_faculty_for_row(r, pseudo_col, prenom_col, nom_col, email_col)
        rows.append(
            {
                "Prénom": r.get(prenom_col, ""),
                "Nom": r.get(nom_col, ""),
                "Fac": FAC_DISPLAY.get(fac, fac) if fac else "",
                "Commentaires": "\n\n".join(comments),
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
            columns=["Prénom", "Nom", "Fac", "Recommandation"]
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
                    "Fac": FAC_DISPLAY.get(fac, fac) if fac else "",
                    "Recommandation": rec.strip(),
                }
            )
    return pd.DataFrame(rows)


# ================== VUE "TOUS LES ÉLÈVES" ================== #

def build_tous_les_eleves(
    df: pd.DataFrame,
    prenom_col: str | None,
    nom_col: str | None,
    email_col: str | None,
    pseudo_col: str | None,
    pairs: dict[str, tuple[str, str]],
) -> pd.DataFrame:
    note_cols = [note for (note, _) in pairs.values()]
    rows = []
    for _, r in df.iterrows():
        notes = []
        for note_col in note_cols:
            val = parse_note(r.get(note_col))
            if val is not None:
                notes.append(val)
        if notes:
            avg = sum(notes) / len(notes)
        else:
            avg = None

        fac = infer_faculty_for_row(r, pseudo_col, prenom_col, nom_col, email_col)
        rows.append(
            {
                "Prénom": r.get(prenom_col, ""),
                "Nom": r.get(nom_col, ""),
                "Fac": FAC_DISPLAY.get(fac, fac) if fac else "",
                "Pseudo": r.get(pseudo_col, ""),
                "Email": r.get(email_col, ""),
                "Moyenne globale /5": round(avg, 2) if avg is not None else None,
            }
        )
    out = pd.DataFrame(rows)
    out = out.sort_values(
        "Moyenne globale /5", ascending=True, na_position="last"
    ).reset_index(drop=True)
    return out


# ================== STREAMLIT UI ================== #

st.set_page_config(page_title="Feedback PASS", layout="wide")

# Petit style pour rendre le tout plus agréable
st.markdown("""
<style>
.block {
    background-color: #f8f9fc;
    padding: 20px;
    border-radius: 14px;
    margin-bottom: 25px;
    box-shadow: 0 4px 12px rgba(0,0,0,0.05);
}
.dataframe { border-radius: 12px; overflow: hidden; }
</style>
""", unsafe_allow_html=True)

st.markdown(
    "<h1 style='text-align:center;'>📊 Dashboard de satisfaction – Diploma Santé</h1>",
    unsafe_allow_html=True,
)
st.write(
    "Dépose l’export Excel du formulaire : la plateforme te donne directement "
    "les moyennes par fac, les élèves en difficulté, tous les commentaires, les recommandations, "
    "un classement des élèves par moyenne globale, et des graphiques de synthèse."
)

uploaded = st.file_uploader("Fichier Excel exporté (.xlsx ou .xls)", type=["xlsx", "xls"])

if not uploaded:
    st.info("▶ Uploade un fichier pour démarrer l’analyse.")
    st.stop()

try:
    df = read_all_sheets(uploaded)
except Exception as e:
    st.error(f"Erreur lors de la lecture du fichier : {e}")
    st.stop()

if df.empty:
    st.error("Le fichier ne contient aucune donnée exploitable.")
    st.stop()

# Détection colonnes
prenom_col, nom_col, email_col = find_identity_columns(df)
pseudo_col = find_pseudo_column(df)
pairs = build_pairs(df)

if not pairs:
    st.error(
        "Impossible de détecter les colonnes de notes/commentaires.\n\n"
        "Vérifie que les questions sont bien du type "
        "'Sur une échelle de 0 à 5...' et que les commentaires sont dans des colonnes 'Commentaire'."
    )
    st.stop()

df_avg = compute_averages_by_fac(df, pairs, pseudo_col, prenom_col, nom_col, email_col)
standard_views = build_views(df, prenom_col, nom_col, email_col, pseudo_col, pairs)
commentaires_df = build_commentaires_view(df, prenom_col, nom_col, email_col, pseudo_col, pairs)
reco_df = build_recommandations_view(df, prenom_col, nom_col, email_col, pseudo_col)
tous_eleves_df = build_tous_les_eleves(df, prenom_col, nom_col, email_col, pseudo_col, pairs)

# ========== HEADER STATS ========== #

col1, col2, col3 = st.columns(3)
with col1:
    st.metric("Nombre de réponses", len(df))
with col2:
    # moyenne globale (toutes notes toutes facs)
    all_notes = []
    for note_col, _ in pairs.values():
        all_notes.extend(df[note_col].map(parse_note).dropna().tolist())
    global_mean = round(sum(all_notes) / len(all_notes), 2) if all_notes else None
    st.metric("Moyenne globale", f"{global_mean}/5" if global_mean is not None else "N/A")
with col3:
    st.metric("Facs détectées", ", ".join(sorted({f for f in tous_eleves_df["Fac"].unique() if f})) or "—")

# ========== TABS ========== #

tab_moyennes, tab_vues, tab_comments, tab_reco, tab_eleves, tab_graphs = st.tabs(
    ["📈 Moyennes par fac", "⚠️ Vues < 3/5", "💬 Commentaires", "📝 Recommandations", "👥 Tous les élèves", "📊 Graphiques"]
)

with tab_moyennes:
    st.markdown("<div class='block'>", unsafe_allow_html=True)
    st.subheader("Moyennes par fac et par catégorie")
    st.dataframe(df_avg, use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)

with tab_vues:
    st.markdown("<div class='block'>", unsafe_allow_html=True)
    st.subheader("Élèves avec note < 3/5 par catégorie")
    vue_name = st.selectbox(
        "Choisis une catégorie",
        ["Coaching", "Fiches de cours", "Professeurs", "Plateforme", "Organisation générale"],
    )
    df_vue = standard_views.get(vue_name)
    if df_vue is None or df_vue.empty:
        st.info("Aucun élève avec note < 3/5 pour cette catégorie.")
    else:
        fac_filter = st.multiselect(
            "Filtrer par fac (optionnel)",
            sorted([f for f in df_vue["Fac"].unique() if f]),
        )
        df_affiche = df_vue.copy()
        if fac_filter:
            df_affiche = df_affiche[df_affiche["Fac"].isin(fac_filter)]
        st.dataframe(df_affiche, use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)

with tab_comments:
    st.markdown("<div class='block'>", unsafe_allow_html=True)
    st.subheader("Tous les élèves ayant laissé au moins un commentaire")
    if commentaires_df.empty:
        st.info("Aucun commentaire détecté.")
    else:
        fac_filter = st.multiselect(
            "Filtrer par fac (optionnel)", sorted([f for f in commentaires_df["Fac"].unique() if f]), key="fac_comments"
        )
        df_affiche = commentaires_df.copy()
        if fac_filter:
            df_affiche = df_affiche[df_affiche["Fac"].isin(fac_filter)]
        st.dataframe(df_affiche, use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)

with tab_reco:
    st.markdown("<div class='block'>", unsafe_allow_html=True)
    st.subheader("Recommandations / besoins avant le concours")
    if reco_df.empty:
        st.info("Aucune recommandation trouvée dans le champ dédié.")
    else:
        fac_filter = st.multiselect(
            "Filtrer par fac (optionnel)", sorted([f for f in reco_df["Fac"].unique() if f]), key="fac_reco"
        )
        df_affiche = reco_df.copy()
        if fac_filter:
            df_affiche = df_affiche[df_affiche["Fac"].isin(fac_filter)]
        st.dataframe(df_affiche, use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)

with tab_eleves:
    st.markdown("<div class='block'>", unsafe_allow_html=True)
    st.subheader("Tous les élèves – classement par moyenne globale (croissant)")
    if tous_eleves_df.empty:
        st.info("Aucune donnée pour calculer les moyennes.")
    else:
        fac_filter = st.multiselect(
            "Filtrer par fac (optionnel)", sorted([f for f in tous_eleves_df["Fac"].unique() if f]), key="fac_eleves"
        )
        df_affiche = tous_eleves_df.copy()
        if fac_filter:
            df_affiche = df_affiche[df_affiche["Fac"].isin(fac_filter)]
        st.dataframe(
            df_affiche[["Prénom", "Nom", "Fac", "Moyenne globale /5", "Pseudo", "Email"]],
            use_container_width=True,
        )
    st.markdown("</div>", unsafe_allow_html=True)

with tab_graphs:
    st.markdown("<div class='block'>", unsafe_allow_html=True)
    st.subheader("Histogramme des moyennes des élèves (A2)")

    moy_series = tous_eleves_df["Moyenne globale /5"].dropna()
    if moy_series.empty:
        st.info("Pas assez de données pour tracer l'histogramme des moyennes.")
    else:
        hist_df = pd.DataFrame({"Moyenne": moy_series})
        chart_hist = (
            alt.Chart(hist_df)
            .mark_bar()
            .encode(
                x=alt.X("Moyenne:Q", bin=alt.Bin(maxbins=15), title="Moyenne sur 5"),
                y=alt.Y("count():Q", title="Nombre d'élèves"),
                tooltip=["count():Q"]
            )
            .properties(height=300)
        )
        st.altair_chart(chart_hist, use_container_width=True)

    st.markdown("---")
    st.subheader("Comparaison des moyennes par catégorie et par fac (D8)")

    long_df = df_avg.melt(id_vars="Catégorie", var_name="Fac", value_name="Moyenne")
    long_df = long_df.dropna(subset=["Moyenne"])
    if long_df.empty:
        st.info("Pas assez de données pour tracer le comparatif par fac.")
    else:
        chart_bar = (
            alt.Chart(long_df)
            .mark_bar()
            .encode(
                x=alt.X("Catégorie:N", sort=None, title="Catégorie"),
                y=alt.Y("Moyenne:Q", title="Moyenne /5"),
                color=alt.Color("Fac:N", title="Fac"),
                tooltip=["Catégorie", "Fac", "Moyenne"]
            )
            .properties(height=350)
        )
        st.altair_chart(chart_bar, use_container_width=True)

    st.markdown("</div>", unsafe_allow_html=True)
