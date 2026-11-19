"""Calculs sur les données (pandas uniquement, sans interface).

Chaque fonction reçoit des DataFrames et renvoie un DataFrame,
ce qui permet de les tester sans ouvrir l'interface.
"""

import pandas as pd

from . import config

COLS_STATS = ["المخلف", "المسجل", "الرائج", "المحكوم", "الباقي"]


# ---------------------------------------------------------------------------
# Lecture et classification
# ---------------------------------------------------------------------------
def lire_tableau(chemin):
    """Lit un fichier Excel ou CSV (séparateur ;) exporté par le tribunal."""
    if str(chemin).lower().endswith(".csv"):
        return pd.read_csv(chemin, encoding="utf-8-sig", delimiter=";")
    return pd.read_excel(chemin)


def extraire_partie(numero, position=1):
    """Renvoie la partie `position` d'un numéro de dossier « n/رمز/سنة », ou None."""
    morceaux = str(numero).split("/")
    if len(morceaux) > position and "/" in str(numero):
        return morceaux[position].strip()
    return None


MOIS_FR = {"janv": "01", "févr": "02", "fevr": "02", "mars": "03", "avr": "04", "mai": "05", "juin": "06",
           "juil": "07", "août": "08", "aout": "08", "sept": "09", "oct": "10", "nov": "11", "déc": "12", "dec": "12"}


def convertir_dates(serie):
    """Convertit une colonne de dates aux formats rencontrés dans les exports du tribunal :
    vraie date Excel, « 12/02/2025 », ou « févr 12 2025 10:30AM »."""
    if pd.api.types.is_datetime64_any_dtype(serie):
        return serie
    dates = pd.to_datetime(serie, format="%d/%m/%Y", errors="coerce")
    reste = dates.isna() & serie.notna()
    if reste.any():
        texte = serie[reste].astype(str).str.strip().str.lower()
        for mois, numero in MOIS_FR.items():
            texte = texte.str.replace(rf"^{mois}\.?\s", f"{numero} ", regex=True)
        texte = texte.str.replace(r"\s+", " ", regex=True).str.upper()
        dates[reste] = pd.to_datetime(texte, format="%m %d %Y %I:%M%p", errors="coerce").dt.normalize()
    reste = dates.isna() & serie.notna()
    if reste.any():  # dernier recours : date ISO (2025-02-12)
        dates[reste] = pd.to_datetime(serie[reste], format="%Y-%m-%d", errors="coerce")
    return dates


def chercher_shi3ba(code, table):
    """Renvoie la شعبة d'un رمز selon une table {شعبة: [codes]}."""
    for shi3ba, codes in table.items():
        if code in codes:
            return shi3ba
    return config.INCONNU_SHI3BA


def chercher_naw3iya(code, table):
    """Renvoie la نوعية d'un رمز selon une table {"c1,c2": نوعية}."""
    code_str = str(code)
    for cles, naw3iya in table.items():
        if code_str in cles.split(","):
            return naw3iya
    return config.INCONNU_NAW3IYA


def ajouter_totaux(pivot):
    """Ajoute la colonne et la ligne « المجموع » à un tableau croisé, puis remet l'index en colonne."""
    pivot = pivot.copy()
    pivot["المجموع"] = pivot.sum(axis=1)
    ligne_total = pivot.sum().to_frame().T
    ligne_total.index = ["المجموع"]
    pivot = pd.concat([pivot, ligne_total]).astype(int)
    return pivot.reset_index()


# ---------------------------------------------------------------------------
# الزجري : إحصائيات عامة
# ---------------------------------------------------------------------------
def preparer_fichier_stats(df):
    """Nettoie un fichier المخلف / المسجل / المحكوم et ajoute la colonne الرمز."""
    df = df.dropna(axis=1, how="all").copy()
    df[config.COL_NUMERO] = df[config.COL_NUMERO].astype(str).str.strip()
    df["الرمز"] = df[config.COL_NUMERO].apply(extraire_partie)
    df = df.dropna(subset=["الرمز"]).drop_duplicates(subset=[config.COL_NUMERO])
    df["الرمز"] = df["الرمز"].astype(str).str.strip()
    return df


def stats_generales(df_mokhalaf, df_mossajal, df_mahkoum):
    """Tableau المخلف / المسجل / الرائج / المحكوم / الباقي par رمز, avec ligne المجموع."""
    freq_mokhalaf = df_mokhalaf.groupby("الرمز").agg(
        المخلف=("الرمز", "count"),
        نوع_القضية=("نوع القضية", "first"),
    ).reset_index()

    freq_mossajal = df_mossajal.drop_duplicates(subset=[config.COL_NUMERO])["الرمز"].value_counts().reset_index()
    freq_mossajal.columns = ["الرمز", "المسجل"]

    df = pd.merge(freq_mokhalaf, freq_mossajal, on="الرمز", how="outer")
    df["المخلف"] = df["المخلف"].fillna(0).astype(int)
    df["المسجل"] = df["المسجل"].fillna(0).astype(int)
    df["الرائج"] = df["المخلف"] + df["المسجل"]

    freq_mahkoum = df_mahkoum.drop_duplicates(subset=[config.COL_NUMERO])["الرمز"].value_counts().reset_index()
    freq_mahkoum.columns = ["الرمز", "المحكوم"]

    df = pd.merge(df, freq_mahkoum, on="الرمز", how="outer")
    df["المحكوم"] = df["المحكوم"].fillna(0).astype(int)
    df["الباقي"] = df["الرائج"] - df["المحكوم"]
    df = df.rename(columns={"نوع_القضية": "نوع القضية"})

    colonnes = ["نوع القضية", "الرمز"] + COLS_STATS
    for col in colonnes:
        if col not in df.columns:
            df[col] = 0 if col in COLS_STATS else ""
    df = df.reindex(columns=colonnes)
    df = df[df["الرمز"] != "المجموع"]
    for col in COLS_STATS:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)

    total = {"نوع القضية": "", "الرمز": "المجموع"}
    total.update({col: df[col].sum() for col in COLS_STATS})
    return pd.concat([df, pd.DataFrame([total])], ignore_index=True)


def stats_par_hay2a(df_mokhalaf, df_mossajal, df_mahkoum):
    """Même tableau que stats_generales, mais regroupé par الهيئة."""
    sans_hay2a = "غير محددة"
    m = df_mokhalaf.assign(الهيئة=df_mokhalaf["الهيئة"].fillna(sans_hay2a))
    s = df_mossajal.assign(الهيئة=df_mossajal["الهيئة"].fillna(sans_hay2a))
    h = df_mahkoum.assign(الهيئة=df_mahkoum["الهيئة"].fillna(sans_hay2a))

    freq_m = m.groupby("الهيئة").size().reset_index(name="المخلف")
    freq_s = s.drop_duplicates(subset=[config.COL_NUMERO]).groupby("الهيئة").size().reset_index(name="المسجل")
    freq_h = h.drop_duplicates(subset=[config.COL_NUMERO]).groupby("الهيئة").size().reset_index(name="المحكوم")

    df = pd.merge(freq_m, freq_s, on="الهيئة", how="outer").merge(freq_h, on="الهيئة", how="outer")
    df = df.fillna(0)
    df[["المخلف", "المسجل", "المحكوم"]] = df[["المخلف", "المسجل", "المحكوم"]].astype(int)
    df["الرائج"] = df["المخلف"] + df["المسجل"]
    df["الباقي"] = df["الرائج"] - df["المحكوم"]

    total = {"الهيئة": "المجموع"}
    total.update({col: df[col].sum() for col in ["المخلف", "المسجل", "الرائج", "المحكوم", "الباقي"]})
    return pd.concat([df, pd.DataFrame([total])], ignore_index=True)


# ---------------------------------------------------------------------------
# الزجري : انتاجيات القضاة et الملفات المزمنة
# ---------------------------------------------------------------------------
def preparer_ahkam(df):
    """Nettoie le سجل الأحكام et ajoute الرمز (entier) et الشعبة."""
    df = df.dropna(axis=1, how="all").copy()
    df[config.COL_NUMERO] = df[config.COL_NUMERO].astype(str).str.strip()
    df["الرمز"] = df[config.COL_NUMERO].apply(extraire_partie)
    df = df.dropna(subset=["الرمز"])
    df["الرمز"] = df["الرمز"].astype(int)
    df["الشعبة"] = df["الرمز"].apply(lambda c: chercher_shi3ba(c, config.CODES_SHI3AB_ZAJRI))
    return df


def pivot_juges(df, col_juge):
    """Nombre de dossiers par شعبة et par juge, avec les totaux. Première colonne : « index »."""
    comptes = df.groupby(["الشعبة", col_juge]).size().reset_index(name="عدد الملفات")
    pivot = pd.pivot_table(comptes, index="الشعبة", columns=col_juge, values="عدد الملفات", fill_value=0)
    return ajouter_totaux(pivot)


def preparer_chronic(df):
    """Nettoie le fichier des الملفات المزمنة et ajoute السنة et الرمز (entiers)."""
    df = df.dropna(subset=[config.COL_NUMERO]).copy()
    df[config.COL_NUMERO] = df[config.COL_NUMERO].astype(str).str.strip()
    df["السنة"] = df[config.COL_NUMERO].apply(lambda x: x.split("/")[-1] if "/" in x else None)
    df["الرمز"] = df[config.COL_NUMERO].apply(lambda x: x.split("/")[1] if "/" in x else None)
    df = df.dropna(subset=["السنة", "الرمز"])
    df["السنة"] = df["السنة"].astype(int)
    df["الرمز"] = df["الرمز"].astype(int)
    return df


def pivot_chronic(df_chronic):
    """Nombre de dossiers restants par شعبة (lignes) et par السنة (colonnes)."""
    df = df_chronic.copy()
    df["الشعبة"] = df["الرمز"].apply(lambda c: chercher_shi3ba(c, config.CODES_SHI3AB_ZAJRI))
    groupes = df.groupby(["الشعبة", "السنة"]).size().reset_index(name="عدد الملفات")
    pivot = pd.pivot_table(groupes, index="الشعبة", columns="السنة", values="عدد الملفات", fill_value=0)
    pivot = ajouter_totaux(pivot)
    return pivot.rename(columns={pivot.columns[0]: "الشعبة"})


def chronic_par_annee(df_chronic):
    """Pour le graphique : nombre de dossiers par السنة (lignes) et par شعبة (colonnes)."""
    df = df_chronic.copy()
    df["الشعبة"] = df["الرمز"].apply(lambda c: chercher_shi3ba(c, config.CODES_SHI3AB_ZAJRI))
    groupes = df.groupby(["السنة", "الشعبة"]).size().reset_index(name="عدد الملفات")
    return pd.pivot_table(groupes, index="السنة", columns="الشعبة", values="عدد الملفات", fill_value=0)


# ---------------------------------------------------------------------------
# المدني : انتاجيات القضاة (fichiers fusionnés المدني / الأسرة)
# ---------------------------------------------------------------------------
def preparer_fichier_fusionne(df, table_shi3ab):
    """Ajoute الرمز (entier) et الشعبة à un fichier fusionné المدني ou الأسرة."""
    df = df.dropna(axis=1, how="all").copy()
    df["الرمز"] = df[config.COL_NUMERO_COMPLET].apply(extraire_partie)
    df = df.dropna(subset=["الرمز"])
    df["الرمز"] = df["الرمز"].astype(int)
    df["الشعبة"] = df["الرمز"].apply(lambda c: chercher_shi3ba(c, table_shi3ab))
    return df


def colonnes_presentes(df, colonnes):
    return all(col in df.columns for col in colonnes)


# ---------------------------------------------------------------------------
# المدني : إحصائيات التنفيذ (المدني et الأسرة)
# ---------------------------------------------------------------------------
def stats_tanfid(fichiers, date_debut, date_fin, table_naw3iya):
    """Tableau المخلف / المسجل / المجموع / المنفذ / الباقي par نوعية القضايا,
    avec le reste par année d'enregistrement et la ligne المجموع."""
    df = pd.concat(fichiers, ignore_index=True)
    df["الرمز"] = df[config.COL_NUMERO_COMPLET].apply(extraire_partie)
    df = df.dropna(subset=["الرمز"])
    for col in ["تاريخ التسجيل بالسجل", "تاريخ الحالة"]:
        df[col] = convertir_dates(df[col])
    df["السنة"] = df["تاريخ التسجيل بالسجل"].dt.year
    df["نوعية القضايا"] = df["الرمز"].apply(lambda c: chercher_naw3iya(c, table_naw3iya))

    etat = df["حالة الملف"]
    execute = etat.notna() & (etat != "غير منفذ")
    date_etat_dans_periode = (df["تاريخ الحالة"] >= date_debut) & (df["تاريخ الحالة"] <= date_fin)
    enregistre_dans_periode = (df["تاريخ التسجيل بالسجل"] >= date_debut) & (df["تاريخ التسجيل بالسجل"] <= date_fin)
    mokhalaf = (
        etat.isna()
        | (etat == "غير منفذ")
        | (df["تاريخ التسجيل بالسجل"] < date_debut)
        | ((etat == "منفذ") & date_etat_dans_periode)
    )

    lignes = []
    for naw3iya in dict.fromkeys(table_naw3iya.values()):  # ordre de la table, sans doublon
        sel = df["نوعية القضايا"] == naw3iya
        if not sel.any():
            lignes.append({"نوعية القضايا": naw3iya, "المخلف": 0, "المسجل": 0, "المجموع": 0, "المنفذ": 0, "الباقي": 0})
            continue
        n_mokhalaf = int((sel & mokhalaf).sum())
        n_mossajal = int((sel & enregistre_dans_periode).sum())
        n_monafath = int((sel & execute & date_etat_dans_periode).sum())
        ligne = {
            "نوعية القضايا": naw3iya,
            "المخلف": n_mokhalaf,
            "المسجل": n_mossajal,
            "المجموع": n_mokhalaf + n_mossajal,
            "المنفذ": n_monafath,
            "الباقي": n_mokhalaf + n_mossajal - n_monafath,
        }
        df_type = df[sel]
        for annee in sorted(df_type["السنة"].dropna().unique()):
            df_annee = df_type[df_type["السنة"] == annee]
            reste = len(df_annee) - int(execute[df_annee.index].sum())
            ligne[str(int(annee))] = max(0, reste)
        lignes.append(ligne)

    resultats = pd.DataFrame(lignes)
    cols_annees = [c for c in resultats.columns if c.isdigit()]
    if cols_annees:
        resultats["مجموع السنوات"] = resultats[cols_annees].sum(axis=1)
    total = {"نوعية القضايا": "المجموع"}
    total.update({c: resultats[c].sum() for c in resultats.columns if c != "نوعية القضايا"})
    return pd.concat([resultats, pd.DataFrame([total])], ignore_index=True)


# ---------------------------------------------------------------------------
# تتبع النشاط العام للمحكمة
# ---------------------------------------------------------------------------
def preparer_nachate(df):
    """Calcule المحكوم quand le fichier ne donne que المخلف المحكوم et المسجل المحكوم."""
    df = df.copy()
    if "المحكوم" not in df.columns and {"المخلف المحكوم", "المسجل المحكوم"} <= set(df.columns):
        df["المحكوم"] = df["المخلف المحكوم"] + df["المسجل المحكوم"]
    return df


def activite_mensuelle(fichiers, debut_mm_aaaa, fin_mm_aaaa):
    """Somme mensuelle de المحكوم, المخلف et المسجل entre deux mois au format « mm-aaaa »."""
    df = pd.concat(fichiers, ignore_index=True)
    df["تاريخ"] = pd.to_datetime(df["تاريخ"], errors="coerce")
    df = df.dropna(subset=["تاريخ"])
    df["شهر"] = df["تاريخ"].dt.strftime("%Y-%m")
    debut = pd.to_datetime("01-" + debut_mm_aaaa, format="%d-%m-%Y")
    fin = pd.to_datetime("01-" + fin_mm_aaaa, format="%d-%m-%Y") + pd.offsets.MonthEnd(0)
    df = df[(df["تاريخ"] >= debut) & (df["تاريخ"] <= fin)]
    return df.groupby("شهر")[["المحكوم", "المخلف", "المسجل"]].sum().sort_index()
