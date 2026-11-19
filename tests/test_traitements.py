import pandas as pd
import pytest

from app import config
from app import traitements as tr


def test_extraire_partie():
    assert tr.extraire_partie("12/2101/2024") == "2101"
    assert tr.extraire_partie("12/2101/2024", position=2) == "2024"
    assert tr.extraire_partie("sans-barre") is None
    assert tr.extraire_partie("12") is None


def test_chercher_shi3ba_et_naw3iya():
    assert tr.chercher_shi3ba(2401, config.CODES_SHI3AB_ZAJRI) == "حوادث السير"
    assert tr.chercher_shi3ba(9999, config.CODES_SHI3AB_ZAJRI) == config.INCONNU_SHI3BA
    assert tr.chercher_naw3iya("6153", config.NAW3IYA_TANFID_MADANI) == "الاستعجالي"
    assert tr.chercher_naw3iya("6150", config.NAW3IYA_TANFID_MADANI) == "الأوامر بالأداء"
    assert tr.chercher_naw3iya("1", config.NAW3IYA_TANFID_MADANI) == config.INCONNU_NAW3IYA


def _stats(numeros):
    return tr.preparer_fichier_stats(pd.DataFrame({
        "رقم الملف": numeros, "نوع القضية": ["جنحي"] * len(numeros), "الهيئة": ["الهيئة 1"] * len(numeros)}))


def test_stats_generales():
    mokhalaf = _stats(["1/2101/2023", "2/2101/2023", "3/2401/2023"])
    mossajal = _stats(["4/2101/2024", "4/2101/2024"])  # doublon supprimé
    mahkoum = _stats(["1/2101/2023"])
    df = tr.stats_generales(mokhalaf, mossajal, mahkoum)
    ligne = df[df["الرمز"] == "2101"].iloc[0]
    assert (ligne["المخلف"], ligne["المسجل"], ligne["الرائج"], ligne["المحكوم"], ligne["الباقي"]) == (2, 1, 3, 1, 2)
    total = df[df["الرمز"] == "المجموع"].iloc[0]
    assert total["الرائج"] == 4 and total["الباقي"] == 3


def test_stats_par_hay2a_total():
    df = tr.stats_par_hay2a(_stats(["1/2101/2023"]), _stats(["2/2101/2024"]), _stats(["1/2101/2023"]))
    total = df[df["الهيئة"] == "المجموع"].iloc[0]
    assert (total["الرائج"], total["الباقي"]) == (2, 1)


def test_pivot_juges_totaux():
    df = pd.DataFrame({"الشعبة": ["أ", "أ", "ب"], "القاضي": ["x", "y", "x"]})
    pivot = tr.pivot_juges(df, "القاضي")
    total = pivot[pivot["index"] == "المجموع"].iloc[0]
    assert (total["x"], total["y"], total["المجموع"]) == (2, 1, 3)


def test_pivot_chronic():
    df = tr.preparer_chronic(pd.DataFrame({"رقم الملف": ["1/2101/2019", "2/2101/2019", "3/2401/2020"]}))
    pivot = tr.pivot_chronic(df)
    assert pivot.columns[0] == "الشعبة"
    assert pivot[pivot["الشعبة"] == "المجموع"]["المجموع"].iloc[0] == 3


def test_convertir_dates_formats_du_tribunal():
    dates = tr.convertir_dates(pd.Series(["févr 12 2025 10:30AM", "juil  3 2025 03:05PM", "12/02/2025", None]))
    assert dates.tolist()[:3] == [pd.Timestamp("2025-02-12"), pd.Timestamp("2025-07-03"), pd.Timestamp("2025-02-12")]
    assert pd.isna(dates.iloc[3])


def test_stats_tanfid():
    df = pd.DataFrame({
        "الرقم الكامل للملف": ["2023/6151/1", "2024/6151/2", "2024/6151/3"],
        "تاريخ التسجيل بالسجل": pd.to_datetime(["2023-05-10", "2024-02-15", "2024-03-20"]),
        "حالة الملف": [None, "منفذ", "غير منفذ"],
        "تاريخ الحالة": [None, "mars  1 2024 10:00AM", None],
    })
    res = tr.stats_tanfid([df], pd.Timestamp("2024-01-01"), pd.Timestamp("2024-12-31"), config.NAW3IYA_TANFID_MADANI)
    ligne = res[res["نوعية القضايا"] == "الأوامر المبنية على طلب"].iloc[0]
    assert ligne["المسجل"] == 2
    assert ligne["المنفذ"] == 1
    assert ligne["الباقي"] == ligne["المجموع"] - ligne["المنفذ"]
    assert res.iloc[-1]["نوعية القضايا"] == "المجموع"


def test_activite_mensuelle():
    df = tr.preparer_nachate(pd.DataFrame({
        "تاريخ": ["2024-01-01", "2024-02-01", "2025-01-01"],
        "المخلف": [10, 20, 30], "المسجل": [1, 2, 3],
        "المخلف المحكوم": [1, 1, 1], "المسجل المحكوم": [2, 2, 2]}))
    par_mois = tr.activite_mensuelle([df], "01-2024", "12-2024")
    assert list(par_mois.index) == ["2024-01", "2024-02"]
    assert par_mois["المحكوم"].tolist() == [3, 3]


def test_export_excel(tmp_path):
    from openpyxl import load_workbook

    from app.export_excel import enregistrer_classeur
    df = pd.DataFrame({"نوع القضية": ["جنحي", ""], "الرمز": ["2101", "المجموع"], "المخلف": [3, 3]})
    chemin = tmp_path / "test.xlsx"
    enregistrer_classeur(chemin, [("النتائج", df)])
    ws = load_workbook(chemin)["النتائج"]
    assert ws.cell(row=6, column=1).value == "نوع القضية"
    assert ws.cell(row=8, column=1).value == "المجموع"
    assert ws.cell(row=8, column=3).value == 3


@pytest.mark.parametrize("table", [config.CODES_SHI3AB_ZAJRI, config.CODES_SHI3AB_MADANI, config.CODES_SHI3AB_ALSARA])
def test_tables_sans_code_en_double(table):
    tous = [c for codes in table.values() for c in codes]
    assert len(tous) == len(set(tous))
