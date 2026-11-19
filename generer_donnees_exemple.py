"""Génère des fichiers Excel FICTIFS, au même format que les exports du tribunal.

Numéros de dossier, juges et formations sont inventés : aucune donnée réelle.
Usage : python generer_donnees_exemple.py
"""

import random
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd

from app import config

random.seed(2025)
DOSSIER = Path(__file__).parent / "donnees_exemple"
JUGES = ["القاضي أ", "القاضي ب", "القاضي ج", "القاضي د"]
HAY2AT = ["شعبة الجنحي العادي جلسة الاثنين", "شعبة الجنحي العادي جلسة الأربعاء",
          "قضايا حوادث السير جلسة الثلاثاء", "هيئة غرفة المشورة و الصلح"]
NOUW3 = ["جنحي عادي ضبطي", "جنحي تلبسي ضبطي سراح", "حوادث السير عادي ضبطي", "جنحي عادي تأديبي"]
MOIS_FR = ["janv", "févr", "mars", "avr", "mai", "juin", "juil", "août", "sept", "oct", "nov", "déc"]


def codes(table):
    return [c for liste in table.values() for c in liste]


def jour(debut, fin):
    return debut + timedelta(days=random.randint(0, (fin - debut).days))


def date_systeme(d):
    """Format de la colonne تاريخ الحالة dans les exports : « févr 12 2025 10:30AM »."""
    heure = datetime(d.year, d.month, d.day, random.randint(8, 16), random.choice([0, 15, 30, 45]))
    return f"{MOIS_FR[d.month - 1]} {d.day:>2} {d.year} {heure:%I:%M%p}"


def fichier_zajri(n):
    """المخلف / المسجل / المحكوم : رقم الملف au format n/رمز/سنة."""
    lignes = []
    for _ in range(n):
        inscrit = jour(date(2023, 1, 1), date(2025, 6, 30))
        lignes.append({
            "رقم الملف": f"{random.randint(1, 3500)}/{random.choice(codes(config.CODES_SHI3AB_ZAJRI))}/{inscrit.year}",
            "نوع القضية": random.choice(NOUW3),
            "تاريخ تسجيل الملف": pd.Timestamp(inscrit),
            "آخر جلسة": pd.Timestamp(jour(inscrit, date(2025, 7, 1))),
            "المقرر": random.choice(JUGES),
            "الهيئة": random.choice(HAY2AT),
            "رقم الحكم": float(random.randint(1, 1200)),
        })
    return pd.DataFrame(lignes)


def fichier_fusion(table, n):
    """سجل الأحكام (المدني / الأسرة) : الرقم الكامل للملف au format سنة/رمز/n."""
    return pd.DataFrame([{
        "الرقم الكامل للملف": f"{random.choice([2024, 2025])}/{random.choice(codes(table))}/{random.randint(1, 99)}",
        "نوع الحكم": random.choice(["قبول الطلب", "رفض الطلب"]),
        "تاريخ الحكم ": jour(date(2025, 1, 1), date(2025, 6, 30)).strftime("%d/%m/%Y"),
        "القاضي أَو المستشار المقرّر": random.choice(JUGES),
    } for _ in range(n)])


def fichier_tanfid(table, n):
    """السجل العام لتسجيل ملفات التنفيذ : الرقم الكامل للملف au format سنة/رمز/n."""
    cles = [c for cle in table for c in cle.split(",") if c.isdigit()]
    lignes = []
    for i in range(1, n + 1):
        inscrit = jour(date(2023, 1, 1), date(2025, 6, 30))
        etat = random.choice(["منفذ", "غير منفذ", "مواصلة التنفيذ", None])
        lignes.append({
            "الرقم التسلسلي للملف": i,
            "الرقم الكامل للملف": f"{inscrit.year}/{random.choice(cles)}/{i}",
            "تاريخ التسجيل بالسجل": pd.Timestamp(inscrit),
            "حالة الملف": etat,
            "تاريخ الحالة": date_systeme(jour(inscrit, date(2025, 7, 1))) if etat else None,
        })
    return pd.DataFrame(lignes)


def fichier_nachate():
    lignes = []
    for mois in range(1, 13):
        mokhalaf, mossajal = random.randint(2000, 2600), random.randint(250, 400)
        m_mahkoum, s_mahkoum = random.randint(150, 300), random.randint(80, 200)
        lignes.append({
            "المخلف": mokhalaf, "المسجل": mossajal, "الرائجة": mokhalaf + mossajal,
            "المخلف المحكوم": m_mahkoum, "المسجل المحكوم": s_mahkoum, "الرائجة المحكومة": m_mahkoum + s_mahkoum,
            "تاريخ": f"2024-{mois:02d}-01",
        })
    return pd.DataFrame(lignes)


def main():
    DOSSIER.mkdir(exist_ok=True)
    fichiers = {
        "1_zajri_mokhalaf.xlsx": fichier_zajri(250),
        "2_zajri_mossajal.xlsx": fichier_zajri(180),
        "3_zajri_mahkoum.xlsx": fichier_zajri(160),
        "4_zajri_ahkam.xlsx": pd.DataFrame([{
            "رقم الملف": f"{random.randint(1, 3500)}/{random.choice(codes(config.CODES_SHI3AB_ZAJRI))}/2025",
            "نوع القضية": random.choice(NOUW3),
            "تاريخ الحكم": pd.Timestamp(jour(date(2025, 1, 1), date(2025, 6, 30))),
            "القاضي المقرر": random.choice(JUGES),
        } for _ in range(200)]),
        "5_zajri_chronic.xlsx": pd.DataFrame([{
            "رقم الملف": f"{random.randint(1, 3500)}/{random.choice(codes(config.CODES_SHI3AB_ZAJRI))}/{annee}",
            "طبيعة الملف": random.choice(NOUW3),
            "الهيئة": random.choice(HAY2AT),
        } for annee in random.choices([2018, 2019, 2020, 2021, 2022], k=150)]),
        "6_madani_fusion.xlsx": fichier_fusion(config.CODES_SHI3AB_MADANI, 140),
        "7_alsara_fusion.xlsx": fichier_fusion(config.CODES_SHI3AB_ALSARA, 110),
        "8_tanfid_madani.xlsx": fichier_tanfid(config.NAW3IYA_TANFID_MADANI, 200),
        "9_tanfid_alsara.xlsx": fichier_tanfid(config.NAW3IYA_TANFID_ALSARA, 120),
        "10_nachate.xlsx": fichier_nachate(),
    }
    for nom, df in fichiers.items():
        df.to_excel(DOSSIER / nom, index=False)
    print(f"{len(fichiers)} fichiers fictifs créés dans {DOSSIER}")


if __name__ == "__main__":
    main()
