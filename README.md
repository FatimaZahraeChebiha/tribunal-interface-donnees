# Interface de traitement des données d'un tribunal

Application de bureau (Python, Tkinter) qui transforme les exports Excel du système de gestion des dossiers d'un tribunal en tableaux statistiques, graphiques et rapports Excel prêts à l'emploi.

Projet réalisé pendant mon **stage d'initiation au Tribunal de Première Instance de Taounate (juillet 2025)**, en 1re année du cycle ingénieur Science des Données et Intelligence Artificielle à l'ENSA Fès. Le code, écrit à l'origine dans un notebook Jupyter, a ensuite été réorganisé en projet Python structuré et testé.

## Le problème

Chaque mois, le greffe prépare des statistiques sur l'activité du tribunal : dossiers restants, enregistrés, jugés, production de chaque juge, dossiers anciens, exécution des jugements civils. Ces chiffres étaient calculés à la main à partir de plusieurs fichiers Excel, ce qui prenait du temps et laissait passer des erreurs.

L'application automatise ces calculs : on charge les fichiers exportés, on clique sur « traiter », et on obtient le tableau, le graphique et le fichier Excel mis en forme.

Pendant le stage, elle a été utilisée sur les exports réels du tribunal : une douzaine de registres représentant plus de 31 000 lignes au total (jusqu'à près de 13 000 lignes pour le seul registre des jugements). Certaines sections acceptent plusieurs fichiers à la fois : les registres de jugements civils et familiaux exportés en plusieurs parties sont fusionnés en un seul tableau, et les registres d'exécution de plusieurs années (2022 à 2025) sont traités ensemble.

## Fonctionnalités

L'interface est en arabe, comme les données et les utilisateurs du tribunal.

**Section pénale (الزجري)**
- **Statistiques générales** : à partir des fichiers des dossiers restants (المخلف), enregistrés (المسجل) et jugés (المحكوم), calcul par code d'affaire et par formation (الهيئة) des dossiers en cours (الرائج) et restants (الباقي).
- **Production des juges** : nombre de dossiers jugés par chaque juge et par chambre (الشعبة), avec un graphique.
- **Dossiers anciens (الملفات المزمنة)** : nombre de dossiers restants par chambre et par année, avec un graphique.

**Section civile (المدني)**
- **Production des juges** (civil et famille) : fusion de plusieurs exports Excel, puis tableau par chambre et par juge.
- **Statistiques d'exécution** (civil et famille) : choix d'une période dans un calendrier, puis calcul par type d'affaire des dossiers restants, enregistrés, exécutés et restants par année.

**Suivi de l'activité générale** : courbes mensuelles des dossiers jugés, restants et enregistrés sur une période choisie.

Chaque tableau peut être exporté vers Excel avec un en-tête, des bordures, la ligne du total mise en valeur et, selon le cas, une feuille contenant le graphique.

## Technologies

Python, Tkinter (interface), pandas (traitement des données), Matplotlib (graphiques), openpyxl (export Excel), arabic-reshaper et python-bidi (affichage correct de l'arabe dans les graphiques), tkcalendar (choix des dates).

## Structure

```
main.py                      lancement de l'application
app/
  config.py                  tables des codes d'affaires, noms des colonnes, styles
  traitements.py             calculs pandas (sans interface, testables)
  graphiques.py              graphiques Matplotlib avec texte arabe
  export_excel.py            mise en forme et export des classeurs Excel
  interface.py               fenêtres et boutons Tkinter
generer_donnees_exemple.py   création de fichiers Excel fictifs
donnees_exemple/             fichiers fictifs pour essayer l'application
tests/                       tests des calculs (pytest)
```

## Installation et lancement

```bash
git clone https://github.com/FatimaZahraeChebiha/tribunal-interface-donnees.git
cd tribunal-interface-donnees
python -m venv venv
venv\Scripts\activate          # Windows (sous Linux ou macOS : source venv/bin/activate)
pip install -r requirements.txt
python main.py
```

Le logo affiché en haut de la fenêtre et dans les fichiers exportés se trouve dans `assets/logo.png`. Il appartient au Ministère de la Justice et est utilisé avec l'autorisation du tribunal ; il ne peut pas être réutilisé en dehors de ce projet.

## Essayer avec les données d'exemple

Le dossier `donnees_exemple/` contient des fichiers Excel **entièrement fictifs** (numéros de dossier, juges et formations inventés), au même format que les exports du tribunal. Ils contiennent volontairement peu de lignes (quelques centaines au plus) : ils servent uniquement à essayer l'application. Ils peuvent être régénérés avec :

```bash
python generer_donnees_exemple.py
```

| Fichier | À charger dans |
|---|---|
| `1_zajri_mokhalaf`, `2_zajri_mossajal`, `3_zajri_mahkoum` | الزجري > إحصائيات عامة |
| `4_zajri_ahkam` | الزجري > انتاجيات القضاة |
| `5_zajri_chronic` | الزجري > الملفات المزمنة |
| `6_madani_fusion`, `7_alsara_fusion` | المدني > انتاجيات القضاة > المدني / الأسرة |
| `8_tanfid_madani`, `9_tanfid_alsara` | المدني > إحصائيات تنفيذ المدني > المدني / الأسرة |
| `10_nachate` | + إضافة ملف نشاط, puis تتبع النشاط العام للمحكمة |

## Tests

```bash
pip install pytest
python -m pytest
```

Les tests vérifient les calculs principaux (statistiques générales, tableaux par juge, dossiers anciens, exécution, activité mensuelle) et l'export Excel.

## Confidentialité

Les données réelles du tribunal ne sont pas publiées. Le fichier `.gitignore` exclut les fichiers Excel et CSV, sauf les exemples fictifs.

## Auteure

Fatima Zahrae Chebiha, élève ingénieure en Science des Données et Intelligence Artificielle, ENSA Fès.
[LinkedIn](https://www.linkedin.com/in/fatima-zahrae-chebiha) · [GitHub](https://github.com/FatimaZahraeChebiha)
