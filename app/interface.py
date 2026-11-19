"""Interface Tkinter de l'application de traitement des données du tribunal."""

import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import matplotlib.pyplot as plt
import pandas as pd
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from PIL import Image, ImageTk
from tkcalendar import DateEntry

from . import config, graphiques
from . import traitements as tr
from .export_excel import enregistrer_classeur, generer_nom_fichier

FILTRE_EXCEL = [("Fichiers Excel", "*.xlsx *.xls")]
FILTRE_EXCEL_CSV = [("Fichiers Excel ou CSV", "*.csv *.xls *.xlsx")]
PERIODE_VIDE = "لم يتم اختيار الفترة بعد"


def basculer(cadre, **options_pack):
    """Affiche le cadre s'il est caché, le cache s'il est affiché."""
    if cadre.winfo_ismapped():
        cadre.pack_forget()
    else:
        cadre.pack(**options_pack)


class SectionTanfid:
    """Bloc « إحصائيات التنفيذ » (utilisé pour المدني et pour الأسرة)."""

    def __init__(self, app, parent, table_naw3iya, titre_chargement, titre_fichiers, titre_periode, nom_export):
        self.app = app
        self.table_naw3iya = table_naw3iya
        self.titre_fichiers = titre_fichiers
        self.titre_periode = titre_periode
        self.nom_export = nom_export
        self.fichiers = []
        self.resultats = pd.DataFrame()
        self.date_debut = None
        self.date_fin = None
        self.statut = tk.StringVar(value="")

        self.cadre = tk.Frame(parent, bg="white")
        petit = {"font": (config.POLICE, 10, "bold"), "fg": "white", "width": 15}
        tk.Button(self.cadre, text=titre_chargement, command=self.charger, bg="#9C27B0", **petit).pack(pady=3)
        tk.Label(self.cadre, textvariable=self.statut, font=(config.POLICE, 10, "bold"), bg="white").pack()
        tk.Button(self.cadre, text="اختيار الفترة الزمنية", command=self.choisir_periode, bg="#FF5722", **petit).pack(pady=3)
        self.label_periode = tk.Label(self.cadre, text=PERIODE_VIDE, font=(config.POLICE, 9), bg="white", fg="blue")
        self.label_periode.pack(pady=2)
        tk.Button(self.cadre, text="حذف الملفات", command=self.vider, bg="#DC3545", **petit).pack(pady=3)
        tk.Button(self.cadre, text="معالجة البيانات", command=self.traiter, bg="#4CAF50", **petit).pack(pady=3)
        tk.Button(self.cadre, text="تصدير إلى Excel", command=self.exporter, bg="#2196F3", **petit).pack(pady=3)

    def charger(self):
        chemins = filedialog.askopenfilenames(title=self.titre_fichiers, filetypes=FILTRE_EXCEL)
        if not chemins:
            self.statut.set("❌ لم يتم اختيار ملفات")
            return
        try:
            self.fichiers = [pd.read_excel(c).dropna(axis=1, how="all") for c in chemins]
            self.statut.set(f"✅ تم تحميل {len(chemins)} ملف")
            messagebox.showinfo("نجح", f"تم تحميل {len(chemins)} ملف بنجاح!")
        except Exception as e:
            self.fichiers = []
            self.statut.set("❌ خطأ في التحميل")
            messagebox.showerror("خطأ", f"حدث خطأ أثناء تحميل الملفات:\n{e}")

    def vider(self):
        self.fichiers = []
        self.date_debut = self.date_fin = None
        self.statut.set("")
        self.label_periode.config(text=PERIODE_VIDE)
        messagebox.showinfo("تم", "تم حذف جميع الملفات المحملة بنجاح!")

    def choisir_periode(self):
        fenetre = tk.Toplevel(self.app.root)
        fenetre.title(self.titre_periode)
        fenetre.geometry("500x400")
        cadre = tk.Frame(fenetre)
        cadre.pack(padx=20, pady=20, fill="both", expand=True)

        calendriers = []
        for libelle in ["تاريخ البداية:", "تاريخ النهاية:"]:
            bloc = tk.Frame(cadre)
            bloc.pack(pady=10, fill="x")
            tk.Label(bloc, text=libelle, font=(config.POLICE, 12)).pack()
            cal = DateEntry(bloc, width=12, background="darkblue", foreground="white", borderwidth=2,
                            date_pattern="dd/mm/yyyy", locale="ar_MA")
            cal.pack(pady=5)
            calendriers.append(cal)

        def valider():
            debut = pd.to_datetime(calendriers[0].get_date())
            fin = pd.to_datetime(calendriers[1].get_date())
            if debut > fin:
                messagebox.showerror("خطأ", "تاريخ البداية يجب أن يكون قبل تاريخ النهاية")
                return
            self.definir_periode(debut, fin)
            fenetre.destroy()

        tk.Button(cadre, text="تأكيد", command=valider, bg="#4CAF50", fg="white",
                  font=(config.POLICE, 12), width=15).pack(pady=20)

    def definir_periode(self, debut, fin):
        self.date_debut, self.date_fin = debut, fin
        self.label_periode.config(text=f"الفترة: من {debut:%d/%m/%Y} إلى {fin:%d/%m/%Y}")

    def traiter(self):
        if not self.fichiers:
            messagebox.showinfo("تنبيه", "يرجى تحميل الملفات أولاً")
            return
        if self.date_debut is None or self.date_fin is None:
            messagebox.showinfo("تنبيه", "يرجى اختيار الفترة الزمنية أولاً")
            return
        try:
            self.resultats = tr.stats_tanfid(self.fichiers, self.date_debut, self.date_fin, self.table_naw3iya)
            self.app.afficher_tableau(self.resultats)
        except Exception as e:
            messagebox.showerror("خطأ", f"حدث خطأ أثناء معالجة البيانات:\n{e}")

    def exporter(self):
        if self.resultats.empty:
            messagebox.showinfo("تنبيه", "لا توجد نتائج للتصدير")
            return
        self.app.exporter([("النتائج", self.resultats)], self.nom_export)


class SectionFusion:
    """Bloc « انتاجيات القضاة » de la section المدني : fusion de fichiers puis tableau par juge."""

    def __init__(self, app, parent, table_shi3ab, ordre_colonnes, titres, nom_export, couleur_export):
        self.app = app
        self.table_shi3ab = table_shi3ab
        self.ordre_colonnes = ordre_colonnes  # "shi3ba_a_droite" (المدني) ou "shi3ba_a_gauche" (الأسرة)
        self.titres = titres
        self.nom_export = nom_export
        self.liste = []
        self.fichier_fusionne = pd.DataFrame()
        self.statut = tk.StringVar(value="")

        self.cadre = tk.Frame(parent, bg="white")
        bouton = {"font": (config.POLICE, 10, "bold"), "fg": "white", "width": 12}
        tk.Button(self.cadre, text="Excel إضافة ملف", command=self.ajouter, bg="#800080", **bouton).pack(pady=5)
        self.label_nombre = tk.Label(self.cadre, text="0", font=(config.POLICE, 10), bg="white", fg="green")
        self.label_nombre.pack(pady=5)
        tk.Button(self.cadre, text="تفريغ قائمة الملفات", command=self.vider, bg="#FF5722", **bouton).pack(pady=5)
        tk.Button(self.cadre, text="Excel دمج ملفات", command=self.fusionner, bg="#800080", **bouton).pack(pady=5)
        tk.Button(self.cadre, text="تحميل الملف المدمج", command=self.charger_fusionne, bg="#800080", **bouton).pack(pady=5)
        tk.Label(self.cadre, textvariable=self.statut, font=(config.POLICE, 12, "bold"), bg="white").pack()
        tk.Button(self.cadre, text="عرض الجدول المدمج", command=self.afficher, bg="#800080", **bouton).pack(pady=5)
        tk.Button(self.cadre, text="Excel تصدير إلى", command=self.exporter, bg=couleur_export, **bouton).pack(pady=5)

    def ajouter(self):
        chemins = filedialog.askopenfilenames(title=self.titres["ajout"], filetypes=FILTRE_EXCEL)
        if not chemins:
            return
        try:
            self.liste.extend(pd.read_excel(c) for c in chemins)
            self.label_nombre.config(text=str(len(self.liste)))
        except Exception as e:
            messagebox.showerror("خطأ", f"خطأ في إضافة الملفات: {e}")

    def vider(self):
        self.liste = []
        self.label_nombre.config(text="0")

    def fusionner(self):
        if not self.liste:
            messagebox.showinfo("تنبيه", "لا توجد ملفات للدمج")
            return
        chemin = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=self.titres["fichier_fusionne"],
                                              filetypes=[("Excel files", "*.xlsx")])
        if not chemin:
            return
        try:
            pd.concat(self.liste, ignore_index=True).to_excel(chemin, index=False)
            messagebox.showinfo("نجح", f"تم حفظ الملف المدموج في:\n{chemin}")
        except Exception as e:
            messagebox.showerror("خطأ", f"خطأ في دمج الملفات: {e}")

    def charger_fusionne(self):
        chemin = filedialog.askopenfilename(title=self.titres["chargement"], filetypes=FILTRE_EXCEL_CSV)
        if not chemin:
            self.statut.set("❌")
            return
        try:
            self.fichier_fusionne = tr.preparer_fichier_fusionne(tr.lire_tableau(chemin), self.table_shi3ab)
            self.statut.set("✅")
        except Exception as e:
            self.statut.set("❌")
            messagebox.showerror("خطأ", f"خطأ في تحميل الملف: {e}")

    def tableau(self):
        """Tableau par شعبة et par juge, ou None si les données manquent (avec un message)."""
        if self.fichier_fusionne.empty:
            messagebox.showinfo("تنبيه", "يرجى تحميل الملف المدمج أولاً")
            return None
        df = self.fichier_fusionne
        if not tr.colonnes_presentes(df, [config.COL_JUGE_MADANI, "الرمز", "الشعبة"]):
            messagebox.showerror("خطأ", "الأعمدة المطلوبة غير موجودة في الملف")
            return None
        pivot = tr.pivot_juges(df, config.COL_JUGE_MADANI)
        return pivot.rename(columns={pivot.columns[0]: "الشعبة"})

    def ordonner(self, pivot):
        """المدني : المجموع, juges, الشعبة (lecture de droite à gauche). الأسرة : الشعبة, المجموع, juges."""
        juges = [c for c in pivot.columns if c not in ["المجموع", "الشعبة"]]
        if self.ordre_colonnes == "shi3ba_a_droite":
            return pivot[["المجموع"] + juges + ["الشعبة"]]
        return pivot[["الشعبة", "المجموع"] + juges]

    def afficher(self):
        try:
            pivot = self.tableau()
            if pivot is not None:
                self.app.afficher_tableau(self.ordonner(pivot))
        except Exception as e:
            messagebox.showerror("خطأ", f"خطأ في عرض البيانات: {e}")

    def exporter(self):
        try:
            pivot = self.tableau()
            if pivot is None:
                return
            if self.ordre_colonnes == "shi3ba_a_gauche":
                pivot = self.ordonner(pivot)
            self.app.exporter([("النتائج", pivot)], self.nom_export)
        except Exception as e:
            messagebox.showerror("خطأ", f"فشل التصدير: {e}")


class Application:
    def __init__(self, root):
        self.root = root
        root.title("تطبيق معالجة بيانات المحكمة")
        root.geometry("1200x750")
        root.configure(bg="white")

        # Données chargées (remplacent les variables globales du notebook)
        self.fichiers_stats = {"المخلف": None, "المسجل": None, "المحكوم": None}
        self.statuts_stats = {nom: tk.StringVar() for nom in self.fichiers_stats}
        self.df_final = pd.DataFrame()
        self.df_par_hay2a = pd.DataFrame()
        self.df_ahkam = pd.DataFrame()
        self.pivot_ahkam = pd.DataFrame()
        self.statut_ahkam = tk.StringVar()
        self.df_chronic = pd.DataFrame()
        self.pivot_chronic = pd.DataFrame()
        self.statut_chronic = tk.StringVar()
        self.fichiers_nachate = []

        self._cadre_defilant()
        self._logo()
        self._sections_principales()
        self._section_madani()
        self._section_zajri()
        self._boutons_nachate()

    # ------------------------------------------------------------------ cadre
    def _cadre_defilant(self):
        canvas = tk.Canvas(self.root, bg="white")
        barre = tk.Scrollbar(self.root, orient="vertical", command=canvas.yview)
        barre.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        canvas.configure(yscrollcommand=barre.set)
        self.main_frame = tk.Frame(canvas, bg="white")
        canvas.create_window((0, 0), window=self.main_frame, anchor="nw")
        self.main_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1 * (e.delta / 120)), "units"))

    def _logo(self):
        if not config.LOGO.exists():
            return
        try:
            image = ImageTk.PhotoImage(Image.open(config.LOGO).resize(config.TAILLE_LOGO))
        except OSError as e:
            print(f"Logo ignoré : {e}")
            return
        label = tk.Label(self.main_frame, image=image, bg="white")
        label.image = image
        label.pack(pady=10, anchor="center")

    def _sections_principales(self):
        cadre = tk.Frame(self.main_frame, bg="white")
        cadre.pack(pady=20)
        grand = {"fg": "white", "font": (config.POLICE, 16, "bold"), "width": 20, "height": 2}
        tk.Button(cadre, text="الــزجــري", bg="#D32F2F", command=self.basculer_zajri, **grand).pack(side="left", padx=150)
        tk.Button(cadre, text="الـمــدنــي", bg="#388E3C", command=self.basculer_madani, **grand).pack(side="left", padx=200)
        self.frame_zajri_buttons = tk.Frame(self.main_frame, bg="white")
        self.frame_madani_buttons = tk.Frame(self.main_frame, bg="white")
        self.zajri_visible = False
        self.madani_visible = False

    # ----------------------------------------------------------------- المدني
    def _section_madani(self):
        parent = self.frame_madani_buttons
        grand = {"font": (config.POLICE, 16, "bold"), "bg": "#007BFF", "fg": "white", "width": 20}
        jaune = {"font": (config.POLICE, 12, "bold"), "bg": "#FFC107", "fg": "white", "width": 15}

        cadre_ahkam = tk.Frame(parent, bg="white")
        cadre_tanfid = tk.Frame(parent, bg="white")
        tk.Button(parent, text="انتاجيات القضاة", command=lambda: basculer(cadre_ahkam, pady=5), **grand).pack(pady=5)
        tk.Button(parent, text="إحصائيات تنفيذ المدني", command=lambda: basculer(cadre_tanfid, pady=5), **grand).pack(pady=5)

        # إحصائيات تنفيذ المدني : deux blocs identiques (المدني, الأسرة)
        self.tanfid_madani = SectionTanfid(self, cadre_tanfid, config.NAW3IYA_TANFID_MADANI, "تحميل ملفات التنفيذ",
                                           "اختر ملفات Excel للتنفيذ المدني", "اختيار الفترة الزمنية",
                                           "إحصائيات_التنفيذ_المدني")
        tk.Button(cadre_tanfid, text="المدني", command=lambda: basculer(self.tanfid_madani.cadre, pady=5), **jaune).pack(pady=5)
        self.tanfid_alsara = SectionTanfid(self, cadre_tanfid, config.NAW3IYA_TANFID_ALSARA, "تحميل ملفات الأسرة",
                                           "اختر ملفات Excel للأسرة", "اختيار الفترة الزمنية للأسرة",
                                           "إحصائيات_الأسرة")
        tk.Button(cadre_tanfid, text="الأسرة", command=lambda: basculer(self.tanfid_alsara.cadre, pady=5), **jaune).pack(pady=5)

        # انتاجيات القضاة : fusion des fichiers puis tableau par juge
        self.fusion_madani = SectionFusion(
            self, cadre_ahkam, config.CODES_SHI3AB_MADANI, "shi3ba_a_droite",
            {"ajout": "اختر ملفات Excel", "fichier_fusionne": "fichier_fusionne.xlsx",
             "chargement": "اختر الملف المدموج"},
            "جدول الملفات المدمجة", "#008000")
        tk.Button(cadre_ahkam, text="المدني", command=lambda: basculer(self.fusion_madani.cadre, pady=5), **jaune).pack(pady=5)
        self.fusion_alsara = SectionFusion(
            self, cadre_ahkam, config.CODES_SHI3AB_ALSARA, "shi3ba_a_gauche",
            {"ajout": "اختر ملفات Excel للأسرة", "fichier_fusionne": "fichier_alsara_fusionne.xlsx",
             "chargement": "اختر الملف المدموج للأسرة"},
            "جدول_الأسرة_المدموج", "#008000")
        tk.Button(cadre_ahkam, text="الأسرة", command=lambda: basculer(self.fusion_alsara.cadre, pady=5), **jaune).pack(pady=5)

    # ----------------------------------------------------------------- الزجري
    def _section_zajri(self):
        self.frame_buttons = tk.Frame(self.main_frame, bg="white")
        self.frame_ahkam = tk.Frame(self.main_frame, bg="white")
        self.frame_result = tk.Frame(self.main_frame, bg="white")
        self.frame_result.pack(fill="both", expand=True, padx=10, pady=10)
        self.frame_chronic = tk.Frame(self.main_frame, bg="white")
        self.etat_sections = {self.frame_buttons: False, self.frame_ahkam: False, self.frame_chronic: False}

        for texte, cadre in [("إحصائيات عامة", self.frame_buttons), ("انتاجيات القضاة", self.frame_ahkam),
                             ("الملفات المزمنة", self.frame_chronic)]:
            tk.Button(self.frame_zajri_buttons, text=texte, command=lambda c=cadre: self.basculer_section(c),
                      **config.BOUTON_SECTION).pack(pady=5)

        # إحصائيات عامة
        for nom in self.fichiers_stats:
            self._ligne_chargement(self.frame_buttons, f"تحميل ملف {nom}", lambda n=nom: self.charger_fichier_stats(n),
                                   self.statuts_stats[nom])
        tk.Button(self.frame_buttons, text="معالجة البيانات", command=self.traiter_stats, bg="#4CAF50", fg="white").pack(pady=10)
        tk.Button(self.frame_buttons, text="عرض حسب الهيئة", command=self.afficher_par_hay2a, bg="#FF9800", fg="white").pack(pady=5)
        tk.Button(self.frame_buttons, text="Excel تصدير إلى", command=self.exporter_stats, bg="#2196F3", fg="white").pack(pady=5)

        # انتاجيات القضاة
        self._ligne_chargement(self.frame_ahkam, "تحميل سجل الأحكام", self.charger_ahkam, self.statut_ahkam)
        tk.Button(self.frame_ahkam, text="عرض الأحكام", command=self.afficher_ahkam, bg="#4CAF50", fg="white").pack(pady=10)
        tk.Button(self.frame_ahkam, text="عرض الرسوم البيانية", command=self.graphique_ahkam, bg="#FF9800", fg="white").pack(pady=10)
        tk.Button(self.frame_ahkam, text="Excel تصدير إلى", command=self.exporter_ahkam, bg="#2196F3", fg="white").pack(pady=5)

        # الملفات المزمنة
        self._ligne_chargement(self.frame_chronic, "تحميل الملفات المزمنة", self.charger_chronic, self.statut_chronic)
        tk.Button(self.frame_chronic, text="عرض الإحصائيات", command=self.afficher_chronic, bg="#4CAF50", fg="white").pack(pady=10)
        tk.Button(self.frame_chronic, text="عرض الرسوم البيانية", command=self.graphique_chronic, bg="#FF9800", fg="white").pack(pady=10)
        tk.Button(self.frame_chronic, text="Excel تصدير إلى", command=self.exporter_chronic, bg="#2196F3", fg="white").pack(pady=5)

    @staticmethod
    def _ligne_chargement(parent, texte, commande, statut):
        ligne = tk.Frame(parent, bg="white")
        ligne.pack(pady=5, fill="x")
        tk.Button(ligne, text=texte, command=commande, bg="#f0f0f0").pack(side="left")
        tk.Label(ligne, textvariable=statut, bg="white", fg="green", font=(config.POLICE, 12)).pack(side="right")

    def _boutons_nachate(self):
        tk.Button(self.root, text="تتبع النشاط العام للمحكمة", command=self.fenetre_nachate,
                  bg="lightblue", font=(config.POLICE, 12)).pack(pady=5)
        tk.Button(self.root, text="+ إضافة ملف نشاط", command=self.ajouter_nachate,
                  bg="lightgreen", font=(config.POLICE, 12)).pack(pady=5)

    # ------------------------------------------------------- affichage / nav
    def basculer_section(self, cadre):
        if not self.etat_sections[cadre]:
            for autre in self.etat_sections:
                autre.pack_forget()
                self.etat_sections[autre] = False
            cadre.pack(pady=10)
            self.etat_sections[cadre] = True
        else:
            cadre.pack_forget()
            self.etat_sections[cadre] = False

    def basculer_zajri(self):
        if not self.zajri_visible:
            self.frame_zajri_buttons.pack(pady=10)
            self.frame_buttons.pack(pady=10)  # إحصائيات عامة ouvert par défaut
            for cadre in (self.frame_ahkam, self.frame_chronic):
                cadre.pack_forget()
                self.etat_sections[cadre] = False
            self.etat_sections[self.frame_buttons] = True
            self.frame_madani_buttons.pack_forget()
            self.zajri_visible, self.madani_visible = True, False
        else:
            self.frame_zajri_buttons.pack_forget()
            self.frame_buttons.pack_forget()
            self.zajri_visible = False

    def basculer_madani(self):
        if not self.madani_visible:
            self.frame_madani_buttons.pack(pady=10)
            self.frame_zajri_buttons.pack_forget()
            self.frame_buttons.pack_forget()
            self.zajri_visible, self.madani_visible = False, True
        else:
            self.frame_madani_buttons.pack_forget()
            self.madani_visible = False

    def afficher_tableau(self, df):
        """Affiche un DataFrame dans un tableau défilant (horizontalement et verticalement)."""
        for widget in self.frame_result.winfo_children():
            widget.destroy()
        if df.empty:
            messagebox.showinfo("Info", "Aucune donnée à afficher.")
            return
        if str(df.columns[0]).lower() in ["index", "unnamed: 0"]:
            df = df.rename(columns={df.columns[0]: "السنة"})

        conteneur = tk.Frame(self.frame_result)
        conteneur.pack(fill="both", expand=True)
        barre_y = tk.Scrollbar(conteneur, orient="vertical")
        barre_x = tk.Scrollbar(conteneur, orient="horizontal")
        barre_y.pack(side="right", fill="y")
        barre_x.pack(side="bottom", fill="x")
        canvas = tk.Canvas(conteneur, xscrollcommand=barre_x.set, yscrollcommand=barre_y.set)
        canvas.pack(side="left", fill="both", expand=True)
        barre_x.config(command=canvas.xview)
        barre_y.config(command=canvas.yview)

        interieur = tk.Frame(canvas)
        canvas.create_window((0, 0), window=interieur, anchor="nw")
        interieur.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        colonnes = [str(c) for c in df.columns]
        arbre = ttk.Treeview(interieur, columns=colonnes, show="headings")
        arbre.pack(fill="both", expand=True)
        for col in colonnes:
            arbre.heading(col, text=col)
            arbre.column(col, anchor="center", width=150, minwidth=100)
        for ligne in df.itertuples(index=False):
            arbre.insert("", "end", values=[str(v) if pd.notna(v) else "" for v in ligne])

        for widget in (arbre, canvas):
            widget.bind("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1 * (e.delta / 120)), "units"))
            widget.bind("<Shift-MouseWheel>", lambda e: canvas.xview_scroll(int(-1 * (e.delta / 120)), "units"))
        arbre.update_idletasks()
        canvas.configure(scrollregion=canvas.bbox("all"))
        canvas.xview_moveto(0)
        canvas.yview_moveto(0)

    def fenetre_graphique(self, titre, figure, taille="1000x600"):
        fenetre = tk.Toplevel(self.root)
        fenetre.title(titre)
        fenetre.geometry(taille)
        canvas = FigureCanvasTkAgg(figure, master=fenetre)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
        return fenetre

    def exporter(self, feuilles, nom_base, figure=None, nom_fichier=None):
        """Demande où enregistrer, puis écrit le classeur Excel."""
        if all(df.empty for _, df in feuilles):
            messagebox.showerror("Erreur", "لا توجد بيانات لتصديرها.")
            return
        dossier = os.path.expanduser("~")
        nom = nom_fichier or os.path.basename(generer_nom_fichier(nom_base, dossier))
        chemin = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=nom, initialdir=dossier,
                                              filetypes=[("Excel files", "*.xlsx")])
        if not chemin:
            return
        try:
            enregistrer_classeur(chemin, feuilles, figure)
            messagebox.showinfo("تم الحفظ", f"تم تصدير الملف إلى:\n{chemin}")
        except Exception as e:
            messagebox.showerror("Erreur export Excel", str(e))
        finally:
            if figure is not None:
                plt.close(figure)

    # ------------------------------------------------------- إحصائيات عامة
    def charger_fichier_stats(self, nom):
        chemin = filedialog.askopenfilename(filetypes=FILTRE_EXCEL_CSV)
        if not chemin:
            return
        try:
            self.fichiers_stats[nom] = tr.preparer_fichier_stats(tr.lire_tableau(chemin))
            self.statuts_stats[nom].set("✅")
        except Exception as e:
            self.statuts_stats[nom].set("❌ خطأ")
            messagebox.showerror("Erreur", str(e))

    def _trois_fichiers(self):
        if any(df is None for df in self.fichiers_stats.values()):
            messagebox.showerror("Erreur", "Veuillez charger les 3 fichiers avant de continuer.")
            return None
        return self.fichiers_stats["المخلف"], self.fichiers_stats["المسجل"], self.fichiers_stats["المحكوم"]

    def traiter_stats(self):
        fichiers = self._trois_fichiers()
        if fichiers is None:
            return
        try:
            self.df_final = tr.stats_generales(*fichiers)
            self.afficher_tableau(self.df_final)
        except Exception as e:
            messagebox.showerror("Erreur traitement", str(e))

    def afficher_par_hay2a(self):
        fichiers = self._trois_fichiers()
        if fichiers is None:
            return
        try:
            self.df_par_hay2a = tr.stats_par_hay2a(*fichiers)
            self.afficher_tableau(self.df_par_hay2a)
        except Exception as e:
            messagebox.showerror("Erreur هيئة", str(e))

    def exporter_stats(self):
        self.exporter([("البيانات العامة", self.df_final), ("حسب الهيئة", self.df_par_hay2a)], "النتائج")

    # ------------------------------------------------- انتاجيات القضاة (الزجري)
    def charger_ahkam(self):
        chemin = filedialog.askopenfilename(filetypes=FILTRE_EXCEL_CSV)
        if not chemin:
            return
        try:
            self.df_ahkam = tr.preparer_ahkam(tr.lire_tableau(chemin))
            self.statut_ahkam.set("✅")
        except Exception as e:
            self.statut_ahkam.set("❌ خطأ")
            messagebox.showerror("Erreur ahkam", str(e))

    def afficher_ahkam(self):
        if self.df_ahkam.empty:
            messagebox.showinfo("Info", "Veuillez d'abord تحميل سجل الأحكام.")
            return
        if not tr.colonnes_presentes(self.df_ahkam, [config.COL_JUGE_ZAJRI, "الرمز", "الشعبة"]):
            messagebox.showerror("Erreur", "Les colonnes nécessaires (القاضي المقرر، الرمز، الشعبة) sont manquantes.")
            return
        try:
            pivot = tr.pivot_juges(self.df_ahkam, config.COL_JUGE_ZAJRI)
            pivot = pivot.rename(columns={pivot.columns[0]: "الشعب"})
            self.pivot_ahkam = pivot
            juges = [c for c in pivot.columns if c not in ["المجموع", "الشعب"]]
            self.afficher_tableau(pivot[["المجموع"] + juges + ["الشعب"]])
        except Exception as e:
            messagebox.showerror("Erreur affichage", str(e))

    def graphique_ahkam(self):
        if self.df_ahkam.empty:
            messagebox.showinfo("Info", "المرجو تحميل سجل الأحكام أولاً.")
            return
        self.fenetre_graphique("الرسوم البيانية - انتاجيات القضاة",
                               graphiques.figure_juges(self.df_ahkam, config.COL_JUGE_ZAJRI))

    def exporter_ahkam(self):
        if self.pivot_ahkam.empty:
            messagebox.showerror("Erreur", "لا توجد بيانات لتصديرها.")
            return
        figure = graphiques.figure_juges(self.df_ahkam, config.COL_JUGE_ZAJRI)
        self.exporter([("الإنتاجيات", self.pivot_ahkam)], "انتاجيات_القضاة_مع_الرسم", figure=figure,
                      nom_fichier="انتاجيات_القضاة_مع_الرسم.xlsx")

    # ----------------------------------------------------- الملفات المزمنة
    def charger_chronic(self):
        chemin = filedialog.askopenfilename(filetypes=FILTRE_EXCEL_CSV)
        if not chemin:
            return
        try:
            self.df_chronic = tr.preparer_chronic(tr.lire_tableau(chemin))
            self.statut_chronic.set("✅")
        except Exception as e:
            self.statut_chronic.set("❌ خطأ")
            messagebox.showerror("Erreur fichier المزمنة", str(e))

    def afficher_chronic(self):
        if self.df_chronic.empty:
            messagebox.showinfo("Info", "Veuillez d'abord تحميل الملفات المزمنة.")
            return
        try:
            self.pivot_chronic = tr.pivot_chronic(self.df_chronic)
            self.afficher_tableau(self.pivot_chronic)
        except Exception as e:
            messagebox.showerror("Erreur chronique", str(e))

    def graphique_chronic(self):
        if self.df_chronic.empty:
            messagebox.showinfo("Info", "المرجو تحميل الملفات المزمنة أولاً.")
            return
        self.fenetre_graphique("الرسوم البيانية - الملفات المزمنة",
                               graphiques.figure_chronic(tr.chronic_par_annee(self.df_chronic)))

    def exporter_chronic(self):
        if self.pivot_chronic.empty:
            messagebox.showerror("Erreur", "لا توجد بيانات لتصديرها.")
            return
        figure = graphiques.figure_chronic(tr.chronic_par_annee(self.df_chronic))
        self.exporter([("الملفات المزمنة", self.pivot_chronic)], "الملفات_المزمنة_مع_الرسم", figure=figure,
                      nom_fichier="الملفات_المزمنة_مع_الرسم.xlsx")

    # ------------------------------------------- تتبع النشاط العام للمحكمة
    def ajouter_nachate(self):
        chemin = filedialog.askopenfilename(title="اختر ملف Excel", filetypes=FILTRE_EXCEL)
        if not chemin:
            return
        try:
            self.fichiers_nachate.append(tr.preparer_nachate(pd.read_excel(chemin)))
            messagebox.showinfo("نجاح", "تمت إضافة الملف بنجاح!")
        except Exception as e:
            messagebox.showerror("خطأ", f"حدث خطأ أثناء قراءة الملف: {e}")

    def fenetre_nachate(self):
        if not self.fichiers_nachate:
            messagebox.showinfo("معلومة", "المرجو إضافة ملفات النشاط أولاً.")
            return
        fenetre = tk.Toplevel(self.root)
        fenetre.title("تتبع النشاط العام للمحكمة")
        periode = tk.Frame(fenetre)
        periode.pack(pady=10)
        tk.Label(periode, text="من:").grid(row=0, column=0)
        debut = tk.Entry(periode, width=10)
        debut.grid(row=0, column=1)
        debut.insert(0, "01-2024")
        tk.Label(periode, text="إلى:").grid(row=0, column=2)
        fin = tk.Entry(periode, width=10)
        fin.grid(row=0, column=3)
        fin.insert(0, "12-2024")

        zone = {"canvas": None}

        def generer():
            try:
                par_mois = tr.activite_mensuelle(self.fichiers_nachate, debut.get(), fin.get())
                if zone["canvas"] is not None:  # remplace le graphique précédent au lieu d'en empiler un autre
                    zone["canvas"].get_tk_widget().destroy()
                canvas = FigureCanvasTkAgg(graphiques.figure_activite(par_mois), master=fenetre)
                canvas.draw()
                canvas.get_tk_widget().pack()
                zone["canvas"] = canvas
            except Exception as e:
                messagebox.showerror("خطأ", f"تعذر إنشاء الرسم البياني: {e}")

        tk.Button(fenetre, text="عرض الرسم", command=generer).pack(pady=5)


def lancer():
    root = tk.Tk()
    Application(root)
    root.mainloop()
