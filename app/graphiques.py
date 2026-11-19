"""Graphiques matplotlib avec un texte arabe correctement affiché."""

import arabic_reshaper
import matplotlib.pyplot as plt
from bidi.algorithm import get_display

plt.rcParams["axes.unicode_minus"] = False


def ar(texte):
    """Met en forme un texte arabe pour matplotlib (liaison des lettres et sens de lecture)."""
    return get_display(arabic_reshaper.reshape(str(texte)))


def _valeurs_sur_barres(ax, decalage=1, taille=8):
    for groupe in ax.containers:
        for barre in groupe:
            hauteur = barre.get_height()
            if hauteur > 0:
                ax.text(barre.get_x() + barre.get_width() / 2, hauteur + decalage, str(int(hauteur)),
                        ha="center", va="bottom", fontsize=taille)


def figure_juges(df_ahkam, col_juge):
    """Nombre d'affaires jugées par chaque juge (barres)."""
    comptes = df_ahkam[col_juge].value_counts()
    comptes.index = [ar(nom) for nom in comptes.index]
    fig, ax = plt.subplots(figsize=(8, 5))
    comptes.plot(kind="bar", ax=ax, color="skyblue")
    _valeurs_sur_barres(ax, taille=10)
    ax.set_title(ar("عدد القضايا التي يحكمها كل قاضي"))
    ax.set_xlabel(ar("القاضي المقرر"))
    ax.set_ylabel(ar("عدد القضايا"))
    ax.tick_params(axis="x", rotation=45)
    fig.tight_layout()
    return fig


def figure_chronic(par_annee):
    """Dossiers restants par année et par شعبة (barres groupées)."""
    donnees = par_annee.copy()
    donnees.columns = [ar(c) for c in donnees.columns]
    donnees.index = donnees.index.astype(str)
    fig, ax = plt.subplots(figsize=(12, 6))
    donnees.plot(kind="bar", stacked=False, ax=ax, width=1)
    _valeurs_sur_barres(ax)
    ax.set_title(ar("عدد الملفات المتبقية في كل شعبة و في كل سنة"))
    ax.set_xlabel(ar("السنة"))
    ax.set_ylabel(ar("عدد الملفات المتبقية"))
    ax.tick_params(axis="x", rotation=0)
    fig.tight_layout()
    return fig


def figure_activite(par_mois):
    """Évolution mensuelle de المحكوم, المخلف et المسجل (courbes)."""
    mois = [ar(m) for m in par_mois.index]
    fig, ax = plt.subplots(figsize=(10, 5))
    for col in ["المحكوم", "المخلف", "المسجل"]:
        ax.plot(mois, par_mois[col], label=ar(col), marker="o")
    ax.legend()
    ax.set_title(ar("تتبع النشاط العام للمحكمة"))
    ax.set_ylabel(ar("العدد"))
    ax.set_xlabel(ar("الأشهر"))
    ax.tick_params(axis="x", rotation=45)
    fig.tight_layout()
    return fig
