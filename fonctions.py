import pandas as pd
import numpy as np


def read_table(file_obj):
    """
    Lit un fichier CSV ou Excel à partir d’un chemin ou d’un fichier Streamlit.
    Args:
        file_obj (str ou UploadedFile): Le chemin du fichier ou l'objet UploadedFile.
    Returns:
        pd.DataFrame: Le DataFrame contenant les données.
    """

    # Si c’est un chemin (str)
    if isinstance(file_obj, str):
        if file_obj.endswith(".csv"):
            return pd.read_csv(file_obj, sep=';', encoding='utf-8')
        elif file_obj.endswith(".xlsx"):
            return pd.read_excel(file_obj)
        else:
            raise ValueError("Le fichier doit être au format .csv ou .xlsx")

    # Si c’est un fichier uploadé (Streamlit UploadedFile)
    elif hasattr(file_obj, "name"):
        filename = file_obj.name
        if filename.endswith(".csv"):
            return pd.read_csv(file_obj, sep=';', encoding='utf-8')
        elif filename.endswith(".xlsx"):
            return pd.read_excel(file_obj)
        else:
            raise ValueError("Le fichier doit être au format .csv ou .xlsx")

    else:
        raise TypeError("Entrée non reconnue : chemin ou fichier Streamlit attendu.")


def nettoyer_et_convertir(col):
    """
    Nettoie une colonne de données en supprimant les espaces (y compris insécables),
    remplaçant les virgules par des points, et convertissant en float.
    Args:
        col (pd.Series): La colonne à nettoyer et convertir.
    Returns:
        pd.Series: La colonne nettoyée et convertie en float.
    """
    return (
        col.astype(str)
        .str.replace("\xa0", "", regex=False)  # supprime les espaces insécables
        .str.replace(" ", "")                  # supprime les espaces classiques
        .str.replace(",", ".", regex=False)    # remplace virgule par point
        .replace("", "0")                      # chaîne vide => 0
        .astype(float)
    )


def preparer_donnees(stocks, ventes):
    """
    Prépare les données pour le calcul du taux de rotation.
    - Nettoie les colonnes numériques
    - Effectue une jointure gauche
    - Remplace les NA
    - Calcule le taux de rotation
    - Sélectionne et renomme les colonnes pertinentes

    Args:
        stocks (pd.DataFrame): Données de stock
        ventes (pd.DataFrame): Données de ventes

    Returns:
        pd.DataFrame: Données finales prêtes à être exportées
    """
    # Nettoyage des colonnes numériques
    ventes["Qté Vendues"] = nettoyer_et_convertir(ventes["Qté Vendues"])
    ventes["Chiffre d'affaires HT"] = nettoyer_et_convertir(
        ventes["Chiffre d'affaires HT"]
    )
    stocks["Qté Stock Réel"] = nettoyer_et_convertir(stocks["Qté Stock Réel"])

    # LEFT MERGE sur "Référence Article"
    df = pd.merge(stocks, ventes, on="Référence Article", how="left")

    # Remplacement des valeurs manquantes
    df["Qté Vendues"] = df["Qté Vendues"].fillna(0)
    df["Chiffre d'affaires HT"] = df["Chiffre d'affaires HT"].fillna(0)

    # Calcul du taux de rotation
    df["Taux de rotation"] = df["Qté Vendues"] / df["Qté Stock Réel"]
    df["Taux de rotation"] = df["Taux de rotation"].replace(
        [float("inf"), -float("inf")], np.nan
    )

    # Colonnes à conserver
    colonnes_selection = [
        "Référence Article",
        "Désignation Article_x",
        "Qté Vendues",
        "Qté Stock Réel",
        "Chiffre d'affaires HT",
        "Taux de rotation"
    ]
    df_final = df[colonnes_selection].copy()

    # Renommage des colonnes
    df_final.rename(
        columns={"Désignation Article_x": "Désignation Article"},
        inplace=True
    )

    return df_final


def preparer_donnees2(stocks, mouvements):
    # Calculer Qté Sortie négative par article
    quant_neg = mouvements[mouvements["Quantité"] < 0].copy()
    quant_neg["Référence Article"] = (
        quant_neg["Référence Article"].astype(str).str.strip()
    )
    stocks["Référence Article"] = stocks["Référence Article"].astype(str).str.strip()

    somme_neg = (
        quant_neg.groupby("Référence Article")
        .agg({
            "Quantité": "sum",
            "Désignation Article": "first",
            "Code - Intitulé Famille": "first"
        })
        .rename(columns={"Quantité": "Qté Sortie"})
        .reset_index()
        )

    # Qté Sortie positive
    somme_neg["Qté Sortie"] = somme_neg["Qté Sortie"].abs()

    # Fusionner directement avec stocks
    df = pd.merge(stocks, somme_neg, on="Référence Article", how="outer")

    # Pour la colonne Désignation Article
    df["Désignation Article"] = df["Désignation Article_x"].fillna(
        df["Désignation Article_y"]
    )

    # Pour la colonne Code - Intitulé Famille
    df["Code - Intitulé Famille"] = (
        df["Code - Intitulé Famille_x"].fillna(df["Code - Intitulé Famille_y"])
    )

    # Ensuite tu peux supprimer les colonnes originales
    df.drop(columns=[
        "Désignation Article_x", "Désignation Article_y",
        "Code - Intitulé Famille_x", "Code - Intitulé Famille_y"
    ], inplace=True)

    # Remplacer NaN par 0 pour les articles sans sortie enregistrée
    df["Qté Sortie"] = df["Qté Sortie"].fillna(0)

    df["Taux de rotation"] = (
        df["Qté Sortie"] / df["Qté Stock Réel"]
    )
    df["Taux de rotation"].replace([np.inf, -np.inf], np.nan, inplace=True)

    return df


def groupby_famille(df_article):
    df_grouped = (
        df_article.groupby("Code - Intitulé Famille")[
            ["Taux de rotation", "Qté Stock Réel", "Qté Sortie"]
        ]
        .mean()
        .reset_index()
    )

    df_grouped.rename(columns={
        "Code - Intitulé Famille": "Famille",
        "Taux de rotation": "Taux rotation moyen",
        "Qté Stock Réel": "Qté Stock Réel moyen",
        "Qté Sortie": "Qté Sortie moyenne"
    }, inplace=True)

    return df_grouped
