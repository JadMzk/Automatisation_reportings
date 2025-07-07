import pandas as pd
import re


def trouver_ligne_header(filepath, mot_clef="N° Pièce", sheet_name=0):
    # On lit les 10 premières lignes du fichier pour chercher le header
    preview = pd.read_excel(
        filepath, sheet_name=sheet_name, nrows=10, header=None
    )
    for i, row in preview.iterrows():
        if row.astype(str).str.contains(mot_clef).any():
            return i
    raise ValueError(f"Impossible de trouver une ligne contenant '{mot_clef}'")


def lire_avec_header_auto(filepath, sheet_name=0, mot_clef="N° Pièce"):
    """    Lit un fichier Excel et détermine automatiquement la ligne d'en-tête
    en cherchant un mot clé spécifique dans les premières lignes."""
    header_line = trouver_ligne_header(
        filepath, mot_clef=mot_clef, sheet_name=sheet_name
    )
    df = pd.read_excel(filepath, sheet_name=sheet_name, skiprows=header_line)
    df.columns = df.columns.str.strip()
    return df


def nettoyer_chaine(s):
    if pd.isna(s):
        return ''
    s = str(s)
    s = re.sub(r'[\s\u00A0]+', '', s)  # supprime espaces classiques et insécables
    # Plus robuste qu'un simple strip()
    return s.upper()


def traiter_fichiers(ancien_path, nouveau_path):
    """
    Transfère les remarques de l'ancien fichier vers le nouveau,
    en utilisant une clé basée sur Référence + Désignation.
    """
    try:
        df_ancien = lire_avec_header_auto(ancien_path)
        df_nouveau = lire_avec_header_auto(nouveau_path)

        # Nettoyage des noms de colonnes
        df_ancien.columns = df_ancien.columns.str.strip()
        df_nouveau.columns = df_nouveau.columns.str.strip()

        # Vérifie que 'Remarques' est bien dans l'ancien fichier
        if "Remarques" not in df_ancien.columns:
            raise ValueError("La colonne 'Remarques' est absente du fichier ancien")

        # Création des clés de jointure
        df_ancien['key'] = df_ancien.apply(
            lambda row: nettoyer_chaine(str(row['Référence'])) + "__" +
                        nettoyer_chaine(str(row['Désignation'])),
            axis=1
        )
        df_nouveau['key'] = df_nouveau.apply(
            lambda row: nettoyer_chaine(str(row['Référence'])) + "__" +
                        nettoyer_chaine(str(row['Désignation'])),
            axis=1
        )

        # Dictionnaire : clé -> remarque
        remarques_map = df_ancien.set_index('key')["Remarques"].to_dict()

        # Ajout des remarques dans le fichier nouveau
        df_nouveau["Remarques_ancien"] = df_nouveau["key"].map(remarques_map)

        # Suppression de la clé temporaire
        df_nouveau.drop(columns=["key"], inplace=True)

        return df_nouveau

    except Exception as e:
        raise Exception(f"Erreur dans traitement fichiers : {e}")

def nettoyer_num_piece(df):
    """    Nettoie la colonne "N° Pièce" d'un DataFrame."""
    df["N° Pièce"] = (
        df["N° Pièce"]
        .astype(str)
        .str.strip()
        .str.upper()
        .str.replace("\xa0", "", regex=False)  # supprime les espaces insécables
    )
    return df
