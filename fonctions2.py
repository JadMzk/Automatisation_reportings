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
    return s.upper()


def traiter_fichiers(ancien_path, nouveau_path):
    try:
        df_ancien = lire_avec_header_auto(ancien_path)
        df_nouveau = lire_avec_header_auto(nouveau_path)

        df_ancien.columns = df_ancien.columns.str.strip()
        df_nouveau.columns = df_nouveau.columns.str.strip()

        if "Remarques" not in df_ancien.columns:
            raise ValueError(
                "La colonne 'Remarques' est absente du fichier ancien"
            )
        if "Remarques" not in df_nouveau.columns:
            raise ValueError(
                "La colonne 'Remarques' est absente du fichier nouveau"
            )

        # Création de la clé pour la jointure dans les deux df avec nettoyage
        df_ancien['key'] = df_ancien.apply(
            lambda row: nettoyer_chaine(row['Référence']) + "__" +
                        nettoyer_chaine(row['Désignation']),
            axis=1
        )
        df_nouveau['key'] = df_nouveau.apply(
            lambda row: nettoyer_chaine(row['Référence']) + "__" +
                        nettoyer_chaine(row['Désignation']),
            axis=1
        )

        # Mapping clé -> remarques ancien
        remarques_map = df_ancien.set_index('key')['Remarques'].to_dict()

        # Ajout colonne 'Remarques_ancien' dans df_nouveau
        df_nouveau['Remarques_ancien'] = df_nouveau['key'].map(remarques_map)

        # Suppression clé temporaire
        df_nouveau.drop(columns=['key'], inplace=True)

        return df_nouveau

    except Exception as e:
        raise Exception(f"Erreur dans traitement fichiers: {e}")
