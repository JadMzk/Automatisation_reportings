import streamlit as st
import pandas as pd
import io
from fonctions import read_table, preparer_donnees

# Initialisation de l'app
st.set_page_config(page_title="Automatisation des reporting", layout="wide")

# Onglets Streamlit
tab1, tab2 = st.tabs(["📦 Taux de rotation", "📋 Bons de livraison"])
with tab1:
    # Initialisation de l'application Streamlit
    st.title("📦 Calcul automatique du taux de rotation")

    # Description de l'application
    st.markdown(
        """
        ### 📘 Description de l'application

        Cette application permet de calculer le **taux de rotation** des articles à partir de deux fichiers :

        - Un fichier **de stocks** (à gauche)
        - Un fichier **de ventes** (à droite)

        ⚠️ Merci de **sélectionner les fichiers dans cet ordre** pour garantir un bon traitement des données.

        ---

        #### 📦 Le fichier de **stocks** doit contenir les colonnes suivantes :

        - `Référence Article`
        - `Désignation Article`
        - `Qté stock réel`

        > Une extraction Sage permet de récupérer ces colonnes ainsi que d'autres,
        > mais **seules celles-ci sont nécessaires**.

        ---

        #### 🛒 Le fichier de **ventes** doit contenir les colonnes suivantes :

        - `Référence Article`
        - `Désignation Article`
        - `Qté vendues`
        - `Chiffre d'affaires HT`

        > Une extraction Sage (ventes par article) permet de récupérer ces colonnes ainsi que d'autres,
        > mais **seules celles-ci sont nécessaires**.

        ⚠️ Les deux tables doivent concerner la **même période** pour un résultat cohérent.
        """
    )

    col1, col2 = st.columns(2)  # Affichage de l'interface en deux colonnes
    with col1:
        fichier_stocks = st.file_uploader(
            "📦 Fichier des **stocks** (.csv ou .xlsx)", type=["csv", "xlsx"], key="stocks"
        )
    with col2:
        fichier_ventes = st.file_uploader(
            "🛒 Fichier des **ventes** (.csv ou .xlsx)", type=["csv", "xlsx"], key="ventes"
        )

    if fichier_ventes and fichier_stocks:
        try:
            ventes = read_table(fichier_ventes)
            stocks = read_table(fichier_stocks)

            df_final = preparer_donnees(stocks, ventes)

            st.subheader("📊 Résultat")
            st.dataframe(df_final)

            buffer = io.BytesIO()  # Création d'un buffer pour l'export Excel
            # (sans sauvegarde sur disque mais en RAM)
            with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                df_final.to_excel(writer, index=False, sheet_name="Taux de rotation")

            st.download_button(
                label="📥 Télécharger le fichier Excel",
                data=buffer,
                file_name="résultat_taux_rotation.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Erreur pendant le traitement : {e}")


with tab2:
    st.header("Fusion des bons de livraison")

    st.markdown("""
    Téléversez deux fichiers Excel contenant les bons de livraison, dans l’ordre :
    - 📄 `etat_blst.xlsx` (mois actuel)
    - 📄 `etat_blstm1.xlsx` (mois précédent)
    """)

    col1, col2 = st.columns(2)

    with col1:
        fichier_bls_t = st.file_uploader("📥 Fichier actuel (bls_t)",
                                         type=["xlsx"], key="bls_t")
    with col2:
        fichier_bls_tm1 = st.file_uploader("📥 Fichier précédent (bls_tm1)",
                                           type=["xlsx"], key="bls_tm1")

    if fichier_bls_t and fichier_bls_tm1:
        try:
            # Lecture avec skiprows=2
            bls_t = pd.read_excel(fichier_bls_t, sheet_name="Feuil2", skiprows=2)
            bls_tm1 = pd.read_excel(fichier_bls_tm1, sheet_name="Feuil2", skiprows=2)

            # Nettoyage
            bls_t.columns.values[-1] = "REMARQUES_t"
            bls_tm1.columns.values[-1] = "REMARQUES_t_1"

            bls_t.columns = bls_t.columns.str.strip()
            bls_tm1.columns = bls_tm1.columns.str.strip()

            # Fusion
            df_bls = pd.merge(
                bls_t,
                bls_tm1[["Référence", "REMARQUES_t_1"]],
                on="Référence",
                how="left"
            )

            st.success("✅ Fusion réussie ! Aperçu ci-dessous :")
            st.dataframe(df_bls)

            # Export Excel
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                df_bls.to_excel(writer, index=False, sheet_name="Fusion BL")
            buffer.seek(0)

            st.download_button(
                label="📥 Télécharger la fusion Excel",
                data=buffer,
                file_name="fusion_bons_livraison.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Erreur pendant la fusion : {e}")
