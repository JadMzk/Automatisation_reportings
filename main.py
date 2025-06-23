import streamlit as st
import pandas as pd
import io
from fonctions import read_table, preparer_donnees

st.set_page_config(page_title="Taux de rotation", layout="wide")
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

    - `Référence article`
    - `Désignation article`
    - `Code - Intitulé famille`
    - `Qté stock réel`

    > Une extraction Sage permet de récupérer ces colonnes ainsi que d'autres,
    > mais **seules celles-ci sont nécessaires**.

    ---

    #### 🛒 Le fichier de **ventes** doit contenir les colonnes suivantes :

    - `Référence article`
    - `Désignation article`
    - `Code - Intitulé famille`
    - `Qté vendues`
    - `Chiffre d'affaires HT`

    > Une extraction Sage (ventes par article) permet de récupérer ces colonnes ainsi que d'autres,
    > mais **seules celles-ci sont nécessaires**.
    """
)


col1, col2 = st.columns(2)
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

        buffer = io.BytesIO()
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
