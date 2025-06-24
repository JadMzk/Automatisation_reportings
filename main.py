import streamlit as st
import pandas as pd
import io
from fonctions import read_table, preparer_donnees
import matplotlib.pyplot as plt
import seaborn as sns
from fonctions2 import traiter_fichiers
import os

# Initialisation de l'app
st.set_page_config(page_title="Automatisation des reporting", layout="wide")

# Onglets Streamlit
tab1, tab2, tab3 = st.tabs(["📦 Taux de rotation", "📋 Bons de livraison",
                            "🧾 Suivi de commandes"])
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

            # ✅ Création de graphs pour avoir des insights pertinents
            col1, col2 = st.columns(2)

            with col1:
                st.subheader("📊 Top 10 par chiffre d'affaires")

                top_CA = df_final.sort_values("Chiffre d'affaires HT",
                                              ascending=False).head(10)

                fig1, ax1 = plt.subplots(figsize=(8, 6))
                sns.barplot(
                    data=top_CA,
                    x="Taux de rotation",
                    y="Désignation Article",
                    palette="tab10",
                    ax=ax1
                )
                ax1.set_title("Top 10 produits par CA - Taux de rotation en abscisse")
                ax1.set_xlabel("Taux de rotation")
                ax1.set_ylabel("Désignation Article")
                ax1.grid(True)
                st.pyplot(fig1)

            with col2:
                st.subheader("📊 Top 10 par taux de rotation")

                top_rotation = df_final.sort_values("Taux de rotation",
                                                    ascending=False).head(10)

                fig2, ax2 = plt.subplots(figsize=(8, 6))
                sns.barplot(
                    data=top_rotation,
                    x="Chiffre d'affaires HT",
                    y="Désignation Article",
                    palette="tab10",
                    ax=ax2
                )
                ax2.set_title("Top 10 par taux de rotation - CA en abscisse")
                ax2.set_xlabel("Chiffre d'affaires HT en MAD")
                ax2.set_ylabel("Désignation Article")
                ax2.grid(True)
                st.pyplot(fig2)

            # ✅ Export Excel
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                df_final.to_excel(writer, index=False, sheet_name="Taux de rotation")
            buffer.seek(0)

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
    - 📄 `Dernier bon de livraison avec remarques`
    - 📄 `Avant dernier bon de livraison avec remarques`
    """)

    col1, col2 = st.columns(2)

    with col1:
        fichier_bls_t = st.file_uploader("📥 Dernier fichier avec remarques",
                                         type=["xlsx"], key="BLS_de_ce_mois")
    with col2:
        fichier_bls_tm1 = st.file_uploader("📥 Avant dernier fichier avec remarques",
                                           type=["xlsx"],
                                           key="BLS_t_m1")
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
                on="Référence"
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

with tab3:
    st.header("🧾 Suivi de commandes Excel")

    st.markdown("""
    Cette fonctionnalité permet de **reporter automatiquement les remarques** d’un ancien fichier Excel de bons de livraison vers un nouveau fichier.

    👉 Assurez-vous que les colonnes "N° Pièce", "Référence" et "Remarques" sont bien présentes.
    """)

    ancien = st.file_uploader("📤 Charger l'ancien fichier", type=["xlsx"], key="ancien_file")
    nouveau = st.file_uploader("📤 Charger le nouveau fichier", type=["xlsx"], key="nouveau_file")

    if ancien and nouveau:
        if st.button("⚙️ Lancer la mise à jour"):
            with st.spinner("Traitement en cours..."):
                try:
                    # Sauvegarde temporaire des fichiers uploadés
                    temp_ancien = os.path.join(os.getcwd(), "ancien_temp.xlsx")
                    temp_nouveau = os.path.join(os.getcwd(), "nouveau_temp.xlsx")

                    with open(temp_ancien, "wb") as f:
                        f.write(ancien.getbuffer())

                    with open(temp_nouveau, "wb") as f:
                        f.write(nouveau.getbuffer())

                    # Appel de la fonction principale
                    path_result = traiter_fichiers(temp_ancien, temp_nouveau)

                    st.success("✅ Traitement terminé avec succès.")

                    # Télécharger le fichier final
                    with open(path_result, "rb") as f:
                        st.download_button(
                            label="📥 Télécharger le fichier final",
                            data=f,
                            file_name=os.path.basename(path_result),
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )

                except Exception as e:
                    st.error(f"❌ Erreur pendant le traitement : {e}")
