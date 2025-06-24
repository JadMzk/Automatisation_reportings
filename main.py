import streamlit as st
import pandas as pd
import io
from fonctions import read_table, preparer_donnees
import matplotlib.pyplot as plt
import seaborn as sns
import fonctions2 as f2

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

        ⚠️ Les deux tables doivent concerner la **même période** (idéalement 1 an pour un résultat cohérent.
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
    ⚠️ Assurez-vous que les deux fichiers n'ont pas de ligne vide en haut, ou de lignes
    qui ne contiennent qu'un titre et que les colonnes sont correctement nommées.
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
            bls_t = pd.read_excel(fichier_bls_t, sheet_name="Feuil2")
            bls_tm1 = pd.read_excel(fichier_bls_tm1, sheet_name="Feuil2")

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
    st.header("🧾 Suivi des commandes")

    st.markdown("""
    Cette fonctionnalité vous permet de **mettre à jour automatiquement les remarques**
    d’un fichier de commandes en comparant avec un fichier précédent.

    - 📄 `Fichier ancien` : contient les anciennes remarques
    - 📄 `Fichier nouveau` : contient les nouvelles remarques
    """)

    col1, col2 = st.columns(2)

    with col1:
        ancien_fichier = st.file_uploader("📥 Fichier ancien",
                                          type=["xlsx"],
                                          key="ancien_suivi")
    with col2:
        nouveau_fichier = st.file_uploader("📥 Fichier nouveau",
                                           type=["xlsx"],
                                           key="nouveau_suivi")

    if ancien_fichier and nouveau_fichier:
        try:

            # Traitement
            df_resultat = f2.traiter_fichiers(ancien_fichier, nouveau_fichier)

            st.success("✅ Mise à jour effectuée avec succès !")
            st.dataframe(df_resultat)

            # Export Excel
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                df_resultat.to_excel(writer, index=False, sheet_name="Suivi commandé")
            buffer.seek(0)

            st.download_button(
                label="📥 Télécharger le fichier mis à jour",
                data=buffer,
                file_name="suivi_commandes_mis_a_jour.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Erreur pendant le traitement : {e}")
