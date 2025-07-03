import streamlit as st
import pandas as pd
import io
from fonctions import preparer_donnees, preparer_donnees2, groupby_famille
import matplotlib.pyplot as plt
import seaborn as sns
import fonctions2 as f2

# Initialisation de l'app
st.set_page_config(page_title="Automatisation des reporting", layout="wide")

# Onglets Streamlit
tab1, tab2, tab3, tab4 = st.tabs([
    "📊 Taux de rotation avec vente", "📋 Bons de livraison",
    "🧾 Suivi de commandes", "📦 Taux de rotation avec mouvement"
])

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
        - `Qté Vendues`
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
            ventes = f2.lire_avec_header_auto(fichier_ventes, mot_clef="Référence Article")
            stocks = f2.lire_avec_header_auto(fichier_stocks, mot_clef="Référence Article")

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
    - 📄 `Fichier des bons de livraison avec nouvelles remarques`
    - 📄 `Fichier des bons de livraison avec anciennes remarques`

    ⚠️ Assurez-vous que les deux fichiers n'ont pas de ligne vide en haut, ou de lignes
    qui ne contiennent qu'un titre et que les colonnes sont correctement nommées.

    Assurez-vous également que la feuille de calcul est nommée
                `Feuil2` dans les deux fichiers.

    Les colonnes doivent contenir au moins :
    - `Référence` : la référence du bon de livraison
    - `REMARQUES` : les remarques des bons de livraison
    - `N° Compte Client` : le numéro de compte client (pour la fusion)

    Les colonnes nommés ainsi sont nécessaires pour que la fusion fonctionne correctement.
    Si les noms de colonnes sont différents, vous pouvez les renommer dans le fichier Excel
    """)

    col1, col2 = st.columns(2)

    with col1:
        fichier_bls_t = st.file_uploader("📥 Fichier avec nouvelles remarques",
                                         type=["xlsx"], key="BLS_de_ce_mois")
    with col2:
        fichier_bls_tm1 = st.file_uploader("📥 Fichier avec anciennes remarques",
                                           type=["xlsx"],
                                           key="BLS_t_m1")
    if fichier_bls_t and fichier_bls_tm1:
        try:
            bls_t = f2.lire_avec_header_auto(fichier_bls_t, sheet_name="Feuil2",
                                             mot_clef="N° Compte Client")
            bls_tm1 = f2.lire_avec_header_auto(fichier_bls_tm1, sheet_name="Feuil2",
                                               mot_clef="N° Compte Client")

            # Nettoyage
            bls_t.columns.values[-1] = "REMARQUES_nouvelles"
            bls_tm1.columns.values[-1] = "REMARQUES_anciennes"

            bls_t.columns = bls_t.columns.str.strip()
            bls_tm1.columns = bls_tm1.columns.str.strip()

            bls_t = f2.nettoyer_num_piece(bls_t)
            bls_tm1 = f2.nettoyer_num_piece(bls_tm1)

            bls_tm1_unique = (
                bls_tm1.groupby("N° Pièce", as_index=False)["REMARQUES_anciennes"]
                .first()
            )

            # Fusion sans doublons
            df_bls = pd.merge(
                bls_t,
                bls_tm1_unique[["N° Pièce", "REMARQUES_anciennes"]],
                on="N° Pièce", how="left"
            )
            # Extraire la colonne REMARQUES_nouvelles
            remarques_nouvelles = df_bls.pop("REMARQUES_nouvelles")

            # La remettre en dernière position
            df_bls["REMARQUES_nouvelles"] = remarques_nouvelles

            # Fusion avec doublons (version brute)
            df_bls2 = pd.merge(
                bls_t,
                bls_tm1[["N° Pièce", "REMARQUES_anciennes"]],
                on="N° Pièce", how="left"
            )

            df_bls.drop(columns = ["Prix Revient Total"], inplace=True, errors='ignore')

            st.success("✅ Fusion réussie ! Aperçu ci-dessous :")
            st.dataframe(df_bls)

            # Préparer les deux buffers
            buffer_sans_doublon = io.BytesIO()
            with pd.ExcelWriter(buffer_sans_doublon, engine="xlsxwriter") as writer:
                df_bls.to_excel(writer, index=False, sheet_name="Fusion BL")
            buffer_sans_doublon.seek(0)

            buffer_avec_doublon = io.BytesIO()
            with pd.ExcelWriter(buffer_avec_doublon, engine="xlsxwriter") as writer:
                df_bls2.to_excel(writer, index=False, sheet_name="Fusion BL (doublons)")
            buffer_avec_doublon.seek(0)

            # Deux colonnes pour deux boutons
            col_g, col_d = st.columns(2)

            with col_g:
                st.download_button(
                    label="📥 Télécharger la fusion SANS doublons",
                    data=buffer_sans_doublon,
                    file_name="fusion_bons_livraison_sans_doublon.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            with col_d:
                st.download_button(
                    label="📥 Télécharger la fusion AVEC doublons",
                    data=buffer_avec_doublon,
                    file_name="fusion_bons_livraison_avec_doublon.xlsx",
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

    Les colonnes nécessaires pour le traitement sont :
    - `Référence` : la référence de la commande
    - `Désignation` : la désignation de l'article commandé
    - `Remarques` : les remarques associées à l'article

    Assurez-vous que les colonnes sont correctement nommées dans les fichiers Excel.
    Si les noms de colonnes sont différents, vous pouvez les renommer dans le fichier Excel
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


with tab4:
    st.header("📊 Taux de rotation avec mouvement")

    st.markdown("""
    Cette fonctionnalité vous permet de **calculer le taux de rotation**
    en tenant compte des mouvements de stock.

    - 📄 `Fichier de stock` : contient les données de stock
    - 📄 `Fichier de mouvement` : contient les données de mouvement

    Les colonnes nécessaires pour le traitement sont :
    - `Référence Article` : la référence de l'article
    - `Qté Stock Réel` : la quantité de stock réel (dans le fichier de stock)
    - `Qté Mouvement` : la quantité de mouvement
    - `Quantité` : la quantité associée à chaque mouvement (dans le fichier de mouvement)
    - `Code - Intitulé Famille` : la famille de l'article (optionnel mais recommandé)

    Assurez-vous que les colonnes sont correctement nommées dans les fichiers Excel.
    Si les noms de colonnes sont différents, vous pouvez les renommer dans le fichier Excel
    """)

    col1, col2 = st.columns(2)

    with col1:
        fichier_stock = st.file_uploader(
            "📥 Fichier de stock",
            type=["xlsx"],
            key="fichier_stock"
        )
    with col2:
        fichier_mouvement = st.file_uploader(
            "📥 Fichier de mouvement",
            type=["xlsx"],
            key="fichier_mouvement"
        )

    if fichier_stock and fichier_mouvement:
        try:
            mouvements = f2.lire_avec_header_auto(
                fichier_mouvement,
                mot_clef="Référence Article"
            )
            stocks = f2.lire_avec_header_auto(
                fichier_stock,
                mot_clef="Référence Article"
            )

            df_resultat = preparer_donnees2(stocks, mouvements)
            st.success("✅ Calcul du taux de rotation effectué avec succès !")
            st.dataframe(df_resultat)

            df_famille = groupby_famille(df_resultat)

            # Préparer buffer Excel pour taux de rotation par article
            buffer_articles = io.BytesIO()
            with pd.ExcelWriter(buffer_articles, engine="xlsxwriter") as writer:
                df_resultat.to_excel(writer, index=False,
                                     sheet_name="Taux de rotation par article")
            buffer_articles.seek(0)

            # Préparer buffer Excel pour taux de rotation par famille
            buffer_familles = io.BytesIO()
            with pd.ExcelWriter(buffer_familles, engine="xlsxwriter") as writer:
                df_famille.to_excel(writer, index=False,
                                    sheet_name="Taux de rotation par famille")
            buffer_familles.seek(0)

            col_download_1, col_download_2 = st.columns(2)

            with col_download_1:
                st.download_button(
                    label="📥 Télécharger taux de rotation par article",
                    data=buffer_articles,
                    file_name="taux_de_rotation_articles.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            with col_download_2:
                st.download_button(
                    label="📥 Télécharger taux de rotation par famille",
                    data=buffer_familles,
                    file_name="taux_de_rotation_familles.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error(f"Erreur pendant le traitement : {e}")
