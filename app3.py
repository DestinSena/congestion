# ================================================
# streamlit_app3.py
# Analyse des cellules récurrentes dans les feuilles S1, S2, S3...
# ================================================

import streamlit as st
import pandas as pd
import io
from collections import defaultdict
from datetime import datetime

# ====================== CONFIGURATION ======================
st.set_page_config(
    page_title="Cellules Récurrentes - Congestion",
    page_icon="📊",
    layout="wide"
)

st.title("📡 Analyse Congestion - Cellules Récurrentes")
st.markdown("""
**Fonctionnement :**  
Cette application identifie les **cellules** qui apparaissent dans **au moins la moitié** des feuilles (S1, S2, S3...) du fichier Excel.  
Ajoutez simplement une nouvelle feuille **Sxx** chaque semaine → l’application s’adapte automatiquement.
""")


# ====================== FONCTION DE TÉLÉCHARGEMENT ======================
def to_excel(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Cellules_Recurrentes')
    output.seek(0)
    return output.getvalue()


# ====================== UPLOAD DU FICHIER ======================
uploaded_file = st.file_uploader(
    "📤 Charger le fichier Excel (congestion analysis.xlsx)",
    type=["xlsx", "xls"],
    help="Le fichier doit contenir des feuilles nommées S1, S2, S3, ..."
)

if uploaded_file is not None:
    try:
        # Lecture du fichier
        xls = pd.ExcelFile(uploaded_file)

        # Filtrer uniquement les feuilles commençant par "S"
        s_sheets = [sheet for sheet in xls.sheet_names if sheet.startswith("S")]

        if not s_sheets:
            st.error("❌ Aucune feuille commençant par 'S' n'a été trouvée dans le fichier.")
            st.stop()

        num_sheets = len(s_sheets)
        threshold = num_sheets / 2  # Moitié des feuilles

        st.success(
            f"✅ **{num_sheets}** feuilles détectées : {', '.join(sorted(s_sheets)[:10])}{'...' if len(s_sheets) > 10 else ''}")
        st.info(f"**Seuil de répétition** : une cellule doit apparaître dans **au moins {threshold}** feuille(s)")

        # ====================== ANALYSE ======================
        with st.spinner("Analyse des cellules en cours..."):
            cellule_dict = defaultdict(lambda: {"nom_site": None, "count": 0})

            for sheet_name in s_sheets:
                df = pd.read_excel(xls, sheet_name=sheet_name, usecols=["nom_site", "cellule"])
                df = df.dropna(subset=["cellule"])
                df["cellule"] = df["cellule"].astype(str).str.strip()
                df["nom_site"] = df["nom_site"].astype(str).str.strip()

                for _, row in df.iterrows():
                    cellule = row["cellule"]
                    if cellule and cellule.lower() != "nan":
                        if cellule_dict[cellule]["nom_site"] is None:
                            cellule_dict[cellule]["nom_site"] = row["nom_site"]
                        cellule_dict[cellule]["count"] += 1

        # ====================== FILTRAGE DES RÉSULTATS ======================
        results = []
        for cellule, info in cellule_dict.items():
            if info["count"] >= threshold:
                results.append({
                    "nom_site": info["nom_site"],
                    "cellules": cellule,
                    "nombre de répétition": info["count"]
                })

        if results:
            result_df = pd.DataFrame(results)
            result_df = result_df.sort_values(by="nombre de répétition", ascending=False).reset_index(drop=True)

            st.subheader(f"📋 **{len(result_df)}** cellules récurrentes détectées")

            # Affichage du tableau
            st.dataframe(
                result_df,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "nombre de répétition": st.column_config.NumberColumn(format="%d")
                }
            )

            # Statistiques
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Cellules trouvées", len(result_df))
            with col2:
                st.metric("Répétition maximale", result_df["nombre de répétition"].max())
            with col3:
                st.metric("Répétition minimale", result_df["nombre de répétition"].min())

            # ====================== TÉLÉCHARGEMENT ======================
            excel_bytes = to_excel(result_df)

            st.download_button(
                label="📥 Télécharger le résultat en Excel",
                data=excel_bytes,
                file_name=f"cellules_recurrentes_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        else:
            st.warning("⚠️ Aucune cellule n'atteint le seuil de répétition (moitié des feuilles).")

    except Exception as e:
        st.error(f"❌ Erreur lors du traitement du fichier : {str(e)}")
        st.info("Vérifiez que votre fichier contient bien les colonnes : **date, nom_site, cellule, %prb**")

else:
    st.info("👆 Veuillez charger votre fichier Excel pour lancer l'analyse.")
    st.caption(
        "💡 Astuce : Ajoutez une nouvelle feuille nommée **S13**, **S14**, etc. chaque semaine. L’application s’adapte automatiquement.")

st.caption("Application développée pour le suivi hebdomadaire de la congestion réseau")