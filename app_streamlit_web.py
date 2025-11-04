import streamlit as st
import pandas as pd
import cv2
from pyzxing import BarCodeReader
from PIL import Image
import os

# === CONFIGURATION ===
st.set_page_config(page_title="Suivi Livraisons Batteries - My Solar Battery", page_icon="🔋", layout="centered")

# === LOGO ===
if os.path.exists("logo.png"):
    st.image("logo.png", width=220)
else:
    st.warning("⚠️ Logo non trouvé (assure-toi que logo.png est bien présent dans le dépôt)")

# === TITRE ===
st.markdown("<h2 style='text-align:center;'>Suivi des livraisons de batteries</h2>", unsafe_allow_html=True)

# === CHOIX DE LA DATE ===
date_livraison = st.date_input("📅 Date de livraison")

# === CHOIX DU FOURNISSEUR ===
fournisseurs = ["Mavisun", "Pylontech", "Dyness", "Autre"]
fournisseur = st.selectbox("🏭 Fournisseur :", fournisseurs)

autre_fournisseur = ""
if fournisseur == "Autre":
    autre_fournisseur = st.text_input("➡️ Nom du fournisseur")

# === CHOIX DE L’IMAGE ===
uploaded_image = st.file_uploader("📸 Importer la photo du code-barres", type=["png", "jpg", "jpeg"])

# === TRAITEMENT DU CODE-BARRES ===
decoded_text = ""
if uploaded_image is not None:
    image_path = "temp_image.png"
    with open(image_path, "wb") as f:
        f.write(uploaded_image.getbuffer())

    reader = BarCodeReader()
    result = reader.decode(image_path)

    if result and len(result) > 0 and "parsed" in result[0]:
        decoded_text = result[0]["parsed"]
        st.success(f"✅ Code barre détecté : {decoded_text}")
    else:
        st.error("❌ Aucun code-barres détecté. Essaie une autre photo.")

# === CHOIX DU FICHIER EXCEL ===
st.markdown("📁 Sélection du fichier Excel pour enregistrer la livraison :")
excel_file = "Livraisons_Batteries.xlsx"

if os.path.exists(excel_file):
    df = pd.read_excel(excel_file)
else:
    df = pd.DataFrame(columns=["Date", "Fournisseur", "Code", "Image"])

# === VALIDATION ===
if st.button("💾 Enregistrer la livraison"):
    fournisseur_final = autre_fournisseur if fournisseur == "Autre" else fournisseur
    new_row = pd.DataFrame([[date_livraison, fournisseur_final, decoded_text, uploaded_image.name if uploaded_image else ""]], 
                           columns=["Date", "Fournisseur", "Code", "Image"])
    df = pd.concat([df, new_row], ignore_index=True)
    df.to_excel(excel_file, index=False)
    st.success("✅ Livraison enregistrée avec succès !")

# === AFFICHAGE DU TABLEAU ===
st.markdown("---")
st.subheader("📊 Historique des livraisons enregistrées :")
st.dataframe(df)
