import streamlit as st
import os
import pytesseract
import openpyxl
import cv2
from pyzbar.pyzbar import decode
from datetime import datetime
import re
import shutil

# === CONFIGURATION ===
st.set_page_config(page_title="Suivi Livraisons Batteries", page_icon="🔋", layout="centered")

# === Authentification basique ===
password = st.text_input("🔒 Entrez le mot de passe pour accéder à l'application :", type="password")
if password != "MySolar2025":  # 🔁 à personnaliser
    st.warning("Accès restreint. Entrez le bon mot de passe.")
    st.stop()
    
# === Titre & Logo ===
if os.path.exists("logo.png"):
    st.image("logo.png", width=180)
st.title("📦 Suivi des livraisons de batteries - My Solar Battery")

# === Dossiers ===
DOSSIER_DROPBOX = os.path.expanduser("~/My Solar Battery Dropbox/Evan Le Bris/Sauvegardes_Batteries")
EXCEL_LOCAL = os.path.expanduser("~/Documents/Livraisons_Batteries_local.xlsx")
EXCEL_DROPBOX = os.path.join(DOSSIER_DROPBOX, "Livraisons_Batteries.xlsx")

os.makedirs(DOSSIER_DROPBOX, exist_ok=True)

# === Initialisation Excel ===
if not os.path.exists(EXCEL_LOCAL):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Livraisons"
    ws.append(["Date de livraison", "Fournisseur", "Numéro de série"])
    wb.save(EXCEL_LOCAL)
    wb.close()

# === OCR / Lecture code barre ===
def extraire_numero_serie(image_path):
    img = cv2.imdecode(image_path, cv2.IMREAD_COLOR)
    if img is None:
        return None
    # Lecture via code-barres
    barcodes = decode(img)
    for b in barcodes:
        data = b.data.decode("utf-8").strip()
        if re.match(r"Y\d{6,8}C\d{6,8}", data):
            return data
    # Fallback OCR
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    texte = pytesseract.image_to_string(gray, config="--oem 1 --psm 6")
    texte = texte.upper().replace(" ", "").replace("\n", "")
    corrections = {"|": "1", "I": "1", "O": "0", "S": "5", "Z": "2"}
    for k, v in corrections.items():
        texte = texte.replace(k, v)
    texte = re.sub(r"[^A-Z0-9]", "", texte)
    match = re.search(r"Y\d{5,8}C\d{5,8}", texte)
    return match.group(0) if match else None

# === Interface ===
st.subheader("🗓️ Informations de livraison")

col1, col2 = st.columns(2)
with col1:
    date_livraison = st.date_input("Date de livraison", datetime.now()).strftime("%d/%m/%Y")
with col2:
    fournisseur = st.selectbox("Fournisseur", ["Mavisun", "Swissgreen", "Allosolar", "Autre..."])
    autre_fournisseur = ""
    if fournisseur == "Autre...":
        autre_fournisseur = st.text_input("Autre fournisseur")
        if autre_fournisseur.strip() != "":
            fournisseur = autre_fournisseur.strip()

st.markdown("---")

st.subheader("📸 Importer les photos de batteries")
uploaded_files = st.file_uploader("Sélectionne une ou plusieurs photos", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

# === Bouton d'analyse ===
if st.button("✅ Analyser et enregistrer dans Excel"):
    if not uploaded_files:
        st.warning("Veuillez importer au moins une photo.")
    else:
        numeros_detectes = []
        for file in uploaded_files:
            file_bytes = bytearray(file.read())
            num = extraire_numero_serie(file_bytes)
            if num:
                numeros_detectes.append(num)

        if not numeros_detectes:
            st.error("Aucun numéro de série détecté.")
        else:
            try:
                wb = openpyxl.load_workbook(EXCEL_LOCAL)
                ws = wb.active
                for numero in numeros_detectes:
                    ws.append([date_livraison, fournisseur, numero])
                wb.save(EXCEL_LOCAL)
                wb.close()
                shutil.copy2(EXCEL_LOCAL, EXCEL_DROPBOX)
                st.success(f"{len(numeros_detectes)} numéro(s) ajouté(s) avec succès !")
            except Exception as e:
                st.error(f"Erreur d'enregistrement : {e}")

st.markdown("---")
st.caption(f"📁 Fichier Excel : {EXCEL_DROPBOX}")
