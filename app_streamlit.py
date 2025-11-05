import os
import re
import cv2
import pytesseract
import openpyxl
import shutil
import tempfile
from PIL import Image
from datetime import datetime
import streamlit as st

# === CONFIGURATION ===
pytesseract.pytesseract.tesseract_cmd = "/opt/homebrew/bin/tesseract"  # pour ton Mac
DOSSIER_DROPBOX = r"/Users/evanl/My Solar Battery Dropbox/Evan Le Bris/Sauvegardes_Batteries"
EXCEL_LOCAL = os.path.expanduser("~/Documents/Livraisons_Batteries_local.xlsx")
EXCEL_DROPBOX = os.path.join(DOSSIER_DROPBOX, "Livraisons_Batteries.xlsx")

os.makedirs(DOSSIER_DROPBOX, exist_ok=True)

if not os.path.exists(EXCEL_LOCAL):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Livraisons"
    ws.append(["Date de livraison", "Fournisseur", "Numéro de série"])
    wb.save(EXCEL_LOCAL)
    wb.close()

# === OCR HYBRIDE (local + cloud) ===
def extraire_numero_serie(image_path):
    """Lecture hybride : pyzbar (local) → pyzxing (cloud) → OCR brut."""
    # 1️⃣ Tentative locale avec pyzbar
    try:
        from pyzbar.pyzbar import decode
        img = cv2.imread(image_path)
        barcodes = decode(img)
        for b in barcodes:
            data = b.data.decode("utf-8").strip()
            if re.match(r"Y\d{6,8}C\d{6,8}", data):
                return data
    except Exception:
        pass

    # 2️⃣ Tentative cloud avec pyzxing
    try:
        from pyzxing import BarCodeReader
        reader = BarCodeReader()
        result = reader.decode(image_path)
        if result and "parsed" in result[0]:
            data = result[0]["parsed"].strip()
            if re.match(r"Y\d{6,8}C\d{6,8}", data):
                return data
    except Exception:
        pass

    # 3️⃣ Fallback OCR texte brut
    try:
        gray = cv2.imread(image_path, cv2.IMREAD_GRAYSCALE)
        custom_config = "--oem 1 --psm 6 -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
        texte = pytesseract.image_to_string(gray, config=custom_config)
        texte = texte.upper().replace(" ", "").replace("\n", "")
        corrections = {"|": "1", "I": "1", "O": "0", "S": "5", "Z": "2"}
        for k, v in corrections.items():
            texte = texte.replace(k, v)
        texte = re.sub(r"[^A-Z0-9]", "", texte)
        match = re.search(r"Y\d{5,8}C\d{5,8}", texte)
        if match:
            return match.group(0)
    except Exception:
        pass

    return None


# === EXCEL ===
def enregistrer_excel(donnees):
    """Enregistre les données dans Excel local et copie vers Dropbox."""
    try:
        wb = openpyxl.load_workbook(EXCEL_LOCAL)
        ws = wb.active
        for ligne in donnees:
            ws.append(ligne)
        wb.save(EXCEL_LOCAL)
        wb.close()
        shutil.copy2(EXCEL_LOCAL, EXCEL_DROPBOX)
        st.success("✅ Données enregistrées et synchronisées avec Dropbox.")
    except Exception as e:
        st.error(f"Erreur lors de l'enregistrement : {e}")


# === INTERFACE STREAMLIT ===
st.set_page_config(page_title="📦 Suivi Livraisons - My Solar Battery", layout="centered")

# Logo
try:
    st.image("logo.png", width=220)
except:
    st.warning("⚠️ Logo non trouvé (logo.png)")

st.title("📦 Suivi des livraisons de batteries")

# Champs
date_livraison = st.date_input("📅 Date de livraison :", datetime.now())
fournisseurs = ["Mavisun", "Swissgreen", "Allosolar", "Autre..."]
fournisseur = st.selectbox("🏭 Fournisseur :", fournisseurs)

autre_fournisseur = ""
if fournisseur == "Autre...":
    autre_fournisseur = st.text_input("➡️ Indique le nom du fournisseur")

# Import des images
uploaded_files = st.file_uploader(
    "📸 Importer les photos des batteries",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True
)

# Bouton d’analyse
if st.button("✅ Enregistrer dans Excel"):
    if not uploaded_files:
        st.warning("Merci d'importer au moins une photo.")
        st.stop()

    fournisseur_final = autre_fournisseur if fournisseur == "Autre..." else fournisseur
    if not fournisseur_final:
        st.warning("Merci d'indiquer un fournisseur.")
        st.stop()

    numeros_detectes = []
    temp_dir = tempfile.mkdtemp()

    for uploaded_file in uploaded_files:
        path_temp = os.path.join(temp_dir, uploaded_file.name)
        with open(path_temp, "wb") as f:
            f.write(uploaded_file.getbuffer())
        numero = extraire_numero_serie(path_temp)
        if numero:
            numeros_detectes.append(numero)

    if not numeros_detectes:
        st.error("❌ Aucun numéro détecté sur les images fournies.")
        st.stop()

    donnees = [[date_livraison.strftime("%d/%m/%Y"), fournisseur_final, num] for num in numeros_detectes]
    enregistrer_excel(donnees)
    st.success(f"{len(numeros_detectes)} numéro(s) ajouté(s) avec succès ✅")
