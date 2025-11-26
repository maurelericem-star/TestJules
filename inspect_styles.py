# inspect_styles.py
from docx import Document

# Charge votre modèle
doc = Document('MAUREL_Eric_CV_RATP.docm')

print("--- LISTE DES STYLES DISPONIBLES ---")
# Affiche tous les styles de type 'Paragraphe' utilisés
for style in doc.styles:
    if style.type.name == 'PARAGRAPH':
        print(f"Nom du style : {style.name}")