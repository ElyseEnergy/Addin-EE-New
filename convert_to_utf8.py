import os
import glob

def convert_to_utf8(file_path):
    try:
        # Lire en cp1252
        with open(file_path, 'r', encoding='cp1252') as f:
            content = f.read()
        
        # Réécrire en UTF-8
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(content)
        print(f"Converti: {file_path}")
    except Exception as e:
        print(f"Erreur sur {file_path}: {str(e)}")

# Convertir tous les .bas et .cls
for pattern in ['*.bas', '*.cls']:
    for file in glob.glob(pattern):
        convert_to_utf8(file) 