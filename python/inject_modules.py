import os
import win32com.client
import time
import psutil
import logging
import sys
from datetime import datetime
import traceback
import shutil
import tempfile

# --- Définition des chemins ---
# Le répertoire où se trouve ce script (.py)
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
# Le répertoire racine du projet, qui est le parent du répertoire du script
WORKSPACE_DIR = os.path.dirname(SCRIPT_DIR)

def setup_logging():
    """Configure le système de logging"""
    log_dir = os.path.join(SCRIPT_DIR, "logs")
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
    
    log_file = os.path.join(log_dir, "inject_modules.log")
    
    # Si le fichier dépasse 2Mo, on le vide
    if os.path.exists(log_file) and os.path.getsize(log_file) > 2*1024*1024:
        open(log_file, 'w').close()
    
    logging.basicConfig(
        level=logging.DEBUG,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8'),
            logging.StreamHandler(sys.stdout)
        ]
    )
    return log_file

def prepare_temp_file(src_file_path, temp_dir):
    """Copie un module VBA dans un dossier temporaire.
    - Pour les .frm, copie brute du .frm et du .frx associé.
    - Pour les .bas/.cls, gère l'encodage (UTF-8 -> CP1252)."""
    base_name = os.path.basename(src_file_path)
    temp_file = os.path.join(temp_dir, base_name)
    
    try:
        # Cas des formulaires (.frm) : copie brute, sans conversion d'encodage.
        if src_file_path.lower().endswith('.frm'):
            # Copie du .frm
            shutil.copy2(src_file_path, temp_file)
            logging.info(f"Fichier .frm '{base_name}' copié tel quel (brut).")
            
            # Copie du .frx associé s'il existe
            frx_file = src_file_path[:-4] + '.frx'
            if os.path.exists(frx_file):
                shutil.copy2(frx_file, os.path.join(temp_dir, os.path.basename(frx_file)))
                logging.info(f"Fichier .frx associé '{os.path.basename(frx_file)}' copié.")
        
        # Cas des modules texte (.bas, .cls).
        else:
            # On s'assure qu'ils sont en encodage CP1252, requis par l'IDE VBA.
            try:
                with open(src_file_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                # Le fichier source est en UTF-8, on le convertit en CP1252 dans le dossier temp.
                with open(temp_file, 'w', encoding='cp1252', errors='replace') as f:
                    f.write(content)
                logging.info(f"Fichier '{base_name}' converti de UTF-8 à CP1252.")
            except UnicodeDecodeError:
                # Le fichier source n'est pas en UTF-8, on suppose qu'il est déjà dans un format compatible
                # (comme CP1252) et on le copie simplement.
                shutil.copy2(src_file_path, temp_file)
                logging.info(f"Fichier '{base_name}' n'est pas en UTF-8, copié tel quel.")
        
        return temp_file
    except Exception as e:
        logging.error(f"Erreur lors de la préparation du fichier temporaire pour '{base_name}': {str(e)}")
        return None

def get_excel_instance(target_file=None):
    """Récupère une instance Excel existante avec notre fichier ou en crée une nouvelle"""
    try:
        # Essayer de se connecter à une instance existante
        excel = win32com.client.GetActiveObject("Excel.Application")
        
        # Si on a un fichier cible, vérifier s'il est déjà ouvert
        if target_file:
            target_file = os.path.abspath(target_file)
            for wb in excel.Workbooks:
                if os.path.abspath(wb.FullName) == target_file:
                    logging.info(f"Fichier {target_file} déjà ouvert dans Excel")
                    return excel, wb
        
        # Si le fichier n'est pas ouvert dans cette instance, on la ferme
        excel.Quit()
        
    except Exception as e:
        logging.debug(f"Pas d'instance Excel active: {str(e)}")
    
    # Créer une nouvelle instance
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    logging.info("Nouvelle instance Excel créée")
    return excel, None

def check_vba_access():
    """Vérifie si l'accès au VBA Project est activé"""
    logging.info("Vérification de l'accès au VBA Project")
    excel = None
    wb = None
    
    try:
        excel = win32com.client.DispatchEx('Excel.Application')
        wb = excel.Workbooks.Add()
        
        try:
            _ = wb.VBProject
            logging.info("Accès au VBA Project confirmé")
            return True
        except Exception as e:
            logging.error(f"Accès au VBA Project refusé: {str(e)}")
            logging.info("""
            Pour activer l'accès au VBA Project:
            1. Ouvrez Excel
            2. Allez dans Fichier > Options > Centre de gestion de la confidentialité
            3. Cliquez sur 'Paramètres du Centre de gestion de la confidentialité'
            4. Sélectionnez 'Paramètres des macros'
            5. Cochez 'Faire confiance à l'accès au modèle d'objet des projets VBA'
            6. Cliquez sur OK et redémarrez Excel
            """)
            return False
    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
            if excel is not None:
                excel.Quit()
        except:
            pass

def remove_all_bas_modules(vba_project):
    """Supprime tous les modules .bas du projet VBA"""
    modules_to_remove = []
    for comp in vba_project.VBComponents:
        # Type 1 = vbext_ct_StdModule (.bas)
        if comp.Type == 1:
            modules_to_remove.append(comp.Name)
    
    for module_name in modules_to_remove:
        try:
            vba_project.VBComponents.Remove(vba_project.VBComponents(module_name))
            logging.info(f"Module .bas supprimé: {module_name}")
        except Exception as e:
            logging.error(f"Erreur lors de la suppression du module {module_name}: {str(e)}")
    
    logging.info(f"Total des modules .bas supprimés: {len(modules_to_remove)}")

def remove_all_vba_modules(vba_project):
    """Supprime tous les modules VBA (StdModule, ClassModule, Form) sauf les modules Document (type 100)."""
    modules_to_remove = []
    for comp in vba_project.VBComponents:
        # Type 1 = StdModule (.bas), 2 = ClassModule (.cls), 3 = MSForm (.frm)
        if comp.Type in (1, 2, 3):
            modules_to_remove.append(comp.Name)
    for module_name in modules_to_remove:
        try:
            vba_project.VBComponents.Remove(vba_project.VBComponents(module_name))
            logging.info(f"Module VBA supprimé: {module_name}")
        except Exception as e:
            logging.error(f"Erreur lors de la suppression du module {module_name}: {str(e)}")
    logging.info(f"Total des modules VBA supprimés: {len(modules_to_remove)}")

def clean_document_code(code):
    """Nettoie le code d'un module Document en enlevant les attributs VB"""
    lines = code.split('\n')
    cleaned_lines = []
    skip_attributes = True
    
    for line in lines:
        if line.strip().startswith('Attribute VB_'):
            continue  # On saute les lignes d'attributs
        cleaned_lines.append(line)
    
    return '\n'.join(cleaned_lines)

def inject_modules():
    log_file = setup_logging()
    logging.info(f"Début de l'injection des modules. Log file: {log_file}")
    
    if not check_vba_access():
        logging.error("Accès VBA non disponible. Arrêt du processus.")
        return

    excel = None
    workbook = None
    excel_file = os.path.join(WORKSPACE_DIR, "Addin Elyse Energy.xlsm")
    temp_dir = None
    
    try:
        logging.info(f"Fichier Excel cible: {excel_file}")
        
        if not os.path.exists(excel_file):
            logging.error(f"Le fichier {excel_file} n'existe pas!")
            return
        
        # Créer un dossier temporaire
        temp_dir = tempfile.mkdtemp(prefix="vba_modules_")
        logging.info(f"Dossier temporaire créé: {temp_dir}")
        
        # Récupérer ou créer une instance Excel
        excel, workbook = get_excel_instance(excel_file)
        
        # Si on n'a pas récupéré le workbook, on l'ouvre
        if not workbook:
            try:
                workbook = excel.Workbooks.Open(excel_file)
            except Exception as e:
                logging.error(f"Erreur lors de l'ouverture du fichier: {str(e)}")
                # Essayer de fermer toutes les instances Excel
                try:
                    os.system('taskkill /F /IM excel.exe')
                    time.sleep(2)  # Attendre que Excel se ferme
                    excel = win32com.client.DispatchEx("Excel.Application")
                    excel.Visible = False
                    workbook = excel.Workbooks.Open(excel_file)
                except Exception as e2:
                    logging.error(f"Impossible d'ouvrir le fichier même après avoir fermé Excel: {str(e2)}")
                    raise
        
        vba_project = workbook.VBProject
        
        # Nettoyage complet : supprimer tous les modules VBA (sauf Document)
        remove_all_vba_modules(vba_project)
        
        # Dictionnaire des composants existants pour des recherches rapides
        all_components = {comp.Name: comp.Type for comp in vba_project.VBComponents}
        logging.info(f"Composants trouvés dans le classeur: {', '.join(all_components.keys())}")

        # Lister les fichiers de module depuis la racine du projet
        module_files = [f for f in os.listdir(WORKSPACE_DIR) if f.endswith(('.bas', '.cls', '.frm'))]
        logging.info(f"Fichiers de module trouvés sur le disque: {', '.join(module_files)}")

        # Étape 1: Supprimer les anciens composants qui vont être remplacés.
        # Cela évite les conflits lors de la ré-importation de modules modifiés.
        # On ne touche PAS aux modules de type Document (Type 100).
        for file_name in module_files:
            module_name = os.path.splitext(file_name)[0]
            if module_name in all_components and all_components[module_name] != 100:
                try:
                    vba_project.VBComponents.Remove(vba_project.VBComponents(module_name))
                    logging.info(f"Ancien composant '{module_name}' supprimé avant la nouvelle injection.")
                except Exception as e:
                    logging.warning(f"Impossible de supprimer le composant existant '{module_name}': {e}")

        # Étape 2: Injecter les modules depuis le disque.
        for file_name in module_files:
            module_name = os.path.splitext(file_name)[0]
            # Construire le chemin complet vers le fichier de module
            file_path = os.path.join(WORKSPACE_DIR, file_name)
            try:
                # Prépare le fichier pour l'import (gère l'encodage et les fichiers .frx)
                temp_file = prepare_temp_file(file_path, temp_dir)
                if not temp_file:
                    continue
                
                # Cas spécial: mise à jour du code d'un module Document existant.
                if module_name in all_components and all_components[module_name] == 100:
                    with open(temp_file, 'r', encoding='cp1252') as f:
                        new_code = f.read()
                    
                    # Nettoyage du code avant injection
                    if new_code.startswith('VERSION 1.0 CLASS'):
                        new_code = new_code.split('END\n', 1)[1]
                    new_code = clean_document_code(new_code)
                    
                    code_module = vba_project.VBComponents(module_name).CodeModule
                    code_module.DeleteLines(1, code_module.CountOfLines)
                    code_module.AddFromString(new_code)
                    logging.info(f"Code du module Document '{module_name}' mis à jour.")
                else:
                    # Cas standard: importation d'un nouveau module (.bas, .cls, .frm).
                    # L'ancienne version a déjà été supprimée à l'étape 1.
                    vba_project.VBComponents.Import(temp_file)
                    logging.info(f"Module '{module_name}' importé avec succès depuis '{file_name}'.")
            
            except Exception as e:
                logging.error(f"Erreur lors du traitement de '{module_name}' depuis '{file_name}': {e}")
                logging.error(traceback.format_exc())
        
        # Sauvegarder
        logging.info("Sauvegarde des modifications...")
        workbook.Save()
        
        # Forcer le refresh du Ribbon
        try:
            excel.Application.CommandBars.ExecuteMso("HideRibbon")
            time.sleep(0.1)
            excel.Application.CommandBars.ExecuteMso("HideRibbon")
            logging.info("Ribbon rafraîchi")
        except:
            logging.warning("Impossible de rafraîchir le Ribbon via CommandBars")
            
        # Rendre Excel visible
        excel.Visible = True
        logging.info("Excel rendu visible")
        
    except Exception as e:
        logging.error(f"Erreur générale: {str(e)}")
        logging.error(traceback.format_exc())
        try:
            if workbook:
                workbook.Close(SaveChanges=False)
            if excel:
                excel.Quit()
        except:
            pass
    finally:
        # Nettoyer le dossier temporaire
        if temp_dir and os.path.exists(temp_dir):
            try:
                shutil.rmtree(temp_dir)
                logging.info(f"Dossier temporaire nettoyé: {temp_dir}")
            except Exception as e:
                logging.error(f"Erreur lors du nettoyage du dossier temporaire: {str(e)}")

if __name__ == "__main__":
    try:
        inject_modules()
    except Exception as e:
        print(f"Erreur fatale: {str(e)}")
        print(traceback.format_exc())
        sys.exit(1)
