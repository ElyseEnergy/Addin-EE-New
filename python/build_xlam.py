import os
import win32com.client
import time
import shutil
import logging
import tempfile
import re
import sys
import subprocess
import zipfile

# --- Configuration ---
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
WORKSPACE_DIR = os.path.dirname(SCRIPT_DIR)

# Fichier source et destination
XLSM_SOURCE = os.path.join(WORKSPACE_DIR, "Addin Elyse Energy.xlsm")

# --- Extraction de la version depuis Utilities.bas ---
def extract_version_from_utilities():
    util_path = os.path.join(WORKSPACE_DIR, "Utilities.bas")
    major = minor = patch = None
    with open(util_path, encoding="utf-8") as f:
        for line in f:
            if "ADDIN_VERSION_MAJOR" in line:
                major = line.split("=")[-1].strip()
            elif "ADDIN_VERSION_MINOR" in line:
                minor = line.split("=")[-1].strip()
            elif "ADDIN_VERSION_PATCH" in line:
                patch = line.split("=")[-1].strip()
            if major and minor and patch:
                break
    if not (major and minor and patch):
        raise RuntimeError("Impossible d'extraire la version depuis Utilities.bas")
    return f"{major}.{minor}.{patch}"

ADDIN_VERSION = extract_version_from_utilities()
XLAM_DESTINATION = os.path.join(WORKSPACE_DIR, "build", f"EE Addin_v{ADDIN_VERSION}.xlam")

# Module VBA à modifier pour désactiver le logging en production
CONFIG_MODULE_BASENAME = "SYS_Logger.bas" 

def setup_logging():
    """Configuration du logging pour le script de build."""
    log_file = os.path.join(WORKSPACE_DIR, "build_xlam.log")
    logging.basicConfig(
        level=logging.DEBUG,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.StreamHandler(),
            logging.FileHandler(log_file, encoding='utf-8', mode='w')
        ]
    )

def kill_excel_processes():
    """Tue tous les processus Excel."""
    try:
        subprocess.run(['taskkill', '/F', '/IM', 'excel.exe'], 
                      capture_output=True, 
                      text=True)
        time.sleep(2)  # Attendre que tous les processus soient bien fermés
        logging.debug("Tous les processus Excel ont été fermés.")
    except Exception as e:
        logging.warning(f"Erreur lors de la fermeture des processus Excel: {e}")

def prepare_production_module(temp_dir):
    """Prépare une version 'production' du module de config en y activant le mode Add-in."""
    src_path = os.path.join(WORKSPACE_DIR, CONFIG_MODULE_BASENAME)
    logging.debug(f"Préparation du module source: {src_path}")
    
    if not os.path.exists(src_path):
        logging.error(f"Le module de configuration '{CONFIG_MODULE_BASENAME}' n'a pas été trouvé.")
        return None

    try:
        # Lire le contenu original du module avec l'encodage par défaut du système ou utf-8
        try:
            with open(src_path, 'r', encoding='utf-8') as f:
                content = f.read()
            logging.debug("Fichier lu avec succès en UTF-8")
        except UnicodeDecodeError:
            logging.debug("Tentative de lecture en latin-1...")
            with open(src_path, 'r', encoding='latin-1') as f:
                content = f.read()
            logging.debug("Fichier lu avec succès en latin-1")
        
        # Modifier la constante de production
        modified_content = re.sub(
            r"(#Const\s+IS_ADDIN\s*=\s*)0", 
            r"\g<1>1", 
            content, 
            flags=re.IGNORECASE
        )

        if modified_content == content:
            logging.warning(f"La constante '#Const IS_ADDIN = 0' n'a pas été trouvée dans {CONFIG_MODULE_BASENAME}. Le logging fichier ne sera pas désactivé.")
        else:
            logging.debug("Constante IS_ADDIN modifiée avec succès")

        # Écrire le contenu modifié dans un fichier temporaire en CP1252
        temp_module_path = os.path.join(temp_dir, CONFIG_MODULE_BASENAME)
        with open(temp_module_path, 'w', encoding='cp1252', errors='replace') as f:
            f.write(modified_content)
            
        logging.info(f"Module '{CONFIG_MODULE_BASENAME}' préparé pour la production.")
        return temp_module_path
    except Exception as e:
        logging.error(f"Erreur lors de la préparation du module de production: {e}")
        return None

def modify_ribbon_tab_name(excel_file):
    """Modifie le nom de l'onglet dans le fichier customUI.xml de manière simple."""
    logging.debug("Modification du nom de l'onglet Ribbon...")
    
    # Créer une copie temporaire
    temp_file = excel_file + ".tmp"
    shutil.copy2(excel_file, temp_file)
    
    try:
        # Ouvrir le fichier comme un zip
        with zipfile.ZipFile(temp_file, 'r') as z:
            # Chercher le fichier customUI.xml
            xml_files = [f for f in z.namelist() if 'customui.xml' in f.lower()]
            if not xml_files:
                logging.warning("Fichier customUI.xml non trouvé.")
                return False
                
            xml_file = xml_files[0]
            xml_content = z.read(xml_file).decode('utf-8')
            
            # Modifier le contenu
            if "Elyse Energy" in xml_content:
                new_content = xml_content.replace("Elyse Energy", "EE")
                
                # Créer un nouveau zip avec le contenu modifié
                with zipfile.ZipFile(excel_file, 'w') as new_zip:
                    # Copier tous les fichiers sauf le XML
                    for item in z.filelist:
                        if item.filename != xml_file:
                            new_zip.writestr(item.filename, z.read(item.filename))
                    # Ajouter le XML modifié
                    new_zip.writestr(xml_file, new_content)
                
                logging.info("Nom de l'onglet modifié avec succès")
                return True
            else:
                logging.warning("Label 'Elyse Energy' non trouvé dans le XML")
                return False
                
    except Exception as e:
        logging.error(f"Erreur lors de la modification du Ribbon: {e}")
        return False
    finally:
        # Nettoyer le fichier temporaire
        if os.path.exists(temp_file):
            os.remove(temp_file)

def build_xlam():
    """Convertit le fichier XLSM en XLAM en mode production."""
    setup_logging()
    logging.info("Démarrage du processus de build XLAM...")
    
    # Tuer tous les processus Excel existants
    kill_excel_processes()
    
    logging.debug(f"Chemin source: {XLSM_SOURCE}")
    logging.debug(f"Chemin destination: {XLAM_DESTINATION}")

    # Vérifier que le fichier source existe
    if not os.path.exists(XLSM_SOURCE):
        logging.error(f"Le fichier source {XLSM_SOURCE} n'existe pas!")
        return

    # Créer le répertoire de build s'il n'existe pas
    build_dir = os.path.dirname(XLAM_DESTINATION)
    os.makedirs(build_dir, exist_ok=True)
    logging.debug(f"Répertoire de build créé/vérifié: {build_dir}")

    excel = None
    workbook = None
    temp_dir = tempfile.mkdtemp(prefix="xlam_build_")
    logging.debug(f"Dossier temporaire créé: {temp_dir}")
    
    try:
        # 1. Préparer le module de configuration pour la production
        prod_module_path = prepare_production_module(temp_dir)
        if not prod_module_path:
            return

        # 2. Démarrer Excel
        logging.debug("Démarrage d'Excel...")
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        logging.debug("Excel démarré avec succès")

        # 3. Ouvrir le fichier source
        try:
            logging.debug(f"Tentative d'ouverture du fichier source: {XLSM_SOURCE}")
            workbook = excel.Workbooks.Open(XLSM_SOURCE)
            logging.debug("Fichier source ouvert avec succès")
        except Exception as e:
            logging.error(f"Erreur lors de l'ouverture du fichier source: {str(e)}")
            raise

        # 4. Remplacer le module de configuration par la version de production
        vba_project = workbook.VBProject
        module_name_in_vba = os.path.splitext(CONFIG_MODULE_BASENAME)[0]
        try:
            # Essayer de trouver le module existant
            existing_module = None
            for component in vba_project.VBComponents:
                if component.Name == module_name_in_vba:
                    existing_module = component
                    break
            
            # Si le module existe, le supprimer
            if existing_module:
                logging.debug(f"Suppression de l'ancien module: {module_name_in_vba}")
                vba_project.VBComponents.Remove(existing_module)
            else:
                logging.debug(f"Module {module_name_in_vba} non trouvé dans le projet, on continue avec l'import")
            
            logging.debug("Import du nouveau module...")
            vba_project.VBComponents.Import(prod_module_path)
            logging.info(f"Module '{module_name_in_vba}' remplacé par la version de production.")
        except Exception as e:
            logging.error(f"Impossible de remplacer le module '{module_name_in_vba}'. Est-il présent dans le projet ? Erreur: {e}")
            raise

        # 5. Configurer le classeur comme un Add-in et sauvegarder
        logging.debug("Configuration du classeur en tant qu'add-in...")
        workbook.IsAddin = True
        logging.info("Propriété 'IsAddin' définie à True.")
        
        logging.debug(f"Sauvegarde du fichier XLAM: {XLAM_DESTINATION}")
        try:
            workbook.SaveAs(XLAM_DESTINATION, FileFormat=55)
            logging.info(f"Fichier XLAM généré avec succès: {XLAM_DESTINATION}")
            
            # 6. Modifier le nom de l'onglet dans le Ribbon
            if modify_ribbon_tab_name(XLAM_DESTINATION):
                logging.info("Modification du Ribbon réussie")
            else:
                logging.warning("La modification du Ribbon a échoué, mais le fichier XLAM a été généré")
                
        except Exception as e:
            logging.error(f"Erreur lors de la sauvegarde du fichier XLAM: {str(e)}")
            raise

    except Exception as e:
        logging.error(f"Une erreur est survenue durant le build: {e}")
        if hasattr(e, 'exc_info'):
            logging.error(f"Détails de l'erreur: {e.exc_info}")
    finally:
        # 7. Nettoyage
        if excel:
            logging.debug("Nettoyage d'Excel...")
            excel.DisplayAlerts = False
            # Ferme tous les classeurs restants sans sauvegarder avant de quitter
            for wb in excel.Workbooks:
                try:
                    wb.Close(SaveChanges=False)
                    logging.debug(f"Classeur fermé: {wb.Name}")
                except Exception as e:
                    logging.error(f"Erreur lors de la fermeture du classeur {wb.Name}: {str(e)}")
            try:
                excel.Quit()
                logging.debug("Excel fermé avec succès")
            except Exception as e:
                logging.error(f"Erreur lors de la fermeture d'Excel: {str(e)}")
            
        if os.path.exists(temp_dir):
            try:
                shutil.rmtree(temp_dir)
                logging.info(f"Dossier temporaire nettoyé: {temp_dir}")
            except Exception as e:
                logging.error(f"Erreur lors du nettoyage du dossier temporaire: {e}")

        logging.info("Processus de build terminé.")

if __name__ == "__main__":
    try:
        build_xlam()
    except Exception as e:
        logging.error(f"Erreur fatale non gérée: {str(e)}")
        sys.exit(1) 