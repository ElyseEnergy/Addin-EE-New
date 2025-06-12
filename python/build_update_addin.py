import os
import win32com.client
import time
import shutil
import logging
import tempfile
import sys
import subprocess

# --- Configuration ---
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
WORKSPACE_DIR = os.path.dirname(SCRIPT_DIR)

# Fichiers source et destination
XLSM_SOURCE = os.path.join(WORKSPACE_DIR, "Update Addin.xlsm")
XLAM_DESTINATION = os.path.join(WORKSPACE_DIR, "build", "EE Addin Update.xlam")

def setup_logging():
    """Configuration du logging pour le script de build."""
    log_file = os.path.join(WORKSPACE_DIR, "build_update_addin.log")
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

def build_update_addin():
    """Convertit le fichier XLSM en XLAM."""
    setup_logging()
    logging.info("Démarrage du build de l'addin d'update...")
    
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
    
    try:
        # 1. Démarrer Excel
        logging.debug("Démarrage d'Excel...")
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        logging.debug("Excel démarré avec succès")

        # 2. Ouvrir le fichier source
        try:
            logging.debug(f"Tentative d'ouverture du fichier source: {XLSM_SOURCE}")
            workbook = excel.Workbooks.Open(XLSM_SOURCE)
            logging.debug("Fichier source ouvert avec succès")
        except Exception as e:
            logging.error(f"Erreur lors de l'ouverture du fichier source: {str(e)}")
            raise

        # 3. Exécuter les tests
        logging.info("Exécution des tests...")
        try:
            workbook.Application.Run("RunAllTests")
            logging.info("Tests exécutés avec succès")
        except Exception as e:
            logging.error(f"Erreur lors de l'exécution des tests: {str(e)}")
            raise

        # 4. Configurer le classeur comme un Add-in et sauvegarder
        logging.debug("Configuration du classeur en tant qu'add-in...")
        workbook.IsAddin = True
        logging.info("Propriété 'IsAddin' définie à True")
        
        logging.debug(f"Sauvegarde du fichier XLAM: {XLAM_DESTINATION}")
        try:
            workbook.SaveAs(XLAM_DESTINATION, FileFormat=55)  # 55 = xlOpenXMLAddIn
            logging.info(f"Fichier XLAM généré avec succès: {XLAM_DESTINATION}")
        except Exception as e:
            logging.error(f"Erreur lors de la sauvegarde du fichier XLAM: {str(e)}")
            raise

    except Exception as e:
        logging.error(f"Une erreur est survenue durant le build: {e}")
        if hasattr(e, 'exc_info'):
            logging.error(f"Détails de l'erreur: {e.exc_info}")
        raise
    finally:
        # 5. Nettoyage
        if excel:
            logging.debug("Nettoyage d'Excel...")
            excel.DisplayAlerts = False
            if workbook:
                try:
                    workbook.Close(SaveChanges=False)
                    logging.debug("Classeur fermé")
                except:
                    pass
            try:
                excel.Quit()
                logging.debug("Excel fermé avec succès")
            except:
                pass

        logging.info("Processus de build terminé.")

if __name__ == "__main__":
    try:
        build_update_addin()
    except Exception as e:
        logging.error(f"Erreur fatale non gérée: {str(e)}")
        sys.exit(1) 