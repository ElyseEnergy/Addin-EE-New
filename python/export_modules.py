import os
import win32com.client
import time
import sys
import traceback
import psutil
from datetime import datetime
import logging

# --- Définition des chemins ---
# Le répertoire où se trouve ce script (.py)
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
# Le répertoire racine du projet, qui est le parent du répertoire du script
WORKSPACE_DIR = os.path.dirname(SCRIPT_DIR)

def setup_logging():
    """Configure le système de logging"""
    # Le répertoire des logs est maintenant DANS le répertoire du script
    log_dir = os.path.join(SCRIPT_DIR, "logs")
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
    
    log_file = os.path.join(log_dir, "export_modules.log")
    logging.basicConfig(
        level=logging.DEBUG,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8'),
            logging.StreamHandler(sys.stdout)
        ]
    )
    return log_file

def kill_excel():
    """Ferme toutes les instances d'Excel"""
    logging.info("Tentative de fermeture de toutes les instances Excel...")
    killed = False
    try:
        for proc in psutil.process_iter(['pid', 'name', 'create_time']):
            try:
                if proc.info['name'] == 'EXCEL.EXE':
                    proc_info = f"PID: {proc.info['pid']}, Créé le: {datetime.fromtimestamp(proc.info['create_time']).strftime('%Y-%m-%d %H:%M:%S')}"
                    logging.debug(f"Instance Excel trouvée - {proc_info}")
                    proc.kill()
                    killed = True
                    logging.info(f"Instance Excel terminée - {proc_info}")
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess) as e:
                logging.warning(f"Erreur lors de la tentative de fermeture d'Excel: {str(e)}")
    except Exception as e:
        logging.error(f"Erreur inattendue lors de la fermeture d'Excel: {str(e)}")
    
    if killed:
        logging.info("Attente de 2 secondes après la fermeture d'Excel...")
        time.sleep(2)
    else:
        logging.info("Aucune instance d'Excel n'a été trouvée")

def check_vba_access():
    """Vérifie si l'accès au VBA Project est activé"""
    logging.info("Vérification de l'accès au VBA Project")
    excel = None
    wb = None
    
    try:
        excel = win32com.client.Dispatch('Excel.Application')
        excel.Visible = False
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
            kill_excel()
        except:
            pass

def convert_to_utf8(file_path):
    """Convertit un fichier de CP1252 vers UTF-8"""
    try:
        # Lire en CP1252
        with open(file_path, 'r', encoding='cp1252') as f:
            content = f.read()
        
        # Réécrire en UTF-8
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(content)
        logging.info(f"Converti en UTF-8: {file_path}")
    except Exception as e:
        logging.error(f"Erreur lors de la conversion UTF-8 de {file_path}: {str(e)}")

def export_modules(export_dir=None):
    """Exporte tous les modules VBA d'un fichier .xlsm vers un répertoire."""
    log_file = setup_logging()
    logging.info(f"Début de l'export des modules. Log file: {log_file}")
    
    if not check_vba_access():
        logging.error("Accès VBA non disponible. Arrêt du processus.")
        return

    # Par défaut, on exporte à la racine du projet
    if export_dir is None:
        export_dir = WORKSPACE_DIR

    excel_file = os.path.join(WORKSPACE_DIR, "Addin Elyse Energy.xlsm")
    
    excel = None
    workbook = None
    
    try:
        # S'assurer que le dossier d'export existe
        os.makedirs(export_dir, exist_ok=True)
        logging.info(f"Fichier Excel source: {excel_file}")
        logging.info(f"Dossier d'export: {export_dir}")
        
        if not os.path.exists(excel_file):
            logging.error(f"Le fichier {excel_file} n'existe pas!")
            return
            
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        workbook = excel.Workbooks.Open(excel_file)
        
        vba_project = workbook.VBProject
        
        # Exporter tous les composants
        for comp in vba_project.VBComponents:
            try:
                # Gérer les différents types de composants et leur extension
                if comp.Type == 1: extension = "bas"
                elif comp.Type == 2: extension = "cls"
                elif comp.Type == 3: extension = "frm"
                elif comp.Type == 100: extension = "cls" # Modules de document
                else: continue # Ne pas exporter les types inconnus

                export_filename = os.path.join(export_dir, f"{comp.Name}.{extension}")
                
                # Exporter le composant
                comp.Export(export_filename)
                logging.info(f"Module exporté: {export_filename}")

                # Pour les .bas et .cls, qui sont du texte pur, on convertit en UTF-8.
                # Pour les .frm, on ne touche à rien pour ne pas corrompre le format.
                if extension in ["bas", "cls"]:
                    convert_to_utf8(export_filename)
            
            except Exception as e:
                logging.error(f"Erreur lors de l'export de {comp.Name}: {str(e)}")
                logging.error(traceback.format_exc())
        
        logging.info("Export terminé avec succès.")
        
    except Exception as e:
        logging.error(f"Erreur générale: {str(e)}")
        logging.error(traceback.format_exc())
    finally:
        try:
            if workbook:
                workbook.Close(SaveChanges=False)
            if excel:
                excel.Quit()
        except:
            pass

if __name__ == "__main__":
    try:
        # Par défaut, on exporte à la racine du projet
        export_modules()
    except Exception as e:
        print(f"Erreur fatale: {str(e)}")
        print(traceback.format_exc())
        sys.exit(1)