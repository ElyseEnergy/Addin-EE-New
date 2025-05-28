import os
import win32com.client
import time
import psutil
import logging
import sys
from datetime import datetime
import traceback

def setup_logging():
    """Configure le système de logging"""
    log_dir = "logs"
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = os.path.join(log_dir, f"inject_modules_{timestamp}.log")
    
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

def start_excel():
    """Démarre une nouvelle instance d'Excel propre"""
    kill_excel()  # On s'assure qu'aucune instance n'est en cours
    time.sleep(1)
    try:
        excel = win32com.client.Dispatch('Excel.Application')
        excel.Visible = False  # On essaie de le cacher dès le début
        time.sleep(1)
        return excel
    except Exception as e:
        logging.error(f"Erreur lors du démarrage d'Excel: {str(e)}")
        return None

def get_module_type(file_name):
    """Détermine le type de module basé sur le nom du fichier"""
    if file_name.startswith('ThisWorkbook'):
        return 100  # vbext_ct_Document
    elif file_name.startswith('UserForm') or file_name.startswith('frm'):
        return 3    # vbext_ct_MSForm
    else:
        return 1    # vbext_ct_StdModule

def process_workbook_content(content):
    """Traite le contenu d'un module Workbook"""
    if not content.startswith('VERSION 1.0 CLASS'):
        content = 'VERSION 1.0 CLASS\nBEGIN\n  MultiUse = -1  \'True\nEND\n' + content
    return content

def read_file_with_fallback_encoding(file_path):
    """Essaie de lire le fichier avec différents encodages"""
    encodings = ['utf-8', 'cp1252', 'latin1', 'utf-8-sig']
    for encoding in encodings:
        try:
            with open(file_path, 'r', encoding=encoding) as f:
                return f.read()
        except UnicodeDecodeError:
            continue
            
    raise ValueError(f"Impossible de lire le fichier {file_path} avec les encodages disponibles.")

def check_vba_access():
    """Vérifie si l'accès au VBA Project est activé"""
    logging.info("Vérification de l'accès au VBA Project")
    excel = None
    wb = None
    
    try:
        excel = start_excel()
        if excel is None:
            return False
        
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
            kill_excel()  # On s'assure de bien nettoyer
        except:
            pass

def inject_modules():
    log_file = setup_logging()
    logging.info(f"Début de l'injection des modules. Log file: {log_file}")
    
    if not check_vba_access():
        logging.error("Accès VBA non disponible. Arrêt du processus.")
        return
    
    excel = None
    workbook = None
    excel_file = os.path.abspath("Addin Elyse Energy.xlsm")
    
    try:
        logging.info(f"Fichier Excel cible: {excel_file}")
        
        if not os.path.exists(excel_file):
            logging.error(f"Le fichier {excel_file} n'existe pas!")
            return
        
        # Démarrer une nouvelle instance propre
        excel = start_excel()
        if excel is None:
            return
            
        logging.info(f"Ouverture du fichier {excel_file}...")
        workbook = excel.Workbooks.Open(excel_file)
            
        vba_project = workbook.VBProject
        logging.info("VBA Project accessible")
        
        # Lister tous les composants actuels
        current_components = [comp.Name for comp in vba_project.VBComponents]
        logging.info(f"Composants actuels: {', '.join(current_components)}")
        
        # Traiter tous les modules
        module_files = [f for f in os.listdir('.') if f.endswith(('.bas', '.cls', '.frm'))]
        logging.info(f"\nModules trouvés: {', '.join(module_files)}")
        
        for file in module_files:
            module_name = os.path.splitext(file)[0]
            module_type = get_module_type(module_name)
            
            try:
                if module_name in current_components:
                    vba_project.VBComponents.Remove(vba_project.VBComponents(module_name))
                    logging.info(f"Ancien module {module_name} supprimé")
                
                if file.endswith('.frm'):
                    frm_path = os.path.abspath(file)
                    imported = vba_project.VBComponents.Import(frm_path)
                else:
                    content = read_file_with_fallback_encoding(file)
                    if module_type == 100:
                        content = process_workbook_content(content)
                    
                    new_module = vba_project.VBComponents.Add(module_type)
                    new_module.Name = module_name
                    new_module.CodeModule.AddFromString(content)
                
                logging.info(f"Module {module_name} injecté avec succès")
            except Exception as e:
                logging.error(f"Erreur lors de l'injection de {module_name}: {str(e)}")
                logging.error(traceback.format_exc())
        
        # Sauvegarder
        logging.info("\nSauvegarde des modifications...")
        workbook.Save()
        time.sleep(1)
        
        # Fermer Excel
        workbook.Close(SaveChanges=True)
        excel.Quit()
        time.sleep(1)
        
        # Réouvrir en mode visible
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True
        workbook = excel.Workbooks.Open(excel_file)
        logging.info("Fichier rouvert avec succès en mode visible")
        
    except Exception as e:
        logging.error(f"Erreur générale: {str(e)}")
        logging.error(f"Stacktrace complet:\n{traceback.format_exc()}")
    finally:
        try:
            if workbook is not None and excel is not None:
                workbook.Close(SaveChanges=True)
            if excel is not None:
                excel.Quit()
        except:
            pass
        kill_excel()  # On s'assure de bien nettoyer à la fin

if __name__ == "__main__":
    try:
        inject_modules()
    except Exception as e:
        print(f"Erreur fatale: {str(e)}")
        print(traceback.format_exc())
        sys.exit(1)
