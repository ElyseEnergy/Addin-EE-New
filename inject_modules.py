import os
import win32com.client
import time
import psutil
import logging
import sys
from datetime import datetime
import traceback

# Configuration du logging
def setup_logging():
    """Configure le système de logging"""
    log_dir = "logs"
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = os.path.join(log_dir, f"inject_modules_{timestamp}.log")
    
    # Configuration du logging vers fichier et console
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

def get_module_type(file_name):
    """Détermine le type de module basé sur le nom du fichier"""
    logging.debug(f"Détermination du type de module pour: {file_name}")
    if file_name.startswith('ThisWorkbook'):
        logging.debug(f"{file_name} -> Type: Document (100)")
        return 100  # vbext_ct_Document
    elif file_name.startswith('UserForm') or file_name.startswith('frm'):
        logging.debug(f"{file_name} -> Type: MSForm (3)")
        return 3    # vbext_ct_MSForm
    else:
        logging.debug(f"{file_name} -> Type: StdModule (1)")
        return 1    # vbext_ct_StdModule

def process_workbook_content(content):
    """Traite le contenu d'un module Workbook"""
    logging.debug("Traitement du contenu du module Workbook")
    if not content.startswith('VERSION 1.0 CLASS'):
        logging.debug("Ajout de l'en-tête CLASS manquant")
        content = 'VERSION 1.0 CLASS\nBEGIN\n  MultiUse = -1  \'True\nEND\n' + content
    return content

def read_file_with_fallback_encoding(file_path):
    """Essaie de lire le fichier avec différents encodages"""
    logging.info(f"Lecture du fichier: {file_path}")
    encodings = [
        'utf-8', 'cp1252', 'iso-8859-1', 'latin1', 'utf-8-sig',
        'cp850', 'cp437', 'ascii', 'mac-roman', 'cp858',
        'iso-8859-15', 'cp1254', 'cp1256', 'cp1257'
    ]
    last_error = None
    file_size = os.path.getsize(file_path)
    logging.debug(f"Taille du fichier: {file_size} octets")
    
    # Lire le contenu binaire du fichier
    with open(file_path, 'rb') as f:
        raw_data = f.read()
        
    # Vérifier si le fichier commence par un BOM
    if raw_data.startswith(b'\xef\xbb\xbf'):
        logging.debug("BOM UTF-8 détecté")
        try:
            content = raw_data.decode('utf-8-sig')
            logging.info("Fichier lu avec succès using utf-8-sig (BOM)")
            return content
        except Exception as e:
            logging.warning(f"Échec de lecture avec BOM UTF-8: {str(e)}")
    
    # D'abord essayer de détecter l'encodage
    try:
        import chardet
        detected = chardet.detect(raw_data)
        logging.debug(f"Encodage détecté: {detected}")
        if detected['confidence'] > 0.8:
            try:
                content = raw_data.decode(detected['encoding'])
                logging.info(f"Fichier lu avec succès using {detected['encoding']}")
                return content
            except Exception as e:
                logging.warning(f"Échec de lecture avec l'encodage détecté: {str(e)}")
    except ImportError:
        logging.warning("Module chardet non disponible")
    
    # Essayer les encodages connus
    for encoding in encodings:
        try:
            logging.debug(f"Tentative de lecture avec l'encodage: {encoding}")
            content = raw_data.decode(encoding)
            
            # Vérification plus souple des caractères invalides
            invalid_chars = ['', '□', '¿', '\x00']
            invalid_count = sum(1 for c in invalid_chars if c in content)
            if invalid_count > len(content) * 0.01:  # Plus de 1% de caractères invalides
                logging.debug(f"Trop de caractères invalides détectés avec {encoding} ({invalid_count} caractères)")
                continue
                
            logging.info(f"Fichier lu avec succès using {encoding}")
            return content
        except UnicodeDecodeError as e:
            logging.debug(f"Échec avec {encoding}: {str(e)}")
            last_error = e
        except Exception as e:
            logging.error(f"Erreur inattendue avec {encoding}: {str(e)}")
            last_error = e
    
    # Dernier recours : essayer de lire en ignorant les erreurs
    try:
        logging.warning("Tentative de lecture en mode permissif (ignore errors)")
        content = raw_data.decode('utf-8', errors='ignore')
        if not any(c in content for c in ['', '□', '¿', '\x00']):
            logging.info("Fichier lu avec succès en mode permissif")
            return content
    except Exception as e:
        logging.error(f"Échec de la lecture en mode permissif: {str(e)}")
        last_error = e
            
    error_msg = f"Impossible de lire le fichier {file_path} avec les encodages disponibles. Dernière erreur: {str(last_error)}"
    logging.error(error_msg)
    raise ValueError(error_msg)

def check_vba_access():
    """Vérifie si l'accès au VBA Project est activé"""
    logging.info("Vérification de l'accès au VBA Project")
    excel = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        logging.debug("Excel démarré en mode invisible")
        
        wb = excel.Workbooks.Add()
        logging.debug("Nouveau classeur créé")
        
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
        if excel:
            try:
                excel.Quit()
                logging.debug("Excel fermé")
            except Exception as e:
                logging.warning(f"Erreur lors de la fermeture d'Excel: {str(e)}")

def get_open_workbook(excel_file):
    """Tente de récupérer le classeur s'il est déjà ouvert"""
    logging.info("Recherche du classeur déjà ouvert...")
    try:
        excel = win32com.client.GetObject(Class="Excel.Application")
        for wb in excel.Workbooks:
            if os.path.abspath(wb.FullName) == os.path.abspath(excel_file):
                logging.info("Classeur trouvé déjà ouvert!")
                return excel, wb
        logging.info("Classeur non trouvé parmi les fichiers ouverts")
        return None, None
    except Exception as e:
        logging.debug(f"Aucune instance d'Excel en cours d'exécution: {str(e)}")
        return None, None

def inject_modules():
    log_file = setup_logging()
    logging.info(f"Début de l'injection des modules. Log file: {log_file}")
    
    if not check_vba_access():
        logging.error("Accès VBA non disponible. Arrêt du processus.")
        return
    
    excel = None
    workbook = None
    need_to_reopen = False
    excel_file = os.path.abspath("Addin Elyse Energy.xlsm")
    
    try:
        logging.info(f"Fichier Excel cible: {excel_file}")
        
        if not os.path.exists(excel_file):
            logging.error(f"Le fichier {excel_file} n'existe pas!")
            return
            
        # Essayer de récupérer le classeur s'il est déjà ouvert
        excel, workbook = get_open_workbook(excel_file)
        
        if excel is None:
            logging.info("Démarrage d'une nouvelle instance d'Excel...")
            kill_excel()
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            logging.debug("Excel démarré en mode invisible, alertes désactivées")
            
            logging.info(f"Ouverture du fichier {excel_file}...")
            workbook = excel.Workbooks.Open(excel_file)
            need_to_reopen = True
        else:
            logging.info("Utilisation de l'instance Excel existante")
            excel.DisplayAlerts = False
            
        vba_project = workbook.VBProject
        logging.info("VBA Project accessible")
        
        # Lister tous les composants actuels
        current_components = [comp.Name for comp in vba_project.VBComponents]
        logging.info(f"Composants actuels: {', '.join(current_components)}")
        
        # D'abord, traiter les UserForms (fichiers .frm)
        frm_files = [f for f in os.listdir('.') if f.endswith('.frm')]
        logging.info(f"UserForms trouvés: {', '.join(frm_files)}")
        
        for file in frm_files:
            module_name = os.path.splitext(file)[0]
            logging.info(f"\nTraitement du UserForm: {module_name}")
            
            try:
                if module_name in current_components:
                    vba_project.VBComponents.Remove(vba_project.VBComponents(module_name))
                    logging.info(f"Ancien UserForm {module_name} supprimé")
            except Exception as e:
                logging.warning(f"Erreur lors de la suppression de {module_name}: {str(e)}")
            
            try:
                frm_path = os.path.abspath(file)
                logging.debug(f"Import de {frm_path}")
                imported = vba_project.VBComponents.Import(frm_path)
                logging.info(f"UserForm {module_name} importé avec succès")
            except Exception as e:
                logging.error(f"Erreur lors de l'import du UserForm {module_name}: {str(e)}")
                logging.error(f"Stacktrace: {traceback.format_exc()}")
        
        # Traiter les modules standards et classes
        module_files = [f for f in os.listdir('.') if f.endswith(('.bas', '.cls'))]
        logging.info(f"\nModules trouvés: {', '.join(module_files)}")
        
        for file in module_files:
            module_name = os.path.splitext(file)[0]
            module_type = get_module_type(module_name)
            logging.info(f"\nTraitement du module: {module_name} (Type: {module_type})")
            
            if module_type == 3:
                logging.debug(f"Module {module_name} est un UserForm, déjà traité")
                continue
            
            try:
                if module_name in current_components:
                    vba_project.VBComponents.Remove(vba_project.VBComponents(module_name))
                    logging.info(f"Ancien module {module_name} supprimé")
            except Exception as e:
                logging.warning(f"Erreur lors de la suppression de {module_name}: {str(e)}")
            
            try:
                content = read_file_with_fallback_encoding(file)
                logging.debug(f"Contenu lu: {len(content)} caractères")
                
                if module_type == 100:
                    content = process_workbook_content(content)
                    logging.debug("Contenu ThisWorkbook traité")
                
                new_module = vba_project.VBComponents.Add(module_type)
                new_module.Name = module_name
                
                logging.debug(f"Injection du code dans {module_name}")
                new_module.CodeModule.AddFromString(content)
                logging.info(f"Module {module_name} injecté avec succès")
                
            except Exception as e:
                logging.error(f"Erreur lors de l'injection de {module_name}: {str(e)}")
                logging.error(f"Stacktrace: {traceback.format_exc()}")
        
        # Sauvegarder
        logging.info("\nSauvegarde des modifications...")
        workbook.Save()
        time.sleep(1)
        
    except Exception as e:
        logging.error(f"Erreur générale: {str(e)}")
        logging.error(f"Stacktrace complet:\n{traceback.format_exc()}")
    finally:
        # Nettoyage et réouverture
        logging.info("Finalisation...")
        try:
            if need_to_reopen or not workbook.Windows(1).Visible:
                # Si on a ouvert une nouvelle instance ou si Excel était invisible
                if workbook is not None:
                    workbook.Close(SaveChanges=True)
                    logging.debug("Workbook fermé")
                if excel is not None:
                    excel.Quit()
                    logging.debug("Excel fermé")
                time.sleep(1)
                
                # Réouvrir en mode visible
                logging.info("Réouverture du fichier en mode visible...")
                excel = win32com.client.Dispatch("Excel.Application")
                excel.Visible = True
                excel.DisplayAlerts = True
                workbook = excel.Workbooks.Open(excel_file)
                logging.info("Fichier rouvert avec succès en mode visible")
            else:
                # Si Excel était déjà ouvert et visible, juste mettre à jour l'affichage
                logging.info("Mise à jour de l'affichage...")
                excel.Visible = True
                excel.DisplayAlerts = True
                workbook.Windows(1).Visible = True
        except Exception as e:
            logging.error(f"Erreur lors de la finalisation: {str(e)}")
            logging.error(traceback.format_exc())

if __name__ == "__main__":
    try:
        inject_modules()
    except Exception as e:
        print(f"Erreur fatale: {str(e)}")
        print(traceback.format_exc())
        sys.exit(1)
