import os
import win32com.client
import time
import sys
import traceback
import psutil
from datetime import datetime
import logging

def setup_logging():
    """Configure le système de logging"""
    log_dir = "logs"
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
        # Lire en cp1252
        with open(file_path, 'r', encoding='cp1252') as f:
            content = f.read()
        
        # Réécrire en UTF-8
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(content)
        logging.info(f"Converti en UTF-8: {file_path}")
    except Exception as e:
        logging.error(f"Erreur lors de la conversion UTF-8 de {file_path}: {str(e)}")

def export_vba_modules(xlsm_path, export_dir="."):
    log_file = setup_logging()
    logging.info(f"Début de l'export des modules. Log file: {log_file}")
    
    if not check_vba_access():
        logging.error("Accès VBA non disponible. Arrêt du processus.")
        return

    kill_excel()  # On s'assure qu'aucune instance n'est en cours
    
    excel = None
    workbook = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        # Convertir en chemin absolu et échapper les espaces
        abs_path = os.path.abspath(xlsm_path).replace('"', '')
        logging.info(f"Ouverture du fichier: {abs_path}")
        workbook = excel.Workbooks.Open(abs_path)

        vba_project = workbook.VBProject
        for component in vba_project.VBComponents:
            name = component.Name
            type_ = component.Type
            # 1 = Module, 2 = Class, 3 = Form, 100 = Document (ThisWorkbook, Feuil1, etc.)
            if type_ == 1:
                ext = ".bas"
            elif type_ == 2:
                ext = ".cls"
            elif type_ == 3:
                ext = ".frm"
            elif type_ == 100:
                ext = ".cls"  # Les objets ThisWorkbook/FeuilX sont des classes
            else:
                ext = ".txt"  # fallback

            # Créer le chemin d'export et s'assurer qu'il est valide
            export_path = os.path.join(os.path.abspath(export_dir), f"{name}{ext}")
            export_path = export_path.replace('"', '')
            os.makedirs(os.path.dirname(export_path), exist_ok=True)
            
            # S'il existe déjà, on l'écrase
            if os.path.exists(export_path):
                os.remove(export_path)
                
            logging.info(f"Export de {name} vers {export_path}")
            component.Export(export_path)
            logging.info(f"Exporté: {export_path}")
            
            # Convertir en UTF-8 si c'est un fichier texte
            if ext in ['.bas', '.cls', '.frm', '.txt']:
                convert_to_utf8(export_path)

        workbook.Close(SaveChanges=False)
        excel.Quit()
        logging.info("Export terminé avec succès.")
    except Exception as e:
        logging.error(f"Erreur: {e}")
        logging.error(traceback.format_exc())
        if workbook:
            workbook.Close(SaveChanges=False)
        if excel:
            excel.Quit()
        sys.exit(1)
    finally:
        kill_excel()  # On s'assure de bien nettoyer à la fin

def export_modules(export_dir="."):
    log_file = setup_logging()
    logging.info(f"Début de l'export des modules. Log file: {log_file}")
    
    excel = None
    workbook = None
    excel_file = r"C:\Users\JulienFernandez\OneDrive\Coding\_Projets de code\Addin Elyse Energy - fonctionnel - v2025.06.06.xlsm"
    
    try:
        # Convertir le dossier d'export en chemin absolu
        export_dir = os.path.abspath(export_dir)
        
        logging.info(f"Fichier Excel source: {excel_file}")
        logging.info(f"Dossier d'export: {export_dir}")
        
        if not os.path.exists(excel_file):
            logging.error(f"Le fichier {excel_file} n'existe pas!")
            return
            
        # S'assurer que le dossier d'export existe
        os.makedirs(export_dir, exist_ok=True)
        
        # Créer une instance Excel
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        workbook = excel.Workbooks.Open(excel_file)
        
        vba_project = workbook.VBProject
        
        # Exporter tous les composants
        for comp in vba_project.VBComponents:
            try:
                if comp.Type in [1, 2, 3]:  # Module standard, Classe, Formulaire
                    extension = {1: "bas", 2: "cls", 3: "frm"}[comp.Type]
                    export_filename = os.path.join(export_dir, f"{comp.Name}.{extension}")
                    comp.Export(export_filename)
                    logging.info(f"Module exporté: {export_filename}")
                else:
                    # Pour les modules Document (ThisWorkbook, feuilles)
                    if comp.Type == 100:
                        export_filename = os.path.join(export_dir, f"{comp.Name}.cls")
                        comp.Export(export_filename)
                        logging.info(f"Module Document exporté: {export_filename}")
            
            except Exception as e:
                logging.error(f"Erreur lors de l'export de {comp.Name}: {str(e)}")
                logging.error(traceback.format_exc())
        
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
        export_modules()
    except Exception as e:
        print(f"Erreur fatale: {str(e)}")
        print(traceback.format_exc())
        sys.exit(1) 