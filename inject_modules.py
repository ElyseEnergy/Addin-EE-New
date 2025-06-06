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

def setup_logging():
    """Configure le système de logging"""
    log_dir = "logs"
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

def prepare_temp_file(src_file, temp_dir):
    """Copie et convertit un fichier en CP1252 dans le dossier temporaire"""
    temp_file = os.path.join(temp_dir, os.path.basename(src_file))
    
    try:
        # D'abord, copier le fichier
        shutil.copy2(src_file, temp_file)
        
        # Essayer de lire en UTF-8
        try:
            with open(src_file, 'r', encoding='utf-8') as f:
                content = f.read()
            # Si on arrive ici, c'est que le fichier est en UTF-8
            logging.info(f"Fichier {src_file} détecté en UTF-8, conversion en CP1252...")
            # Réécrire le fichier temporaire en CP1252
            with open(temp_file, 'w', encoding='cp1252') as f:
                f.write(content)
            logging.info(f"Converti en CP1252 dans {temp_file}")
        except UnicodeDecodeError:
            # Si on ne peut pas lire en UTF-8, c'est probablement déjà du CP1252
            logging.info(f"Fichier {src_file} déjà en CP1252")
        
        return temp_file
    except Exception as e:
        logging.error(f"Erreur lors de la préparation de {src_file}: {str(e)}")
        return None

def get_excel_instance(target_file=None):
    """Récupère une instance Excel existante avec notre fichier ou en crée une nouvelle"""
    try:
        # Essayer de se connecter à une instance existante
        excel = win32com.client.GetObject(Class="Excel.Application")
        
        # Si on a un fichier cible, vérifier s'il est déjà ouvert
        if target_file:
            target_file = os.path.abspath(target_file)
            for wb in excel.Workbooks:
                if os.path.abspath(wb.FullName) == target_file:
                    logging.info(f"Fichier {target_file} déjà ouvert dans Excel")
                    return excel, wb
        
        # Si on arrive ici, soit pas d'instance Excel, soit fichier non ouvert
        raise Exception("Pas d'instance Excel utilisable")
        
    except:
        # Créer une nouvelle instance
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        logging.info("Nouvelle instance Excel créée")
        return excel, None

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

def inject_modules():
    log_file = setup_logging()
    logging.info(f"Début de l'injection des modules. Log file: {log_file}")
    
    if not check_vba_access():
        logging.error("Accès VBA non disponible. Arrêt du processus.")
        return
    
    excel = None
    workbook = None
    excel_file = os.path.abspath("Addin Elyse Energy.xlsm")
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
        if not workbook:
            workbook = excel.Workbooks.Open(excel_file)
        
        vba_project = workbook.VBProject
        
        # Lister tous les composants actuels
        current_components = {comp.Name: comp.Type for comp in vba_project.VBComponents}
        logging.info(f"Composants actuels: {', '.join(current_components.keys())}")
        
        # Supprimer tous les modules .bas d'un coup
        remove_all_bas_modules(vba_project)
        
        # Traiter tous les modules
        module_files = [f for f in os.listdir('.') if f.endswith(('.bas', '.cls', '.frm'))]
        logging.info(f"Modules trouvés: {', '.join(module_files)}")
        
        for file in module_files:
            module_name = os.path.splitext(file)[0]
            
            try:
                # Préparer le fichier temporaire
                temp_file = prepare_temp_file(file, temp_dir)
                if not temp_file:
                    continue
                
                # Vérifier si c'est un module Document (Sheet, ThisWorkbook)
                is_document = (module_name in current_components and 
                             current_components[module_name] == 100)  # vbext_ct_Document
                
                if not is_document and module_name in current_components:
                    # On peut supprimer les modules non-Document
                    vba_project.VBComponents.Remove(vba_project.VBComponents(module_name))
                    logging.info(f"Ancien module {module_name} supprimé")
                
                if is_document:
                    # Pour les modules Document, on met à jour le code
                    with open(temp_file, 'r', encoding='cp1252') as f:
                        new_code = f.read()
                    # Enlever l'en-tête CLASS si présent
                    if new_code.startswith('VERSION 1.0 CLASS'):
                        new_code = new_code.split('END\n', 1)[1]
                    vba_project.VBComponents(module_name).CodeModule.DeleteLines(
                        1, vba_project.VBComponents(module_name).CodeModule.CountOfLines)
                    vba_project.VBComponents(module_name).CodeModule.AddFromString(new_code)
                    logging.info(f"Code du module Document {module_name} mis à jour")
                else:
                    # Pour les autres modules, import normal
                    vba_project.VBComponents.Import(temp_file)
                    logging.info(f"Module {module_name} injecté avec succès")
                
            except Exception as e:
                logging.error(f"Erreur lors de l'injection de {module_name}: {str(e)}")
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
        logging.error(f"Stacktrace complet:\n{traceback.format_exc()}")
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
