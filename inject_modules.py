import os
import win32com.client
import time
import psutil

def kill_excel():
    """Ferme toutes les instances d'Excel"""
    for proc in psutil.process_iter(['name']):
        if proc.info['name'] == 'EXCEL.EXE':
            proc.kill()
    time.sleep(2)  # Attendre que Excel se ferme complètement

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
    encodings = ['utf-8', 'cp1252', 'iso-8859-1', 'latin1', 'utf-8-sig']
    last_error = None
    
    # D'abord essayer de détecter l'encodage
    try:
        import chardet
        with open(file_path, 'rb') as f:
            raw_data = f.read()
        detected = chardet.detect(raw_data)
        if detected['confidence'] > 0.8:
            try:
                with open(file_path, 'r', encoding=detected['encoding']) as f:
                    return f.read()
            except:
                pass
    except ImportError:
        pass  # chardet n'est pas installé
    
    # Essayer les encodages connus
    for encoding in encodings:
        try:
            with open(file_path, 'r', encoding=encoding) as f:
                content = f.read()
                # Vérifier si le contenu semble valide
                if any(invalid_char in content for invalid_char in ['�', '□', '¿']):
                    continue
                return content
        except UnicodeDecodeError as e:
            last_error = e
            continue
        except Exception as e:
            last_error = e
            continue
            
    raise ValueError(f"Impossible de lire le fichier {file_path} avec les encodages disponibles. Dernière erreur: {str(last_error)}")

def check_vba_access():
    """Vérifie si l'accès au VBA Project est activé"""
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        wb = excel.Workbooks.Add()
        try:
            _ = wb.VBProject
            return True
        except:
            print("ERREUR: L'accès au VBA Project n'est pas activé dans Excel.")
            print("Pour activer l'accès:")
            print("1. Ouvrez Excel")
            print("2. Allez dans Fichier > Options > Centre de gestion de la confidentialité > Paramètres du Centre de gestion de la confidentialité")
            print("3. Cliquez sur 'Paramètres des macros'")
            print("4. Cochez 'Faire confiance à l'accès au modèle d'objet des projets VBA'")
            print("5. Cliquez sur OK et redémarrez Excel")
            return False
        finally:
            try:
                wb.Close(False)
                excel.Quit()
            except:
                pass
    except Exception as e:
        print(f"Erreur lors de la vérification de l'accès VBA: {str(e)}")
        return False

def inject_modules():
    excel = None
    workbook = None
    try:
        excel_file = os.path.abspath("Addin Elyse Energy.xlsm")
        kill_excel()
        
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        print(f"Ouverture du fichier {excel_file}...")
        workbook = excel.Workbooks.Open(excel_file)
        vba_project = workbook.VBProject
        
        # D'abord, traiter les UserForms (fichiers .frm)
        for file in os.listdir('.'):
            if file.endswith('.frm'):
                module_name = os.path.splitext(file)[0]
                
                # Supprimer l'ancien UserForm s'il existe
                try:
                    vba_project.VBComponents.Remove(vba_project.VBComponents(module_name))
                    print(f"Ancien UserForm {module_name} supprimé")
                except:
                    pass
                
                # Importer le UserForm via le fichier .frm
                try:
                    imported = vba_project.VBComponents.Import(os.path.abspath(file))
                    print(f"UserForm {module_name} importé avec succès via {file}")
                except Exception as e:
                    print(f"Erreur lors de l'import du UserForm {module_name}: {str(e)}")
        
        # Ensuite, traiter les modules standards et classes
        for file in os.listdir('.'):
            if file.endswith('.bas') or file.endswith('.cls'):
                module_name = os.path.splitext(file)[0]
                module_type = get_module_type(module_name)
                
                # Si c'est un UserForm, on l'a déjà traité
                if module_type == 3:
                    continue
                
                # Supprimer l'ancien module s'il existe
                try:
                    vba_project.VBComponents.Remove(vba_project.VBComponents(module_name))
                    print(f"Ancien module {module_name} supprimé")
                except:
                    pass
                
                # Lire le contenu du fichier avec différents encodages
                try:
                    content = read_file_with_fallback_encoding(file)
                except ValueError as e:
                    print(f"Erreur lors de la lecture du fichier {file}: {str(e)}")
                    continue
                
                # Traiter le contenu si c'est ThisWorkbook
                if module_type == 100:
                    content = process_workbook_content(content)
                
                # Créer et remplir le nouveau module
                new_module = vba_project.VBComponents.Add(module_type)
                new_module.Name = module_name
                new_module.CodeModule.AddFromString(content)
                print(f"Module {module_name} injecté avec succès")
        
        # Sauvegarder et fermer
        print("\nSauvegarde des modifications...")
        workbook.Save()
        time.sleep(1)  # Attendre que la sauvegarde soit terminée
        
        print("Fermeture du classeur...")
        workbook.Close(SaveChanges=True)  # Forcer la sauvegarde à la fermeture
        time.sleep(1)  # Attendre la fermeture
        
        print("Fermeture d'Excel...")
        excel.Quit()
        time.sleep(1)  # Attendre la fermeture d'Excel
        
        print("\nInjection terminée avec succès!")
        
    except Exception as e:
        print(f"Erreur générale: {str(e)}")
    finally:
        # S'assurer que tout est bien fermé
        try:
            if workbook is not None:
                workbook.Close(SaveChanges=True)
        except:
            pass
            
        try:
            if excel is not None:
                excel.Quit()
        except:
            pass
            
        # Forcer la fermeture d'Excel si nécessaire
        kill_excel()

if __name__ == "__main__":
    inject_modules()