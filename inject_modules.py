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

def inject_modules():
    excel_file = os.path.abspath("Addin Elyse Energy.xlsm")
    kill_excel()
    
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
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
                
                # Lire le contenu du fichier
                with open(file, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                # Traiter le contenu si c'est ThisWorkbook
                if module_type == 100:
                    content = process_workbook_content(content)
                
                # Créer et remplir le nouveau module
                new_module = vba_project.VBComponents.Add(module_type)
                new_module.Name = module_name
                new_module.CodeModule.AddFromString(content)
                print(f"Module {module_name} injecté avec succès")
        
        # Sauvegarder et fermer
        workbook.Save()
        workbook.Close()
        excel.Quit()
        print("\nInjection terminée avec succès!")
        
    except Exception as e:
        print(f"Erreur générale: {str(e)}")
        try:
            workbook.Close(False)
            excel.Quit()
        except:
            pass

if __name__ == "__main__":
    inject_modules()