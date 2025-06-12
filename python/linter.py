import os
import re
import xml.etree.ElementTree as ET
from collections import defaultdict
import datetime
import sys
from typing import Dict, List, Tuple

# --- Configuration ---
VBA_SRC_DIR = "./"
XML_RIBBON_FILE = "customUI.xml"
LOG_FILE = "python/logs/linter.log"
CRITICAL = "CRITICAL"
WARNING = "WARNING"

# --- Regex et Constantes ---
PASCAL_CASE_RE = re.compile(r"^[A-Z][a-zA-Z0-9]*$")
ALL_CAPS_RE = re.compile(r"^[A-Z0-9_]+$")
PROC_DEF_RE = re.compile(r"^\s*(?:Public|Private|Friend)?\s*(?:Sub|Function)\s+([a-zA-Z0-9_]+)\s*\((.*?)\)", re.IGNORECASE)
CONST_RE = re.compile(r"^\s*(?:Public|Private)?\s*Const\s+([a-zA-Z0-9_]+)", re.IGNORECASE)
ERROR_HANDLER_RE = re.compile(r"On Error GoTo \w+", re.IGNORECASE)
CALLBACK_CANDIDATE_RE = re.compile(r"^\s*Public\s+(Sub|Function)\s+([a-zA-Z0-9_]+)\(.*\)", re.IGNORECASE)
ATTRIBUTE_RE = re.compile(r'Attribute VB_Name = "([^"]+)"')
PROC_CALL_RE = re.compile(r'(?:Call\s+)?([a-zA-Z0-9_\.]+)\s*\((.*?)\)', re.IGNORECASE)
SYS_LOGGER_RE = re.compile(r'Log\s+"[^"]+"\s*,\s*"[^"]+"\s*,\s*[A-Z_]+_LEVEL\s*,\s*PROC_NAME\s*,\s*MODULE_NAME', re.IGNORECASE)
ERROR_HANDLER_BLOCK_RE = re.compile(r'ErrorHandler:', re.IGNORECASE)
OPTION_EXPLICIT_RE = re.compile(r'^\s*Option\s+Explicit\s*$', re.IGNORECASE)

# Liste des attributs XML qui définissent un callback
CALLBACK_ATTRIBUTES = [
    'onAction', 'getEnabled', 'getVisible', 'getLabel', 'getText', 
    'getSupertip', 'onLoad', 'getScreentip', 'getSize', 'getImage'
]

# Exceptions pour la gestion d'erreur (fonctions courtes/wrappers)
ERROR_HANDLER_EXCEPTIONS = ["Main", "Click", "Callback"]

# --- Fonctions d'analyse (précédemment dans fix_argument_errors.py) ---

def parse_procedure_signature(signature_line: str) -> Tuple[str, List[str], int, int, bool]:
    """Parse une signature de procédure VBA pour extraire le nom et les paramètres.
    Retourne (nom_procedure, liste_parametres, min_args, max_args, est_une_fonction)
    """
    signature = signature_line.replace(" _\n", " ").strip()
    
    match = re.search(r"(?:Sub|Function)\s+([a-zA-Z0-9_]+)", signature, re.IGNORECASE)
    if not match:
        return "", [], 0, 0, False
    name = match.group(1)
    
    is_function = "Function" in signature.split("(")[0]
    
    params_str_match = re.search(r"\((.*)\)", signature)
    if not params_str_match:
        return name, [], 0, 0, is_function
    params_str = params_str_match.group(1)

    if not params_str.strip():
        return name, [], 0, 0, is_function
        
    param_parts = re.split(r',\s*(?![^()]*\))', params_str)
    
    all_params = []
    min_args = 0
    max_args = 0
    
    has_param_array = False
    
    for param in param_parts:
        param_clean = param.strip()
        if not param_clean:
            continue

        param_name_match = re.search(r"(?:ByVal|ByRef|Optional|ParamArray)?\s*([a-zA-Z0-9_]+)", param_clean, re.IGNORECASE)
        if param_name_match:
            all_params.append(param_name_match.group(1))

        if "ParamArray" in param_clean:
            has_param_array = True
        elif "Optional" not in param_clean:
            min_args += 1
    
    max_args = len(all_params)
    if has_param_array:
        max_args = float('inf') # Un ParamArray peut prendre un nombre infini d'arguments

    return name, all_params, min_args, max_args, is_function

def analyze_vba_file_signatures(file_content: str) -> Dict[str, Tuple[List[str], int, int, bool, int]]:
    """Analyse un fichier VBA pour extraire les signatures de procédure.
    Retourne un dict mappant nom_procedure à (liste_parametres, min_args, max_args, est_une_fonction, numero_ligne)
    """
    procedures = {}
    lines = file_content.split('\n')
    
    proc_start_re = re.compile(r"^\s*(?:Public|Private|Friend)?\s*(?!(?:Declare)\s)(Sub|Function)\s+([a-zA-Z0-9_]+)", re.IGNORECASE)

    for i, line in enumerate(lines):
        if proc_start_re.match(line):
            # Reconstituer la signature complète qui peut s'étendre sur plusieurs lignes
            signature = ""
            for j in range(i, len(lines)):
                line_part = lines[j].strip()
                signature += line_part.rstrip(' _')
                if not line_part.endswith("_"):
                    break
            
            try:
                name, params, min_args, max_args, is_func = parse_procedure_signature(signature)
                if name:
                    procedures[name.lower()] = (params, min_args, max_args, is_func, i + 1)
            except Exception:
                pass # Ignorer les erreurs de parsing

    return procedures

# --- Fonctions de Logging ---
def setup_logging():
    os.makedirs(os.path.dirname(LOG_FILE), exist_ok=True)
    with open(LOG_FILE, "w", encoding="utf-8") as f:
        f.write(f"--- Linter Log - {datetime.datetime.now()} ---\n\n")

def log_message(message, level):
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(f"[{level}] {message}\n")
    print(f"[{level}] {message}")

# --- Fonctions de Parsing ---
def find_vba_files(directory):
    vba_files = []
    for root, _, files in os.walk(directory):
        if "python" in root or ".git" in root:
            continue
        for file in files:
            if file.lower().endswith(('.bas', '.cls', '.frm')):
                vba_files.append(os.path.join(root, file))
    return vba_files

def parse_procedure_calls(lines: List[str], start_line: int, end_line: int) -> List[Tuple[str, int, int]]:
    """Analyse un bloc de code pour trouver les appels de procédure."""
    calls = []
    # Regex simplifié pour trouver les appels de fonction ou de sub
    # `(?:Call\s+)?` : `Call` est optionnel
    # `([a-zA-Z0-9_\.]+)`: Capture le nom de la procédure (y compris les appels de membre comme `Utilities.Sanitize`)
    # `(?:\(([^)]*)\))?`: Capture optionnellement les arguments entre parenthèses
    proc_call_re = re.compile(r'(?:Call\s+)?([a-zA-Z0-9_\.]+)(?:\(([^)]*)\))?', re.IGNORECASE)

    proc_def_re = re.compile(r"^\s*(?:Public|Private|Friend)?\s*(?:Sub|Function)", re.IGNORECASE)

    for i in range(start_line, end_line):
        line = lines[i].strip()

        # Ignorer les lignes de déclaration pour ne pas les compter comme des appels
        if proc_def_re.match(line):
            continue

        for match in proc_call_re.finditer(line):
            proc_name = match.group(1)
            args_str = match.group(2)

            # Ne pas compter les assignations de variable comme des appels
            # ex: `maVar = MaFonction(x)` est un appel, mais `maVar = x` ne l'est pas
            if not match.group(2) and "=" in line and proc_name.lower() not in ["select", "case"]:
                 # S'il n'y a pas de parenthèses, ce n'est probablement pas un appel de fonction,
                 # sauf si c'est une Sub appelée sans `Call`. C'est ambigu.
                 # Pour l'instant, on suppose que sans parenthèses, ce n'est pas un appel,
                 # sauf si le mot-clé `Call` est présent.
                 if "call " not in line.lower():
                     continue

            if args_str is None:
                # Pas de parenthèses (ex: `Call MySub` ou `MySub arg1, arg2`)
                # Compter les arguments en se basant sur les virgules après le nom
                following_code = line[match.end(1):].strip()
                if not following_code:
                    args_count = 0
                else:
                    # C'est une approximation, car `MySub "a,b", "c"` serait mal compté
                    args_count = len(following_code.split(','))
            elif args_str.strip() == "":
                 # Parenthèses vides (ex: `MyFunc()`)
                args_count = 0
            else:
                # Parenthèses avec des arguments
                args_count = len(re.split(r',(?![^"]*"(?:(?:[^"]*"){2})*[^"]*$)', args_str))

            calls.append((proc_name.lower(), args_count, i + 1))
    return calls

def check_error_handler_with_logger(lines, start_line, end_line):
    has_on_error = False
    has_handler_label = False
    has_log_call = False
    has_handle_error_call = False

    # 1. Vérifier la présence de "On Error GoTo"
    for i in range(start_line, end_line):
        line = lines[i].strip()
        if ERROR_HANDLER_RE.match(line):
            has_on_error = True
            break
    if not has_on_error:
        return False

    # 2. Chercher le bloc ErrorHandler: et les appels requis
    in_error_handler_block = False
    for i in range(start_line, end_line):
        line = lines[i].strip()
        if ERROR_HANDLER_BLOCK_RE.search(line):
            in_error_handler_block = True
        
        if in_error_handler_block:
            if "SYS_Logger.Log" in line or "Log " in line:
                has_log_call = True
            if "HandleError" in line:
                has_handle_error_call = True
    
    return has_on_error and in_error_handler_block and has_log_call and has_handle_error_call

def parse_vba_file(file_path):
    module_name_attr = None
    procedures = {}
    consts = []
    lines = []
    has_option_explicit = False
    try:
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
            lines = f.readlines()
    except Exception as e:
        log_message(f"Could not read file {file_path}: {e}", CRITICAL)
        return None

    if not lines:
        return {'name': os.path.basename(file_path), 'procedures': {}, 'consts': [], 'has_option_explicit': False}

    match = ATTRIBUTE_RE.match(lines[0])
    if match:
        module_name_attr = match.group(1)
    
    # Cherche Option Explicit dans les premières lignes jusqu'à la première déclaration
    for line in lines:
        if OPTION_EXPLICIT_RE.match(line):
            has_option_explicit = True
            break
        # Si on trouve une déclaration avant Option Explicit, on arrête la recherche
        if PROC_DEF_RE.match(line) or CONST_RE.match(line):
            break

    current_proc = None
    proc_start_line = 0
    
    for i, line in enumerate(lines):
        proc_match = PROC_DEF_RE.match(line)
        if proc_match:
            if current_proc:
                # Analyse la procédure précédente
                procedures[current_proc.name] = {
                    'info': current_proc,
                    'line': proc_start_line + 1,
                    'calls': parse_procedure_calls(lines, proc_start_line, i),
                    'has_proper_error_handler': check_error_handler_with_logger(lines, proc_start_line, i)
                }
            
            proc_name = proc_match.group(1)
            params_str = proc_match.group(2)
            current_proc = ProcedureInfo(proc_name, params_str, module_name_attr, i + 1)
            current_proc.is_public = "Public" in line
            proc_start_line = i
        
        const_match = CONST_RE.match(line)
        if const_match:
            consts.append({'name': const_match.group(1), 'line': i + 1})
            
    # Analyse la dernière procédure
    if current_proc:
        procedures[current_proc.name] = {
            'info': current_proc,
            'line': proc_start_line + 1,
            'calls': parse_procedure_calls(lines, proc_start_line, len(lines)),
            'has_proper_error_handler': check_error_handler_with_logger(lines, proc_start_line, len(lines))
        }
    
    return {
        'name': module_name_attr or os.path.basename(file_path),
        'procedures': procedures,
        'consts': consts,
        'has_option_explicit': has_option_explicit,
        'lines': lines
    }

def get_xml_callbacks(xml_file):
    callbacks = set()
    try:
        tree = ET.parse(xml_file)
        root = tree.getroot()
        # Special case for onLoad on the root element
        if 'onLoad' in root.attrib:
            callbacks.add(root.get('onLoad'))
        
        for attr in CALLBACK_ATTRIBUTES:
            for elem in root.findall(f".//*[@{attr}]"):
                callbacks.add(elem.get(attr))
    except (ET.ParseError, FileNotFoundError) as e:
        log_message(f"Could not parse XML file {xml_file}: {e}", CRITICAL)
    return callbacks

# --- Linter Principal ---

class VBALinter:
    def __init__(self):
        self.public_procedures = defaultdict(list)
        self.xml_callbacks = get_xml_callbacks(XML_RIBBON_FILE)
        self.all_procedures = {} # Dictionnaire global des signatures

    def analyze_all_files(self, vba_files: List[str]):
        """Première passe pour collecter toutes les signatures de procédure."""
        for file_path in vba_files:
            try:
                with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                    content = f.read()
                
                # Récupérer le nom du module de l'attribut VB_Name
                module_name_match = ATTRIBUTE_RE.search(content)
                module_name = module_name_match.group(1) if module_name_match else os.path.basename(file_path)

                file_procs = analyze_vba_file_signatures(content)
                for name, sig in file_procs.items():
                    # Ajouter le nom du module à la signature
                    self.all_procedures[name] = (sig[0], sig[1], sig[2], sig[3], sig[4], module_name)
            except Exception as e:
                log_message(f"Error analyzing signatures in {file_path}: {e}", CRITICAL)

    def _check_file_content(self, file_path: str, content: str):
        lines = content.split('\n')
        filename = os.path.basename(file_path)

        # 1. Vérifier la présence de "Attribute VB_Name"
        if not ATTRIBUTE_RE.search(content):
            log_message(f"Dans '{filename}': L'attribut 'VB_Name' est manquant en début de fichier.", CRITICAL)

        # 2. Vérifier "Option Explicit"
        if not OPTION_EXPLICIT_RE.search(content):
            log_message(f"Dans '{filename}': 'Option Explicit' est manquant ou mal placé.", WARNING)

        # 3. Parcourir les procédures pour les vérifications détaillées
        proc_start_re = re.compile(r"^\s*(?:Public|Private|Friend)?\s*(?!(?:Declare)\s)(Sub|Function)\s+([a-zA-Z0-9_]+)", re.IGNORECASE)
        
        current_proc_name = None
        current_proc_start_line = 0
        
        for i, line in enumerate(lines):
            match = proc_start_re.match(line)
            if match:
                # Fin de la procédure précédente, on l'analyse
                if current_proc_name:
                    self._analyze_procedure_block(lines, current_proc_start_line, i, filename, current_proc_name)

                # Début d'une nouvelle procédure
                current_proc_name = match.group(2).lower()
                current_proc_start_line = i
        
        # Analyser la dernière procédure du fichier
        if current_proc_name:
            self._analyze_procedure_block(lines, current_proc_start_line, len(lines), filename, current_proc_name)


    def _analyze_procedure_block(self, lines, start_line, end_line, filename, proc_name):
        """Analyse un bloc de code correspondant à une seule procédure."""
        
        # 4. Vérifier la gestion d'erreur
        # Exclure les fonctions de test et les wrappers simples
        is_exception = any(exc.lower() in proc_name.lower() for exc in ERROR_HANDLER_EXCEPTIONS)
        
        if not is_exception and not check_error_handler_with_logger(lines, start_line, end_line):
            log_message(f"Dans '{filename}', procédure '{proc_name}' (ligne {start_line + 1}): semble ne pas avoir de gestion d'erreur complète (On Error GoTo, ErrorHandler:, Log, HandleError).", CRITICAL)

        # 5. Vérifier les appels de procédure
        calls = parse_procedure_calls(lines, start_line, end_line)
        for call_name, arg_count, line_num in calls:
            if call_name in self.all_procedures:
                params, min_args, max_args, is_func, _, module_name = self.all_procedures[call_name]
                
                # Si c'est une sub appelée sans `Call` et sans `()`, VBA autorise à omettre les parenthèses
                # C'est un cas complexe que le linter ne gère pas parfaitement.
                # Pour l'instant, on se fie au compte d'arguments.
                if not (min_args <= arg_count <= max_args):
                    log_message(f"Dans '{filename}', procédure '{proc_name}' (ligne {line_num}): Appel à '{call_name}' avec {arg_count} arguments, mais entre {min_args} et {max_args} attendus.", CRITICAL)


    def check_directory(self, directory: str):
        vba_files = find_vba_files(directory)
        
        # 1. Analyser toutes les signatures d'abord
        self.analyze_all_files(vba_files)

        # 2. Analyser le contenu de chaque fichier
        for file_path in vba_files:
            try:
                with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                    content = f.read()
                self._check_file_content(file_path, content)
            except Exception as e:
                log_message(f"Error checking file {file_path}: {e}", CRITICAL)
        
        # 3. Vérifier la cohérence XML-VBA après avoir tout analysé
        self._check_xml_vba_consistency()

    def _check_xml_vba_consistency(self):
        vba_procs = {k.lower() for k in self.all_procedures.keys()}
        for callback in self.xml_callbacks:
            if callback.lower() not in vba_procs:
                log_message(f"Le callback XML '{callback}' n'a pas de procédure VBA correspondante.", CRITICAL)


def main():
    setup_logging()
    linter = VBALinter()
    linter.check_directory(VBA_SRC_DIR)
    print(f"\nAnalyse terminée. Rapport sauvegardé dans {LOG_FILE}")

if __name__ == "__main__":
    main() 