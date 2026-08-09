# Base architecture and logic adapted from ai_bridge_New Ollama V14.py
from flask import Flask, request, jsonify
from sentence_transformers import SentenceTransformer
import difflib
import torch
import numpy as np
import pandas as pd
import os
import re
import json
import threading
import requests
import glob
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

app = Flask(__name__)

# ==========================================
# 1. INITIALIZE PURE MATH RADAR & WHITELIST
# ==========================================
print("Loading Pure Semantic Math Engine...")

device = "cuda" if torch.cuda.is_available() else "cpu"
print(f"Engine forcing hardware acceleration on: [{device.upper()}]")

embedder = SentenceTransformer('all-MiniLM-L6-v2', device=device)

script_dir = os.path.dirname(os.path.abspath(__file__))
appdata = os.environ.get('APPDATA', '')
if not appdata:
    appdata = os.path.expanduser('~') 

# Thread locks for safe hot-reloading
custom_rules_lock = threading.Lock()
knowledge_base_lock = threading.Lock()
typo_cache_lock = threading.Lock()

# Global variables
lookup_phrases = []
lookup_phrases_original = [] 
canonical_casing = {}        
lookup_words_sets = []
lookup_embeddings = None

industry_translations = {}
list_translations = {} 
broad_categories = []

compiled_expert_rules = [] 
compiled_list_rules = [] 
compiled_broad_categories = [] 

uniformat_valid_words = set()
stemmed_whitelist = set() 
master_rule_pattern = None

# Custom Learned Rules memory
custom_rules = {}
compiled_custom_rules = [] 

# Memoization cache for typo correction
typo_cache = {}

def recompile_custom_rules_global():
    """Compiles all learned mappings into regex objects globally to avoid overhead per API request."""
    global compiled_custom_rules
    with custom_rules_lock:
        sorted_custom_rules = sorted(custom_rules.items(), key=lambda x: len(x[0]), reverse=True)
        compiled_custom_rules = [
            (re.compile(r'(?<![a-z0-9])' + re.escape(k) + r'(?:s|es)?(?![a-z0-9])', re.IGNORECASE), v, get_stemmed_words(k)) 
            for k, v in sorted_custom_rules if k
        ]

# ==========================================
# PRE-COMPILED REGEX PATTERNS
# ==========================================
COMPILED_LOCATION_PATTERNS = [
    re.compile(r'\b(?:elevator )?lobby\b', re.IGNORECASE),
    re.compile(r'\b(?:stair|stairwell|staircase)(?:\s+\d+[a-z]?|\s+[a-z]\b)?\b', re.IGNORECASE),
    re.compile(r'\b(?:vestibule|corridor|hallway|hall|concourse|mezzanine|mezz|penthouse|ph|basement|roof|ceiling|atrium|closet)\b', re.IGNORECASE),
    re.compile(r'\b(?:parking )?garage\b', re.IGNORECASE),
    re.compile(r'\b(?:level|floor|fl)\s*\d+[a-z]*\b', re.IGNORECASE),
    re.compile(r'\b\d+(?:st|nd|rd|th)?\s+(?:floor|fl|level)\b', re.IGNORECASE),
    re.compile(r'\bp\d+\b', re.IGNORECASE), 
    re.compile(r'\b(?:mechanical|elec|electrical|boiler|pump|utility|telecom|it|server|data|storage|mail|garbage|comm)\s+room\b', re.IGNORECASE),
    re.compile(r'\bl\d{2,3}\b', re.IGNORECASE),
    re.compile(r'\b(?:block|zone|sector|phase|stage|bldg)\s*[a-z0-9-]+\b', re.IGNORECASE),
    re.compile(r'\b(?:male|female|mens|womens|family)\b', re.IGNORECASE),
    re.compile(r'\b\d+(?:st|nd|rd|th)\b', re.IGNORECASE),
    re.compile(r'\b(?:northeast|northwest|southeast|southwest|north|south|east|west|corner|side|area)\b', re.IGNORECASE),
    re.compile(r'\b(?:panel|unit|board|bldg|b|c|d)\s+\d+\b', re.IGNORECASE)
]

HASH_TAG_PATTERN = re.compile(r'#\s*[A-Z0-9]+', re.IGNORECASE)
ROOM_PATTERN = re.compile(r'\b(?:room|rm)\s*[A-Z0-9]+', re.IGNORECASE)
TRAILING_NUM_PATTERN = re.compile(r'(\d+[A-Z]?)$')
CONTROL_CHAR_PATTERN = re.compile(r'[\x00-\x09\x0B-\x0C\x0E-\x1F\u200B-\u200D\uFEFF]')
NON_ALNUM_PATTERN = re.compile(r'[^a-z0-9\s]')
MULTI_SPACE_PATTERN = re.compile(r'\s+')
TRAILING_DIGIT_PATTERN = re.compile(r'\s+(?<![a-z0-9])\d+(?![a-z0-9])$')

# Pre-compiled Hierarchy Rules
HIERARCHY_RULES = [
    {
        "group": "Life Safety / Fire",
        "parents": [r'\bfire alarm control', r'\bfire panel', r'\bfacp\b', r'\bannunciator', r'\bmain fire'],
        "children": [r'\bsmoke', r'\bdetector', r'\bstrobe', r'\bhorn', r'\bpull', r'\bbell', r'\btamper switch', r'\bflow switch', r'\bfire damper', r'\binterface device', r'\bfire box']
    },
    {
        "group": "Emergency Power",
        "parents": [r'\bgenerator', r'\bgenset', r'\bdiesel gen', r'\bgas gen'],
        "children": [r'\bats\b', r'\bautomatic transfer switch', r'\btransfer switch', r'\bload bank'] 
    },
    {
        "group": "Vertical Transport",
        "parents": [r'\belevator', r'\bcab\b', r'\bescalator', r'\bmoving walk', r'\bwheelchair lift', r'\bdumbwaiter'],
        "children": [r'\belevator controller', r'\bhoistway', r'\bpit equipment', r'\belevator machine room']
    },
    {
        "group": "Hydronic HVAC",
        "parents": [r'\bhydronic fan coil'],
        "children": []
    },
    {
        "group": "Split HVAC",
        "parents": [r'\bcondensing unit', r'\bcondenser', r'\bcu\b', r'\boutdoor unit', r'\bvrf outdoor', r'\bvrv outdoor', r'\bheat pump outdoor'],
        "children": [r'\bfan coil', r'\bfcu\b', r'\bevaporator', r'\bvrf indoor', r'\bvrv indoor', r'\bcassette', r'\bwall mount split', r'\bindoor unit']
    },
    {
        "group": "Fire Pump System",
        "parents": [r'\bfire pump', r'\belectric fire pump', r'\bdiesel fire pump'],
        "children": [r'\bjockey pump', r'\bpressure maintenance pump', r'\bfire pump controller']
    },
    {
        "group": "Fire Sprinkler",
        "parents": [r'\bsprinkler system', r'\bwet pipe', r'\bdry pipe', r'\bpre-action', r'\bdeluge'],
        "children": [r'\bsprinkler head', r'\bdrop\b']
    },
    {
        "group": "Water Cooled Heat Pump",
        "parents": [r'\bwater cooled heat pump'],
        "children": []
    },
    {
        "group": "Ambiguous Fuel System",
        "parents": [],
        "children": [r'\bday tank', r'\bfuel polishing', r'\bfuel tank', r'\bfuel oil tank', r'\btransfer pump']
    }
]

COMPILED_HIERARCHY_RULES = [
    {
        "group": rule["group"],
        "parents": [re.compile(p, re.IGNORECASE) for p in rule["parents"]],
        "children": [re.compile(c, re.IGNORECASE) for c in rule["children"]]
    }
    for rule in HIERARCHY_RULES
]

def get_stemmed_words(text):
    words = text.split()
    stemmed = set()
    protected_words = {"heating", "cooling", "piping", "building", "ceiling", "wiring", "lighting"}
    for w in words:
        if len(w) > 4 and w not in protected_words:
            if w.endswith('ing'): w = w[:-3]
            elif w.endswith('ed'): w = w[:-2]
            elif w.endswith('es'): w = w[:-2]
            elif w.endswith('s') and not w.endswith('ss'): w = w[:-1]
        stemmed.add(w)
    return stemmed

# STRICT LOOKUP ENFORCER: Guarantees ONLY exact phrases from the lookup sheet survive
def validate_and_format(match_str):
    if not match_str or match_str in ["No good match", "REQUIRES HUMAN", "SKIP", "Out of Scope (Architectural/Civil)", "Subtype missing in input"]:
        return match_str
    
    if "\n" in match_str:
        valid_parts = []
        for part in match_str.split("\n"):
            part_clean = part.strip()
            clean = canonical_casing.get(part_clean.lower())
            if clean and clean not in valid_parts: 
                valid_parts.append(clean)
        return "\n".join(valid_parts) if valid_parts else "No good match"
    
    match_clean = match_str.strip()
    clean_match = canonical_casing.get(match_clean.lower())
    return clean_match if clean_match else "No good match"

def load_knowledge_base():
    global lookup_phrases, lookup_phrases_original, canonical_casing
    global lookup_words_sets, lookup_embeddings
    global industry_translations, list_translations, broad_categories, uniformat_valid_words
    global stemmed_whitelist, compiled_expert_rules, compiled_list_rules, compiled_broad_categories
    global master_rule_pattern

    print("--- Initializing Engine ---")

    appdata = os.environ.get('APPDATA', '')
    if not appdata: appdata = os.path.expanduser('~')
        
    t_lookup_phrases, t_lookup_phrases_original, t_canonical_casing = [], [], {}
    t_industry_translations, t_list_translations, t_broad_categories = {}, {}, []
    t_uniformat_valid_words = set()
    t_stemmed_whitelist = set()
    t_compiled_expert_rules = []
    t_compiled_list_rules = []
    t_compiled_broad_categories = []
    paired = []

    # 1. Locate the newest Excel template
    custom_template_dir = r"C:\Users\pcarr\OneDrive\Documents\Custom Office Templates"
    search_dirs = [appdata, os.path.dirname(os.path.abspath(__file__)), custom_template_dir]
    excel_files = []
    for d in search_dirs:
        for ext in ['*.xltm', '*.xlsm', '*.xlsx']:
            found_files = glob.glob(os.path.join(d, ext))
            for f in found_files:
                if not os.path.basename(f).startswith('~$'):
                    excel_files.append(f)
            
    if excel_files:
        latest_file = max(excel_files, key=os.path.getmtime)
        print(f"Reading configuration from newest Excel template: {latest_file}...")
        
        try:
            print("Reading Lookups from Column A...")
            df_lookups = pd.read_excel(latest_file, sheet_name="Uniformat RS Means Lookup", header=None, engine='openpyxl')
            df_lookups_records = df_lookups.dropna(how='all').to_dict('records')
            
            for row in df_lookups_records:
                if 0 in row and pd.notna(row[0]):
                    val = str(row[0]).strip()
                    if val.lower() == "asset sub type": continue # Skip header
                    
                    clean_phrase = val
                    
                    if clean_phrase:
                        lower_phrase = clean_phrase.lower()
                        stripped_phrase = NON_ALNUM_PATTERN.sub(' ', lower_phrase)
                        stripped_phrase = MULTI_SPACE_PATTERN.sub(' ', stripped_phrase).strip()
                        
                        paired.append((stripped_phrase, clean_phrase))
                        t_canonical_casing[lower_phrase] = clean_phrase
                        
                        for word in stripped_phrase.split():
                            t_uniformat_valid_words.add(word)
                            
            print(f"Loaded lookup phrases from Excel!")
        except Exception as e:
            print(f"ERROR reading Uniformat RS Means Lookup: {e}")

        try:
            print("Reading Rules...")
            df_rules = pd.read_excel(latest_file, sheet_name="Rules Sheet", header=None, engine='openpyxl')
            df_rules_records = df_rules.dropna(how='all').to_dict('records')
            
            for row in df_rules_records:
                if 0 in row and pd.notna(row[0]):
                    rule_type = str(row[0]).strip().lower()
                    from_text = str(row.get(1, "")).strip().lower()
                    to_text = str(row.get(2, "")).strip().replace('|', '\n')
                    
                    if rule_type in ["alias", "phrase"]:
                        clean_from = MULTI_SPACE_PATTERN.sub(' ', from_text).strip()
                        if clean_from:
                            t_industry_translations[clean_from] = to_text
                            for word in re.findall(r'[a-z0-9]+', to_text.lower()):
                                t_uniformat_valid_words.add(word)
                            for word in re.findall(r'[a-z0-9]+', clean_from):
                                t_uniformat_valid_words.add(word)
                    
                    elif rule_type == "list":
                        clean_from = MULTI_SPACE_PATTERN.sub(' ', from_text).strip()
                        if clean_from:
                            t_list_translations[clean_from] = to_text
                            for word in re.findall(r'[a-z0-9]+', to_text.lower()):
                                t_uniformat_valid_words.add(word)
                            for word in re.findall(r'[a-z0-9]+', clean_from):
                                t_uniformat_valid_words.add(word)
                                
                    elif rule_type == "core":
                        clean_from = MULTI_SPACE_PATTERN.sub(' ', from_text).strip()
                        if clean_from:
                            if clean_from not in t_broad_categories:
                                t_broad_categories.append(clean_from)
                            for word in re.findall(r'[a-z0-9]+', clean_from):
                                t_uniformat_valid_words.add(word)
            print(f"Loaded {len(t_industry_translations)} translations, {len(t_list_translations)} strict lists, and {len(t_broad_categories)} core categories!")
        except Exception as e:
            print(f"ERROR reading Rules sheet: {e}")

        try:
            print("Reading Learned Mappings...")
            df_learned = pd.read_excel(latest_file, sheet_name="LearnedMappings", header=None, engine='openpyxl')
            df_learned_records = df_learned.dropna(how='all').to_dict('records')
            
            with custom_rules_lock:
                custom_rules.clear()
                for row in df_learned_records:
                    if 0 in row and 1 in row and pd.notna(row[0]) and pd.notna(row[1]):
                        rule_k = str(row[0]).strip()
                        rule_v = str(row[1]).strip()
                        if rule_k and rule_v:
                            clean_k = NON_ALNUM_PATTERN.sub(' ', rule_k)
                            clean_k = MULTI_SPACE_PATTERN.sub(' ', clean_k).strip()
                            if clean_k:
                                custom_rules[clean_k] = rule_v
            print(f"Loaded {len(custom_rules)} custom learned rules from Excel!")
        except Exception as e:
            print(f"ERROR reading Learned Mappings sheet: {e}")
    else:
        print("WARNING: No valid Excel templates (.xltm, .xlsm, .xlsx) found.")

    # 2. Overlay live text files from %APPDATA%
    lookups_file = os.path.join(appdata, "Uniformat_Lookups.txt")
    rules_file = os.path.join(appdata, "Uniformat_Rules.txt")
    custom_rules_file = os.path.join(appdata, "Uniformat_Learned.txt")

    if os.path.exists(lookups_file):
        try:
            print(f"Applying real-time fallbacks/lookups from {lookups_file}...")
            with open(lookups_file, "r", encoding="utf-8-sig") as f:
                for line in f:
                    parts = line.strip('\n').split('\t')
                    if parts:
                        clean_phrase = parts[0].strip()
                        if clean_phrase and clean_phrase.lower() != "asset sub type":
                            lower_phrase = clean_phrase.lower()
                            stripped_phrase = NON_ALNUM_PATTERN.sub(' ', lower_phrase)
                            stripped_phrase = MULTI_SPACE_PATTERN.sub(' ', stripped_phrase).strip()
                            
                            if lower_phrase not in t_canonical_casing:
                                paired.append((stripped_phrase, clean_phrase))
                                t_canonical_casing[lower_phrase] = clean_phrase
                                for word in stripped_phrase.split():
                                    t_uniformat_valid_words.add(word)
        except Exception as e:
            print(f"ERROR reading Lookups text file: {e}")
            
    paired.sort(key=lambda x: len(x[0]), reverse=True)
    t_lookup_phrases = [p[0] for p in paired]
    t_lookup_phrases_original = [p[1] for p in paired]
    print(f"Total fully-validated lookup phrases ready: {len(t_lookup_phrases)}")

    if os.path.exists(rules_file):
        try:
            print(f"Applying real-time overrides from {rules_file}...")
            with open(rules_file, "r", encoding="utf-8-sig") as f:
                for line in f:
                    parts = line.strip('\n').split('\t')
                    if len(parts) >= 3:
                        rule_type = parts[0].strip().lower()
                        from_text = parts[1].strip().lower()
                        to_text = parts[2].strip().replace('|', '\n')
                        
                        if rule_type in ["alias", "phrase"]:
                            clean_from = MULTI_SPACE_PATTERN.sub(' ', from_text).strip()
                            if clean_from:
                                t_industry_translations[clean_from] = to_text
                                for word in re.findall(r'[a-z0-9]+', to_text.lower()):
                                    t_uniformat_valid_words.add(word)
                                for word in re.findall(r'[a-z0-9]+', clean_from):
                                    t_uniformat_valid_words.add(word)
                                    
                        elif rule_type == "list":
                            clean_from = MULTI_SPACE_PATTERN.sub(' ', from_text).strip()
                            if clean_from:
                                t_list_translations[clean_from] = to_text
                                for word in re.findall(r'[a-z0-9]+', to_text.lower()):
                                    t_uniformat_valid_words.add(word)
                                for word in re.findall(r'[a-z0-9]+', clean_from):
                                    t_uniformat_valid_words.add(word)
                                    
                        elif rule_type == "core":
                            clean_from = MULTI_SPACE_PATTERN.sub(' ', from_text).strip()
                            if clean_from:
                                if clean_from not in t_broad_categories:
                                    t_broad_categories.append(clean_from)
                                for word in re.findall(r'[a-z0-9]+', clean_from):
                                    t_uniformat_valid_words.add(word)
        except Exception as e:
            print(f"ERROR reading Rules text file: {e}")

    if os.path.exists(custom_rules_file):
        try:
            print(f"Applying real-time overrides from {custom_rules_file}...")
            with open(custom_rules_file, "r", encoding="utf-8-sig") as f:
                with custom_rules_lock:
                    for line in f:
                        line = line.strip()
                        if line:
                            parts = line.split("|")
                            if len(parts) >= 2:
                                rule_k = parts[0].strip()
                                rule_v = parts[1].strip()
                                if rule_k and rule_v:
                                    clean_k = NON_ALNUM_PATTERN.sub(' ', rule_k)
                                    clean_k = MULTI_SPACE_PATTERN.sub(' ', clean_k).strip()
                                    if clean_k:
                                        custom_rules[clean_k] = rule_v
        except Exception as e:
            print(f"ERROR reading LearnedMappings text file: {e}")

    for rule in HIERARCHY_RULES:
        for p in rule['parents'] + rule['children']:
            clean_p = p.replace(r'\b', '')
            for w in re.findall(r'[a-z0-9]+', clean_p.lower()):
                t_uniformat_valid_words.add(w)

    if not t_broad_categories:
        t_broad_categories = ["pump", "fan", "boiler", "furnace", "transformer", "compressor", "chiller", "motor", "sump", "sprinkler", "valve", "hydrant", "tower", "exchanger", "tank"]
        for cat in t_broad_categories:
            for word in re.findall(r'[a-z0-9]+', cat):
                t_uniformat_valid_words.add(word)

    sorted_messy_terms = sorted(t_industry_translations.keys(), key=len, reverse=True)
    if sorted_messy_terms:
        escaped_keys = [r'(?<![a-z0-9])' + re.escape(k) + r'(?:s|es)?(?![a-z0-9])' for k in sorted_messy_terms]
        t_master_rule_pattern = re.compile('(' + '|'.join(escaped_keys) + ')', flags=re.IGNORECASE)
    else:
        t_master_rule_pattern = None

    # Compile Expert Rules (Substrings)
    for messy_term in sorted_messy_terms:
        clean_term = t_industry_translations[messy_term]
        if clean_term:
            clean_messy = NON_ALNUM_PATTERN.sub(' ', messy_term)
            rule_tokens = get_stemmed_words(clean_messy)
            compiled_pat = re.compile(r'(?<![a-z0-9])' + re.escape(messy_term) + r'(?:s|es)?(?![a-z0-9])', re.IGNORECASE)
            t_compiled_expert_rules.append((compiled_pat, messy_term, rule_tokens, clean_term))

    # Compile Strict List Rules (Exact Match Only)
    for messy_term, clean_term in t_list_translations.items():
        if clean_term:
            clean_messy = NON_ALNUM_PATTERN.sub(' ', messy_term)
            rule_tokens = get_stemmed_words(clean_messy)
            compiled_pat = re.compile(r'^' + re.escape(messy_term) + r'(?:s|es)?$', re.IGNORECASE)
            t_compiled_list_rules.append((compiled_pat, messy_term, rule_tokens, clean_term))

    t_compiled_broad_categories = [re.compile(r'(?<![a-z0-9])' + re.escape(cat) + r'(?![a-z0-9])', re.IGNORECASE) for cat in t_broad_categories]
    t_lookup_words_sets = [get_stemmed_words(phrase) for phrase in t_lookup_phrases]

    t_lookup_embeddings = None
    if t_lookup_phrases:
        embeddings = embedder.encode(t_lookup_phrases, batch_size=256)
        norms = np.linalg.norm(embeddings, axis=1, keepdims=True)
        t_lookup_embeddings = embeddings / (norms + 1e-9)

    for w in t_uniformat_valid_words:
        t_stemmed_whitelist.add(w)
        t_stemmed_whitelist.update(get_stemmed_words(w))
        
    with typo_cache_lock:
        typo_cache.clear()

    with knowledge_base_lock:
        lookup_phrases = t_lookup_phrases
        lookup_phrases_original = t_lookup_phrases_original
        canonical_casing = t_canonical_casing
        industry_translations = t_industry_translations
        list_translations = t_list_translations
        broad_categories = t_broad_categories
        uniformat_valid_words = t_uniformat_valid_words
        stemmed_whitelist = t_stemmed_whitelist
        master_rule_pattern = t_master_rule_pattern
        compiled_expert_rules = t_compiled_expert_rules
        compiled_list_rules = t_compiled_list_rules
        compiled_broad_categories = t_compiled_broad_categories
        lookup_words_sets = t_lookup_words_sets
        lookup_embeddings = t_lookup_embeddings

    recompile_custom_rules_global()
    print(f"Engine Ready! Whitelist contains {len(uniformat_valid_words)} authorized words.")

load_knowledge_base()

# ==========================================
# 2. THE UNIFIED PRE-SCRUB & LOCAL AI ENGINE
# ==========================================
def extract_locations(raw_text):
    if not isinstance(raw_text, str): 
        return "", raw_text
    locations = []
    clean_base = raw_text.strip()
    
    for pattern in COMPILED_LOCATION_PATTERNS:
        matches = pattern.findall(clean_base)
        if matches:
            locations.extend(matches)
            clean_base = pattern.sub('', clean_base)
            
    clean_base = MULTI_SPACE_PATTERN.sub(' ', clean_base).strip()
    extracted = " | ".join(sorted(set([loc.title() for loc in locations])))
    return extracted, clean_base

def extract_asset_identifiers(raw_text):
    if not isinstance(raw_text, str): 
        return "", raw_text
    identifiers = []
    clean_base = raw_text.strip()
    
    hash_tags = HASH_TAG_PATTERN.findall(clean_base)
    identifiers.extend(hash_tags)
    clean_base = HASH_TAG_PATTERN.sub('', clean_base)
        
    rooms = ROOM_PATTERN.findall(clean_base)
    identifiers.extend(rooms)
    clean_base = ROOM_PATTERN.sub('', clean_base)
        
    clean_base = clean_base.strip()
    trailing_num = TRAILING_NUM_PATTERN.search(clean_base)
    if trailing_num and not hash_tags:
        identifiers.append(trailing_num.group(1))
        clean_base = clean_base[:trailing_num.start()]
        
    return " ".join(identifiers).strip(), clean_base.strip()

def purify_text(scrubbed_text, current_valid_words, current_stemmed_whitelist):
    if not isinstance(scrubbed_text, str) or not scrubbed_text:
        return ""
    if not current_valid_words:
        return scrubbed_text
        
    global typo_cache
    pure_words = []
    for w in scrubbed_text.split():
        if w in current_valid_words:
            pure_words.append(w)
        else:
            stem_set = get_stemmed_words(w)
            w_stem = list(stem_set)[0] if stem_set else ""
            
            if w_stem in current_stemmed_whitelist:
                pure_words.append(w)
            elif len(w) >= 3:
                with typo_cache_lock:
                    cached = w in typo_cache
                    cache_val = typo_cache.get(w)
                    
                if cached:
                    if cache_val:
                        pure_words.append(cache_val)
                else:
                    matches = difflib.get_close_matches(w, current_valid_words, n=1, cutoff=0.8)
                    with typo_cache_lock:
                        if matches:
                            typo_cache[w] = matches[0]
                            pure_words.append(matches[0])
                        else:
                            typo_cache[w] = None
                    
    pure = " ".join(pure_words)
    return pure if pure else scrubbed_text

def prescrub_text(raw_text, use_whitelist=True, current_valid_words=None, current_stemmed_whitelist=None, current_pattern=None):
    if not isinstance(raw_text, str): return ""
        
    clean = CONTROL_CHAR_PATTERN.sub('', raw_text)
    clean = clean.replace('\xa0', ' ').strip().lower()
    
    clean = re.sub(r'\b(?:blue|green|yellow|red|orange|purple|brown|grey|gray|white|black|east|west|north|south|behind|beside|across|inside|freight|retail|tenant|dock|bay)\b', '', clean)
    
    if current_pattern:
        def replace_logic(m):
            match_word = m.group(0).lower()
            target_text = industry_translations.get(match_word)
            
            if target_text is None:
                if match_word.endswith('es') and match_word[:-2] in industry_translations:
                    target_text = industry_translations.get(match_word[:-2])
                elif match_word.endswith('s') and match_word[:-1] in industry_translations:
                    target_text = industry_translations.get(match_word[:-1])
                else:
                    target_text = ""
                    
            if '\n' in target_text:
                return match_word
            return target_text.split('\n')[0].lower()
        
        clean = current_pattern.sub(replace_logic, clean)

    clean = NON_ALNUM_PATTERN.sub(' ', clean)
    clean = MULTI_SPACE_PATTERN.sub(' ', clean).strip()
    clean = TRAILING_DIGIT_PATTERN.sub('', clean).strip()
    
    if use_whitelist and current_valid_words is not None:
        clean = purify_text(clean, current_valid_words, current_stemmed_whitelist)
    return clean

# ==========================================
# PM SUITABILITY RULES
# ==========================================
PM_SUITABILITY_RULES = {
    "Comprehensive PM": [
        r'\bchiller\b', r'\bboiler\b', r'\bgenerator\b', r'\belevator\b', r'\bescalator\b', 
        r'\bair handling unit\b', r'\bartu\b', r'\bcooling tower\b', r'\bmake up air\b', r'\bmau\b',
        r'\bpackage ac\b', r'\bsplit system\b', r'\bair conditioner\b', r'\bhumidifier\b',
        r'\bvav\b', r'\bfan coil\b', r'\bfcu\b', r'\bfan\b', 
        r'\bunit heater\b', r'\bradiant heater\b', r'\bair terminal\b',
        r'\bpump\b', r'\bair compressor\b', r'\bsteam trap\b', r'\bprv\b', 
        r'\bstandpipe\b', r'\briser\b', r'\bbackflow\b',
        r'\bpanel\b', r'\btransformer\b', r'\bbreaker\b', r'\bswitch\b', 
        r'\bemergency lighting\b', r'\bengine\b', r'\bcircuits\b',
        r'\bfire extinguisher\b', r'\bfire system\b', r'\bsprinkler\b', 
        r'\bfire hose cabinet\b', r'\bfire alarm\b', r'\bdetector\b', 
        r'\bpull station\b', r'\bsecurity system\b', r'\baccess control\b', 
        r'\bvideo surveillance\b', r'\bfirefighter\b', r'\beyewash\b', r'\beye wash\b',
        r'\bair quality\b', r'\bgas\s?(?:sensor|detector|monitor)\b', 
        r'\b(?:co|co2|oxygen)\s?(?:sensor|detector|monitor)\b'
    ],
    "Non-PM": [
        r'\bwindow\b', r'\bexit sign\b', r'\bpipe\b', r'\bmanual valve\b', 
        r'\bductwork\b', r'\bwall\b', r'\bwashroom\b', r'\bjanitor\b'
    ],
    "Questionable": [
        r'\bday tank\b', r'\bmeter\b', r'\bsensor\b', r'\bdoor\b', r'\bportable\b'
    ]
}

COMPILED_PM_SUITABILITY_RULES = {}
for status, patterns in PM_SUITABILITY_RULES.items():
    compiled_patterns = []
    for pattern in patterns:
        strict_pattern = pattern.replace(r'\b', '')
        strict_pattern = r'(?<![a-z0-9])' + strict_pattern + r'(?![a-z0-9])'
        compiled_patterns.append(re.compile(strict_pattern, re.IGNORECASE))
    COMPILED_PM_SUITABILITY_RULES[status] = compiled_patterns

def determine_pm_suitability(scrubbed_phrase):
    for status, compiled_patterns in COMPILED_PM_SUITABILITY_RULES.items():
        for pattern in compiled_patterns:
            if pattern.search(scrubbed_phrase):
                return status
    return "Questionable"

def categorize_asset(scrubbed_phrase):
    for rule in COMPILED_HIERARCHY_RULES:
        for p_regex in rule['parents']:
            if p_regex.search(scrubbed_phrase): return rule['group'], "Parent"
        for c_regex in rule['children']:
            if c_regex.search(scrubbed_phrase): return rule['group'], "Child"
    return "N/A", "N/A"

def ask_maintenance_director_ollama(raw_phrase, lookups_sample):
    url = "http://localhost:11434/api/generate"
    sample = lookups_sample if lookups_sample else []
    
    prompt = f"""
You are an expert Maintenance Director processing a raw, messy client asset registry.
Your job is to identify what the asset actually is and align it with a standard CMMS asset lookup type.

RAW CLIENT INPUT: "{raw_phrase}"

SAMPLE VALID UNIFORMAT LOOKUPS FOR REFERENCE:
{json.dumps(sample, indent=2)}

INSTRUCTIONS:
1. Extract any explicit location, room, unit, or identifying tag (e.g., "#1", "ROOM 101", "UNIT A"). YOU MUST NOT invent, guess, or hallucinate locations. If no tag exists in the raw input, leave it blank.
2. Deduce the true underlying core equipment type, ignoring messy abbreviations.
3. Select the best matching asset classification from the provided reference list.

Respond ONLY in valid JSON with these exact keys:
{{
    "matched_asset": "Clean Asset Type Here",
    "asset_tag": "Extracted Location or Tag Here (or empty string)",
    "confidence": "High/Medium/Low"
}}
"""
    payload = {
        "model": "qwen2.5:7b",
        "prompt": prompt,
        "format": "json",
        "stream": False
    }
    try:
        response = requests.post(url, json=payload, timeout=30)
        if response.status_code == 200:
            return json.loads(response.json().get("response", ""))
    except Exception as e:
        print(f"Ollama API Error: {e}")
    
    return None

# ==========================================
# CENTRALIZED BATCH LOGIC
# ==========================================
def run_ai_batch_processing(items, is_batch_file=False):
    results = []
    ai_queue_phrases = []
    ai_queue_rows = []
    original_phrases_list = []
    clean_base_phrases_list = []

    k_row = "Row" if is_batch_file else "row"
    k_match = "Match" if is_batch_file else "match"
    k_id = "Id" if is_batch_file else "id"

    with custom_rules_lock:
        local_compiled_custom_rules = compiled_custom_rules
        local_custom_rules = custom_rules.copy()

    with knowledge_base_lock:
        local_lookup_phrases = lookup_phrases
        local_lookup_phrases_original = lookup_phrases_original
        local_lookup_embeddings = lookup_embeddings
        local_lookup_words_sets = lookup_words_sets
        local_compiled_expert_rules = compiled_expert_rules
        local_compiled_list_rules = compiled_list_rules
        local_compiled_broad_categories = compiled_broad_categories
        local_valid_words = uniformat_valid_words
        local_stemmed_whitelist = stemmed_whitelist
        local_master_rule_pattern = master_rule_pattern

    for row in items:
        if is_batch_file:
            row_id = str(row.get('Row', ''))
            original_phrase = str(row.get('Phrase', '')).strip()
            if original_phrase == 'nan': original_phrase = ''
        else:
            row_id = row.get('row')
            original_phrase = row.get('phrase', '').strip()
            current_match = row.get('current_match', '').strip().lower()
            current_id = row.get('current_id', '').strip().lower()

            is_broken = False
            if current_match in ["", "no good match", "no matching words"]: is_broken = True
            if any(x in current_id for x in ["none", "category", "requires human", "requires ai"]): is_broken = True

            if not is_broken:
                results.append({k_row: row_id, k_match: "SKIP", k_id: "SKIP"})
                continue

        _, text_no_loc = extract_locations(original_phrase)
        _, clean_base_phrase = extract_asset_identifiers(text_no_loc)

        signature = prescrub_text(clean_base_phrase, use_whitelist=True, 
                                  current_valid_words=local_valid_words, 
                                  current_stemmed_whitelist=local_stemmed_whitelist,
                                  current_pattern=local_master_rule_pattern)
        
        semantic_phrase = prescrub_text(clean_base_phrase, use_whitelist=False, 
                                        current_valid_words=local_valid_words, 
                                        current_stemmed_whitelist=local_stemmed_whitelist,
                                        current_pattern=local_master_rule_pattern)

        original_clean = NON_ALNUM_PATTERN.sub(' ', original_phrase.lower())

        match_found = local_custom_rules.get(signature)
        if not match_found:
            for rule_regex, rule_val, _ in local_compiled_custom_rules:
                if rule_regex.search(signature):
                    match_found = rule_val
                    break
        if not match_found:
            sig_words = get_stemmed_words(signature)
            best_mem_score = 0.0
            for _, rule_val, rule_words in local_compiled_custom_rules:
                if not rule_words: continue
                intersection = len(sig_words.intersection(rule_words))
                total_len = len(sig_words) + len(rule_words)
                overlap = (2.0 * intersection) / total_len if total_len > 0 else 0.0
                
                if overlap > best_mem_score and overlap >= 0.90:
                    best_mem_score = overlap
                    match_found = rule_val

        if match_found:
            validated = validate_and_format(match_found)
            if validated != "No good match":
                results.append({k_row: row_id, k_match: validated, k_id: "USER_LEARNED"})
                continue

        matched_rules = []
        for pattern, messy_term, rule_tokens, clean_term in local_compiled_expert_rules:
            if pattern.search(original_clean):
                matched_rules.append((messy_term, clean_term))
                
        # FIX: Ensure list dropdowns trigger perfectly by evaluating against the stripped signature
        for pattern, messy_term, rule_tokens, clean_term in local_compiled_list_rules:
            if pattern.search(signature): 
                matched_rules.append((messy_term, clean_term))
                
        expert_matches = []
        for i, (term1, clean1) in enumerate(matched_rules):
            is_subset = False
            for j, (term2, clean2) in enumerate(matched_rules):
                if i != j and term1 in term2 and len(term1) < len(term2):
                    is_subset = True
                    break
            if not is_subset:
                validated = validate_and_format(clean1.strip())
                if validated != "No good match":
                    for part in validated.split("\n"):
                        if part not in expert_matches:
                            expert_matches.append(part)

        s_pad = f" {semantic_phrase} "
        if " pump " in s_pad:
            specific_pumps = ["sump", "submersible", "fuel", "condensate", "hydraulic",
                              "fire", "well", "ejector", "metering", "vacuum",
                              "sewage", "lift", "jockey", "booster", "rotary",
                              "centrifugal", "circ", "circulation", "dosing", "chem"]
            if not any(f" {p} " in s_pad for p in specific_pumps):
                validated = validate_and_format("Centrifugal Pump")
                if validated != "No good match" and validated not in expert_matches:
                    expert_matches.append(validated)

        if expert_matches:
            final_match = "\n".join(expert_matches[:8])
            results.append({k_row: row_id, k_match: final_match, k_id: "EXPERT_RULE"})
        else:
            EXCLUSION_PATTERN = re.compile(r'\b(?:structural|furniture|masonry|glazing|pavement|interlocking paver|playground|mural|pest control|washroom accessories|curbing|bollard|planter|ceiling|floors?|flooring|walls?|landscape|landscaping|painting|paint|garbage recep|waste recep|ash urn|sidewalks?|parking|decks?|sweeper|window washing|recycling|charcoal|ramp plates?)\b', re.IGNORECASE)
            
            if EXCLUSION_PATTERN.search(original_phrase):
                results.append({k_row: row_id, k_match: "Out of Scope (Architectural/Civil)", k_id: "AUTO_EXCLUDED"})
            else:
                # NEW GROUPING LOGIC: Group AI queues tightly via signature to restore speed
                queue_phrase = signature if len(signature) >= 2 else semantic_phrase
                if len(queue_phrase) >= 2:
                    ai_queue_phrases.append(queue_phrase)
                    ai_queue_rows.append(row_id)
                    original_phrases_list.append(original_phrase)
                    clean_base_phrases_list.append(clean_base_phrase)
                else:
                    results.append({k_row: row_id, k_match: "No good match", k_id: "REQUIRES HUMAN"})

    if ai_queue_phrases:
        grouped_tasks = {}
        for i, sig_phrase in enumerate(ai_queue_phrases):
            if sig_phrase not in grouped_tasks:
                clean_base = clean_base_phrases_list[i]
                scrubbed_sig = prescrub_text(clean_base, use_whitelist=True, current_valid_words=local_valid_words, current_stemmed_whitelist=local_stemmed_whitelist, current_pattern=local_master_rule_pattern)
                
                grouped_tasks[sig_phrase] = {
                    "rows": [],
                    "sample_original": original_phrases_list[i],
                    "pure_signature": scrubbed_sig,
                    "strict_signature": scrubbed_sig
                }
            grouped_tasks[sig_phrase]["rows"].append(ai_queue_rows[i])
            
        unique_sem_phrases = list(grouped_tasks.keys())
        total_groups = len(unique_sem_phrases)
        
        unique_candidates = set()
        item_candidates_list = []

        for semantic_phrase in unique_sem_phrases:
            candidates_set = {semantic_phrase}
            words = semantic_phrase.split()
            for n in range(1, 6):
                if len(words) >= n:
                    for j in range(len(words) - n + 1):
                        candidates_set.add(' '.join(words[j:j+n]))
            candidates = list(candidates_set)
            unique_candidates.update(candidates)
            item_candidates_list.append(candidates)

        unique_candidates_list = list(unique_candidates)
        cand_to_idx = {cand: i for i, cand in enumerate(unique_candidates_list)}
        unique_vectors = embedder.encode(unique_candidates_list, batch_size=256)
        unique_norms = unique_vectors / (np.linalg.norm(unique_vectors, axis=1, keepdims=True) + 1e-9)

        all_pure_signatures = [grouped_tasks[sp]["pure_signature"] for sp in unique_sem_phrases]
        
        unique_pure = list(set(all_pure_signatures))
        pure_to_idx = {p: i for i, p in enumerate(unique_pure)}
        unique_pure_vecs = embedder.encode(unique_pure, batch_size=256)
        unique_pure_norms = unique_pure_vecs / (np.linalg.norm(unique_pure_vecs, axis=1, keepdims=True) + 1e-9)
        
        indices = [pure_to_idx[p] for p in all_pure_signatures]
        pure_norms_array = unique_pure_norms[indices]

        top_k = min(20, len(local_lookup_phrases))

        def process_single_ai_group(i_queue):
            semantic_phrase = unique_sem_phrases[i_queue]
            group_data = grouped_tasks[semantic_phrase]
            
            strict_signature = group_data["strict_signature"]
            pure_signature = group_data["pure_signature"]
            
            candidates = item_candidates_list[i_queue]
            num_candidates = len(candidates)
            indices_cand = [cand_to_idx[c] for c in candidates]
            candidate_norms = unique_norms[indices_cand]

            base_words = get_stemmed_words(pure_signature)
            
            pure_norm = pure_norms_array[i_queue:i_queue+1]
            pure_semantic_scores = np.dot(pure_norm, local_lookup_embeddings.T)[0]

            all_semantic_scores = np.dot(candidate_norms, local_lookup_embeddings.T)

            collected_matches = []
            overlap_cache = {}

            input_core_regexes = [
                pattern for pattern in local_compiled_broad_categories 
                if pattern.search(strict_signature)
            ]

            for c_idx in range(num_candidates):
                semantic_scores = all_semantic_scores[c_idx]
                combined_max_scores = np.maximum(semantic_scores, pure_semantic_scores)
                
                if len(combined_max_scores) > top_k:
                    idx_part = np.argpartition(combined_max_scores, -top_k)[-top_k:]
                    top_indices = idx_part[np.argsort(combined_max_scores[idx_part])[::-1]]
                else:
                    top_indices = np.argsort(combined_max_scores)[::-1]

                for idx in top_indices:
                    sem_score = combined_max_scores[idx]
                    lookup_candidate_original = local_lookup_phrases_original[idx]

                    if idx not in overlap_cache:
                        lookup_words = local_lookup_words_sets[idx]
                        intersection = len(base_words.intersection(lookup_words))
                        overlap = (2.0 * intersection) / (len(base_words) + len(lookup_words)) if (len(base_words) + len(lookup_words)) > 0 else 0.0
                        overlap_cache[idx] = (overlap * 0.35) + (0.15 if overlap >= 0.80 else 0.0)

                    combined_score = (sem_score * 0.50) + overlap_cache[idx]

                    if input_core_regexes:
                        candidate_lower = lookup_candidate_original.lower()
                        has_matching_core = any(core_re.search(candidate_lower) for core_re in input_core_regexes)
                        if not has_matching_core:
                            combined_score *= 0.40 

                    collected_matches.append((combined_score, lookup_candidate_original))

            collected_matches.sort(key=lambda x: x[0], reverse=True)
            unique_matches = []
            seen_phrases = set()
            for score, phrase in collected_matches:
                if phrase not in seen_phrases:
                    unique_matches.append((score, phrase))
                    seen_phrases.add(phrase)

            best_combined_score = unique_matches[0][0] if unique_matches else 0.0
            valid_options = []

            if best_combined_score >= 0.75:
                v = validate_and_format(unique_matches[0][1])
                if v != "No good match":
                    valid_options.append(v)
            elif best_combined_score >= 0.40:
                for m in unique_matches:
                    if m[0] >= 0.40 and (best_combined_score - m[0] <= 0.15):
                        v = validate_and_format(m[1])
                        if v != "No good match" and v not in valid_options:
                            valid_options.append(v)
                            if len(valid_options) >= 6: break

            if len(valid_options) > 1:
                final_match = "\n".join(valid_options[:8])
                final_id = "AI_HYBRID_MATCH"
            elif len(valid_options) == 1:
                final_match = valid_options[0]
                if best_combined_score >= 0.75:
                    final_id = "AI_SMART_VECTOR"
                else:
                    final_id = "AI_HYBRID_MATCH"
            else:
                broad_match_found = False
                for pattern in local_compiled_broad_categories:
                    if pattern.search(strict_signature):
                        final_match, final_id = "Subtype missing in input", "REQUIRES HUMAN"
                        broad_match_found = True
                        break
                if not broad_match_found:
                    final_match, final_id = "No good match", "REQUIRES HUMAN"

            return semantic_phrase, final_match, final_id

        processed_count = [0]
        progress_lock = threading.Lock()

        def worker(i_queue):
            semantic_phrase, final_match, final_id = process_single_ai_group(i_queue)
            sample_original = grouped_tasks[semantic_phrase]["sample_original"]
            
            with progress_lock:
                processed_count[0] += 1
                current = processed_count[0]
            
            if final_id == "REQUIRES HUMAN" or final_match == "No good match":
                ai_result = ask_maintenance_director_ollama(sample_original, local_lookup_phrases_original)
                
                if ai_result and ai_result.get("matched_asset"):
                    conf = str(ai_result.get("confidence", "")).strip().lower()
                    raw_ollama_guess = ai_result["matched_asset"]
                    strict_match = validate_and_format(raw_ollama_guess)
                    
                    if conf == "low" or strict_match == "No good match":
                        print(f"[{current}/{total_groups}] Ollama loosely guessed group '{semantic_phrase}': '{raw_ollama_guess}'. Rejecting.")
                        final_match = "No good match"
                        final_id = "REQUIRES HUMAN"
                    else:
                        final_match = strict_match
                        final_id = "AI_OLLAMA_DIRECTOR"
                        print(f"[{current}/{total_groups}] Ollama fixed group '{semantic_phrase}' ({conf}) -> {final_match}")
                else:
                    print(f"[{current}/{total_groups}] Vector/Ollama unresolved for group '{semantic_phrase}'")
            else:
                print(f"[{current}/{total_groups}] Vector matched group '{semantic_phrase}' -> {final_match}")

            group_results = []
            for r_id in grouped_tasks[semantic_phrase]["rows"]:
                group_results.append({k_row: r_id, k_match: final_match, k_id: final_id})
            
            return group_results

        with ThreadPoolExecutor(max_workers=8) as executor:
            futures = {executor.submit(worker, i): i for i in range(total_groups)}
            for future in as_completed(futures):
                results.extend(future.result())

    return results

# ==========================================
# 3. ENDPOINTS
# ==========================================
@app.route('/reload', methods=['POST'])
def reload_knowledge_base():
    try:
        load_knowledge_base()
        return jsonify({"status": "success"}), 200
    except Exception as e:
        return jsonify({"status": "error"}), 500

@app.route('/delete_rule', methods=['POST'])
def delete_rule():
    data = request.json
    raw_phrase = data.get('phrase', '').strip()
    signature = prescrub_text(raw_phrase, use_whitelist=True, current_valid_words=uniformat_valid_words, current_stemmed_whitelist=stemmed_whitelist, current_pattern=master_rule_pattern)

    if not signature:
        return jsonify({"status": "ignored", "message": "Invalid signature"}), 400

    deleted = False
    with custom_rules_lock:
        if signature in custom_rules:
            del custom_rules[signature]
            deleted = True
            
    if deleted:
        recompile_custom_rules_global()
        return jsonify({"status": "success", "message": "Rule deleted from memory"}), 200

    return jsonify({"status": "not_found"}), 404

@app.route('/batch_learn', methods=['POST'])
def batch_learn():
    data = request.json
    items = data.get('items', [])
    
    if not items:
        return jsonify({"status": "ignored", "message": "No items provided"}), 400

    learned_count = 0

    with custom_rules_lock:
        for item in items:
            raw_phrase = item.get('phrase', '').strip()
            clean_match = item.get('match', '').strip()

            corrected_match = canonical_casing.get(clean_match.lower(), clean_match)
            signature = prescrub_text(raw_phrase, use_whitelist=True, current_valid_words=uniformat_valid_words, current_stemmed_whitelist=stemmed_whitelist, current_pattern=master_rule_pattern)

            if signature and corrected_match:
                custom_rules[signature] = corrected_match
                learned_count += 1

    if learned_count > 0:
        recompile_custom_rules_global()

    return jsonify({"status": "success", "learned_count": learned_count}), 200

@app.route('/batch_lookup', methods=['POST'])
def batch_lookup():
    items = request.json.get('items', [])
    results = run_ai_batch_processing(items, is_batch_file=False)
    return jsonify(results)

@app.route('/batch_file', methods=['POST'])
def batch_file():
    data = request.json
    input_file = data.get('input_file')
    output_file = data.get('output_file')

    if not os.path.exists(input_file):
        return jsonify({"status": "error", "message": "Input file not found"}), 400

    df = pd.read_csv(input_file, sep='\t', dtype=str, encoding='utf-8-sig')
    records = df.to_dict('records')
    
    results = run_ai_batch_processing(records, is_batch_file=True)

    for r in results:
        if isinstance(r.get("Match"), str):
            r["Match"] = r["Match"].replace('\n', '\\n')

    df_out = pd.DataFrame(results)
    df_out.to_csv(output_file, sep='\t', index=False, encoding='utf-8-sig')

    return jsonify({"status": "success", "processed": len(results)}), 200

@app.route('/prescrub_file', methods=['POST'])
def prescrub_file():
    data = request.json
    input_file = data.get('input_file')
    output_file = data.get('output_file')

    if not os.path.exists(input_file):
        return jsonify({"status": "error", "message": "Input file not found"}), 400

    df = pd.read_csv(input_file, sep='\t', dtype=str, encoding='utf-8-sig')
    
    records = df.to_dict('records')
    processed_data = []
    
    with custom_rules_lock:
        local_compiled_custom_rules = compiled_custom_rules
        
    with knowledge_base_lock:
        local_valid_words = uniformat_valid_words
        local_stemmed_whitelist = stemmed_whitelist
        local_master_rule_pattern = master_rule_pattern
    
    for row in records:
        row_id = str(row.get('Row', ''))
        building = str(row.get('Building', '')).strip()
        original_phrase = str(row.get('Phrase', '')).strip()
        if original_phrase == 'nan': original_phrase = ''
        
        extracted_loc, text_no_loc = extract_locations(original_phrase)
        asset_id, base_phrase = extract_asset_identifiers(text_no_loc)
        
        scrubbed = prescrub_text(base_phrase, use_whitelist=True, 
                                 current_valid_words=local_valid_words, 
                                 current_stemmed_whitelist=local_stemmed_whitelist,
                                 current_pattern=local_master_rule_pattern)
        
        eval_phrase = scrubbed
        match_found = custom_rules.get(scrubbed)
        
        if not match_found:
            for rule_regex, rule_val, _ in local_compiled_custom_rules:
                if rule_regex.search(scrubbed):
                    match_found = rule_val
                    break
        
        if match_found:
            eval_phrase = match_found 
        
        sys_group, hierarchy = categorize_asset(eval_phrase)
        pm_suitability = determine_pm_suitability(eval_phrase)
        
        processed_data.append({
            "Row": row_id,
            "Building": building,
            "Original Phrase": original_phrase,
            "Scrubbed Phrase": scrubbed, 
            "Asset Tag": asset_id,
            "Location": extracted_loc,
            "System Group": sys_group,
            "Hierarchy": hierarchy,
            "Audit Flag": "",
            "PM Suitability": pm_suitability
        })
        
    df_processed = pd.DataFrame(processed_data)
    
    df_processed['__has_boiler'] = df_processed['Scrubbed Phrase'].str.contains(r'\bboiler\b|\bblr\b', case=False, regex=True)
    
    for bldg in df_processed['Building'].unique():
        if not bldg or str(bldg).strip().lower() == 'nan': continue
        bldg_mask = df_processed['Building'] == bldg
        
        ambiguous_mask = bldg_mask & (df_processed['System Group'] == 'Ambiguous Fuel System')
        
        if ambiguous_mask.any():
            has_fp = (bldg_mask & (df_processed['System Group'] == 'Fire Pump System') & (df_processed['Hierarchy'] == 'Parent')).any()
            has_gen = (bldg_mask & (df_processed['System Group'] == 'Emergency Power') & (df_processed['Hierarchy'] == 'Parent')).any()
            has_boiler = (bldg_mask & df_processed['__has_boiler']).any()
            
            phrase_lower = df_processed.loc[ambiguous_mask, 'Scrubbed Phrase'].str.lower()
            is_day_tank = phrase_lower.str.contains('day tank', regex=False)
            
            day_tank_indices = ambiguous_mask & is_day_tank
            other_indices = ambiguous_mask & ~is_day_tank
            
            if has_fp and not has_gen:
                df_processed.loc[day_tank_indices, ['System Group', 'Hierarchy']] = ['Fire Pump System', 'Child']
            if has_gen and not has_fp:
                df_processed.loc[day_tank_indices, ['System Group', 'Hierarchy']] = ['Emergency Power', 'Child']
                
            if has_fp and not has_gen and not has_boiler:
                df_processed.loc[other_indices, ['System Group', 'Hierarchy']] = ['Fire Pump System', 'Child']
            elif has_gen and not has_fp and not has_boiler:
                df_processed.loc[other_indices, ['System Group', 'Hierarchy']] = ['Emergency Power', 'Child']
            elif has_boiler and not has_fp and not has_gen:
                df_processed.loc[other_indices, ['System Group', 'Hierarchy']] = ['Boiler System', 'Child']

    for bldg in df_processed['Building'].unique():
        if not bldg or str(bldg).strip().lower() == 'nan': continue
        bldg_mask = df_processed['Building'] == bldg
        
        for rule in HIERARCHY_RULES:
            group_name = rule['group']
            group_mask = bldg_mask & (df_processed['System Group'] == group_name)
            
            if not group_mask.any(): continue
                
            parents = df_processed[group_mask & (df_processed['Hierarchy'] == 'Parent')]
            children = df_processed[group_mask & (df_processed['Hierarchy'] == 'Child')]
            
            p_count, c_count = len(parents), len(children)
            audit_msg = ""
            
            if group_name == "Split HVAC":
                if p_count > 0 and c_count == 0: 
                    audit_msg = "⚠️ Split HVAC: Missing Child (e.g., Fan Coil, Indoor Unit)"
                elif c_count > 0 and p_count == 0: 
                    audit_msg = "⚠️ Split HVAC: Missing Parent (e.g., Condenser, Outdoor Unit)"
                elif p_count > 0 and c_count > 0: 
                    audit_msg = "Matched Split HVAC System"
                    
            elif group_name in ["Life Safety / Fire", "Emergency Power", "Fire Pump System", "Vertical Transport", "Fire Sprinkler"]:
                parent_labels = {
                    "Life Safety / Fire": "Parent (e.g., Main FACP, Fire Panel)",
                    "Emergency Power": "Parent (e.g., Emergency Generator)",
                    "Fire Pump System": "Parent (e.g., Main Fire Pump)",
                    "Vertical Transport": "Parent (e.g., Elevator Cab, Escalator)",
                    "Fire Sprinkler": "Parent (e.g., Sprinkler Riser System)"
                }
                
                flag_text = parent_labels.get(group_name, "Parent Asset")
                
                if c_count > 0 and p_count == 0: 
                    audit_msg = f"⚠️ {group_name}: Missing {flag_text}"
                elif p_count > 0: 
                    audit_msg = f"Valid {group_name} System"
                
            if audit_msg: df_processed.loc[group_mask, 'Audit Flag'] = audit_msg

    df_processed.fillna("", inplace=True)
    df_processed = df_processed[["Row", "Building", "Original Phrase", "Scrubbed Phrase", "Asset Tag", "Location", "System Group", "Hierarchy", "Audit Flag", "PM Suitability"]]
    df_processed.to_csv(output_file, sep='\t', index=False, encoding='utf-8-sig')

    return jsonify({"status": "success", "processed": len(df_processed)}), 200

if __name__ == '__main__':
    app.run(port=5000, threaded=True)