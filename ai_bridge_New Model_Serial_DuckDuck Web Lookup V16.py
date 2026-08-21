# Base architecture and logic adapted for Enterprise Bulk Processing
from flask import Flask, request, jsonify
from sentence_transformers import SentenceTransformer
from functools import lru_cache
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
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
import time
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

app = Flask(__name__)

# ==========================================
# 1. INITIALIZE ENGINE, CACHE, & SESSIONS
# ==========================================
print("Loading Pure Semantic Math Engine...")

device = "cuda" if torch.cuda.is_available() else "cpu"
print(f"Engine forcing hardware acceleration on: [{device.upper()}]")

embedder = SentenceTransformer('all-MiniLM-L6-v2', device=device)

# --- NEW: ENTERPRISE HTTP CONNECTION POOLING ---
http_session = requests.Session()
retries = Retry(total=3, backoff_factor=1, status_forcelist=[500, 502, 503, 504])
adapter = HTTPAdapter(pool_connections=100, pool_maxsize=100, max_retries=retries)
http_session.mount('http://', adapter)
http_session.mount('https://', adapter)

script_dir = os.path.dirname(os.path.abspath(__file__))
appdata = os.environ.get('APPDATA', '')
if not appdata:
    appdata = os.path.expanduser('~') 

custom_rules_lock = threading.Lock()
knowledge_base_lock = threading.Lock()
typo_cache_lock = threading.Lock()

# Multi-Tier Persistent Caching
model_cache_lock = threading.Lock()
global_model_cache = {}  

serial_cache_lock = threading.Lock()
global_serial_cache = {}

# --- NEW: DUCKDUCKGO RATE LIMIT SAFETY ---
ddg_rate_limited = False
ddg_lock = threading.Lock()

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

custom_rules = {}
compiled_custom_rules = [] 
typo_cache = {}

def recompile_custom_rules_global():
    global compiled_custom_rules
    with custom_rules_lock:
        sorted_custom_rules = sorted(custom_rules.items(), key=lambda x: len(x[0]), reverse=True)
        compiled_custom_rules = [
            (re.compile(r'(?<![a-z0-9])' + re.escape(k) + r'(?:s|es)?(?![a-z0-9])', re.IGNORECASE), v, get_stemmed_words(k)) 
            for k, v in sorted_custom_rules if k
        ]

# ==========================================
# PERSISTENT MEMORY FUNCTIONS
# ==========================================
def load_persistent_caches():
    global global_model_cache, global_serial_cache
    model_cache_file = os.path.join(appdata, "Uniformat_Model_Cache.json")
    serial_cache_file = os.path.join(appdata, "Uniformat_Serial_Cache.json")
    
    with model_cache_lock:
        if os.path.exists(model_cache_file):
            try:
                with open(model_cache_file, 'r', encoding='utf-8') as f:
                    global_model_cache.update(json.load(f))
            except Exception as e:
                print(f"Error loading model cache: {e}")
                
    with serial_cache_lock:
        if os.path.exists(serial_cache_file):
            try:
                with open(serial_cache_file, 'r', encoding='utf-8') as f:
                    global_serial_cache.update(json.load(f))
            except Exception as e:
                print(f"Error loading serial cache: {e}")

def save_persistent_caches():
    model_cache_file = os.path.join(appdata, "Uniformat_Model_Cache.json")
    serial_cache_file = os.path.join(appdata, "Uniformat_Serial_Cache.json")
    
    with model_cache_lock:
        try:
            with open(model_cache_file, 'w', encoding='utf-8') as f:
                json.dump(global_model_cache, f, indent=4)
        except Exception:
            pass
            
    with serial_cache_lock:
        try:
            with open(serial_cache_file, 'w', encoding='utf-8') as f:
                json.dump(global_serial_cache, f, indent=4)
        except Exception:
            pass

# ==========================================
# PRE-COMPILED REGEX PATTERNS
# ==========================================
COMPILED_LOCATION_PATTERNS = [
    re.compile(r'\b(?:elevator )?lobby\b', re.IGNORECASE),
    re.compile(r'\b(?:stair|stairwell|staircase)(?:\s+\d+[a-z]?|\s+[a-z]\b)?\b', re.IGNORECASE),
    re.compile(r'\b(?:vestibule|corridor|hallway|hall|concourse|mezzanine|mezz|penthouse|ph|basement|roof|ceiling|atrium|closet)\b', re.IGNORECASE),
    re.compile(r'\b(?:parking )?garage\b', re.IGNORECASE),
    re.compile(r'\b(?:level|floor|fl)\s*(?:qty\.?|#|no\.?|quantity)?\s*\d+[a-z]*\b', re.IGNORECASE),
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

HIERARCHY_RULES = [
    {"group": "Life Safety / Fire", "parents": [r'\bfire alarm control', r'\bfire panel', r'\bfacp\b', r'\bannunciator', r'\bmain fire'], "children": [r'\bsmoke', r'\bdetector', r'\bstrobe', r'\bhorn', r'\bpull', r'\bbell', r'\btamper switch', r'\bflow switch', r'\binterface device']},
    {"group": "Emergency Power", "parents": [r'\bgenerator', r'\bgenset', r'\bdiesel gen', r'\bgas gen'], "children": [r'\bats\b', r'\bautomatic transfer switch', r'\btransfer switch', r'\bload bank']},
    {"group": "Vertical Transport", "parents": [r'\belevator', r'\bcab\b', r'\bescalator', r'\bmoving walk', r'\bwheelchair lift', r'\bdumbwaiter'], "children": [r'\belevator controller', r'\bhoistway', r'\bpit equipment', r'\belevator machine room']},
    {"group": "Hydronic HVAC", "parents": [r'\bhydronic fan coil'], "children": []},
    {"group": "Split HVAC", "parents": [r'\bcondensing unit', r'\bcondenser', r'\bcu\b', r'\boutdoor unit', r'\bvrf outdoor', r'\bvrv outdoor', r'\bheat pump outdoor'], "children": [r'\bfan coil', r'\bfcu\b', r'\bevaporator', r'\bvrf indoor', r'\bvrv indoor', r'\bcassette', r'\bwall mount split', r'\bindoor unit']},
    {"group": "Fire Pump System", "parents": [r'\bfire pump', r'\belectric fire pump', r'\bdiesel fire pump'], "children": [r'\bjockey pump', r'\bpressure maintenance pump', r'\bfire pump controller']},
    {"group": "Fire Sprinkler", "parents": [r'\bsprinkler system', r'\bwet pipe', r'\bdry pipe', r'\bpre-action', r'\bdeluge'], "children": [r'\bsprinkler head', r'\bdrop\b']},
    {"group": "Water Cooled Heat Pump", "parents": [r'\bwater cooled heat pump'], "children": []},
    {"group": "Ambiguous Fuel System", "parents": [], "children": [r'\bday tank', r'\bfuel polishing', r'\bfuel tank', r'\bfuel oil tank', r'\btransfer pump']}
]

COMPILED_HIERARCHY_RULES = []
for rule in HIERARCHY_RULES:
    compiled_parents = [re.compile(r'(?<![a-z0-9])' + p.replace(r'\b', '') + r'(?![a-z0-9])', re.IGNORECASE) for p in rule["parents"]]
    compiled_children = [re.compile(r'(?<![a-z0-9])' + c.replace(r'\b', '') + r'(?![a-z0-9])', re.IGNORECASE) for c in rule["children"]]
    COMPILED_HIERARCHY_RULES.append({"group": rule["group"], "parents": compiled_parents, "children": compiled_children})

@lru_cache(maxsize=10000)
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
    return frozenset(stemmed)

def validate_and_format(match_str):
    if not match_str or match_str in ["No good match", "REQUIRES HUMAN", "SKIP", "Out of Scope (Architectural/Civil)", "Subtype missing in input"]:
        return match_str
    
    if "\n" in match_str:
        valid_parts = []
        for part in match_str.split("\n"):
            part_clean = part.strip()
            clean = canonical_casing.get(part_clean.lower())
            if clean and clean not in valid_parts: valid_parts.append(clean)
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

    t_lookup_phrases, t_lookup_phrases_original, t_canonical_casing = [], [], {}
    t_industry_translations, t_list_translations, t_broad_categories = {}, {}, []
    t_uniformat_valid_words = set()
    t_stemmed_whitelist = set()
    t_compiled_expert_rules = []
    t_compiled_list_rules = []
    t_compiled_broad_categories = []
    paired = []

    custom_template_dir = os.path.expanduser(r"~\OneDrive\Documents\Custom Office Templates")
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
            xls = pd.ExcelFile(latest_file, engine='openpyxl')
            
            if "Uniformat RS Means Lookup" in xls.sheet_names:
                print("Reading Lookups from memory...")
                df_lookups = pd.read_excel(xls, sheet_name="Uniformat RS Means Lookup", header=None)
                df_lookups_records = df_lookups.dropna(how='all').to_dict('records')
                for row in df_lookups_records:
                    if 0 in row and pd.notna(row[0]):
                        val = str(row[0]).strip()
                        if val.lower() == "asset sub type": continue 
                        clean_phrase = val
                        if clean_phrase:
                            lower_phrase = clean_phrase.lower()
                            stripped_phrase = NON_ALNUM_PATTERN.sub(' ', lower_phrase)
                            stripped_phrase = MULTI_SPACE_PATTERN.sub(' ', stripped_phrase).strip()
                            paired.append((stripped_phrase, clean_phrase))
                            t_canonical_casing[lower_phrase] = clean_phrase
                            for word in stripped_phrase.split():
                                t_uniformat_valid_words.add(word)

            if "Rules Sheet" in xls.sheet_names:
                print("Reading Rules from memory...")
                df_rules = pd.read_excel(xls, sheet_name="Rules Sheet", header=None)
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
                                for word in re.findall(r'[a-z0-9]+', to_text.lower()): t_uniformat_valid_words.add(word)
                                for word in re.findall(r'[a-z0-9]+', clean_from): t_uniformat_valid_words.add(word)
                        
                        elif rule_type == "list":
                            clean_from = MULTI_SPACE_PATTERN.sub(' ', from_text).strip()
                            if clean_from:
                                t_list_translations[clean_from] = to_text
                                for word in re.findall(r'[a-z0-9]+', to_text.lower()): t_uniformat_valid_words.add(word)
                                for word in re.findall(r'[a-z0-9]+', clean_from): t_uniformat_valid_words.add(word)
                                    
                        elif rule_type == "core":
                            clean_from = MULTI_SPACE_PATTERN.sub(' ', from_text).strip()
                            if clean_from:
                                if clean_from not in t_broad_categories: t_broad_categories.append(clean_from)
                                for word in re.findall(r'[a-z0-9]+', clean_from): t_uniformat_valid_words.add(word)

            if "LearnedMappings" in xls.sheet_names:
                print("Reading Learned Mappings from memory...")
                df_learned = pd.read_excel(xls, sheet_name="LearnedMappings", header=None)
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
                                if clean_k: custom_rules[clean_k] = rule_v

            xls.close()

        except Exception as e:
            print(f"Excel Memory Load Error: {e}")

    paired.sort(key=lambda x: len(x[0]), reverse=True)
    t_lookup_phrases = [p[0] for p in paired]
    t_lookup_phrases_original = [p[1] for p in paired]

    for rule in HIERARCHY_RULES:
        for p in rule['parents'] + rule['children']:
            clean_p = p.replace(r'\b', '')
            for w in re.findall(r'[a-z0-9]+', clean_p.lower()): t_uniformat_valid_words.add(w)

    if not t_broad_categories:
        t_broad_categories = ["pump", "fan", "boiler", "furnace", "transformer", "compressor", "chiller", "motor", "sump", "sprinkler", "valve", "hydrant", "tower", "exchanger", "tank"]
        for cat in t_broad_categories:
            for word in re.findall(r'[a-z0-9]+', cat): t_uniformat_valid_words.add(word)

    sorted_messy_terms = sorted(t_industry_translations.keys(), key=len, reverse=True)
    if sorted_messy_terms:
        escaped_keys = [r'(?<![a-z0-9])' + re.escape(k) + r'(?:s|es)?(?![a-z0-9])' for k in sorted_messy_terms]
        t_master_rule_pattern = re.compile('(' + '|'.join(escaped_keys) + ')', flags=re.IGNORECASE)
    else:
        t_master_rule_pattern = None

    for messy_term in sorted_messy_terms:
        clean_term = t_industry_translations[messy_term]
        if clean_term:
            clean_messy = NON_ALNUM_PATTERN.sub(' ', messy_term)
            rule_tokens = get_stemmed_words(clean_messy)
            compiled_pat = re.compile(r'(?<![a-z0-9])' + re.escape(messy_term) + r'(?:s|es)?(?![a-z0-9])', re.IGNORECASE)
            t_compiled_expert_rules.append((compiled_pat, messy_term, rule_tokens, clean_term))

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

    load_persistent_caches()
    recompile_custom_rules_global()
    print(f"Engine Ready! Whitelist contains {len(uniformat_valid_words)} authorized words.")

load_knowledge_base()

# ==========================================
# 2. THE UNIFIED PRE-SCRUB & LOCAL AI ENGINE
# ==========================================
def extract_locations(raw_text):
    if not isinstance(raw_text, str): return "", raw_text
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
    if not isinstance(raw_text, str): return "", raw_text
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
    if not isinstance(scrubbed_text, str) or not scrubbed_text: return ""
    if not current_valid_words: return scrubbed_text
        
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
                cached = w in typo_cache
                cache_val = typo_cache.get(w)
                    
                if cached:
                    if cache_val: pure_words.append(cache_val)
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
            if '\n' in target_text: return match_word
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
    "Non-PM": [r'\bwindow\b', r'\bexit sign\b', r'\bpipe\b', r'\bmanual valve\b', r'\bductwork\b', r'\bwall\b', r'\bwashroom\b', r'\bjanitor\b'],
    "Questionable": [r'\bday tank\b', r'\bmeter\b', r'\bsensor\b', r'\bdoor\b', r'\bportable\b']
}

COMPILED_PM_SUITABILITY_RULES = {}
for status, patterns in PM_SUITABILITY_RULES.items():
    compiled_patterns = []
    for pattern in patterns:
        strict_pattern = r'(?<![a-z0-9])' + pattern.replace(r'\b', '') + r'(?![a-z0-9])'
        compiled_patterns.append(re.compile(strict_pattern, re.IGNORECASE))
    COMPILED_PM_SUITABILITY_RULES[status] = compiled_patterns

def determine_pm_suitability(scrubbed_phrase):
    for status, compiled_patterns in COMPILED_PM_SUITABILITY_RULES.items():
        for pattern in compiled_patterns:
            if pattern.search(scrubbed_phrase): return status
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
    
    prompt = f"""You are an expert mechanical engineer and HVAC equipment specialist.
Analyze the Client Asset Label to classify the equipment. Do NOT guess or estimate technical specifications.

INPUT DATA:
- CLIENT LABEL / DESCRIPTION: "{raw_phrase}"

INSTRUCTIONS:
1. Identify the core equipment type (e.g., Boiler, Air Handling Unit, Centrifugal Pump).
2. DO NOT hallucinate voltages, capacities, or drive types. 
3. If specific specs are not explicitly present in the label, you MUST output "DATA_MISSING".

Respond ONLY in valid JSON with these exact keys:
{{
    "matched_asset": "[Clean Asset Type Here]",
    "confidence": "[High, Medium, or Low]",
    "technical_specs": "DATA_MISSING",
    "consumables": "DATA_MISSING",
    "lifecycle_data": "DATA_MISSING"
}}
"""
    payload = {"model": "qwen2.5:7b", "prompt": prompt, "format": "json", "stream": False, "options": {"temperature": 0.0}}
    try:
        response = http_session.post(url, json=payload, timeout=300)
        if response.status_code == 200:
            raw_res = response.json().get("response", "").strip()
            
            if raw_res.startswith("```"):
                raw_res = re.sub(r"^```(?:json)?\n?", "", raw_res)
                raw_res = re.sub(r"\n?```$", "", raw_res)
            return json.loads(raw_res.strip())
            
    except Exception as e:
        print(f"Ollama API Error: {e}")
    return None

KNOWN_MANUFACTURERS = {
    "carrier", "trane", "lennox", "york", "daikin", "greenheck", "rheem", 
    "ruud", "ao smith", "bradford white", "mcquay", "dunham bush", "liebert", 
    "mitsubishi", "fujitsu", "goodman", "bryant", "payne", "cleaver brooks",
    "viessmann", "haakon", "cook", "unknown", "n/a", "none", "nan", "unit", "model"
}

def validate_model_and_serial(manufacturer, model_num, serial_num):
    clean_mfg = str(manufacturer).strip()
    clean_model = str(model_num).strip()
    clean_serial = str(serial_num).strip()
    
    def is_real(val):
        return val and val.lower() not in ["nan", "none", "n/a", "unknown", "null", "undefined"]

    clean_mfg = clean_mfg if is_real(clean_mfg) else ""
    clean_model = clean_model if is_real(clean_model) else ""
    clean_serial = clean_serial if is_real(clean_serial) else ""

    if clean_mfg:
        clean_mfg = re.sub(r'^(?:MAKE|MFG|BRAND)[:\-]?\s*', '', clean_mfg, flags=re.IGNORECASE).strip()
        clean_mfg = re.sub(r'-[A-Za-z0-9]+$', '', clean_mfg).strip()
        
    if clean_model:
        clean_model = re.sub(r'^MODEL[:\-]?\s*', '', clean_model, flags=re.IGNORECASE).strip()

    brand_context = clean_mfg

    if clean_model and clean_model.lower() in KNOWN_MANUFACTURERS:
        if not brand_context:
            brand_context = clean_model
        clean_model = "" 

    return clean_model, clean_serial, brand_context

# ==========================================
# TIER 1: GLOBALLY CACHED MODEL LOOKUP
# ==========================================
def extract_model_data_global(brand_context, clean_model, base_phrase):
    global ddg_rate_limited
    url = "http://localhost:11434/api/generate"
    search_context = ""
    
    if clean_model and not ddg_rate_limited:
        try:
            from ddgs import DDGS
            search_query_parts = [brand_context, clean_model, base_phrase, "HVAC", "specifications"]
            query = " ".join([p for p in search_query_parts if p]).strip()
            
            with DDGS() as ddgs_client:
                results = list(ddgs_client.text(query, max_results=3)) 
                if results:
                    snippets = [f"- {r.get('title', '')}: {r.get('body', '')}" for r in results]
                    search_context = "\n".join(snippets)
        except Exception as e:
            if "rate" in str(e).lower() or "202" in str(e):
                with ddg_lock:
                    if not ddg_rate_limited:
                        ddg_rate_limited = True
                        print("\n[WARNING] DuckDuckGo Rate Limit Hit! Engaging safety lock.")
                        print("Falling back to pure Ollama intelligence for remaining models to prevent IP ban.\n")

    prompt = f"""You are an expert mechanical engineer. Analyze the inputs to determine exact technical specs and consumables for the model provided.

INPUT DATA:
- MANUFACTURER: "{brand_context if brand_context else 'Not Specified'}"
- MODEL NUMBER: "{clean_model if clean_model else 'Not Specified'}"
- CLIENT LABEL: "{base_phrase}"
- WEB CONTEXT: "{search_context}"

INSTRUCTIONS:
1. Identify the equipment identity.
2. Extract technical specs strictly from the web context.
3. If web context is missing or empty, use your internal HVAC engineering knowledge to decode the Model Nomenclature. 
4. DO NOT GUESS OR HALLUCINATE standard defaults if the attribute cannot be logically deduced from the model string or web context.
5. If an attribute (Voltage, Type, Drive, Heating, Filters, Belts) is completely unidentifiable, output "DATA_MISSING".

Respond ONLY in valid JSON with these exact keys:
{{
    "equipment_identity": "[Asset Type]",
    "technical_specs": "Voltage: [X] | Type: [X] | Drive: [X] | Heating: [X]",
    "consumables": "Filters: [X] | Belts: [X]",
    "debug_source": "[State if you used Web Context, Nomenclature Decode, or None]"
}}
"""
    payload = {"model": "qwen2.5:7b", "prompt": prompt, "format": "json", "stream": False, "options": {"temperature": 0.0}}
    
    try:
        response = http_session.post(url, json=payload, timeout=120)
        if response.status_code == 200:
            raw_res = response.json().get("response", "").strip()
            if raw_res.startswith("```"):
                raw_res = re.sub(r"^```(?:json)?\n?", "", raw_res)
                raw_res = re.sub(r"\n?```$", "", raw_res)
                
            data = json.loads(raw_res.strip())
            identity = data.get("equipment_identity", "").strip()
            
            bad_phrases = ["Insert", "Unknown", "Clean Asset", "Equipment Type"]
            if any(bp in identity for bp in bad_phrases) or not identity:
                identity = base_phrase

            return {
                "identity": identity,
                "technical_specs": data.get("technical_specs", "DATA_MISSING").strip(),
                "consumables": data.get("consumables", "DATA_MISSING").strip()
            }
    except Exception as e:
        pass
        
    return {
        "identity": base_phrase, 
        "technical_specs": "DATA_MISSING",
        "consumables": "DATA_MISSING"
    }

# ==========================================
# TIER 2: INDIVIDUAL ROW SERIAL LOOKUP
# ==========================================
def decode_serial_data_local(brand_context, clean_serial):
    if not clean_serial:
        return "DATA_MISSING"
        
    url = "http://localhost:11434/api/generate"
    prompt = f"""You are an HVAC serial number decoder. 
Determine the year of manufacture from the given serial number based on standard industry manufacturer coding logic.

INPUT DATA:
- MANUFACTURER: "{brand_context}"
- SERIAL NUMBER: "{clean_serial}"

INSTRUCTIONS: 
1. Determine the year of manufacture.
2. If it is impossible to decode the year, output exactly "DATA_MISSING".

Respond ONLY in valid JSON with this exact key:
{{
    "lifecycle_data": "Manufactured: [Year] | Warranty: [Status]"
}}
"""
    payload = {"model": "qwen2.5:7b", "prompt": prompt, "format": "json", "stream": False, "options": {"temperature": 0.0}}
    try:
        response = http_session.post(url, json=payload, timeout=120)
        if response.status_code == 200:
            raw_res = response.json().get("response", "").strip()
            if raw_res.startswith("```"):
                raw_res = re.sub(r"^```(?:json)?\n?", "", raw_res)
                raw_res = re.sub(r"\n?```$", "", raw_res)
            data = json.loads(raw_res.strip())
            return data.get("lifecycle_data", "DATA_MISSING").strip()
    except Exception as e:
        pass
    return "DATA_MISSING"

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

    unique_models = {}
    for row in items:
        if is_batch_file:
            mfg = str(row.get('Manufacturer', '')).strip()
            m_num = str(row.get('Model', '')).strip()
            o_phr = str(row.get('Phrase', '')).strip()
        else:
            mfg = str(row.get('manufacturer', '')).strip()
            m_num = str(row.get('model', '')).strip()
            o_phr = str(row.get('phrase', '')).strip()
            
        extracted_loc, text_no_loc = extract_locations(o_phr)
        _, base_phrase = extract_asset_identifiers(text_no_loc)
        clean_model, _, brand_context = validate_model_and_serial(mfg, m_num, "")

        if clean_model or brand_context or base_phrase:
            model_key = f"{brand_context}|{clean_model}|{base_phrase}"
            with model_cache_lock:
                if model_key not in global_model_cache:
                    unique_models[model_key] = (brand_context, clean_model, base_phrase)

    if unique_models:
        def fetch_model_tier(m_key, brand, model_str, b_phrase):
            tier1_data = extract_model_data_global(brand, model_str, b_phrase)
            with model_cache_lock:
                global_model_cache[m_key] = tier1_data

        with ThreadPoolExecutor(max_workers=5) as executor: 
            futures = [executor.submit(fetch_model_tier, k, b, m, p) for k, (b, m, p) in unique_models.items()]
            for future in as_completed(futures): pass

    for row in items:
        if is_batch_file:
            row_id = str(row.get('Row', ''))
            original_phrase = str(row.get('Phrase', '')).strip()
            manufacturer = str(row.get('Manufacturer', '')).strip()
            model_num = str(row.get('Model', '')).strip()
        else:
            row_id = row.get('row')
            original_phrase = row.get('phrase', '').strip()
            manufacturer = str(row.get('manufacturer', '')).strip()
            model_num = str(row.get('model', '')).strip()
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

        clean_model, _, brand_context = validate_model_and_serial(manufacturer, model_num, "")
        model_key = f"{brand_context}|{clean_model}|{clean_base_phrase}"
        
        with model_cache_lock:
            cached_data = global_model_cache.get(model_key, {"identity": ""})
            
        llm_identity = cached_data.get("identity", "")
        resolved_phrase = llm_identity if llm_identity else clean_base_phrase

        signature = prescrub_text(resolved_phrase, use_whitelist=True, 
                                  current_valid_words=local_valid_words, 
                                  current_stemmed_whitelist=local_stemmed_whitelist,
                                  current_pattern=local_master_rule_pattern)
        
        semantic_phrase = prescrub_text(resolved_phrase, use_whitelist=False, 
                                        current_valid_words=local_valid_words, 
                                        current_stemmed_whitelist=local_stemmed_whitelist,
                                        current_pattern=local_master_rule_pattern)

        original_clean = NON_ALNUM_PATTERN.sub(' ', resolved_phrase.lower())

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
            if len(expert_matches) > 1:
                exact_hits = []
                for opt in expert_matches:
                    opt_clean = NON_ALNUM_PATTERN.sub(' ', opt.lower())
                    opt_clean = MULTI_SPACE_PATTERN.sub(' ', opt_clean).strip()
                    if opt_clean in original_clean or opt_clean in semantic_phrase:
                        exact_hits.append(opt)
                
                if len(exact_hits) == 1:
                    final_match = exact_hits[0]
                    results.append({k_row: row_id, k_match: final_match, k_id: "EXPERT_RULE (Auto-Resolved)"})
                else:
                    final_match = "\n".join(expert_matches[:8])
                    results.append({k_row: row_id, k_match: final_match, k_id: "EXPERT_RULE (Multiple Options)"})
            else:
                final_match = expert_matches[0]
                results.append({k_row: row_id, k_match: final_match, k_id: "EXPERT_RULE"})
        else:
            EXCLUSION_PATTERN = re.compile(r'\b(?:structural|furniture|masonry|glazing|pavement|interlocking paver|playground|mural|pest control|washroom accessories|curbing|bollard|planter|ceiling tiles?|flooring|drywall|retaining walls?|landscape|landscaping|painting|paint|garbage recep|waste recep|ash urn|sidewalks?|parking lot|decks?|sweeper|window washing|recycling|charcoal|ramp plates?)\b', re.IGNORECASE)
            
            if EXCLUSION_PATTERN.search(clean_base_phrase):
                results.append({k_row: row_id, k_match: "Out of Scope (Architectural/Civil)", k_id: "AUTO_EXCLUDED"})
            else:
                queue_phrase = signature if len(signature) >= 2 else semantic_phrase
                if len(queue_phrase) >= 2:
                    ai_queue_phrases.append(queue_phrase)
                    ai_queue_rows.append(row_id)
                    original_phrases_list.append(resolved_phrase) 
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
                        
                        if len(lookup_words) >= 2 and intersection == len(lookup_words):
                            overlap_cache[idx] = 0.50 
                        else:
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
            
            if final_id == "REQUIRES HUMAN" or final_match == "No good match":
                ai_result = ask_maintenance_director_ollama(sample_original, local_lookup_phrases_original)
                
                if ai_result and ai_result.get("matched_asset"):
                    conf = str(ai_result.get("confidence", "")).strip().lower()
                    raw_ollama_guess = ai_result["matched_asset"]
                    strict_match = validate_and_format(raw_ollama_guess)
                    
                    if conf == "low" or strict_match == "No good match":
                        final_match = "No good match"
                        final_id = "REQUIRES HUMAN"
                    else:
                        final_match = strict_match
                        final_id = "AI_OLLAMA_DIRECTOR"

            group_results = []
            for r_id in grouped_tasks[semantic_phrase]["rows"]:
                group_results.append({k_row: r_id, k_match: final_match, k_id: final_id})
            
            return group_results

        with ThreadPoolExecutor(max_workers=5) as executor:
            futures = {executor.submit(worker, i): i for i in range(total_groups)}
            for future in as_completed(futures):
                results.extend(future.result())

    return results

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
    df.fillna("", inplace=True)
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
    df.fillna("", inplace=True)
    records = df.to_dict('records')
    processed_data = []
    
    with custom_rules_lock:
        local_compiled_custom_rules = compiled_custom_rules
        
    with knowledge_base_lock:
        local_valid_words = uniformat_valid_words
        local_stemmed_whitelist = stemmed_whitelist
        local_master_rule_pattern = master_rule_pattern

    unique_models = {}
    for row in records:
        manufacturer = str(row.get('Manufacturer', '')).strip()
        model_num = str(row.get('Model', '')).strip()
        original_phrase = str(row.get('Phrase', '')).strip()
        
        extracted_loc, text_no_loc = extract_locations(original_phrase)
        _, base_phrase = extract_asset_identifiers(text_no_loc)

        clean_model, _, brand_context = validate_model_and_serial(manufacturer, model_num, "")
        
        if clean_model or brand_context or base_phrase:
            model_key = f"{brand_context}|{clean_model}|{base_phrase}"
            
            with model_cache_lock:
                if model_key not in global_model_cache:
                    unique_models[model_key] = (brand_context, clean_model, base_phrase)

    total_models = len(unique_models)
    model_counter = [0]
    model_count_lock = threading.Lock()

    if total_models > 0:
        print(f"\n--- Batching {total_models} unique equipment models (Tier 1) ---")
        def fetch_model_tier(m_key, brand, model_str, b_phrase):
            tier1_data = extract_model_data_global(brand, model_str, b_phrase)
            with model_cache_lock:
                global_model_cache[m_key] = tier1_data
            
            with model_count_lock:
                model_counter[0] += 1
                print(f"[Tier 1: {model_counter[0]}/{total_models}] Processed: {brand} {model_str} -> {tier1_data.get('identity', '')}")

        with ThreadPoolExecutor(max_workers=5) as executor: 
            futures = [executor.submit(fetch_model_tier, k, b, m, p) for k, (b, m, p) in unique_models.items()]
            for future in as_completed(futures): pass
        print("\n--- Tier 1 Cache Complete! ---")

    unique_serials = {}
    for row in records:
        manufacturer = str(row.get('Manufacturer', '')).strip()
        serial_num = str(row.get('Serial', '')).strip()
        
        _, clean_serial, brand_context = validate_model_and_serial(manufacturer, "", serial_num)
        
        if clean_serial and brand_context:
            serial_key = f"{brand_context}|{clean_serial}"
            with serial_cache_lock:
                if serial_key not in global_serial_cache:
                    unique_serials[serial_key] = (brand_context, clean_serial)

    total_serials = len(unique_serials)
    serial_counter = [0]
    serial_count_lock = threading.Lock()

    if total_serials > 0:
        print(f"\n--- Batching {total_serials} unique serial numbers (Tier 2) ---")
        def fetch_serial_tier(s_key, brand, serial_str):
            tier2_data = decode_serial_data_local(brand, serial_str)
            with serial_cache_lock:
                global_serial_cache[s_key] = tier2_data
            
            with serial_count_lock:
                serial_counter[0] += 1
                print(f"[Tier 2: {serial_counter[0]}/{total_serials}] Decoded Serial: {brand} {serial_str} -> {tier2_data}")

        with ThreadPoolExecutor(max_workers=10) as executor: 
            futures = [executor.submit(fetch_serial_tier, k, b, s) for k, (b, s) in unique_serials.items()]
            for future in as_completed(futures): pass
        print("\n--- Tier 2 Cache Complete! Building final rows... ---\n")

    # ==========================================
    # FINAL PASS: Combine Data instantly (STRICT 1-to-1 INDEX)
    # ==========================================
    for row in records:
        row_id = str(row.get('Row', ''))
        building = str(row.get('Building', '')).strip()
        original_phrase = str(row.get('Phrase', '')).strip()
        manufacturer = str(row.get('Manufacturer', '')).strip()
        model_num = str(row.get('Model', '')).strip()
        serial_num = str(row.get('Serial', '')).strip()
        
        input_loc = str(row.get('Location', '')).strip()
        extracted_loc, text_no_loc = extract_locations(original_phrase)
        asset_id, base_phrase = extract_asset_identifiers(text_no_loc)
        
        final_loc = input_loc if input_loc else extracted_loc
        
        clean_model, clean_serial, brand_context = validate_model_and_serial(manufacturer, model_num, serial_num)
        
        model_key = f"{brand_context}|{clean_model}|{base_phrase}"
        with model_cache_lock:
            cached_data = global_model_cache.get(model_key, {
                "identity": "", 
                "technical_specs": "DATA_MISSING",
                "consumables": "DATA_MISSING"
            })
            
        llm_identity = cached_data.get("identity", "")
        technical_specs = cached_data.get("technical_specs", "")
        consumables = cached_data.get("consumables", "")

        serial_key = f"{brand_context}|{clean_serial}"
        with serial_cache_lock:
            lifecycle_data = global_serial_cache.get(serial_key, "DATA_MISSING") if clean_serial else "DATA_MISSING"
        
        resolved_phrase = llm_identity if llm_identity else base_phrase
        scrubbed = prescrub_text(resolved_phrase, use_whitelist=True, 
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
            "LLM Model Identity": llm_identity,
            "Technical Specs": technical_specs,  
            "Consumables": consumables,          
            "Lifecycle Data": lifecycle_data,    
            "Asset Tag": asset_id,
            "Location": final_loc,
            "System Group": sys_group,
            "Hierarchy": hierarchy,
            "Audit Flag": "",
            "PM Suitability": pm_suitability,
            "Yearly Effort (Hrs)": 0.0,
            "Yearly Material ($)": 0.0,
            "NFPA Group": ""
        })
        
    save_persistent_caches()
    df_processed = pd.DataFrame(processed_data)
    
    df_processed['__has_boiler'] = df_processed['Scrubbed Phrase'].str.contains(r'\bboiler\b|\bblr\b', case=False, regex=True, na=False)
    
    # ---------------------------------------------------------
    # PART 1: ORIGINAL HIERARCHY & AUDIT FLAG LOGIC (sort=False)
    # ---------------------------------------------------------
    for bldg, bldg_df in df_processed.groupby('Building', sort=False):
        if not bldg or str(bldg).strip().lower() == 'nan': continue
        
        ambiguous_mask = bldg_df['System Group'] == 'Ambiguous Fuel System'
        
        if ambiguous_mask.any():
            has_fp = ((bldg_df['System Group'] == 'Fire Pump System') & (bldg_df['Hierarchy'] == 'Parent')).any()
            has_gen = ((bldg_df['System Group'] == 'Emergency Power') & (bldg_df['Hierarchy'] == 'Parent')).any()
            has_boiler = bldg_df['__has_boiler'].any()
            
            phrase_lower = bldg_df.loc[ambiguous_mask, 'Scrubbed Phrase'].str.lower()
            is_day_tank = phrase_lower.str.contains('day tank', regex=False)
            
            day_tank_indices = bldg_df[ambiguous_mask][is_day_tank].index
            other_indices = bldg_df[ambiguous_mask][~is_day_tank].index
            
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

        for rule in HIERARCHY_RULES:
            group_name = rule['group']
            group_mask = bldg_df['System Group'] == group_name
            
            if not group_mask.any(): continue
                
            p_count = (bldg_df.loc[group_mask, 'Hierarchy'] == 'Parent').sum()
            c_count = (bldg_df.loc[group_mask, 'Hierarchy'] == 'Child').sum()
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
                
            if audit_msg: 
                group_indices = bldg_df[group_mask].index
                df_processed.loc[group_indices, 'Audit Flag'] = audit_msg

    # ---------------------------------------------------------
    # PART 2: NEW EFFORT HOUR & MATERIAL ROLL-UP LOGIC (sort=False)
    # ---------------------------------------------------------
    for bldg, bldg_df in df_processed.groupby('Building', sort=False):
        if not bldg or str(bldg).strip().lower() == 'nan': continue
            
        for group_name in ["Life Safety / Fire", "Emergency Power", "Fire Pump System", "Fire Sprinkler"]:
            group_mask = bldg_df['System Group'] == group_name
            if not group_mask.any(): continue
                
            parents = bldg_df[(group_mask) & (bldg_df['Hierarchy'] == 'Parent')]
            children = bldg_df[(group_mask) & (bldg_df['Hierarchy'] == 'Child')]
            
            p_count = len(parents)
            c_count = len(children)
            
            if p_count > 0:
                total_effort = 0.0
                total_mat = 0.0
                
                if group_name == "Life Safety / Fire":
                    total_effort = (p_count * 1.5) + (c_count * 0.05)
                    total_mat = (p_count * 100.0)
                elif group_name == "Fire Pump System":
                    diesel_mask = parents['Scrubbed Phrase'].str.contains(r'diesel|engine', case=False, regex=True, na=False)
                    diesel_count = diesel_mask.sum()
                    elec_count = p_count - diesel_count
                    total_effort = (diesel_count * 72.6) + (elec_count * 12.0) + (c_count * 2.78)
                    total_mat = (diesel_count * 327.0) + (elec_count * 142.0) + (c_count * 88.0)
                elif group_name == "Fire Sprinkler":
                    dry_mask = parents['Scrubbed Phrase'].str.contains(r'dry', case=False, regex=True, na=False)
                    dry_count = dry_mask.sum()
                    wet_count = p_count - dry_count
                    total_effort = (dry_count * 16.5) + (wet_count * 7.5)
                    total_mat = (dry_count * 191.0) + (wet_count * 110.0)
                elif group_name == "Emergency Power":
                    total_effort = (p_count * 10.0) + (c_count * 2.0)
                    total_mat = (p_count * 250.0)
                
                df_processed.loc[parents.index, 'Yearly Effort (Hrs)'] = round(total_effort / p_count, 2)
                df_processed.loc[parents.index, 'Yearly Material ($)'] = round(total_mat / p_count, 2)
                df_processed.loc[children.index, 'Yearly Effort (Hrs)'] = 0.0
                df_processed.loc[children.index, 'Yearly Material ($)'] = 0.0
            else:
                if group_name == "Life Safety / Fire":
                    df_processed.loc[children.index, 'Yearly Effort (Hrs)'] = 0.05
                elif group_name == "Fire Pump System":
                    df_processed.loc[children.index, 'Yearly Effort (Hrs)'] = 2.78
                elif group_name == "Emergency Power":
                    df_processed.loc[children.index, 'Yearly Effort (Hrs)'] = 2.0

    # ---------------------------------------------------------
    # PART 2.5: STANDALONE ROUTE ASSET EFFORT HOURS
    # ---------------------------------------------------------
    phrase_col = df_processed['Scrubbed Phrase'].str.lower()
    df_processed.loc[phrase_col.str.contains('hose cabinet', regex=False, na=False), 'Yearly Effort (Hrs)'] = 0.26
    df_processed.loc[phrase_col.str.contains('extinguisher', regex=False, na=False), 'Yearly Effort (Hrs)'] = 0.10
    df_processed.loc[phrase_col.str.contains('fire door', regex=False, na=False), 'Yearly Effort (Hrs)'] = 0.73
    df_processed.loc[phrase_col.str.contains('damper', regex=False, na=False), 'Yearly Effort (Hrs)'] = 1.05
    df_processed.loc[phrase_col.str.contains('damper', regex=False, na=False), 'Yearly Material ($)'] = 14.30

    # ---------------------------------------------------------
    # PART 3: NFPA FIRE GROUPINGS & AUDIT FLAG ENHANCEMENT
    # ---------------------------------------------------------
    def get_nfpa_group(row_data):
        phrase = str(row_data.get('Scrubbed Phrase', '')).lower()
        sys_grp = str(row_data.get('System Group', ''))
        
        if any(x in phrase for x in ['damper', 'door', 'extinguisher', 'hose cabinet', 'fire box']):
            return "Group 3: Passive Fire Defenses & Portable Equipment (NFPA 80 / 10)"
            
        if sys_grp == "Life Safety / Fire":
            return "Group 1: Fire Alarm & Detection (NFPA 72)"
        elif sys_grp in ["Fire Pump System", "Fire Sprinkler"]:
            return "Group 2: Water-Based & Active Suppression (NFPA 25 / 17A)"
            
        return ""

    df_processed['NFPA Group'] = df_processed.apply(get_nfpa_group, axis=1)

    def enhance_audit_flag(row_data):
        flag = str(row_data.get('Audit Flag', '')).strip()
        nfpa = str(row_data.get('NFPA Group', '')).strip()
        
        if nfpa:
            if "Group 3" in nfpa:
                if flag:
                    return f"Valid Passive Fire Asset - {nfpa} | {flag}"
                else:
                    return f"Valid Passive Fire Asset - {nfpa}"
            else:
                if flag:
                    parts = flag.split(" | ")
                    parts[0] = f"{parts[0]} - {nfpa}"
                    return " | ".join(parts)
        return flag

    df_processed['Audit Flag'] = df_processed.apply(enhance_audit_flag, axis=1)

    # ---------------------------------------------------------
    # PART 4: BUILDING-LEVEL COMPLIANCE CHECKS (sort=False)
    # ---------------------------------------------------------
    for bldg, bldg_df in df_processed.groupby('Building', sort=False):
        if not bldg or str(bldg).strip().lower() == 'nan': continue
        
        fire_parents = bldg_df[(bldg_df['NFPA Group'].str.contains('Group 1|Group 2', na=False)) & (bldg_df['Hierarchy'] == 'Parent')]
        has_fire_parent = len(fire_parents) > 0
        has_backflow = bldg_df['Scrubbed Phrase'].str.contains(r'\bbackflow\b', case=False, regex=True, na=False).any()
        
        building_alerts = []
        if not has_fire_parent:
            building_alerts.append("⚠️ Missing Building-Wide Fire System (Parent)")
        if not has_backflow:
            building_alerts.append("⚠️ Missing Building Backflow Device")
            
        if building_alerts:
            first_idx = bldg_df.index[0]
            existing_flag = str(df_processed.loc[first_idx, 'Audit Flag']).strip()
            new_flag = " | ".join(building_alerts)
            
            if existing_flag and existing_flag.lower() != "nan":
                df_processed.loc[first_idx, 'Audit Flag'] = existing_flag + " | " + new_flag
            else:
                df_processed.loc[first_idx, 'Audit Flag'] = new_flag

    # ---------------------------------------------------------
    # FINAL CLEANUP & EXPORT (Strict 1-to-1 Row Preservation)
    # ---------------------------------------------------------
    df_processed.drop(columns=['__has_boiler'], inplace=True, errors='ignore')
    df_processed.fillna("", inplace=True)
    
    df_processed = df_processed[[
        "Row", "Building", "Original Phrase", "Scrubbed Phrase", 
        "LLM Model Identity", "Technical Specs", "Consumables", "Lifecycle Data", "Asset Tag", 
        "Location", "System Group", "Hierarchy", "Audit Flag", "PM Suitability", 
        "Yearly Effort (Hrs)", "Yearly Material ($)", "NFPA Group"
    ]]
    df_processed.to_csv(output_file, sep='\t', index=False, encoding='utf-8-sig')

    return jsonify({"status": "success", "processed": len(df_processed)}), 200

if __name__ == '__main__':
    app.run(port=5000, threaded=True)