from flask import Flask, request, jsonify
from sentence_transformers import SentenceTransformer
import torch
import numpy as np
import pandas as pd
import os
import re
import json
import threading
import requests
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

custom_rules_file = os.path.join(appdata, "Uniformat_Learned.txt")
custom_rules = {}
custom_rules_lock = threading.Lock()

lookup_phrases = []
lookup_phrases_original = []
canonical_casing = {}
lookup_words_sets = []
lookup_embeddings = None

industry_translations = {}
broad_categories = []

compiled_expert_rules = []
compiled_broad_categories = []

uniformat_valid_words = set()
stemmed_whitelist = set()
master_rule_pattern = None

# PRE-COMPILED REGEXES FOR PRE-SCRUB AND IDENTIFIERS
RE_CONTROL_CHARS = re.compile(r'[\x00-\x09\x0B-\x0C\x0E-\x1F\u200B-\u200D\uFEFF]')
RE_NON_ALPHANUM = re.compile(r'[^a-z0-9\s]')
RE_MULTI_SPACE = re.compile(r'\s+')
RE_TRAILING_DIGIT = re.compile(r'\s+\b\d+\b$')

RE_HASHTAGS = re.compile(r'#\s*[A-Z0-9]+', re.IGNORECASE)
RE_ROOMS = re.compile(r'\b(?:room|rm)\s*[A-Z0-9]+', re.IGNORECASE)
RE_TRAILING_NUM = re.compile(r'(\d+[A-Z]?)$')

# PRE-COMPILED AT BOOT TO SAVE CPU CYCLES
COMPILED_LOCATION_PATTERNS = [
    re.compile(r'\b(?:elevator )?lobby\b', re.IGNORECASE),
    re.compile(r'\b(?:stair|stairwell|staircase)\s*[a-z0-9]*\b', re.IGNORECASE),
    re.compile(r'\b(?:vestibule|corridor|hallway|hall|concourse|mezzanine|mezz|penthouse|ph|basement|roof|ceiling|atrium|closet)\b', re.IGNORECASE),
    re.compile(r'\b(?:parking )?garage\b', re.IGNORECASE),
    re.compile(r'\b(?:level|floor|fl)\s*\d+[a-z]*\b', re.IGNORECASE),
    re.compile(r'\b\d+(?:st|nd|rd|th)?\s+(?:floor|fl|level)\b', re.IGNORECASE),
    re.compile(r'\bp\d+\b', re.IGNORECASE),
    re.compile(r'\b(?:mechanical|elec|electrical|boiler|pump|utility|telecom|it|server|data|storage|mail)\s+room\b', re.IGNORECASE)
]

# Define hierarchy categories based on standard nomenclature
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
        "parents": [],
        "children": [r'\bfire pump', r'\belectric fire pump', r'\bdiesel fire pump', r'\bjockey pump', r'\bpressure maintenance pump', r'\bfire pump controller']
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

def validate_and_format(match_str):
    if not match_str or match_str in ["No good match", "REQUIRES HUMAN", "SKIP"]:
        return match_str

    if "\n" in match_str:
        valid_parts = []
        for part in match_str.split("\n"):
            clean = canonical_casing.get(part.strip().lower(), part.strip())
            if clean: valid_parts.append(clean)
        if valid_parts: return "\n".join(valid_parts)
        return "No good match"

    clean_match = canonical_casing.get(match_str.strip().lower(), match_str.strip())
    return clean_match if clean_match else "No good match"

def load_knowledge_base():
    global lookup_phrases, lookup_phrases_original, canonical_casing
    global lookup_words_sets, lookup_embeddings
    global industry_translations, broad_categories, uniformat_valid_words
    global stemmed_whitelist
    global compiled_expert_rules, compiled_broad_categories
    global custom_rules, master_rule_pattern

    print("--- Initializing Engine ---")

    appdata = os.environ.get('APPDATA', '')
    if not appdata: appdata = os.path.expanduser('~')

    lookups_file = os.path.join(appdata, "Uniformat_Lookups.txt")
    excel_file = "Book1.xlsx"

    lookup_phrases, lookup_phrases_original, canonical_casing = [], [], {}
    industry_translations, broad_categories = {}, []
    uniformat_valid_words.clear()
    stemmed_whitelist.clear()

    # 1. LOAD LOOKUPS FROM APPDATA
    if os.path.exists(lookups_file):
        try:
            print(f"Reading Lookups from {lookups_file}...")
            paired = []
            with open(lookups_file, "r", encoding="utf-8-sig") as f:
                for line in f:
                    clean_phrase = line.strip()
                    if clean_phrase:
                        lower_phrase = clean_phrase.lower()
                        stripped_phrase = re.sub(r'[^a-z0-9\s]', ' ', lower_phrase)
                        stripped_phrase = re.sub(r'\s+', ' ', stripped_phrase).strip()

                        paired.append((stripped_phrase, clean_phrase))
                        canonical_casing[lower_phrase] = clean_phrase

                        for word in stripped_phrase.split():
                            uniformat_valid_words.add(word)

            paired.sort(key=lambda x: len(x[0]), reverse=True)
            lookup_phrases = [p[0] for p in paired]
            lookup_phrases_original = [p[1] for p in paired]
            print(f"Loaded {len(lookup_phrases)} lookup phrases!")
        except Exception as e:
            print(f"ERROR reading lookups: {e}")
    else:
        print(f"WARNING: Lookups file not found.")

    # FORCE HIERARCHY WORDS INTO WHITELIST
    for rule in HIERARCHY_RULES:
        for p in rule['parents'] + rule['children']:
            clean_p = p.replace(r'\b', '')
            for w in re.findall(r'[a-z0-9]+', clean_p.lower()):
                uniformat_valid_words.add(w)

    # 2. LOAD RULES FROM EXCEL WORKBOOK DIRECTLY
    if os.path.exists(excel_file):
        try:
            print(f"Reading Rules directly from active workbook: {excel_file}...")
            df_rules = pd.read_excel(excel_file, sheet_name="Rules Sheet")
            for _, row in df_rules.iterrows():
                rule_type = str(row.get("Type", "")).strip().lower()
                from_text = str(row.get("From", "")).strip().lower()
                to_text = str(row.get("To", "")).strip().replace('|', '\n')
                if to_text == 'nan': to_text = ""

                if rule_type in ["alias", "phrase", "list"]:
                    industry_translations[from_text] = to_text
                    for word in re.findall(r'[a-z0-9]+', to_text.lower()):
                        uniformat_valid_words.add(word)
                    for word in re.findall(r'[a-z0-9]+', from_text.lower()):
                        uniformat_valid_words.add(word)
                elif rule_type == "core":
                    broad_categories.append(from_text)
                    for word in re.findall(r'[a-z0-9]+', from_text):
                        uniformat_valid_words.add(word)

            print(f"Loaded {len(industry_translations)} translations and {len(broad_categories)} core categories from Excel!")
        except Exception as e:
            print(f"ERROR reading Rules from Excel: {e}")

        # 3. LOAD CUSTOM LEARNED MAPPINGS FROM EXCEL
        try:
            if "Learned Mappings" in pd.ExcelFile(excel_file).sheet_names:
                df_learned = pd.read_excel(excel_file, sheet_name="Learned Mappings")
                with custom_rules_lock:
                    custom_rules.clear()
                    for _, row in df_learned.iterrows():
                        rule_k = str(row.iloc[0]).strip()
                        rule_v = str(row.iloc[1]).strip()
                        if rule_k and rule_v and rule_k != 'nan':
                            custom_rules[rule_k] = rule_v
                print(f"Loaded {len(custom_rules)} custom learned rules from Excel!")
        except Exception as e:
            print(f"ERROR reading Learned Mappings from Excel: {e}")
    else:
        print(f"WARNING: {excel_file} not found.")

    if not broad_categories:
        broad_categories = ["pump", "fan", "boiler", "furnace", "transformer", "compressor", "chiller", "motor", "sump", "sprinkler", "valve", "hydrant", "tower", "exchanger", "tank"]
        for cat in broad_categories:
            for word in re.findall(r'[a-z0-9]+', cat):
                uniformat_valid_words.add(word)

    sorted_messy_terms = sorted(industry_translations.keys(), key=len, reverse=True)
    if sorted_messy_terms:
        escaped_keys = [r'\b' + re.escape(k) + r'\b' for k in sorted_messy_terms]
        master_rule_pattern = re.compile('(' + '|'.join(escaped_keys) + ')', flags=re.IGNORECASE)
    else:
        master_rule_pattern = None

    compiled_expert_rules.clear()
    for messy_term in sorted_messy_terms:
        clean_term = industry_translations[messy_term]
        if clean_term:
            clean_messy = re.sub(r'[^a-z0-9\s]', ' ', messy_term)
            rule_tokens = get_stemmed_words(clean_messy)
            compiled_expert_rules.append((messy_term, rule_tokens, clean_term))

    compiled_broad_categories = [re.compile(r'\b' + re.escape(cat) + r'\b') for cat in broad_categories]
    lookup_words_sets = [get_stemmed_words(phrase) for phrase in lookup_phrases]

    if lookup_phrases:
        embeddings = embedder.encode(lookup_phrases, batch_size=256)
        norms = np.linalg.norm(embeddings, axis=1, keepdims=True)
        lookup_embeddings = embeddings / (norms + 1e-9)

    # 4. BUILD SMART STEMMED WHITELIST FOR PLURALS
    for w in uniformat_valid_words:
        stemmed_whitelist.add(w)
        stemmed_whitelist.update(get_stemmed_words(w))

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
        for m in matches:
            locations.append(m)
            clean_base = re.sub(r'\b' + re.escape(m) + r'\b', '', clean_base, flags=re.IGNORECASE)

    clean_base = re.sub(r'\s+', ' ', clean_base).strip()
    extracted = " | ".join(sorted(set([loc.title() for loc in locations])))

    return extracted, clean_base

def extract_asset_identifiers(raw_text):
    if not isinstance(raw_text, str):
        return "", raw_text

    identifiers = []
    clean_base = raw_text.strip()

    hash_tags = RE_HASHTAGS.findall(clean_base)
    identifiers.extend(hash_tags)
    for tag in hash_tags:
        clean_base = clean_base.replace(tag, '')

    rooms = RE_ROOMS.findall(clean_base)
    identifiers.extend(rooms)
    for room in rooms:
        clean_base = clean_base.replace(room, '')

    clean_base = clean_base.strip()
    trailing_num = RE_TRAILING_NUM.search(clean_base)
    if trailing_num and not hash_tags:
        identifiers.append(trailing_num.group(1))
        clean_base = clean_base[:trailing_num.start()]

    return " ".join(identifiers).strip(), clean_base.strip()

def purify_text(scrubbed_text):
    if not isinstance(scrubbed_text, str) or not scrubbed_text:
        return ""
    if not uniformat_valid_words:
        return scrubbed_text

    pure_words = []
    for w in scrubbed_text.split():
        if w in uniformat_valid_words:
            pure_words.append(w)
        else:
            stem_set = get_stemmed_words(w)
            w_stem = list(stem_set)[0] if stem_set else ""
            if w_stem in stemmed_whitelist:
                pure_words.append(w)

    pure = " ".join(pure_words)
    return pure if pure else scrubbed_text

def prescrub_text(raw_text, use_whitelist=True):
    if not isinstance(raw_text, str): return ""

    clean = RE_CONTROL_CHARS.sub('', raw_text)
    clean = clean.replace('\xa0', ' ').strip().lower()

    if master_rule_pattern:
        def replace_logic(m):
            match_word = m.group(0).lower()
            target_text = industry_translations.get(match_word, "")
            if '\n' in target_text:
                return match_word
            return target_text.split('\n')[0].lower()

        clean = master_rule_pattern.sub(replace_logic, clean)

    clean = RE_NON_ALPHANUM.sub(' ', clean)
    clean = RE_MULTI_SPACE.sub(' ', clean).strip()
    clean = RE_TRAILING_DIGIT.sub('', clean).strip()

    if use_whitelist:
        clean = purify_text(clean)

    return clean

def categorize_asset(scrubbed_phrase):
    for rule in HIERARCHY_RULES:
        for p in rule['parents']:
            if re.search(p, scrubbed_phrase): return rule['group'], "Parent"
        for c in rule['children']:
            if re.search(c, scrubbed_phrase): return rule['group'], "Child"
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
            result_text = response.json().get("response", "")
            return json.loads(result_text)
    except Exception as e:
        print(f"Ollama API Error: {e}")

    return None

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
    signature = prescrub_text(raw_phrase)

    if not signature:
        return jsonify({"status": "ignored", "message": "Invalid signature"}), 400

    with custom_rules_lock:
        if signature in custom_rules:
            del custom_rules[signature]
            print(f"🗑️ AI FORGOT RULE: [{signature}]")
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
            signature = prescrub_text(raw_phrase)

            if signature and corrected_match:
                custom_rules[signature] = corrected_match
                learned_count += 1
                print(f"🧠 AI BATCH LEARNED IN MEMORY: [{signature}] -> [{corrected_match}]")

    return jsonify({"status": "success", "learned_count": learned_count}), 200

@app.route('/batch_lookup', methods=['POST'])
def batch_lookup():
    data = request.json
    items = data.get('items', [])
    results = []

    ai_queue_phrases = []
    ai_queue_rows = []
    original_phrases_list = []

    with custom_rules_lock:
        local_custom_rules = custom_rules.copy()
        
    compiled_custom_rules = []
    sorted_rules = sorted(local_custom_rules.items(), key=lambda x: len(x[0]), reverse=True)
    for k, v in sorted_rules:
        if k:
            compiled_custom_rules.append((
                re.compile(r'\b' + re.escape(k) + r'\b'),
                get_stemmed_words(k),
                v
            ))

    for item in items:
        original_phrase = item.get('phrase', '').strip()
        row_id = item.get('row')
        current_match = item.get('current_match', '').strip().lower()
        current_id = item.get('current_id', '').strip().lower()

        signature = prescrub_text(original_phrase, use_whitelist=True)
        semantic_phrase = prescrub_text(original_phrase, use_whitelist=False)

        original_clean = RE_NON_ALPHANUM.sub(' ', original_phrase.lower())
        original_base_words = get_stemmed_words(original_clean)

        is_broken = False
        if current_match in ["", "no good match", "no matching words"]: is_broken = True
        if any(x in current_id for x in ["none", "category", "requires human", "requires ai"]): is_broken = True

        if not is_broken:
            results.append({"row": row_id, "match": "SKIP", "id": "SKIP"})
            continue

        match_found = local_custom_rules.get(signature)
        if not match_found:
            for rule_re, _, rule_val in compiled_custom_rules:
                if rule_re.search(signature):
                    match_found = rule_val
                    break
        if not match_found:
            sig_words = get_stemmed_words(signature)
            best_mem_score = 0.0
            for _, rule_words, rule_val in compiled_custom_rules:
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
                results.append({"row": row_id, "match": validated, "id": "USER_LEARNED"})
                continue

        matched_rules = []
        for messy_term, rule_tokens, clean_term in compiled_expert_rules:
            if rule_tokens.issubset(original_base_words):
                matched_rules.append((messy_term, clean_term))

        expert_matches = []
        for i, (term1, clean1) in enumerate(matched_rules):
            is_subset = False
            for j, (term2, clean2) in enumerate(matched_rules):
                if i != j and f" {term1} " in f" {term2} ":
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
            final_id = "EXPERT_RULE"
            results.append({"row": row_id, "match": final_match, "id": final_id})
        elif len(semantic_phrase) >= 2:
            ai_queue_phrases.append(semantic_phrase)
            ai_queue_rows.append(row_id)
            original_phrases_list.append(original_phrase)
        else:
            results.append({"row": row_id, "match": "No good match", "id": "REQUIRES HUMAN"})

    return jsonify(results)

@app.route('/batch_file', methods=['POST'])
def batch_file():
    data = request.json
    input_file = data.get('input_file')
    output_file = data.get('output_file')

    if not os.path.exists(input_file):
        return jsonify({"status": "error", "message": "Input file not found"}), 400

    df = pd.read_csv(input_file, sep='\t', dtype=str, encoding='utf-8-sig')

    results = []
    ai_queue_phrases = []
    ai_queue_rows = []
    original_phrases_list = []

    with custom_rules_lock:
        local_custom_rules = custom_rules.copy()
        
    compiled_custom_rules = []
    sorted_rules = sorted(local_custom_rules.items(), key=lambda x: len(x[0]), reverse=True)
    for k, v in sorted_rules:
        if k:
            compiled_custom_rules.append((
                re.compile(r'\b' + re.escape(k) + r'\b'),
                get_stemmed_words(k),
                v
            ))

    for row_dict in df.to_dict('records'):
        row_id = str(row_dict.get('Row', ''))
        original_phrase = str(row_dict.get('Phrase', '')).strip()

        signature = prescrub_text(original_phrase, use_whitelist=True)
        semantic_phrase = prescrub_text(original_phrase, use_whitelist=False)

        original_clean = RE_NON_ALPHANUM.sub(' ', original_phrase.lower())
        original_base_words = get_stemmed_words(original_clean)

        match_found = local_custom_rules.get(signature)
        if not match_found:
            for rule_re, _, rule_val in compiled_custom_rules:
                if rule_re.search(signature):
                    match_found = rule_val
                    break
        if not match_found:
            sig_words = get_stemmed_words(signature)
            best_mem_score = 0.0
            for _, rule_words, rule_val in compiled_custom_rules:
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
                results.append({"Row": row_id, "Match": validated, "Id": "USER_LEARNED"})
                continue

        matched_rules = []
        for messy_term, rule_tokens, clean_term in compiled_expert_rules:
            if rule_tokens.issubset(original_base_words):
                matched_rules.append((messy_term, clean_term))

        expert_matches = []
        for i, (term1, clean1) in enumerate(matched_rules):
            is_subset = False
            for j, (term2, clean2) in enumerate(matched_rules):
                if i != j and f" {term1} " in f" {term2} ":
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
            final_id = "EXPERT_RULE"
            results.append({"Row": row_id, "Match": final_match, "Id": final_id})
        elif len(semantic_phrase) >= 2:
            ai_queue_phrases.append(semantic_phrase)
            ai_queue_rows.append(row_id)
            original_phrases_list.append(original_phrase)
        else:
            results.append({"Row": row_id, "Match": "No good match", "Id": "REQUIRES HUMAN"})

    if ai_queue_phrases:
        unique_candidates = set()
        item_candidates_list = []

        for semantic_phrase in ai_queue_phrases:
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

        all_pure_signatures = [purify_text(prescrub_text(orig, use_whitelist=True)) for orig in original_phrases_list]
        unique_pure = list(set(all_pure_signatures))
        pure_to_idx = {p: i for i, p in enumerate(unique_pure)}

        unique_pure_vecs = embedder.encode(unique_pure, batch_size=256)
        unique_pure_norms = unique_pure_vecs / (np.linalg.norm(unique_pure_vecs, axis=1, keepdims=True) + 1e-9)

        pure_norms_array = unique_pure_norms[[pure_to_idx[p] for p in all_pure_signatures]]

        top_k = min(20, len(lookup_phrases))

        def process_single_ai_item(i_queue):
            row_id = ai_queue_rows[i_queue]
            original_phrase = original_phrases_list[i_queue]

            candidates = item_candidates_list[i_queue]
            num_candidates = len(candidates)
            indices = [cand_to_idx[c] for c in candidates]
            candidate_norms = unique_norms[indices]

            strict_signature = prescrub_text(original_phrase, use_whitelist=True)
            pure_signature = all_pure_signatures[i_queue]
            base_words = get_stemmed_words(pure_signature)

            pure_norm = pure_norms_array[i_queue:i_queue+1]
            pure_semantic_scores = np.dot(pure_norm, lookup_embeddings.T)[0]

            all_semantic_scores = np.dot(candidate_norms, lookup_embeddings.T)

            collected_matches = []
            overlap_cache = {}

            input_core_regexes = []
            for cat in broad_categories:
                if re.search(r'\b' + re.escape(cat) + r'\b', strict_signature):
                    input_core_regexes.append(re.compile(r'\b' + re.escape(cat) + r'\b'))

            for c_idx in range(num_candidates):
                semantic_scores = all_semantic_scores[c_idx]

                combined_max_scores = np.maximum(semantic_scores, pure_semantic_scores)
                top_indices = np.argsort(combined_max_scores)[-top_k:][::-1]

                for idx in top_indices:
                    sem_score = combined_max_scores[idx]
                    lookup_candidate_original = lookup_phrases_original[idx]

                    if idx not in overlap_cache:
                        lookup_words = lookup_words_sets[idx]
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
                for pattern in compiled_broad_categories:
                    if pattern.search(signature):
                        final_match, final_id = "Subtype missing in input", "REQUIRES HUMAN"
                        broad_match_found = True
                        break
                if not broad_match_found:
                    final_match, final_id = "No good match", "REQUIRES HUMAN"

            return row_id, original_phrase, final_match, final_id

        ollama_cache = {}
        ollama_cache_lock = threading.Lock()

        def worker(i_queue):
            row_id, original_phrase, final_match, final_id = process_single_ai_item(i_queue)

            if final_id == "REQUIRES HUMAN" or final_match == "No good match":
                with ollama_cache_lock:
                    if original_phrase in ollama_cache:
                        ai_result = ollama_cache[original_phrase]
                    else:
                        print(f"Vector engine failed on [{original_phrase}]. Asking Local AI...")
                        ai_result = ask_maintenance_director_ollama(original_phrase, lookup_phrases_original)
                        ollama_cache[original_phrase] = ai_result

                if ai_result and ai_result.get("matched_asset"):
                    conf = str(ai_result.get("confidence", "")).strip().lower()
                    if conf == "low":
                        print("Ollama guessed loosely (Low Confidence). Rejecting guess.")
                        final_match = "No good match"
                        final_id = "REQUIRES HUMAN"
                    else:
                        final_match = ai_result["matched_asset"]
                        final_id = "AI_OLLAMA_DIRECTOR"
                        print(f"Ollama fixed it ({conf} confidence) -> {final_match}")

            return row_id, final_match, final_id

        with ThreadPoolExecutor(max_workers=8) as executor:
            futures = {executor.submit(worker, i): i for i in range(len(ai_queue_phrases))}
            for future in as_completed(futures):
                row_id, final_match, final_id = future.result()
                results.append({"Row": row_id, "Match": final_match, "Id": final_id})

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

    processed_data = []
    for row_dict in df.to_dict('records'):
        row_id = str(row_dict.get('Row', ''))
        building = str(row_dict.get('Building', '')).strip()
        original_phrase = str(row_dict.get('Phrase', '')).strip()

        extracted_loc, text_no_loc = extract_locations(original_phrase)
        asset_id, base_phrase = extract_asset_identifiers(text_no_loc)

        scrubbed = prescrub_text(base_phrase, use_whitelist=True)
        sys_group, hierarchy = categorize_asset(scrubbed)

        processed_data.append({
            "Row": row_id,
            "Building": building,
            "Original Phrase": original_phrase,
            "Scrubbed Phrase": scrubbed,
            "Asset Tag": asset_id,
            "Location": extracted_loc,
            "System Group": sys_group,
            "Hierarchy": hierarchy,
            "Audit Flag": ""
        })

    df_processed = pd.DataFrame(processed_data)

    for bldg in df_processed['Building'].unique():
        if not bldg or str(bldg).strip().lower() == 'nan': continue
        bldg_mask = df_processed['Building'] == bldg

        ambiguous_mask = bldg_mask & (df_processed['System Group'] == 'Ambiguous Fuel System')

        if ambiguous_mask.any():
            has_fp = (bldg_mask & (df_processed['System Group'] == 'Fire Pump System') & (df_processed['Hierarchy'] == 'Parent')).any()
            has_gen = (bldg_mask & (df_processed['System Group'] == 'Emergency Power') & (df_processed['Hierarchy'] == 'Parent')).any()
            has_boiler = (bldg_mask & df_processed['Scrubbed Phrase'].str.contains(r'\bboiler\b|\bblr\b', case=False, regex=True)).any()

            for idx in df_processed[ambiguous_mask].index:
                phrase = str(df_processed.at[idx, 'Scrubbed Phrase']).lower()
                is_day_tank = 'day tank' in phrase

                if is_day_tank:
                    if has_fp and not has_gen:
                        df_processed.at[idx, 'System Group'] = 'Fire Pump System'
                        df_processed.at[idx, 'Hierarchy'] = 'Child'
                    elif has_gen and not has_fp:
                        df_processed.at[idx, 'System Group'] = 'Emergency Power'
                        df_processed.at[idx, 'Hierarchy'] = 'Child'
                else:
                    if has_fp and not has_gen and not has_boiler:
                        df_processed.at[idx, 'System Group'] = 'Fire Pump System'
                        df_processed.at[idx, 'Hierarchy'] = 'Child'
                    elif has_gen and not has_fp and not has_boiler:
                        df_processed.at[idx, 'System Group'] = 'Emergency Power'
                        df_processed.at[idx, 'Hierarchy'] = 'Child'
                    elif has_boiler and not has_fp and not has_gen:
                        df_processed.at[idx, 'System Group'] = 'Boiler System'
                        df_processed.at[idx, 'Hierarchy'] = 'Child'

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
    df_processed = df_processed[["Row", "Building", "Original Phrase", "Scrubbed Phrase", "Asset Tag", "Location", "System Group", "Hierarchy", "Audit Flag"]]
    df_processed.to_csv(output_file, sep='\t', index=False, encoding='utf-8-sig')

    return jsonify({"status": "success", "processed": len(df_processed)}), 200

if __name__ == '__main__':
    app.run(port=5000, threaded=True)