from flask import Flask, request, jsonify
from sentence_transformers import SentenceTransformer
import torch
import numpy as np
import pandas as pd
import os
import glob
import re
import json
import threading

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
compiled_multi_rules = []
compiled_single_rules = []
compiled_broad_categories = [] # RESTORED
uniformat_valid_words = set()
master_rule_pattern = None

# Define hierarchy categories based on standard nomenclature
HIERARCHY_RULES = [
    {
        "group": "Life Safety / Fire",
        "parents": [r'\bpanel\b', r'\bfacp\b', r'\bannunciator\b', r'\bfire alarm control\b', r'\bmain fire\b'],
        "children": [r'\bsmoke\b', r'\bdetector\b', r'\bstrobe\b', r'\bhorn\b', r'\bpull station\b', r'\bbell\b', r'\btamper switch\b', r'\bflow switch\b', r'\bfire damper\b']
    },
    {
        "group": "Emergency Power",
        "parents": [r'\bgenerator\b', r'\bgenset\b', r'\bdiesel gen\b', r'\bgas gen\b'],
        "children": [r'\bats\b', r'\btransfer switch\b', r'\bday tank\b', r'\bload bank\b', r'\bfuel polishing\b']
    },
    {
        "group": "Vertical Transport",
        "parents": [r'\belevator\b', r'\bcab\b', r'\bescalator\b', r'\bmoving walk\b', r'\bwheelchair lift\b', r'\bdumbwaiter\b'],
        "children": [r'\belevator controller\b', r'\bhoistway\b', r'\bpit equipment\b', r'\belevator machine room\b']
    },
    {
        "group": "Split HVAC",
        "parents": [r'\bcondensing unit\b', r'\bcondenser\b', r'\bcu\b', r'\boutdoor unit\b', r'\bvrf outdoor\b', r'\bvrv outdoor\b', r'\bheat pump outdoor\b'],
        "children": [r'\bfan coil\b', r'\bfcu\b', r'\bevaporator\b', r'\bvrf indoor\b', r'\bvrv indoor\b', r'\bcassette\b', r'\bwall mount split\b', r'\bindoor unit\b']
    },
    {
        "group": "Fire Pump System",
        "parents": [r'\bfire pump\b', r'\belectric fire pump\b', r'\bdiesel fire pump\b'],
        "children": [r'\bjockey pump\b', r'\bpressure maintenance pump\b', r'\bfire pump controller\b']
    }
]

def get_stemmed_words(text):
    words = text.split()
    stemmed = set()
    for w in words:
        if len(w) > 4:
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
    global compiled_multi_rules, compiled_single_rules, compiled_broad_categories # RESTORED
    global custom_rules, master_rule_pattern

    print("--- Initializing Engine ---")

    appdata = os.environ.get('APPDATA', '')
    if not appdata: appdata = os.path.expanduser('~')
        
    rules_file = os.path.join(appdata, "Uniformat_Rules.txt")

    lookup_phrases, lookup_phrases_original, canonical_casing = [], [], {}
    industry_translations, broad_categories = {}, []
    uniformat_valid_words.clear()

    # DYNAMIC SEARCH
    template_folder = os.environ.get("TEMPLATE_FOLDER", r"C:\Users\pcarr\OneDrive\Documents\Custom Office Templates")
    search_pattern_xltm = os.path.join(template_folder, "*.xltm")
    search_pattern_xlsm = os.path.join(template_folder, "*.xlsm")
    search_pattern_xlsx = os.path.join(template_folder, "*.xlsx")
    
    list_of_files = glob.glob(search_pattern_xltm) + glob.glob(search_pattern_xlsm) + glob.glob(search_pattern_xlsx)
    excel_file = max(list_of_files, key=os.path.getmtime) if list_of_files else ""
    
    if excel_file:
        print(f"Dynamically locked onto newest template: {os.path.basename(excel_file)}")

    # 1. LOAD LOOKUPS & BUILD WHITELIST
    if excel_file and os.path.exists(excel_file):
        try:
            print(f"Reading Holy Grail phrases directly from {os.path.basename(excel_file)}...")
            df_lookup = pd.read_excel(excel_file, sheet_name="Uniformat RS Means Lookup")
            if 'Asset Sub Type' in df_lookup.columns:
                raw_phrases = df_lookup['Asset Sub Type'].dropna().astype(str).tolist()
                paired = []
                for phrase in raw_phrases:
                    clean_phrase = str(phrase).strip()
                    if clean_phrase:
                        lower_phrase = clean_phrase.lower()
                        paired.append((lower_phrase, clean_phrase))
                        canonical_casing[lower_phrase] = clean_phrase
                        for word in re.findall(r'[a-z0-9]+', lower_phrase):
                            uniformat_valid_words.add(word)

                paired.sort(key=lambda x: len(x[0]), reverse=True)
                lookup_phrases = [p[0] for p in paired]
                lookup_phrases_original = [p[1] for p in paired]
                print(f"Successfully locked onto {len(lookup_phrases)} phrases from Excel!")
        except Exception as e:
            print(f"ERROR reading lookups: {e}")

    # FORCE HIERARCHY WORDS INTO WHITELIST
    for rule in HIERARCHY_RULES:
        for p in rule['parents'] + rule['children']:
            clean_p = p.replace(r'\b', '')
            for w in re.findall(r'[a-z0-9]+', clean_p.lower()):
                uniformat_valid_words.add(w)

    # 2. LOAD RULES
    if excel_file and os.path.exists(excel_file):
        try:
            print(f"Reading Rules Sheet directly from {os.path.basename(excel_file)}...")
            df_rules = pd.read_excel(excel_file, sheet_name="Rules Sheet")
            for index, row in df_rules.iterrows():
                if pd.notna(row.iloc[0]) and pd.notna(row.iloc[1]) and pd.notna(row.iloc[2]):
                    rule_type = str(row.iloc[0]).strip().lower()
                    from_text = str(row.iloc[1]).strip().lower()
                    to_text = str(row.iloc[2]).strip().replace('|', '\n')
                    
                    if rule_type in ["alias", "phrase", "list"]:
                        industry_translations[from_text] = to_text
                        for word in re.findall(r'[a-z0-9]+', to_text.lower()):
                            uniformat_valid_words.add(word)
                        if '\n' in to_text:
                            for word in re.findall(r'[a-z0-9]+', from_text.lower()):
                                uniformat_valid_words.add(word)
                    elif rule_type == "core":
                        broad_categories.append(from_text)
                        for word in re.findall(r'[a-z0-9]+', from_text):
                            uniformat_valid_words.add(word)
            print(f"Loaded {len(industry_translations)} translations and {len(broad_categories)} core categories from Excel!")
        except Exception as e:
            print(f"ERROR reading Rules: {e}")
    
    # Compile Master Regex
    sorted_messy_terms = sorted(industry_translations.keys(), key=len, reverse=True)
    if sorted_messy_terms:
        escaped_keys = [r'\b' + re.escape(k) + r'\b' for k in sorted_messy_terms]
        master_rule_pattern = re.compile('(' + '|'.join(escaped_keys) + ')', flags=re.IGNORECASE)
    else:
        master_rule_pattern = None

    compiled_multi_rules.clear()
    compiled_single_rules.clear()
    for messy_term in sorted_messy_terms:
        clean_term = industry_translations[messy_term]
        if clean_term:
            rule_tokens = get_stemmed_words(messy_term)
            if '\n' in clean_term:
                compiled_multi_rules.append((rule_tokens, clean_term))
            else:
                compiled_single_rules.append((rule_tokens, clean_term))

    compiled_broad_categories = [re.compile(r'\b' + re.escape(cat) + r'\b') for cat in broad_categories]
    lookup_words_sets = [get_stemmed_words(phrase) for phrase in lookup_phrases]

    if lookup_phrases:
        embeddings = embedder.encode(lookup_phrases, batch_size=256)
        norms = np.linalg.norm(embeddings, axis=1, keepdims=True)
        lookup_embeddings = embeddings / (norms + 1e-9)

    # 3. LOAD LEARNED MAPPINGS
    if os.path.exists(custom_rules_file):
        try:
            print(f"Reading Learned Mappings from {custom_rules_file}...")
            with open(custom_rules_file, "r", encoding="utf-8") as f:
                with custom_rules_lock:
                    custom_rules.clear()
                    for line in f:
                        line = line.strip()
                        if line:
                            parts = line.split("|")
                            if len(parts) >= 2:
                                custom_rules[parts[0].strip()] = parts[1].strip()
            print(f"Loaded {len(custom_rules)} custom learned rules from memory!")
        except Exception as e:
            pass
    elif excel_file and os.path.exists(excel_file):
        try:
            print(f"Reading LearnedMappings directly from {os.path.basename(excel_file)}...")
            df_learned = pd.read_excel(excel_file, sheet_name="LearnedMappings")
            with custom_rules_lock:
                custom_rules.clear()
                for index, row in df_learned.iterrows():
                    if pd.notna(row.iloc[2]) and pd.notna(row.iloc[1]):
                        signature = str(row.iloc[2]).strip()
                        match = str(row.iloc[1]).strip()
                        if signature and match:
                            custom_rules[signature] = match
            print(f"Loaded {len(custom_rules)} custom learned rules from Excel memory!")
        except Exception as e:
            print(f"ERROR reading LearnedMappings from Excel: {e}")

    print(f"Engine Ready! Whitelist contains {len(uniformat_valid_words)} authorized words.")

load_knowledge_base()

# ==========================================
# 2. THE UNIFIED PRE-SCRUB ENGINE
# ==========================================
def prescrub_text(raw_text):
    if not isinstance(raw_text, str): return ""
        
    clean = re.sub(r'[\x00-\x09\x0B-\x0C\x0E-\x1F\u200B-\u200D\uFEFF]', '', raw_text)
    clean = clean.replace('\xa0', ' ').strip().lower()
    clean = re.sub(r'[^a-z0-9\s]', ' ', clean)
    
    if master_rule_pattern:
        def replace_logic(m):
            match_word = m.group(0).lower()
            target_text = industry_translations.get(match_word, "")
            if '\n' in target_text:
                return match_word
            return target_text.split('\n')[0].lower()
            
        clean = master_rule_pattern.sub(replace_logic, clean)
            
    if uniformat_valid_words:
        words = clean.split()
        kept_words = [w for w in words if w in uniformat_valid_words]
        clean = " ".join(kept_words)
        
    clean = re.sub(r'\s+', ' ', clean).strip()
    clean = re.sub(r'\s+\b\d+\b$', '', clean).strip()
    
    return clean

def categorize_asset(scrubbed_phrase):
    for rule in HIERARCHY_RULES:
        for p in rule['parents']:
            if re.search(p, scrubbed_phrase): return rule['group'], "Parent"
        for c in rule['children']:
            if re.search(c, scrubbed_phrase): return rule['group'], "Child"
    return "N/A", "N/A"

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

@app.route('/learn', methods=['POST'])
def learn_rule():
    data = request.json
    raw_phrase = data.get('phrase', '').strip()
    clean_match = data.get('match', '').strip()

    corrected_match = canonical_casing.get(clean_match.lower(), clean_match)
    signature = prescrub_text(raw_phrase)

    if signature and corrected_match:
        with custom_rules_lock:
            custom_rules[signature] = corrected_match
            diskMap = {}
            if os.path.exists(custom_rules_file):
                with open(custom_rules_file, "r", encoding="utf-8") as f:
                    for line in f:
                        line = line.strip()
                        if line:
                            parts = line.split("|")
                            if len(parts) >= 2: diskMap[parts[0]] = line
            
            from datetime import datetime
            now_str = datetime.now().strftime("%m/%d/%Y %I:%M:%S %p")
            diskMap[signature] = f"{signature}|{corrected_match}|{signature}|{now_str}"
            
            temp_file = custom_rules_file + ".tmp"
            with open(temp_file, "w", encoding="utf-8") as f:
                for k, v in diskMap.items(): f.write(f"{v}\n")
            os.replace(temp_file, custom_rules_file)

        print(f"🧠 AI LEARNED PURIFIED RULE: [{signature}] -> [{corrected_match}]")
        return jsonify({"status": "success"}), 200

    return jsonify({"status": "ignored"}), 400

@app.route('/batch_lookup', methods=['POST'])
def batch_lookup():
    data = request.json
    items = data.get('items', [])
    results = []

    ai_queue_phrases = []
    ai_queue_rows = []
    original_phrases_list = []

    for item in items:
        original_phrase = item.get('phrase', '').strip()
        row_id = item.get('row')
        current_match = item.get('current_match', '').strip().lower()
        current_id = item.get('current_id', '').strip().lower()

        signature = prescrub_text(original_phrase)
        
        original_clean = re.sub(r'[^a-z0-9\s]', ' ', original_phrase.lower())
        original_base_words = get_stemmed_words(original_clean)

        is_broken = False
        if current_match in ["", "no good match", "no matching words"]: is_broken = True
        if any(x in current_id for x in ["none", "category", "requires human", "requires ai"]): is_broken = True

        if not is_broken:
            results.append({"row": row_id, "match": "SKIP", "id": "SKIP"})
            continue

        with custom_rules_lock:
            match_found = custom_rules.get(signature)
            if not match_found:
                sorted_rules = sorted(custom_rules.items(), key=lambda x: len(x[0]), reverse=True)
                for rule_key, rule_val in sorted_rules:
                    if re.search(r'\b' + re.escape(rule_key) + r'\b', signature):
                        match_found = rule_val
                        break
            if not match_found:
                sig_words = get_stemmed_words(signature)
                best_mem_score = 0.0
                for rule_key, rule_val in custom_rules.items():
                    rule_words = get_stemmed_words(rule_key)
                    if not rule_words: continue
                    intersection = len(sig_words.intersection(rule_words))
                    overlap = intersection / len(rule_words) if len(rule_words) > 0 else 0.0
                    if overlap > best_mem_score and overlap >= 0.75:
                        best_mem_score = overlap
                        match_found = rule_val

        if match_found:
            validated = validate_and_format(match_found)
            if validated != "No good match":
                results.append({"row": row_id, "match": validated, "id": "USER_LEARNED"})
                continue

        multi_match_found = False
        for rule_tokens, clean_term in compiled_multi_rules:
            if rule_tokens.issubset(original_base_words):
                validated = validate_and_format(clean_term.strip())
                if validated != "No good match":
                    results.append({"row": row_id, "match": validated, "id": "EXPERT_RULE"})
                    multi_match_found = True
                    break
        if multi_match_found:
            continue

        if len(signature) >= 2:
            ai_queue_phrases.append(signature)
            ai_queue_rows.append(row_id)
            original_phrases_list.append(original_phrase)
        else:
            results.append({"row": row_id, "match": "No good match", "id": "REQUIRES HUMAN"})

    if ai_queue_phrases:
        unique_candidates = set()
        item_candidates_list = []

        for signature in ai_queue_phrases:
            candidates_set = {signature}
            words = signature.split()
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

        top_k = min(20, len(lookup_phrases))

        for i_queue, signature in enumerate(ai_queue_phrases):
            row_id = ai_queue_rows[i_queue]
            original_phrase_for_rule = original_phrases_list[i_queue]
            
            original_clean_for_rule = re.sub(r'[^a-z0-9\s]', ' ', original_phrase_for_rule.lower())
            original_base_words_for_rule = get_stemmed_words(original_clean_for_rule)

            candidates = item_candidates_list[i_queue]
            num_candidates = len(candidates)
            indices = [cand_to_idx[c] for c in candidates]
            candidate_norms = unique_norms[indices]

            base_words = get_stemmed_words(signature)
            all_semantic_scores = np.dot(candidate_norms, lookup_embeddings.T)

            collected_matches = []
            overlap_cache = {}

            for c_idx in range(num_candidates):
                semantic_scores = all_semantic_scores[c_idx]
                top_indices = np.argsort(semantic_scores)[-top_k:][::-1]

                for idx in top_indices:
                    sem_score = semantic_scores[idx]
                    lookup_candidate_original = lookup_phrases_original[idx]

                    if idx not in overlap_cache:
                        lookup_words = lookup_words_sets[idx]
                        intersection = len(base_words.intersection(lookup_words))
                        overlap = (2.0 * intersection) / (len(base_words) + len(lookup_words)) if (len(base_words) + len(lookup_words)) > 0 else 0.0
                        overlap_cache[idx] = (overlap * 0.35) + (0.15 if overlap >= 0.80 else 0.0)

                    combined_score = (sem_score * 0.50) + overlap_cache[idx]
                    collected_matches.append((combined_score, lookup_candidate_original))

            collected_matches.sort(key=lambda x: x[0], reverse=True)
            unique_matches = []
            seen_phrases = set()
            for score, phrase in collected_matches:
                if phrase not in seen_phrases:
                    unique_matches.append((score, phrase))
                    seen_phrases.add(phrase)

            best_combined_score = unique_matches[0][0] if unique_matches else 0.0

            if best_combined_score >= 0.75:
                proposed_match = unique_matches[0][1]
                validated = validate_and_format(proposed_match)
                final_match = validated if validated != "No good match" else "No good match"
                final_id = "AI_SMART_VECTOR" if validated != "No good match" else "REQUIRES HUMAN"
            else:
                domain_match_found = False
                for rule_tokens, clean_term in compiled_single_rules:
                    if rule_tokens.issubset(original_base_words_for_rule):
                        validated = validate_and_format(clean_term.strip())
                        if validated != "No good match":
                            final_match, final_id = validated, "EXPERT_RULE"
                            domain_match_found = True
                            break

                if not domain_match_found:
                    if best_combined_score >= 0.40:
                        close_matches = [m[1] for m in unique_matches if m[0] >= 0.40 and (best_combined_score - m[0] <= 0.15)][:5]
                        if len(close_matches) > 1:
                            final_match, final_id = validate_and_format("\n".join(close_matches)), "AI_SUGGESTED_LIST"
                        else:
                            final_match, final_id = validate_and_format(unique_matches[0][1]), "AI_HYBRID_MATCH"
                    else:
                        broad_match_found = False
                        for pattern in compiled_broad_categories:
                            if pattern.search(signature):
                                final_match, final_id = "Subtype missing in input", "REQUIRES HUMAN"
                                broad_match_found = True
                                break
                        if not broad_match_found:
                            final_match, final_id = "No good match", "REQUIRES HUMAN"

            results.append({"row": row_id, "match": final_match, "id": final_id})

    return jsonify(results)


@app.route('/batch_file', methods=['POST'])
def batch_file():
    data = request.json
    input_file = data.get('input_file')
    output_file = data.get('output_file')

    if not os.path.exists(input_file):
        return jsonify({"status": "error", "message": "Input file not found"}), 400

    df = pd.read_csv(input_file, sep='\t', dtype=str)

    results = []
    ai_queue_phrases = []
    ai_queue_rows = []
    original_phrases_list = []

    for index, row in df.iterrows():
        row_id = str(row['Row'])
        original_phrase = str(row.get('Phrase', '')).strip()

        signature = prescrub_text(original_phrase)
        original_clean = re.sub(r'[^a-z0-9\s]', ' ', original_phrase.lower())
        original_base_words = get_stemmed_words(original_clean)

        with custom_rules_lock:
            match_found = custom_rules.get(signature)
            if not match_found:
                sorted_rules = sorted(custom_rules.items(), key=lambda x: len(x[0]), reverse=True)
                for rule_key, rule_val in sorted_rules:
                    if re.search(r'\b' + re.escape(rule_key) + r'\b', signature):
                        match_found = rule_val
                        break
            if not match_found:
                sig_words = get_stemmed_words(signature)
                best_mem_score = 0.0
                for rule_key, rule_val in custom_rules.items():
                    rule_words = get_stemmed_words(rule_key)
                    if not rule_words: continue
                    intersection = len(sig_words.intersection(rule_words))
                    overlap = intersection / len(rule_words) if len(rule_words) > 0 else 0.0
                    if overlap > best_mem_score and overlap >= 0.75:
                        best_mem_score = overlap
                        match_found = rule_val

        if match_found:
            validated = validate_and_format(match_found)
            if validated != "No good match":
                results.append({"Row": row_id, "Match": validated, "Id": "USER_LEARNED"})
                continue

        multi_match_found = False
        for rule_tokens, clean_term in compiled_multi_rules:
            if rule_tokens.issubset(original_base_words):
                validated = validate_and_format(clean_term.strip())
                if validated != "No good match":
                    results.append({"Row": row_id, "Match": validated, "Id": "EXPERT_RULE"})
                    multi_match_found = True
                    break
        if multi_match_found:
            continue

        if len(signature) >= 2:
            ai_queue_phrases.append(signature)
            ai_queue_rows.append(row_id)
            original_phrases_list.append(original_phrase)
        else:
            results.append({"Row": row_id, "Match": "No good match", "Id": "REQUIRES HUMAN"})

    if ai_queue_phrases:
        unique_candidates = set()
        item_candidates_list = []

        for signature in ai_queue_phrases:
            candidates_set = {signature}
            words = signature.split()
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

        top_k = min(20, len(lookup_phrases))

        for i_queue, signature in enumerate(ai_queue_phrases):
            row_id = ai_queue_rows[i_queue]
            original_phrase = original_phrases_list[i_queue]
            
            original_clean = re.sub(r'[^a-z0-9\s]', ' ', original_phrase.lower())
            original_base_words = get_stemmed_words(original_clean)

            candidates = item_candidates_list[i_queue]
            num_candidates = len(candidates)
            indices = [cand_to_idx[c] for c in candidates]
            candidate_norms = unique_norms[indices]

            base_words = get_stemmed_words(signature)
            all_semantic_scores = np.dot(candidate_norms, lookup_embeddings.T)

            collected_matches = []
            overlap_cache = {}

            for c_idx in range(num_candidates):
                semantic_scores = all_semantic_scores[c_idx]
                top_indices = np.argsort(semantic_scores)[-top_k:][::-1]

                for idx in top_indices:
                    sem_score = semantic_scores[idx]
                    lookup_candidate_original = lookup_phrases_original[idx]

                    if idx not in overlap_cache:
                        lookup_words = lookup_words_sets[idx]
                        intersection = len(base_words.intersection(lookup_words))
                        overlap = (2.0 * intersection) / (len(base_words) + len(lookup_words)) if (len(base_words) + len(lookup_words)) > 0 else 0.0
                        overlap_cache[idx] = (overlap * 0.35) + (0.15 if overlap >= 0.80 else 0.0)

                    combined_score = (sem_score * 0.50) + overlap_cache[idx]
                    collected_matches.append((combined_score, lookup_candidate_original))

            collected_matches.sort(key=lambda x: x[0], reverse=True)
            unique_matches = []
            seen_phrases = set()
            for score, phrase in collected_matches:
                if phrase not in seen_phrases:
                    unique_matches.append((score, phrase))
                    seen_phrases.add(phrase)

            best_combined_score = unique_matches[0][0] if unique_matches else 0.0

            if best_combined_score >= 0.75:
                proposed_match = unique_matches[0][1]
                validated = validate_and_format(proposed_match)
                final_match = validated if validated != "No good match" else "No good match"
                final_id = "AI_SMART_VECTOR" if validated != "No good match" else "REQUIRES HUMAN"
            else:
                domain_match_found = False
                for rule_tokens, clean_term in compiled_single_rules:
                    if rule_tokens.issubset(original_base_words):
                        validated = validate_and_format(clean_term.strip())
                        if validated != "No good match":
                            final_match, final_id = validated, "EXPERT_RULE"
                            domain_match_found = True
                            break

                if not domain_match_found:
                    if best_combined_score >= 0.40:
                        close_matches = [m[1] for m in unique_matches if m[0] >= 0.40 and (best_combined_score - m[0] <= 0.15)][:5]
                        if len(close_matches) > 1:
                            final_match, final_id = validate_and_format("\n".join(close_matches)), "AI_SUGGESTED_LIST"
                        else:
                            final_match, final_id = validate_and_format(unique_matches[0][1]), "AI_HYBRID_MATCH"
                    else:
                        broad_match_found = False
                        for pattern in compiled_broad_categories:
                            if pattern.search(signature):
                                final_match, final_id = "Subtype missing in input", "REQUIRES HUMAN"
                                broad_match_found = True
                                break
                        if not broad_match_found:
                            final_match, final_id = "No good match", "REQUIRES HUMAN"

            results.append({"Row": row_id, "Match": final_match, "Id": final_id})

    for r in results:
        if isinstance(r.get("Match"), str):
            r["Match"] = r["Match"].replace('\n', '\\n')

    df_out = pd.DataFrame(results)
    df_out.to_csv(output_file, sep='\t', index=False)

    return jsonify({"status": "success", "processed": len(results)}), 200


@app.route('/prescrub_file', methods=['POST'])
def prescrub_file():
    data = request.json
    input_file = data.get('input_file')
    output_file = data.get('output_file')

    if not os.path.exists(input_file):
        return jsonify({"status": "error", "message": "Input file not found"}), 400

    df = pd.read_csv(input_file, sep='\t', dtype=str)
    
    processed_data = []
    for index, row in df.iterrows():
        row_id = str(row.get('Row', ''))
        building = str(row.get('Building', '')).strip()
        original_phrase = str(row.get('Phrase', '')).strip()
        
        scrubbed = prescrub_text(original_phrase)
        sys_group, hierarchy = categorize_asset(scrubbed)
        
        processed_data.append({
            "Row": row_id,
            "Building": building,
            "Original Phrase": original_phrase,
            "Scrubbed Phrase": scrubbed,
            "System Group": sys_group,
            "Hierarchy": hierarchy,
            "Audit Flag": ""
        })
        
    df_processed = pd.DataFrame(processed_data)
    
    for bldg in df_processed['Building'].unique():
        if not bldg: continue
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
                if p_count > 0 and c_count == 0: audit_msg = "⚠️ Missing Indoor Child"
                elif c_count > 0 and p_count == 0: audit_msg = "⚠️ Missing Outdoor Parent"
                elif p_count > 0 and c_count > 0: audit_msg = "Matched Split System"
            elif group_name in ["Life Safety / Fire", "Emergency Power", "Fire Pump System", "Vertical Transport"]:
                flag_text = "Main Fire Panel" if group_name == "Life Safety / Fire" else "Generator Parent" if group_name == "Emergency Power" else "Main Fire Pump" if group_name == "Fire Pump System" else "Elevator/Escalator Parent"
                if c_count > 0 and p_count == 0: audit_msg = f"⚠️ Missing {flag_text}"
                elif p_count > 0: audit_msg = "Valid Parent System"
                
            if audit_msg: df_processed.loc[group_mask, 'Audit Flag'] = audit_msg

    df_processed.fillna("", inplace=True)
    df_processed.to_csv(output_file, sep='\t', index=False)

    return jsonify({"status": "success", "processed": len(df_processed)}), 200

if __name__ == '__main__':
    app.run(port=5000, threaded=True)