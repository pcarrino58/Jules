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
# 1. INITIALIZE PURE MATH RADAR
# ==========================================
print("Loading Pure Semantic Math Engine...")

device = "cuda" if torch.cuda.is_available() else "cpu"
print(f"Engine forcing hardware acceleration on: [{device.upper()}]")

embedder = SentenceTransformer('all-MiniLM-L6-v2', device=device)

script_dir = os.path.dirname(os.path.abspath(__file__))
custom_rules_file = os.path.join(script_dir, "custom_rules.json")

custom_rules = {}
custom_rules_lock = threading.Lock()

# Global variables for lookup data
lookup_phrases = []
lookup_phrases_original = [] 
canonical_casing = {}        
lookup_words_sets = []
lookup_embeddings = None

industry_translations = {}
broad_categories = []
compiled_industry_translations = []
compiled_broad_categories = []

# Pre-compile noise removal regexes
RE_PARENS = re.compile(r'\(.*?\)')
RE_NON_ALPHA = re.compile(r'[^a-z0-9]')
RE_SPACES = re.compile(r'\s+')
RE_TRAILING_NOISE = re.compile(r'\s+\b([0-9]+[a-z]*|[a-z])\b$')

# --- THE NEW GRAMMAR ENGINE (STEMMER) ---
def get_stemmed_words(text):
    """Removes plural 's' so 'tanks' mathematically intersects with 'tank'"""
    words = text.split()
    stemmed = set()
    for w in words:
        # Don't strip 's' off words like 'glass', 'bypass', 'gas'
        if len(w) > 3 and w.endswith('s') and not w.endswith('ss'):
            stemmed.add(w[:-1])
        else:
            stemmed.add(w)
    return stemmed

def get_generalized_signature(text):
    text = re.sub(r'([a-z])([0-9]+[a-z]*)$', r'\1 \2', text)
    text = re.sub(r'^[a-z][0-9]+\s+', '', text).strip()
    while True:
        new_text = RE_TRAILING_NOISE.sub('', text).strip()
        if new_text == text:
            break
        text = new_text
    return text.strip()

def validate_and_format(match_str):
    if not match_str or match_str in ["No good match", "REQUIRES HUMAN", "SKIP"]:
        return match_str
    
    if "\n" in match_str:
        valid_parts = []
        for part in match_str.split("\n"):
            clean = canonical_casing.get(part.strip().lower())
            if clean: valid_parts.append(clean)
        if valid_parts: return "\n".join(valid_parts)
        return "No good match"
    
    clean_match = canonical_casing.get(match_str.strip().lower())
    return clean_match if clean_match else "No good match"

def load_knowledge_base():
    global lookup_phrases, lookup_phrases_original, canonical_casing
    global lookup_words_sets, lookup_embeddings
    global industry_translations, broad_categories
    global compiled_industry_translations, compiled_broad_categories
    global custom_rules

    print("--- Initializing Engine ---")

    template_folder = os.environ.get("TEMPLATE_FOLDER", r"C:\Users\pcarr\OneDrive\Documents\Custom Office Templates")
    search_pattern = os.path.join(template_folder, "*.xltm")
    list_of_files = glob.glob(search_pattern)

    if list_of_files:
        excel_file = max(list_of_files, key=os.path.getmtime)
        print(f"Dynamically locked onto newest template: {os.path.basename(excel_file)}")
    else:
        excel_file = ""
        print(f"WARNING: No .xltm files found in {template_folder}")

    lookup_phrases, lookup_phrases_original, canonical_casing = [], [], {}
    industry_translations, broad_categories = {}, []

    if os.path.exists(excel_file):
        try:
            print(f"Reading Holy Grail phrases from {excel_file}...")
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

                paired.sort(key=lambda x: len(x[0]), reverse=True)
                lookup_phrases = [p[0] for p in paired]
                lookup_phrases_original = [p[1] for p in paired]
                print(f"Successfully locked onto {len(lookup_phrases)} phrases!")
            else:
                print("WARNING: Could not find 'Asset Sub Type' column.")

            print(f"Reading Rules Sheet from {excel_file}...")
            df_rules = pd.read_excel(excel_file, sheet_name="Rules Sheet", header=None)
            for index, row in df_rules.iterrows():
                if pd.isna(row[0]) or pd.isna(row[1]): continue
                rule_type = str(row[0]).strip().lower()
                from_text = str(row[1]).strip().lower()
                to_text = str(row[2]).strip().replace('|', '\n') if not pd.isna(row[2]) else ""

                if to_text:
                    corrected_options = []
                    for opt in to_text.split('\n'):
                        corrected_options.append(canonical_casing.get(opt.strip().lower(), opt.strip()))
                    to_text = '\n'.join(corrected_options)

                if rule_type in ["alias", "phrase", "list"]:
                    industry_translations[from_text] = to_text
                elif rule_type == "core":
                    broad_categories.append(from_text)
            print(f"Loaded {len(industry_translations)} translations and {len(broad_categories)} core categories!")
        except Exception as e:
            print(f"ERROR reading Excel file: {e}")
    
    if not lookup_phrases:
        fallback_phrases = ["Air Handler", "Door Manual Swing", "Fan Coil Unit"]
        paired = []
        for phrase in fallback_phrases:
            lower_phrase = phrase.lower()
            paired.append((lower_phrase, phrase))
            canonical_casing[lower_phrase] = phrase
        paired.sort(key=lambda x: len(x[0]), reverse=True)
        lookup_phrases = [p[0] for p in paired]
        lookup_phrases_original = [p[1] for p in paired]

    if not industry_translations:
        industry_translations = {
            "sprinkler air compressor": canonical_casing.get("reciprocating air compressor", "Reciprocating Air Compressor")
        }
    if not broad_categories: broad_categories = ["boiler", "pump", "chiller", "compressor", "valve", "fan", "motor", "compactor"]

    compiled_industry_translations = []
    sorted_messy_terms = sorted(industry_translations.keys(), key=len, reverse=True)
    for messy_term in sorted_messy_terms:
        clean_term = industry_translations[messy_term]
        if clean_term:
            rule_tokens = get_stemmed_words(messy_term)
            compiled_industry_translations.append((rule_tokens, clean_term))

    compiled_broad_categories = [re.compile(r'\b' + re.escape(cat) + r'\b') for cat in broad_categories]

    print("Pre-computing stemmed word sets for overlap scoring...")
    # --- UPGRADED: Apply stemming to Holy Grail list ---
    lookup_words_sets = [get_stemmed_words(phrase) for phrase in lookup_phrases]

    print("Calculating vectors for lookup phrases...")
    embeddings = embedder.encode(lookup_phrases, batch_size=256)
    norms = np.linalg.norm(embeddings, axis=1, keepdims=True)
    lookup_embeddings = embeddings / (norms + 1e-9)

    if os.path.exists(custom_rules_file):
        try:
            with open(custom_rules_file, "r") as f:
                with custom_rules_lock:
                    custom_rules.clear()
                    raw_rules = json.load(f)
                    for k, v in raw_rules.items():
                        corrected_v = []
                        for opt in str(v).split('\n'):
                            corrected_v.append(canonical_casing.get(opt.strip().lower(), opt.strip()))
                        custom_rules[k] = '\n'.join(corrected_v)
            print(f"Loaded {len(custom_rules)} custom learned rules from memory!")
        except Exception as e:
            print(f"Failed to read custom_rules.json: {e}")

    print("Knowledge Base Engine Ready!")

# Run initialization at startup
load_knowledge_base()

# ==========================================
# 2. THE RULEBOOK CLEANERS
# ==========================================
def clean_contractor_noise(text):
    text = RE_PARENS.sub(' ', text)
    text = RE_NON_ALPHA.sub(' ', text)
    return RE_SPACES.sub(' ', text).strip()

# ==========================================
# 3. THE ENDPOINTS
# ==========================================
@app.route('/reload', methods=['POST'])
def reload_knowledge_base():
    try:
        load_knowledge_base()
        return jsonify({"status": "success", "message": "Knowledge base reloaded"}), 200
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/learn', methods=['POST'])
def learn_rule():
    data = request.json
    raw_phrase = data.get('phrase', '').strip().lower()
    clean_match = data.get('match', '').strip()

    corrected_match = canonical_casing.get(clean_match.lower(), clean_match)
    base_phrase = clean_contractor_noise(raw_phrase)
    signature = get_generalized_signature(base_phrase)

    if signature and corrected_match:
        with custom_rules_lock:
            custom_rules[signature] = corrected_match
            temp_file = custom_rules_file + ".tmp"
            with open(temp_file, "w") as f:
                json.dump(custom_rules, f, indent=4)
            os.replace(temp_file, custom_rules_file)

        print(f"🧠 AI LEARNED GENERAL RULE: [{signature}] -> [{corrected_match}]")
        return jsonify({"status": "success"}), 200

    return jsonify({"status": "ignored"}), 400

@app.route('/batch_lookup', methods=['POST'])
def batch_lookup():
    data = request.json
    items = data.get('items', [])
    results = []

    ai_queue_phrases = []
    ai_queue_rows = []
    ai_original_bases = []

    for item in items:
        original_phrase = item.get('phrase', '').strip().lower()
        row_id = item.get('row')
        current_match = item.get('current_match', '').strip().lower()
        current_id = item.get('current_id', '').strip().lower()

        base_phrase = clean_contractor_noise(original_phrase)
        signature = get_generalized_signature(base_phrase)

        is_broken = False
        if current_match in ["", "no good match", "no matching words"]: is_broken = True
        if any(x in current_id for x in ["none", "category", "requires human", "requires ai"]): is_broken = True

        if not is_broken:
            results.append({"row": row_id, "match": "SKIP", "id": "SKIP"})
            continue

        with custom_rules_lock:
            match_found = custom_rules.get(base_phrase) or custom_rules.get(signature)
            if not match_found:
                # --- UPGRADED: Apply stemming to memory check ---
                sig_words = get_stemmed_words(signature)
                best_mem_score = 0.0

                for rule_key, rule_val in custom_rules.items():
                    rule_words = get_stemmed_words(rule_key)
                    if not rule_words: continue

                    intersection = len(sig_words.intersection(rule_words))
                    if len(sig_words) + len(rule_words) > 0:
                        overlap = (2.0 * intersection) / (len(sig_words) + len(rule_words))
                    else:
                        overlap = 0.0

                    if overlap > best_mem_score and overlap >= 0.80:
                        best_mem_score = overlap
                        match_found = rule_val

        if match_found:
            validated = validate_and_format(match_found)
            if validated != "No good match":
                results.append({"row": row_id, "match": validated, "id": "USER_LEARNED"})
                continue

        if len(signature) > 2:
            ai_queue_phrases.append(signature)
            ai_queue_rows.append(row_id)
            ai_original_bases.append(base_phrase)
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

            candidates = item_candidates_list[i_queue]
            num_candidates = len(candidates)
            indices = [cand_to_idx[c] for c in candidates]
            candidate_norms = unique_norms[indices]

            # --- UPGRADED: Apply stemming to input words ---
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

                        if len(base_words) + len(lookup_words) > 0:
                            overlap = (2.0 * intersection) / (len(base_words) + len(lookup_words))
                        else:
                            overlap = 0.0

                        hidden_pattern_bonus = 0.15 if overlap == 1.0 else 0.0
                        overlap_cache[idx] = (overlap * 0.35) + hidden_pattern_bonus

                    combined_score = (sem_score * 0.50) + overlap_cache[idx]
                    collected_matches.append((combined_score, lookup_candidate_original))

            collected_matches.sort(key=lambda x: x[0], reverse=True)

            unique_matches = []
            seen_phrases = set()
            for score, phrase in collected_matches:
                if phrase not in seen_phrases:
                    unique_matches.append((score, phrase))
                    seen_phrases.add(phrase)

            final_match = ""
            final_id = ""

            best_combined_score = unique_matches[0][0] if unique_matches else 0.0

            if best_combined_score >= 0.85:
                proposed_match = unique_matches[0][1]
                validated = validate_and_format(proposed_match)
                if validated != "No good match":
                    final_match = validated
                    final_id = "AI_SMART_VECTOR"
                else:
                    final_match = "No good match"
                    final_id = "REQUIRES HUMAN"
            else:
                domain_match_found = False
                for rule_tokens, clean_term in compiled_industry_translations:
                    if rule_tokens.issubset(base_words):
                        validated = validate_and_format(clean_term.strip())
                        if validated != "No good match":
                            final_match = validated
                            final_id = "EXPERT_RULE"
                            domain_match_found = True
                            break

                if not domain_match_found:
                    if best_combined_score >= 0.40:
                        close_matches = [m[1] for m in unique_matches if m[0] >= 0.40 and (best_combined_score - m[0] <= 0.15)]
                        close_matches = close_matches[:5]

                        if len(close_matches) > 1:
                            final_match = validate_and_format("\n".join(close_matches))
                            final_id = "AI_SUGGESTED_LIST"
                        else:
                            final_match = validate_and_format(unique_matches[0][1])
                            final_id = "AI_HYBRID_MATCH"
                    else:
                        broad_match_found = False
                        for pattern in compiled_broad_categories:
                            if pattern.search(signature):
                                final_match = "Subtype missing in input"
                                final_id = "REQUIRES HUMAN"
                                broad_match_found = True
                                break

                        if not broad_match_found:
                            final_match = "No good match"
                            final_id = "REQUIRES HUMAN"

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
    ai_original_bases = []

    for index, row in df.iterrows():
        row_id = str(row['Row'])
        original_phrase = str(row['Phrase']).strip().lower()

        base_phrase = clean_contractor_noise(original_phrase)
        signature = get_generalized_signature(base_phrase)

        if len(signature) > 2:
            ai_queue_phrases.append(signature)
            ai_queue_rows.append(row_id)
            ai_original_bases.append(base_phrase)
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
            original_base = ai_original_bases[i_queue]

            candidates = item_candidates_list[i_queue]
            num_candidates = len(candidates)
            indices = [cand_to_idx[c] for c in candidates]
            candidate_norms = unique_norms[indices]

            # --- UPGRADED: Apply stemming to input words ---
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

                        if len(base_words) + len(lookup_words) > 0:
                            overlap = (2.0 * intersection) / (len(base_words) + len(lookup_words))
                        else:
                            overlap = 0.0

                        hidden_pattern_bonus = 0.15 if overlap == 1.0 else 0.0
                        overlap_cache[idx] = (overlap * 0.35) + hidden_pattern_bonus

                    combined_score = (sem_score * 0.50) + overlap_cache[idx]
                    collected_matches.append((combined_score, lookup_candidate_original))

            collected_matches.sort(key=lambda x: x[0], reverse=True)

            unique_matches = []
            seen_phrases = set()
            for score, phrase in collected_matches:
                if phrase not in seen_phrases:
                    unique_matches.append((score, phrase))
                    seen_phrases.add(phrase)

            final_match = ""
            final_id = ""

            best_combined_score = unique_matches[0][0] if unique_matches else 0.0

            if best_combined_score >= 0.85:
                proposed_match = unique_matches[0][1]
                validated = validate_and_format(proposed_match)
                if validated != "No good match":
                    final_match = validated
                    final_id = "AI_SMART_VECTOR"
                else:
                    final_match = "No good match"
                    final_id = "REQUIRES HUMAN"
            else:
                with custom_rules_lock:
                    match_found = custom_rules.get(original_base) or custom_rules.get(signature)
                    if not match_found:
                        # --- UPGRADED: Apply stemming to memory check ---
                        sig_words = get_stemmed_words(signature)
                        best_mem_score = 0.0

                        for rule_key, rule_val in custom_rules.items():
                            rule_words = get_stemmed_words(rule_key)
                            if not rule_words: continue

                            intersection = len(sig_words.intersection(rule_words))
                            if len(sig_words) + len(rule_words) > 0:
                                overlap = (2.0 * intersection) / (len(sig_words) + len(rule_words))
                            else:
                                overlap = 0.0

                            if overlap > best_mem_score and overlap >= 0.80:
                                best_mem_score = overlap
                                match_found = rule_val

                if match_found:
                    validated = validate_and_format(match_found)
                    if validated != "No good match":
                        final_match = validated
                        final_id = "USER_LEARNED"
                    else:
                        final_match = "No good match"
                        final_id = "REQUIRES HUMAN"
                else:
                    domain_match_found = False
                    for rule_tokens, clean_term in compiled_industry_translations:
                        if rule_tokens.issubset(base_words):
                            validated = validate_and_format(clean_term.strip())
                            if validated != "No good match":
                                final_match = validated
                                final_id = "EXPERT_RULE"
                                domain_match_found = True
                                break

                    if not domain_match_found:
                        if best_combined_score >= 0.40:
                            close_matches = [m[1] for m in unique_matches if m[0] >= 0.40 and (best_combined_score - m[0] <= 0.15)]
                            close_matches = close_matches[:5]

                            if len(close_matches) > 1:
                                final_match = validate_and_format("\n".join(close_matches))
                                final_id = "AI_SUGGESTED_LIST"
                            else:
                                final_match = validate_and_format(unique_matches[0][1])
                                final_id = "AI_HYBRID_MATCH"
                        else:
                            broad_match_found = False
                            for pattern in compiled_broad_categories:
                                if pattern.search(signature):
                                    final_match = "Subtype missing in input"
                                    final_id = "REQUIRES HUMAN"
                                    broad_match_found = True
                                    break

                            if not broad_match_found:
                                final_match = "No good match"
                                final_id = "REQUIRES HUMAN"

            results.append({"Row": row_id, "Match": final_match, "Id": final_id})

    for r in results:
        if isinstance(r["Match"], str):
            r["Match"] = r["Match"].replace('\n', '\\n')

    df_out = pd.DataFrame(results)
    df_out.to_csv(output_file, sep='\t', index=False)

    return jsonify({"status": "success", "processed": len(results)}), 200

if __name__ == '__main__':
    app.run(port=5000, threaded=True)