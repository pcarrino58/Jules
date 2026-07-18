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

# --- NEW: FORCE HARDWARE ACCELERATION ---
device = "cuda" if torch.cuda.is_available() else "cpu"
print(f"Engine forcing hardware acceleration on: [{device.upper()}]")

embedder = SentenceTransformer('all-MiniLM-L6-v2', device=device)

script_dir = os.path.dirname(os.path.abspath(__file__))
custom_rules_file = os.path.join(script_dir, "custom_rules.json")

# Lock for thread-safety when editing or reading `custom_rules`
custom_rules = {}
custom_rules_lock = threading.Lock()

# Global variables for lookup data
lookup_phrases = []
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

# --- UPGRADED GENERALIZER ---
RE_TRAILING_NOISE = re.compile(r'\s+\b([0-9]+[a-z]*|[a-z])\b$')

def get_generalized_signature(text):
    """Aggressively strips sequence numbers, trailing letters, and codes (e.g., '#1', '1A', 'A', 'Tank2')"""
    # Detach numbers that are crammed against words (e.g., "unit2a" -> "unit 2a")
    text = re.sub(r'([a-z])([0-9]+[a-z]*)$', r'\1 \2', text)

    # Repeatedly shave off trailing numbers, single letters, or alphanumeric codes
    while True:
        new_text = RE_TRAILING_NOISE.sub('', text).strip()
        if new_text == text:
            break
        text = new_text

    return text.strip()

def load_knowledge_base():
    """Loads all templates, rules, and vectorizes them."""
    global lookup_phrases, lookup_words_sets, lookup_embeddings
    global industry_translations, broad_categories
    global compiled_industry_translations, compiled_broad_categories
    global custom_rules

    # --- NEW DYNAMIC FILE LOCATOR ---
    template_folder = os.environ.get(
        "TEMPLATE_FOLDER",
        r"C:\Users\pcarr\OneDrive\Documents\Custom Office Templates"
    )

    search_pattern = os.path.join(template_folder, "*.xltm")
    list_of_files = glob.glob(search_pattern)

    if list_of_files:
        excel_file = max(list_of_files, key=os.path.getmtime)
        print(f"Dynamically locked onto newest template: {os.path.basename(excel_file)}")
    else:
        excel_file = ""
        print(f"WARNING: No .xltm files found in {template_folder}")

    # Reset lists
    lookup_phrases = []
    industry_translations = {}
    broad_categories = []

    if os.path.exists(excel_file):
        try:
            print(f"Reading Holy Grail phrases from {excel_file}...")
            df_lookup = pd.read_excel(excel_file, sheet_name="Uniformat RS Means Lookup")

            if 'Asset Sub Type' in df_lookup.columns:
                raw_phrases = df_lookup['Asset Sub Type'].dropna().astype(str).tolist()
                lookup_phrases = [str(phrase).strip().lower() for phrase in raw_phrases if str(phrase).strip()]
                lookup_phrases.sort(key=len, reverse=True)
                print(f"Successfully locked onto {len(lookup_phrases)} phrases!")
            else:
                print("WARNING: Could not find 'Asset Sub Type' column.")

            print(f"Reading Rules Sheet from {excel_file}...")
            try:
                df_rules = pd.read_excel(excel_file, sheet_name="Rules Sheet", header=None)
                for index, row in df_rules.iterrows():
                    if pd.isna(row[0]) or pd.isna(row[1]):
                        continue

                    rule_type = str(row[0]).strip().lower()
                    from_text = str(row[1]).strip().lower()
                    # Safely parse pipes into newlines for VBA dropdowns
                    to_text = str(row[2]).strip().replace('|', '\n') if not pd.isna(row[2]) else ""

                    if rule_type in ["alias", "phrase", "list"]:
                        industry_translations[from_text] = to_text
                    elif rule_type == "core":
                        broad_categories.append(from_text)

                print(f"Loaded {len(industry_translations)} translations and {len(broad_categories)} core categories!")
            except ValueError:
                print("WARNING: 'Rules Sheet' tab not found.")
        except Exception as e:
            print(f"ERROR reading Excel file: {e}")
    else:
        print(f"WARNING: {excel_file} not found!")

    # Fallbacks
    if not lookup_phrases:
        lookup_phrases = ["air handler", "door manual swing", "fan coil unit"]
    if not industry_translations:
        industry_translations = {
            "sprinkler air compressor": "reciprocating air compressor",
            "ah": "Air Handling Unit 25 to 50 Tons",
            "rtu": "Air Cooled Package AC Unit 25 to 50 Tons",
            "crac": "Computer Room Direct Expansion Package AC Unit"
        }
    if not broad_categories:
        broad_categories = ["boiler", "pump", "chiller", "compressor", "valve", "fan", "motor", "compactor"]

    # Pre-compile regex patterns for massive performance gains
    compiled_industry_translations = []

    # NEW: Sort rules by length (longest first) to prevent short words from hijacking long phrases
    sorted_messy_terms = sorted(industry_translations.keys(), key=len, reverse=True)

    for messy_term in sorted_messy_terms:
        clean_term = industry_translations[messy_term]
        if clean_term:
            pattern = re.compile(r'\b' + re.escape(messy_term) + r'(?:s|es)?\b')
            compiled_industry_translations.append((pattern, clean_term))

    compiled_broad_categories = [re.compile(r'\b' + re.escape(cat) + r'\b') for cat in broad_categories]

    # Pre-compute word sets for fast O(1) intersection during batch lookups
    print("Pre-computing word sets for overlap scoring...")
    lookup_words_sets = [set(phrase.split()) for phrase in lookup_phrases]

    print("Calculating vectors for lookup phrases...")
    embeddings = embedder.encode(lookup_phrases, batch_size=256)
    norms = np.linalg.norm(embeddings, axis=1, keepdims=True)
    lookup_embeddings = embeddings / (norms + 1e-9)

    if os.path.exists(custom_rules_file):
        try:
            with open(custom_rules_file, "r") as f:
                with custom_rules_lock:
                    custom_rules.clear()
                    custom_rules.update(json.load(f))
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

def get_generalized_signature(text):
    """Aggressively strips sequence numbers, trailing letters, and codes (e.g., '#1', '1A', 'A', 'Tank2')"""
    # Detach numbers that are crammed against words (e.g., "unit2a" -> "unit 2a")
    text = re.sub(r'([a-z])([0-9]+[a-z]*)$', r'\1 \2', text)

    # Repeatedly shave off trailing numbers, single letters, or alphanumeric codes
    while True:
        new_text = RE_TRAILING_NOISE.sub('', text).strip() # Fixed the name here
        if new_text == text:
            break
        text = new_text

    return text.strip()

# ==========================================
# 3. THE ENDPOINTS
# ==========================================
@app.route('/reload', methods=['POST'])
def reload_knowledge_base():
    """Endpoint to dynamically reload data without restarting the server."""
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

    base_phrase = clean_contractor_noise(raw_phrase)
    signature = get_generalized_signature(base_phrase)

    if signature and clean_match:
        with custom_rules_lock:
            custom_rules[signature] = clean_match
            # Safely write to a temporary file, then atomic replace
            temp_file = custom_rules_file + ".tmp"
            with open(temp_file, "w") as f:
                json.dump(custom_rules, f, indent=4)
            os.replace(temp_file, custom_rules_file)

        print(f"🧠 AI LEARNED GENERAL RULE: [{signature}] -> [{clean_match}]")
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

        # RULING 1: THE SPEED PASS
        is_broken = False
        if current_match in ["", "no good match", "no matching words"]: is_broken = True
        if any(x in current_id for x in ["none", "category", "requires human", "requires ai"]): is_broken = True

        if not is_broken:
            results.append({"row": row_id, "match": "SKIP", "id": "SKIP"})
            continue

        # RULING 1.5: USER'S OWN LEARNED RULES!
        with custom_rules_lock:
            match_found = custom_rules.get(base_phrase) or custom_rules.get(signature)

        if match_found:
            results.append({"row": row_id, "match": match_found, "id": "USER_LEARNED"})
            continue

        # RULING 2: EXPERT OVERRIDE
        domain_match_found = False
        for pattern, clean_term in compiled_industry_translations:
            if pattern.search(signature):
                results.append({"row": row_id, "match": clean_term.strip(), "id": "EXPERT_RULE"})
                domain_match_found = True
                break
        if domain_match_found: continue

        # RULING 3: BROAD CATEGORY OVERRIDE
        broad_match_found = False
        for pattern in compiled_broad_categories:
            if pattern.search(signature):
                results.append({"row": row_id, "match": "Subtype missing in input", "id": "REQUIRES HUMAN"})
                broad_match_found = True
                break
        if broad_match_found: continue

        # RULING 4: QUEUE FOR AI MATH
        if len(signature) > 2:
            ai_queue_phrases.append(signature)
            ai_queue_rows.append(row_id)
            ai_original_bases.append(base_phrase)
        else:
            results.append({"row": row_id, "match": "No good match", "id": "REQUIRES HUMAN"})

    # ==========================================
    # 4. BATCH HYBRID MATH
    # ==========================================
    if ai_queue_phrases:
        all_candidates = []
        item_candidate_counts = []

        for signature in ai_queue_phrases:
            candidates_set = {signature}
            words = signature.split()
            for n in range(2, 6):
                if len(words) >= n:
                    for j in range(len(words) - n + 1):
                        candidates_set.add(' '.join(words[j:j+n]))

            candidates = list(candidates_set)
            all_candidates.extend(candidates)
            item_candidate_counts.append(len(candidates))

        all_candidate_vectors = embedder.encode(all_candidates, batch_size=256)
        all_candidate_norms = all_candidate_vectors / (np.linalg.norm(all_candidate_vectors, axis=1, keepdims=True) + 1e-9)

        offset = 0
        top_k = min(20, len(lookup_phrases))

        for i_queue, signature in enumerate(ai_queue_phrases):
            row_id = ai_queue_rows[i_queue]
            original_base = ai_original_bases[i_queue]
            num_candidates = item_candidate_counts[i_queue]

            candidate_norms = all_candidate_norms[offset:offset+num_candidates]
            offset += num_candidates

            base_words = set(signature.split())

            best_idx = -1
            best_combined_score = 0.0

            all_semantic_scores = np.dot(candidate_norms, lookup_embeddings.T)

            for c_idx in range(num_candidates):
                semantic_scores = all_semantic_scores[c_idx]
                top_indices = np.argsort(semantic_scores)[-top_k:][::-1]

                for idx in top_indices:
                    sem_score = semantic_scores[idx]
                    lookup_candidate = lookup_phrases[idx]
                    lookup_words = lookup_words_sets[idx]

                    substring_bonus = 0.1 if lookup_candidate in signature else 0.0

                    if len(lookup_words) > 0:
                        overlap = len(base_words.intersection(lookup_words)) / len(lookup_words)
                    else:
                        overlap = 0.0

                    combined_score = (sem_score * 0.60) + (overlap * 0.30) + substring_bonus

                    if combined_score > best_combined_score:
                        best_combined_score = combined_score
                        best_idx = idx

            if best_combined_score >= 0.85:
                final_match = lookup_phrases[best_idx]
                final_id = "AI_SMART_VECTOR"
            elif best_combined_score >= 0.40:
                final_match = lookup_phrases[best_idx]
                final_id = "AI_HYBRID_MATCH"
            else:
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

        # QUEUE EVERYTHING FOR HOLY GRAIL MATH
        # This acts as a safeguard to ensure single-word locations
        # or tiny strings aren't forced into aggressive vector guessing
        if len(signature) > 2:
            ai_queue_phrases.append(signature)
            ai_queue_rows.append(row_id)
            ai_original_bases.append(base_phrase)
        else:
            results.append({"Row": row_id, "Match": "No good match", "Id": "REQUIRES HUMAN"})

    # ==========================================
    # 4. BATCH HYBRID MATH (Holy Grail Priority)
    # ==========================================
    if ai_queue_phrases:
        all_candidates = []
        item_candidate_counts = []

        for signature in ai_queue_phrases:
            candidates_set = {signature}
            words = signature.split()
            for n in range(2, 6):
                if len(words) >= n:
                    for j in range(len(words) - n + 1):
                        candidates_set.add(' '.join(words[j:j+n]))

            candidates = list(candidates_set)
            all_candidates.extend(candidates)
            item_candidate_counts.append(len(candidates))

        all_candidate_vectors = embedder.encode(all_candidates, batch_size=256)
        all_candidate_norms = all_candidate_vectors / (np.linalg.norm(all_candidate_vectors, axis=1, keepdims=True) + 1e-9)

        offset = 0
        top_k = min(20, len(lookup_phrases))

        for i_queue, signature in enumerate(ai_queue_phrases):
            row_id = ai_queue_rows[i_queue]
            original_base = ai_original_bases[i_queue]
            num_candidates = item_candidate_counts[i_queue]

            candidate_norms = all_candidate_norms[offset:offset+num_candidates]
            offset += num_candidates

            base_words = set(signature.split())
            best_idx = -1
            best_combined_score = 0.0

            all_semantic_scores = np.dot(candidate_norms, lookup_embeddings.T)

            for c_idx in range(num_candidates):
                semantic_scores = all_semantic_scores[c_idx]
                top_indices = np.argsort(semantic_scores)[-top_k:][::-1]

                for idx in top_indices:
                    sem_score = semantic_scores[idx]
                    lookup_candidate = lookup_phrases[idx]
                    lookup_words = lookup_words_sets[idx]

                    substring_bonus = 0.1 if lookup_candidate in signature else 0.0

                    if len(lookup_words) > 0:
                        overlap = len(base_words.intersection(lookup_words)) / len(lookup_words)
                    else:
                        overlap = 0.0

                    combined_score = (sem_score * 0.60) + (overlap * 0.30) + substring_bonus

                    if combined_score > best_combined_score:
                        best_combined_score = combined_score
                        best_idx = idx

            # --- THE PURE AI HIERARCHY ---
            final_match = ""
            final_id = ""

            # PRIORITY 1: Holy Grail AI Match (Strict 85%)
            if best_combined_score >= 0.85:
                final_match = lookup_phrases[best_idx]
                final_id = "AI_SMART_VECTOR"
            else:
                # PRIORITY 2: User's Learned Memory
                with custom_rules_lock:
                    # Check for older exact rules first, then the new generalized signature
                    match_found = custom_rules.get(original_base) or custom_rules.get(signature)

                if match_found:
                    final_match = match_found
                    final_id = "USER_LEARNED"
                else:
                    # PRIORITY 3: Spreadsheet Expert Rules
                    domain_match_found = False
                    for pattern, clean_term in compiled_industry_translations:
                        if pattern.search(signature):
                            final_match = clean_term.strip()
                            final_id = "EXPERT_RULE"
                            domain_match_found = True
                            break

                    if not domain_match_found:
                        # PRIORITY 4: AI Hybrid Guess (Let the AI try first!)
                        if best_combined_score >= 0.40:
                            final_match = lookup_phrases[best_idx]
                            final_id = "AI_HYBRID_MATCH"
                        else:
                            # PRIORITY 5: Broad Categories (The final safety net)
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

    # Protect the TSV file from rogue newlines
    for r in results:
        if isinstance(r["Match"], str):
            r["Match"] = r["Match"].replace('\n', '\\n')

    df_out = pd.DataFrame(results)
    df_out.to_csv(output_file, sep='\t', index=False)

    return jsonify({"status": "success", "processed": len(results)}), 200

if __name__ == '__main__':
    app.run(port=5000, threaded=True)
