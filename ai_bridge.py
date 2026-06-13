from flask import Flask, request
from sentence_transformers import SentenceTransformer
import numpy as np
import pandas as pd
import os
import glob  # <-- Added glob to search for files
import re
import json
import threading

app = Flask(__name__)

# ==========================================
# 1. INITIALIZE PURE MATH RADAR
# ==========================================
print("Loading Pure Semantic Math Engine...")
embedder = SentenceTransformer('all-MiniLM-L6-v2')

script_dir = os.path.dirname(os.path.abspath(__file__))
custom_rules_file = os.path.join(script_dir, "custom_rules.json")

# --- NEW DYNAMIC FILE LOCATOR ---
template_folder = r"C:\Users\pcarr\OneDrive\Documents\Custom Office Templates"

# Updated to specifically target your .xltm template files
search_pattern = os.path.join(template_folder, "*.xltm") 
list_of_files = glob.glob(search_pattern)

if list_of_files:
    # Grab the file with the most recent modification time
    excel_file = max(list_of_files, key=os.path.getmtime)
    print(f"Dynamically locked onto newest template: {os.path.basename(excel_file)}")
else:
    excel_file = ""
    print(f"WARNING: No .xltm files found in {template_folder}")
# --------------------------------

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
                to_text = str(row[2]).strip() if not pd.isna(row[2]) else ""
                
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

# Pre-compile regex patterns for massive performance gains in `/batch_lookup`
compiled_industry_translations = []
for messy_term, clean_term in industry_translations.items():
    if clean_term:
        pattern = re.compile(r'\b' + re.escape(messy_term) + r'(?:s|es)?\b')
        compiled_industry_translations.append((pattern, clean_term))

compiled_broad_categories = [re.compile(r'\b' + re.escape(cat) + r'\b') for cat in broad_categories]

# Pre-compile noise removal regexes
RE_PARENS = re.compile(r'\(.*?\)')
RE_NON_ALPHA = re.compile(r'[^a-z]')
RE_SPACES = re.compile(r'\s+')

print("Calculating vectors for lookup phrases...")
lookup_embeddings = embedder.encode(lookup_phrases)
# Adding small epsilon (1e-9) to prevent division by zero
lookup_norms = np.linalg.norm(lookup_embeddings, axis=1, keepdims=True)
lookup_embeddings = lookup_embeddings / (lookup_norms + 1e-9)

# Lock for thread-safety when editing or reading `custom_rules`
custom_rules = {}
custom_rules_lock = threading.Lock()

if os.path.exists(custom_rules_file):
    try:
        with open(custom_rules_file, "r") as f:
            custom_rules = json.load(f)
        print(f"Loaded {len(custom_rules)} custom learned rules from memory!")
    except Exception as e:
        print(f"Failed to read custom_rules.json: {e}")

print("Engine Ready!")

# ==========================================
# 2. THE RULEBOOK CLEANER
# ==========================================
def clean_contractor_noise(text):
    text = RE_PARENS.sub(' ', text)
    text = RE_NON_ALPHA.sub(' ', text)
    return RE_SPACES.sub(' ', text).strip()

# ==========================================
# 3. THE ENDPOINTS
# ==========================================
@app.route('/learn', methods=['POST'])
def learn_rule():
    data = request.json
    raw_phrase = data.get('phrase', '').strip().lower()
    clean_match = data.get('match', '').strip()
    
    base_phrase = clean_contractor_noise(raw_phrase)
    
    if base_phrase and clean_match:
        with custom_rules_lock:
            custom_rules[base_phrase] = clean_match
            # Safely write to a temporary file, then atomic replace
            # This ensures custom_rules_file isn't corrupted if server crashes mid-write.
            temp_file = custom_rules_file + ".tmp"
            with open(temp_file, "w") as f:
                json.dump(custom_rules, f, indent=4)
            os.replace(temp_file, custom_rules_file)
            
        print(f"🧠 AI LEARNED: [{base_phrase}] -> [{clean_match}]")
        return {"status": "success"}, 200
        
    return {"status": "ignored"}, 400


@app.route('/batch_lookup', methods=['POST'])
def batch_lookup():
    data = request.json
    items = data.get('items', [])
    results = []
    
    ai_queue_phrases = []
    ai_queue_rows = []

    for item in items:
        original_phrase = item.get('phrase', '').strip().lower()
        row_id = item.get('row')
        current_match = item.get('current_match', '').strip().lower()
        current_id = item.get('current_id', '').strip().lower()
        
        base_phrase = clean_contractor_noise(original_phrase)

        # RULING 1: THE SPEED PASS
        is_broken = False
        if current_match in ["", "no good match", "no matching words"]: is_broken = True
        if any(x in current_id for x in ["none", "category", "requires human", "requires ai"]): is_broken = True
        
        if not is_broken:
            results.append({"row": row_id, "match": "SKIP", "id": "SKIP"})
            continue

        # RULING 1.5: USER'S OWN LEARNED RULES!
        with custom_rules_lock:
            match_found = custom_rules.get(base_phrase)
        
        if match_found:
            results.append({"row": row_id, "match": match_found, "id": "USER_LEARNED"})
            continue

        # RULING 2: EXPERT OVERRIDE (Using Pre-compiled regex)
        domain_match_found = False
        for pattern, clean_term in compiled_industry_translations:
            if pattern.search(base_phrase):
                results.append({"row": row_id, "match": clean_term.strip(), "id": "EXPERT_RULE"})
                domain_match_found = True
                break 
        if domain_match_found: continue 

        # RULING 3: BROAD CATEGORY OVERRIDE (Using Pre-compiled regex)
        broad_match_found = False
        for pattern in compiled_broad_categories:
            if pattern.search(base_phrase):
                results.append({"row": row_id, "match": "Subtype missing in input", "id": "REQUIRES HUMAN"})
                broad_match_found = True
                break
        if broad_match_found: continue

        # RULING 4: QUEUE FOR AI MATH
        if len(base_phrase) > 2:
            ai_queue_phrases.append(base_phrase)
            ai_queue_rows.append(row_id)
        else:
            results.append({"row": row_id, "match": "No good match", "id": "REQUIRES HUMAN"})

    # ==========================================
    # 4. BATCH HYBRID MATH 
    # ==========================================
    if ai_queue_phrases:
        # Flatten all candidates to encode in a single batch for maximum GPU/CPU efficiency
        all_candidates = []
        item_candidate_counts = []
        
        for base_phrase in ai_queue_phrases:
            candidates = [base_phrase]
            words = base_phrase.split()
            for n in range(2, 6):
                if len(words) >= n:
                    for j in range(len(words) - n + 1):
                        candidates.append(' '.join(words[j:j+n]))
            all_candidates.extend(candidates)
            item_candidate_counts.append(len(candidates))
            
        # Encode all candidates in one shot
        all_candidate_vectors = embedder.encode(all_candidates)
        all_candidate_norms = all_candidate_vectors / (np.linalg.norm(all_candidate_vectors, axis=1, keepdims=True) + 1e-9)
        
        offset = 0
        for i_queue, base_phrase in enumerate(ai_queue_phrases):
            row_id = ai_queue_rows[i_queue]
            num_candidates = item_candidate_counts[i_queue]
            
            candidate_norms = all_candidate_norms[offset:offset+num_candidates]
            offset += num_candidates
            
            base_words = set(base_phrase.split())
            
            best_idx = -1
            best_combined_score = 0.0
            
            # Matrix multiplication for semantic scores of all candidates in this batch item
            all_semantic_scores = np.dot(candidate_norms, lookup_embeddings.T)
            
            for c_idx in range(num_candidates):
                semantic_scores = all_semantic_scores[c_idx]
                top_indices = np.argsort(semantic_scores)[-20:][::-1]
                
                for idx in top_indices:
                    sem_score = semantic_scores[idx]
                    lookup_candidate = lookup_phrases[idx]
                    lookup_words = set(lookup_candidate.split())
                    
                    substring_bonus = 0.1 if lookup_candidate in base_phrase else 0.0
                    
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
            elif best_combined_score >= 0.55:
                final_match = lookup_phrases[best_idx]
                final_id = "AI_HYBRID_MATCH"
            else:
                final_match = "No good match"
                final_id = "REQUIRES HUMAN"
                
            results.append({"row": row_id, "match": final_match, "id": final_id})

    tight_json = json.dumps(results, separators=(',', ':'))
    return app.response_class(response=tight_json, mimetype='application/json')

if __name__ == '__main__':
    app.run(port=5000, threaded=True)