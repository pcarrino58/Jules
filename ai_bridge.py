from flask import Flask, request, jsonify
import ollama
import chromadb
import hashlib

app = Flask(__name__)

# 1. Connect to your local AI vault (Using your original, highly accurate database)
client = chromadb.PersistentClient(path="./asset_vault")
collection = client.get_or_create_collection(name="facility_assets")

# ==========================================
# ROUTE 1: The "Thinking" Lookup Function (Untouched & 95% Accurate)
# ==========================================
@app.route('/lookup', methods=['POST'])
def lookup():
    data = request.json
    messy_text = data.get("phrase", "")
    
    if not messy_text:
        return jsonify({"error": "No phrase provided"}), 400

    try:
        # STEP 1: THE LIBRARIAN (Get top 5 fuzzy matches)
        response = ollama.embed(model='mxbai-embed-large', input=messy_text)
        query_vector = response['embeddings'][0]
        
        results = collection.query(
            query_embeddings=[query_vector],
            n_results=5 
        )
        
        if not results['documents'] or len(results['documents'][0]) == 0:
            return jsonify({"match": "No match found", "id": "UNKNOWN"})
            
        # Extract those 5 options into a clean list
        top_5_options = results['documents'][0]
        
        # Remove duplicates just in case the vault has multiples
        unique_options = list(set(top_5_options))
        options_string = "\n".join(f"- {opt}" for opt in unique_options)

        # STEP 2: THE MECHANIC (Reasoning Layer)
        system_prompt = f"""You are an expert facility asset manager. 
Your task is to identify the primary piece of equipment in a messy vendor string. 
Ignore location codes, building numbers, and secondary parts.

Messy String: "{messy_text}"

Available Options:
{options_string}

Pick the ONE option from the list that best represents the primary equipment. 
Return ONLY the exact text of that option. Do not include any punctuation, explanations, or conversational text."""

        # Ask the reasoning model to make the final choice
        llm_response = ollama.chat(model='phi3', messages=[
            {
                'role': 'user',
                'content': system_prompt
            }
        ])
        
        # Clean up the output to ensure it matches exactly
        final_winner = llm_response['message']['content'].strip()

        # Fallback: If the LLM goes rogue and types a paragraph, return the #1 vector match
        if final_winner not in unique_options:
            final_winner = top_5_options[0]

        return jsonify({"match": final_winner, "id": "AI_REASONED"})
            
    except Exception as e:
        print(f"Error during lookup: {e}")
        return jsonify({"error": str(e)}), 500


# ==========================================
# ROUTE 2: The Universal "Learn" Function (Updated)
# ==========================================
@app.route('/teach', methods=['POST'])
@app.route('/learn', methods=['POST'])
def learn():
    data = request.json
    print(f"\n📥 INCOMING LEARNING DATA: {data}")
    
    # 1. Try to grab the data using standard labels
    messy_text = data.get("messy_phrase") or data.get("phrase") or data.get("input", "")
    official_name = data.get("official_name") or data.get("category") or data.get("output", "")
    
    # 2. If the labels don't match, force extract the raw values
    if not messy_text or not official_name:
        values = list(data.values())
        if len(values) >= 2:
            messy_text = str(values[0]).strip()
            official_name = str(values[1]).strip()

    if not messy_text or not official_name:
        print("❌ ERROR: Missing data for learning")
        return jsonify({"error": "Missing data"}), 400

    try:
        response = ollama.embed(model='mxbai-embed-large', input=messy_text)
        learned_vector = response['embeddings'][0]
        
        phrase_hash = hashlib.md5(messy_text.encode('utf-8')).hexdigest()
        unique_learned_id = f"learned_{phrase_hash}"
        
        collection.add(
            ids=[unique_learned_id],
            embeddings=[learned_vector],
            documents=[official_name]
        )
        
        print(f"✅ Successfully learned: '{messy_text}' -> '{official_name}'")
        return jsonify({"status": "success"})
        
    except Exception as e:
        print(f"❌ Error during learn: {e}")
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    print("=========================================")
    print(" RAG AI Bridge is LIVE. Waiting for Excel... ")
    print("=========================================")
    app.run(port=5000)