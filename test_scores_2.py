from sentence_transformers import SentenceTransformer
import numpy as np

embedder = SentenceTransformer('all-MiniLM-L6-v2')

lookup_phrases = ["pump", "chilled water pump", "exhaust fan", "fan", "air handling unit", "rtu", "fan coil unit", "door manual swing"]
lookup_embeddings = embedder.encode(lookup_phrases)
lookup_norms = lookup_embeddings / (np.linalg.norm(lookup_embeddings, axis=1, keepdims=True) + 1e-9)

def test_phrase(signature):
    print(f"\n--- Testing Signature: '{signature}' ---")
    candidates_set = {signature}
    words = signature.split()
    for n in range(2, 6):
        if len(words) >= n:
            for j in range(len(words) - n + 1):
                candidates_set.add(' '.join(words[j:j+n]))
    candidates = list(candidates_set)
    candidate_vectors = embedder.encode(candidates)
    candidate_norms = candidate_vectors / (np.linalg.norm(candidate_vectors, axis=1, keepdims=True) + 1e-9)

    all_semantic_scores = np.dot(candidate_norms, lookup_norms.T)

    base_words = set(signature.split())

    results = []

    for c_idx, candidate in enumerate(candidates):
        for l_idx, lookup in enumerate(lookup_phrases):
            sem_score = all_semantic_scores[c_idx, l_idx]
            lookup_words = set(lookup.split())

            substring_bonus = 0.1 if lookup in signature else 0.0

            # Old overlap
            if len(lookup_words) > 0:
                old_overlap = len(base_words.intersection(lookup_words)) / len(lookup_words)
            else:
                old_overlap = 0.0

            old_combined = (sem_score * 0.60) + (old_overlap * 0.30) + substring_bonus

            # New overlap (Dice / F1)
            intersection = len(base_words.intersection(lookup_words))
            if len(base_words) + len(lookup_words) > 0:
                new_overlap = (2.0 * intersection) / (len(base_words) + len(lookup_words))
            else:
                new_overlap = 0.0

            new_combined = (sem_score * 0.60) + (new_overlap * 0.30) + substring_bonus

            results.append({
                "candidate": candidate,
                "lookup": lookup,
                "sem_score": sem_score,
                "old_overlap": old_overlap,
                "new_overlap": new_overlap,
                "old_combined": old_combined,
                "new_combined": new_combined
            })

    # Print top 3 old and new
    results_old = sorted(results, key=lambda x: x["old_combined"], reverse=True)
    results_new = sorted(results, key=lambda x: x["new_combined"], reverse=True)

    print("BEST OLD:")
    for i in range(3):
        res = results_old[i]
        print(f"  {i+1}: '{res['lookup']}' via cand '{res['candidate']}' (score {res['old_combined']:.3f}, sem: {res['sem_score']:.3f}, ol: {res['old_overlap']:.3f})")

    print("BEST NEW:")
    for i in range(3):
        res = results_new[i]
        print(f"  {i+1}: '{res['lookup']}' via cand '{res['candidate']}' (score {res['new_combined']:.3f}, sem: {res['sem_score']:.3f}, ol: {res['new_overlap']:.3f})")

test_phrase("chilled water pump roof")
test_phrase("fcu")
test_phrase("door 123")
