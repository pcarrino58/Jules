def test_overlap(base, lookup):
    base_words = set(base.split())
    lookup_words = set(lookup.split())

    intersection = len(base_words.intersection(lookup_words))

    old_overlap = intersection / len(lookup_words) if lookup_words else 0

    precision = intersection / len(lookup_words) if lookup_words else 0
    recall = intersection / len(base_words) if base_words else 0

    if precision + recall > 0:
        f1_overlap = (2 * precision * recall) / (precision + recall)
    else:
        f1_overlap = 0.0

    print(f"Base: '{base}' | Lookup: '{lookup}'")
    print(f"  Old Overlap: {old_overlap:.2f}")
    print(f"  F1 Overlap:  {f1_overlap:.2f}\n")

test_overlap("chilled water pump roof", "pump")
test_overlap("chilled water pump roof", "chilled water pump")
test_overlap("fan", "exhaust fan")
test_overlap("fan", "fan")
test_overlap("return fan 1", "return fan")
test_overlap("condenser water pump", "pump")
