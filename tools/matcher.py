import difflib
import re
import unicodedata
from typing import Dict, List, Any

def normalize(text: str) -> str:
    """
    Cleans and standardizes an input string for accurate matching.
    """
    if not text:
        return ""
    
    # 1. Convert to lowercase
    text = text.lower()
    
    # 2. Normalize unicode characters (e.g., converting 'é' to 'e')
    text = unicodedata.normalize('NFKD', text).encode('ascii', 'ignore').decode('utf-8')
    
    # 3. Remove punctuation and special characters (keep alphanumeric and spaces)
    text = re.sub(r'[^\w\s]', ' ', text)
    
    # 4. Collapse multiple spaces into a single space and strip whitespace from ends
    text = re.sub(r'\s+', ' ', text).strip()
    
    return text

from typing import List, Set

def tokenize(text: str, method: str = "word", n: int = 3) -> List[str]:
    """
    Splits a normalized string into a list of tokens.
    Supported methods: 'word' or 'ngram'.
    """
    if not text:
        return []
    
    # Method 1: Standard word tokenization (split by whitespace)
    if method == "word":
        return text.split()
    
    # Method 2: Character N-grams (great for fuzzy matching and identifiers)
    elif method == "ngram":
        # Example: "sony" with n=3 -> ["son", "ony"]
        return [text[i:i+n] for i in range(len(text) - n + 1)]
    
    else:
        raise ValueError(f"Unknown tokenization method: {method}")
    
def extract_features(tokens: List[str], text_id: str, tfidf_model=None) -> Dict[str, Any]:
    """
    Transforms tokens into a structured dictionary of features for the similarity engine.
    """
    joined_text = " ".join(tokens)
    numbers = set(re.findall(r'\d+', joined_text))

    tfidf_vector = None
    if tfidf_model and joined_text:
        tfidf_vector = tfidf_model.transform([joined_text]).toarray()[0]

    return {
        "id": text_id,
        "raw_text": joined_text,
        "tokens": tokens,
        "token_set": set(tokens),
        "numerical_attributes": numbers,
        "vector": tfidf_vector
    }

def identifier_score(str1: str, str2: str) -> float:
    """Extracts and compares alphanumeric serial/model codes (e.g., 'XBR-65X900F')."""
    # Find words containing both letters and numbers
    ids1 = set(re.findall(r'\b(?=\w*\d)(?=\w*[a-zA-Z])\w+\b', str1))
    ids2 = set(re.findall(r'\b(?=\w*\d)(?=\w*[a-zA-Z])\w+\b', str2))
    
    if not ids1 or not ids2:
        return 0.5  # Neutral score if no specific identifiers exist
    
    # Jaccard similarity between identifier sets
    intersection = ids1.intersection(ids2)
    union = ids1.union(ids2)
    return len(intersection) / len(union)


def attribute_score(tokens1: list, tokens2: list) -> float:
    """Measures overlapping descriptive tokens (colors, specs) using Jaccard Similarity."""
    set1, set2 = set(tokens1), set(tokens2)
    if not set1 or not set2:
        return 0.0
    return len(set1.intersection(set2)) / len(set1.union(set2))


def fuzzy_score(str1: str, str2: str) -> float:
    """Handles typos and structural shifts using SequenceMatcher (Gestalt Error Rate)."""
    return difflib.SequenceMatcher(None, str1, str2).ratio()


def price_score(p1: float, p2: float) -> float:
    """Compares numeric values, like prices or weights. High variance = low score."""
    if p1 <= 0 or p2 <= 0:
        return 0.5  # Neutral if price is missing
    # Ratio of min to max price
    return min(p1, p2) / max(p1, p2)


# --- The Main Engine ---

def similarity_engine(str_a: str, str_b: str) -> float:
    """
    Combines sub-scores using weights to determine if two items are a match.
    """

    str_a = normalize(str_a)
    features_a = extract_features(tokenize(str_a), text_id="string_a")
    str_b = normalize(str_b)
    features_b = extract_features(tokenize(str_b), text_id="string_b")

    # 1. Calculate individual scores
    f_score = fuzzy_score(features_a["raw_text"], features_b["raw_text"])
    a_score = attribute_score(features_a["tokens"], features_b["tokens"])
    i_score = identifier_score(features_a["raw_text"], features_b["raw_text"])
    
    # 2. Apply weights (Identifiers and price matching are heavily weighted)
    weights = {
        "fuzzy": 0.2,
        "attribute": 0.2,
        "identifier": 0.6,
    }
    
    # If a critical identifier mismatch happens, we can force a lower score
    if i_score == 0.0:
        return 0.1  # Fail-safe: Model numbers don't match, they aren't the same item!

    final_score = (
        (f_score * weights["fuzzy"]) +
        (a_score * weights["attribute"]) +
        (i_score * weights["identifier"])
    )

    print(f"String 1 Parsed Rate: ${features_a['tokens']}")
    print(f"String 2 Parsed Rate: ${features_b['tokens']}")
   
    return round(final_score, 4)
