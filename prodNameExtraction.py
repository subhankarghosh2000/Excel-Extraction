import pandas as pd
import re

def clean_text(text):
    if not isinstance(text, str):
        return ""
    
    patterns_to_remove = [
        r'ETA[\s\-]*NO\.\s*[\w\-]+',
        r'BIS[\s\-]*NO\.\s*\w+',
        r'DT\.\d{2}\.\d{2}\.\d{2}',
        r'\d{2}\.\d{2}\.\d{2}',
        r'\bR-\d{8,}\b',
        r'CONTENT\s*[:\-]?\s*\d+\.?\d*%',
        r'AVERAGE\s+OF\s+CONTENT\s*\d+\.?\d*%',
        r'\d+\.?\d*%',
    ]
    for pattern in patterns_to_remove:
        text = re.sub(pattern, '', text, flags=re.IGNORECASE)
    
    return re.sub(r'\s+', ' ', text).strip()

def is_model_code(word):
    # Exclude patterns like A362-019-5017-00 or similar model codes
    if bool(re.match(r'^[A-Z0-9\-]{5,}$', word)) and any(char.isdigit() for char in word):
        return True
    # Exclude words with excessive hyphens or numbers
    if word.count('-') > 2 or sum(c.isdigit() for c in word) > 4:
        return True
    return False

def extract_product_name_description(text):
    if not isinstance(text, str) or not text.strip():
        return "", text

    text = clean_text(text)
    text = ' '.join(dict.fromkeys(text.split()))  # Remove exact repeated words

    # If starts with known prefix, keep it
    known_prefixes = [
        "GRID CONNECTED INVERTER",
        "SOLAR INVERTER",
        "GRID TIE SOLAR PV INVERTER",
        "UTILITY INTERCONNECTED PHOTOVOLTAIC INVERTERS",
        "Crystalline Silicon Terrestrial Photovoltaic PV Module",
    ]
    for prefix in known_prefixes:
        if text.upper().startswith(prefix.upper()):
            return prefix.strip(), text

    # Split into segments to avoid model/BIS/ETA pollution
    segments = re.split(r'\b(?:MODEL|ETA|BIS|NO\.|R-|CODE|MAX|INV\.|P\.LIST|BL)\b', text, flags=re.IGNORECASE)
    candidates = []
    for segment in segments:
        words = segment.strip().split()
        clean_words = [w for w in words if not is_model_code(w)]
        candidate = ' '.join(clean_words).strip()
        if 3 <= len(candidate.split()) <= 20:  # Ensure reasonable length
            candidates.append(candidate)

    if candidates:
        # Choose the candidate with the most relevant keywords or longest length
        best = max(candidates, key=lambda x: len(x.split()))
        return best.strip(), text

    return text.strip(), text  # fallback

# === Main Code ===
def process_file(input_path, output_path):
    # Read the Excel file in chunks
    chunks = pd.read_excel(input_path, chunksize=1000)  # Process 1000 rows at a time
    processed_chunks = []

    desc_col = 'Product Description'

    for chunk in chunks:
        if desc_col not in chunk.columns:
            raise ValueError(f"Column '{desc_col}' not found in the uploaded file")

        # Add the Product Name column
        chunk['Product Name'] = chunk[desc_col].apply(lambda x: extract_product_name_description(x)[0])

        # Reorder columns to place Product Name before Product Description
        cols = list(chunk.columns)
        cols.insert(cols.index(desc_col), cols.pop(cols.index('Product Name')))
        chunk = chunk[cols]

        processed_chunks.append(chunk)

    # Concatenate all processed chunks and save to output
    processed_df = pd.concat(processed_chunks)
    processed_df.to_excel(output_path, index=False)