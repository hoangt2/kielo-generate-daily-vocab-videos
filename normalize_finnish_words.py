import google.generativeai as genai
import json
import os
from dotenv import load_dotenv
from typing import Any

load_dotenv()
genai.configure(api_key=os.getenv("GEMINI_API_KEY"))
model = genai.GenerativeModel(os.getenv("GEMINI_MODEL", "gemini-2.5-flash"))

def _strip_markdown_code_fences(text: str) -> str:
    text = text.strip()
    if text.startswith("```"):
        lines = text.splitlines()
        if lines:
            lines = lines[1:]
        if lines and lines[-1].strip().startswith("```"):
            lines = lines[:-1]
        return "\n".join(lines).strip()
    return text

def _json_loads_loose(text: str) -> Any:
    cleaned = _strip_markdown_code_fences(text)
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        first_bracket = cleaned.find("[")
        if first_bracket == -1:
            raise
        start = first_bracket
        end = cleaned.rfind("]")
        if start == -1 or end == -1 or end <= start:
            raise
        return json.loads(cleaned[start : end + 1])

def main():
    with open("finnish_common_1000.json", "r", encoding="utf-8") as f:
        data = json.load(f)
    
    words = data.get("words", [])
    batch_size = 100
    normalized_words = []
    
    for i in range(0, len(words), batch_size):
        batch = words[i:i+batch_size]
        print(f"Processing batch {i} to {i+len(batch)}...")
        input_json = json.dumps(batch, ensure_ascii=False)
        
        prompt = f"""
You are a Finnish language expert. I will provide a JSON array of objects with 'finnish_word' and 'english_translation'.
The 'finnish_word' may be in an inflected form (e.g., 'hänen', 'kohti', 'olivat', 'tekee', 'meille', 'kissan', jne).
Your task is to:
1. Convert the 'finnish_word' into its basic, dictionary form (nominative singular for nouns/adjectives, first infinitive/A-infinitive for verbs, nominative for pronouns).
   - If the word is already in basic form, keep it as is.
   - For example: 'hänen' -> 'hän', 'kohti' -> 'kohti' (adposition, unchanging), 'olivat' -> 'olla', 'tekee' -> 'tehdä', 'meille' -> 'me', 'kaksi' -> 'kaksi', 'puhtaita' -> 'puhdas', 'lapset' -> 'lapsi'.
2. Update the 'english_translation' if necessary so it accurately matches the basic form.
   - For example: 'hänen' (his) -> 'hän' (he/she), 'olivat' (were) -> 'olla' (be), 'puhtaita' (clean ones) -> 'puhdas' (clean).

Return EXACTLY a JSON array of objects with 'finnish_word' and 'english_translation' keys, in the exact same order as the input. No markdown fences around the response, just the JSON array.

INPUT:
{input_json}
"""
        try:
            response = model.generate_content(prompt)
            result = _json_loads_loose(getattr(response, "text", "") or "")
            if len(result) == len(batch):
                normalized_words.extend(result)
            else:
                print(f"Warning: batch size mismatch in return. Expected {len(batch)}, got {len(result)}")
                normalized_words.extend(result)
        except Exception as e:
            print(f"Failed to parse batch {i}: {e}")
            if hasattr(response, "text"):
                print("Raw response:", response.text)
            normalized_words.extend(batch)

    # Remove duplicates
    seen = set()
    unique_words = []
    for item in normalized_words:
        fw = str(item.get("finnish_word", "")).strip().lower()
        if not fw or fw in seen:
            continue
        seen.add(fw)
        item["finnish_word"] = fw
        unique_words.append(item)
    
    data["words"] = unique_words
    
    with open("finnish_common_1000.json", "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
        
    print(f"Done. Original count: {len(words)}. New unique count: {len(unique_words)}")

if __name__ == "__main__":
    main()
