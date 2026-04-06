import google.generativeai as genai
from google.oauth2.service_account import Credentials
import gspread
# 🛠️ CORRECTED IMPORTS: Removed set_text_format, set_data_validation, get_default_format
from gspread_formatting import set_row_height, cellFormat, format_cell_range 
from datetime import datetime
import json
import os
import random
from typing import Any
from dotenv import load_dotenv
import requests
from bs4 import BeautifulSoup

# Load environment variables
load_dotenv()

# Configuration
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
# Ensure GEMINI_MODEL is set or defaults to a good one
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-2.5-flash") 
GOOGLE_SHEETS_CREDENTIALS_FILE = os.getenv("GOOGLE_SHEETS_CREDENTIALS_FILE", "credentials.json")
SPREADSHEET_NAME = os.getenv("SPREADSHEET_NAME", "Daily Vocabulary")
# Update column index since 'Starting Date' is removed and 'Date Added' is now column 1 ('A')
FINNISH_WORD_COLUMN_INDEX = 2 # 'B' column for 'Finnish Word' in 1-based index
VOCAB_COUNT=int(os.getenv("VOCAB_COUNT", 20)) # Default to 20 words if not set
VOCAB_SOURCE=os.getenv("VOCAB_SOURCE", "common1000").strip().lower()  # common1000 | gemini
COMMON_WORDS_URL=os.getenv(
    "COMMON_WORDS_URL",
    "https://1000mostcommonwords.com/1000-most-common-finnish-words/",
).strip()
COMMON_WORDS_CACHE_FILE=os.getenv("COMMON_WORDS_CACHE_FILE", "finnish_common_1000.json").strip()
ENRICH_BATCH_SIZE=int(os.getenv("ENRICH_BATCH_SIZE", 25))


def _strip_markdown_code_fences(text: str) -> str:
    text = text.strip()
    if text.startswith("```"):
        lines = text.splitlines()
        # Remove opening fence line: ``` or ```json
        if lines:
            lines = lines[1:]
        # Remove closing fence line if present
        if lines and lines[-1].strip().startswith("```"):
            lines = lines[:-1]
        return "\n".join(lines).strip()
    return text


def _json_loads_loose(text: str) -> Any:
    """
    Parse JSON from model output that may include code fences or extra text.
    """
    cleaned = _strip_markdown_code_fences(text)
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        # Fallback: try to find the first JSON array/object region.
        first_brace = cleaned.find("{")
        first_bracket = cleaned.find("[")
        if first_bracket == -1 and first_brace == -1:
            raise
        if first_bracket != -1 and (first_brace == -1 or first_bracket < first_brace):
            start = first_bracket
            end = cleaned.rfind("]")
        else:
            start = first_brace
            end = cleaned.rfind("}")
        if start == -1 or end == -1 or end <= start:
            raise
        return json.loads(cleaned[start : end + 1])


def _load_cached_common_words(cache_file: str) -> tuple[datetime | None, list[dict[str, str]] | None]:
    if not os.path.exists(cache_file):
        return None, None
    try:
        with open(cache_file, "r", encoding="utf-8") as f:
            payload = json.load(f)
        fetched_at_raw = payload.get("fetched_at")
        words = payload.get("words")
        if not isinstance(words, list):
            return None, None
        fetched_at = None
        if isinstance(fetched_at_raw, str):
            try:
                fetched_at = datetime.fromisoformat(fetched_at_raw.replace("Z", "+00:00"))
            except ValueError:
                fetched_at = None
        # Only keep well-formed items
        cleaned: list[dict[str, str]] = []
        for item in words:
            if not isinstance(item, dict):
                continue
            fi = str(item.get("finnish_word", "")).strip()
            en = str(item.get("english_translation", "")).strip()
            if fi and en:
                cleaned.append({"finnish_word": fi, "english_translation": en})
        return fetched_at, cleaned or None
    except Exception:
        return None, None


def scrape_common_finnish_words(url: str) -> list[dict[str, str]]:
    """
    Scrape the "1000 most common Finnish words" table.
    Returns list of objects: {finnish_word, english_translation}
    """
    resp = requests.get(
        url,
        timeout=30,
        headers={
            "User-Agent": (
                "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/120.0.0.0 Safari/537.36"
            )
        },
    )
    resp.raise_for_status()

    soup = BeautifulSoup(resp.text, "html.parser")
    tables = soup.find_all("table")
    target_table = None

    for table in tables:
        first_row = table.find("tr")
        if not first_row:
            continue
        cells = [c.get_text(" ", strip=True).lower() for c in first_row.find_all(["td", "th"])]
        if len(cells) >= 3 and "finnish" in cells[1] and "english" in cells[2]:
            target_table = table
            break

    if target_table is None:
        raise RuntimeError("Could not find Finnish/English word table on the page.")

    words: list[dict[str, str]] = []
    seen: set[str] = set()

    for row in target_table.find_all("tr")[1:]:
        tds = row.find_all("td")
        if len(tds) < 3:
            continue
        finnish_word = tds[1].get_text(" ", strip=True)
        english_translation = tds[2].get_text(" ", strip=True)
        key = finnish_word.strip().lower()
        if not key or key in seen:
            continue
        seen.add(key)
        words.append(
            {
                "finnish_word": finnish_word.strip(),
                "english_translation": english_translation.strip(),
            }
        )

    if not words:
        raise RuntimeError("Scraped word list was empty.")

    return words


def get_common_finnish_words(
    *,
    url: str = COMMON_WORDS_URL,
    cache_file: str = COMMON_WORDS_CACHE_FILE,
) -> list[dict[str, str]]:
    """
    Load the 1000-most-common Finnish words list, using a local cache file.
    This scrapes ONLY ONCE: if the cache file exists, it is always used.
    Delete the cache file to force a re-scrape.
    """
    _, cached_words = _load_cached_common_words(cache_file)
    if cached_words:
        return cached_words

    try:
        print(f"Fetching common Finnish words from: {url}")
        words = scrape_common_finnish_words(url)
        payload = {
            "fetched_at": datetime.utcnow().isoformat(timespec="seconds") + "Z",
            "url": url,
            "words": words,
        }
        with open(cache_file, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
        print(f"Cached {len(words)} words to '{cache_file}'.")
        return words
    except Exception as e:
        print(f"❌ Failed to fetch common words list from {url}: {e}")
        raise

def setup_gemini():
    """Initialize Gemini API"""
    genai.configure(api_key=GEMINI_API_KEY)
    model = genai.GenerativeModel(GEMINI_MODEL)
    return model

def setup_google_sheets():
    """
    Initialize Google Sheets connection and set up the sheet.
    This version ensures headers are always present and correct.
    """
    # 🆕 UPDATED HEADERS: Added 'Video Caption'
    headers = [
        'Date Added', 'Finnish Word', 'English Translation', 
        'Category', 'Level', 'Example Sentence', 'Video Prompt', 'Video Caption'
    ]
    scope = [
        'https://spreadsheets.google.com/feeds',
        'https://www.googleapis.com/auth/drive'
    ]
    
    # Load credentials for service account
    creds = Credentials.from_service_account_file(
        GOOGLE_SHEETS_CREDENTIALS_FILE, 
        scopes=scope
    )
    client = gspread.authorize(creds)
    
    # Open or create spreadsheet
    try:
        # 1. Sheet exists: open it
        spreadsheet = client.open(SPREADSHEET_NAME)
        sheet = spreadsheet.sheet1
        print(f"Spreadsheet '{SPREADSHEET_NAME}' opened.")
        
    except gspread.SpreadsheetNotFound:
        # 2. Sheet does not exist: create it
        print(f"Spreadsheet '{SPREADSHEET_NAME}' not found. Creating a new one.")
        spreadsheet = client.create(SPREADSHEET_NAME)
        sheet = spreadsheet.sheet1
        
    
    # --- Check and set headers in the first row (A1 to H1) ---
    try:
        # Read the current first row's values
        current_headers = sheet.row_values(1)
    except IndexError:
        # Handles a completely empty sheet
        current_headers = []
        
    if current_headers != headers:
        print("Adding/Correcting headers in the first row.")
        
        # Use gspread's batch update to set headers in row 1
        cell_list = sheet.range(f'A1:{gspread.utils.rowcol_to_a1(1, len(headers))}')
        for i, val in enumerate(headers):
            cell_list[i].value = val
        sheet.update_cells(cell_list, value_input_option='USER_ENTERED')
        
        if len(current_headers) == 0:
              print("Spreadsheet created with correct headers.")
              
    # -----------------------------------------------------------
    
    return sheet

def get_existing_words(sheet):
    """Retrieves all existing Finnish words from the sheet to prevent duplicates."""
    try:
        # Get all values from the 'Finnish Word' column (column B, index 2)
        column_data = sheet.col_values(FINNISH_WORD_COLUMN_INDEX)
        # Skip the header row and return the set of existing words for fast lookups
        return set(word.strip().lower() for word in column_data[1:])
    except Exception as e:
        print(f"Could not retrieve existing words: {e}")
        # Return an empty set if an error occurs
        return set() 


def pick_new_words_from_common_list(
    common_words: list[dict[str, str]], existing_words: set[str], count: int
) -> list[dict[str, str]]:
    candidates: list[dict[str, str]] = []
    for item in common_words:
        finnish_word = str(item.get("finnish_word", "")).strip()
        english_translation = str(item.get("english_translation", "")).strip()
        if not finnish_word or not english_translation:
            continue
        if finnish_word.lower() in existing_words:
            continue
        candidates.append({"finnish_word": finnish_word, "english_translation": english_translation})

    if len(candidates) < count:
        raise RuntimeError(
            f"Not enough new words left in the common-word list. "
            f"Need {count}, but only found {len(candidates)} that are not in the sheet."
        )

    selected = random.sample(candidates, count)
    for item in selected:
        existing_words.add(item["finnish_word"].strip().lower())
    return selected


def _normalize_level(level_raw: str) -> str:
    val = (level_raw or "").strip().upper()
    if val in {"A1", "A2", "B1"}:
        return val
    for candidate in ("A1", "A2", "B1"):
        if candidate in val:
            return candidate
    return ""


def enrich_vocabulary_details(model, words: list[dict[str, str]]) -> list[dict[str, str]]:
    """
    Given [{finnish_word, english_translation}], ask Gemini to add:
    category, level, example_finnish, example_english.
    Runs in batches to avoid huge prompts.
    """
    enriched: list[dict[str, str]] = []

    for i in range(0, len(words), max(1, ENRICH_BATCH_SIZE)):
        chunk = words[i : i + max(1, ENRICH_BATCH_SIZE)]
        input_json = json.dumps(chunk, ensure_ascii=False)

        prompt = f"""
You are a Finnish teacher. For each vocabulary item provided, add learning metadata.

INPUT (JSON array):
{input_json}

For EACH item, output a JSON array of objects with EXACTLY these keys:
- finnish_word
- english_translation
- category (noun, verb, adjective, adverb, pronoun, preposition, conjunction, phrase, abbreviation, etc.)
- level (A1, A2, or B1 only)
- example_finnish (one short, natural Finnish sentence using the word; the word may be inflected if needed)
- example_english (English translation of the example sentence)

Rules:
- Keep finnish_word and english_translation exactly as in the input.
- No Markdown, no commentary. Output ONLY the raw JSON array.
""".strip()

        response = model.generate_content(prompt)
        data = _json_loads_loose(getattr(response, "text", "") or "")
        if not isinstance(data, list):
            raise RuntimeError("Gemini enrichment did not return a JSON array.")

        by_word: dict[str, dict[str, Any]] = {}
        for item in data:
            if not isinstance(item, dict):
                continue
            fi = str(item.get("finnish_word", "")).strip()
            if not fi:
                continue
            by_word[fi.lower()] = item

        for original in chunk:
            fi_key = original["finnish_word"].strip().lower()
            item = by_word.get(fi_key, {})
            enriched_item = {
                "finnish_word": original["finnish_word"],
                "english_translation": original["english_translation"],
                "category": str(item.get("category", "")).strip(),
                "level": _normalize_level(str(item.get("level", ""))),
                "example_finnish": str(item.get("example_finnish", "")).strip(),
                "example_english": str(item.get("example_english", "")).strip(),
            }
            enriched.append(enriched_item)

    return enriched


def generate_finnish_vocabulary_from_common_words(model, existing_words, count=10):
    """
    Select random words from the "1000 most common Finnish words" list,
    skipping words already in the Google Sheet, then enrich with Gemini.
    """
    common_words = get_common_finnish_words()
    selected = pick_new_words_from_common_list(common_words, existing_words, count)
    return enrich_vocabulary_details(model, selected)


def generate_finnish_vocabulary_with_gemini(model, existing_words, count=10, max_attempts=8):
    """
    Generate random Finnish words using Gemini, filtering out existing ones.
    Tries up to max_attempts to find the requested number of unique words.
    """
    new_vocabulary = []
    attempts = 0
    
    # Keep iterating until we have exactly the requested number of words
    while len(new_vocabulary) < count and attempts < max_attempts:
        needed = count - len(new_vocabulary)
        attempts += 1
        print(f"Generating {needed} more words (Attempt {attempts}, have {len(new_vocabulary)}/{count})...")

        if existing_words:
            # Convert set to list and take a sample to avoid huge prompts, 
            # but enough to give the model a hint of what's there.
            # Taking last 50 might be better than random 5 if we want to avoid recent ones,
            # but random sample is okay. Let's just show a few.
            existing_sample = list(existing_words)[-20:] if len(existing_words) > 20 else list(existing_words)
            exclude_hint = f" Do NOT use any of these words: {', '.join(existing_sample)}."
        else:
            exclude_hint = ""

        # UPDATED PROMPT: Requesting Level (A1-B1)
        prompt = f"""Generate {needed} random Finnish vocabulary words suitable for daily learning between levels A1 to B1.
        {exclude_hint}
        For each word, provide:
        1. The Finnish word
        2. English translation
        3. Category (e.g., noun, verb, adjective, daily life, food, nature, etc.)
        4. **Level (A1, A2, or B1)**
        5. A simple example sentence in Finnish with English translation
        
        Format the response as JSON array with objects containing: 
        finnish_word, english_translation, category, **level**, example_finnish, example_english
        
        Make the words varied and useful for beginners to intermediate learners."""
        
        try:
            response = model.generate_content(prompt)
            batch_vocabulary = _json_loads_loose(getattr(response, "text", "") or "")
            
            if not isinstance(batch_vocabulary, list):
                print("Model did not return a list. Retrying...")
                continue

            # POST-PROCESSING: Filter out duplicates and ensure data structure
            for item in batch_vocabulary:
                if 'finnish_word' not in item:
                    continue
                    
                finnish_word = item['finnish_word'].strip().lower()
                
                # Check against global existing words AND words generated in this session
                if finnish_word not in existing_words and not any(w['finnish_word'].lower() == finnish_word for w in new_vocabulary):
                    new_vocabulary.append(item)
                    existing_words.add(finnish_word) # Add to set to prevent duplicates in next loop iteration
                else:
                    print(f" ⚠️ Skipping duplicate word: {item['finnish_word']}")
        
        except Exception as e:
            print(f"Error during generation attempt {attempts}: {e}")

    if len(new_vocabulary) < count:
        raise RuntimeError(
            f"Only generated {len(new_vocabulary)} unique words out of {count} "
            f"after {max_attempts} attempts."
        )

    return new_vocabulary


def generate_finnish_vocabulary(model, existing_words, count=10):
    if VOCAB_SOURCE == "gemini":
        return generate_finnish_vocabulary_with_gemini(model, existing_words, count=count)
    if VOCAB_SOURCE == "common1000":
        return generate_finnish_vocabulary_from_common_words(model, existing_words, count=count)
    raise ValueError(f"Unknown VOCAB_SOURCE='{VOCAB_SOURCE}'. Use 'common1000' or 'gemini'.")

def generate_video_prompt(model, word_data):
    """Generate a video generation prompt for the vocabulary word"""
    prompt = f"""
    You are a creative TikTok scriptwriter. Your task is to generate a video prompt for Finnish word of the day. 
    The word is: {word_data['finnish_word']} which means "{word_data['english_translation']}". 
    
    IMPORTANT - Choose the BEST illustration approach to maximize visual impact:
    
    **STEP 1: Evaluate if the word can be illustrated visually**
    Ask yourself: Can this word be shown through a clear, engaging visual image or action?
    
    Examples that work well visually:
    - Concrete objects: apple, car, house, book
    - Actions/verbs: jump, run, swim, dance, eat, sleep
    - Visual states: happy (smiling face), tired (yawning), cold (shivering), hot (sweating)
    - Visual adjectives: big, small, colorful, clean, dirty
    - Places: park, library, kitchen, beach
    
    **STEP 2: Choose your approach**
    
    ✅ **USE VISUAL DESCRIPTION** if the word can be effectively shown through visuals:
       - Create a detailed visual description showing the object/action/state
       - Focus on visual elements: colors, textures, setting, environment, body language, expressions
       - A character can perform the action or demonstrate the state while saying the word
       - Keep it simple, dynamic, and visually focused
       - Example: For "jump" → show a character mid-jump with dynamic motion
       - Example: For "happy" → show a character with a bright smile and joyful expression
    
    ❌ **USE SCENE/CONVERSATION** only if the word needs context to be understood:
       - Use this for complex emotions (nostalgia, anxiety), abstract concepts (possibility, freedom), or social situations
       - Create a short scene or conversation that illustrates the meaning through context
       - Use characters in a daily life situation that demonstrates the concept
       - The dialogue should naturally include the word in context
       - Make the scene relatable and easy to understand
    
    General Guidelines:
    - Characters speak clearly in Finnish and grammatically correct
    - The scene should be about common daily life situations and easy to illustrate
    - Make any conversation sound natural
    - The main word only needs to appear once in the speech, no need to repeat it
    - The scene should begin right in the first second to get the audience attention
    - Strictly no text or subtitles included in the video
    - **Prioritize visual descriptions whenever possible** - they create better, more engaging content!
    
    **YOUR RESPONSE MUST START WITH THE ILLUSTRATION STYLE DESCRIPTION**, followed by the scene details and audio.
    
    REQUIRED FORMAT:
    
    **Illustration Style:**
    Use a warm, modern flat-vector illustration style with soft pastel colors, clean lines, and simple but expressive facial features. Think of a style that could be used in educational flashcards or language-learning apps—playful yet clear, conveying both the action and the meaning.
    
    **Scene (8 seconds):**
    [Describe the visual scene here in detail]
    
    **Audio:**
    [Describe what characters say in Finnish with Helsinki region accent, matching their actions]
    **IMPORTANT: Do NOT include any background music. Only include contextual sound effects if relevant to the scene (e.g., a doorbell, phone ringing, kitchen sounds, birds chirping).**
    """
    
    response = model.generate_content(prompt)
    return response.text.strip()

def check_and_fix_finnish_speech(model, video_prompt, word_data, max_iterations=3):
    """Check and fix Finnish grammar and naturalness in the video prompt.
    
    Args:
        model: Gemini model instance
        video_prompt: The generated video prompt containing Finnish speech
        word_data: Dictionary with word information
        max_iterations: Maximum number of fix attempts
        
    Returns:
        Corrected video prompt
    """
    finnish_word = word_data['finnish_word']
    english_translation = word_data['english_translation']
    
    print(f"    Checking Finnish grammar and naturalness...")
    
    for iteration in range(max_iterations):
        check_prompt = f"""
        You are a Finnish language expert specializing in natural, conversational Finnish. 
        
        Analyze the following video script for the Finnish word "{finnish_word}" (meaning "{english_translation}").
        
        VIDEO SCRIPT:
        {video_prompt}
        
        **CRITICAL INSTRUCTION:** Make MINIMAL changes. Only fix actual grammatical errors. Preserve all original text, structure, dialogue, and content.
        
        Check for:
        1. **Grammar errors** (case endings, verb conjugations, word order, etc.)
        2. **Naturalness** (Does it sound like how a native Finnish speaker would actually talk in daily conversation?)
        3. **Appropriate use of the target word** in context
        4. **Helsinki region accent compatibility** (avoid overly formal or archaic expressions)
        5. **No direct literal translations from English** (e.g., "what a beautiful house" should not be translated directly as "Mikä kaunis talo", but rather using an authentic Finnish expression).
        
        **When providing corrections:**
        - Keep the ENTIRE original text structure (illustration style, scene description, audio sections)
        - Only change the specific words/phrases that have grammatical errors
        - Do NOT rewrite or rephrase content that is already correct
        - Do NOT change the meaning, tone, or creative elements
        - Preserve character names, actions, and scene descriptions exactly
        
        Respond in JSON format with:
        {{
            "is_correct": true/false,
            "issues": ["list of specific issues found, if any"],
            "corrected_script": "the FULL script with ONLY grammatical errors fixed (only if is_correct is false)",
            "explanation": "brief explanation of what specific words/phrases were fixed (only if corrections were made)"
        }}
        
        If the Finnish is already perfect, set is_correct to true and leave corrected_script empty.
        """
        
        try:
            response = model.generate_content(check_prompt)
            text = response.text
            
            if "```json" in text:
                text = text.split("```json")[1].split("```")[0]
            elif "```" in text:
                text = text.split("```")[1].split("```")[0]
            
            result = json.loads(text.strip())
            
            if result.get('is_correct', False):
                print(f"    ✅ Finnish speech verified (attempt {iteration + 1}/{max_iterations})")
                return video_prompt
            else:
                issues = result.get('issues', [])
                print(f"    ⚠️ Issues found (attempt {iteration + 1}/{max_iterations}):")
                for issue in issues:
                    print(f"       - {issue}")
                
                corrected_script = result.get('corrected_script', '')
                if corrected_script:
                    video_prompt = corrected_script
                    explanation = result.get('explanation', '')
                    if explanation:
                        print(f"    🔧 Fixed: {explanation}")
                else:
                    print(f"    ⚠️ No correction provided, using original script")
                    break
                    
        except Exception as e:
            print(f"    ❌ Error during grammar check (attempt {iteration + 1}/{max_iterations}): {e}")
            return video_prompt
    
    print(f"    ✅ Finnish speech finalized after {max_iterations} checks")
    return video_prompt

def generate_video_caption(model, word_data):
    """Generate an engaging TikTok caption for the vocabulary word."""
    finnish_word = word_data['finnish_word']
    english_translation = word_data['english_translation']
    level = str(word_data.get('level') or "B1").strip().upper()
    if level not in {"A1", "A2", "B1"}:
        level = "B1"
    
    prompt = f"""
    You are a TikTok content strategist. 
    Create a short, engaging TikTok caption for a video teaching the Finnish word **{finnish_word}** (meaning "{english_translation}"). 
    The target audience is A1-{level} Finnish learners.

    Do not use JSON formatting. Just provide the raw caption text.
    
    Below is the example as template, put the whole post with proper format in a "caption".
    Make sure the Quick Tip feels fresh and varied each day (sometimes grammar endings, sometimes synonyms, related words, cultural notes, fun facts, etc.).

    ✨ Finnish Word of the Day ✨

    📖 **kirjasto** (noun) → library

    💬 Example:
    [fi]: "Mennään tänään kirjastoon."
    [en]: “Let’s go to the library today.”

    🔎 Quick Tip
    Finnish place endings change the meaning:
    - kirjastossa = in the library
    - kirjastoon = to the library
    - kirjastosta = from the library

    🎭 What’s the best thing you’ve borrowed from a library? 📚

    📌 #FinnishWordOfTheDay #LearnFinnish #suomi #suomenkieli  #LanguageTok #FinnishLanguage #kirjasto #LanguageLearning
    """
    
    response = model.generate_content(prompt)
    return response.text.strip()


def save_to_sheets(sheet, vocabulary_data):
    """Save vocabulary data to Google Sheets, using a single 'Date Added' column."""
    current_date = datetime.now().strftime("%Y-%m-%d")
    
    rows = []
    for item in vocabulary_data:
        example_sentence = f"{item.get('example_finnish', '')} ({item.get('example_english', '')})"
        
        row = [
            current_date, 
            item.get('finnish_word', ''),
            item.get('english_translation', ''),
            item.get('category', ''),
            item.get('level', ''), 
            example_sentence,
            item.get('video_prompt', ''),
            item.get('video_caption', '')
        ]
        rows.append(row)
    
    if rows:
        sheet.append_rows(rows)
        print(f"Successfully added {len(rows)} **new** vocabulary words to Google Sheets")
    else:
        print("No new vocabulary words to add.")

def apply_fixed_row_height(sheet, pixel_size=50):
    """
    Sets a fixed row height for all data rows (starting from row 2) 
    and ensures text wrapping is disabled.
    """
    
    # Set fixed row height (e.g., 50 pixels) for all rows from row 2 onwards
    print(f"Applying fixed row height of {pixel_size} pixels...")
    
    # Set row height for rows 2 to the current last row
    
    # Get the total number of rows currently in the sheet
    max_rows = sheet.row_count
def main():
    """Main function to run the vocabulary generator"""
    
    BACKUP_FILE = "backup_vocab.json"

    # --- 1. Check for Backup ---
    if os.path.exists(BACKUP_FILE):
        print(f"⚠️ Found backup file '{BACKUP_FILE}'. Resuming from previous failed run...")
        try:
            with open(BACKUP_FILE, "r", encoding="utf-8") as f:
                vocabulary = json.load(f)
            
            if vocabulary:
                print(f"Loaded {len(vocabulary)} words from backup.")
                print("Connecting to Google Sheets...")
                sheet = setup_google_sheets()
                
                print("Saving to Google Sheets...")
                save_to_sheets(sheet, vocabulary)
                apply_fixed_row_height(sheet, pixel_size=50)
                
                # If successful, remove backup
                os.remove(BACKUP_FILE)
                print(f"✅ Backup file '{BACKUP_FILE}' deleted after successful save.")
                print("\n🎉 **Complete!** Vocabulary restored from backup and saved.")
                return
            else:
                print("Backup file was empty. Proceeding with normal generation.")
        except Exception as e:
            print(f"Error reading backup file: {e}. Proceeding with normal generation.")

    # --- 2. Normal Generation Flow ---
    if not GEMINI_API_KEY:
        print("Error: GEMINI_API_KEY is not set in environment variables/dotenv file.")
        return
        
    print("Initializing Gemini...")
    model = setup_gemini()
    
    print("Connecting to Google Sheets...")
    sheet = setup_google_sheets()
    
    print("Checking for existing vocabulary words to prevent duplicates...")
    existing_words = get_existing_words(sheet)
    print(f"Found {len(existing_words)} existing words in the sheet.")
    
    print(f"Generating {VOCAB_COUNT} Finnish vocabulary words...")
    vocabulary = generate_finnish_vocabulary(model, existing_words, count=VOCAB_COUNT) 
    
    if not vocabulary:
        print("No new unique words were generated. Exiting.")
        return
        
    print(f"Generated {len(vocabulary)} unique words.")
    
    print("Generating video prompts and captions...") 
    for item in vocabulary:
        word = item.get('finnish_word', 'Unknown Word')
        print(f"  Processing prompt for: {word}")
        
        # 1. Generate Video Prompt
        video_prompt = generate_video_prompt(model, item)
        
        # 2. Check and Fix Finnish Grammar/Naturalness
        print(f"  Checking Finnish speech for: {word}")
        video_prompt = check_and_fix_finnish_speech(model, video_prompt, item)
        item['video_prompt'] = video_prompt
        
        # 3. Generate Video Caption
        print(f"  Processing caption for: {word}")
        video_caption = generate_video_caption(model, item)
        item['video_caption'] = video_caption 
        
    # --- 3. Save Backup before Sheets ---
    print(f"Saving backup to '{BACKUP_FILE}' before uploading...")
    try:
        with open(BACKUP_FILE, "w", encoding="utf-8") as f:
            json.dump(vocabulary, f, indent=4, ensure_ascii=False)
    except Exception as e:
        print(f"⚠️ Warning: Could not save backup file: {e}")

    # --- 4. Save to Sheets ---
    print("Saving to Google Sheets...")
    try:
        save_to_sheets(sheet, vocabulary)
        apply_fixed_row_height(sheet, pixel_size=50)
        
        # If successful, remove backup
        if os.path.exists(BACKUP_FILE):
            os.remove(BACKUP_FILE)
            print(f"Backup file '{BACKUP_FILE}' deleted.")
            
        print("\n🎉 **Complete!** Your Finnish vocabulary has been saved to Google Sheets.")
        print(f"Generated {len(vocabulary)} new, unique words with video prompts and captions.")
        
    except Exception as e:
        print(f"\n❌ Error saving to Google Sheets: {e}")
        print(f"The generated data is saved in '{BACKUP_FILE}'. Fix the issue and run the script again to retry saving.")

if __name__ == "__main__":
    main()
