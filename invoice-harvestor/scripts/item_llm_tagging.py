import json
import pandas as pd
from google import genai
import time
from pathlib import Path
from google.api_core import exceptions

from config import (
    DEFAULT_INVOICE_ITEMS_CSV,
    DEFAULT_TAG_CACHE_FILE,
    DEFAULT_TAGGED_EXPENSES_CSV,
    load_project_env,
)

INPUT_CSV = DEFAULT_INVOICE_ITEMS_CSV
OUTPUT_CSV = DEFAULT_TAGGED_EXPENSES_CSV
TAG_CACHE_FILE = DEFAULT_TAG_CACHE_FILE

load_project_env()


def normalize_description(description) -> str:
    if pd.isna(description):
        return ""
    return " ".join(str(description).lower().split())


def load_tag_cache(cache_path: Path) -> dict[str, str]:
    if not cache_path.exists():
        return {}

    try:
        cache = json.loads(cache_path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError) as e:
        print(f"Could not load tag cache from {cache_path}: {e}")
        return {}

    if not isinstance(cache, dict):
        print(f"Ignoring tag cache at {cache_path}: expected a JSON object.")
        return {}

    return {
        str(description): str(tag).strip().lower()
        for description, tag in cache.items()
        if str(description).strip() and str(tag).strip()
    }


def save_tag_cache(cache_path: Path, cache: dict[str, str]) -> None:
    cache_path.parent.mkdir(parents=True, exist_ok=True)
    cache_path.write_text(
        json.dumps(dict(sorted(cache.items())), indent=2) + "\n",
        encoding="utf-8",
    )


def get_tag_from_llm(description):

    client = genai.Client()

    """Calls the LLM to classify the description into a specific tag."""

    # Define your allowed categories here
    prompt = f"""
    You are a helpful categorization assistant. 
    Categorize the following item description into EXACTLY ONE of these tags: 
    [vegetables, fruits, masala, yogurt (or curd), milk, paneer, utility, snacks, groceries, dining, aata, rice, unknown].
    
    If it doesn't clearly fit, choose the closest one or 'unknown'.
    Respond ONLY with the single word tag, nothing else, no markdown.
    
    Description: "{description}"
    Tag:
    """
    max_retries = 5
    retry_delay = 15  # Base wait time in seconds if a 429 occurs

    for attempt in range(max_retries):
        try:
            response = client.models.generate_content(
                model="gemini-3.1-flash-lite",
                contents=prompt,
            )
            tag = response.text.strip().lower()
            return tag

        except exceptions.ResourceExhausted as e:
            # This catches the 429 error explicitly
            print(
                f"\n[Rate Limit hit!] 429 Resource Exhausted on attempt {attempt + 1}/{max_retries}."
            )
            print(f"Sleeping for {retry_delay} seconds before trying again...")
            time.sleep(retry_delay)
            retry_delay *= 2

        except Exception as e:
            print(f"Error processing '{description}': {e}")
            return "unknown"
    print(
        f"Failed to process '{description}' after {max_retries} attempts due to rate limits."
    )
    return "unknown"


def tag_expenses_core(
    input_csv: str | Path = INPUT_CSV,
    output_csv: str | Path = OUTPUT_CSV,
    cache_file: str | Path = TAG_CACHE_FILE,
) -> Path:
    input_path = Path(input_csv)
    output_path = Path(output_csv)
    cache_path = Path(cache_file)

    print(f"Loading '{input_path}' for processing...")
    df = pd.read_csv(input_path)
    # Ensure the description column exists
    if "description" not in df.columns:
        raise ValueError("CSV must contain a 'description' column.")

    description_tag_map = load_tag_cache(cache_path)
    cache_updated = False
    print(
        f"Starting LLM tagging process with {len(description_tag_map)} cached description(s)..."
    )

    tags = []
    for index, row in df.iterrows():
        description = row["description"]

        # Skip empty descriptions
        if pd.isna(description):
            tags.append("unknown")
            continue

        description_key = normalize_description(description)
        if description_key in description_tag_map:
            tag = description_tag_map[description_key]
            print(f"Row {index + 1}: '{description}' -> [{tag}] cached")
        else:
            tag = get_tag_from_llm(description)
            description_tag_map[description_key] = tag
            cache_updated = True
            print(f"Row {index + 1}: '{description}' -> [{tag}]")
            # Add a tiny sleep to avoid hitting API rate limits if you have hundreds of rows
            time.sleep(7)

        tags.append(tag)

    # Add the new tags to the dataframe
    df["tag"] = tags

    # Save to a new CSV
    output_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(output_path, index=False)
    if cache_updated:
        save_tag_cache(cache_path, description_tag_map)
        print(f"Updated tag cache at {cache_path}")
    print(f"\nSuccess! Tagged data saved to {output_path}")
    return output_path


def main():
    tag_expenses_core()


if __name__ == "__main__":
    main()
