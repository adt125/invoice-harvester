import pandas as pd
from google import genai
import time
from pathlib import Path
from dotenv import find_dotenv, load_dotenv
from google.api_core import exceptions
import shutil

SCRIPT_DIR = Path(__file__).resolve().parent  # Points to 'scripts/'
ROOT = SCRIPT_DIR.parent
ASSETS_DIR = ROOT / "assets"
TEMP_DIR = ASSETS_DIR / "temp"
INPUT_CSV = TEMP_DIR / "invoice_items.csv"
OUTPUT_CSV = TEMP_DIR / "tagged_expenses.csv"

load_dotenv(ROOT / ".env")


def get_tag_from_llm(description):

    client = genai.Client()

    """Calls the LLM to classify the description into a specific tag."""

    # Define your allowed categories here
    prompt = f"""
    You are a helpful categorization assistant. 
    Categorize the following item description into EXACTLY ONE of these tags: 
    [vegetables, masala, milk, paneer, utility, snacks, groceries, dining, unknown].
    
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


def tag_expenses_core():
    print(f"Initializing pipeline: Copying '{INPUT_CSV}' to '{OUTPUT_CSV}'...")
    shutil.copy(INPUT_CSV, OUTPUT_CSV)

    # Read and operate directly on the newly created working copy
    print(f"Loading '{OUTPUT_CSV}' for processing...")
    df = pd.read_csv(OUTPUT_CSV)
    # Ensure the description column exists
    if "description" not in df.columns:
        raise ValueError("CSV must contain a 'description' column.")

    print(
        "Starting LLM tagging process. This might take a moment depending on the number of rows..."
    )

    tags = []
    for index, row in df.iterrows():
        description = row["description"]

        # Skip empty descriptions
        if pd.isna(description):
            tags.append("unknown")
            continue

        tag = get_tag_from_llm(description)
        tags.append(tag)
        print(f"Row {index + 1}: '{description}' -> [{tag}]")

        # Add a tiny sleep to avoid hitting API rate limits if you have hundreds of rows
        time.sleep(0.5)

    # Add the new tags to the dataframe
    df["tag"] = tags

    # Save to a new CSV
    df.to_csv(OUTPUT_CSV, index=False)
    print(f"\nSuccess! Tagged data saved to {OUTPUT_CSV}")


def main():
    print(f"Initializing pipeline: Copying '{INPUT_CSV}' to '{OUTPUT_CSV}'...")
    shutil.copy(INPUT_CSV, OUTPUT_CSV)

    # Read and operate directly on the newly created working copy
    print(f"Loading '{OUTPUT_CSV}' for processing...")
    df = pd.read_csv(OUTPUT_CSV)
    # Ensure the description column exists
    if "description" not in df.columns:
        raise ValueError("CSV must contain a 'description' column.")

    print(
        "Starting LLM tagging process. This might take a moment depending on the number of rows..."
    )

    tags = []
    for index, row in df.iterrows():
        description = row["description"]

        # Skip empty descriptions
        if pd.isna(description):
            tags.append("unknown")
            continue

        tag = get_tag_from_llm(description)
        tags.append(tag)
        print(f"Row {index + 1}: '{description}' -> [{tag}]")

        # Add a tiny sleep to avoid hitting API rate limits if you have hundreds of rows
        time.sleep(0.5)

    # Add the new tags to the dataframe
    df["tag"] = tags

    # Save to a new CSV
    df.to_csv(OUTPUT_CSV, index=False)
    print(f"\nSuccess! Tagged data saved to {OUTPUT_CSV}")


if __name__ == "__main__":
    main()
