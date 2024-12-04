import os
import xml.etree.ElementTree as ET
import json
import uuid
import zipfile
import shutil
import argparse
from openai import OpenAI

api_key = "your_api_key"

# Configuration section for default values
DEFAULTS = {
    "input_file": "document.docx",  # path to the input .docx file
    "model": "o1-preview",  # OpenAI model to use for translation; e.g., o1-preview, gpt-4
    "chunk_size_safety_margin": 0.9,  # Safety margin for chunk size calculation
    "api_key": os.getenv("OPENAI_API_KEY", api_key)  # API key for OpenAI
}


"""
translate_openai.py

This script translates the text content of a .docx file from one language to another using OpenAI's API. 
It extracts text from the document, replaces it with unique identifiers, translates the text, and then 
replaces the identifiers with the translated text. The script can be run from the command line or directly 
in an IDE with default settings.

Inputs:
- A .docx file containing the text to be translated.
- Optional command-line arguments to specify the model and chunk size.

Outputs:
- A new .docx file with the translated text.
- A JSON file containing the mapping of text identifiers to original text.

Configuration:
- The script uses an API key for OpenAI, which can be set via an environment variable or directly in the script.
- Default values for input file, model, and chunk size can be configured at the top of the script.

Functions:
- determine_chunk_size: Calculates the chunk size based on the model's token limit and a safety margin.
- translate_text_batch: Translates text in batches using the OpenAI API.
- extract_and_replace_text_with_ids: Extracts text from XML and replaces it with unique identifiers.
- replace_ids_with_translated_text: Replaces identifiers in XML with translated text.
- process_docx: Main function to process the .docx file, handle translation, and save the output.
- main: Command-line interface for running the script.
"""



# Model-specific token limits
MODEL_LIMITS = {
    "gpt-3.5": 4097,
    "gpt-4": 8192,
    "gpt-4-32k": 32768,
    "gpt-4o": 128000,
    "gpt-4o-mini": 128000,
    "o1-preview": 128000,
    "gpt-4-turbo": 30000,
    "gpt-3.5-turbo": 200000
}

# Function to determine chunk size based on model with a safety margin
def determine_chunk_size(model, max_chunk_size=None):
    token_limit = MODEL_LIMITS.get(model, 8192)  # Default to GPT-4 limit if model not found
    calculated_chunk_size = int(token_limit * 4 * DEFAULTS["chunk_size_safety_margin"])  # Convert tokens to characters with safety margin
    if max_chunk_size is not None:
        return min(calculated_chunk_size, max_chunk_size)
    return calculated_chunk_size

# Initialize the OpenAI client
client = OpenAI(api_key=DEFAULTS["api_key"])

# Function to translate a batch of text using OpenAI API
def translate_text_batch(text_map, source_language, target_language, model, chunk_size):
    print("Preparing to translate text batch...")
    instruction = (
        f"You are a translator. Translate the following text from {source_language} to {target_language}. "
        "Only translate the text content and do not alter any code, tags, or placeholders."
    )
    try:
        # Split the text map into chunks of approximately chunk_size characters
        print("Splitting text into manageable chunks...")
        chunks = []
        current_chunk = []
        current_length = 0
        for key, value in text_map.items():
            entry = f"{key}: {value}"
            if current_length + len(entry) > chunk_size:
                chunks.append("\n\n".join(current_chunk))
                current_chunk = []
                current_length = 0
            current_chunk.append(entry)
            current_length += len(entry)
        if current_chunk:
            chunks.append("\n\n".join(current_chunk))

        translated_text_map = {}
        for i, chunk in enumerate(chunks):
            attempt = 0
            while attempt < 3:
                try:
                    print(f"Translating chunk {i+1} of {len(chunks)} with chunk size {chunk_size}...")
                    response = client.chat.completions.create(
                        model=model,
                        messages=[
                            {"role": "user", "content": f"{instruction}\n\n{chunk}"}
                        ]
                    )
                    translated = response.choices[0].message.content.strip()
                    
                    # Split the translated text back into individual entries
                    translated_entries = translated.split("\n\n")
                    for entry in translated_entries:
                        key, value = entry.split(": ", 1)
                        translated_text_map[key] = value
                    break  # Exit the retry loop if successful
                except Exception as e:
                    error_message = str(e)
                    print(f"Error during translation: {error_message}")
                    if "context_length_exceeded" in error_message:
                        attempt += 1
                        chunk_size //= 2
                        print(f"Reducing chunk size to {chunk_size} and retrying...")
                    else:
                        raise  # Raise the exception if it's not a context length error
            else:
                print("Failed to translate chunk after 3 attempts. Exiting.")
                exit(1)
        
        print("Completed translation of all chunks.")
        return translated_text_map
    except Exception as e:
        print(f"Unexpected error during translation: {e}")
        exit(1)  # Exit if an unexpected error occurs

# Function to extract text and replace with unique identifiers
def extract_and_replace_text_with_ids(element, text_map):
    if element.text and element.text.strip():
        original_text = element.text
        unique_id = str(uuid.uuid4())[:12]  # Shorten the UUID to 12 characters
        text_map[unique_id] = original_text
        element.text = unique_id
        print(f"Replaced text with ID: {unique_id}")

    for child in element:
        extract_and_replace_text_with_ids(child, text_map)

    if element.tail and element.tail.strip():
        original_tail = element.tail
        unique_id = str(uuid.uuid4())[:12]  # Shorten the UUID to 12 characters
        text_map[unique_id] = original_tail
        element.tail = unique_id
        print(f"Replaced tail with ID: {unique_id}")

# Function to replace identifiers with translated text
def replace_ids_with_translated_text(element, text_map):
    if element.text and element.text.strip() in text_map:
        element.text = text_map[element.text.strip()]
        print(f"Replaced ID with translated text: {element.text[:100]}...")

    for child in element:
        replace_ids_with_translated_text(child, text_map)

    if element.tail and element.tail.strip() in text_map:
        element.tail = text_map[element.tail.strip()]
        print(f"Replaced tail ID with translated text: {element.tail[:100]}...")

# Main function to process the .docx file
def process_docx(input_docx_path, source_language="fr", target_language="nl", model="gpt-4o", chunk_size=None):
    print(f"Starting processing of .docx file: {input_docx_path}")
    # Determine chunk size based on model if not provided
    chunk_size = determine_chunk_size(model, chunk_size)
    print(f"Using chunk size: {chunk_size} for model: {model}")

    # Create a temporary directory to extract the .docx contents
    temp_dir = "temp_docx"
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)
    os.makedirs(temp_dir)
    print(f"Temporary directory created: {temp_dir}")

    try:
        # Extract the .docx file
        print("Extracting .docx file contents...")
        with zipfile.ZipFile(input_docx_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
        print("Extraction complete.")

        # Path to the document.xml file
        document_xml_path = os.path.join(temp_dir, 'word', 'document.xml')
        print(f"Document XML path: {document_xml_path}")

        # Parse the XML file
        print("Parsing XML file...")
        tree = ET.parse(document_xml_path)
        root = tree.getroot()
        print("XML parsing complete.")

        # Extract text and replace with unique identifiers
        print("Extracting text and replacing with unique identifiers...")
        text_map = {}
        extract_and_replace_text_with_ids(root, text_map)
        print("Text extraction and replacement complete.")

        # Write the text map to a JSON file
        json_output_path = "text_map.json"
        print(f"Writing text map to JSON file: {json_output_path}")
        with open(json_output_path, 'w', encoding='utf-8') as json_file:
            json.dump(text_map, json_file, ensure_ascii=False, indent=4)
        print("Text map written to JSON file.")

        # Translate the entire text map using the OpenAI API
        print("Starting translation of text map...")
        translated_text_map = translate_text_batch(text_map, source_language, target_language, model, chunk_size)
        print("Translation of text map complete.")

        # Replace identifiers with translated text
        print("Replacing identifiers with translated text...")
        replace_ids_with_translated_text(root, translated_text_map)
        print("Replacement of identifiers complete.")

        # Write the modified XML back to the document.xml file
        print("Writing modified XML back to document.xml...")
        tree.write(document_xml_path, encoding='utf-8', xml_declaration=True)
        print("Modified XML written back to document.xml.")

        # Create a new .docx file from the modified contents
        output_docx_path = os.path.splitext(input_docx_path)[0] + "_translated.docx"
        print(f"Creating new .docx file: {output_docx_path}")
        with zipfile.ZipFile(output_docx_path, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
            for foldername, subfolders, filenames in os.walk(temp_dir):
                for filename in filenames:
                    file_path = os.path.join(foldername, filename)
                    arcname = os.path.relpath(file_path, temp_dir)
                    zip_ref.write(file_path, arcname)
        print(f"New .docx file created: {output_docx_path}")

    except FileNotFoundError:
        print(f"Error: Input file {input_docx_path} not found.")
    except ET.ParseError as pe:
        print(f"Error parsing XML file: {pe}")
    except Exception as e:
        print(f"Unexpected error: {e}")
    finally:
        # Clean up the temporary directory
        print("Cleaning up temporary directory...")
        shutil.rmtree(temp_dir)
        print("Temporary directory cleaned up.")

# Command-line interface
def main():
    parser = argparse.ArgumentParser(description="Translate a .docx file using OpenAI's API.")
    parser.add_argument("input", nargs='?', default=None, help="Path to the input .docx file")
    parser.add_argument("--model", default=DEFAULTS["model"], help="OpenAI model to use for translation")
    parser.add_argument("--chunk_size", type=int, default=None, help="Maximum chunk size for translation in characters")
    args = parser.parse_args()

    # Use command-line arguments if provided, otherwise use defaults
    input_file = args.input if args.input else DEFAULTS["input_file"]
    model = args.model if args.model else DEFAULTS["model"]
    chunk_size = args.chunk_size if args.chunk_size else None

    process_docx(input_file, model=model, chunk_size=chunk_size)

if __name__ == "__main__":
    main()
