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
    "input_file": "document.docx",  # Path to the input .docx file
    "model": "gpt-4o",  # OpenAI model to use for translation
    "chunk_size_safety_margin": 0.9,  # Safety margin for chunk size calculation
    "api_key": os.getenv("OPENAI_API_KEY", api_key),  # API key for OpenAI
    "source_language": "English",  # Default source language
    "target_language": "French"  # Default target language
}

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
    calculated_chunk_size = int(token_limit * 4 * DEFAULTS["chunk_size_safety_margin"])
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
                    print(f"Translating chunk {i+1} of {len(chunks)}...")
                    response = client.chat.completions.create(
                        model=model,
                        messages=[
                            {"role": "user", "content": f"{instruction}\n\n{chunk}"}
                        ]
                    )
                    translated = response.choices[0].message.content.strip()
                    translated_entries = translated.split("\n\n")
                    for entry in translated_entries:
                        key, value = entry.split(": ", 1)
                        translated_text_map[key] = value
                    break
                except Exception as e:
                    error_message = str(e)
                    if "context_length_exceeded" in error_message:
                        attempt += 1
                        chunk_size //= 2
                        print(f"Reducing chunk size to {chunk_size} and retrying...")
                    else:
                        raise
            else:
                print("Failed to translate chunk after 3 attempts.")
                exit(1)
        
        return translated_text_map
    except Exception as e:
        print(f"Error during translation: {e}")
        exit(1)

# Function to extract text and replace with unique identifiers
def extract_and_replace_text_with_ids(element, text_map):
    if element.text and element.text.strip():
        unique_id = str(uuid.uuid4())[:12]
        text_map[unique_id] = element.text
        element.text = unique_id

    for child in element:
        extract_and_replace_text_with_ids(child, text_map)

    if element.tail and element.tail.strip():
        unique_id = str(uuid.uuid4())[:12]
        text_map[unique_id] = element.tail
        element.tail = unique_id

# Function to replace identifiers with translated text
def replace_ids_with_translated_text(element, text_map):
    if element.text and element.text.strip() in text_map:
        element.text = text_map[element.text.strip()]

    for child in element:
        replace_ids_with_translated_text(child, text_map)

    if element.tail and element.tail.strip() in text_map:
        element.tail = text_map[element.tail.strip()]

# Main function to process the .docx file
def process_docx(input_docx_path, source_language, target_language, model="gpt-4o", chunk_size=None):
    chunk_size = determine_chunk_size(model, chunk_size)

    temp_dir = "temp_docx"
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)
    os.makedirs(temp_dir)

    try:
        with zipfile.ZipFile(input_docx_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        document_xml_path = os.path.join(temp_dir, 'word', 'document.xml')
        tree = ET.parse(document_xml_path)
        root = tree.getroot()

        text_map = {}
        extract_and_replace_text_with_ids(root, text_map)

        json_output_path = "text_map.json"
        with open(json_output_path, 'w', encoding='utf-8') as json_file:
            json.dump(text_map, json_file, ensure_ascii=False, indent=4)

        translated_text_map = translate_text_batch(text_map, source_language, target_language, model, chunk_size)

        replace_ids_with_translated_text(root, translated_text_map)
        tree.write(document_xml_path, encoding='utf-8', xml_declaration=True)

        output_docx_path = os.path.splitext(input_docx_path)[0] + "_translated.docx"
        with zipfile.ZipFile(output_docx_path, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
            for foldername, subfolders, filenames in os.walk(temp_dir):
                for filename in filenames:
                    file_path = os.path.join(foldername, filename)
                    arcname = os.path.relpath(file_path, temp_dir)
                    zip_ref.write(file_path, arcname)

        os.remove(json_output_path)
    finally:
        shutil.rmtree(temp_dir)

# Command-line interface
def main():
    parser = argparse.ArgumentParser(description="Translate a .docx file using OpenAI's API.")
    parser.add_argument("input", help="Path to the input .docx file")
    parser.add_argument("--source_language", default=DEFAULTS["source_language"], help="Source language (e.g., 'English')")
    parser.add_argument("--target_language", default=DEFAULTS["target_language"], help="Target language (e.g., 'French')")
    parser.add_argument("--model", default=DEFAULTS["model"], help="OpenAI model to use for translation")
    parser.add_argument("--chunk_size", type=int, default=None, help="Maximum chunk size for translation in characters")
    args = parser.parse_args()

    process_docx(
        args.input,
        source_language=args.source_language,
        target_language=args.target_language,
        model=args.model,
        chunk_size=args.chunk_size
    )

if __name__ == "__main__":
    main()
