
# Translate OpenAI Script

Translates text content from `.docx` files using OpenAI's API. It processes the document, replaces text with unique identifiers, translates the text, and reconstructs the document with the translated content. The script can be customized by modifying default values or by using command-line arguments.


## Compatibility

- This script has been tested on Python 3.12, OpenAI Python library version 1.56.2, on Kubuntu Linux 24. It should work on other systems.


## Installation

Install Python and the required openai package:

```bash
pip install "openai>=1.0.0"
```

## Usage

### Command Line Interface

Run the script from the command line:

```bash
python translate_openai.py <input_docx> --model <model> --chunk_size <chunk_size>
```

**Example**:
```bash
python translate_openai.py document.docx --model gpt-4o --chunk_size 400000
```

- `<input_docx>`: Path to the input `.docx` file (optional; defaults to `document.docx`).
- `--model`: OpenAI model for translation (e.g., `gpt-4o`, `gpt-4`).
- `--chunk_size`: Maximum chunk size for translation in characters (optional, useful for very large files).

### Direct Script Execution

To run the script directly without input parameters, modify the `DEFAULTS` section at the top of the script to set values for:

- `input_file`: Path to the `.docx` file.
- `model`: OpenAI model to use.
- `chunk_size_safety_margin`: Safety margin for chunk size.

At least input file and api key has to be specified, then run:

```bash
python translate_openai.py
```


## Output

- A translated `.docx` file saved in the same directory as the input file with `_translated` appended to the name.
- A JSON file named `text_map.json` containing original text mappings.

## Notes

- Ensure your OpenAI API key is set as an environment variable (`OPENAI_API_KEY`) or included in the script.
- The script handles chunk size dynamically to avoid exceeding token limits.

## License

This project is licensed under the GNU General Public License v3.0. See the `LICENSE` file for details.
