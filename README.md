
# Translate word documents with OpenAI

Translates a `.docx` word file using OpenAI's API. It temporarily unzips the `.docx` file into its elements, examines `document.xml`, which contains the document text, temporarily replaces these text strings with unique identifiers, creates a `.json` file with text strings and identifiers, translates the text, and reconstructs the document with the translated content. 


## Disclaimer

This script comes with zero guarantees for correctness and you should not trust the translation blindly, or at all. Verify with human eyes.


## Compatibility

This script has been tested on Kubuntu Linux 24 with Python 3.12 and OpenAI Python library version 1.56.2, but should run on other systems.


## Installation

Download the repository:

```bash
git clone https://github.com/paulzerr/translate_openai
```

Install Python 3 and the required openai package:

```bash
pip install "openai>=1.0.0"
```
Get an account and OpenAI API key if you don't have one yet: https://platform.openai.com/api-keys 

See section below for how to export your API key as environment variable or include it at the top of the script.


## Usage

### Option 1: CLI

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

### Option 2: Direct script execution

To run the script directly without input parameters, modify the `DEFAULTS` section at the top of the script to set values for:

- `input_file`: Path to the `.docx` file.
- `model`: OpenAI model to use (optional).
- `chunk_size_safety_margin`: Safety margin for chunk size (optional).

At least input file and api key have to be specified, then run:

```bash
python translate_openai.py
```


## Output

- A translated `.docx` file saved in the same directory as the input file with `_translated` appended to the name.


## License

This project is licensed under the GNU General Public License v3.0. See the `LICENSE` file for details.





## Make API Key available as environment variable (safe option)

### Linux/MacOS
1. Open your terminal and edit your shell configuration file (`~/.bashrc` or `~/.zshrc`):
   ```bash
   export OPENAI_API_KEY="your_api_key_here"
   ```
2. Apply the changes:
   ```bash
   source ~/.bashrc  # or source ~/.zshrc
   ```

### Windows
1. Open Command Prompt and run:
   ```cmd
   setx OPENAI_API_KEY "your_api_key_here"
   ```
