
# `ai_linugist` : Translate large, complex text documents with OpenAI LLM's

`ai_linugist` translates a `.docx` word file using OpenAI's API. This script unzips the `.docx` file into its elements, examines `document.xml`, which contains the document text paragraphs, temporarily replaces these text strings with unique identifiers, creates a `.json` file with said text strings and identifiers, translates the text, and reconstructs the document with the translated content. 

- Completely preserves document formatting (e.g., images, fonts, ... ).
- Can translate documents of any length.
- Efficiently translates as much text per prompt as possible.

This script has been tested on Kubuntu Linux 24.04 with Python 3.12 and OpenAI Python library 1.56.2, but should run on other systems.

## Disclaimer

This software comes with zero guarantees for correctness or otherwise, and you should not trust the translation blindly, or at all. Verify with human eyes. 


## Installation

Download the repository:

```bash
git clone https://github.com/paulzerr/ai_linguistai
```

Install Python 3 and the required openai package:

```bash
pip install "openai>=1.0.0"
```
Get an account and OpenAI API key if you don't have one yet: https://platform.openai.com/api-keys 

See section below for how to export your API key as environment variable (recommended) or include it at the top of the script.


## Usage

### Option 1: CLI

Ensure the API key is available as an environment variable (see below) and run the script from a command line interface:

**Example**
```bash
python ai_linguist.py path/to/your/document.docx --model gpt-4o --chunk_size 400000 --source_language English --target_language French
```

- `<input_docx>`: Path to the input `.docx` file. By default, the script looks for `document.docx` in the current working directory.
- `--model`: OpenAI model for translation (e.g.,`o1-preview`, `gpt-4o`). Default is `GPT-4o`. `o1-preview` is suggested for best translation performance.
- `--chunk_size`: Maximum chunk size for translation segment (in characters). This determines the maximum size per segment of the document that will be uploaded for translation. This is optional, as a safe chunk limit is determined automatically depending on the model.
- `--source_language`: The language of the input document (e.g., `English`, `Spanish`, `German`). Default is `English`.
- `--target_language`: The language to translate the document into (e.g., `French`, `Italian`, `Dutch`). Default is `French`.


### Option 2: Direct script execution

To run the script directly without input parameters, modify the `DEFAULTS` section at the top of the script, then run:

```bash
python ai_linguist.py
```


### Output

A translated `.docx` file saved in the same directory as the input file with `_translated` appended to the name. The formatting of the original document should be preserved.


## License

This project is licensed under the GNU General Public License v3.0. See the `LICENSE` file for details.





## Make API key available as environment variable (safe option)

### Linux/MacOS
Open your terminal and edit the shell configuration file `~/.bashrc` or `~/.zshrc` 

Add the following line and apply the changes:
```bash
export OPENAI_API_KEY="your_api_key"

source ~/.bashrc
```

### Windows
Open command window (cmd) and run:
```cmd
setx OPENAI_API_KEY "your_api_key"
```


## TODO:

- make non-openai models available
- allow any text document to be translated