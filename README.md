# X-to-TXT: Command-Line Text Extraction Tool

A command-line tool for extracting text from PDF, DOCX, and PPTX files with robust support for all formats.

## Features

- Extract text from PDF, DOCX, and PPTX files
- Automatic format detection based on file extension
- Output to console or save to file/directory
- Process multiple files in a single command
- Verbose mode with processing details
- Cross-platform (Windows, macOS, Linux)

## Installation

### Option 1: Quick Install (Local Development)

1. Ensure you have Python 3.x installed
2. Install dependencies:

```
pip install -r requirements.txt
```

3. Make the script executable (Unix-based systems):

```
chmod +x extract_text.py
```

4. Run the script with:

```
./extract_text.py [options] <file_path>
```

### Option 2: System-wide Installation (Recommended)

1. Ensure you have Python 3.x installed
2. Install the package:

```
python -m pip install -e .
```

3. Now you can run the command from anywhere:

```
extract_text [options] <file_path>
```

## Usage

Basic usage:

```
extract_text [options] <file_path> [<file_path> ...]
```

Options:

- `-f, --format FORMAT`: Specify file format (`pdf`, `docx`, `pptx`) - optional
- `-o, --output OUTPUT`: Specify output file or directory - optional
- `-v, --verbose`: Enable verbose output - optional
- `-h, --help`: Display help message

## Examples

Extract text from a PPTX file to the console:

```
extract_text presentation.pptx
```

Extract text from multiple DOCX files to a directory:

```
extract_text -o output_dir/ document1.docx document2.docx
```

Extract text from a PDF with verbose output:

```
extract_text -v document.pdf
```

Specify format explicitly:

```
extract_text -f pptx -o output.txt presentation.pptx
```

## Limitations

- PDF: Only extracts text from text-based PDFs (not image-based/scanned)
- DOCX: Limited to paragraph text (no headers, footers, tables)
- PPTX: Extracts text from slide shapes only (no notes or master slides)
- All formats: Returns plain text without formatting

## License

MIT
